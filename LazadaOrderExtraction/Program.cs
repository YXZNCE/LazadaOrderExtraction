using System.Diagnostics;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using LazadaOrderExtraction.Data;
using PuppeteerSharp;

namespace LazadaOrderExtraction
{
    public static class Program
    {
        public static async Task Main(string[] args)
        {
            Console.Title = "Lazada [PH] Total Order Extraction";
            await EnsureChromiumAsync();

            var launchOptions = new LaunchOptions
            {
                Headless = false,
            };

            await using var browser = await Puppeteer.LaunchAsync(launchOptions);
            await using var page = await browser.NewPageAsync();

            await GoToLazadaHomeAsync(page);

            LogWarning("Please manually log in. Press any key once you're logged in and at the homepage.");
            Console.ReadKey(intercept: true);

            LogInfo("Navigating to orders page...");
            await page.GoToAsync("https://my.lazada.com.ph/customer/order/index/");

            int totalPages = await GetTotalPages(page);
            LogSuccess($"Detected total pages: {totalPages}");

            var (allOrders, overallTotal, cancelledSum, receivedSum) =
                await ProcessAllPages(page, totalPages);


            var statusGroups = allOrders
                .GroupBy(o => o.DeliveryStatus)
                .Select(g => new
                {
                    Status = g.Key,
                    Count = g.Count(),
                    Sum = g.Sum(x => x.TotalOrderPrice)
                })
                .OrderByDescending(x => x.Count)
                .ToList();

            Console.WriteLine("---- Status Breakdown ----");
            foreach (var sg in statusGroups)
            {
                Console.WriteLine($"Status: '{sg.Status}', Count: {sg.Count}, Sum: {sg.Sum}");
            }

            PrintSummary(allOrders, overallTotal, cancelledSum, receivedSum);

            ExportToExcel(allOrders, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Orders.xlsx"));

            LogInfo("Extraction complete. Press any key to exit...");
            Console.ReadKey(intercept: true);
        }

        #region ==== Initialization / Setup ====

        private static async Task EnsureChromiumAsync()
        {
            var browserFetcher = new BrowserFetcher();
            LogInfo("Checking/Downloading Chromium if not present...");
            await browserFetcher.DownloadAsync();
            LogSuccess("Chromium is ready.");
        }

        private static async Task GoToLazadaHomeAsync(IPage page)
        {
            LogInfo("Navigating to Lazada.com.ph...");
            await page.GoToAsync("https://lazada.com.ph/");
        }

        #endregion

        #region ==== Pagination and Page Navigation ====

        private static async Task<int> GetTotalPages(IPage page)
        {
            return await page.EvaluateFunctionAsync<int>(@"
                () => {
                    const lastPageButton = document.querySelector('.next-pagination-list button:last-child');
                    return lastPageButton ? parseInt(lastPageButton.innerText, 10) : 1;
                }
            ");
        }

        private static async Task GoToNextPage(IPage page)
        {
            var nextButton = await page.QuerySelectorAsync(".next-pagination-item.next");
            if (nextButton != null)
            {
                LogInfo("Clicking 'Next' to move to next page...");
                await nextButton.ClickAsync();
                await Task.Delay(2000);
            }
            else
            {
                LogWarning("'Next' button not found, possibly on the last page.");
            }
        }

        #endregion

        #region ==== Main Data Extraction ====

        private static async Task<(List<Order> allOrders, double overallTotal, double cancelledSum, double receivedSum)> ProcessAllPages(IPage page, int totalPages)
        {
            var stopwatch = Stopwatch.StartNew();

            var allOrders = new List<Order>();
            double overallTotal = 0;
            double cancelledSum = 0;
            double receivedSum = 0;

            for (int currentPage = 1; currentPage <= totalPages; currentPage++)
            {
                LogInfo($"Processing Page {currentPage}/{totalPages}...");

                var ordersOnThisPage = await ScrapeOrdersOnPage(page);
                LogInfo($"Fetched {ordersOnThisPage.Count} orders from Page {currentPage}.");

                allOrders.AddRange(ordersOnThisPage);

                double pageTotal = ordersOnThisPage.Sum(o => o.TotalOrderPrice);
                overallTotal += pageTotal;

                foreach (var order in ordersOnThisPage)
                {
                    if (order.DeliveryStatus?.Equals("Cancelled", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        cancelledSum += order.TotalOrderPrice;
                    }
                    else if (order.DeliveryStatus?.Equals("Received", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        receivedSum += order.TotalOrderPrice;
                    }
                    else if (order.DeliveryStatus?.Equals("Delivered", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        receivedSum += order.TotalOrderPrice;
                    }
                }

                if (currentPage < totalPages)
                {
                    await GoToNextPage(page);
                }

                // (Optional) Log details for this page's orders
                // PrintOrderDetails(ordersOnThisPage, currentPage);
            }

            stopwatch.Stop();
            LogSuccess($"All pages processed in {stopwatch.Elapsed.TotalSeconds:F2} seconds.");

            return (allOrders, overallTotal, cancelledSum, receivedSum);
        }
        private static async Task<List<Order>> ScrapeOrdersOnPage(IPage page)
        { 
            var orderElements = await page.QuerySelectorAllAsync(".order-list [tag=\"order-component\"]");
            var orders = new List<Order>();
            foreach (var orderEl in orderElements)
            {
                var shopNameEl = await orderEl.QuerySelectorAsync(".shop-left-info-name");
                string shopName = "";
                if (shopNameEl != null)
                    shopName = await shopNameEl.EvaluateFunctionAsync<string>("el => el.innerText.trim()") ?? "";

                var statusEl = await orderEl.QuerySelectorAsync(".shop-right-status");
                string deliveryStatus = "";
                if (statusEl != null)
                    deliveryStatus = await statusEl.EvaluateFunctionAsync<string>("el => el.innerText.trim()") ?? "";
                

                var itemElements = await orderEl.QuerySelectorAllAsync(".order-item");
                var items = new List<OrderItem>();
                foreach (var itemEl in itemElements)
                {
                    string itemName = "";
                    var titleEl = await itemEl.QuerySelectorAsync(".text.title.item-title");
                    if (titleEl != null)
                        itemName = await titleEl.EvaluateFunctionAsync<string>("el => el.innerText.trim()") ?? "";
                    
                    double price = 0.0;
                    var priceEl = await itemEl.QuerySelectorAsync(".item-price");
                    if (priceEl != null)
                    {
                        var rawPriceText = await priceEl.EvaluateFunctionAsync<string>("el => el.innerText.trim()");
                        var cleanedPrice = rawPriceText?.Replace("₱", "").Replace(",", "").Trim() ?? "0";
                        if (double.TryParse(cleanedPrice, out double parsed))
                            price = parsed;
                    }
                    
                    int quantity = 0;
                    var qtyEl = await itemEl.QuerySelectorAsync(".item-quantity .text.desc.info.multiply + .text");
                    if (qtyEl != null)
                    {
                        var rawQtyText = await qtyEl.EvaluateFunctionAsync<string>("el => el.innerText");
                        var cleanedQty = Regex.Replace(rawQtyText ?? "0", @"[^\d]", "");
                        if (int.TryParse(cleanedQty, out int parsedQty))
                            quantity = parsedQty;
                        
                    }

                    bool isRefunded = false;
                    bool isCancelled = false;
                    var capsuleEl = await itemEl.QuerySelectorAsync(".item-status.item-capsule");
                    if (capsuleEl != null)
                    {
                        var capsuleText =
                            await capsuleEl.EvaluateFunctionAsync<string>("el => el.innerText.trim().toLowerCase()");
                        switch (string.IsNullOrEmpty(capsuleText))
                        {
                            case false when capsuleText.Contains("refund"):
                                isRefunded = true;
                                break;
                            case false when capsuleText.Contains("cancel"):
                                isCancelled = true;
                                break;
                        }
                    }

                    items.Add(new OrderItem
                    {
                        ItemName = itemName,
                        Price = price,
                        Quantity = quantity,
                        IsRefunded = isRefunded,
                        IsCancelled = isCancelled
                    });
                }
                
                double totalOrderPrice = items.Where(it => !it.IsRefunded).Sum(it => it.Price * it.Quantity);
                orders.Add(new Order
                {
                    ShopName = shopName,
                    DeliveryStatus = deliveryStatus,
                    Items = items,
                    TotalOrderPrice = totalOrderPrice
                });
            }

            return orders;
        }

        #endregion

        #region ==== Printing / Logging Helpers ====

        private static void PrintSummary(
            List<Order> allOrders,
            double overallTotal,
            double cancelledSum,
            double receivedSum
        )
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\n========== FINAL SUMMARY ==========");
            Console.ResetColor();

            Console.WriteLine($"All Orders Count: {allOrders.Count}");

            int receivedCount = allOrders.Count(o =>
                o.DeliveryStatus?.Equals("Received", StringComparison.OrdinalIgnoreCase) == true
            );
            int deliveredCount = allOrders.Count(o =>
                o.DeliveryStatus?.Equals("Delivered", StringComparison.OrdinalIgnoreCase) == true
            );
            int cancelledCount = allOrders.Count(o =>
                o.DeliveryStatus?.Equals("Cancelled", StringComparison.OrdinalIgnoreCase) == true
            );

            Console.WriteLine($"Received + Delivered Orders Count: {receivedCount + deliveredCount}");
            Console.WriteLine($"Cancelled Orders Count: {cancelledCount}");

            Console.WriteLine($"Overall Total       : PHP {overallTotal}");
            Console.WriteLine($"Received Total      : PHP {receivedSum}");
            Console.WriteLine($"Cancelled Total     : PHP {cancelledSum}");


            Console.WriteLine("===================================\n");
        }

        private static void ExportToExcel(List<Order> allOrders, string filePath)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Lazada Orders");

            worksheet.Cell("A1").Value = "Item Name";
            worksheet.Cell("B1").Value = "Item Price";
            worksheet.Cell("C1").Value = "Item Count (Quantity)";
            worksheet.Cell("D1").Value = "Cancelled / Refunded";

            worksheet.Row(1).Style.Font.Bold = true;

            int currentRow = 2;  
            double grandTotal = 0.0;

            foreach (var item in allOrders.SelectMany(order => order.Items))
            {
                // if skip refunded items uncomment the code below
                // if (item.IsRefunded) continue;

                worksheet.Cell(currentRow, 1).Value = item.ItemName;
                worksheet.Cell(currentRow, 2).Value = item.Price;
                worksheet.Cell(currentRow, 3).Value = item.Quantity;
                worksheet.Cell(currentRow, 4).Value = item.IsRefunded || item.IsCancelled;

                grandTotal += (item.IsRefunded ? 0 : item.Price * item.Quantity);

                currentRow++;
            }

            currentRow++;
            worksheet.Cell(currentRow, 1).Value = "Grand Total:";
            worksheet.Cell(currentRow, 2).Value = grandTotal;
            worksheet.Range($"A{currentRow}:B{currentRow}").Style.Font.Bold = true;

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);

            LogInfo($"Excel file saved to: {filePath}");
        }

        // debug
        private static void PrintOrderDetails(List<Order> ordersOnThisPage, int currentPage)
        {
            Console.WriteLine($"\n--- Page {currentPage} Orders ---");
            foreach (var order in ordersOnThisPage)
            {
                Console.WriteLine(new string('-', 50));
                Console.WriteLine($"Shop: {order.ShopName}");
                Console.WriteLine($"Delivery Status: {order.DeliveryStatus}");
                Console.WriteLine("Items:");
                foreach (var i in order.Items)
                {
                    Console.WriteLine($"  - ItemName: {i.ItemName}");
                    Console.WriteLine($"    Price   : {i.Price}");
                    Console.WriteLine($"    Quantity: {i.Quantity}");
                    Console.WriteLine($"    Refunded: {i.IsRefunded}");
                }

                Console.WriteLine($"Order Total: {order.TotalOrderPrice}");
            }

            Console.WriteLine();
        }

        #endregion

        #region ==== Console Logging ====

        private static void LogInfo(string message)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"[INFO] {message}");
            Console.ResetColor();
        }

        private static void LogWarning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"[WARN] {message}");
            Console.ResetColor();
        }

        private static void LogSuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[OK]   {message}");
            Console.ResetColor();
        }

        private static void LogError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERR]  {message}");
            Console.ResetColor();
        }

        #endregion
    }
}
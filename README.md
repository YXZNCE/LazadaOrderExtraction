# Lazada PH Order Extraction

A simple .NET console application that uses **PuppeteerSharp** to scrape order details like price, quantity, name from Lazada Philippines.  
this collects order data across multiple pages, and can export it to an Excel file for easy viewing or bookkeeping.

## Disclaimer
**⚠️ Important Notice ⚠️**

This software is provided **"as is"** without any warranties, express or implied. The author **does not** accept any responsibility or liability for any damages, losses, or issues arising from the use or misuse of this tool.

### Intended Use

This tool is **exclusively** designed for **personal use** to help users manage and optimize their own Lazada [PH] orders. It is **not** intended for commercial purposes, data mining, or any activities that violate Lazada's [Terms & Conditions](https://www.lazada.com.ph/terms-conditions/).

### Compliance

Users are **responsible** for ensuring that their use of this software complies with all applicable laws, regulations, and Lazada's [Platform Engagement Tools Terms & Conditions](https://www.lazada.com.ph/terms-conditions/). The author **does not** endorse or support any actions that infringe upon these terms.

### No Warranty

The author makes **no guarantees** regarding the functionality, accuracy, or reliability of this tool. Use it at your **own risk**.

### Limitation of Liability

Under no circumstances shall the author be liable for any direct, indirect, incidental, special, or consequential damages resulting from the use or inability to use this software.
 

## Features

- **Manual Login**: Allows you to log in via a real browser window (headless = false).  
- **Scrapes All Pages**: Automatically navigates through pagination, collecting each order.  
- **Item-Level Refund Detection**: Skips or flags refunded items so they don’t inflate totals.  
- **Exports to Excel**: Writes out item details (name, price, quantity, date) and calculates grand total.  

## Requirements

1. [.NET 6+](https://dotnet.microsoft.com/en-us/download/dotnet) installed.  
2. [PuppeteerSharp](https://github.com/hardkoded/puppeteer-sharp) via NuGet:
   ```bash
   dotnet add package PuppeteerSharp
   ```
3. [ClosedXML](https://github.com/ClosedXML/ClosedXML) (if you plan to export to `.xlsx`):
   ```bash
   dotnet add package ClosedXML
   ```

## Usage

1. **Clone/Download** this repository.
2. **Open** the project in your preferred IDE (Visual Studio, Rider, VSCode) or via command line.
3. In **`Program.cs`**, make sure `Headless = false` so you can log in manually:
   ```csharp
   var launchOptions = new LaunchOptions
   {
       Headless = false
   };
   ```
4. **Run** the application:
   ```bash
   dotnet run
   ```
5. A Chromium browser will launch. **Log in** to your Lazada PH account.  
6. **Press any key** in the console once you are logged in. The scraper then navigates to your orders page and collects data.
7. After scraping all pages, **view** the summary in the console.  
8. (Optional) Export the results to an Excel file via a helper method such as `ExportToExcel(allOrders, "output.xlsx")`.

## Preview Image

![Preview Image](https://github.com/YXZNCE/LazadaOrderExtraction/blob/master/Images/preview.png?raw=true)

## Additional Notes

- **Timeouts**: if pages take longer to load, adjust the `Task.Delay(...)` or use `WaitForSelectorAsync(...)` to make sure content is fully loaded.  
- **Data Privacy**: the code logs into Lazada using your credentials, so keep your code private and handle any saved credentials securely.

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/YXZNCE/LazadaOrderExtraction/blob/master/LICENSE.txt) file for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LazadaOrderExtraction.Data
{
    public class Order
    {
        [JsonPropertyName("shopName")]
        public string ShopName { get; set; }

        [JsonPropertyName("deliveryStatus")]
        public string DeliveryStatus { get; set; }

        [JsonPropertyName("items")]
        public List<OrderItem> Items { get; set; } = new List<OrderItem>();

        [JsonPropertyName("totalOrderPrice")]
        public double TotalOrderPrice { get; set; }
    }
}

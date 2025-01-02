using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LazadaOrderExtraction.Data
{
    public class OrderItem
    {
        [JsonPropertyName("itemName")]
        public string ItemName { get; set; }

        [JsonPropertyName("price")]
        public double Price { get; set; }

        [JsonPropertyName("quantity")]
        public int Quantity { get; set; }
        
        
        public bool IsRefunded { get; set; }
        public bool IsCancelled { get; set; }

    }
}

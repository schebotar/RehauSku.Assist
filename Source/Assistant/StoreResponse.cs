using System.Collections.Generic;

namespace RehauSku.Assistant
{
    public class StoreResponce
    {
        public Ecommerce Ecommerce { get; set; }
    }

    public class Ecommerce
    {
        public List<Product> Impressions { get; set; }
    }

    public class Product : IProduct
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Price { get; set; }
    }
}
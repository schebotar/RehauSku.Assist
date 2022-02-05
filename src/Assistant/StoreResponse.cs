using System.Collections.Generic;

namespace RehauSku.Assistant
{
    class StoreResponce
    {
        public Ecommerce Ecommerce { get; set; }
    }

    class Ecommerce
    {
        public List<Product> Impressions { get; set; }
    }

    class Product : IProduct
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Price { get; set; }
    }
}
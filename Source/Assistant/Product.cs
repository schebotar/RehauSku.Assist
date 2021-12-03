using System;

namespace Rehau.Sku.Assist
{
    public class Product : IProduct
    {
        public string Sku { get; }
        public string Name { get; }
        public Uri Uri { get; }

        public Product(string sku, string name)
        {
            Sku = sku;
            Name = name;
        }

        public Product(string sku, string name, string uri)
        {
            Sku = sku;
            Name = name;
            Uri = new Uri(uri);
        }

        public override string ToString()
        {
            return $"{this.Name} ({this.Sku})";
        }
    }
}
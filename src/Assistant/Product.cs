namespace Rehau.Sku.Assist
{
    public class Product : IProduct
    {
        public string Sku { get; }
        public string Name { get; }

        public string Uri => throw new System.NotImplementedException();

        public Product(string sku, string name)
        {
            Sku = sku;
            Name = name;
        }

        public override string ToString()
        {
            return $"{this.Name} ({this.Sku})";
        }
    }
}
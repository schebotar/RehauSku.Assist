namespace RehauSku.PriceListTools
{
    public class Position
    {
        public string Group { get; private set; }
        public string Sku { get; private set; }
        public string Name { get; private set; }

        public Position(string group, string sku, string name)
        {
            Group = group;
            Sku = sku;
            Name = name;
        }
    }
}


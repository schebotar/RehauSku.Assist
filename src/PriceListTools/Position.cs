using System.Linq;

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

        public override bool Equals(object obj)
        {
            if (obj as Position == null)
                return false;

            Position other = obj as Position;

            return Group == other.Group &&
                Sku == other.Sku &&
                Name == other.Name;
        }

        public override int GetHashCode()
        {
            string[] properties = new[]
            {
                Group,
                Sku,
                Name
            };

            return string.Concat(properties.Where(p => p != null)).GetHashCode();
        }
    }
}
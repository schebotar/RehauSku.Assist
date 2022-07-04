using System.Text.RegularExpressions;

namespace RehauSku
{
    internal class RauSku
    {
        public string Sku { get; private set; }
        public string Variant { get; private set; }

        public RauSku(string sku, string variant)
        {
            Sku = sku;
            Variant = variant;
        }

        public static bool TryParse(string line, out RauSku rehauSku)
        {
            Match match;
            match = Regex.Match(line, @"\b[1]\d{6}[1]\d{3}\b");
            if (match.Success)
            {
                string sku = match.Value.Substring(1, 6);
                string variant = match.Value.Substring(8, 3);
                rehauSku = new RauSku(sku, variant);
                return true;
            }

            match = Regex.Match(line, @"\b\d{6}\D\d{3}\b");
            if (match.Success)
            {
                string sku = match.Value.Substring(0, 6);
                string variant = match.Value.Substring(7, 3);
                rehauSku = new RauSku(sku, variant);
                return true;
            }

            match = Regex.Match(line, @"\b\d{9}\b");
            if (match.Success)
            {
                string sku = match.Value.Substring(0, 6);
                string variant = match.Value.Substring(6, 3);
                rehauSku = new RauSku(sku, variant);
                return true;
            }

            match = Regex.Match(line, @"\b\d{6}\b");
            if (match.Success)
            {
                string sku = match.Value.Substring(0, 6);
                string variant = "001";
                rehauSku = new RauSku(sku, variant);
                return true;
            }

            else
            {
                rehauSku = null;
                return false; 
            }
        }

        public override string ToString()
        {
            return $"1{Sku}1{Variant}";
        }
    }
}
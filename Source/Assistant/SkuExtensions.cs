using System.Text.RegularExpressions;

namespace RehauSku.Assistant
{
    static class SkuExtensions
    {
        public static bool IsRehauSku(this string line)
        {
            return Regex.IsMatch(line, @"^[1]\d{6}[1]\d{3}$");
        }
    }
}
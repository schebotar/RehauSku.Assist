using ExcelDna.Integration;

namespace Rehau.Sku.Assist
{
    public class Functions
    {
        [ExcelFunction]
        public static object RAUNAME(string request)
            => SkuAssist.GetProduct(request, ProductField.Name);

        [ExcelFunction]
        public static object RAUSKU(string request)
            => SkuAssist.GetProduct(request, ProductField.Id);

        [ExcelFunction]
        public static object RAUPRICE(string request)
            => SkuAssist.GetProduct(request, ProductField.Price);
    }
}
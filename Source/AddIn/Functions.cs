using ExcelDna.Integration;
using RehauSku.Assistant;

namespace RehauSku
{
    public class Functions
    {
        [ExcelFunction(Description = "Получение названия первого продукта по запросу в интернет-магазин REHAU")]
        public static object RAUNAME([ExcelArgument(Name = "Запрос", Description = "Запрос в свободной форме или ячейка с запросом")] string request)
            => MakeRequest(request, ProductField.Name);

        [ExcelFunction(Description = "Получение артикула первого продукта по запросу в интернет-магазин REHAU")]
        public static object RAUSKU([ExcelArgument(Name = "Запрос", Description = "Запрос в свободной форме или ячейка с запросом")] string request)
            => MakeRequest(request, ProductField.Id);

        [ExcelFunction(Description = "Получение цены первого продукта по запросу в интернет-магазин REHAU")]
        public static object RAUPRICE([ExcelArgument(Name = "Запрос", Description = "Запрос в свободной форме или ячейка с запросом")] string request)
            => MakeRequest(request, ProductField.Price);

        private static object MakeRequest(string request, ProductField field)
        {
            object result;

            if (request.IsCached())
                result = request.GetFromCache();

            else
            {
                result = ExcelAsyncUtil.Run("Request", request, delegate
                {
                    return request.RequestAndCache().GetAwaiter().GetResult();
                });
            }

            if (result == null)
                return "Не найдено :(";

            if (result.Equals(ExcelError.ExcelErrorNA))
                return "Загрузка...";

            IProduct product = result as IProduct;

            switch (field)
            {
                case ProductField.Name:
                    return product.Name;
                case ProductField.Id:
                    return product.Id;
                case ProductField.Price:
                    return double.Parse(product.Price, System.Globalization.CultureInfo.InvariantCulture);
                default:
                    return null;
            }
        }
    }
}
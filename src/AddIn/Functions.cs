using ExcelDna.Integration;
using RehauSku.Assistant;

namespace RehauSku
{
    public class Functions
    {
        [ExcelFunction(description: "Получение названия первого продукта в поиске")]
        public static object RAUNAME([ExcelArgument(Name = "\"Запрос\"", Description = "в свободной форме или ячейка с запросом")] string request)
            => MakeRequest(request, ProductField.Name);

        [ExcelFunction(Description = "Получение артикула первого продукта в поиске")]
        public static object RAUSKU([ExcelArgument(Name = "\"Запрос\"", Description = "в свободной форме или ячейка с запросом")] string request)
            => MakeRequest(request, ProductField.Id);

        [ExcelFunction(Description = "Получение цены первого продукта в поиске")]
        public static object RAUPRICE([ExcelArgument(Name = "\"Запрос\"", Description = "в свободной форме или ячейка с запросом")] string request)
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

        [ExcelFunction(Description = "Получение корректного артикула из строки")]
        public static object GETRAUSKU([ExcelArgument(Name = "\"Строка\"", Description = "строка, содержащая актикул")] string line)
        {
            RauSku rausku;

            if (RauSku.TryParse(line, out rausku))
            {
                return rausku.ToString();
            }

            else return ExcelError.ExcelErrorNA;
        }
    }
}
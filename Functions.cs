using ExcelDna.Integration;
using System.Runtime.Caching;
using System.Net.Http;

namespace Rehau.Sku.Assist
{
    public class Functions : IExcelAddIn
    {
        private static HttpClient _httpClient;
        private static ObjectCache _resultCache = MemoryCache.Default;

        public void AutoClose()
        {
        }
        public void AutoOpen()
        {
            _httpClient = new HttpClient();
        }

        [ExcelFunction]
        public static object RAUNAME(string request)
        {
            string cachedResult = _resultCache[request] as string;

            if (cachedResult != null)
            {
                return cachedResult;
            }

            else
            {
                object result = ExcelAsyncUtil.Run("RAUNAME", null,
                    delegate
                {
                    var document = SkuAssist.GetDocumentAsync(request, _httpClient).Result;
                    var product = SkuAssist.GetProductFromDocument(document);
                    return product.ToString();
                });

                if (result.Equals(ExcelError.ExcelErrorNA))
                {
                    return "Загрузка...";
                }

                else
                {
                    _resultCache.Add(request, result, System.DateTime.Now.AddMinutes(20));
                    return result.ToString();
                }
            }
        }
    }
}
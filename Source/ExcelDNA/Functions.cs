using ExcelDna.Integration;
using System.Threading.Tasks;
using System.Runtime.Caching;

namespace Rehau.Sku.Assist
{
    public class Functions
    {
        [ExcelFunction]
        public static object RAUNAME(string request)
        {
            if (MemoryCache.Default.Contains(request))
                return MemoryCache.Default[request].ToString();

            else
            {
                object result = ExcelAsyncUtil.Run("Rauname", new[] { request },
                    delegate
                    {
                        Task<IProduct> product = Task.Run(() => SkuAssist.GetProduct(request));
                        return product.Result;
                    });

                if (Equals(result, ExcelError.ExcelErrorNA))
                {
                    return "Загрузка...";
                }
                else
                {
                    MemoryCache.Default.Add(request, result, System.DateTime.Now.AddMinutes(10));
                    return result == null ? "Не найдено" : result.ToString();
                }
            }
        }
    }
}
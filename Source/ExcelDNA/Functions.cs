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
            {
                IProduct product = MemoryCache.Default[request] as IProduct;
                return product.Name;
            }

            else
            {
                object result = ExcelAsyncUtil.Run("Rauname", new[] { request },
                    delegate
                    {
                        Task<IProduct> p = Task.Run(() => SkuAssist.GetProduct(request));
                        return p.Result;
                    });

                if (result == null)
                    return "Не найдено";

                if (result.Equals(ExcelError.ExcelErrorNA))
                    return "Загрузка...";

                IProduct product = result as IProduct;
                MemoryCache.Default.Add(request, product, System.DateTime.Now.AddMinutes(10));
                //MemoryCache.Default.Add(product.Name, product, System.DateTime.Now.AddMinutes(10));
                return product.Name;
            }
        }

        [ExcelFunction]
        public static object RAUSKU(string request)
        {
            if (MemoryCache.Default.Contains(request))
            {
                IProduct result = MemoryCache.Default[request] as IProduct;
                return result.Sku;
            }
            else
            {
                object result = ExcelAsyncUtil.Run("RauSku", new[] { request },
                 delegate
                 {
                     Task<IProduct> p = Task.Run(() => SkuAssist.GetProduct(request));
                     return p.Result;
                 });

                if (result == null)
                    return "Не найдено";

                if (result.Equals(ExcelError.ExcelErrorNA))
                    return "Загрузка...";

                IProduct product = result as IProduct;
                MemoryCache.Default.Add(request, product, System.DateTime.Now.AddMinutes(10));
                //MemoryCache.Default.Add(product.Sku, product, System.DateTime.Now.AddMinutes(10));
                return product.Sku;
            }
        }
    }
}
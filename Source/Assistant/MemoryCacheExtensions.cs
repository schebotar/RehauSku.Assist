using System;
using System.Runtime.Caching;
using System.Threading.Tasks;

namespace RehauSku.Assistant
{
    static class MemoryCacheExtensions
    {
        public static bool IsCached(this string request)
        {
            return MemoryCache.Default.Contains(request);
        }

        public static IProduct GetFromCache(this string request)
        {
            return MemoryCache.Default[request] as IProduct;
        }

        public static async Task<IProduct> RequestAndCache(this string request)
        {
            IProduct product = await SkuAssist.GetProductAsync(request);

            if (product == null)
                return null;

            MemoryCache.Default.Add(request, product, DateTime.Now.AddMinutes(10));
            return product;
        }
    }
}
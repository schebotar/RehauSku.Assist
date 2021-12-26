using System;
using System.Runtime.Caching;
using System.Threading.Tasks;
using RehauSku.Assistant;

namespace RehauSku
{
    static class MemoryCacheUtil
    {
        public static bool IsCached(this string request)
        {
            return AddIn.memoryCache.Contains(request);
        }

        public static IProduct GetFromCache(this string request)
        {
            return AddIn.memoryCache[request] as IProduct;
        }

        public static async Task<IProduct> RequestAndCache(this string request)
        {
            IProduct product = await SkuAssist.GetProductAsync(request);

            if (product == null)
                return null;

            AddIn.memoryCache.Add(request, product, DateTime.Now.AddMinutes(10));
            return product;
        }

        public static void ClearCache()
        {
            AddIn.memoryCache.Dispose();
            AddIn.memoryCache = new MemoryCache("RehauSku");
        }
    }
}
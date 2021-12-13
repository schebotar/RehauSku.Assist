using System.Threading.Tasks;

namespace RehauSku.Assistant
{
    public enum ProductField
    {
        Name,
        Id,
        Price
    }

    static class SkuAssist
    {
        public static async Task<IProduct> GetProductAsync(string request)
        {
            var content = await HttpClientUtil.GetContentByRequest(request);
            var document = await ParseUtil.ContentToDocAsync(content);

            return ParseUtil.GetProduct(document);
        }
    }
}
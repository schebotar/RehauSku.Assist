using AngleSharp.Dom;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    public enum ResponseOrder
    {
        NoSettings,
        Relevance,
        Name,
        Price,
        Series
    }

    static class SkuAssist
    {
        public static async Task<IProduct> GetProduct(string request)
        {
            Uri uri = request.ConvertToUri(ResponseOrder.NoSettings);

            Task<string> contentTask = Task.Run(() => HttpClientUtil.GetContentByUriAsync(uri));
            Task<IDocument> documentTask = await contentTask.ContinueWith(content => HttpClientUtil.ContentToDocAsync(content));

            return GetProduct(documentTask.Result);
        }

        public static IProduct GetProduct(IDocument d)
        {
            string script = d.Scripts
                   .Where(s => s.InnerHtml.Contains("dataLayer"))
                   .First()
                   .InnerHtml;

            string json = script
                .Substring(script.IndexOf("push(") + 5)
                .TrimEnd(new[] { ')', ';', '\n', ' ' });

            StoreResponce storeResponse = JsonConvert.DeserializeObject<StoreResponce>(json);
            IProduct product = storeResponse
                .Ecommerce
                .Impressions
                .Where(i => Regex.IsMatch(i.Id, @"\d{11}", RegexOptions.None))
                .FirstOrDefault();

            return product;
        }
    }
}
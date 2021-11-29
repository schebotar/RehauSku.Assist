using AngleSharp;
using AngleSharp.Dom;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    static class SkuAssist
    {
        public async static Task<string> GetContent(string request, HttpClient httpClient)
        {
            string uri = "https://shop-rehau.ru/catalogsearch/result/?q=" + request._CleanRequest();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            return await httpClient.GetStringAsync(uri);
        }

        public async static Task<IDocument> GetDocument(Task<string> source)
        {
            IConfiguration config = Configuration.Default;
            IBrowsingContext context = BrowsingContext.New(config);

            return await context.OpenAsync(req => req.Content(source.Result));
        }

        public static IProduct GetProductFromDocument(IDocument document)
        {
            return document
                .All
                .Where(e => e.ClassName == "product-item__desc-top")
                .Select(e => new Product(e.Children[0].TextContent, e.Children[1].TextContent.Trim(new[] { '\n', ' ' })))
                .FirstOrDefault();
        }

        private static string _CleanRequest(this string input)
        {
            return input.Replace("+", " plus ");
        }
    }
}



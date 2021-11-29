using AngleSharp;
using AngleSharp.Dom;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    static class SkuAssist
    {
        private static HttpClient _httpClient;
        private enum ResponseOrder
        {
            NoSettings,
            Relevance,
            Name,
            Price,
            Series
        }
        private static void _EnsureHttpClientRunning()
        {
            if (_httpClient == null)
                _httpClient = new HttpClient();
        }

        public async static Task<string> GetContent(string request)
        {
            Uri uri = _ConvertToUri(request, ResponseOrder.NoSettings);
            _EnsureHttpClientRunning();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            return await _httpClient.GetStringAsync(uri);
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

        private static Uri _ConvertToUri(this string request, ResponseOrder order)
        {
            string cleanedRequest = request._CleanRequest();
            switch (order)
            {
                case ResponseOrder.Relevance:
                    return new Uri("https://shop-rehau.ru/catalogsearch/result/index/?dir=asc&order=relevance&q=" + cleanedRequest);
                case ResponseOrder.Name:
                    return new Uri("https://shop-rehau.ru/catalogsearch/result/index/?dir=asc&order=name&q=" + cleanedRequest);
                case ResponseOrder.Price:
                    return new Uri("https://shop-rehau.ru/catalogsearch/result/index/?dir=asc&order=price&q=" + cleanedRequest);
                case ResponseOrder.Series:
                    return new Uri("https://shop-rehau.ru/catalogsearch/result/index/?dir=asc&order=sch_product_series&q=" + cleanedRequest);
                case ResponseOrder.NoSettings:
                    return new Uri("https://shop-rehau.ru/catalogsearch/result/?q=" + cleanedRequest);
                default:
                    throw new ArgumentException();
            }
        }
        private static string _CleanRequest(this string input)
        {
            return input.Replace("+", " plus ");
        }
    }
}



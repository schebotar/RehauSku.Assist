using AngleSharp;
using AngleSharp.Dom;
using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;

namespace Rehau.Sku.Assist
{
    static class HttpClientUtil
    {
        private static HttpClient _httpClient = AddIn.httpClient;

        public async static Task<string> GetContentByUriAsync(Uri uri)
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12 |
                SecurityProtocolType.Tls11 |
                SecurityProtocolType.Tls;

            return await _httpClient.GetStringAsync(uri);
        }

        public async static Task<IDocument> ContentToDocAsync(Task<string> content)
        {
            IConfiguration config = Configuration.Default;
            IBrowsingContext context = BrowsingContext.New(config);

            return await context.OpenAsync(req => req.Content(content.Result));
        }

        public static Uri ConvertToUri(this string request)
        {
            UriBuilder baseUri = new UriBuilder("https", "shop-rehau.ru");

            baseUri.Path = "/catalogsearch/result/index/";
            string cleanedRequest = request._CleanRequest();

            switch (AddIn.responseOrder)
            {
                case ResponseOrder.Relevance:
                    baseUri.Query = "dir=asc&order=relevance&q=" + cleanedRequest;
                    break;
                case ResponseOrder.Name:
                    baseUri.Query = "dir=asc&order=name&q=" + cleanedRequest;
                    break;
                case ResponseOrder.Price:
                    baseUri.Query = "dir=asc&order=price&q=" + cleanedRequest;
                    break;
                case ResponseOrder.Series:
                    baseUri.Query = "dir=asc&order=sch_product_series&q=" + cleanedRequest;
                    break;
                default:
                    baseUri.Query = "q=" + cleanedRequest;
                    break;
            }

            return baseUri.Uri;
        }

        private static string _CleanRequest(this string input)
        {
            return new StringBuilder(input)
                .Replace("+", " plus ")
                .Replace("РХ", "")
                .Replace("º", " ")
                .Replace(".", " ")
                .Replace("Ø", " ")
                .ToString();
        }
    }
}
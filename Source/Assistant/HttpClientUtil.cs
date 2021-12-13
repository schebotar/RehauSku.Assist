using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace RehauSku.Assistant
{
    static class HttpClientUtil
    {
        private static HttpClient _httpClient = AddIn.httpClient;

        public async static Task<string> GetContentByRequest(string request)
        {
            Uri uri = request.ConvertToUri();

            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12 |
                SecurityProtocolType.Tls11 |
                SecurityProtocolType.Tls;

            return await _httpClient.GetStringAsync(uri);
        }

        private static Uri ConvertToUri(this string request)
        {
            UriBuilder baseUri = new UriBuilder("https", "shop-rehau.ru");

            baseUri.Path = "/catalogsearch/result/index/";
            string cleanedRequest = request.CleanRequest();

            switch (AddIn.StoreResponse)
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
    }
}
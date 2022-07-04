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
            baseUri.Query = "q=" + cleanedRequest;

            return baseUri.Uri;
        }
    }
}
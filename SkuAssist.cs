using System.Net.Http;
using System.Threading.Tasks;
using AngleSharp;
using System.Linq;
using System.Net;

namespace Rehau.Sku.Assist
{
    static class SkuAssist
    {
        static private HttpClient _httpClient;

        public static void EnsureHttpInitialized()
        {
            if (_httpClient == null)
            {
                _httpClient = new HttpClient();
            }
        }

        public static string GetSku(string request)
        {
            string url = "https://shop-rehau.ru/catalogsearch/result/?q=" + request;
            HttpResponseMessage response = GetResponse(url).Result;
            var document = GetDocument(response).Result;

            var name = document
                .All
                .Where(e => e.ClassName == "product-item__desc-top")
                .Select(e => new { sku = e.Children[0].TextContent, name = e.Children[1].TextContent.Trim(new[] { '\n', ' ' }) })
                .Where(t => !t.sku.Any(c => char.IsLetter(c)))
                .FirstOrDefault();

            return name == null ? "Не найдено" : $"{name.name} ({name.sku})";
        }

        private static async Task<HttpResponseMessage> GetResponse(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            HttpResponseMessage response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();
            return response;
        }

        private static async Task<AngleSharp.Dom.IDocument> GetDocument(HttpResponseMessage response)
        {
            IConfiguration config = Configuration.Default;
            IBrowsingContext context = BrowsingContext.New(config);

            string source = await response.Content.ReadAsStringAsync();
            return await context.OpenAsync(req => req.Content(source));
        }
    }
}



using AngleSharp;
using AngleSharp.Dom;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RehauSku.Assistant
{
    static class ParseUtil
    {
        public async static Task<IDocument> ContentToDocAsync(string content)
        {
            IConfiguration config = Configuration.Default;
            IBrowsingContext context = BrowsingContext.New(config);

            return await context.OpenAsync(req => req.Content(content));
        }

        public static IProduct GetProduct(IDocument document)
        {
            try
            {
                string script = document
                    .Scripts
                    .Where(s => s.InnerHtml.Contains("dataLayer"))
                    .FirstOrDefault()
                    .InnerHtml;

                string json = script
                    .Substring(script.IndexOf("push(") + 5)
                    .TrimEnd(new[] { ')', ';', '\n', ' ' });

                if (!json.Contains("impressions"))
                    return null;

                StoreResponce storeResponse = JsonConvert.DeserializeObject<StoreResponce>(json);
                IProduct product = storeResponse
                    .Ecommerce
                    .Impressions
                    .Where(p => p.Id.IsRehauSku())
                    .FirstOrDefault();

                return product;
            }

            catch (NullReferenceException e)
            {
                MessageBox.Show(e.Message, "Ошибка получения данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
    }
}
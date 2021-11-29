using AngleSharp.Dom;
using ExcelDna.Integration;
using System.Net.Http;
using System.Threading.Tasks;

namespace Rehau.Sku.Assist
{
    public class Functions
    {
        private static HttpClient httpClient = new HttpClient();

        [ExcelFunction]
        public static async Task<string> RAUNAME(string request)
        {
            Task<string> contentTask = Task.Run(() => SkuAssist.GetContent(request, httpClient));
            Task<IDocument> documentTask = await contentTask.ContinueWith(content => SkuAssist.GetDocument(content));
            IProduct product = await documentTask.ContinueWith(doc => SkuAssist.GetProductFromDocument(doc.Result));
            return product == null ? ExcelError.ExcelErrorNull.ToString() : product.ToString();
        }
    }
}
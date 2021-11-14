using ExcelDna.Integration;
using System.Net.Http;

namespace Rehau.Sku.Assist
{
    public class Functions : IExcelAddIn
    {
        static readonly HttpClient httpClient = new HttpClient();

        public static object RAUNAME(string request)
        {
            return ExcelAsyncUtil.Run("RAUNAME", request, delegate
            {
                var document = SkuAssist.GetDocumentAsync(request, httpClient).Result;
                return SkuAssist.GetResultFromDocument(document);
            });
        }

        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
        }
    }
}
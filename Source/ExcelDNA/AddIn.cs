using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Net.Http;

namespace Rehau.Sku.Assist
{
    public class AddIn : IExcelAddIn
    {
        public static HttpClient httpClient;

        public void AutoOpen()
        {
            RegisterFunctions();
            httpClient = new HttpClient();
        }

        public void AutoClose()
        {
        }

        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                             .RegisterFunctions();
        }
    }
}

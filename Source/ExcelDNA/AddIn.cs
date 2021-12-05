using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Net.Http;

namespace Rehau.Sku.Assist
{
    public class AddIn : IExcelAddIn
    {
        public static readonly HttpClient httpClient = new HttpClient();

        public void AutoOpen()
        {
            RegisterFunctions();
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

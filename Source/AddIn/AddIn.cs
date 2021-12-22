using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using System.Net.Http;

namespace RehauSku
{
    public enum ResponseOrder
    {
        Default,
        Relevance,
        Name,
        Price,
        Series
    }

    public class AddIn : IExcelAddIn
    {
        public static HttpClient httpClient = new HttpClient();

        public void AutoOpen()
        {
            RegisterFunctions();
            IntelliSenseServer.Install();
            RegistryUtil.Initialize();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                             .RegisterFunctions();
        }
    }
}

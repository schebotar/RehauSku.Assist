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
        public static readonly HttpClient httpClient = new HttpClient();
        public static ResponseOrder StoreResponseOrder = RegistryUtil.StoreResponseOrder;
        public static string PriceListPath = RegistryUtil.PriceListPath;

        public void AutoOpen()
        {
            RegisterFunctions();
            IntelliSenseServer.Install();
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

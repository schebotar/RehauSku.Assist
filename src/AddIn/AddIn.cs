using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Runtime.Caching;


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
        public static HttpClient httpClient;
        public static MemoryCache memoryCache;
        public static Application Excel;

        public void AutoOpen()
        {
            httpClient = new HttpClient();
            memoryCache = new MemoryCache("RehauSku");
            RegisterFunctions();
            IntelliSenseServer.Install();
            RegistryUtil.Initialize();
            Excel = (Application)ExcelDnaUtil.Application;
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
            RegistryUtil.Uninitialize();
            memoryCache.Dispose();
        }

        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                             .RegisterFunctions();
        }
    }
}

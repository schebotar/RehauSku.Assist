using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Runtime.Caching;


namespace RehauSku
{
    enum ResponseOrder
    {
        Default,
        Relevance,
        Name,
        Price,
        Series
    }

    class AddIn : IExcelAddIn
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
            AddEvents();
        }

        private void AddEvents()
        {
            Excel.SheetSelectionChange += RefreshExportButton;
            Excel.SheetActivate += RefreshConvertButton;
            Excel.WorkbookActivate += RefreshConvertButton;
        }

        private void RefreshConvertButton(object sh)
        {
            Interface.RibbonController.RefreshControl("convertPrice");
        }

        private void RefreshExportButton(object sh, Range target)
        {
            Interface.RibbonController.RefreshControl("exportToPrice");
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

using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Runtime.Caching;

namespace RehauSku
{
    class AddIn : IExcelAddIn
    {
        public static Application Excel;

        public void AutoOpen()
        {
            Excel = (Application)ExcelDnaUtil.Application;
            RegisterFunctions();
            IntelliSenseServer.Install();
            RegistryUtil.Initialize();
            EventsUtil.Initialize();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
            RegistryUtil.Uninitialize();
            EventsUtil.Uninitialize();
        }

        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                             .RegisterFunctions();
        }
    }
}

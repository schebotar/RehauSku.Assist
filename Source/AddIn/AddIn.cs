using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Win32;
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
        public static ResponseOrder StoreResponse { get; set; }
        public static string PriceListPath { get; set; }

        public void AutoOpen()
        {
            RegisterFunctions();
            GetRegistryKeys();
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

        void GetRegistryKeys()
        {
            RegistryKey addInKeys = Registry
                .CurrentUser
                .OpenSubKey("SOFTWARE")
                .OpenSubKey("REHAU")
                .OpenSubKey("SkuAssist");

            StoreResponse = (ResponseOrder)addInKeys.GetValue("ResponseOrder");
            PriceListPath = (string)addInKeys.GetValue("PriceListPath");
        }
    }
}

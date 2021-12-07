using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Win32;
using System.Net.Http;

namespace Rehau.Sku.Assist
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
        public static ResponseOrder responseOrder;
        public static string priceListPath;

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

            responseOrder = (ResponseOrder)addInKeys.GetValue("ResponseOrder");
            priceListPath = (string)addInKeys.GetValue("PriceListPath");
        }

    }
}

using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Net.Http;

namespace Rehau.Sku.Assist
{
    public enum ResponseOrder
    {
        NoSettings,
        Relevance,
        Name,
        Price,
        Series
    }

    public class AddIn : IExcelAddIn
    {
        public static readonly HttpClient httpClient = new HttpClient();
        public static ResponseOrder responseOrder;

        public void AutoOpen()
        {
            RegisterFunctions();
            responseOrder = ResponseOrder.NoSettings;
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

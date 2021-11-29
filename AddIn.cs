using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Rehau.Sku.Assist
{
    public class AddIn : IExcelAddIn
    {
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

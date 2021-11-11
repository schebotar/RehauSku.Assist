using ExcelDna.Integration;

namespace Rehau.Sku.Assist
{
    public class Functions : IExcelAddIn
    {
        [ExcelFunction(description: "Получение наименования и артикула позиции")]
        public static string RAUNAME(string request)
        {
            SkuAssist.EnsureHttpInitialized();
            return SkuAssist.GetSku(request);
        }

        public void AutoClose()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                delegate (object ex) { return string.Format("!!!ERROR: {0}", ex.ToString()); });
        }

        public void AutoOpen()
        {
        }
    }
}
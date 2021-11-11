using ExcelDna.Integration;

namespace Rehau.Sku.Assist
{
    public class Functions : IExcelAddIn
    {
        [ExcelFunction(description: "Получение наименования и артикула позиции")]
        public static object RAUNAME(string request)
        {
            SkuAssist.EnsureHttpInitialized();

            return ExcelTaskUtil.Run("RAUNAME ASYNC", request, async token =>
            {
                var document = await SkuAssist.GetDocumentAsync(request);
                return SkuAssist.GetResultFromDocument(document);
            });
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
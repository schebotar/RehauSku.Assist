using Microsoft.Office.Interop.Excel;

namespace RehauSku
{
    internal static class EventsUtil
    {
        private static readonly Application Excel = AddIn.Excel;

        public static void Initialize()
        {
            Excel.SheetSelectionChange += RefreshExportButton;
            Excel.SheetActivate += RefreshConvertButton;
            Excel.WorkbookActivate += RefreshConvertButton;
        }

        public static void Uninitialize()
        {
            Excel.SheetSelectionChange -= RefreshExportButton;
            Excel.SheetActivate -= RefreshConvertButton;
            Excel.WorkbookActivate -= RefreshConvertButton;
        }

        private static void RefreshConvertButton(object sh)
        {
            Interface.RibbonController.RefreshControl("convert");
        }

        private static void RefreshExportButton(object sh, Range target)
        {
            Interface.RibbonController.RefreshControl("export");
        }
    }
}

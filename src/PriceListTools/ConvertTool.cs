namespace RehauSku.PriceListTools
{
    internal class ConvertTool : PriceListTool
    {
        private Source Current;

        public void GetCurrent()
        {
            Current = new Source(ExcelApp.ActiveWorkbook);
        }

        public void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;            
            FillColumn(Current.SkuAmount, TargetFile.amountCell.Column);
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}
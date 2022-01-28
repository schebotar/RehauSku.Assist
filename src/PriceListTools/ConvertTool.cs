using System;

namespace RehauSku.PriceListTools
{
    internal class ConvertTool : PriceListTool
    {
        private Source Current;

        public void GetCurrent()
        {
            try
            {
                Current = new Source(ExcelApp.ActiveWorkbook);
            }

            catch (Exception exception)
            {
                System.Windows.Forms.MessageBox.Show
                    (exception.Message,
                    "Ошибка распознавания",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                throw exception;
            }
        }

        public void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;            
            FillColumn(Current.PositionAmount, TargetFile.amountCell.Column);
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}
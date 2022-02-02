using RehauSku.Interface;
using System;

namespace RehauSku.PriceListTools
{
    internal class ConvertTool : AbstractTool
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
            ProgressBar bar = new ProgressBar(Current.PositionAmount.Count);

            foreach (var kvp in Current.PositionAmount)
            {
                FillColumnsWithDictionary(kvp, TargetFile.amountCell.Column);
                bar.DoProgress();
            }

            FilterByAmount();

            Dialog.SaveWorkbookAs();
        }
    }
}
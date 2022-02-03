using RehauSku.Interface;
using System;

namespace RehauSku.PriceListTools
{
    internal class ConvertTool : AbstractTool
    {
        private SourcePriceList Current;

        public void GetCurrent()
        {
            try
            {
                Current = new SourcePriceList(ExcelApp.ActiveWorkbook);
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
            ProgressBar bar = new ProgressBar("Заполняю строки...", Current.PositionAmount.Count);

            foreach (var kvp in Current.PositionAmount)
            {
                FillPositionAmountToColumns(kvp, TargetFile.amountCell.Column);
                bar.DoProgress();
            }

            FilterByAmount();

            Dialog.SaveWorkbookAs();
        }
    }
}
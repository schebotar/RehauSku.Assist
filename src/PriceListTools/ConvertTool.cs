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
            ProgressBar = new ProgressBar("Заполняю строки...", Current.PositionAmount.Count);
            ResultBar = new ResultBar();

            foreach (var kvp in Current.PositionAmount)
            {
                FillPositionAmountToColumns(kvp, TargetFile.amountCell.Column);
                ProgressBar.Update();
            }

            FilterByAmount();
            ResultBar.Update();

            //Dialog.SaveWorkbookAs();
            ExcelApp.StatusBar = false;
        }
    }
}
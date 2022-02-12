using RehauSku.Interface;
using System;

namespace RehauSku.PriceListTools
{
    internal class ConvertTool : AbstractTool
    {
        private SourcePriceList Current { get; set; }

        public ConvertTool()
        {
            Current = new SourcePriceList(ExcelApp.ActiveWorkbook);
        }

        public override void FillTarget()
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

            Dialog.SaveWorkbookAs();
            ExcelApp.StatusBar = false;
        }
    }
}
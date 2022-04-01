﻿using RehauSku.Interface;

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
                FillPositionAmountToColumns(kvp, TargetFile.AmountCell.Column);
                ProgressBar.Update();
            }

            FilterByAmount();
            ResultBar.Update();
        }
    }
}
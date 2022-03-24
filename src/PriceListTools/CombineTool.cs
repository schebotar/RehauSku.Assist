using Microsoft.Office.Interop.Excel;
using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Dialog = RehauSku.Interface.Dialog;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : AbstractTool
    {
        private List<SourcePriceList> SourceFiles { get; set; }

        public CombineTool()
        {
            string[] files = Dialog.GetMultiplyFiles();

            if (files != null)
            {
                SourceFiles = SourcePriceList.GetSourceLists(files);
            }

            else
            {
                throw new Exception("Не выбраны файлы");
            }
        }

        public override async void FillTarget()
        {
            ProgressBar = new ProgressBar("Заполняю строки...", SourceFiles.Sum(file => file.PositionAmount.Count));
            ResultBar = new ResultBar();

            foreach (SourcePriceList source in SourceFiles)
            {
                TargetFile.Sheet.Columns[TargetFile.AmountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                Range newColumnHeader = TargetFile.Sheet.Cells[TargetFile.AmountCell.Row, TargetFile.AmountCell.Column - 1];
                newColumnHeader.Value2 = $"{source.Name}";
                newColumnHeader.WrapText = true;

                foreach (var kvp in source.PositionAmount)
                {
                    FillPositionAmountToColumns(kvp, TargetFile.AmountCell.Column - 1, TargetFile.AmountCell.Column);
                    ProgressBar.Update();
                }
            }

            FilterByAmount();
            ResultBar.Update();

            await Task.Delay(new TimeSpan(0, 0, 5));
            ExcelApp.StatusBar = false;
        }
    }
}

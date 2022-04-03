using Microsoft.Office.Interop.Excel;
using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
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

        public override void FillTarget()
        {
            using (ProgressBar = new ProgressBar("Заполняю строки...", SourceFiles.Sum(file => file.PositionAmount.Count)))
            using (ResultBar = new ResultBar())
            {
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
            }
        }
    }
}

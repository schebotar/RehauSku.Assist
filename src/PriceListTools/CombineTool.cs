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
            ProgressBar = new ProgressBar("Заполняю строки...", SourceFiles.Sum(file => file.PositionAmount.Count));
            ResultBar = new ResultBar();

            foreach (SourcePriceList source in SourceFiles)
            {
                TargetFile.Sheet.Columns[TargetFile.amountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                Range newColumnHeader = TargetFile.Sheet.Cells[TargetFile.amountCell.Row, TargetFile.amountCell.Column - 1];
                newColumnHeader.Value2 = $"{source.Name}";
                newColumnHeader.WrapText = true;

                foreach (var kvp in source.PositionAmount)
                {
                    FillPositionAmountToColumns(kvp, TargetFile.amountCell.Column - 1, TargetFile.amountCell.Column);
                    ProgressBar.Update();
                }
            }

            FilterByAmount();
            ResultBar.Update();

            Interface.Dialog.SaveWorkbookAs();
            ExcelApp.StatusBar = false;
        }
    }
}

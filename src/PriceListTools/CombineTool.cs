using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using RehauSku.Interface;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : AbstractTool
    {
        public List<SourcePriceList> SourceFiles;

        public void FillTarget()
        {
            ProgressBar bar = new ProgressBar(SourceFiles.Sum(file => file.PositionAmount.Count));

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
                    bar.DoProgress();
                }
            }

            FilterByAmount();

            Interface.Dialog.SaveWorkbookAs();
        }
    }
}

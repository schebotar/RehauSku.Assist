using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : PriceListTool
    {
        public List<Source> SourceFiles;

        public void FillTarget()
        {
            foreach (Source source in SourceFiles)
            {
                TargetFile.Sheet.Columns[TargetFile.amountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                Range newColumnHeader = TargetFile.Sheet.Cells[TargetFile.amountCell.Row, TargetFile.amountCell.Column - 1];
                newColumnHeader.Value2 = $"{source.Name}";
                newColumnHeader.WrapText = true;

                foreach(var kvp in source.PositionAmount)
                    FillColumnsWithDictionary(kvp, TargetFile.amountCell.Column - 1, TargetFile.amountCell.Column);
            }

            FilterByAmount();

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}

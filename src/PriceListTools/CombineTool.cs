using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : PriceListTool
    {
        public List<Source> SourceFiles;

        public void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;

            foreach (Source source in SourceFiles)
            {
                TargetFile.Sheet.Columns[TargetFile.amountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                TargetFile.Sheet.Cells[TargetFile.amountCell.Row, TargetFile.amountCell.Column - 1].Value2 = $"{source.Name}";

                FillColumn(source.PositionAmount, TargetFile.amountCell.Column - 1);
                FillColumn(source.PositionAmount, TargetFile.amountCell.Column);
            }

            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}

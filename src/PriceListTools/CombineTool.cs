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
            FillAmountColumn(SourceFiles.Select(x => x.SkuAmount).ToArray());
            AddAndFillSourceColumns();
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }

        private void AddAndFillSourceColumns()
        {
            foreach (var source in SourceFiles)
            {
                if (source.SkuAmount.Count == 0)
                    continue;

                TargetFile.Sheet.Columns[TargetFile.amountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                TargetFile.Sheet.Cells[TargetFile.amountCell.Row, TargetFile.amountCell.Column - 1].Value2 = $"{source.Name}";

                foreach (var kvp in source.SkuAmount)
                {
                    Range cell = TargetFile.Sheet.Columns[TargetFile.skuCell.Column].Find(kvp.Key);

                    if (cell == null)
                    {
                        continue;
                    }

                    else
                    {
                        TargetFile.Sheet.Cells[cell.Row, TargetFile.amountCell.Column - 1].Value2 = kvp.Value;
                    }
                }
            }
        }
    }
}

using Microsoft.Office.Interop.Excel;
using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using ProgressBar = RehauSku.Interface.ProgressBar;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractTool
    {
        protected Application ExcelApp = AddIn.Excel;
        protected TargetPriceList TargetFile { get; set; }
        protected ResultBar ResultBar { get; set; }
        protected ProgressBar ProgressBar { get; set; }

        public abstract void FillTarget();

        public void OpenNewPrice()
        {
            if (ExcelApp.Workbooks
                .Cast<Workbook>()
                .FirstOrDefault(w => w.FullName == RegistryUtil.PriceListPath) != null)
            {
                throw new ArgumentException("Шаблонный файл редактируется в другом месте");
            }

            Workbook wb = ExcelApp.Workbooks.Open(RegistryUtil.PriceListPath, null, true);

            try
            {
                TargetFile = new TargetPriceList(wb);
            }

            catch (Exception exception)
            {
                if (wb != null)
                {
                    wb.Close();
                }

                throw exception;
            }
        }

        protected void FillPositionAmountToColumns(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            Range worksheetCells = TargetFile.Sheet.Cells;
            Range skuColumn = TargetFile.SkuCell.EntireColumn;
            Range oldSkuColumn = TargetFile.OldSkuCell.EntireColumn;

            int? row = GetPositionRow(skuColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = worksheetCells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementSuccess();
                return;
            }

            if (TargetFile.OldSkuCell != null)
            {
                row = GetPositionRow(oldSkuColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

                if (row != null)
                {
                    foreach (int column in columns)
                    {
                        Range cell = worksheetCells[row, column];
                        cell.AddValue(positionAmount.Value);
                    }

                    ResultBar.IncrementReplaced();
                    return;
                }
            }

            string sku = positionAmount.Key.Sku.Substring(1, 6);
            row = GetPositionRow(skuColumn, sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = worksheetCells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementReplaced();
                return;
            }

            FillMissing(positionAmount, columns);
            ResultBar.IncrementNotFound();
        }

        protected void FillMissing(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            Range worksheetCells = TargetFile.Sheet.Cells;
            Range worksheetRows = TargetFile.Sheet.Rows;
            int skuColumn = TargetFile.SkuCell.Column;
            int groupColumn = TargetFile.GroupCell.Column;
            int nameColumn = TargetFile.NameCell.Column;

            int row = worksheetCells[worksheetRows.Count, skuColumn]
                .End[XlDirection.xlUp]
                .Row + 1;

            worksheetRows[row]
                .EntireRow
                .Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            Range previous = worksheetRows[row - 1];
            Range current = worksheetRows[row];

            previous.Copy(current);
            current.ClearContents();

            worksheetCells[row, groupColumn].Value2 = positionAmount.Key.Group;
            worksheetCells[row, nameColumn].Value2 = positionAmount.Key.Name;

            if (TargetFile.OldSkuCell != null)
            {
                worksheetCells[row, skuColumn].Value2 = "Не найден";
                worksheetCells[row, TargetFile.OldSkuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            else
            {
                worksheetCells[row, skuColumn].Value2 = positionAmount.Key.Sku;
            }

            foreach (int column in columns)
            {
                Range cell = worksheetCells[row, column];
                cell.AddValue(positionAmount.Value);
            }
        }

        protected int? GetPositionRow(Range range, string sku, string group)
        {
            Range found = range.Find(sku);
            string foundGroupValue;

            if (found == null)
            {
                return null;
            }

            int firstFoundRow = found.Row;

            if (string.IsNullOrEmpty(group))
            {
                return found.Row;
            }

            while (true)
            {
                foundGroupValue = TargetFile.Sheet.Cells[found.Row, TargetFile.GroupCell.Column].Value2.ToString();

                if (group.Equals(foundGroupValue))
                {
                    return found.Row;
                }

                found = range.FindNext(found);

                if (found.Row == firstFoundRow)
                {
                    return null;
                }
            }
        }

        protected void FilterByAmount()
        {
            AutoFilter filter = TargetFile.Sheet.AutoFilter;
            int startColumn = filter.Range.Column;

            filter.Range.AutoFilter(TargetFile.AmountCell.Column - startColumn + 1, "<>");
            TargetFile.Sheet.Range["A1"].Activate();
        }
    }
}
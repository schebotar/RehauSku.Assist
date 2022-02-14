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
            int? row = GetPositionRow(TargetFile.SkuCell.EntireColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = TargetFile.Sheet.Cells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementSuccess();
                return;
            }

            if (TargetFile.OldSkuCell != null)
            {
                row = GetPositionRow(TargetFile.OldSkuCell.EntireColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

                if (row != null)
                {
                    foreach (int column in columns)
                    {
                        Range cell = TargetFile.Sheet.Cells[row, column];
                        cell.AddValue(positionAmount.Value);
                    }

                    ResultBar.IncrementReplaced();
                    return;
                }
            }

            string sku = positionAmount.Key.Sku.Substring(1, 6);
            row = GetPositionRow(TargetFile.SkuCell.EntireColumn, sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = TargetFile.Sheet.Cells[row, column];
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
            int row = TargetFile.Sheet.Cells[TargetFile.Sheet.Rows.Count, TargetFile.SkuCell.Column]
                .End[XlDirection.xlUp]
                .Row + 1;

            TargetFile.Sheet.Rows[row]
                .EntireRow
                .Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            Range previous = TargetFile.Sheet.Rows[row - 1];
            Range current = TargetFile.Sheet.Rows[row];

            previous.Copy(current);
            current.ClearContents();

            TargetFile.Sheet.Cells[row, TargetFile.GroupCell.Column].Value2 = positionAmount.Key.Group;
            TargetFile.Sheet.Cells[row, TargetFile.NameCell.Column].Value2 = positionAmount.Key.Name;

            if (TargetFile.OldSkuCell != null)
            {
                TargetFile.Sheet.Cells[row, TargetFile.SkuCell.Column].Value2 = "Не найден";
                TargetFile.Sheet.Cells[row, TargetFile.OldSkuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            else
            {
                TargetFile.Sheet.Cells[row, TargetFile.SkuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            foreach (int column in columns)
            {
                Range cell = TargetFile.Sheet.Cells[row, column];
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

            while (true)
            {
                foundGroupValue = TargetFile.Sheet.Cells[found.Row, TargetFile.GroupCell.Column].Value2.ToString();

                if (string.IsNullOrEmpty(group) || group.Equals(foundGroupValue))
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
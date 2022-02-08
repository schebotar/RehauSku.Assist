using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using ProgressBar = RehauSku.Interface.ProgressBar;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractTool
    {
        protected private Application ExcelApp = (Application)ExcelDnaUtil.Application;
        protected private TargetPriceList TargetFile;
        protected private ResultBar ResultBar { get; set; }
        protected private ProgressBar ProgressBar { get; set; }

        public void OpenNewPrice()
        {
            Workbook wb = ExcelApp.Workbooks.Open(RegistryUtil.PriceListPath, null, true);

            try
            {
                TargetFile = new TargetPriceList(wb);
            }

            catch (Exception ex)
            {
                MessageBox.Show
                    (ex.Message,
                    "Ошибка открытия шаблонного прайс-листа",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                wb.Close();
                throw ex;
            }
        }

        protected private void FillPositionAmountToColumns(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            int? row = GetPositionRow(positionAmount.Key.Sku, positionAmount.Key.Group, TargetFile.skuCell.Column);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = TargetFile.Sheet.Cells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementSuccess();
            }

            else if (TargetFile.oldSkuCell != null)
            {
                row = GetPositionRow(positionAmount.Key.Sku, positionAmount.Key.Group, TargetFile.oldSkuCell.Column);

                if (row != null)
                {
                    foreach (int column in columns)
                    {
                        Range cell = TargetFile.Sheet.Cells[row, column];
                        cell.AddValue(positionAmount.Value);
                    }

                    ResultBar.IncrementReplaced();
                }
            }

            else
            {
                string sku = positionAmount.Key.Sku.Substring(1, 6);
                row = GetPositionRow(sku, positionAmount.Key.Group, TargetFile.skuCell.Column);

                if (row != null)
                {
                    foreach (int column in columns)
                    {
                        Range cell = TargetFile.Sheet.Cells[row, column];
                        cell.AddValue(positionAmount.Value);
                    }

                    ResultBar.IncrementReplaced();
                }

                else
                {
                    FillMissing(positionAmount, columns);
                    ResultBar.IncrementNotFound();
                }
            }
        }

        protected private void FillMissing(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            int row = TargetFile.Sheet.Cells[TargetFile.Sheet.Rows.Count, TargetFile.skuCell.Column]
                .End[XlDirection.xlUp]
                .Row + 1;

            TargetFile.Sheet.Rows[row]
                .EntireRow
                .Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            Range previous = TargetFile.Sheet.Rows[row - 1];
            Range current = TargetFile.Sheet.Rows[row];

            previous.Copy(current);
            current.ClearContents();

            TargetFile.Sheet.Cells[row, TargetFile.groupCell.Column].Value2 = positionAmount.Key.Group;
            TargetFile.Sheet.Cells[row, TargetFile.nameCell.Column].Value2 = positionAmount.Key.Name;

            if (TargetFile.oldSkuCell != null)
            {
                TargetFile.Sheet.Cells[row, TargetFile.skuCell.Column].Value2 = "Не найден";
                TargetFile.Sheet.Cells[row, TargetFile.oldSkuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            else
            {
                TargetFile.Sheet.Cells[row, TargetFile.skuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            foreach (int column in columns)
            {
                Range cell = TargetFile.Sheet.Cells[row, column];
                cell.AddValue(positionAmount.Value);
            }
        }

        protected private int? GetPositionRow(string sku, string group, int column)
        {
            int? row = null;
            Range foundCell = TargetFile.Sheet.Columns[column].Find(sku);
            string foundGroupValue;

            if (foundCell == null) return null;

            else
            {
                row = foundCell.Row;
                foundGroupValue = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();
            }

            if (string.IsNullOrEmpty(group) || group.Equals(foundGroupValue))
                return row;

            else
                while (true)
                {
                    foundCell = TargetFile.skuCell.EntireColumn.FindNext(foundCell);
                    if (foundCell == null) return row;

                    foundGroupValue = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();
                    if (group.Equals(foundGroupValue)) return foundCell.Row;
                }
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = TargetFile.Sheet.AutoFilter;
            int startColumn = filter.Range.Column;

            filter.Range.AutoFilter(TargetFile.amountCell.Column - startColumn + 1, "<>");
            TargetFile.Sheet.Range["A1"].Activate();
        }
    }
}
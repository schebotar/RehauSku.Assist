using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace RehauSku.PriceListTools
{
    internal abstract class PriceListTool
    {
        protected private Application ExcelApp = (Application)ExcelDnaUtil.Application;
        protected private Target TargetFile;

        public void OpenNewPrice()
        {
            Workbook wb = ExcelApp.Workbooks.Open(RegistryUtil.PriceListPath);

            try
            {
                TargetFile = new Target(wb);
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

        protected private void FillColumnsWithDictionary(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            int? row = GetPositionRow(positionAmount.Key.Sku, positionAmount.Key.Group, TargetFile.skuCell.Column);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range sumCell = TargetFile.Sheet.Cells[row, column];

                    if (sumCell.Value2 == null)
                    {
                        sumCell.Value2 = positionAmount.Value;
                    }

                    else
                    {
                        sumCell.Value2 += positionAmount.Value;
                    }
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
                        Range amountCell = TargetFile.Sheet.Cells[row, column];

                        if (amountCell.Value2 == null)
                        {
                            amountCell.Value2 = positionAmount.Value;
                        }

                        else
                        {
                            amountCell.Value2 += positionAmount.Value;
                        }

                        Range oldSkuCell = TargetFile.Sheet.Cells[row, TargetFile.oldSkuCell.Column];
                        oldSkuCell.Value2 = positionAmount.Key.Sku;
                    }
                }

                else
                {
                    FillMissing(positionAmount, columns);
                }
            }
        }

        protected private void FillMissing(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            Range foundCell = TargetFile.oldSkuCell.EntireColumn.Find(positionAmount.Key.Sku);
            int row;

            if (foundCell == null)
            {
                row = TargetFile.Sheet.Cells[TargetFile.Sheet.Rows.Count, TargetFile.skuCell.Column]
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
                TargetFile.Sheet.Cells[row, TargetFile.oldSkuCell.Column].Value2 = positionAmount.Key.Sku;
                TargetFile.Sheet.Cells[row, TargetFile.nameCell.Column].Value2 = positionAmount.Key.Name;
                TargetFile.Sheet.Cells[row, TargetFile.skuCell.Column].Value2 = "Не найден";
            }

            else
            {
                row = foundCell.Row;
            }

            foreach (int column in columns)
            {
                if (TargetFile.Sheet.Cells[row, column].Value2 == null)
                {
                    TargetFile.Sheet.Cells[row, column].Value2 = positionAmount.Value;
                }

                else
                {
                    TargetFile.Sheet.Cells[row, column].Value2 += positionAmount.Value;
                }
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
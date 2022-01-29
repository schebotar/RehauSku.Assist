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
        protected private List<KeyValuePair<Position, double>> Missing;

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

        protected private void FillColumnsWithDictionary(IEnumerable<KeyValuePair<Position, double>> dictionary, params int[] columns)
        {
            Missing = new List<KeyValuePair<Position, double>>();

            foreach (var positionAmount in dictionary)
            {
                FillPositionAmountToColumns(positionAmount, columns);
            }

            if (Missing.Count > 0)
            { 
                DialogResult result = 
                MessageBox.Show
                    ($"{Missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath} Попробовать найти новый вариант?",
                    "Отсутствует позиция в конечной таблице заказов",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    var lookUp = new List<KeyValuePair<Position, double>>(Missing);

                    foreach (var missingPosition in lookUp)
                    {
                        TryFillVariantlessSkuToColumns(missingPosition, columns);
                    }
                }
            }

            if (Missing.Count > 0)
            {
                FillMissing(columns);
                MessageBox.Show
                    ($"{Missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath}\n" +
                    $"Под основной таблицей составлен список не найденных артикулов",
                    "Отсутствует позиция в конечной таблице заказов",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        protected private void FillPositionAmountToColumns(KeyValuePair<Position, double> positionAmount, int[] columns)
        {
            Range foundCell = TargetFile.skuCell.EntireColumn.Find(positionAmount.Key.Sku);

            if (foundCell == null)
            {
                Missing.Add(positionAmount);
                return;
            }

            string foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();

            while (foundCell != null && foundCellGroup != positionAmount.Key.Group)
            {
                foundCell = TargetFile.skuCell.EntireColumn.FindNext(foundCell);
                foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();
            }

            if (foundCell == null)
            {
                Missing.Add(positionAmount);
            }

            else
            {
                foreach (var column in columns)
                {
                    Range sumCell = TargetFile.Sheet.Cells[foundCell.Row, column];

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
        }

        protected private void TryFillVariantlessSkuToColumns(KeyValuePair<Position, double> positionAmount, int[] columns)
        {
            string sku = positionAmount.Key.Sku.Substring(1, 6);

            Range foundCell = TargetFile.skuCell.EntireColumn.Find(sku);

            if (foundCell == null)
            {
                return;
            }

            string foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();

            while (foundCell != null && foundCellGroup != positionAmount.Key.Group)
            {
                foundCell = TargetFile.skuCell.EntireColumn.FindNext(foundCell);
                foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();
            }

            if (foundCell == null)
            {
                return;
            }

            foreach (var column in columns)
            {
                Range sumCell = TargetFile.Sheet.Cells[foundCell.Row, column];

                if (sumCell.Value2 == null)
                {
                    sumCell.Value2 = positionAmount.Value;
                }

                else
                {
                    sumCell.Value2 += positionAmount.Value;
                }
            }

            Missing.Remove(positionAmount);
        }

        protected private void FillMissing(int[] columns)
        {
            int startRow =
                TargetFile.Sheet.AutoFilter.Range.Row + 
                TargetFile.Sheet.AutoFilter.Range.Rows.Count + 5;

            for (int i = 0; i < Missing.Count; i++)
            {
                Range group = TargetFile.Sheet.Cells[startRow + i, TargetFile.groupCell.Column];
                Range sku = TargetFile.Sheet.Cells[startRow + i, TargetFile.skuCell.Column];
                Range name = TargetFile.Sheet.Cells[startRow + i, TargetFile.nameCell.Column];

                group.Value2 = Missing[i].Key.Group;
                sku.Value2 = Missing[i].Key.Sku;
                name.Value2 = Missing[i].Key.Name;

                group.ClearFormats();
                sku.ClearFormats();
                name.ClearFormats();

                foreach (int column in columns)
                {
                    Range amount = TargetFile.Sheet.Cells[startRow + i, column];
                    amount.Value2 = Missing[i].Value;
                    amount.ClearFormats();
                }
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
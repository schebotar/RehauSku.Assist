using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

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
                System.Windows.Forms.MessageBox.Show
                    (ex.Message,
                    "Ошибка открытия шаблонного прайс-листа",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                wb.Close();
                throw ex;
            }
        }

        protected private void FillColumn(IEnumerable<KeyValuePair<Position, double>> dictionary, params int[] columns)
        {
            Missing = new List<KeyValuePair<Position, double>>();
            object[,] groupColumn = TargetFile.groupCell.EntireColumn.Value2;

            foreach (var kvp in dictionary)
            {
                FillPosition(kvp, columns);
            }

            if (Missing.Count > 0)
            {
                System.Windows.Forms.MessageBox.Show
                    ($"{Missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath} Попробовать найти новый вариант?",
                    "Отсутствует позиция в конечной таблице заказов",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
        }

        protected private void FillPosition(KeyValuePair<Position, double> kvp, int[] columns)
        {
            Range foundCell = TargetFile.skuCell.EntireColumn.Find(kvp.Key.Sku);
            if (foundCell == null)
            {
                Missing.Add(kvp);
                return;
            }

            string foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();

            while (foundCell != null && foundCellGroup != kvp.Key.Group)
            {
                foundCell = TargetFile.skuCell.EntireColumn.FindNext(foundCell);
                foundCellGroup = TargetFile.Sheet.Cells[foundCell.Row, TargetFile.groupCell.Column].Value2.ToString();
            }

            if (foundCell == null)
            {
                Missing.Add(kvp);
            }

            else
            {
                foreach (var column in columns)
                {
                    Range sumCell = TargetFile.Sheet.Cells[foundCell.Row, column];
                    if (sumCell.Value2 == null)
                    {
                        sumCell.Value2 = kvp.Value;
                    }

                    else
                    {
                        sumCell.Value2 += kvp.Value;
                    }
                }
            }
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = TargetFile.Sheet.AutoFilter;

            filter.Range.AutoFilter(TargetFile.amountCell.Column, "<>");
            TargetFile.Sheet.Range["A1"].Activate();
        }
    }
}
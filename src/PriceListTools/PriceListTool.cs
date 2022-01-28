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

        protected private void FillColumn(IEnumerable<KeyValuePair<Position, double>> dictionary, int column)
        {
            List<KeyValuePair<Position, double>> missing = new List<KeyValuePair<Position, double>>();
            object[,] groupColumn = TargetFile.groupCell.EntireColumn.Value2;

            foreach (var kvp in dictionary)
            {
                Range foundCell = TargetFile.skuCell.EntireColumn.Find(kvp.Key.Sku);
                string foundCellGroup = groupColumn[foundCell.Row, 1].ToString();

                while (foundCell != null && foundCellGroup != kvp.Key.Group)
                {
                    foundCell = TargetFile.skuCell.EntireColumn.FindNext(foundCell);
                    foundCellGroup = groupColumn[foundCell.Row, 1].ToString();
                }

                if (foundCell == null)
                {
                    missing.Add(kvp);
                }

                else
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

            if (missing.Count > 0)
            {
                System.Windows.Forms.MessageBox.Show
                    ($"{missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath} Попробовать найти новый вариант?",
                    "Отсутствует позиция в конечной таблице заказов",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Information);
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
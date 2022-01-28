using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

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

        protected private void FillColumn(Dictionary<string, double> dictionary, int column)
        {
            List<KeyValuePair<string, double>> missing = new List<KeyValuePair<string, double>>();

            foreach (var kvp in dictionary)
            {
                Range cell = TargetFile.Sheet.Columns[TargetFile.skuCell.Column].Find(kvp.Key);

                if (cell == null)
                {
                    missing.Add(kvp);
                }

                else
                {
                    Range sumCell = TargetFile.Sheet.Cells[cell.Row, column];

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

            string values = string.Join("\n", missing.Select(kvp => kvp.Key).ToArray());
            System.Windows.Forms.MessageBox.Show
                ($"{missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath}",
                "Отсутствует позиция в конечной таблице заказов",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = TargetFile.Sheet.AutoFilter;

            filter.Range.AutoFilter(TargetFile.amountCell.Column, "<>");
            TargetFile.Sheet.Range["A1"].Activate();
        }
    }
}
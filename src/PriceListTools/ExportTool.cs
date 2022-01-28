using Microsoft.Office.Interop.Excel;
using RehauSku.Assistant;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class ExportTool : PriceListTool
    {
        private Dictionary<string, double> SkuAmount { get; set; }
        private Range Selection;

        public void TryGetSelection()
        {
            Selection = ExcelApp.Selection;

            if (Selection == null || Selection.Columns.Count != 2)
            {
                throw new Exception("Неверный диапазон");
            }
        }

        public void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;
            GetSelected();
            FillColumn(SkuAmount, TargetFile.amountCell.Column);
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }

        private void FillColumn(IEnumerable<KeyValuePair<string, double>> dictionary, int column)
        {
            List<KeyValuePair<string, double>> missing = new List<KeyValuePair<string, double>>();

            foreach (var kvp in dictionary)
            {
                Range cell = TargetFile.skuCell.EntireColumn.Find(kvp.Key);

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

            if (missing.Count > 0)
            {
                System.Windows.Forms.MessageBox.Show
                    ($"{missing.Count} артикулов отсутствует в таблице заказов {RegistryUtil.PriceListPath} Попробовать найти новый вариант?",
                    "Отсутствует позиция в конечной таблице заказов",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
        }


        private void GetSelected()
        {
            object[,] cells = Selection.Value2;
            SkuAmount = new Dictionary<string, double>();
            int rowsCount = Selection.Rows.Count;

            for (int row = 1; row <= rowsCount; row++)
            {
                if (cells[row, 1] == null || cells[row, 2] == null)
                    continue;

                string sku = null;
                double? amount = null;

                for (int column = 1; column <= 2; column++)
                {
                    object current = cells[row, column];

                    if (current.ToString().IsRehauSku())
                    {
                        sku = current.ToString();
                    }

                    else if (current.GetType() == typeof(string)
                        && double.TryParse(current.ToString(), out _))
                    {
                        amount = double.Parse((string)current);
                    }

                    else if (current.GetType() == typeof(double))
                    {
                        amount = (double)current;
                    }
                }

                if (sku == null || amount == null)
                {
                    continue;
                }

                if (SkuAmount.ContainsKey(sku))
                {
                    SkuAmount[sku] += amount.Value;
                }
                else
                {
                    SkuAmount.Add(sku, amount.Value);
                }
            }
        }
    }
}


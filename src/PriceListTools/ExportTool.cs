using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RehauSku.Assistant;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class ExportTool : AbstractPriceListTool, IDisposable
    {
        private Dictionary<string, double> SkuAmount { get; set; }
        private Range Selection;

        public ExportTool()
        {
            ExcelApp = (Application)ExcelDnaUtil.Application;
            Selection = ExcelApp.Selection;
        }

        public override void GetSource()
        {
            if (Selection != null && Selection.Columns.Count == 2)
                FillSkuAmountDict();

            else throw new Exception("Неверный диапазон");
        }

        private void FillSkuAmountDict()
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
                    continue;

                if (SkuAmount.ContainsKey(sku))
                    SkuAmount[sku] += amount.Value;
                else
                    SkuAmount.Add(sku, amount.Value);
            }
        }

        public override void FillPriceList()
        {
            if (SkuAmount.Count < 1) return;

            PriceListSheet offer = NewPriceList.OfferSheet;
            offer.Sheet.Activate();

            int exportedValues = 0;

            foreach (var kvp in SkuAmount)
            {
                Range cell = offer.Sheet.Columns[offer.skuColumnNumber].Find(kvp.Key);

                if (cell == null)
                {
                    System.Windows.Forms.MessageBox.Show
                        ($"Артикул {kvp.Key} отсутствует в таблице заказов {RegistryUtil.PriceListPath}",
                        "Отсутствует позиция в конечной таблице заказов",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }

                else
                {
                    Range sumCell = offer.Sheet.Cells[cell.Row, offer.amountColumnNumber];

                    if (sumCell.Value2 == null)
                        sumCell.Value2 = kvp.Value;
                    else
                        sumCell.Value2 += kvp.Value;

                    exportedValues++;
                }
            }

            AutoFilter filter = offer.Sheet.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(offer.amountColumnNumber - firstFilterColumn + 1, "<>");
            offer.Sheet.Range["A1"].Activate();
            offer.Sheet.Application.StatusBar = $"Экспортировано {exportedValues} строк из {SkuAmount.Count}";
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}


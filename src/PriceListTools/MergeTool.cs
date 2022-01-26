using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : AbstractPriceListTool, IDisposable
    {
        public override void FillPriceList()
        {
            PriceListSheet offer = NewPriceList.Sheet;
            offer.Sheet.Activate();

            int exportedValues = 0;

            foreach (var priceList in sourcePriceLists)
            {
                PriceListSheet sheet = priceList.Sheet;

                if (sheet.SkuAmount.Count == 0)
                    continue;

                foreach (var kvp in sheet.SkuAmount)
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
            }

            AutoFilter filter = offer.Sheet.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(offer.amountColumnNumber - firstFilterColumn + 1, "<>");
            offer.Sheet.Range["A1"].Activate();
            AddIn.Excel.StatusBar = $"Экспортировано {exportedValues} строк из {sourcePriceLists.Count} файлов";

            Forms.Dialog.SaveWorkbookAs();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}

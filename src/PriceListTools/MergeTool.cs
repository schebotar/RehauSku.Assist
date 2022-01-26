using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : AbstractPriceListTool, IDisposable
    {
        public override void FillPriceList()
        {
            int exportedValues = 0;

            foreach (var sheet in sourcePriceLists)
            {
                if (sheet.SkuAmount.Count == 0)
                    continue;

                foreach (var kvp in sheet.SkuAmount)
                {
                    Range cell = NewPriceList.Sheet.Columns[NewPriceList.skuCell.Column].Find(kvp.Key);

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
                        Range sumCell = NewPriceList.Sheet.Cells[cell.Row, NewPriceList.amountCell.Column];

                        if (sumCell.Value2 == null)
                            sumCell.Value2 = kvp.Value;
                        else
                            sumCell.Value2 += kvp.Value;

                        exportedValues++;
                    }
                }
            }

            FilterByAmount();
            AddIn.Excel.StatusBar = $"Экспортировано {exportedValues} строк из {sourcePriceLists.Count} файлов";
            Forms.Dialog.SaveWorkbookAs();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}

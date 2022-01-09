using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : AbstractPriceListTool, IDisposable
    {
        public override void FillPriceList()
        {
            PriceListSheet offer = NewPriceList.OfferSheet;
            offer.Sheet.Activate();

            int exportedValues = 0;
            int exportedLists = 0;

            foreach (var priceList in sourcePriceLists)
            {
                foreach (var sheet in priceList.Sheets)
                {
                    if (sheet.SkuAmount.Count == 0)
                        continue;

                    offer.Sheet.Columns[offer.amountColumnNumber]
                        .EntireColumn
                        .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                    exportedLists++;

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
                            offer.Sheet.Cells[cell.Row, offer.amountColumnNumber].Value2 = kvp.Value;
                            Range sumCell = offer.Sheet.Cells[cell.Row, offer.amountColumnNumber + exportedLists];

                            if (sumCell.Value2 == null)
                                sumCell.Value2 = kvp.Value;
                            else
                                sumCell.Value2 += kvp.Value;

                            exportedValues++;
                        }

                        offer.Sheet.Cells[offer.headerRowNumber, offer.amountColumnNumber].Value2 = $"{priceList.Name}\n{sheet.Name}";
                    }
                }
            }

            AutoFilter filter = offer.Sheet.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(offer.amountColumnNumber - firstFilterColumn + 1 + exportedLists, "<>");
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

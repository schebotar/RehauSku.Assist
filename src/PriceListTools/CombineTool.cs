﻿using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.PriceListTools
{
    internal class CombineTool : AbstractPriceListTool, IDisposable
    {
        public override void FillPriceList()
        {
            PriceList offer = NewPriceList;
            offer.Sheet.Activate();

            int exportedValues = 0;

            foreach (var sheet in sourcePriceLists)
            {
                if (sheet.SkuAmount.Count == 0)
                    continue;

                offer.Sheet.Columns[offer.amountCell.Column]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                foreach (var kvp in sheet.SkuAmount)
                {
                    Range cell = offer.Sheet.Columns[offer.skuCell.Column].Find(kvp.Key);

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
                        offer.Sheet.Cells[cell.Row, offer.amountCell.Column - 1].Value2 = kvp.Value;
                        Range sumCell = offer.Sheet.Cells[cell.Row, offer.amountCell.Column];

                        if (sumCell.Value2 == null)
                            sumCell.Value2 = kvp.Value;
                        else
                            sumCell.Value2 += kvp.Value;

                        exportedValues++;
                    }

                    offer.Sheet.Cells[offer.amountCell.Row, offer.amountCell.Column - 1].Value2 = $"{sheet.Name}";
                }
            }

            AutoFilter filter = offer.Sheet.AutoFilter;

            filter.Range.AutoFilter(offer.amountCell.Column, "<>");
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
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace RehauSku.PriceListTools
{
    class PriceList
    {
        public readonly Workbook Workbook;
        public readonly PriceListSheet OfferSheet;
        public readonly PriceListSheet ActiveSheet;

        private const string amountHeader = "Кол-во";
        private const string skuHeader = "Актуальный материал";
        private const string offerSheetHeader = "КП";

        public PriceList(Workbook workbook)
        {
            Workbook = workbook;
            OfferSheet = new PriceListSheet(workbook.Sheets[offerSheetHeader]);

            Worksheet active = workbook.ActiveSheet;

            if (active.Name == offerSheetHeader)
                ActiveSheet = OfferSheet;

            else
                ActiveSheet = new PriceListSheet(active);
        }

        public void FillWithValues(Dictionary<string, double> values)
        {
            Worksheet ws = OfferSheet.sheet;
            ws.Activate();

            int amountColumn = OfferSheet.amountColumn.Value;
            int skuColumn = OfferSheet.skuColumn.Value;
            int exportedValues = 0;

            foreach (KeyValuePair<string, double> kvp in values)
            {
                Range cell = ws.Columns[skuColumn].Find(kvp.Key);

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
                    ws.Cells[cell.Row, amountColumn].Value = kvp.Value;
                    exportedValues++;
                }
            }

            AutoFilter filter = ws.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(amountColumn - firstFilterColumn + 1, "<>");
            ws.Range["A1"].Activate();
            ws.Application.StatusBar = $"Экспортировано {exportedValues} строк из {values.Count}";
        }

        public void FillWithValues(Dictionary<string, double>[] values, string[] filenames)
        {
            Worksheet ws = OfferSheet.sheet;
            ws.Activate();

            int amountColumn = OfferSheet.amountColumn.Value;
            int skuColumn = OfferSheet.skuColumn.Value;
            int headerColumn = OfferSheet.headerRow.Value;

            int exportedValues = 0;

            for (int i = 0; i < values.Length; i++)
            {
                ws.Columns[amountColumn]
                    .EntireColumn
                    .Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                foreach (var kvp in values[i])
                {
                    Range cell = ws.Columns[skuColumn].Find(kvp.Key);

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
                        ws.Cells[cell.Row, amountColumn].Value2 = kvp.Value;
                        Range sumCell = ws.Cells[cell.Row, amountColumn + i + 1];

                        if (sumCell.Value2 == null)
                            sumCell.Value2 = kvp.Value;
                        else
                            sumCell.Value2 += kvp.Value;

                        exportedValues++;
                    }
                }

                ws.Cells[headerColumn, amountColumn].Value2 = filenames[i];
            }

            AutoFilter filter = ws.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(amountColumn - firstFilterColumn + 1 + values.Length, "<>");
            ws.Range["A1"].Activate();
            ws.Application.StatusBar = $"Экспортировано {exportedValues} строк из {values.Sum(x => x.Count)}";
        }

        public class PriceListSheet
        {
            public readonly Worksheet sheet;
            public readonly int? headerRow;

            public readonly int? skuColumn;
            public readonly int? amountColumn;

            public object[,] skuCells;
            public object[,] amountCells;

            public PriceListSheet(Worksheet sheet)
            {
                this.sheet = sheet;
                headerRow = sheet.Cells.Find(amountHeader).Row;
                amountColumn = sheet.Cells.Find(amountHeader).Column;
                skuColumn = sheet.Cells.Find(skuHeader).Column;

                amountCells = sheet.Columns[amountColumn].Value2;
                skuCells = sheet.Columns[skuColumn].Value2;
            }            
        }
    }
}


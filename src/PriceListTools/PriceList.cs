using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    class PriceList
    {
        public readonly Workbook Workbook;
        public readonly PriceListSheet OfferSheet;
        public readonly PriceListSheet ActiveSheet;

        private const string _amountHeader = "Кол-во";
        private const string _skuHeader = "Актуальный материал";
        private const string _offerSheetHeader = "КП";

        public PriceList(Workbook workbook)
        {
            Workbook = workbook;
            OfferSheet = new PriceListSheet(workbook.Sheets[_offerSheetHeader]);

            Worksheet active = workbook.ActiveSheet;

            if (active.Name == _offerSheetHeader)
                ActiveSheet = OfferSheet;

            else
                ActiveSheet = new PriceListSheet(active);                
        }

        public bool IsValid()
        {
            return OfferSheet.IsValid() &&
                ActiveSheet.IsValid();
        }

        public void Fill(Dictionary<string, double> values)
        {
            Worksheet ws = OfferSheet.sheet;
            ws.Activate();

            int amountColumn = OfferSheet.amountColumn.Value;
            int skuColumn = OfferSheet.skuColumn.Value;

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
                    ws.Cells[cell.Row, amountColumn].Value = kvp.Value;
            }

            AutoFilter filter = ws.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(amountColumn - firstFilterColumn + 1, "<>");
            ws.Range["A1"].Activate();
        }

        public class PriceListSheet
        {
            public readonly Worksheet sheet;
            public readonly int? headerRow;
            public readonly int? amountColumn;
            public readonly int? skuColumn;
            public object[,] amountCells;
            public object[,] skuCells;

            public PriceListSheet(Worksheet sheet)
            {
                this.sheet = sheet;
                headerRow = sheet.Cells.Find(_amountHeader).Row;
                amountColumn = sheet.Cells.Find(_amountHeader).Column;
                skuColumn = sheet.Cells.Find(_skuHeader).Column;

                amountCells = sheet.Columns[amountColumn].Value2;
                skuCells = sheet.Columns[skuColumn].Value2;
            }

            public bool IsValid()
            {
                return sheet != null &&
                    headerRow != null &&
                    amountColumn != null &&
                    skuColumn != null;
            }
        }
    }
}


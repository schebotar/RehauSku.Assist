using Microsoft.Office.Interop.Excel;
using RehauSku.PriceListTools;
using System.Linq;

namespace RehauSku
{
    public static class WorksheetExtensions
    {
        public static bool IsRehauSource(this Worksheet worksheet)
        {
            Range amountCell;
            Range skuCell;
            Range groupCell;
            Range nameCell;

            Range[] cells = new[]
            {
                amountCell = worksheet.Cells.Find(PriceListHeaders.Amount),
                skuCell = worksheet.Cells.Find(PriceListHeaders.Sku),
                groupCell = worksheet.Cells.Find(PriceListHeaders.Group),
                nameCell = worksheet.Cells.Find(PriceListHeaders.Name)
            };

            return cells.All(x => x != null);
        }

        public static void AddValue(this Range range, double value)
        {
            if (range.Value2 == null)
            {
                range.Value2 = value;
            }

            else
            {
                range.Value2 += value;
            }
        }
    }
}


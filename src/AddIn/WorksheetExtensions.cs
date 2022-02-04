using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace RehauSku
{
    public static class WorksheetExtensions
    {
        private static string amountHeader = "Кол-во";
        private static string skuHeader = "Актуальный материал";
        private static string groupHeader = "Программа";
        private static string nameHeader = "Наименование";

        public static bool IsRehauSource(this Worksheet worksheet)
        {
            Range amountCell;
            Range skuCell;
            Range groupCell;
            Range nameCell;

            Range[] cells = new[]
            {
                amountCell = worksheet.Cells.Find(amountHeader),
                skuCell = worksheet.Cells.Find(skuHeader),
                groupCell = worksheet.Cells.Find(groupHeader),
                nameCell = worksheet.Cells.Find(nameHeader)
            };

            return cells.All(x => x != null);
        }
    }
}


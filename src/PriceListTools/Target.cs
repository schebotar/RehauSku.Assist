using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class Target : AbstractPriceList
    {
        private const string oldSkuHeader = "Прежний материал";
        public Range oldSkuCell { get; private set; }

        public Target(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.FullName;

            Range[] cells = new[]
            {
                amountCell = Sheet.Cells.Find(amountHeader),
                skuCell = Sheet.Cells.Find(skuHeader),
                groupCell = Sheet.Cells.Find(groupHeader),
                nameCell = Sheet.Cells.Find(nameHeader),
                oldSkuCell = Sheet.Cells.Find(oldSkuHeader)
            };

            if (cells.Any(x => x == null))
            {
                throw new ArgumentException($"Шаблон { Name } не является прайс-листом");
            }
        }
    }
}


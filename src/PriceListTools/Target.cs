using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.PriceListTools
{
    internal class Target : PriceList
    {
        public Target(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.FullName;

            amountCell = Sheet.Cells.Find(amountHeader);
            skuCell = Sheet.Cells.Find(skuHeader);
            groupCell = Sheet.Cells.Find(groupHeader);

            if (amountCell == null || skuCell == null || groupCell == null)
            {
                throw new ArgumentException($"Шаблон { Name } не является прайс-листом");
            }
        }
    }
}


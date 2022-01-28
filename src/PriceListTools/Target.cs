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
            nameCell = Sheet.Cells.Find(nameHeader);

            if (amountCell == null || skuCell == null || groupCell == null || nameCell == null)
            {
                throw new ArgumentException($"Шаблон { Name } не является прайс-листом");
            }
        }
    }
}


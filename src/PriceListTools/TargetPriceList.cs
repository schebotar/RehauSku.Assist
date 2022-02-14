using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class TargetPriceList : AbstractPriceList
    {
        public Range OldSkuCell { get; private set; }

        public TargetPriceList(Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentException("Невозможно открыть книгу шаблонного файла. " +
                    "Возможно открыт файл с именем, совпадающим с именем шаблонного файла.");
            }

            Sheet = workbook.ActiveSheet;
            Name = workbook.FullName;

            Range[] cells = new[]
            {
                AmountCell = Sheet.Cells.Find(PriceListHeaders.Amount),
                SkuCell = Sheet.Cells.Find(PriceListHeaders.Sku),
                GroupCell = Sheet.Cells.Find(PriceListHeaders.Group),
                NameCell = Sheet.Cells.Find(PriceListHeaders.Name)
            };

            OldSkuCell = Sheet.Cells.Find(PriceListHeaders.OldSku);

            if (cells.Any(x => x == null))
            {
                throw new ArgumentException($"Шаблон { Name } не является прайс-листом");
            }
        }
    }
}


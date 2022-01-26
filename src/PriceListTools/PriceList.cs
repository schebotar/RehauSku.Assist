using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class PriceList
    {
        public readonly string Name;
        public PriceListSheet Sheet { get; private set; }

        public PriceList(Workbook workbook)
        {
            Name = workbook.Name;
            Sheet = new PriceListSheet(workbook.ActiveSheet);
        }
    }
}


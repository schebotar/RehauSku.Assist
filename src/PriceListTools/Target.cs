using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class Target : PriceList
    {
        public Dictionary<PriceListPosition, Range> Map { get; private set; }

        public Target(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.Name;

            amountCell = Sheet.Cells.Find(amountHeader);
            skuCell = Sheet.Cells.Find(skuHeader);
            groupCell = Sheet.Cells.Find(groupHeader);

            if (amountCell == null || skuCell == null || groupCell == null)
            {
                throw new ArgumentException($"Лист { Name } не распознан");
            }

            CreateMap();
        }

        private void CreateMap()
        {
            Range amountCell = Sheet.Cells.Find(amountHeader);
            Range skuCell = Sheet.Cells.Find(skuHeader);
            Range groupCell = Sheet.Cells.Find(groupHeader);

            //headerRowNumber = amountCell.Row;
            //skuColumnNumber = skuCell.Column;
            //amountColumnNumber = amountCell.Column;
            //groupColumnNumber = groupCell.Column;

            //for (int row = headerRowNumber + 1; row < skuCell.Rows.Count; row++)
            //{
            //    string sku =
            //}
        }
    }
}


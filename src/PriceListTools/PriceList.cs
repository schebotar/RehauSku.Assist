using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System;

namespace RehauSku.PriceListTools
{
    internal class PriceList
    {
        private const string amountHeader = "Кол-во";
        private const string skuHeader = "Актуальный материал";
        private const string groupHeader = "Программа";

        public readonly Worksheet Sheet;
        public readonly string Name;
        public Dictionary<string, double> SkuAmount { get; private set; }

        public readonly Range amountCell;
        public readonly Range skuCell;
        public readonly Range groupCell;

        //public Dictionary<PriceListPosition, Range> Map { get; private set; }

        public PriceList(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.Name + '\n' + Sheet.Name;

            amountCell = Sheet.Cells.Find(amountHeader);
            skuCell = Sheet.Cells.Find(skuHeader);
            groupCell = Sheet.Cells.Find(groupHeader);

            if (amountCell == null || skuCell == null || groupCell == null)
            {
                throw new ArgumentException($"Лист { Name } не распознан");
            }

            FillSkuAmount();
        }

        private void FillSkuAmount()
        {
            SkuAmount = new Dictionary<string, double>();

            object[,] amountColumn = Sheet.Columns[amountCell.Column].Value2;
            object[,] skuColumn = Sheet.Columns[skuCell.Column].Value2;

            for (int row = amountCell.Row + 1; row < amountColumn.GetLength(0); row++)
            {
                object amount = amountColumn[row, 1];
                object sku = skuColumn[row, 1];

                if (amount != null && (double)amount != 0)
                {
                    if (SkuAmount.ContainsKey(sku.ToString()))
                        SkuAmount[sku.ToString()] += (double)amount;

                    else
                        SkuAmount.Add(sku.ToString(), (double)amount);
                }
            }
        }

        //public void CreateMap()
        //{
        //    Range amountCell = Sheet.Cells.Find(amountHeader);
        //    Range skuCell = Sheet.Cells.Find(skuHeader);
        //    Range groupCell = Sheet.Cells.Find(groupHeader);

        //    headerRowNumber = amountCell.Row;
        //    skuColumnNumber = skuCell.Column;
        //    amountColumnNumber = amountCell.Column;
        //    groupColumnNumber = groupCell.Column;

        //    for (int row = headerRowNumber + 1; row < skuCell.Rows.Count; row++)
        //    {
        //        string sku = 
        //    }

        //}
    }

}


using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class Source : PriceList
    {
        public Dictionary<string, double> SkuAmount { get; private set; }

        public Source(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.Name;

            amountCell = Sheet.Cells.Find(amountHeader);
            skuCell = Sheet.Cells.Find(skuHeader);
            groupCell = Sheet.Cells.Find(groupHeader);

            if (amountCell == null || skuCell == null || groupCell == null)
            {
                throw new ArgumentException($"Файл {Name} не распознан");
            }

            CreateAmountDict();
        }

        private void CreateAmountDict()
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
    }
}


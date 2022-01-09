using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class PriceListSheet
    {
        private const string amountHeader = "Кол-во";
        private const string skuHeader = "Актуальный материал";

        public readonly Worksheet Sheet;
        public readonly string Name;
        public Dictionary<string, double> SkuAmount { get; private set; }
        public int headerRowNumber { get; private set; }
        public int amountColumnNumber { get; private set; }
        public int skuColumnNumber { get; private set; }

        public PriceListSheet(Worksheet sheet)
        {
            Sheet = sheet;
            Name = sheet.Name;
            SkuAmount = new Dictionary<string, double>();

            FillSkuAmount();
        }

        public bool FillSkuAmount()
        {
            Range amountCell = Sheet.Cells.Find(amountHeader);
            Range skuCell = Sheet.Cells.Find(skuHeader);

            if (amountCell == null || skuCell == null)
            {
                Sheet.Application.StatusBar = $"Лист {Name} не распознан";
                return false;
            }

            headerRowNumber = amountCell.Row;
            skuColumnNumber = skuCell.Column;
            amountColumnNumber = amountCell.Column;

            object[,] amountColumn = Sheet.Columns[amountColumnNumber].Value2;
            object[,] skuColumn = Sheet.Columns[skuColumnNumber].Value2;

            for (int row = headerRowNumber + 1; row < amountColumn.GetLength(0); row++)
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
            return true;
        }
    }

}


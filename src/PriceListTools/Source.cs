using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class Source : PriceList
    {
        public Dictionary<Position, double> PositionAmount { get; private set; }

        public Source(Workbook workbook)
        {
            Sheet = workbook.ActiveSheet;
            Name = workbook.Name;

            amountCell = Sheet.Cells.Find(amountHeader);
            skuCell = Sheet.Cells.Find(skuHeader);
            groupCell = Sheet.Cells.Find(groupHeader);
            nameCell = Sheet.Cells.Find(nameHeader);

            if (amountCell == null || skuCell == null || groupCell == null || nameCell == null)
            {
                throw new ArgumentException($"Файл {Name} не распознан");
            }

            CreatePositionsDict();
        }

        private void CreatePositionsDict()
        {
            PositionAmount = new Dictionary<Position, double>();

            var aColumn = amountCell.EntireColumn;

            object[,] amountColumn = amountCell.EntireColumn.Value2;
            object[,] skuColumn = skuCell.EntireColumn.Value2;
            object[,] nameColumn = nameCell.EntireColumn.Value2;
            object[,] groupColumn = groupCell.EntireColumn.Value2;

            for (int row = amountCell.Row + 1; row < amountColumn.GetLength(0); row++)
            {
                object amount = amountColumn[row, 1];
                object group = groupColumn[row, 1];
                object name = nameColumn[row, 1];
                object sku = skuColumn[row, 1];

                if (amount != null && (double)amount != 0)
                {
                    Position p = new Position(group.ToString(), sku.ToString(), name.ToString());

                    if (PositionAmount.ContainsKey(p))
                    {
                        PositionAmount[p] += (double)amount;
                    }
                    else
                    {
                        PositionAmount.Add(p, (double)amount);
                    }
                }
            }
        }
    }
}


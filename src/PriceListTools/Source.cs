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
            if (workbook == null)
            {
                throw new ArgumentException($"Нет рабочего файла");
            }

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

            for (int row = amountCell.Row + 1; row < Sheet.AutoFilter.Range.Rows.Count; row++)
            {
                object amount = Sheet.Cells[row, amountCell.Column].Value2;

                if (amount != null && (double)amount != 0)
                {
                    object group = Sheet.Cells[row, groupCell.Column].Value2;
                    object name = Sheet.Cells[row, nameCell.Column].Value2;
                    object sku = Sheet.Cells[row, skuCell.Column].Value2;

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


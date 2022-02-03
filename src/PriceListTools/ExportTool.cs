using Microsoft.Office.Interop.Excel;
using RehauSku.Assistant;
using System;
using System.Collections.Generic;
using RehauSku.Interface;

namespace RehauSku.PriceListTools
{
    internal class ExportTool : AbstractTool
    {
        private Dictionary<Position, double> PositionAmount;
        private Range Selection;

        public void TryGetSelection()
        {
            Selection = ExcelApp.Selection;

            if (Selection == null || Selection.Columns.Count != 2)
            {
                throw new Exception("Неверный диапазон");
            }
        }

        public void FillTarget()
        {
            GetSelected();
            ProgressBar bar = new ProgressBar(PositionAmount.Count);
            
            foreach (var kvp in PositionAmount)
            {
                FillPositionAmountToColumns(kvp, TargetFile.amountCell.Column);
                bar.DoProgress();
            }

            FilterByAmount();

            Interface.Dialog.SaveWorkbookAs();
        }

        private void GetSelected()
        {
            object[,] cells = Selection.Value2;            
            PositionAmount = new Dictionary<Position, double>();

            int rowsCount = Selection.Rows.Count;

            for (int row = 1; row <= rowsCount; row++)
            {
                if (cells[row, 1] == null || cells[row, 2] == null)
                    continue;

                string sku = null;
                double? amount = null;

                for (int column = 1; column <= 2; column++)
                {
                    object current = cells[row, column];

                    if (current.ToString().IsRehauSku())
                    {
                        sku = current.ToString();
                    }

                    else if (current.GetType() == typeof(string)
                        && double.TryParse(current.ToString(), out _))
                    {
                        amount = double.Parse((string)current);
                    }

                    else if (current.GetType() == typeof(double))
                    {
                        amount = (double)current;
                    }
                }

                if (sku == null || amount == null)
                {
                    continue;
                }

                Position position = new Position(null, sku, null);

                if (PositionAmount.ContainsKey(position))
                {
                    PositionAmount[position] += amount.Value;
                }

                else
                {
                    PositionAmount.Add(position, amount.Value);
                }
            }
        }
    }
}


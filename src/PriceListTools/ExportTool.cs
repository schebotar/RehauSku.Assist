using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using RehauSku.Interface;
using System.Threading.Tasks;

namespace RehauSku.PriceListTools
{
    internal class ExportTool : AbstractTool
    {
        private Dictionary<Position, double> PositionAmount;
        private readonly Range Selection;

        public ExportTool()
        {
            Selection = ExcelApp.Selection;
            GetSelected();

            if (PositionAmount.Count == 0)
            {
                throw new Exception("В выделенном диапазоне не найдены позиции для экспорта");
            }
        }

        public override async void FillTarget()
        {
            ProgressBar = new ProgressBar("Заполняю строки...", PositionAmount.Count);
            ResultBar = new ResultBar();

            foreach (var kvp in PositionAmount)
            {
                FillPositionAmountToColumns(kvp, TargetFile.AmountCell.Column);
                ProgressBar.Update();
            }

            FilterByAmount();
            ResultBar.Update();

            await Task.Delay(new TimeSpan(0, 0, 5));
            ExcelApp.StatusBar = false;
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


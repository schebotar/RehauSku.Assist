using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using RehauSku.Interface;

namespace RehauSku.PriceListTools
{
    internal class SourcePriceList : AbstractPriceList
    {
        public Dictionary<Position, double> PositionAmount { get; private set; }

        public SourcePriceList(Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentException($"Нет рабочего файла");
            }

            Sheet = workbook.ActiveSheet;
            Name = workbook.Name;

            Range[] cells = new[]
            {
                AmountCell = Sheet.Cells.Find(PriceListHeaders.Amount),
                SkuCell = Sheet.Cells.Find(PriceListHeaders.Sku),
                GroupCell = Sheet.Cells.Find(PriceListHeaders.Group),
                NameCell = Sheet.Cells.Find(PriceListHeaders.Name)
            };

            if (cells.Any(x => x == null))
            {
                throw new ArgumentException($"Файл {Name} не распознан");
            }

            CreatePositionsDict();
        }

        public static List<SourcePriceList> GetSourceLists(string[] files)
        {
            var ExcelApp = (Application)ExcelDnaUtil.Application;
            ProgressBar bar = new ProgressBar("Открываю исходные файлы...", files.Length);

            List<SourcePriceList> sourceFiles = new List<SourcePriceList>();

            foreach (string file in files)
            {
                ExcelApp.ScreenUpdating = false;
                Workbook wb = ExcelApp.Workbooks.Open(file);
                try
                {
                    SourcePriceList priceList = new SourcePriceList(wb);
                    sourceFiles.Add(priceList);
                    wb.Close();
                    bar.Update();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show
                        (ex.Message,
                        "Ошибка открытия исходного прайс-листа",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    wb.Close();
                    bar.Update();
                }
                ExcelApp.ScreenUpdating = true;
            }

            return sourceFiles;
        }

        private void CreatePositionsDict()
        {
            PositionAmount = new Dictionary<Position, double>();

            for (int row = AmountCell.Row + 1; row <= Sheet.Cells[Sheet.Rows.Count, AmountCell.Column].End[XlDirection.xlUp].Row; row++)
            {
                object amount = Sheet.Cells[row, AmountCell.Column].Value2;

                if (amount != null && (double)amount != 0)
                {
                    object group = Sheet.Cells[row, GroupCell.Column].Value2;
                    object name = Sheet.Cells[row, NameCell.Column].Value2;
                    object sku = Sheet.Cells[row, SkuCell.Column].Value2;

                    if (group == null || name == null || sku == null)
                        continue;

                    if (!sku.ToString().IsRehauSku())
                        continue;

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


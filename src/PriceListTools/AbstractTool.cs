using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using ProgressBar = RehauSku.Interface.ProgressBar;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractTool
    {
        protected private Application ExcelApp = (Application)ExcelDnaUtil.Application;
        protected private TargetPriceList TargetFile;
        protected private ResultBar ResultBar { get; set; }
        protected private ProgressBar ProgressBar { get; set; }

        public void OpenNewPrice()
        {
            if (ExcelApp.Workbooks
                .Cast<Workbook>()
                .FirstOrDefault(w => w.FullName == RegistryUtil.PriceListPath) != null)
            {
                MessageBox.Show
                    ("Шаблонный файл редактируется в другом месте",
                    "Ошибка открытия шаблонного прайс-листа",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                throw new ArgumentException("Шаблонный файл редактируется в другом месте");
            }

            Workbook wb = ExcelApp.Workbooks.Open(RegistryUtil.PriceListPath, null, true);

            try
            {
                TargetFile = new TargetPriceList(wb);
            }

            catch (Exception exception)
            {
                MessageBox.Show
                    (exception.Message,
                    "Ошибка открытия шаблонного прайс-листа",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (wb != null)
                {
                    wb.Close();
                }

                throw exception;
            }
        }

        protected private void FillPositionAmountToColumns(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            int? row = GetPositionRow(TargetFile.skuCell.EntireColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = TargetFile.Sheet.Cells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementSuccess();
                return;
            }

            if (TargetFile.oldSkuCell != null)
            {
                row = GetPositionRow(TargetFile.oldSkuCell.EntireColumn, positionAmount.Key.Sku, positionAmount.Key.Group);

                if (row != null)
                {
                    foreach (int column in columns)
                    {
                        Range cell = TargetFile.Sheet.Cells[row, column];
                        cell.AddValue(positionAmount.Value);
                    }

                    ResultBar.IncrementReplaced();
                    return;
                }
            }

            string sku = positionAmount.Key.Sku.Substring(1, 6);
            row = GetPositionRow(TargetFile.skuCell.EntireColumn, sku, positionAmount.Key.Group);

            if (row != null)
            {
                foreach (int column in columns)
                {
                    Range cell = TargetFile.Sheet.Cells[row, column];
                    cell.AddValue(positionAmount.Value);
                }

                ResultBar.IncrementReplaced();
                return;
            }

            FillMissing(positionAmount, columns);
            ResultBar.IncrementNotFound();
        }

        protected private void FillMissing(KeyValuePair<Position, double> positionAmount, params int[] columns)
        {
            int row = TargetFile.Sheet.Cells[TargetFile.Sheet.Rows.Count, TargetFile.skuCell.Column]
                .End[XlDirection.xlUp]
                .Row + 1;

            TargetFile.Sheet.Rows[row]
                .EntireRow
                .Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            Range previous = TargetFile.Sheet.Rows[row - 1];
            Range current = TargetFile.Sheet.Rows[row];

            previous.Copy(current);
            current.ClearContents();

            TargetFile.Sheet.Cells[row, TargetFile.groupCell.Column].Value2 = positionAmount.Key.Group;
            TargetFile.Sheet.Cells[row, TargetFile.nameCell.Column].Value2 = positionAmount.Key.Name;

            if (TargetFile.oldSkuCell != null)
            {
                TargetFile.Sheet.Cells[row, TargetFile.skuCell.Column].Value2 = "Не найден";
                TargetFile.Sheet.Cells[row, TargetFile.oldSkuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            else
            {
                TargetFile.Sheet.Cells[row, TargetFile.skuCell.Column].Value2 = positionAmount.Key.Sku;
            }

            foreach (int column in columns)
            {
                Range cell = TargetFile.Sheet.Cells[row, column];
                cell.AddValue(positionAmount.Value);
            }
        }

        protected private int? GetPositionRow(Range range, string sku, string group)
        {
            Range found = range.Find(sku);
            string foundGroupValue;

            if (found == null)
            {
                return null;
            }

            int firstFoundRow = found.Row; 

            while (true)
            {
                foundGroupValue = TargetFile.Sheet.Cells[found.Row, TargetFile.groupCell.Column].Value2.ToString();

                if (string.IsNullOrEmpty(group) || group.Equals(foundGroupValue))
                {
                    return found.Row;
                }

                found = range.FindNext(found);

                if (found.Row == firstFoundRow)
                {
                    return null;
                }
            }
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = TargetFile.Sheet.AutoFilter;
            int startColumn = filter.Range.Column;

            filter.Range.AutoFilter(TargetFile.amountCell.Column - startColumn + 1, "<>");
            TargetFile.Sheet.Range["A1"].Activate();
        }
    }
}
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using RehauSku.Assistant;

namespace RehauSku.DataExport
{
    public class ExportTool : IDisposable
    {
        private Application xlApp;
        private Dictionary<string, double> SkuAmount { get; set; }
        private Range Selection { get; set; }
        private string CurrentFilePath { get; set; }

        public ExportTool()
        {
            this.xlApp = (Application)ExcelDnaUtil.Application;
            this.CurrentFilePath = xlApp.ActiveWorkbook.FullName;

            _GetSelectedCells();
        }

        private void _GetSelectedCells()
        {
            Selection = xlApp.Selection;
        }

        public bool IsRangeValid()
        {
            return Selection.Columns.Count == 2;
        }

        private void FillSkuAmountDict()
        {
            object[,] cells = Selection.Value2;
            SkuAmount = new Dictionary<string, double>();
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

                    if (current.GetType() == typeof(string)
                        && ((string)current).IsRehauSku())
                        sku = (string)current;

                    else if (current.GetType() == typeof(string)
                        && double.TryParse((string)current, out _))
                        amount = double.Parse((string)current);

                    else if (current.GetType() == typeof(double))
                        amount = (double)current;
                }

                if (sku == null || amount == null)
                    continue;

                if (SkuAmount.ContainsKey(sku))
                    SkuAmount[sku] += amount.Value;
                else
                    SkuAmount.Add(sku, amount.Value);
            }
        }

        public void FillNewPriceList()
        {
            const string amountHeader = "Кол-во";
            const string skuHeader = "Актуальный материал";

            FillSkuAmountDict();
            string exportFile = _GetExportFullPath();
            File.Copy(RegistryUtil.PriceListPath, exportFile, true);

            Workbook wb = xlApp.Workbooks.Open(exportFile);
            Worksheet ws = wb.Sheets["КП"];
            ws.Activate();

            int amountColumn = ws.Cells.Find(amountHeader).Column;
            int skuColumn = ws.Cells.Find(skuHeader).Column;

            foreach (KeyValuePair<string, double> kvp in SkuAmount)
            {
                Range cell = ws.Columns[skuColumn].Find(kvp.Key);
                ws.Cells[cell.Row, amountColumn].Value = kvp.Value;
            }

            AutoFilter filter = ws.AutoFilter;
            int firstFilterColumn = filter.Range.Column;

            filter.Range.AutoFilter(amountColumn - firstFilterColumn + 1, "<>");
            ws.Range["A1"].Activate();
        }

        private string _GetExportFullPath()
        {
            string fileExtension = Path.GetExtension(RegistryUtil.PriceListPath);

            return Path.GetTempFileName() + fileExtension;
        }


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {

        }
    }

    class SelectionCheck
    {

    }
}


using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RehauSku.Assistant;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    class ExportTool : IDisposable
    {
        private Application ExcelApp;
        private Dictionary<string, double> SkuAmount { get; set; }
        private Range Selection { get; set; }

        public ExportTool()
        {
            this.ExcelApp = (Application)ExcelDnaUtil.Application;
            Selection = ExcelApp.Selection;

            if (IsRangeValid())
                _FillSkuAmountDict();
        }

        public bool IsRangeValid()
        {
            return Selection != null &&
                Selection.Columns.Count == 2;
        }

        private void _FillSkuAmountDict()
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
                    continue;

                if (SkuAmount.ContainsKey(sku))
                    SkuAmount[sku] += amount.Value;
                else
                    SkuAmount.Add(sku, amount.Value);
            }
        }

        public void ExportToNewFile()
        {
            if (SkuAmount.Count < 1)
            {
                return;
            }

            string exportFile = PriceListUtil.CreateNewExportFile();
            Workbook wb = ExcelApp.Workbooks.Open(exportFile);

            try
            {
                PriceList priceList = new PriceList(wb);
                priceList.Fill(SkuAmount);
            }

            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show
                    ($"{RegistryUtil.PriceListPath} не является файлом прайслиста \n\n {ex.Message}",
                    "Неверный файл прайс-листа!",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

                wb.Close();
            }
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
}


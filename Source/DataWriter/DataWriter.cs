using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Rehau.Sku.Assist;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Rehau.Sku
{
    public class DataWriter : IDisposable
    {
        private Application xlApp;
        private Dictionary<string, double> SkuAmount { get; set; }
        private object[,] SelectedCells { get; set; }
        private string FileName { get; set; }
        private string ExportFileName { get; set; }

        public DataWriter()
        {
            this.xlApp = (Application)ExcelDnaUtil.Application;
            this.FileName = this.xlApp.ActiveWorkbook.FullName;

            GetSelectedCells();
        }

        private void GetSelectedCells()
        {
            Range selection = xlApp.Selection;
            this.SelectedCells = (object[,])selection.Value2;
        }

        public bool IsRangeValid()
        {
            return SelectedCells != null &&
                SelectedCells.GetLength(1) == 2;
        }

        public void FillSkuAmountDict()
        {
            SkuAmount = new Dictionary<string, double>();
            int rowsCount = SelectedCells.GetLength(0);

            for (int row = 1; row <= rowsCount; row++)
            {
                if (SelectedCells[row, 1] == null || SelectedCells[row, 2] == null)
                    continue;

                string sku = null;
                double? amount = null;

                for (int column = 1; column <= 2; column++)
                {
                    object current = SelectedCells[row, column];

                    if (current.GetType() == typeof(string)
                        && SkuAssist.IsRehauSku((string)current))
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

        public void DownloadPriceList()
        {
            Uri linkToPrice = new Uri(@"https://www.rehau.com/downloads/831332/%D0%BF%D1%80%D0%B0%D0%B9%D1%81-%D0%BB%D0%B8%D1%81%D1%82-%D0%B2%D0%B8%D1%81-exel-2021.xlsm");

            if (FileName.Contains(':'))
                ExportFileName = string.Join(".", FileName.Split('.').Select((x, i) => i == 0 ? string.Concat(x, "~") : "xlsm"));
            else
                ExportFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Прайс-лист REHAU.xlsm";

            HttpClientUtil.DownloadFile(linkToPrice, ExportFileName);
        }

        public void GetPriceListWB()
        {            
            Workbook wb = xlApp.Workbooks.Open(ExportFileName);
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


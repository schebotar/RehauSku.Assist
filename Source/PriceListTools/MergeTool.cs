using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    class MergeTool : IDisposable
    {
        private Application ExcelApp;
        private Dictionary<string, double> SkuAmount { get; set; }

        public MergeTool()
        {
            this.ExcelApp = (Application)ExcelDnaUtil.Application;
            this.SkuAmount = new Dictionary<string, double>();
        }

        public void AddSkuAmountToDict(string[] files)
        {
            ExcelApp.ScreenUpdating = false;
            foreach (string file in files)
            {
                Workbook wb = ExcelApp.Workbooks.Open(file);
                PriceList priceList = new PriceList(wb);

                if (priceList.IsValid())
                    SkuAmount.AddValues(priceList);

                wb.Close();
            }
            ExcelApp.ScreenUpdating = true;
        }

        public void ExportToNewFile(string exportFile)
        {
            Workbook wb = ExcelApp.Workbooks.Open(exportFile);
            PriceList priceList = new PriceList(wb);

            if (priceList.IsValid())
                priceList.Fill(SkuAmount);
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

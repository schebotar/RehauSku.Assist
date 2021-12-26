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

                try
                {
                    PriceList priceList = new PriceList(wb);
                    SkuAmount.AddValues(priceList);
                }

                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show
                        ( $"{wb.Name} не является файлом прайс-листа \n\n {ex.Message}",
                        "Неверный файл прайс-листа!",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }

                finally
                {
                    wb.Close();
                }
            }
            ExcelApp.ScreenUpdating = true;
        }

        public void ExportToNewFile(string exportFile)
        {
            if (SkuAmount.Count < 1)
            {
                return;
            }

            Workbook wb = ExcelApp.Workbooks.Open(exportFile);
            PriceList priceList;

            try
            {
                priceList = new PriceList(wb);
                priceList.Fill(SkuAmount);
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show
                    ($"{RegistryUtil.PriceListPath} не является файлом прайс-листа \n\n {ex.Message}",
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

using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RehauSku.PriceListTools
{
    class CombineTool : ConjoinTool, IDisposable, IConjoinTool
    {
        private Dictionary<string, double>[] SkuAmount { get; set; }
        private string[] FileNames { get; set; }

        public CombineTool()
        {
            ExcelApp = (Application)ExcelDnaUtil.Application;
        }

        public void CollectSkuAmount(string[] files)
        {
            FileNames = files.Select(x => Path.GetFileNameWithoutExtension(x)).ToArray();
            SkuAmount = new Dictionary<string, double>[files.Length];

            ExcelApp.ScreenUpdating = false;

            for (int i = 0;  i < files.Length; i++)
            {
                Workbook wb = ExcelApp.Workbooks.Open(files[i]);

                try
                {
                    PriceList priceList = new PriceList(wb);
                    SkuAmount[i] = new Dictionary<string, double>();
                    SkuAmount[i].AddValuesFromPriceList(priceList);
                }

                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show
                        ($"{wb.Name} не является файлом прайс-листа \n\n {ex.Message}",
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

        public void ExportToFile(string exportFile)
        {
            if (SkuAmount.Sum(d => d.Count) < 1)
            {
                return;
            }

            Workbook wb = ExcelApp.Workbooks.Open(exportFile);
            PriceList priceList;

            try
            {
                priceList = new PriceList(wb);
                priceList.FillWithValues(SkuAmount, FileNames);
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
            //Dispose(true);
            GC.SuppressFinalize(this);
        }

        //protected virtual void Dispose(bool disposing)
        //{

        //}
    }
}

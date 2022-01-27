using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal static class SourceUtil
    {
        public static List<Source> GetSourceLists(string[] files)
        {
            var ExcelApp = (Application)ExcelDnaUtil.Application;

            List<Source> sourceFiles = new List<Source>();

            ExcelApp.ScreenUpdating = false;
            foreach (string file in files)
            {
                Workbook wb = ExcelApp.Workbooks.Open(file);
                try
                {
                    Source priceList = new Source(wb);
                    sourceFiles.Add(priceList);
                    wb.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show
                        (ex.Message,
                        "Ошибка открытия исходного прайс-листа",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    wb.Close();
                }
            }
            ExcelApp.ScreenUpdating = true;

            return sourceFiles;
        }
    }
}

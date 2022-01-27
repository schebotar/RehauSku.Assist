using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractPriceListTool
    {
        protected private Application ExcelApp;
        protected private SourceFile NewPriceList;
        protected private List<SourceFile> sourcePriceLists;

        public AbstractPriceListTool()
        {
            ExcelApp = (Application)ExcelDnaUtil.Application;
            sourcePriceLists = new List<SourceFile>();
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = NewPriceList.Sheet.AutoFilter;

            filter.Range.AutoFilter(NewPriceList.amountCell.Column, "<>");
            NewPriceList.Sheet.Range["A1"].Activate();
        }

        public void OpenNewPrice(string path)
        {
            Workbook wb = ExcelApp.Workbooks.Open(path);

            try
            {
                NewPriceList = new SourceFile(wb);
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show
                    (ex.Message,
                    "Ошибка открытия шаблонного прайс-листа",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                wb.Close();
            }
        }

        public virtual void GetSource()
        {
            throw new NotImplementedException();
        }

        public virtual void GetSource(string[] files)
        {
            ExcelApp.ScreenUpdating = false;
            foreach (string file in files)
            {
                Workbook wb = ExcelApp.Workbooks.Open(file);
                try
                {
                    SourceFile priceList = new SourceFile(wb);
                    sourcePriceLists.Add(priceList);
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
        }

        public virtual void FillPriceList()
        {
            throw new NotImplementedException();
        }
    }
}
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractPriceListTool
    {
        protected private Application ExcelApp;
        protected private PriceList NewPriceList;
        protected private List<PriceList> sourcePriceLists;

        public AbstractPriceListTool()
        {
            ExcelApp = (Application)ExcelDnaUtil.Application;
            sourcePriceLists = new List<PriceList>();
        }

        public void OpenNewPrice(string path)
        {
            Workbook wb = ExcelApp.Workbooks.Open(path);

            try
            {
                NewPriceList = new PriceList(wb);
            }

            catch (Exception ex)
            {
                wb.Close();
                throw ex;
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
                PriceList priceList = new PriceList(wb);
                sourcePriceLists.Add(priceList);
                wb.Close();
            }
            ExcelApp.ScreenUpdating = true;
        }

        public virtual void FillPriceList()
        {
            throw new NotImplementedException();
        }
    }
}
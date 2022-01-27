using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal abstract class PriceListTool : IDisposable
    {
        protected private Application ExcelApp;
        protected private Target NewPriceList;
        protected private List<Source> sourcePriceLists;

        public PriceListTool()
        {
            ExcelApp = (Application)ExcelDnaUtil.Application;
            sourcePriceLists = new List<Source>();
        }

        public void OpenNewPrice(string path)
        {
            Workbook wb = ExcelApp.Workbooks.Open(path);

            try
            {
                NewPriceList = new Target(wb);
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

        public virtual void GetSourceLists(string[] files)
        {
            ExcelApp.ScreenUpdating = false;
            foreach (string file in files)
            {
                Workbook wb = ExcelApp.Workbooks.Open(file);
                try
                {
                    Source priceList = new Source(wb);
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

        public virtual void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;
            FillAmountColumn();
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }

        protected private void FillAmountColumn()
        {
            int exportedValues = 0;
            foreach (var sheet in sourcePriceLists)
            {
                if (sheet.SkuAmount.Count == 0)
                    continue;

                foreach (var kvp in sheet.SkuAmount)
                {
                    Range cell = NewPriceList.Sheet.Columns[NewPriceList.skuCell.Column].Find(kvp.Key);

                    if (cell == null)
                    {
                        System.Windows.Forms.MessageBox.Show
                            ($"Артикул {kvp.Key} отсутствует в таблице заказов {RegistryUtil.PriceListPath}",
                            "Отсутствует позиция в конечной таблице заказов",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                    }

                    else
                    {
                        Range sumCell = NewPriceList.Sheet.Cells[cell.Row, NewPriceList.amountCell.Column];

                        if (sumCell.Value2 == null)
                            sumCell.Value2 = kvp.Value;
                        else
                            sumCell.Value2 += kvp.Value;

                        exportedValues++;
                    }
                }
            }
        }

        protected private void FilterByAmount()
        {
            AutoFilter filter = NewPriceList.Sheet.AutoFilter;

            filter.Range.AutoFilter(NewPriceList.amountCell.Column, "<>");
            NewPriceList.Sheet.Range["A1"].Activate();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}
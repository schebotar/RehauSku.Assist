﻿using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Rehau.Sku.Assist
{
    public class DataWriter : IDisposable
    {
        private Application xlApp;
        private Dictionary<string, double> SkuAmount { get; set; }
        private object[,] SelectedCells { get; set; }
        private string ExportFileName { get; set; }

        public DataWriter()
        {
            this.xlApp = (Application)ExcelDnaUtil.Application;
            this.ExportFileName = AddIn.priceListPath;

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

        public void FillPriceList()
        {
            Workbook wb = xlApp.Workbooks.Open(ExportFileName);
            Worksheet ws = wb.ActiveSheet;

            Range amountCell = ws.Cells.Find("Кол-во");

            foreach (KeyValuePair<string,double> kvp in SkuAmount)
            {
                Range cell = ws.Cells.Find(kvp.Key);
                ws.Cells[cell.Row, amountCell.Column].Value = kvp.Value;
            }

            //Range filter = ws.Cells[amountCell.Row + 1, amountCell.Column];
            //filter.AutoFilter(1, "<>");

            wb.Save();
            wb.Close();
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


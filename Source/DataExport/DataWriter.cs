﻿using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace RehauSku.Assist
{
    public class DataWriter : IDisposable
    {
        private Application xlApp;
        private Dictionary<string, double> SkuAmount { get; set; }
        private object[,] SelectedCells { get; set; }
        private string WorkingFileName { get; set; }

        public DataWriter()
        {
            this.xlApp = (Application)ExcelDnaUtil.Application;
            this.WorkingFileName = xlApp.ActiveWorkbook.FullName;

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

        //public void FillPriceList()
        //{
        //    string exportFileName = "rehau-export_" + DateTime.Now + ".xlsm";
        //    string workingDir = xlApp.ActiveWorkbook.Path;

        //    //File.Copy(Path.GetFullPath(PriceListFilePath), Path.Combine(WorkingFileName, exportFileName + ".xlsm"));


        //    Workbook wb = xlApp.Workbooks.Open(PriceListFilePath);
        //    Worksheet ws = wb.ActiveSheet;

        //    Range amountCell = ws.Cells.Find("Кол-во");

        //    foreach (KeyValuePair<string, double> kvp in SkuAmount)
        //    {
        //        Range cell = ws.Cells.Find(kvp.Key);
        //        ws.Cells[cell.Row, amountCell.Column].Value = kvp.Value;
        //    }

        //    //Range filter = ws.Range["H16:H4058"];
        //    ws.Cells.AutoFilter(7, "<>");

        //    //wb.Save();
        //    //wb.Close();
        //}

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


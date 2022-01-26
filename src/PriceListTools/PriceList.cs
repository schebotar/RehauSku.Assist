using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class PriceList
    {
        public readonly string Name;
        //public readonly PriceListSheet OfferSheet;
        public PriceListSheet Sheet { get; private set; }

        
        //private const string offerSheetHeader = "КП";

        public PriceList(Workbook workbook)
        {
            Name = workbook.Name;
            Sheet = new PriceListSheet(workbook.ActiveSheet);

            //foreach (Worksheet worksheet in workbook.Sheets)
            //{
            //    try
            //    {
            //        PriceListSheet priceListSheet = new PriceListSheet(worksheet);
            //        //priceListSheet.FillSkuAmount();
            //        Sheets.Add(priceListSheet);
            //    }
            //    catch (Exception ex)
            //    {
            //        throw ex;
            //    }
            //}



            //OfferSheet = Sheet.Where(s => s.Name == offerSheetHeader).FirstOrDefault();
        }
    }
}


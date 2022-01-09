using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class PriceList
    {
        public readonly string Name;
        public readonly PriceListSheet OfferSheet;
        public List<PriceListSheet> Sheets { get; private set; }

        private const string offerSheetHeader = "КП";

        public PriceList(Workbook workbook)
        {
            Name = workbook.Name;
            Sheets = new List<PriceListSheet>();

            foreach (Worksheet worksheet in workbook.Sheets)
            {
                PriceListSheet priceListSheet = new PriceListSheet(worksheet);

                if (priceListSheet.FillSkuAmount())
                    Sheets.Add(priceListSheet);
            }

            OfferSheet = Sheets.Where(s => s.Name == offerSheetHeader).FirstOrDefault();
        }

        public static string CreateNewFile()
        {
            string fileExtension = Path.GetExtension(RegistryUtil.PriceListPath);
            string path = Path.GetTempFileName() + fileExtension;

            File.Copy(RegistryUtil.PriceListPath, path);
            return path;
        }
    }
}


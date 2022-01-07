using System.Collections.Generic;
using System.IO;

namespace RehauSku.PriceListTools
{
    static class PriceListUtil
    {
        public static string CreateNewExportFile()
        {
            string fileExtension = Path.GetExtension(RegistryUtil.PriceListPath);
            string path = Path.GetTempFileName() + fileExtension;

            File.Copy(RegistryUtil.PriceListPath, path);
            return path;
        }

        public static void AddValuesFromPriceList(this Dictionary<string, double> SkuAmount, PriceList priceList)
        {
            object[,] amountCells = priceList.ActiveSheet.amountCells;
            object[,] skuCells = priceList.ActiveSheet.skuCells;

            for (int row = priceList.ActiveSheet.headerRow.Value + 1; row < amountCells.GetLength(0); row++)
            {
                object amount = amountCells[row, 1];
                object sku = skuCells[row, 1];

                if (amount != null && (double)amount != 0)
                {
                    if (SkuAmount.ContainsKey(sku.ToString()))
                    {
                        SkuAmount[sku.ToString()] += (double)amount;
                    }

                    else
                        SkuAmount.Add(sku.ToString(), (double)amount);
                }
            }
        }
    }
}


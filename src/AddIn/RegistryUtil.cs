using Microsoft.Win32;
using System.IO;
using RehauSku.Forms;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace RehauSku
{
    static class RegistryUtil
    {
        private static string priceListPath;
        private static int? storeResponseOrder;
        private static RegistryKey RootKey { get; set; }

        public static void Initialize()
        {
            RootKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\REHAU\SkuAssist"); 
            priceListPath = RootKey.GetValue("PriceListPath") as string;
            storeResponseOrder = RootKey.GetValue("StoreResponseOrder") as int?;
        }

        public static void Uninitialize()
        {
            RootKey.Close();            
        }

        public static bool IsPriceListPathEmpty()
        {
            return string.IsNullOrEmpty(priceListPath);
        }

        public static string PriceListPath
        {
            get
            {
                if (IsPriceListPathEmpty() || !File.Exists(priceListPath))
                {
                    //MessageBox.Show("Прайс-лист отсутствует или неверный файл прайс-листа", "Укажите файл прайс-листа", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string fileName = Dialog.GetFilePath();
                    priceListPath = fileName;
                    RootKey.SetValue("PriceListPath", fileName);
                    return priceListPath;
                }

                else
                {
                    return priceListPath;
                }
            }

            set
            {
                priceListPath = value;
                RootKey.SetValue("PriceListPath", value);
            }
        }

        public static ResponseOrder StoreResponseOrder
        {
            get
            {
                if (storeResponseOrder == null)
                {
                    RootKey.SetValue("StoreResponseOrder", (int)ResponseOrder.Default);
                    storeResponseOrder = (int)ResponseOrder.Default;
                    return (ResponseOrder)storeResponseOrder.Value;
                }

                else
                {
                    return (ResponseOrder)storeResponseOrder.Value;
                }
            }
        }
    }
}

using Microsoft.Win32;
using System.IO;
using RehauSku.Forms;
using System.Windows.Forms;

namespace RehauSku
{
    static class RegistryUtil
    {
        private static string _priceListPath;
        private static int? _storeResponseOrder;
        private static RegistryKey _RootKey { get; set; }

        public static void Initialize()
        {
            _RootKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\REHAU\SkuAssist"); 
            _priceListPath = _RootKey.GetValue("PriceListPath") as string;
            _storeResponseOrder = _RootKey.GetValue("StoreResponseOrder") as int?;
        }

        public static void Uninitialize()
        {
            _RootKey.Close();
            
        }

        public static bool IsPriceListPathEmpty()
        {
            return string.IsNullOrEmpty(_priceListPath);
        }

        public static string PriceListPath
        {
            get
            {
                if (IsPriceListPathEmpty() || !File.Exists(_priceListPath))
                {
                    MessageBox.Show("Прайс-лист отсутствует или неверный файл прайс-листа", "Укажите файл прайс-листа", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string fileName = Dialog.GetFilePath();
                    _priceListPath = fileName;
                    _RootKey.SetValue("PriceListPath", fileName);
                    return _priceListPath;
                }

                else
                {
                    return _priceListPath;
                }
            }

            set
            {
                _priceListPath = value;
                _RootKey.SetValue("PriceListPath", value);
            }
        }

        public static ResponseOrder StoreResponseOrder
        {
            get
            {
                if (_storeResponseOrder == null)
                {
                    _RootKey.SetValue("StoreResponseOrder", (int)ResponseOrder.Default);
                    _storeResponseOrder = (int)ResponseOrder.Default;
                    return (ResponseOrder)_storeResponseOrder.Value;
                }

                else
                {
                    return (ResponseOrder)_storeResponseOrder.Value;
                }
            }
        }
    }
}

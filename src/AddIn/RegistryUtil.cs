using Microsoft.Win32;
using RehauSku.Interface;
using System;
using System.IO;
using System.Windows.Forms;

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

        public static string PriceListPath
        {
            get
            {
                if (string.IsNullOrEmpty(priceListPath) || !File.Exists(priceListPath))
                {
                    DialogResult result = MessageBox.Show("Прайс-лист отсутствует или неверный файл шаблона прайс-листа. " +
                        "Укажите файл шаблона прайс-листа.",
                        "Нет файла шаблона",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    if (result == DialogResult.OK)
                    {
                        string fileName = Dialog.GetFilePath();

                        if (string.IsNullOrEmpty(fileName))
                        {
                            throw new Exception("Нет файла шаблона");
                        }

                        priceListPath = fileName;
                        RootKey.SetValue("PriceListPath", fileName);
                        return priceListPath;
                    }

                    else
                        throw new Exception("Нет файла шаблона");
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
                RibbonController.RefreshControl("setPriceList");
            }
        }

        public static string GetPriceListName()
        {
            return Path.GetFileName(priceListPath);
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

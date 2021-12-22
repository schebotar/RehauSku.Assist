using Microsoft.Win32;

namespace RehauSku
{
    static class RegistryUtil
    {
        public static string PriceListPath
        {
            get => (string)_RootKey.GetValue("PriceListPath");
        }

        public static ResponseOrder StoreResponseOrder
        {
            get => (ResponseOrder)_RootKey.GetValue("StoreResponseOrder");
        }

        private static RegistryKey _RootKey
        {
            get
            {
                return _OpenRootKey() ?? _CreateRootKey();
            }
        }

        private static RegistryKey _OpenRootKey()
        {
            return Registry.CurrentUser
                .OpenSubKey(@"SOFTWARE\REHAU\SkuAssist");
        }

        private static RegistryKey _CreateRootKey()
        {
            RegistryKey key = Registry.CurrentUser
                .CreateSubKey(@"SOFTWARE\REHAU\SkuAssist");

            key.SetValue("PriceListPath", @"D:\Dropbox\Рабочее\Таблица заказов ИС EAE_2021.xlsm");
            key.SetValue("StoreResponseOrder", 0);

            return key;
        }
    }
}

using Microsoft.Win32;

namespace RehauSku
{
    static class RegistryUtil
    {
        public static string PriceListPath
        {
            get
            {
                _GetRootKey();

                if (_RootKey == null)
                {
                    return @"D:\Dropbox\Рабочее\Таблица заказов ИС EAE_2021.xlsm";
                }

                else return (string)_RootKey.GetValue("PriceListPath");
            }

            private set
            {
                _GetRootKey();

                if (_RootKey == null)
                {
                    RegistryKey PriceListPath = Registry.CurrentUser
                        .CreateSubKey("SOFTWARE")
                        .CreateSubKey("REHAU")
                        .CreateSubKey("SkuAssist");
                }

                _RootKey.SetValue("PriceListPath", value);
            }
        }

        public static ResponseOrder StoreResponseOrder
        {
            get
            {
                _GetRootKey();

                if (_RootKey == null)
                {
                    return ResponseOrder.Default;
                }

                return (ResponseOrder)_RootKey.GetValue("ResponseOrder");
            }

            private set
            {
                if (_RootKey == null)
                {
                    RegistryKey PriceListPath = Registry.CurrentUser
                        .CreateSubKey("SOFTWARE")
                        .CreateSubKey("REHAU")
                        .CreateSubKey("SkuAssist");
                }

                _RootKey.SetValue("ResponseOrder", value);
            }
        }

        private static RegistryKey _RootKey { get; set; }

        private static void _GetRootKey()
        {
            _RootKey = Registry
                .CurrentUser
                .OpenSubKey("SOFTWARE")
                .OpenSubKey("REHAU")
                .OpenSubKey("SkuAssist");
        }
    }
}

using RehauSku.Assistant;
using System;

namespace RehauSku.PriceListTools
{
    internal class PriceListPosition
    {
        public readonly string Group;
        public readonly string Sku;

        public PriceListPosition(string group, string sku)
        {
            if (!sku.IsRehauSku())
                throw new ArgumentException("Wrong SKU");

            else
            {
                Group = group;
                Sku = sku;
            }
        }
    }
}

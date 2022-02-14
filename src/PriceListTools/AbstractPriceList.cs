using Microsoft.Office.Interop.Excel;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractPriceList
    {
        public Range AmountCell { get; protected set; }
        public Range SkuCell { get; protected set; }
        public Range GroupCell { get; protected set; }
        public Range NameCell { get; protected set; }

        public Worksheet Sheet { get; protected set; }
        public string Name { get; protected set; }
    }
}
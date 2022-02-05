using Microsoft.Office.Interop.Excel;

namespace RehauSku.PriceListTools
{
    internal abstract class AbstractPriceList
    {
        protected const string amountHeader = "Кол-во";
        protected const string skuHeader = "Актуальный материал";
        protected const string groupHeader = "Программа";
        protected const string nameHeader = "Наименование";

        public Range amountCell { get; protected set; }
        public Range skuCell { get; protected set; }
        public Range groupCell { get; protected set; }
        public Range nameCell { get; protected set; }

        public Worksheet Sheet { get; protected set; }
        public string Name { get; protected set; }
    }
}
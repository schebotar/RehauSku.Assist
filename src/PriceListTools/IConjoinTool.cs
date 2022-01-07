namespace RehauSku.PriceListTools
{
    internal interface IConjoinTool
    {
        void CollectSkuAmount(string[] files);
        void ExportToFile(string exportFile);
    }
}
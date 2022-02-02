using RehauSku.Interface;
using System.Collections.Generic;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : PriceListTool
    {
        public List<Source> SourceFiles;

        public void FillTarget()
        {
            foreach (Source source in SourceFiles)
            {
                foreach (var kvp in source.PositionAmount)
                    FillColumnsWithDictionary(kvp, TargetFile.amountCell.Column);
            }

            FilterByAmount();

            Dialog.SaveWorkbookAs();
        }
    }
}

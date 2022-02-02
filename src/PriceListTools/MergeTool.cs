using RehauSku.Interface;
using System.Collections.Generic;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : AbstractTool
    {
        public List<Source> SourceFiles;

        public void FillTarget()
        {
            ProgressBar bar = new ProgressBar(SourceFiles.Sum(x => x.PositionAmount.Count));

            foreach (Source source in SourceFiles)
            {
                foreach (var kvp in source.PositionAmount)
                {
                    FillColumnsWithDictionary(kvp, TargetFile.amountCell.Column);
                    bar.DoProgress();
                }
            }

            FilterByAmount();

            Dialog.SaveWorkbookAs();
        }
    }
}

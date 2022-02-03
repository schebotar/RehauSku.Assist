using RehauSku.Interface;
using System.Collections.Generic;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : AbstractTool
    {
        public List<SourcePriceList> SourceFiles;

        public void FillTarget()
        {
            ProgressBar bar = new ProgressBar("Заполняю строки...", SourceFiles.Sum(x => x.PositionAmount.Count));

            foreach (SourcePriceList source in SourceFiles)
            {
                foreach (var kvp in source.PositionAmount)
                {
                    FillPositionAmountToColumns(kvp, TargetFile.amountCell.Column);
                    bar.DoProgress();
                }
            }

            FilterByAmount();

            Dialog.SaveWorkbookAs();
        }
    }
}

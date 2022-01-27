using System.Collections.Generic;
using System.Linq;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : PriceListTool
    {
        public List<Source> SourceFiles;

        public void FillTarget()
        {
            ExcelApp.ScreenUpdating = false;
            FillAmountColumn(SourceFiles.Select(x => x.SkuAmount).ToArray());
            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}

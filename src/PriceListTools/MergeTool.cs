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

            foreach (Source source in SourceFiles)
            {
                FillColumn(source.PositionAmount, TargetFile.amountCell.Column);
            }

            FilterByAmount();
            ExcelApp.ScreenUpdating = true;

            Forms.Dialog.SaveWorkbookAs();
        }
    }
}

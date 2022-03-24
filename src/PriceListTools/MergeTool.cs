using RehauSku.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RehauSku.PriceListTools
{
    internal class MergeTool : AbstractTool
    {
        private List<SourcePriceList> SourceFiles { get; set; }

        public MergeTool()
        {
            string[] files = Dialog.GetMultiplyFiles();

            if (files != null)
            {
                SourceFiles = SourcePriceList.GetSourceLists(files);
            }

            else
            {
                throw new Exception("Не выбраны файлы");
            }
        }

        public override async void FillTarget()
        {
            ProgressBar = new ProgressBar("Заполняю строки...", SourceFiles.Sum(x => x.PositionAmount.Count));
            ResultBar = new ResultBar();

            foreach (SourcePriceList source in SourceFiles)
            {
                foreach (var kvp in source.PositionAmount)
                {
                    FillPositionAmountToColumns(kvp, TargetFile.AmountCell.Column);
                    ProgressBar.Update();
                }
            }

            FilterByAmount();
            ResultBar.Update();

            await Task.Delay(new TimeSpan(0, 0, 5));
            ExcelApp.StatusBar = false;
        }
    }
}

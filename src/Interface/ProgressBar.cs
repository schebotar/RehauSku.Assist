using Microsoft.Office.Interop.Excel;

namespace RehauSku.Interface
{
    internal class ProgressBar
    {
        private Application Excel = AddIn.Excel;
        private double CurrentProgress { get; set; }
        private readonly double TaskWeight;
        private readonly string Message;

        public ProgressBar(string message, int weight)
        {
            Message = message;
            TaskWeight = weight;
            CurrentProgress = 0;
        }

        public void DoProgress()
        {
            double percent = (++CurrentProgress / TaskWeight) * 100;

            if (percent < 100)
            {
                Excel.StatusBar = $"{Message} Выполнено {percent.ToString("#.##")} %";
            }

            else
            {
                Excel.StatusBar = false;
            }
        }
    }
}

using System;
using System.Text;

namespace RehauSku.Interface
{
    internal class ResultBar : AbstractBar
    {
        private int Success { get; set; }
        private int Replaced { get; set; }
        private int NotFound { get; set; }

        public ResultBar()
        {
            Success = 0;
            Replaced = 0;
            NotFound = 0;
        }

        public void IncrementSuccess() => Success++;
        public void IncrementReplaced() => Replaced++;
        public void IncrementNotFound() => NotFound++;

        public override void Update()
        {
            StringBuilder sb = new StringBuilder();

            if (Success > 0)
            {
                sb.Append($"Успешно экспортировано {Success} артикулов. ");
            }

            if (Replaced > 0)
            {
                sb.Append($"Заменено {Replaced} артикулов. ");
            }

            if (NotFound > 0)
            {
                sb.Append($"Не найдено {NotFound} артикулов.");
            }

            Excel.StatusBar = sb.ToString(); 
            AddIn.Excel.OnTime(DateTime.Now + new TimeSpan(0, 0, 5), "ResetStatusBar");
        }
    }
}

using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;

namespace RehauSku.Interface
{
    internal abstract class AbstractBar : IDisposable
    {
        protected Application Excel = AddIn.Excel;

        public abstract void Update();

        [ExcelFunction]
        public static void ResetStatusBar()
        {
            AddIn.Excel.StatusBar = false;
        }

        public void Dispose()
        {
            AddIn.Excel.OnTime(DateTime.Now + new TimeSpan(0, 0, 5), "ResetStatusBar");
        }
    }
}

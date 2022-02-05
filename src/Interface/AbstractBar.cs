using Microsoft.Office.Interop.Excel;

namespace RehauSku.Interface
{
    internal abstract class AbstractBar
    {
        protected Application Excel = AddIn.Excel;

        public abstract void Update();
    }
}

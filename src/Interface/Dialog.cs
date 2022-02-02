using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RehauSku.Interface
{
    static class Dialog
    {
        public static string GetFilePath()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Файлы Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }

            return string.Empty;
        }

        public static string[] GetMultiplyFiles()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Файлы Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileNames;
                }

                else return null;
            }
        }

        public static void SaveWorkbookAs()
        {
            Workbook workbook = AddIn.Excel.ActiveWorkbook;

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.FileName = workbook.Name;
                dialog.Filter = "Файлы Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";

                if (dialog.ShowDialog() == DialogResult.Cancel)
                {
                    workbook.Close(false);
                }

                else
                {
                    string fileName = dialog.FileName;
                    workbook.SaveAs(fileName);
                }
            }
        }
    }
}

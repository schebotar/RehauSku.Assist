using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RehauSku.Interface
{
    static class Dialog
    {
        public static string GetFilePath()
        {
            string filePath = string.Empty;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Файлы Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = dialog.FileName;
                }
            }

            return filePath;
        }

        public static string[] GetMultiplyFiles()
        {
            List<string> fileNames = new List<string>();

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Файлы Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in dialog.FileNames)
                    {
                        fileNames.Add(file);
                    }
                }
            }

            return fileNames.ToArray();
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

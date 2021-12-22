using System.Windows.Forms;

namespace RehauSku
{
    static class FileDialog
    {
        public static string GetFilePath()
        {
            string filePath = string.Empty;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Все файлы (*.*)|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = dialog.FileName;
                }
            }

            return filePath;
        }
    }
}

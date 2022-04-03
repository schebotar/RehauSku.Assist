using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using RehauSku.PriceListTools;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace RehauSku.Interface
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private static IRibbonUI ribbonUi;

        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI onLoad='RibbonLoad' xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='rau' label='REHAU'>
            <group id='priceList' label='Прайс-лист'>
                <button id='export' getEnabled='GetExportEnabled' label='Экспорт в новый файл' size='normal' imageMso='PivotExportToExcel' onAction='OnToolPressed'/> 
                <button id='convert' getEnabled='GetConvertEnabled' label='Актуализировать' size='normal' imageMso='FileUpdate' onAction='OnToolPressed'/> 
                <menu id='conjoinMenu' label='Объединить' imageMso='Copy'>
                    <button id='merge' label='Сложить' onAction='OnToolPressed'/>    
                    <button id='combine' label='По колонкам' onAction='OnToolPressed'/>   
                </menu>
            </group>
            <group id='rausettings' getLabel='GetVersionLabel'>
                <button id='setPriceList' getLabel='GetPriceListPathLabel' size='large' imageMso='TableExcelSpreadsheetInsert' onAction='OnSetPricePressed'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void RibbonLoad(IRibbonUI sender)
        {
            ribbonUi = sender;
        }

        public static void RefreshControl(string id)
        {
            if (ribbonUi != null)
            {
                ribbonUi.InvalidateControl(id);
            }
        }
        public void OnSetPricePressed(IRibbonControl control)
        {
            string path = Dialog.GetFilePath();

            if (!string.IsNullOrEmpty(path))
            {
                RegistryUtil.PriceListPath = path;
            }
        }

        public void OnToolPressed(IRibbonControl control)
        {
            try
            {
                AbstractTool tool;
                switch (control.Id)
                {
                    case "export":
                        tool = new ExportTool();
                        break;
                    case "convert":
                        tool = new ConvertTool();
                        break;
                    case "merge":
                        tool = new MergeTool();
                        break;
                    case "combine":
                        tool = new CombineTool();
                        break;
                    default:
                        throw new Exception("Неизвестный инструмент");
                }

                tool.OpenNewPrice();
                tool.FillTarget();
            }

            catch (Exception exception)
            {
                MessageBox.Show(exception.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                AddIn.Excel.StatusBar = false;
                return;
            }
        }

        public bool GetConvertEnabled(IRibbonControl control)
        {
            if (AddIn.Excel.ActiveWorkbook == null)
                return false;

            else
            {
                Worksheet worksheet = AddIn.Excel.ActiveWorkbook.ActiveSheet;
                return worksheet.IsRehauSource();
            }
        }

        public bool GetExportEnabled(IRibbonControl control)
        {
            if (AddIn.Excel.ActiveWorkbook == null)
                return false;

            else
            {
                Range selection = AddIn.Excel.Selection;
                return selection.Columns.Count == 2;
            }
        }

        public string GetVersionLabel(IRibbonControl control)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            return $"v{version}";
        }

        public string GetPriceListPathLabel(IRibbonControl control)
        {
            string name = RegistryUtil.GetPriceListName();
            return string.IsNullOrEmpty(name) ? "Нет файла шаблона!" : name;
        }
    }
}

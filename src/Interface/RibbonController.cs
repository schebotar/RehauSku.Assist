using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using RehauSku.PriceListTools;
using System;
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
                <button id='exportToPrice' getEnabled='GetExportEnabled' label='Экспорт в новый файл' size='normal' imageMso='PivotExportToExcel' onAction='OnToolPressed'/> 
                <button id='convertPrice' getEnabled='GetConvertEnabled' label='Актуализировать' size='normal' imageMso='FileUpdate' onAction='OnToolPressed'/> 
                <menu id='conjoinMenu' label='Объединить' imageMso='Copy'>
                    <button id='mergeFiles' label='Сложить' onAction='OnToolPressed'/>    
                    <button id='combineFiles' label='По колонкам' onAction='OnToolPressed'/>   
                </menu>
            </group>
            <group id='rausettings' label='Настройки'>
                <button id='setPriceList' label='Указать путь к шаблону' size='large' imageMso='CurrentViewSettings' onAction='OnSetPricePressed'/>
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
                    case "exportToPrice":
                        tool = new ExportTool();
                        break;
                    case "convertPrice":
                        tool = new ConvertTool();
                        break;
                    case "mergeFiles":
                        tool = new MergeTool();
                        break;
                    case "combineFiles":
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
    }
}

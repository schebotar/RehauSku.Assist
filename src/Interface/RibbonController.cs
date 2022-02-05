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
                <button id='exportToPrice' getEnabled='GetExportEnabled' label='Экспорт в новый файл' size='normal' imageMso='PivotExportToExcel' onAction='OnExportPressed'/> 
                <button id='convertPrice' getEnabled='GetConvertEnabled' label='Актуализировать' size='normal' imageMso='FileUpdate' onAction='OnConvertPressed'/> 
                <menu id='conjoinMenu' label='Объединить' imageMso='Copy'>
                    <button id='mergeFiles' label='Сложить' onAction='OnMergePressed'/>    
                    <button id='combineFiles' label='По колонкам' onAction='OnCombinePressed'/>   
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

        public void OnMergePressed(IRibbonControl control)
        {
            MergeTool mergeTool = new MergeTool();
            string[] files = Dialog.GetMultiplyFiles();

            if (files != null)
            {
                mergeTool.SourceFiles = SourcePriceList.GetSourceLists(files);
                mergeTool.OpenNewPrice();
                mergeTool.FillTarget();
            }
        }

        public void OnCombinePressed(IRibbonControl control)
        {
            CombineTool combineTool = new CombineTool();
            string[] files = Dialog.GetMultiplyFiles();

            if (files != null)
            {
                combineTool.SourceFiles = SourcePriceList.GetSourceLists(files);
                combineTool.OpenNewPrice();
                combineTool.FillTarget();
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

        public void OnExportPressed(IRibbonControl control)
        {
            try
            {
                ExportTool exportTool = new ExportTool();
                exportTool.OpenNewPrice();
                exportTool.FillTarget();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
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

        public void OnConvertPressed(IRibbonControl control)
        {
            ConvertTool convertTool = new ConvertTool();

            convertTool.GetCurrent();
            convertTool.OpenNewPrice();
            convertTool.FillTarget();
        }

        public void OnSetPricePressed(IRibbonControl control)
        {
            string path = Dialog.GetFilePath();

            if (!string.IsNullOrEmpty(path))
            {
                RegistryUtil.PriceListPath = path;
            }
        }
    }
}

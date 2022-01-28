using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using RehauSku.PriceListTools;
using RehauSku.Forms;
using System;
using System.Collections.Generic;

namespace RehauSku.Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='rau' label='REHAU'>
            <group id='priceList' label='Прайс-лист'>
                <button id='exportToPrice' label='Экспорт в новый файл' size='normal' imageMso='PivotExportToExcel' onAction='OnExportPressed'/> 
                <button id='convertPrice' label='Актуализировать' size='normal' imageMso='FileUpdate' onAction='OnConvertPressed'/> 
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

        public void OnMergePressed(IRibbonControl control)
        {
            MergeTool mergeTool = new MergeTool();
            string[] files = Dialog.GetMultiplyFiles();

            if (files.Length != 0)
            {
                mergeTool.SourceFiles = SourceUtil.GetSourceLists(files);
                mergeTool.OpenNewPrice();
                mergeTool.FillTarget();
            }
        }

        public void OnCombinePressed(IRibbonControl control)
        {
            CombineTool combineTool = new CombineTool();
            string[] files = Dialog.GetMultiplyFiles();

            if (files.Length != 0)
            {
                combineTool.SourceFiles = SourceUtil.GetSourceLists(files);
                combineTool.OpenNewPrice();
                combineTool.FillTarget();
            }
        }

        public void OnExportPressed(IRibbonControl control)
        {
            try
            {
                ExportTool exportTool = new ExportTool();
                exportTool.TryGetSelection();
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
            RegistryUtil.PriceListPath = path;
        }
    }
}

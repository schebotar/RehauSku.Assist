using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using RehauSku.PriceListTools;
using RehauSku.Forms;
using System;

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
                <menu id='conjoinMenu' label='Объединить' imageMso='Copy'>
                    <button id='mergeFiles' label='Сложить' onAction='OnMergePressed'/>    
                    <button id='combineFiles' label='По колонкам' onAction='OnCombinePressed'/>   
                </menu>
            </group>
            <group id='rausettings' label='Настройки'>
                <button id='setPriceList' label='Файл прайс-листа' size='large' imageMso='CurrentViewSettings' onAction='OnSetPricePressed'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        // <dropDown id = 'dd1' label = 'Drop dynamic' getItemCount = 'fncGetItemCountDrop' getItemLabel = 'fncGetItemLabelDrop' onAction = 'fncOnActionDrop'/>

        public void OnMergePressed(IRibbonControl control)
        {
            using (MergeTool mergeTool = new MergeTool())
            {
                string[] files = Dialog.GetMultiplyFiles();
                if (files.Length != 0)
                {
                    mergeTool.GetSourceLists(files);
                    string exportFile = RegistryUtil.PriceListPath;
                    mergeTool.OpenNewPrice(exportFile);
                    mergeTool.FillTarget();
                }
            }
        }

        public void OnCombinePressed(IRibbonControl control)
        {
            using (CombineTool combineTool = new CombineTool())
            {
                string[] files = Dialog.GetMultiplyFiles();
                if (files.Length != 0)
                {
                    combineTool.GetSourceLists(files);
                    string exportFile = RegistryUtil.PriceListPath;
                    combineTool.OpenNewPrice(exportFile);
                    combineTool.FillTarget();
                }
            }
        }

        public void OnExportPressed(IRibbonControl control)
        {
            try
            {
                using (ExportTool exportTool = new ExportTool())
                {
                    exportTool.GetSource();
                    string exportFile = RegistryUtil.PriceListPath;
                    exportTool.OpenNewPrice(exportFile);
                    exportTool.FillTarget();
                }
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

        public void OnSetPricePressed(IRibbonControl control)
        {
            string path = Dialog.GetFilePath();
            RegistryUtil.PriceListPath = path;
        }
    }
}

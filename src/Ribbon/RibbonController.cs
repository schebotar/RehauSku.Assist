﻿using System.Runtime.InteropServices;
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
                mergeTool.GetSource(files);
                string exportFile = PriceList.CreateNewFile();
                mergeTool.OpenNewPrice(exportFile);
                mergeTool.FillPriceList();
            }
        }

        public void OnCombinePressed(IRibbonControl control)
        {
            using (CombineTool combineTool = new CombineTool())
            {
                string[] files = Dialog.GetMultiplyFiles();
                combineTool.GetSource(files);
                string exportFile = PriceList.CreateNewFile();
                combineTool.OpenNewPrice(exportFile);
                combineTool.FillPriceList();
            }
        }

        public void OnExportPressed(IRibbonControl control)
        {
            try
            {
                using (ExportTool exportTool = new ExportTool())
                {
                    exportTool.GetSource(null);
                    string exportFile = PriceList.CreateNewFile();
                    exportTool.OpenNewPrice(exportFile);
                    exportTool.FillPriceList();
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

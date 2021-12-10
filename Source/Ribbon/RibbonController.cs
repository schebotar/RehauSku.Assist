﻿using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using RehauSku.DataExport;

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
                <button id='exportToPrice' label='Экспорт' size='large' imageMso='PivotExportToExcel' onAction='OnButtonPressed'/>
            </group>
            <group id='rausettings' label='Настройки'>
                <button id='set' label='Настройки' size='large' imageMso='CurrentViewSettings' onAction='OnButtonPressed'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            using (Exporter dw = new Exporter())
            {
                if (!dw.IsRangeValid())
                {
                    MessageBox.Show("Выделен неверный диапазон!", 
                        "Неверный диапазон", 
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                else
                {
                    dw.FillSkuAmountDict();
                    dw.FillPriceList();
                }
            }
        }
    }
}

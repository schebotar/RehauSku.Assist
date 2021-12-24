using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using RehauSku.PriceListTools;
using RehauSku.Forms;

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
                <button id='exportToPrice' label='Экспорт' size='large' imageMso='PivotExportToExcel' onAction='OnExportPressed'/>
                <button id='mergeFiles' label='Объединить' size='large' imageMso='Copy' onAction='OnMergePressed'/>                
            </group>
            <group id='rausettings' label='Настройки'>
                <button id='set' label='Настройки' size='large' imageMso='CurrentViewSettings' onAction='OnSettingsPressed'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnMergePressed(IRibbonControl control)
        {
            using (MergeTool mergeTool = new MergeTool())
            {
                string[] files = Dialog.GetMultiplyFiles();
                mergeTool.AddSkuAmountToDict(files);
                string exportFile = PriceListUtil.CreateNewExportFile();
                mergeTool.ExportToNewFile(exportFile);
            }
        }

        public void OnExportPressed(IRibbonControl control)
        {
            using (ExportTool exportTool = new ExportTool())
            {
                if (!exportTool.IsRangeValid())
                {
                    MessageBox.Show("Выделен неверный диапазон!",
                        "Неверный диапазон",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                else
                {
                    exportTool.ExportToNewFile();
                }
            }
        }

        public void OnSettingsPressed(IRibbonControl control)
        {
            Form settingsForm = new SettingsForm();

            settingsForm.Show();

        }
    }
}

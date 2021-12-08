using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using Rehau.Sku.Assist;

namespace Ribbon
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
          <tab id='tab1' label='REHAU'>
            <group id='group1' label='Прайс-лист'>
              <button id='button1' label='Экспорт' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            using (DataWriter dw = new DataWriter())
            {
                if (!dw.IsRangeValid())
                {
                    MessageBox.Show("Выделен неверный диапазон!", "Неверный диапазон", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                else
                {
                    dw.FillSkuAmountDict();
                    //dw.FillPriceList();
                }
            }
        }
    }
}

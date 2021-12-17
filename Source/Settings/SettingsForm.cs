using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RehauSku.Settings
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();

            FormClosing += (sender, eventArgs) =>
            {
                MessageBox.Show("ok");
            };


        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}

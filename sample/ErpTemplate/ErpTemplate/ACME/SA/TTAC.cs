using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class TTAC : Form
    {
        public TTAC()
        {
            InitializeComponent();
        }

        private void sATT_ACCBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sATT_ACCBindingSource.EndEdit();
            this.sATT_ACCTableAdapter.Update(this.sa.SATT_ACC);

        }

        private void TTAC_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'sa.SATT_ACC' 資料表。您可以視需要進行移動或移除。
            this.sATT_ACCTableAdapter.Fill(this.sa.SATT_ACC);

        }

        private void sATT_ACCDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
    }
}

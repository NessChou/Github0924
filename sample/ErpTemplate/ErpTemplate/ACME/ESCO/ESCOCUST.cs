using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class ESCOCUST : Form
    {
        public ESCOCUST()
        {
            InitializeComponent();
        }

        private void eSCO_CUSTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.eSCO_CUSTBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.eSCO);

        }

        private void eSCO_CUSTBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.eSCO_CUSTBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.eSCO);

        }

        private void ESCOCUST_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'eSCO.ESCO_CUST' 資料表。您可以視需要進行移動或移除。
            this.eSCO_CUSTTableAdapter.Fill(this.eSCO.ESCO_CUST);

        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class WH_ITEM1 : Form
    {
        public WH_ITEM1()
        {
            InitializeComponent();
        }

        private void wH_ITM1BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.wH_ITM1BindingSource.EndEdit();
            this.wH_ITM1TableAdapter.Update(this.wh.WH_ITM1);
            MessageBox.Show("存檔完成");

        }

        private void WH_ITEM1_Load(object sender, EventArgs e)
        {

            this.wH_ITM1TableAdapter.Fill(this.wh.WH_ITM1);

        }
    }
}

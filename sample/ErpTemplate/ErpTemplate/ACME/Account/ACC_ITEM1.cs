using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class ACC_ITEM1 : Form
    {
        public ACC_ITEM1()
        {
            InitializeComponent();
        }

        private void account_ITEM1BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.account_ITEM1BindingSource.EndEdit();
            this.account_ITEM1TableAdapter.Update(this.accBank.Account_ITEM1);

            MessageBox.Show("存檔成功");

        }

        private void ACC_ITEM1_Load(object sender, EventArgs e)
        {
            this.account_ITEM1TableAdapter.Fill(this.accBank.Account_ITEM1);

        }
    }
}

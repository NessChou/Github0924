using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class CheckMoney2 : Form
    {
        public CheckMoney2()
        {
            InitializeComponent();
        }

        private void account_JEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.account_JEBindingSource.EndEdit();
            this.account_JETableAdapter.Update(this.accBank.Account_JE);

            MessageBox.Show("儲存成功");


        }

        private void CheckMoney2_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'accBank.Account_JE' 資料表。您可以視需要進行移動或移除。
            this.account_JETableAdapter.Fill(this.accBank.Account_JE);

        
        }
    }
}
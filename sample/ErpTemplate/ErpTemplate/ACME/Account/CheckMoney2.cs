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

            MessageBox.Show("�x�s���\");


        }

        private void CheckMoney2_Load(object sender, EventArgs e)
        {
            // TODO: �o��{���X�|�N��Ƹ��J 'accBank.Account_JE' ��ƪ�C�z�i�H���ݭn�i�沾�ʩβ����C
            this.account_JETableAdapter.Fill(this.accBank.Account_JE);

        
        }
    }
}
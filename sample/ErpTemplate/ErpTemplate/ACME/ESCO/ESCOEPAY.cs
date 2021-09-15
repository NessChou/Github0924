using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class ESCOEPAY : Form
    {
        public ESCOEPAY()
        {
            InitializeComponent();
        }

        private void eSCO_PAYBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.eSCO_PAYBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.eSCO);

            MessageBox.Show("更新成功");

        }

        private void ESCOEPAY_Load(object sender, EventArgs e)
        {
     
            this.eSCO_PAYTableAdapter.Fill(this.eSCO.ESCO_PAY);

        }
    }
}

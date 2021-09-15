using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class GB_ARLOCK : Form
    {
        public GB_ARLOCK()
        {
            InitializeComponent();
        }

        private void gB_DATELOCKBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            dOCDATETextBox.Text = GetMenu.Day();
            lOGUSERTextBox.Text = fmLogin.LoginID.ToString();

            this.Validate();
            this.gB_DATELOCKBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);

            MessageBox.Show("已存檔");

        }

        private void GB_ARLOCK_Load(object sender, EventArgs e)
        {
            this.gB_DATELOCKTableAdapter.Fill(this.pOTATO.GB_DATELOCK);

   
        }
    }
}

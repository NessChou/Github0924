using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class GB_FRIGHT : Form
    {
        public GB_FRIGHT()
        {
            InitializeComponent();
        }

        private void gB_FRIGHTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_FRIGHTBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);

            MessageBox.Show("更新成功");

        }


        private void GB_FRIGHT_Load(object sender, EventArgs e)
        {
            try
            {
                this.gB_FRIGHTTableAdapter.Fill(this.pOTATO.GB_FRIGHT);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void gB_FRIGHTDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

 
    }
}

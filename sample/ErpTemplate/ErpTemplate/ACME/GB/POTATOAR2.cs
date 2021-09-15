using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class POTATOAR2 : Form
    {
        public POTATOAR2()
        {
            InitializeComponent();
        }

        private void gB_INVTRACKBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_INVTRACKBindingSource.EndEdit();
            this.gB_INVTRACKTableAdapter.Update(this.POTATO.GB_INVTRACK);

            MessageBox.Show("Àx¦s¦¨¥\");
        }

        private void POTATOAR2_Load(object sender, EventArgs e)
        {
            this.gB_INVTRACKTableAdapter.Fill(this.POTATO.GB_INVTRACK);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.gB_INVTRACKTableAdapter.Fill(this.POTATO.GB_INVTRACK);
        }
    }
}
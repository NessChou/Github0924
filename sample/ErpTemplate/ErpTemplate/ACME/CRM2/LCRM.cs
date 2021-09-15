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
    public partial class LCRM : Form
    {
        public LCRM()
        {
            InitializeComponent();
        }

        private void aCME_LEADBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {


        }

        private void cRM_OCRDBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.cRM_OCRDBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.nCRM);

        }

        private void LCRM_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'nCRM.CRM_ITEM' 資料表。您可以視需要進行移動或移除。
            this.cRM_ITEMTableAdapter.Fill(this.nCRM.CRM_ITEM);
            this.cRM_OCRDTableAdapter.Fill(this.nCRM.CRM_OCRD);
        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            this.cRM_OCRDTableAdapter.FillBy(this.nCRM.CRM_OCRD, toolStripTextBox1.Text);
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.cRM_ITEMBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.nCRM);
        }

    

 

  

     

 
    }
}

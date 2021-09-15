using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class fmCrmSec : Form
    {
        public fmCrmSec()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

       

        private void fmCrmSec_Load(object sender, EventArgs e)
        {        this.aCME_CRM_SECTableAdapter.Fill(this.cRM.ACME_CRM_SEC);

        }

        private void aCME_CRM_SECBindingSource_ListChanged(object sender, ListChangedEventArgs e)
        {

        }


        //沒有 Delete 
        //需要用 DataTable 的刪除 
        //private void aCME_CRM_SECBindingSource_CurrentItemChanged(object sender, EventArgs e)
        //{
        //    DataRow ThisDataRow = ((DataRowView)((BindingSource)sender).Current).Row;
        //    if (ThisDataRow.RowState == DataRowState.Modified || ThisDataRow.RowState == DataRowState.Added
        //        || ThisDataRow.RowState == DataRowState.Deleted)
        //    {
        //        aCME_CRM_SECTableAdapter.Update(ThisDataRow);

        //    }
        //}

        //private void aCME_CRM_SECBindingSource_PositionChanged(object sender, EventArgs e)
        //{
        //    DataRow ThisDataRow = ((DataRowView)((BindingSource)sender).Current).Row;
        //    if (ThisDataRow.RowState == DataRowState.Modified || ThisDataRow.RowState == DataRowState.Added
        //        || ThisDataRow.RowState == DataRowState.Deleted)
        //    {
        //        aCME_CRM_SECTableAdapter.Update(ThisDataRow);

        //    }

        //}

       

        private void aCME_CRM_SECBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aCME_CRM_SECBindingSource.EndEdit();
            this.aCME_CRM_SECTableAdapter.Update(this.cRM.ACME_CRM_SEC);
            MessageBox.Show("存檔完成");
        }
    }
}
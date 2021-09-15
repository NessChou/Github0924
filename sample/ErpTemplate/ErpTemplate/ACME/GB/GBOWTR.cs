using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class GBOWTR : ACME.fmBase1
    {
        public GBOWTR()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.CHOICEConnection;
            stkModAdjMainTableAdapter.Connection = MyConnection;
            stkModAdjSubTableAdapter.Connection = MyConnection;

        }

        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                cHOICE.stkModAdjMain.RejectChanges();
                cHOICE.stkModAdjSub.RejectChanges();

            }
            catch
            {
            }
            return true;

        }
        public override void SetInit()
        {

            MyBS = stkModAdjMainBindingSource;
            MyTableName = "stkModAdjMain";
            MyIDFieldName = "ModAdjNO";

        }
        public override void FillData()
        {
            try
            {
                stkModAdjMainTableAdapter.Fill(cHOICE.stkModAdjMain, MyID);
                stkModAdjSubTableAdapter.Fill(cHOICE.stkModAdjSub, MyID);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = stkModAdjSubDataGridView.Rows.Count - 1; i >= 0; i--)
            {
                row = stkModAdjSubDataGridView.Rows[i];

                //string S = comboBox1.SelectedValue.ToString();
                //string T = comboBox2.SelectedValue.ToString();
                row.Cells[4].Value = (Convert.ToInt16(row.Cells[4].Value) * Convert.ToInt16(textBox1.Text)).ToString();
            }
        }
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {

                Validate();
  

                stkModAdjMainTableAdapter.Connection.Open();


                stkModAdjMainBindingSource.EndEdit();
                stkModAdjSubBindingSource.EndEdit();


                tx = stkModAdjMainTableAdapter.Connection.BeginTransaction();

                SqlDataAdapter Adapter = util.GetAdapter(stkModAdjMainTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(stkModAdjSubTableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;


                stkModAdjMainTableAdapter.Update(cHOICE.stkModAdjMain);
                stkModAdjSubTableAdapter.Update(cHOICE.stkModAdjSub);


                cHOICE.stkModAdjMain.AcceptChanges();
                cHOICE.stkModAdjSub.AcceptChanges();


                this.MyID = this.modAdjNOTextBox.Text;
                tx.Commit();


                UpdateData = true;





            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }

                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.stkModAdjMainTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
    }
}

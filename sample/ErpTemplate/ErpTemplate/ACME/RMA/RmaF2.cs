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
    public partial class RmaF2 : Form
    {
        public string q1;
        public string q2;
        public RmaF2()
        {
            InitializeComponent();
        }

        private void RmaF2_Load(object sender, EventArgs e)
        {
            try
            {
                this.rMA_CTR1TableAdapter.Fill(this.rm.RMA_CTR1, q1, q2);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

           
        }
        private System.Data.DataTable GETSZ2(string U_RMA_NO, string MANUFSN)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_RMA_NO,MANUFSN,SUM(CASE WHEN U_U_ACME_JUDGE IN ('NDF','OK') THEN 1 END) OK,");
            sb.Append(" SUM(CASE WHEN U_U_ACME_JUDGE IN ('NG') THEN 1 END) NG FROM RMA_CTR1 WHERE U_RMA_NO=@U_RMA_NO AND MANUFSN=@MANUFSN GROUP BY U_RMA_NO,MANUFSN ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
            command.Parameters.Add(new SqlParameter("@MANUFSN", MANUFSN));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public void UPDATESZ(string OK, string NG, string DOCDATE, string RMANO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE RMA_INVOICEF SET OK=@OK,NG=@NG  FROM rma_MAINF T0 LEFT JOIN RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE DOCDATE=@DOCDATE AND RMANO=@RMANO", connection);


            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OK", OK));
            command.Parameters.Add(new SqlParameter("@NG", NG));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@RMANO", RMANO));




            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public System.Data.DataTable DD23(string U_RMA_NO, string manufsn)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_RMA_NO 'RMA NO',U_S_SEQ 'S/N',U_MODEL Model,U_U_VER Ver");
            sb.Append(" ,U_U_MONTH_SEQ 'W/C',U_U_IQC 'IQC/CLR/FR',U_U_C_COMPLAIN 'Customer Complain'");
            sb.Append(" ,U_U_ACME_CONFIRM 'ACMEPOINT Confirm',U_U_ACME_JUDGE 'ACMEPOINT Judge',U_U_PLACE_1 產地");
            sb.Append(" ,U_RREMARK REMARK  FROM RMA_CTR1 where U_RMA_NO=@U_RMA_NO and manufsn=@manufsn");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
            command.Parameters.Add(new SqlParameter("@manufsn", manufsn));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }

        private void rMA_CTR1BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rMA_CTR1BindingSource.EndEdit();
            this.rMA_CTR1TableAdapter.Update(this.rm.RMA_CTR1);

            MessageBox.Show("資料已刪除");
            System.Data.DataTable J1 = GETSZ2(q1, q2);
            if (J1.Rows.Count > 0)
            {
                string OK = J1.Rows[0]["OK"].ToString();
                string NG = J1.Rows[0]["NG"].ToString();

                UPDATESZ(OK, NG, q2, q1);
            }
        }

        private void rMA_CTR1DataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

    
    }
}

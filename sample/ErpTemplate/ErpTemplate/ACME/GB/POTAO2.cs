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
    public partial class POTAO2 : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public POTAO2()
        {
            InitializeComponent();
        }

        private void gB_POTATOBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_POTATOBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);

        }

        private void fillToolStripButton_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    this.gB_POTATOTableAdapter.Fill(this.pOTATO.GB_POTATO, createDateToolStripTextBox.Text, createDate2ToolStripTextBox.Text, tRANSMARKToolStripTextBox.Text, oRDERPINToolStripTextBox.Text);
            //}
            //catch (System.Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show(ex.Message);
            //}

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void POTAO2_Load(object sender, EventArgs e)
        {

            gB_POTATOTableAdapter.Connection = globals.Connection;
            gB_FRIENDTableAdapter.Connection = globals.Connection;
            gB_POTATO2TableAdapter.Connection = globals.Connection;

            toolStripTextBox1.Text = GetMenu.DFirst();
            toolStripTextBox2.Text = GetMenu.DLast();

            this.gB_POTATOTableAdapter.Fill(this.pOTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox1.Text, toolStripTextBox3.Text);
            this.gB_FRIENDTableAdapter.Fill(this.pOTATO.GB_FRIEND);
            this.gB_POTATO2TableAdapter.Fill(this.pOTATO.GB_POTATO2);
            DELETECC();
            DELETEDD();


            toolStripComboBox1.ComboBox.DataSource = GetOslp1();
            toolStripComboBox1.ComboBox.ValueMember = "DataValue";
            toolStripComboBox1.ComboBox.DisplayMember = "DataValue";

            toolStripComboBox1.Text = "快遞單號";
            if (globals.GroupID.ToString().Trim() == "WH" || globals.GroupID.ToString().Trim() == "GB" || globals.GroupID.ToString().Trim() == "GBT" || globals.GroupID.ToString().Trim() == "EEP")
            {

            }
            else
            {
                gB_POTATOBindingNavigatorSaveItem.Visible = false;

            }


            BILLNO();

        }

        public static System.Data.DataTable GetOslp1()
        {

            SqlConnection con = globals.Connection;
            string sql = "SELECT distinct rtrim(isnull(TransMark,'')) DataValue FROM dbo.GB_POTATO  ORDER BY rtrim(isnull(TransMark,''))";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["oslp"];
        }
        public void BILLNO()
        {
            if (gB_POTATODataGridView.Rows.Count > 0)
            {

                for (int i = 0; i <= gB_POTATODataGridView.Rows.Count - 2; i++)
                {
                    DataGridViewRow row;

                    row = gB_POTATODataGridView.Rows[i];
                    string T1 = row.Cells["ID"].Value.ToString();
                    string T2 = row.Cells["PROJECT"].Value.ToString();

                    if (String.IsNullOrEmpty(T2))
                    {
                        System.Data.DataTable G1 = GETBILLNO(T1);
                        if (G1.Rows.Count > 0)
                        {
                            string BILLNO = G1.Rows[0][0].ToString();

                            UPDATBILLNO(BILLNO, T1);
                        }

                    }
                }

            }
        }
        public System.Data.DataTable GETBILLNO(string CustBillNo)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BillNO   FROM ordBillMain  WHERE CustBillNo=@CustBillNo ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustBillNo", CustBillNo));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public void UPDATBILLNO(string PROJECT, string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;


            command = new SqlCommand("UPDATE GB_POTATO SET PROJECT=@PROJECT WHERE ID=@ID ", connection);

            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@ID", ID));



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
        private void DELETECC()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE GB_POTATO2 WHERE ID2 IN (SELECT ID2 FROM GB_POTATO2 T0 LEFT JOIN GB_POTATO T1 ON (T0.ID=T1.ID) WHERE ISNULL(T1.ID,'') = '')");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

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


        private void DELETEDD()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE GB_FRIEND WHERE DOCID IN (SELECT T0.DOCID FROM GB_FRIEND T0 LEFT JOIN GB_POTATO T1 ON (T0.DOCID=T1.ID) WHERE ISNULL(T1.ID,'') = '') ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

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

    
    }
}

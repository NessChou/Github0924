 using System;
 using System.Collections.Generic;
 using System.ComponentModel;
 using System.Data;
 using System.Drawing;
 using System.Text;
 using System.Windows.Forms;
 using System.Data.SqlClient;
 using System.IO;
using System.Reflection;



namespace ACME
{
    public partial class AP_WHS_List : Form
    {

        public AP_WHS_List()
        {
            InitializeComponent();
            gvData.AutoGenerateColumns = false;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {

          

        }

        private void gvData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "colEdit")
                {

                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;
                    if (row != null)
                    {
                        AP_WHS form = new AP_WHS(Convert.ToInt32(row["ID"]));
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            RefreshData();
                            try
                            {
                                (sender as DataGridView).CurrentCell = (sender as DataGridView)[0, scrollPosition];
                            }
                            catch
                            {

                            }
                        }

                    }
                }

            }
        }

        private void RefreshData()
        {
            System.Data.DataTable dt = GetACME_CREDIT_UNLOCK_Condition();

            gvData.DataSource = dt;

        }

        private void btnAdd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            AP_WHS form = new AP_WHS(0);

            if (form.ShowDialog() == DialogResult.OK)
            {
                RefreshData();
            }
        }


        // Condition 版本
        public System.Data.DataTable GetACME_CREDIT_UNLOCK_Condition()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;


            sb.Append("SELECT * FROM Shipping_WHS WHERE 1= 1 ");
            if (textBox1.Text != "")
            {
                sb.Append(" AND WHSCODE LIKE '%" + textBox1.Text + "%' ");
     
            }
            sb.Append(" order by WHSCODE ");

            command.CommandText = sb.ToString();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }






        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }


        private string FormatDateStr(string sDate)
        {

            try
            {
                return sDate.Substring(0, 4) + "/" + sDate.Substring(4, 2) + "/" + sDate.Substring(6, 2);
            }
            catch
            {
                return "";
            }
        }






        private void UNLOCK_List_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt =  GetACME_CREDIT_UNLOCK_Condition();

            gvData.DataSource = dt;

            gvData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gvData);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetACME_CREDIT_UNLOCK_Condition();

            gvData.DataSource = dt;

            gvData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

        }
    }
}


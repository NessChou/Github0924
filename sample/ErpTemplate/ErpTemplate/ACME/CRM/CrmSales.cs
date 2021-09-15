using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Data.SqlClient;

namespace ACME.CRM
{
    public partial class CrmSales : Form
    {
        private string ConnStrSP ="server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";

        public string EmpID = "";

        public CrmSales()
        {
            InitializeComponent();
        }

        private void CrmSales_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetEmpID();
        }


        public DataTable GetEmpID()
        {
            SqlConnection connection = new SqlConnection(ConnStrSP);
            string sql = "SELECT name as LoginName,EmpID,SapName from employee where kind like '%sales%' ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
           // command.Parameters.Add(new SqlParameter("@name", LoginId));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EmpID = Convert.ToString(dataGridView1.CurrentRow.Cells["LoginName"].Value);
            DialogResult = DialogResult.OK;
        }

      

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            button1_Click(sender, e);
        }




    }
}
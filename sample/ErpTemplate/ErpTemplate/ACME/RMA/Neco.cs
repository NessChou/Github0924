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
    public partial class Neco : Form
    {
        public Neco()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = GetOrderData6(textBox1.Text.Trim().ToString());
            DataRow drw = dt.Rows[0];
            string av = drw["atcentry"].ToString();
            for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
            {
                string aa = dataGridView1.Rows[i].Cells[0].Value.ToString();
                if(!String.IsNullOrEmpty(av))
                {
                    UpdateSQL(av, aa);
                }
            }
           
            MessageBox.Show("更新成功");
          
        }
        private void UpdateSQL(string ATCENTRY, string CONTRACTID)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update OCTR set ATCENTRY=@ATCENTRY where CONTRACTID=@CONTRACTID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ATCENTRY", ATCENTRY));
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
     
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
        private System.Data.DataTable GetOrderData6(string contractid)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("select atcentry from OCTR where contractid=@contractid");
 
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@contractid", contractid));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    }
}
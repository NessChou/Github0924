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
    public partial class APInvoice : Form
    {
        public APInvoice()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBatchPayment5();
        }
        private void ViewBatchPayment5()
        {
            //合計 AS 銷售金額
           // string SAPConnStr_AcmeSql02 = "server=acmesrv13;pwd=@rmas;uid=SapReader;database=acmesql98";
            SqlConnection connection = new SqlConnection("server=acmesrv13;pwd=@rmas;uid=SapReader;database=acmesql98");

            StringBuilder sb = new StringBuilder();


            sb.Append(" select linenum,itemcode,dscription,quantity,u_acme_inv from pdn1 where docentry=@docentry ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "pdn1");
            }
            finally
            {
                connection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 1)
                {
                    for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        string aa = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        string dd = dataGridView1.Rows[i].Cells[4].Value.ToString();
              
                        AddTRACKER_LOG(dd, Convert.ToInt16(textBox1.Text), Convert.ToInt16(aa));
                    }

                }
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(string u_acme_inv, int docentry, int linenum)
        {



            SqlConnection connection = new SqlConnection("server=acmesrv13;pwd=@rmas;uid=Sapdbo;database=acmesql98");
            StringBuilder sb = new StringBuilder();
            sb.Append(" update pdn1 set u_acme_inv=@u_acme_inv where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@u_acme_inv", u_acme_inv));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));


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
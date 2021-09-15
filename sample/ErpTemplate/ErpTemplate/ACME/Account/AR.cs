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
    public partial class AR : Form
    {
        public AR()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBatchPayment598();
        }
        private void ViewBatchPayment5()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" select linenum,itemcode,dscription,u_acme_dscription,quantity u_acme_pqty,price u_acme_cost from inv1 where docentry=@docentry ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment598()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" select linenum,itemcode,dscription,U_base_doc u_acme_dscription,quantity u_acme_pqty,price u_acme_cost from inv1 where docentry=@docentry ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

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
                        string bb = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        string cc = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        string dd = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        string ee = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        string ff = dataGridView1.Rows[i].Cells[5].Value.ToString();

                        AddTRACKER_LOG98(dd, Convert.ToInt32(textBox1.Text), Convert.ToInt16(aa));
                    }

                }
                AddTRACKER_LOG1(textBox2.Text, textBox3.Text, textBox4.Text,textBox5.Text, Convert.ToInt16(textBox1.Text));
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(string u_acme_dscription,int docentry,int linenum)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update inv1 set u_acme_dscription=@u_acme_dscription where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@u_acme_dscription", u_acme_dscription));
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

        private void AddTRACKER_LOG98(string U_base_doc,  int docentry, int linenum)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update inv1 set U_base_doc=@U_base_doc where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_base_doc", U_base_doc));
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
        private void AddTRACKER_LOG1(string u_acme_etd, string u_acme_etc, string u_acme_eta, string u_acme_shipworkday, int docentry)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update oinv set u_acme_etd=@u_acme_etd , u_acme_etc=@u_acme_etc , u_acme_eta=@u_acme_eta,u_acme_shipworkday=@u_acme_shipworkday where docentry=@docentry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@u_acme_etd", u_acme_etd));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@u_acme_etc", u_acme_etc));
            command.Parameters.Add(new SqlParameter("@u_acme_eta", u_acme_eta));
            command.Parameters.Add(new SqlParameter("@u_acme_shipworkday", u_acme_shipworkday));

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
        private void linkLabel1_Click(object sender, EventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\修改AR發票品名.doc";
            System.Diagnostics.Process.Start(aa);
        }
    }
}
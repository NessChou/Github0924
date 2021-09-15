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
    public partial class ChangeCommission : Form
    {
        public ChangeCommission()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBatchPayment5();
        }
        private void ViewBatchPayment5()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT linenum 'Lineno',T1.DOCENTRY 單號,T0.CARDNAME 客戶名稱,itemcode 項目號碼,T1.QUANTITY 數量,T1.U_COMMISSION 傭金 FROM OINV T0 ");
            sb.Append(" LEFT JOIN INV1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  where t0.docentry=@Docentry ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

            System.Data.DataTable DF = GetOrderData8();
            if (DF.Rows.Count > 0)
            {
                textBox2.Text = DF.Rows[0][0].ToString();
            }
          
        }
        private System.Data.DataTable GetOrderData8()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select U_ACME_DOC_RATE  from OINV where docentry=@Docentry  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));


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
                        string cc = dataGridView1.Rows[i].Cells[5].Value.ToString();

                        if (String.IsNullOrEmpty(cc))
                        {
                            cc = "0";
                        }
                        decimal gg = Convert.ToDecimal(cc);
                        AddTRACKER_LOG(gg, bb, aa);
                    }

                }

                string f1 = textBox2.Text;
                if (String.IsNullOrEmpty(f1))
                {
                    f1 = "0";
                }
                decimal g2 = Convert.ToDecimal(f1);
                UPDATE2(g2);
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(decimal U_COMMISSION, string docentry, string linenum)
        {



            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append(" update INV1 set U_COMMISSION=@U_COMMISSION  where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_COMMISSION", U_COMMISSION));
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

        private void UPDATE2(decimal U_ACME_DOC_RATE)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update OINV set U_ACME_DOC_RATE=@U_ACME_DOC_RATE  where docentry=@docentry");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_ACME_DOC_RATE", U_ACME_DOC_RATE));
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));


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
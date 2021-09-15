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
    public partial class APORTT : Form
    {
        public APORTT()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GETORTT(DateTime RATEDATE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT RATE FROM ORTT WHERE CURRENCY='USD' AND RATEDATE=@RATEDATE  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RATEDATE", RATEDATE));


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

        private System.Data.DataTable GETORTT2(DateTime RATEDATE1, DateTime RATEDATE2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT Convert(varchar(8),RATEDATE,112) 日期,CURRENCY,RATE FROM ORTT WHERE RATEDATE BETWEEN @RATEDATE1 AND @RATEDATE2  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RATEDATE1", RATEDATE1));
            command.Parameters.Add(new SqlParameter("@RATEDATE2", RATEDATE2));

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
        private void INSERTORTT(DateTime RateDate,decimal Rate)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Insert into [dbo].[ORTT]([RateDate],[Currency],[Rate],[DataSource],[UserSign]) values(@RateDate,@Currency,@Rate,@DataSource,@UserSign)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@RateDate", RateDate));
            command.Parameters.Add(new SqlParameter("@Currency", "USD"));
            command.Parameters.Add(new SqlParameter("@Rate", Rate));
            command.Parameters.Add(new SqlParameter("@DataSource", "I"));
            command.Parameters.Add(new SqlParameter("@UserSign", 10));

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

        private void UPDATEORTT(DateTime RateDate, decimal Rate)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE T0 SET T0.[Rate] = @Rate  FROM [dbo].[ORTT] T0 WHERE T0.[RateDate] = @RateDate  AND  T0.[Currency] ='USD' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@RateDate", RateDate));
            command.Parameters.Add(new SqlParameter("@Rate", Rate));

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

        private void APORTT_Load(object sender, EventArgs e)
        {
            if (globals.GroupID.ToString().Trim() == "SHI" || globals.GroupID.ToString().Trim() == "RMA")
            {
                button2.Visible = false;
            }
            textBox1.Text = GetMenu.Day();
            textBox3.Text = GetMenu.DFirst();
            textBox4.Text = GetMenu.Day();
            GRDRATE();
            GRDRATE2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GRDRATE();
        }

        private void GRDRATE()
        {
            string D1 = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
            DateTime T1 = Convert.ToDateTime(D1);

            System.Data.DataTable G1 = GETORTT(T1);
            if (G1.Rows.Count > 0)
            {
                textBox2.Text = G1.Rows[0][0].ToString();
            }
            else
            {
                textBox2.Text = "0";
            }
        }
        private void GRDRATE2()
        {
            string D1 = textBox3.Text.Substring(0, 4) + "/" + textBox3.Text.Substring(4, 2) + "/" + textBox3.Text.Substring(6, 2);
            string D2 = textBox4.Text.Substring(0, 4) + "/" + textBox4.Text.Substring(4, 2) + "/" + textBox4.Text.Substring(6, 2);
            DateTime T1 = Convert.ToDateTime(D1);
            DateTime T2 = Convert.ToDateTime(D2);
            System.Data.DataTable G1 = GETORTT2(T1, T2);
            dataGridView1.DataSource = G1; 
        }

       
        private void button2_Click(object sender, EventArgs e)
        {
                        decimal n;
                        if (!decimal.TryParse(textBox2.Text, out n))
                        {
                            MessageBox.Show("請輸入正確匯率");
                            return;
                        }
                      
                        string D1 = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
            DateTime T1 = Convert.ToDateTime(D1);

            System.Data.DataTable G1 = GETORTT(T1);
            if (G1.Rows.Count > 0)
            {

                UPDATEORTT(T1, Convert.ToDecimal(textBox2.Text));
                MessageBox.Show("匯率已更新");
            }
            else
            {
                INSERTORTT(T1, Convert.ToDecimal(textBox2.Text));
                MessageBox.Show("匯率已新增");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GRDRATE2();
        }
    }
}

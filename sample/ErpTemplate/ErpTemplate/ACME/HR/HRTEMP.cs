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
    public partial class HRTEMP : Form
    {
        public HRTEMP()
        {
            InitializeComponent();
        }

        private void HRTEMP_Load(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GetMenu.GETORTT();
            if (G1.Rows.Count > 0)
            {
                textBox1.Text = G1.Rows[0][0].ToString();
              
            }
            textBox2.Text = DateTime.Now.ToString("yyyyMM");
        }
        private void INSERTORTT()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Insert into [dbo].[HR_TEMP]([DOCDATE],[USERS],[YTEMP]) values(@DOCDATE,@USERS,@YTEMP)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@YTEMP", textBox1.Text));
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

        private void UPDATEORTT()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE T0 SET T0.[YTEMP] = @YTEMP  FROM [dbo].[HR_TEMP] T0 WHERE DOCDATE=@DOCDATE AND USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@YTEMP", textBox1.Text));

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
     
        private void button1_Click(object sender, EventArgs e)
        {
            decimal  num1;
            if (decimal.TryParse(textBox1.Text, out num1) == false) //使用int.tryparse出來會是布林值
            {
                MessageBox.Show("請輸入正確格式");
                return;
            }
            System.Data.DataTable G1 = GetMenu.GETORTT();
            if (G1.Rows.Count > 0)
            {
                textBox1.Text = G1.Rows[0][0].ToString();
                UPDATEORTT();

            }
            else
            {
                INSERTORTT();
                MessageBox.Show("體溫已新增");
            }

            Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetMenu.GETORTT2(textBox2.Text);
        }
    }
}

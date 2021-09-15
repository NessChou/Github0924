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
    public partial class ChangeOpch2 : Form
    {
        public ChangeOpch2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "AP")
            {
                ViewOPCH();
            }
            if (comboBox1.Text == "AP貸項")
            {
                ViewORPC();
            }

        }
        private void ViewOPCH()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY 單號,T0.CARDNAME 客戶名稱,U_acme_pi 美金金額 FROM OPCH T0  where docentry=@Docentry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
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

        }
        private void ViewORPC()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY 單號,T0.CARDNAME 客戶名稱,U_acme_pi 美金金額 FROM ORPC T0  where docentry=@Docentry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
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

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                    {
                        string aa = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        string cc = dataGridView1.Rows[i].Cells[2].Value.ToString();

                        if (comboBox1.Text == "AP")
                        {
                            UPDATEOPCH(cc, aa);
                        }
                        if (comboBox1.Text == "AP貸項")
                        {
                            UPDATEORPC(cc, aa);
                        }
                    }

                }
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void UPDATEOPCH(string U_acme_pi, string docentry)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update OPCH set U_acme_pi =@U_acme_pi  where docentry=@docentry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_acme_pi", U_acme_pi));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));



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

        private void UPDATEORPC(string U_acme_pi, string docentry)
        {



            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append(" update orpc set U_acme_pi =@U_acme_pi  where docentry=@docentry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_acme_pi", U_acme_pi));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));



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

        private void ChangeOpch2_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "AP";
        }
 

    }
}
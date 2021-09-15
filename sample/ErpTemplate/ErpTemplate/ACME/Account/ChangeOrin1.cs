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
    public partial class ChangeOrin1 : Form
    {
        public ChangeOrin1()
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
            sb.Append(" SELECT linenum 排序,T1.DOCENTRY 單號,T0.CARDNAME 客戶名稱,itemcode 項目號碼,T1.DSCRIPTION 說明,T1.QUANTITY 數量,T1.PRICE 台幣單價,T1.U_ACME_INV 美金單價 FROM ORIN T0 ");
            sb.Append(" LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  where t0.docentry=@Docentry ");


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
                        string bb = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        string cc = dataGridView1.Rows[i].Cells[7].Value.ToString();
                        string dd = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        AddTRACKER_LOG(cc,dd,bb,aa);
                    }

                }
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(string U_ACME_INV,string DSCRIPTION, string docentry, string linenum)
        {



            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append(" update RIN1 set U_ACME_INV =@U_ACME_INV,DSCRIPTION=@DSCRIPTION  where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
            command.Parameters.Add(new SqlParameter("@DSCRIPTION", DSCRIPTION));
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
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
    public partial class ChangePay : Form
    {
        public ChangePay()
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
            sb.Append(" SELECT linenum 排序,T1.DOCENTRY 單號,T0.CARDNAME 客戶名稱,t1.itemcode 項目號碼,cast(t1.quantity as int) 數量,cast(t1.price as int) 單價,t1.u_acme_dscription 付款方法 FROM opor T0 ");
            sb.Append(" LEFT JOIN por1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  where t0.docentry=@Docentry ");


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
       
                        string dd = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        AddTRACKER_LOG(dd,bb,aa);
                    }

                }
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG( string U_ACME_dscription, string docentry, string linenum)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update por1 set U_ACME_dscription =@U_ACME_dscription  where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_ACME_dscription", U_ACME_dscription));
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
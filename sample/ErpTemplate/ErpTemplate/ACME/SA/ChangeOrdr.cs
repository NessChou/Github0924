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
    public partial class ChangeOrdr : Form
    {
        public ChangeOrdr()
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


            sb.Append(" SELECT linenum 排序,T1.DOCENTRY 單號,T0.CARDNAME 客戶名稱,t1.itemcode 項目號碼,cast(t1.quantity as int) 數量,t1.price 單價,Convert(varchar(10),t1.u_acme_work,112)  排程日期,t1.u_pay 付款,t1.u_shipday 押出貨日,t1.u_SHIPSTATUS 貨況,t1.U_MARK 特殊賣頭,T1.U_MEMO 注意事項,t1.u_acme_dscription 備註,T1.U_FINALPRICE 最終售價  FROM Ordr T0 ");
            sb.Append(" LEFT JOIN rdr1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  where t0.docentry=@Docentry ");


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
                        string a1 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        string a2 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        string a6 = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        string a7 = dataGridView1.Rows[i].Cells[7].Value.ToString();
                        string a8 = dataGridView1.Rows[i].Cells[8].Value.ToString();
                        string a9 = dataGridView1.Rows[i].Cells[9].Value.ToString();
                        string a10 = dataGridView1.Rows[i].Cells[10].Value.ToString();
                        string a11 = dataGridView1.Rows[i].Cells[11].Value.ToString();
                        string a12 = dataGridView1.Rows[i].Cells[12].Value.ToString();
                        string a13 = dataGridView1.Rows[i].Cells[13].Value.ToString();
                        if (String.IsNullOrEmpty(a13))
                        {
                            a13 = "0";
                        }
                        AddTRACKER_LOG(a6, a7, a8, a9, a10, a11, a12, a13, a2, a1);
                    }

                }
                MessageBox.Show("更新成功");
                ViewBatchPayment5();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(string U_ACME_work,string u_pay,string u_shipday,string u_SHIPSTATUS,string U_MARK,string U_MEMO,string U_ACME_dscription,string U_FINALPRICE, string docentry, string linenum)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rdr1 set U_ACME_work =CAST(@U_ACME_work AS DATETIME),u_pay =@u_pay,u_shipday=@u_shipday,u_SHIPSTATUS=@u_SHIPSTATUS,U_MARK=@U_MARK,U_MEMO=@U_MEMO,U_ACME_dscription =@U_ACME_dscription,U_FINALPRICE=@U_FINALPRICE   where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_ACME_work", U_ACME_work));
            command.Parameters.Add(new SqlParameter("@u_pay", u_pay));
            command.Parameters.Add(new SqlParameter("@u_shipday", u_shipday));
            command.Parameters.Add(new SqlParameter("@u_SHIPSTATUS", u_SHIPSTATUS));
            command.Parameters.Add(new SqlParameter("@U_MARK", U_MARK));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@U_ACME_dscription", U_ACME_dscription));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            command.Parameters.Add(new SqlParameter("@U_FINALPRICE", U_FINALPRICE));

            //U_FINALPRICE 

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
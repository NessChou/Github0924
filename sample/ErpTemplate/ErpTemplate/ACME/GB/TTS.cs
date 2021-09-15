using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class TTS : Form 
    {
       string  strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public string cardcode;
        public string usd;
        public string a,c;
        public TTS()
        {
            InitializeComponent();
        }

        private void TTS_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();


            ViewBatchPayment();
        }

        
        private void ViewBatchPayment()
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT BILLNO 訂單號碼,BillDate 日期,P.PersonName 業務,T0.SumAmtATax 訂單金額,");
            sb.Append("			  CASE WHEN  CHARINDEX('外部訂單單號:',Remark)-CHARINDEX('4.外部訂單總金額:',Remark)>0 THEN '' ELSE ");
            sb.Append("			  REPLACE(SUBSTRING(Remark,CHARINDEX('外部訂單單號:',Remark),CHARINDEX('4.外部訂單總金額:',Remark)-CHARINDEX('外部訂單單號:',Remark)),'外部訂單單號:','') END 外部訂單單號,");
            sb.Append("			T2.ShortName 客戶簡稱  FROM OrdBillMain T0 ");
            sb.Append("                   left join comPerson P ON (T0.Salesman=P.PersonID) ");
            sb.Append("                    Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID)  ");

            sb.Append(" WHERE BillDate BETWEEN @BillDate AND @BillDate2  AND T2.ClassID NOT IN ('000','026','020','029','014','022','027','028')");
            
            
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@BillDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " POR1");
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
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;
                    StringBuilder sb = new StringBuilder();

                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];

                        sb.Append("'" + row.Cells["訂單號碼"].Value.ToString() + "',");
                    }
                 
             


                    sb.Remove(sb.Length - 1, 1);

                    //linenum
                    string q = sb.ToString();
                


                   
                    a = q;
         
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {

                ViewBatchPayment();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    }
}
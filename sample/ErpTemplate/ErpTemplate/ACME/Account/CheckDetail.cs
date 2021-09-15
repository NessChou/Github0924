using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace ACME
{
    public partial class CheckDetail : Form
    {

        public CheckDetail()
        {
            InitializeComponent();
        }

   
        private void ViewBatchPayment2(string RefDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select CASE TRANSTYPE ");
            sb.Append("       WHEN -2 THEN '期初開帳' ");
            sb.Append("        WHEN 13 THEN 'AR' ");
            sb.Append("   WHEN 14 THEN 'AR貸項' ");
            sb.Append("   WHEN 15 THEN '交貨單' ");
            sb.Append("   WHEN 16 THEN '銷售退貨' ");
            sb.Append("   WHEN 18 THEN 'AP' ");
            sb.Append("   WHEN 19 THEN 'AP貸項' ");
            sb.Append("   WHEN 20 THEN '收貨採購單' ");
            sb.Append("   WHEN 21 THEN '採購退貨' ");
            sb.Append("   WHEN 59 THEN '收貨單' ");
            sb.Append("   WHEN 60 THEN '發貨單' ");
            sb.Append("   WHEN 67 THEN '庫存調撥' ");
            sb.Append("      ELSE CAST(TRANSTYPE AS VARCHAR)   END 單據總類,BASE_REF 單號,CAST(TRANSVALUE AS INT) 存貨價值 ");
            sb.Append("  from oinm t0    left JOIN OITM T11 ON T0.ITEMCODE = T11.ITEMCODE    where  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append(" AND TRANSVALUE <> 0");
            sb.Append(" AND  Convert(varchar(8),docdate,112) = @RefDate ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RefDate", RefDate));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment3(string RefDate)
        {
            SqlConnection connection = globals.shipConnection;



            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT transid 傳票號碼,cast((T0.[Debit])- (T0.[Credit]) as int) 科目餘額 ");
            sb.Append("             FROM  [dbo].[JDT1] T0 inner join OACT T1 on T0.Account = T1.AcctCode  ");
            sb.Append("             where Convert(varchar(8),refdate,112)= @RefDate  ");
            sb.Append("             AND  T0.[Account]  like '12000%'  AND T0.[Account] <> '12000201' ");

            sb.Append(" order by transid");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RefDate", RefDate));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            dataGridView2.DataSource = ds.Tables[0];

        }
        private void CheckDetail_Load(object sender, EventArgs e)
        {
 
            textBox2.Text = DateTime.Now.ToString("yyyyMMdd");
            label1.Text = "";
            label3.Text = "";
        }



        private void button4_Click(object sender, EventArgs e)
        {

            ViewBatchPayment2(textBox2.Text.Trim());
            ViewBatchPayment3(textBox2.Text.Trim());

            decimal iTotal = 0;

            decimal iTotal2 = 0;
            try
            {


                int i = this.dataGridView1.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    iTotal += Convert.ToDecimal(dataGridView1.Rows[iRecs].Cells["存貨價值"].Value);



                }
            }
            catch (Exception ex)
            {
            }


            try
            {


                int i = this.dataGridView2.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    iTotal2 += Convert.ToDecimal(dataGridView2.Rows[iRecs].Cells["科目餘額"].Value);



                }
            }
            catch (Exception ex)
            {
            }
            label1.Text = "存貨價值 " + iTotal.ToString("#,##0");
            label3.Text = "科目餘額 " + iTotal2.ToString("#,##0"); ;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
            ExcelReport.GridViewToExcel(dataGridView1);
        }

    }
}
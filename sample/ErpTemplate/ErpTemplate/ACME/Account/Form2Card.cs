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
    public partial class Form2Card : Form
    {
        public Form2Card()
        {
            InitializeComponent();
        }

       

        private void Form2Card_Load(object sender, EventArgs e)
        {

        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("請輸入資訊");
                }
                else
                {
                    if (radioButton1.Checked)
                    {

                        ViewBatchPayment();
                    }
                    else if (radioButton3.Checked)
                    {

                        ViewBatchPayment2();
                    }
     
                    else
                    {
                        ViewBatchPayment1();
                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ViewBatchPayment()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  Convert(varchar(10),T0.[DocDate],111)  採購日期,T0.[Docnum] 單據號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy 未稅金額,t1.linevat 稅額,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP 付款條件,T7.WHSNAME 倉庫");
            sb.Append(" FROM OPDN T0 ");
            sb.Append(" INNER JOIN PDN1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OWHS T7 ON (T7.WHSCODE=T1.WHSCODE)");
            sb.Append(" where Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append(" AND SUBSTRING(T0.CARDCODE,1,1)='U' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }
        private void ViewBatchPayment1()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  Convert(varchar(10),T0.[DocDate],111)  採購日期,T0.[Docnum] 單據號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy 未稅金額,t1.linevat 稅額,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP 付款條件,T7.WHSNAME 倉庫");
            sb.Append(" FROM OPOR T0 ");
            sb.Append(" INNER JOIN POR1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OWHS T7 ON (T7.WHSCODE=T1.WHSCODE)");
            sb.Append(" where Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append(" AND SUBSTRING(T0.CARDCODE,1,1)='U' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void ViewBatchPayment2()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  Convert(varchar(10),T0.[DocDate],111)  採購日期,T0.[Docnum] 單據號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy 未稅金額,t1.linevat 稅額,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP 付款條件,T7.WHSNAME 倉庫");
            sb.Append(" FROM OPCH T0 ");
            sb.Append(" INNER JOIN PCH1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OWHS T7 ON (T7.WHSCODE=T1.WHSCODE)");
            sb.Append(" where Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append(" AND SUBSTRING(T0.CARDCODE,1,1)='U' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

    }
}
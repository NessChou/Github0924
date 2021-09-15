using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;

//HashTable
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ACME
{
    public partial class RmaForm2 : Form
    {
        string COM = "";
        public RmaForm2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();

            if (LookupValues != null)
            {
                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                if (cardNameTextBox.Text == "")
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
                    else if (radioButton2.Checked)
                    {
                        ViewBatchPayment1();
                    }
                    else if (radioButton4.Checked)
                    {
                        ViewBatchPayment4();
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

            sb.Append(" SELECT T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T4.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append("      FROM OPDN T0 ");
            sb.Append("          INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append("          LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" LEFT JOIN PDN1 T4 ON (T0.DOCENTRY=T4.DOCENTRY) ");
            sb.Append(" WHERE (SUBSTRING(T4.U_SHIPPING_NO,0,4) in ('RMA' , 'RMR', 'RMS') OR SUBSTRING(T4.U_SHIPPING_NO,1,1) = 'A' )  ");
            sb.Append(" and T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append("  and T0.[DocNum]  not in (select isnull(baseref,0) from  acmesql02.dbo.rpd1)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@tt", cardCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }
        private void ViewBatchPayment1()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T4.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append("      FROM OPOR T0 ");
            sb.Append("          INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append("          LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" LEFT JOIN POR1 T4 ON (T0.DOCENTRY=T4.DOCENTRY) ");
            sb.Append(" WHERE (SUBSTRING(T4.U_SHIPPING_NO,0,4) in ('RMA' , 'RMR', 'RMS') OR SUBSTRING(T4.U_SHIPPING_NO,1,1) = 'A' )  ");
            sb.Append(" AND T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@tt", cardCodeTextBox.Text));
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
            dataGridView4.DataSource = bindingSource1;

        }

        private void ViewBatchPayment2()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T4.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append("      FROM OPCH T0 ");
            sb.Append("          INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append("          LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" LEFT JOIN pch1 T4 ON (T0.DOCENTRY=T4.DOCENTRY) ");
            sb.Append(" WHERE (SUBSTRING(T4.U_SHIPPING_NO,0,4) in ('RMA' , 'RMR', 'RMS') OR SUBSTRING(T4.U_SHIPPING_NO,1,1) = 'A' )  ");
            sb.Append(" AND T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append("  and T0.[DocNum]  not in (select isnull(baseref,0) from  acmesql02.dbo.rpc1)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@tt", cardCodeTextBox.Text));
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
            dataGridView4.DataSource = bindingSource1;

        }
        private void ViewBatchPayment4()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T4.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append("      FROM ORPC T0 ");
            sb.Append("          INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append("          LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" LEFT JOIN RPC1 T4 ON (T0.DOCENTRY=T4.DOCENTRY) ");
            sb.Append(" WHERE (SUBSTRING(T4.U_SHIPPING_NO,0,4) in ('RMA' , 'RMR', 'RMS') OR SUBSTRING(T4.U_SHIPPING_NO,1,1) = 'A' )  ");
            sb.Append(" AND T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append("  and T0.[DocNum]  not in (select isnull(baseref,0) from  acmesql02.dbo.rpc1)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@tt", cardCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ORPC");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow row;

                StringBuilder sb = new StringBuilder();

                for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = dataGridView4.SelectedRows[i];



                  sb.Append("'" + row.Cells["單據號碼"].Value.ToString() + "',");

                }

            

                sb.Remove(sb.Length - 1, 1);
                if (radioButton1.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPDN(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackOPDN2(sb.ToString());
                    frm.ShowDialog();
                   
                   
                  
                }
                else if (radioButton3.Checked)
                {
                    //AP發票
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPCH(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackOPCH2(sb.ToString());
                    frm.ShowDialog();

                }
                else if (radioButton2.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPOR(sb.ToString(), textBox3.Text, textBox4.Text,COM);
                    frm.dt2 = PayFormat.PackOPOR2(sb.ToString());
                    frm.ShowDialog();
                }
                else if (radioButton4.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackORPC(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackORPC2(sb.ToString());
                    frm.ShowDialog();
                }
    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = DateTime.Now.ToString("yyyyMMdd");

            if (globals.DBNAME != "進金生能源服務")
            {
                COM = "進金生實業股份有限公司";
            }
            else
            {
                COM = "進金生能源服務股份有限公司";
            }
        }

    }
}
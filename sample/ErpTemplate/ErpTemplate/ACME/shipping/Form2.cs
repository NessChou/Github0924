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
    public partial class Form2 : Form
    {
        string COM = "";
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton5.Checked || radioButton7.Checked || radioButton8.Checked)
            {
                object[] LookupValues = GetMenu.GetMenuListS1();

                if (LookupValues != null)
                {
                    cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                    cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            else
            {


                object[] LookupValues = GetMenu.GetMenuListS();

                if (LookupValues != null)
                {
                    cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                    cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            
            
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
                        //AP發票
                    else if (radioButton3.Checked)
                    {

                        ViewBatchPayment2();
                    }
                    //AP貸項
                    else if (radioButton4.Checked)
                    {

                        ViewBatchPaymentORPC();
                    }
                    else if (radioButton5.Checked)
                    {

                        ViewBatchPayment5();
                    }
                    else if (radioButton6.Checked)
                    {

                        ViewBatchPaymentAP();
                    }
                    else if (radioButton7.Checked)
                    {

                        ViewBatchPaymentAR();
                    }
                    else if (radioButton8.Checked)
                    {

                        ViewBatchPaymentAR2();
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

            sb.Append(" SELECT  T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM OPDN T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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

            sb.Append(" SELECT  T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM OPOR T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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

            sb.Append(" SELECT  T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM OPCH T0  ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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

        private void ViewBatchPaymentAP()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  'AP發票' 總類,T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM OPCH T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT  'AP貸項' 總類,T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM ORPC T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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
                da.Fill(ds, "OPCH");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }
        private void ViewBatchPaymentORPC()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  'AP貸項' 總類,T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM ORPC T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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
                da.Fill(ds, "OPCH");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }
        private void ViewBatchPaymentAR()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  'AR發票' 總類,T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM OINV T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT  'AR貸項' 總類,T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append(" FROM ORIN T0  ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append(" where T0.[CardCode]=@tt");
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
                da.Fill(ds, "OINV");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }
        private void ViewBatchPaymentAR2()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                                  SELECT 'AR發票' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單據號碼,Convert(varchar(10),(t0.docdate),112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                                  T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱,T8.Gtotalfc 美金小計,");
            sb.Append("                                  isnull(cast((case when  (row_number()  over (partition by T0.[doctotal] order by t0.DOCENTRY))=1 ");
            sb.Append("                                  then T0.[doctotal]else null end) as varchar),'') as 台幣小計  FROM acmesql02.dbo.OINV T0  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("                                  WHERE T0.[DocType] ='I' ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(8),t0.docdate,112)  between '" + textBox1.Text.ToString() + "' and  '" + textBox2.Text.ToString() + "'");
            }
            if (cardCodeTextBox.Text != "")
            {
                sb.Append(" and  T0.[CardCode]  = '" + cardCodeTextBox.Text.ToString() + "'");
            }
            sb.Append("                                 UNION ALL");
            sb.Append("                                  SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單據號碼,Convert(varchar(10),(t0.docdate),112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                                  T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱,0 美金小計,                      ");
            sb.Append("                                  isnull(cast((case when  (row_number()  over (partition by T0.[doctotal] order by t0.DOCENTRY))=1 ");
            sb.Append("                                  then T0.[doctotal]else null end) as varchar),'') as 台幣小計  FROM acmesql02.dbo.ORIN T0  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                                                     WHERE T0.[DocType] ='I' ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(8),t0.docdate,112)  between '" + textBox1.Text.ToString() + "' and  '" + textBox2.Text.ToString() + "'");
            }
            if (cardCodeTextBox.Text != "")
            {
                sb.Append(" and  T0.[CardCode]  = '" + cardCodeTextBox.Text.ToString() + "'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView4.DataSource = bindingSource1;

        }

        private void ViewBatchPayment5()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  T0.[DocDate] 採購日期,T0.[Docnum] 單據號碼,T0.U_SHIPPING_NO 工單號碼, cast(T0.[doctotal] as int)  金額,T3.PYMNTGROUP 付款條件,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱");
            sb.Append("              FROM Orin T0  ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM)");
            sb.Append("              where T0.[CardCode]=@tt");
            sb.Append("               and Convert(varchar(8),t0.docdate,112) between @startdate and @enddate ");
            //    sb.Append("  and a.closeday between @startdate and @enddate ");
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
                da.Fill(ds, "Orin");
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

                        //AP發票
                if (radioButton3.Checked)
                {
          
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPCH(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackOPCH2(sb.ToString());

                    frm.ShowDialog();
                }

                //收貨採購單
                else if (radioButton1.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPDN(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackOPDN2(sb.ToString());
                    frm.ShowDialog();
                   
                  
                }
                //採購單
                else if (radioButton2.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackOPOR(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackOPOR2(sb.ToString());
                    frm.ShowDialog();


                }
                //AR貸項通知單
                else if (radioButton5.Checked)
                {
                    Form4Rpt7 frm = new Form4Rpt7();
                    frm.dt = PayFormat.PackORIN(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                    frm.dt2 = PayFormat.PackORIN2(sb.ToString());
                    frm.ShowDialog();
                }
                //AP貸項
                else if (radioButton4.Checked)
                {
      
                        Form4Rpt7 frm = new Form4Rpt7();
                        frm.dt = PayFormat.PackORPC(sb.ToString(), textBox3.Text, textBox4.Text, COM);
                        frm.dt2 = PayFormat.PackORPC2(sb.ToString());
                        frm.ShowDialog();
            
                }
                    //AP發票+AP貸項
                else if (radioButton6.Checked)
                {
                    try
                    {
                        Form4Rpt7 frm = new Form4Rpt7();
                        frm.dt = PackAP();
                        frm.dt2 = PackAP2();
                        frm.ShowDialog();
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("請選擇AP貸項");
                    }
                }
                //AR發票+AR貸項
                else if (radioButton7.Checked)
                {
                    try
                    {
                        Form4Rpt7 frm = new Form4Rpt7();
                        frm.dt = PackAR();
                        frm.dt2 = PackAR3();
                        frm.ShowDialog();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("請選擇AR貸項");
                    }
                }

                    //財務對帳單
                else if (radioButton8.Checked)
                {
                    try
                    {
                        Form4Rpt7AR2 frm = new Form4Rpt7AR2();
                        frm.dt = PackAR2();
                        frm.ShowDialog();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
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


        public System.Data.DataTable PackAP()
        {
            
              
            DataGridViewRow row2;
            DataGridViewRow row3;
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row2 = dataGridView4.SelectedRows[i];


                if (row2.Cells["總類"].Value.ToString() == "AP發票")
                {
                    sb2.Append("'" + row2.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row3 = dataGridView4.SelectedRows[i];


                if (row3.Cells["總類"].Value.ToString() == "AP貸項")
                {
                    sb3.Append("'" + row3.Cells["單據號碼"].Value.ToString() + "',");
                }

            }
    
            sb2.Remove(sb2.Length - 1, 1);


            sb3.Remove(sb3.Length - 1, 1);


          
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  'AP/貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)-ISNULL(TE.SSS,0) SS,ISNULL(TG.dd,0)-ISNULL(TE.ddd,0) dd,ISNULL(TG.ee,0)-ISNULL(TE.eee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,aa=('" + textBox3.Text.ToString() + "'),bb=('" + textBox4.Text.ToString() + "')");
            sb.Append(" ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM acmesql02.dbo.OPCH  T0 ");
            sb.Append(" INNER JOIN dbo.PCH1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN dbo.OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM acmesql02.dbo.OPCH AA WHERE docentry  IN (" + sb2 + ")");
            sb.Append(" ) TG");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SSS, CAST(SUM(vatsum) AS INT) ddd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as eee  FROM acmesql02.dbo.ORPC AA WHERE docentry  IN (" + sb3 + ")");
            sb.Append(" ) TE");
            sb.Append(" WHERE   t1.docentry  IN (" + sb2 + ")");
            sb.Append(" UNION ALL");
            sb.Append("  SELECT  'AP貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int)*-1 金額,cast(t1.totalsumsy as int)*-1 totalsumsy ,cast(t1.linevat as int)*-1 linevat,(t1.totalsumsy+t1.linevat)*-1 加總,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)-ISNULL(TE.SSS,0) SS,ISNULL(TG.dd,0)-ISNULL(TE.ddd,0) dd,ISNULL(TG.ee,0)-ISNULL(TE.eee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,aa=('" + textBox3.Text.ToString() + "'),bb=('" + textBox4.Text.ToString() + "')");
            sb.Append(" ,t5.[name] ,(t0.vatsum)*-1 s2,(t0.doctotal-t0.vatsum)*-1 s3,(t0.doctotal)*-1 s4,T7.ocrNAME  部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱  FROM acmesql02.dbo.ORPC  T0 ");
            sb.Append(" INNER JOIN RPC1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN dbo.OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM acmesql02.dbo.OPCH AA WHERE docentry  IN (" + sb2 + ")");
            sb.Append(" ) TG");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SSS, CAST(SUM(vatsum) AS INT) ddd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as eee  FROM acmesql02.dbo.ORPC AA WHERE docentry  IN (" + sb3 + ")");
            sb.Append(" ) TE");
            sb.Append(" WHERE   t1.docentry  IN (" + sb3 + ")");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
            return ds.Tables["OPOR"];

        }

        public  System.Data.DataTable PackAP2()
        {
            DataGridViewRow row2;
            DataGridViewRow row3;
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row2 = dataGridView4.SelectedRows[i];


                if (row2.Cells["總類"].Value.ToString() == "AP發票")
                {
                    sb2.Append("'" + row2.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row3 = dataGridView4.SelectedRows[i];


                if (row3.Cells["總類"].Value.ToString() == "AP貸項")
                {
                    sb3.Append("'" + row3.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            sb2.Remove(sb2.Length - 1, 1);


            sb3.Remove(sb3.Length - 1, 1);
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT 科目,MAX(科目名稱) 科目名稱,ROUND(SUM(totalsumsy),0) totalsumsy,ROUND(SUM(linevat),0) linevat,ROUND(SUM(加總),0) 加總 FROM ( SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,SUM(t1.totalsumsy) totalsumsy,SUM(t1.linevat) linevat,SUM(t1.totalsumsy+t1.linevat)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + sb2 + ") GROUP BY T1.ACCTCODE");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,SUM(t1.totalsumsy)*-1 totalsumsy,SUM(t1.linevat)*-1 linevat,SUM(t1.totalsumsy+t1.linevat)*-1  加總 FROM RPC1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + sb3 + ") GROUP BY T1.ACCTCODE ) AS A GROUP BY 科目");
            sb.Append("   UNION ALL");
            sb.Append("             SELECT '四捨五入差異','',0,0,ROUND(SUM(加總),0)-(            SELECT ROUND(SUM(加總),0)  FROM ( SELECT SUM(t1.totalsumsy+t1.linevat)  加總 FROM PCH1 T1  WHERE   t1.docentry  IN (" + sb2 + ") ");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT SUM(t1.totalsumsy+t1.linevat)*-1  加總 FROM RPC1 T1  WHERE   t1.docentry  IN (" + sb3 + ") ) AS A ");
            sb.Append(" ) 加總 FROM ( SELECT  SUM(t1.DOCTOTAL)  加總 FROM OPCH T1  WHERE   t1.docentry  IN (" + sb2 + ") ");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT  SUM(t1.DOCTOTAL)*-1  加總 FROM ORPC T1  WHERE   t1.docentry  IN (" + sb3 + ")  ) AS A ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
    

        public System.Data.DataTable PackAR()
        {

            DataGridViewRow row2;
            DataGridViewRow row3;
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row2 = dataGridView4.SelectedRows[i];


                if (row2.Cells["總類"].Value.ToString() == "AR發票")
                {
                    sb2.Append("'" + row2.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row3 = dataGridView4.SelectedRows[i];


                if (row3.Cells["總類"].Value.ToString() == "AR貸項")
                {
                    sb3.Append("'" + row3.Cells["單據號碼"].Value.ToString() + "',");
                }

            }
   


            sb2.Remove(sb2.Length - 1, 1);
            sb3.Remove(sb3.Length - 1, 1);


            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  'AR/貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)-ISNULL(TE.SSS,0) SS,ISNULL(TG.dd,0)-ISNULL(TE.ddd,0) dd,ISNULL(TG.ee,0)-ISNULL(TE.eee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,aa=('" + textBox3.Text.ToString() + "'),bb=('" + textBox4.Text.ToString() + "')");
            sb.Append(" ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM acmesql02.dbo.OINV  T0 ");
            sb.Append(" INNER JOIN dbo.INV1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN dbo.OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM acmesql02.dbo.OINV AA WHERE docentry  IN (" + sb2 + ")");
            sb.Append(" ) TG");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SSS, CAST(SUM(vatsum) AS INT) ddd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as eee  FROM acmesql02.dbo.ORIN AA WHERE docentry  IN (" + sb3 + ")");
            sb.Append(" ) TE");
            sb.Append(" WHERE   t1.docentry  IN (" + sb2 + ")");
            sb.Append(" UNION ALL");
            sb.Append("  SELECT  'AR/貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int)*-1 金額,cast(t1.totalsumsy as int)*-1 totalsumsy ,cast(t1.linevat as int)*-1 linevat,(t1.totalsumsy+t1.linevat)*-1 加總,");
            sb.Append(" T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)-ISNULL(TE.SSS,0) SS,ISNULL(TG.dd,0)-ISNULL(TE.ddd,0) dd,ISNULL(TG.ee,0)-ISNULL(TE.eee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,aa=('" + textBox3.Text.ToString() + "'),bb=('" + textBox4.Text.ToString() + "')");
            sb.Append(" ,t5.[name] ,(t0.vatsum)*-1 s2,(t0.doctotal-t0.vatsum)*-1 s3,(t0.doctotal)*-1 s4,T7.ocrNAME  部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱  FROM acmesql02.dbo.ORIN  T0 ");
            sb.Append(" INNER JOIN RIN1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN dbo.OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append(" LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append(" LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM acmesql02.dbo.OINV AA WHERE docentry  IN (" + sb2 + ")");
            sb.Append(" ) TG");
            sb.Append(" ,(SELECT CAST(SUM(doctotal) AS INT) SSS, CAST(SUM(vatsum) AS INT) ddd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as eee  FROM acmesql02.dbo.ORIN AA WHERE docentry  IN (" + sb3 + ")");
            sb.Append(" ) TE");
            sb.Append(" WHERE   t1.docentry  IN (" + sb3 + ")");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OINV"];
        }
        public System.Data.DataTable PackAR3()
        {
            DataGridViewRow row2;
            DataGridViewRow row3;
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row2 = dataGridView4.SelectedRows[i];


                if (row2.Cells["總類"].Value.ToString() == "AR發票")
                {
                    sb2.Append("'" + row2.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row3 = dataGridView4.SelectedRows[i];


                if (row3.Cells["總類"].Value.ToString() == "AR貸項")
                {
                    sb3.Append("'" + row3.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            sb2.Remove(sb2.Length - 1, 1);


            sb3.Remove(sb3.Length - 1, 1);
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT 科目,MAX(科目名稱) 科目名稱,ROUND(SUM(totalsumsy),0) totalsumsy,ROUND(SUM(linevat),0) linevat,ROUND(SUM(加總),0) 加總 FROM ( SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,SUM(t1.totalsumsy) totalsumsy,SUM(t1.linevat) linevat,SUM(t1.totalsumsy+t1.linevat)  加總 FROM INV1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + sb2 + ") GROUP BY T1.ACCTCODE");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,SUM(t1.totalsumsy)*-1 totalsumsy,SUM(t1.linevat)*-1 linevat,SUM(t1.totalsumsy+t1.linevat)*-1  加總 FROM RIN1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + sb3 + ") GROUP BY T1.ACCTCODE ) AS A GROUP BY 科目");
            sb.Append("   UNION ALL");
            sb.Append("             SELECT '四捨五入差異','',0,0,ROUND(SUM(加總),0)-(            SELECT ROUND(SUM(加總),0)  FROM ( SELECT SUM(t1.totalsumsy+t1.linevat)  加總 FROM INV1 T1  WHERE   t1.docentry  IN (" + sb2 + ") ");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT SUM(t1.totalsumsy+t1.linevat)*-1  加總 FROM RIN1 T1  WHERE   t1.docentry  IN (" + sb3 + ") ) AS A ");
            sb.Append(" ) 加總 FROM ( SELECT  SUM(t1.DOCTOTAL)  加總 FROM OINV T1  WHERE   t1.docentry  IN (" + sb2 + ") ");
            sb.Append("             UNION ALL");
            sb.Append("              SELECT  SUM(t1.DOCTOTAL)*-1  加總 FROM ORIN T1  WHERE   t1.docentry  IN (" + sb3 + ")  ) AS A ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
        public System.Data.DataTable PackAR2()
        {

            DataGridViewRow row2;
            DataGridViewRow row3;
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row2 = dataGridView4.SelectedRows[i];


                if (row2.Cells["總類"].Value.ToString() == "AR發票")
                {
                    sb2.Append("'" + row2.Cells["單據號碼"].Value.ToString() + "',");
                }

            }

            for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
            {

                row3 = dataGridView4.SelectedRows[i];


                if (row3.Cells["總類"].Value.ToString() == "AR貸項")
                {
                    sb3.Append("'" + row3.Cells["單據號碼"].Value.ToString() + "',");
                }

            }
            ArrayList al2 = new ArrayList();

    


            sb2.Remove(sb2.Length - 1, 1);


        
            if (sb3.ToString() == "")
            {

            }
            else
            {
                sb3.Remove(sb3.Length - 1, 1);
            }

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
      
            sb.Append("                                  SELECT 'AR發票' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單據號碼,Convert(varchar(10),(t0.docdate),112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                                  T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱,cast(T8.Gtotalfc as int) 美金小計,");
            sb.Append("                                  isnull(cast((case when  (row_number()  over (partition by T0.[doctotal] order by t0.DOCENTRY))=1 ");
            sb.Append("                                  then T0.[doctotal]else null end) as int),0) as 台幣小計  FROM acmesql02.dbo.OINV T0  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("                                  WHERE T0.[DocType] ='I' ");
         
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(8),t0.docdate,112)  between '" + textBox1.Text.ToString() + "' and  '" + textBox2.Text.ToString() + "'");
            }
            if (cardCodeTextBox.Text != "")
            {
                sb.Append(" and  T0.[CardCode]  = '" + cardCodeTextBox.Text.ToString() + "'");
            }
            if (sb2.ToString() != "")
            {
                sb.Append(" and   t1.docentry  IN (" + sb2 + ")");
            }
            sb.Append("                                 UNION ALL");
            sb.Append("                                  SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單據號碼,Convert(varchar(10),(t0.docdate),112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                                  T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱,0 美金小計,                      ");
            sb.Append("                                  isnull(cast((case when  (row_number()  over (partition by T0.[doctotal] order by t0.DOCENTRY))=1 ");
            sb.Append("                                  then T0.[doctotal]else null end) as int),'') as 台幣小計  FROM acmesql02.dbo.ORIN T0  ");
            sb.Append("                                  LEFT JOIN acmesql02.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                                                     WHERE T0.[DocType] ='I' ");
         
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(8),t0.docdate,112)  between '" + textBox1.Text.ToString() + "' and  '" + textBox2.Text.ToString() + "'");
            }
            if (cardCodeTextBox.Text != "")
            {
                sb.Append(" and  T0.[CardCode]  = '" + cardCodeTextBox.Text.ToString() + "'");
            }
            if (sb3.ToString() != "")
            {
                sb.Append(" and   t1.docentry  IN (" + sb3 + ")");
            }
            else

            {
                sb.Append(" and   t1.docentry='0' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OINV"];
        }

      
     

    }
}
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
    public partial class AP : Form 
    {
        public string cardcode;
        public string usd;
        public string a,c;
        public AP()
        {
            InitializeComponent();
        }

        private void AP_Load(object sender, EventArgs e)
        {
    
                    if (radioButton1.Checked == true)
                    {
                        if (usd == "USD")
                        {
                            ViewBatchPayment4(cardcode);
                        }
                        else
                        {
                            ViewBatchPayment3(cardcode);
                        }
                       
                    }
                    if (radioButton2.Checked == true)
                    {
                        ViewBatchPayment2(cardcode);
                    }
            if (radioButton4.Checked == true)
            {
                ViewBatchPayment24(cardcode);
            }
            if (radioButton2.Checked == true)
                    {
                        ViewBatchPayment2(cardcode);
                    }
                    if (radioButton3.Checked == true)
                    {
                        if (usd == "USD")
                        {
                            ViewBatchPayment42(cardcode);
                        }
                        else
                        {
                            ViewBatchPayment32(cardcode);
                        }

                    }
              
        }
        private void ViewBatchPayment(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA");
            sb.Append("               from acmesql02.dbo.POR1 T0 ");
            sb.Append("              left join acmesql02.dbo.OPOR T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購單') ");
            sb.Append("              where T0.Docentry=@Docentry and T1.cardcode=@cardcode ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

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

        private void ViewBatchPayment4S(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA");
            sb.Append("               from acmesql02.dbo.PQT1 T0 ");
            sb.Append("              left join acmesql02.dbo.OPQT T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購報價') ");
            sb.Append("              where T0.Docentry=@Docentry and T1.cardcode=@cardcode ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");
            
            
            
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

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
        private void ViewBatchPayment1(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV");
            sb.Append("               from acmesql02.dbo.PCH1 T0 ");
            sb.Append("              left join acmesql02.dbo.OPCH T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='AP發票') ");
            sb.Append("              where T1.U_ACME_INV=@Docentry  and T1.cardcode=@cardcode ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity,t1.U_ACME_INV");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void ViewBatchPayment12(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                         select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV ");
            sb.Append("                             from acmesql02.dbo.PDN1 T0  ");
            sb.Append("                            left join acmesql02.dbo.OPDN T1 on (T0.docentry=T1.docentry) ");
            sb.Append("                            left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購')  ");
            sb.Append("                            where T1.Docentry=@Docentry  and T1.cardcode=@cardcode  ");
            sb.Append("               group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity,t1.U_ACME_INV ");
            sb.Append("               having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0 ");
            sb.Append("               order by T0.Docentry desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
            }
            finally
            {
                connection.Close();
            }

           int G1 = ds.Tables[0].Rows.Count;
            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment2(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA");
            sb.Append("               from acmesql02.dbo.POR1 T0 ");
            sb.Append("              left join acmesql02.dbo.OPOR T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購單') ");
            sb.Append("              where T1.cardcode=@cardcode  ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void ViewBatchPayment24(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA");
            sb.Append("               from acmesql02.dbo.PQT1 T0 ");
            sb.Append("              left join acmesql02.dbo.OPQT T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購報價') ");
            sb.Append("              where T1.cardcode=@cardcode  ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment3(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("           select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV");
            sb.Append("               from acmesql02.dbo.pch1 T0 ");
            sb.Append("              left join acmesql02.dbo.opch T1 on (T0.docentry=T1.docentry)");
            sb.Append("              left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='AP發票') ");
            sb.Append("              where T1.cardcode=@cardcode  ");
            sb.Append(" group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity,t1.U_ACME_INV");
            sb.Append(" having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append(" order by T0.Docentry desc");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void ViewBatchPayment32(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                         select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV ");
            sb.Append("                             from acmesql02.dbo.pdn1 T0  ");
            sb.Append("                            left join acmesql02.dbo.opdn T1 on (T0.docentry=T1.docentry) ");
            sb.Append("                            left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購')  ");
            sb.Append("                           where T1.cardcode=@cardcode   ");
            sb.Append("               group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity,t1.U_ACME_INV ");
            sb.Append("               having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0 ");
            sb.Append("               order by T0.Docentry desc ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment4(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("            select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T5.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV");
            sb.Append("                           from acmesql02.dbo.pch1 T0 ");
            sb.Append("                          left join acmesql02.dbo.opch T1 on (T0.docentry=T1.docentry)");
            sb.Append(" left join PDN1 t4 on (t0.baseentry=T4.docentry and  t0.baseline=t4.linenum )");
            sb.Append(" left join Por1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum )");
            sb.Append("                          left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='AP發票') ");
            sb.Append("                          where T1.cardcode=@cardcode  ");
            sb.Append("             group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T5.PRICE,T0.quantity,t1.U_ACME_INV");
            sb.Append("             having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0");
            sb.Append("             order by T0.Docentry desc");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void ViewBatchPayment42(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                   select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T5.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV ");
            sb.Append("                                         from acmesql02.dbo.pdn1 T0  ");
            sb.Append("                                        left join acmesql02.dbo.opdn T1 on (T0.docentry=T1.docentry) ");
            sb.Append("               left join Por1 t5 on (t0.baseentry=T5.docentry and  t0.baseline=t5.linenum ) ");
            sb.Append("                                       left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購')  ");
            sb.Append("                                       where T1.cardcode=@cardcode   ");
            sb.Append("                           group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T5.PRICE,T0.quantity,t1.U_ACME_INV ");
            sb.Append("                           having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0 ");
            sb.Append("                           order by T0.Docentry desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " INV1");
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

                        sb.Append("'" + row.Cells["Docentry"].Value.ToString() + " "+ row.Cells["LINENUM"].Value.ToString() + "',");
                    }
                 
             


                    sb.Remove(sb.Length - 1, 1);

                    //linenum
                    string q = sb.ToString();
                    string g = string.Empty;


                    if (radioButton1.Checked == true)
                    {
                        g = "1";
                    }
                    if (radioButton2.Checked == true)
                    {
                        g = "2";
                    }
                    if (radioButton3.Checked == true)
                    {
                        g = "3";
                    }
                    if (radioButton4.Checked == true)
                    {
                        g = "4";
                    }
                    a = q;
                    c = g;
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

              
                    if (radioButton1.Checked == true)
                    {
                        ViewBatchPayment3(cardcode);
                    }
                    if (radioButton2.Checked == true)
                    {
                        ViewBatchPayment2(cardcode);
                    }
                    if (radioButton3.Checked == true)
                    {
                        ViewBatchPayment32(cardcode);
                    }
                if (radioButton4.Checked == true)
                {
                    ViewBatchPayment24(cardcode);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
             
                    if (radioButton1.Checked == true)
                    {
                        ViewBatchPayment1(cardcode);
                    }
                    if (radioButton2.Checked == true)
                    {
                        ViewBatchPayment(cardcode);
                    }

                    if (radioButton3.Checked == true)
                    {
                        ViewBatchPayment12(cardcode);
                    }
                if (radioButton4.Checked == true)
                {
                    ViewBatchPayment4S(cardcode);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void radioButton1_Click(object sender, EventArgs e)
        {
            if (usd == "USD")
            {
                ViewBatchPayment4(cardcode);
            }
            else
            {
                ViewBatchPayment3(cardcode);
            }
            
  
            textBox1.Text = "";
            label1.Text = "AU INV";
        }

        private void radioButton2_Click(object sender, EventArgs e)
        {
            ViewBatchPayment2(cardcode);
            textBox1.Text = "";
            label1.Text = "SAP單號";
        }

        private void radioButton3_Click(object sender, EventArgs e)
        {
            if (usd == "USD")
            {
                ViewBatchPayment42(cardcode);
            }
            else
            {
                ViewBatchPayment32(cardcode);
            }


            textBox1.Text = "";
            label1.Text = "SAP單號";
        }
    }
}
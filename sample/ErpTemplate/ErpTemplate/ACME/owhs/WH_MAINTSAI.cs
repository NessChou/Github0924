using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class WH_MAINTSAI : Form
    {
        public WH_MAINTSAI()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("CARDNAME", typeof(string));
            dt.Columns.Add("INVNO", typeof(string));
            dt.Columns.Add("INVDATE", typeof(string));
            dt.Columns.Add("DOCENTRY", typeof(string));
            dt.Columns.Add("RENT", typeof(string));
            dt.Columns.Add("RUSH", typeof(string));
            dt.Columns.Add("CARDINFO", typeof(string));
            dt.Columns.Add("WHNO", typeof(string));
            dt.Columns.Add("WH", typeof(string));
            dt.Columns.Add("DDATE", typeof(string));
            dt.Columns.Add("RDATE", typeof(string));
            dt.Columns.Add("MEMO", typeof(string));
            dt.Columns.Add("PACK", typeof(string));
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "" && textBox6.Text == "" && textBox7.Text == "" && comboBox2.Text =="")
            {

                MessageBox.Show("請輸入條件");
                return;
            }
            System.Data.DataTable K2 = GetOPDN2();
            {
                if (K2.Rows.Count > 0)
                {
                    for (int i = 0; i <= K2.Rows.Count - 1; i++)
                    {
                        string InvoiceNo = K2.Rows[i][0].ToString();
                        string SIZE = K2.Rows[i][1].ToString();
                        string LB = InvoiceNo.Substring(0, 2);
                        if (LB == "LB")
                        {
                            System.Data.DataTable M1 = util.GETPACLD(InvoiceNo);

                            if (M1.Rows.Count > 0)
                            {
                                for (int i2 = 0; i2 <= M1.Rows.Count - 1; i2++)
                                {
                                    string CCBM = M1.Rows[i2]["CBM"].ToString();

                                    string[] sArray = CCBM.Split('*');
                                    int F2 = 0;
                                    foreach (string F in sArray)
                                    {
                                        F2++;
                                    }
                                    if (F2 > 3)
                                    {
                                        int D = CCBM.LastIndexOf("*");
                                        string CC = CCBM.Substring(0, D);
                                        string PLT = sArray[3];


                                        util.UPDATEPACKLB(CC, PLT, InvoiceNo, CCBM);
                                    }
                                }
                            }
                            string CBMM = "";
                            System.Data.DataTable GF2 = util.GETPACLS2(InvoiceNo);
                            if (GF2.Rows.Count > 0)
                            {
                                CBMM = GF2.Rows[0][0].ToString();
                            }
                            util.GETCBM(InvoiceNo, CBMM);
                        }
                        else
                        {
                            string[] splitStr = { "CM" };
                            string[] arrurl = SIZE.Split(splitStr, StringSplitOptions.RemoveEmptyEntries);
                            string PLT = "";
                            foreach (string ESi in arrurl)
                            {
                                string[] arrurl2 = ESi.Split(new Char[] { '@' });
                                int F = 0;
                                string PLATENO = "";
                                string CBM = "";

                                foreach (string ESi2 in arrurl2)
                                {
                                    F++;
                                    string EA = ESi2;
                                    if (F == 1)
                                    {
                                        PLATENO = EA.Replace(":", "").Replace("No.", "").Trim();
                                    }
                                    if (F == 2)
                                    {
                                        CBM = EA;

                                    }
                                }

                                int pall = PLATENO.IndexOf("PALLET");
                                if (pall != -1)
                                {
                                    try
                                    {
                                        System.Data.DataTable GF1 = util.GETPACLS(PLATENO);
                                        if (GF1.Rows.Count > 0)
                                        {
                                            PLT = GF1.Rows[0][0].ToString();
                                        }
                                    }
                                    catch { }
                                }
                                else
                                {
                                    PLT = "0";
                                }

                               // util.UPDATEPACK(CBM, PLT, InvoiceNo, PLATENO);
                                //util.UPDATEPACK2(CBM, PLT, InvoiceNo, PLATENO);
                            }
                            string CBMM = "";
                            if (PLT != "0")
                            {
                                System.Data.DataTable GF2 = util.GETPACLS2(InvoiceNo);
                                if (GF2.Rows.Count > 0)
                                {
                                    CBMM = GF2.Rows[0][0].ToString();
                                }
                                util.GETCBM(InvoiceNo, CBMM);
                            }

                            System.Data.DataTable GF3 = util.GETPACLS3(InvoiceNo);
                            System.Data.DataTable GF3W = util.GETPACLS3W(InvoiceNo);
                            string F3 = GF3.Rows[0][0].ToString();
                            if (GF3W.Rows.Count > 0)
                            {
                                F3 = GF3W.Rows[0][0].ToString();
                            }
                           
                            util.UPDATEPACKH(F3, InvoiceNo);
                        }
                    }
                }
            }
            System.Data.DataTable dtCost = MakeTable();
            System.Data.DataTable K1 = GetOPDN();
            DataRow dr = null;
            if (K1.Rows.Count > 0)
            {
                for (int i = 0; i <= K1.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    DataRow dd = K1.Rows[i];
                    string JOBNO = dd["JOBNO"].ToString();
                    string INVNO = dd["INVNO"].ToString();
                    dr["JOBNO"] = JOBNO;
                    dr["CARDNAME"] = dd["CARDNAME"].ToString();
                    dr["INVNO"] = INVNO;
                    dr["INVDATE"] = dd["INVDATE"].ToString();
                    dr["DOCENTRY"] = dd["DOCENTRY"].ToString();
                    System.Data.DataTable G1 = GetOPDN2(JOBNO, INVNO);
                    if (G1.Rows.Count > 0)
                    {
                        dr["RENT"] = G1.Rows[0]["RENT"].ToString();
                        dr["RUSH"] = G1.Rows[0]["RUSH"].ToString();
                        dr["DDATE"] = G1.Rows[0]["DDATE"].ToString();
                        dr["RDATE"] = G1.Rows[0]["RDATE"].ToString();
                        dr["MEMO"] = G1.Rows[0]["MEMO"].ToString();
                    }
                    else
                    {
                        string RENT = dd["RENT"].ToString().ToUpper();
                        if (RENT == "CHECKED")
                        {
                            dr["RENT"] = "V";
                        }

                        string RUSH = dd["RUSH"].ToString().ToUpper();
                        if (RUSH == "CHECKED")
                        {
                            dr["RUSH"] = "V";
                        }
                    }

                    dr["CARDINFO"] = dd["CARDINFO"].ToString();
                    dr["WHNO"] = dd["WHNO"].ToString();
                    dr["WH"] = dd["WH"].ToString();
                    dr["PACK"] = dd["PACK"].ToString();
                    dtCost.Rows.Add(dr);
                }

                if (checkBox1.Checked || checkBox2.Checked)
                {
                    if (checkBox1.Checked && checkBox2.Checked)
                    {
                        dtCost.DefaultView.RowFilter = " RENT = 'V' AND RUSH = 'V'";
                    }
                    else if (checkBox1.Checked && checkBox2.Checked ==false)
                    {
                        dtCost.DefaultView.RowFilter = " RENT = 'V' ";
                    }
                    else if (checkBox2.Checked && checkBox1.Checked == false)
                    {
                        dtCost.DefaultView.RowFilter = " RUSH = 'V' ";
                    }
                }
                dataGridView1.DataSource = dtCost;
            }
        }

       
        public System.Data.DataTable GetOPDN()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT T0.DOCENTRY,U_Shipping_no JOBNO,U_ACME_INV INVNO,Convert(varchar(10),U_ACME_INVOICE,102)   INVDATE,T4.WH, T2.WHNO,T2.CARDNAME,T3.CARDINFO ,T2.ADD10 RENT,T4.RUSH,T5.PACK   FROM OPDN T0   ");
            sb.Append("              LEFT JOIN ( SELECT MEMO3 WHNO,SHIPPINGCODE,CardName,ADD10   FROM ACMESQLSP.DBO.SHIPPING_MAIN ) T2 ON (T0.U_Shipping_no=T2.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append("              LEFT JOIN (  SELECT MAX(Remark) CARDINFO,SHIPPINGCODE  FROM ACMESQLSP.DBO.SHIPPING_ITEM GROUP BY SHIPPINGCODE) T3 ON (T0.U_Shipping_no=T3.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append("              LEFT JOIN ( SELECT MAX(RUSH) RUSH,MAX(WHSCODE) WH,SHIPPINGCODE  FROM ACMESQLSP.DBO.LcInstro  GROUP BY SHIPPINGCODE) T4 ON (T0.U_Shipping_no=T4.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("              LEFT JOIN ( SELECT MAX(PLT) PACK,INVOICENO  FROM ACMESQLSP.DBO.rpa_packingH WHERE ISNULL(PLT,'') <>'' GROUP BY INVOICENO) T5 ON ( (T0.U_ACME_INV=T5.INVOICENO COLLATE  Chinese_Taiwan_Stroke_CI_AS OR  SUBSTRING(U_ACME_INV,1,10) =T5.INVOICENO COLLATE  Chinese_Taiwan_Stroke_CI_AS))  ");
            sb.Append(" WHERE 1=1   ");

            if (textBox1.Text != "")
            {
                sb.Append("  AND  T0.U_Shipping_no like '%" + textBox1.Text + "%'");
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND  T0.U_ACME_INV like '%" + textBox2.Text + "%'");
   
            }
            if (textBox3.Text != "")
            {
                sb.Append("  AND  T0.DOCENTRY=@DOCENTRY");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  AND  T2.WHNO like '%" + textBox4.Text + "%'");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("  AND  T4.WH = '" + comboBox2.Text + "'");
            }
            if (textBox6.Text != "" && textBox7.Text != "")
            {
                sb.Append("  AND SUBSTRING(U_Shipping_no,3,8)  BETWEEN @DAY1 AND @DAY2");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
       
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DAY1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@DAY2", textBox7.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public  System.Data.DataTable GetOPDN2()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT U_ACME_INV,SIZE    FROM OPDN T0  ");
            sb.Append(" LEFT JOIN ( SELECT MEMO3 WHNO,SHIPPINGCODE,CardName,ADD10   FROM ACMESQLSP.DBO.SHIPPING_MAIN ) T2 ON (T0.U_Shipping_no=T2.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" LEFT JOIN ( SELECT MAX(RUSH) RUSH,MAX(WHSCODE) WH,SHIPPINGCODE  FROM ACMESQLSP.DBO.LcInstro  GROUP BY SHIPPINGCODE) T4 ON (T0.U_Shipping_no=T4.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" LEFT JOIN ( SELECT MAX(SIZE) SIZE,INVOICENO  FROM ACMESQLSP.DBO.rpa_packingH WHERE ISNULL(SIZE,'') <>'' GROUP BY INVOICENO) T5 ON (T0.U_ACME_INV=T5.INVOICENO COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" WHERE  ISNULL(U_ACME_INV,'') <> ''  ");

            if (textBox1.Text != "")
            {
                sb.Append("  AND  T0.U_Shipping_no like '%" + textBox1.Text + "%'");
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND  T0.U_ACME_INV=@U_ACME_INV");
            }
            if (textBox3.Text != "")
            {
                sb.Append("  AND  T0.DOCENTRY=@DOCENTRY");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  AND  T2.WHNO like '%" + textBox4.Text + "%'");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("  AND  T4.WH = '" + comboBox2.Text + "'");
            }
            if (textBox6.Text != "" && textBox7.Text != "")
            {
                sb.Append("  AND SUBSTRING(U_Shipping_no,3,8)  BETWEEN @DAY1 AND @DAY2");
            }
  
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DAY1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@DAY2", textBox7.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public System.Data.DataTable GetOPDN2(string SHIPPINGCODE,string INVOICE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT * FROM WH_MAINTSAI WHERE SHIPPINGCODE =@SHIPPINGCODE AND INVOICE =@INVOICE ");
       
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                DataGridViewRow row;
                row = dataGridView1.Rows[i];
                string JOBNO = row.Cells["JOBNO"].Value.ToString();
                string INVNO = row.Cells["INVNO"].Value.ToString();
                string RENT = row.Cells["RENT"].Value.ToString();
                string RUSH = row.Cells["RUSH"].Value.ToString();
                string DDATE = row.Cells["DDATE"].Value.ToString();
                string MEMO = row.Cells["MEMO"].Value.ToString();
                string RDATE = row.Cells["RDATE"].Value.ToString();
                DELWH_MAINTSAI(JOBNO, INVNO);
                InsertWH_MAINTSAI(JOBNO, INVNO, RENT, RUSH, DDATE, MEMO, RDATE);
            }
            MessageBox.Show("更新成功");
        }
        
        private void UPDATEINVOICE(string SHIPPINGCODE, string INVOICE, string RENT, string RUSH, string DDATE, string MEMO, string RDATE)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OPDN (SHIPPINGCODE,INVOICE,RENT,RUSH,DDATE,MEMO,RDATE) VALUES(@SHIPPINGCODE,@INVOICE,@RENT,@RUSH,@DDATE,@MEMO,@RDATE)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@RENT", RENT));
            command.Parameters.Add(new SqlParameter("@RUSH", RUSH));
            command.Parameters.Add(new SqlParameter("@DDATE", DDATE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@RDATE", RDATE));
            
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
        private void InsertWH_MAINTSAI(string SHIPPINGCODE, string INVOICE, string RENT, string RUSH, string DDATE, string MEMO, string RDATE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO WH_MAINTSAI (SHIPPINGCODE,INVOICE,RENT,RUSH,DDATE,MEMO,RDATE) VALUES(@SHIPPINGCODE,@INVOICE,@RENT,@RUSH,@DDATE,@MEMO,@RDATE)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@RENT", RENT));
            command.Parameters.Add(new SqlParameter("@RUSH", RUSH));
            command.Parameters.Add(new SqlParameter("@DDATE", DDATE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@RDATE", RDATE));
            
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

        private void DELWH_MAINTSAI(string SHIPPINGCODE, string INVOICE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE WH_MAINTSAI WHERE  SHIPPINGCODE=@SHIPPINGCODE AND INVOICE=@INVOICE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));


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

        private void WH_MAINTSAI_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt4 = GetShipping_WHS();
            comboBox2.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
            textBox6.Text = GetMenu.DFirst();
            textBox7.Text = GetMenu.DLast();
        }
        public static System.Data.DataTable GetShipping_WHS()
        {
            SqlConnection con = globals.Connection;
            string sql = "SELECT WHSCODE DataValue FROM Shipping_WHS union all  select ''  order by WHSCODE  ";
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }



    }
}

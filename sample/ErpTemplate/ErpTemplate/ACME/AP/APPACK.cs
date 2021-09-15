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
    public partial class APPACK : Form
    {

        StringBuilder sbS = new StringBuilder();
        StringBuilder sbS2 = new StringBuilder();
        public APPACK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "" && textBox2.Text == "")
            {
                MessageBox.Show("請輸入條件");
                return;
            }

            System.Data.DataTable G1 = GETPACL();

            if (G1.Rows.Count > 0)
            {
                string FINV = G1.Rows[0]["INVOICENO"].ToString();
                for (int i = 0; i <= G1.Rows.Count - 1; i++)
                {

                    string SIZE = G1.Rows[i]["版數"].ToString();
                    string InvoiceNo = G1.Rows[i]["INVOICENO"].ToString();
                    string LB = InvoiceNo.Substring(0, 2);
                    System.Data.DataTable FF1 = F1(InvoiceNo);
                    if (FF1.Rows.Count > 0)
                    {
                        string QTY = FF1.Rows[0]["QTY"].ToString();
                        string NW = FF1.Rows[0]["NW"].ToString();
                        string GW = FF1.Rows[0]["GW"].ToString();
                        util.UPDATEM(QTY, GW, NW, InvoiceNo);
                    }
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
                                //PALLET24-8-19
                                System.Data.DataTable GF1 = util.GETPACLS(PLATENO);
                                if (GF1.Rows.Count > 0)
                                {
                                    PLT = GF1.Rows[0][0].ToString();
                                }
                            }
                            else
                            {
                                PLT = "0";
                            }

                            util.UPDATEPACK(CBM, PLT, InvoiceNo, PLATENO);
                            util.UPDATEPACK2(CBM, PLT, InvoiceNo, PLATENO);
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

                 

                    }
                }

          

                    System.Data.DataTable GG1 = GETPACL();
                    try
                    {
                    decimal[] TotalG = new decimal[GG1.Columns.Count - 1];

                    for (int i = 0; i <= GG1.Rows.Count - 1; i++)
                    {

                        for (int j = 7; j <= 9; j++)
                        {
                            TotalG[j - 1] += Convert.ToDecimal(GG1.Rows[i][j]);

                        }

                        for (int j = 2; j <= 4; j++)
                        {
                            TotalG[j - 1] += Convert.ToDecimal(GG1.Rows[i][j]);

                        }
                    }


                    DataRow rowG;
                    rowG = GG1.NewRow();
                    rowG[1] = "合計";

                    for (int j = 7; j <= 9; j++)
                    {
                        rowG[j] = TotalG[j - 1];

                    }
                    for (int j = 2; j <= 4; j++)
                    {
                        rowG[j] = TotalG[j - 1];

                    }
                    GG1.Rows.Add(rowG);

                    }
                    catch { }

                    dataGridView1.DataSource = GG1;
              
                System.Data.DataTable G2 = util.GETPACL2B(FINV);
                dataGridView2.DataSource = G2;
            }
        }



        public void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }
        private void SBS()
        {
            string[] arrurl = textBox1.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS.Append("'" + i + "',");
            }
            sbS.Remove(sbS.Length - 1, 1);
        }
        private void SBS2()
        {
            string[] arrurl = textBox2.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS2.Append("'" + i + "',");
            }
            sbS2.Remove(sbS2.Length - 1, 1);
        }

        private System.Data.DataTable F1(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select SUM(CAST(ISNULL(REPLACE(QTY,',',''),0) AS decimal(12,2))) QTY");
            sb.Append("   ,SUM(CAST(ISNULL(REPLACE(NWEIGHT,',',''),0) AS decimal(12,2))) NW");
            sb.Append("   ,SUM(CAST(ISNULL(REPLACE(GWEIGHT,',',''),0) AS DECIMAL(12,2))) GW from RPA_PackingD");
            sb.Append("  WHERE NWEIGHT  LIKE '%[^0-9]%' AND GWEIGHT  LIKE '%[^0-9]%'  AND QTY NOT LIKE '%@%'");
            sb.Append("    AND  InvoiceNo=@InvoiceNo ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETPACL()
        {
            Clear(sbS);
            Clear(sbS2);
            SBS();
            SBS2();
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select DISTINCT INVOICENO,InvoiceDate,TotalQty Qty,TotalNW NW,TotalGW GW,SayTotal,SIZE 版數,CBM,PLT,CARTON from ACMESQLSP.DBO.rpa_packingH T0");
            sb.Append("  LEFT JOIN OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("   WHERE 1=1");

            if (sbS.ToString() != "''")
            {
                sb.Append(" AND InvoiceNo  IN (" + sbS.ToString() + ")   ");
            }
            if (sbS2.ToString() != "''")
            {
                sb.Append("  AND U_Shipping_no  IN (" + sbS2.ToString() + ")  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                string FINV = dataGridView1.SelectedRows[0].Cells["INVOICENO"].Value.ToString();

                System.Data.DataTable G2 = util.GETPACL2(FINV);
                dataGridView2.DataSource = G2;
            }
            catch { }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void APPACK_Load(object sender, EventArgs e)
        {
            label3.Text = "2筆以上請加，逗號";
        }



        
  

    }
}

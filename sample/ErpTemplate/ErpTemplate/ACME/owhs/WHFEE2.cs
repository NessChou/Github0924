using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.IO;

namespace ACME
{
    public partial class WHFEE2 : Form
    {
        string CYEAR = Convert.ToString(Convert.ToInt32(DateTime.Now.Year) - 1911);
        string CMON = Convert.ToString(Convert.ToInt16(DateTime.Now.Month));
        string 進貨費用 = "";
        string SIN = "";
        string FA = "acmesql02";
        System.Data.DataTable dtCost = null;
        System.Data.DataTable dtCostF = null;
        string TOTALA = "";
        string TOTALB = "";
        string TOTALC = "";

        string 聯倉租 = "";
        string 聯倉坪 = "";
        string 聯出倉明細 = "";
        string 聯加班費 = "";
        string 新得利進出費用 = "";
        string 新得利加班費用 = "";
        string 出倉理貨費 = "";
        int NANFF1 = 0;
        int NANFF2 = 0;
        private string FileName;

        string FTAI = "";
        string FTAI2 = "";
        string FTU = "";
        string FTU2 = "";
        string FZU = "";
        string FZU2 = "";
        string FRU = "";
        string FRU2 = "";
        decimal FRU3 = 0;
        public WHFEE2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable TempDt = MakeTable();
            System.Data.DataTable dt = GetG1();
            DataRow dr = null;
            string DUP = "";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = TempDt.NewRow();

                string WHNO = dt.Rows[i]["放貨單號"].ToString();
                dr["出貨日期"] = dt.Rows[i]["出貨日期"].ToString();
                dr["放貨單號"] = WHNO;
                dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                dr["數量"] = dt.Rows[i]["數量"].ToString();
                if (DUP != WHNO)
                {
                    dr["出貨客戶"] = dt.Rows[i]["出貨客戶"].ToString();
                    dr["板數"] = dt.Rows[i]["板數"].ToString();
                    dr["車種"] = dt.Rows[i]["車種"].ToString();
                    dr["箱數"] = dt.Rows[i]["箱數"].ToString();
                    dr["快遞寄件數"] = dt.Rows[i]["快遞寄件數"].ToString();
                    dr["運費"] = dt.Rows[i]["運費"].ToString();
                    dr["打版數"] = dt.Rows[i]["打版數"].ToString();
                    dr["貼麥頭數"] = dt.Rows[i]["貼麥頭數"].ToString();
                    dr["打邊條數"] = dt.Rows[i]["打邊條數"].ToString();
                    dr["裝櫃費"] = dt.Rows[i]["裝櫃費"].ToString();
                    dr["加班費"] = dt.Rows[i]["加班費"].ToString();
                    dr["出倉費用"] = dt.Rows[i]["出倉費用"].ToString();
                    dr["進倉費用"] = dt.Rows[i]["進倉費用"].ToString();
                }


                DUP = WHNO;
            //string DOCDATE = dt.Rows[i]["出貨日期"].ToString();
            //string WHNO = dt.Rows[i]["放貨單號"].ToString();
            //string ITEMCODE = dt.Rows[i]["產品編號"].ToString();
            //                string CARDNAME = dt.Rows[i]["出貨客戶"].ToString();
            //string QTY = dt.Rows[i]["數量"].ToString();
            //string PQTY = dt.Rows[i]["版數"].ToString();
            //string CT = dt.Rows[i]["車種"].ToString();
            //string CQTY = dt.Rows[i]["箱數"].ToString();
            //string PQTY3 = dt.Rows[i]["快遞寄件數"].ToString();
            //string FEE = dt.Rows[i]["運費"].ToString();
            //string CBIN = dt.Rows[i]["CBIN"].ToString();
             //   AddATC1(DOCDATE, WHNO, ITEMCODE, QTY, CARDNAME, PQTY, CT, PQTY3, FEE,CBIN,"");
     
                TempDt.Rows.Add(dr);
            }

            dataGridView1.DataSource = TempDt;

            decimal[] Total = new decimal[TempDt.Columns.Count - 1];

            for (int i = 0; i <= TempDt.Rows.Count - 1; i++)
            {

                for (int j = 9; j <= TempDt.Columns.Count - 1; j++)
                {
                    string DA = TempDt.Rows[i][j].ToString();
                    if (String.IsNullOrEmpty(DA))
                    {
                        DA = "0";
                    }
                    Total[j - 1] += Convert.ToDecimal(DA);

                }
            }

            DataRow row;

            row = TempDt.NewRow();
            row[8] = "合計";
            for (int j = 9; j <= TempDt.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            TempDt.Rows.Add(row);
        }

        public void AddATC1(string DOCDATE, string WHNO, string ITEMCODE, string QTY, string CARDNAME, string PQTY, string CQTY, string PQTY3, string FEE, string AWHNO, string OWHNO)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into WH_FEE2(DOCDATE,WHNO,ITEMCODE,QTY,CARDNAME,PQTY,CQTY,PQTY3,FEE,AWHNO,OWHNO,USERS) values(@DOCDATE,@WHNO,@ITEMCODE,@QTY,@CARDNAME,@PQTY,@CQTY,@PQTY3,@FEE,@AWHNO,@OWHNO,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@PQTY", PQTY));
            command.Parameters.Add(new SqlParameter("@CQTY", CQTY));
            command.Parameters.Add(new SqlParameter("@PQTY3", PQTY3));
            command.Parameters.Add(new SqlParameter("@FEE", FEE));
            command.Parameters.Add(new SqlParameter("@AWHNO", AWHNO));
            command.Parameters.Add(new SqlParameter("@OWHNO", OWHNO));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

     
 
        private DataTable GetG1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.QUANTITY 出貨日期,T0.ShippingCode 放貨單號,T2.ItemCode 產品編號,ISNULL(T2.Quantity,0) 數量,T0.CARDNAME 出貨客戶  ");
            sb.Append(" ,PQTY 板數,cT 車種,T1.KQTY 箱數,kQTY3 快遞寄件數,kT2 運費,CBIN,lOUD 打版數,lOUT 貼麥頭數,lOUB 打邊條數,lOUF 裝櫃費");
            sb.Append(" ,lOUG 加班費,lOUFEE 出倉費用,lINFEE  進倉費用  FROM WH_MAIN T0   ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)  ");
            sb.Append(" LEFT JOIN  WH_ITEM T2 ON (T0.ShippingCode=T2.ShippingCode)  ");
            sb.Append(" WHERE (C1=@C1 OR LOUC=@C1 OR LINC =@C1) AND  SUBSTRING(T0.ShippingCode,3,8) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" ORDER BY T0.ShippingCode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@C1", comboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private DataTable GetG2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM Rma_mainR WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }

        private DataTable GetG3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }


        private DataTable GetG4(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("放貨單號", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("數量", typeof(int));
            dt.Columns.Add("出貨客戶", typeof(string));
            dt.Columns.Add("板數", typeof(string));
            dt.Columns.Add("車種", typeof(string));
            dt.Columns.Add("箱數", typeof(string));
            dt.Columns.Add("快遞寄件數", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("打版數", typeof(string));
            dt.Columns.Add("貼麥頭數", typeof(string));
            dt.Columns.Add("打邊條數", typeof(string));
            dt.Columns.Add("裝櫃費", typeof(string));
            dt.Columns.Add("加班費", typeof(string));
            dt.Columns.Add("出倉費用", typeof(string));
            dt.Columns.Add("進倉費用", typeof(string));
            dt.Columns.Add("備註", typeof(string));


            return dt;
        }

        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("放貨單號", typeof(string));
            dt.Columns.Add("項次", typeof(string));
            dt.Columns.Add("產品編號", typeof(int));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("AUINV", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單位", typeof(string));
            dt.Columns.Add("出貨客戶", typeof(string));
            dt.Columns.Add("派送地", typeof(string));
            dt.Columns.Add("板數", typeof(string));
            dt.Columns.Add("箱數", typeof(string));
            dt.Columns.Add("快遞寄件數", typeof(string));
            dt.Columns.Add("卡車噸數", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("打板數", typeof(string));
            dt.Columns.Add("打邊條數", typeof(string));
            dt.Columns.Add("貼麥頭數", typeof(string));
            dt.Columns.Add("搬運費", typeof(string));
            dt.Columns.Add("加班費", typeof(string));
            return dt;
        }


        private void WHFEE_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "21";
            textBox2.Text = DateTime.Now.ToString("yyyyMM") + "20";

            textBox3.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "21";
            textBox4.Text = DateTime.Now.ToString("yyyyMM") + "20";

            comboBox1.Text = "新得利";
            comboBox2.Text = "新得利";
            dtCostF = MakeTableF2();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataRow dr = null;
            //深圳宏高
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
             
                if (comboBox2.Text == "聯倉")
                {
                    TRUNTABLE();
                    WriteExcelProduct4(FileName);

                    System.Data.DataTable GF3 = Getdata9B();
                    dtCost = MakeTableF();
                    if (GF3.Rows.Count > 0)
                    {

                        string 產品編號 = GF3.Rows[0]["產品編號"].ToString();
                        string 產品名稱 = GF3.Rows[0]["產品名稱"].ToString();
                        string 費用 = GF3.Rows[0]["費用"].ToString();
                        string 倉庫 = GF3.Rows[0]["倉庫"].ToString();


                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0193";
                        dr["廠商名稱"] = "聯倉交通股份有限公司";
                        dr["產品編號"] = 產品編號;
                        dr["費用名稱"] = 產品名稱;
                        dr["輔助金額"] = TOTALC;
                        //           dr["輔助金額"] = 費用;
                        dr["EXCEL金額"] = TOTALC;
                        dr["SAP金額"] = TOTALC;
                        dr["倉庫"] = 倉庫;

                        dtCost.Rows.Add(dr);
                    }


                    System.Data.DataTable GF4 = Getdata9B2();
        
                    if (GF4.Rows.Count > 0)
                    {

                        string 產品編號 = GF4.Rows[0]["產品編號"].ToString();
                        string 產品名稱 = GF4.Rows[0]["產品名稱"].ToString();
                        string 費用 = GF4.Rows[0]["費用"].ToString();
                        string 倉庫 = GF4.Rows[0]["倉庫"].ToString();


                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0193";
                        dr["廠商名稱"] = "聯倉交通股份有限公司";
                        dr["產品編號"] = 產品編號;
                        dr["費用名稱"] = 產品名稱;
                        dr["輔助金額"] = 聯出倉明細;
                        dr["EXCEL金額"] = 聯出倉明細;
                        dr["SAP金額"] = 聯出倉明細;
                        dr["倉庫"] = 倉庫;

                        dtCost.Rows.Add(dr);
                    }

                    if (!String.IsNullOrEmpty(聯加班費))
                    {
                        if (聯加班費 != "0")
                        {
                            dr = dtCost.NewRow();
                            dr["廠商編號"] = "U0193";
                            dr["廠商名稱"] = "聯倉交通股份有限公司";
                            dr["產品編號"] = "ZA0SZ0701";
                            dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 加班/急單費用";
                            dr["輔助金額"] = 聯加班費;
                            dr["EXCEL金額"] = 聯加班費;
                            dr["SAP金額"] = 聯加班費;
                            dr["倉庫"] = "Z0011";

                            dtCost.Rows.Add(dr);
                        }
                    }
                    dr = dtCost.NewRow();
                    dr["廠商編號"] = "U0193";
                    dr["廠商名稱"] = "聯倉交通股份有限公司";
                    dr["產品編號"] = "ZA0SF0005";
                    dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 倉租費-" + 聯倉坪+ "坪 TW LOCAL";
                    dr["輔助金額"] = 聯倉租;
                    dr["EXCEL金額"] = 聯倉租;
                    dr["SAP金額"] = 聯倉租;
                    dr["倉庫"] = "Z0009";
                    dtCost.Rows.Add(dr);


                    dr = dtCost.NewRow();
                    dr["廠商編號"] = "U0193";
                    dr["廠商名稱"] = "聯倉交通股份有限公司";
                    dr["產品編號"] = "ZA0SZ0400";
                    dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 出倉理貨費";
                    dr["輔助金額"] = 出倉理貨費;
                    dr["EXCEL金額"] = 出倉理貨費;
                    dr["SAP金額"] = 出倉理貨費;
                    dr["倉庫"] = "Z0011";
                    dtCost.Rows.Add(dr);


                    string SIN2 = Math.Round((Convert.ToDouble(SIN) * 0.95),0, MidpointRounding.AwayFromZero).ToString();
                    string SIN3 = Math.Round((Convert.ToDouble(SIN) * 0.05), 0, MidpointRounding.AwayFromZero).ToString();
                    textBox5.Text = "TFT-" + CMON + "月份倉儲理貨費" +
                                             Environment.NewLine + "4.原價運費為" + SIN + " 打95折為" + SIN2 + "=省" + SIN3 + "#" +
           Environment.NewLine + "發票號碼:";
                }

                if (comboBox2.Text == "台南聯倉")
                {

                    WriteExcelNAN(FileName);

   
                    dtCost = MakeTableF();
           

             
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0193";
                        dr["廠商名稱"] = "聯倉交通股份有限公司";
                        dr["產品編號"] = "ZA0SB0005";

                        dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 出倉卡車費-TW LOCAL";
                        dr["輔助金額"] = NANFF2;
                        dr["EXCEL金額"] = NANFF1;
                        dr["SAP金額"] = NANFF1;
                        dr["倉庫"] = "Z0014";

                        dtCost.Rows.Add(dr);
        
                }


                if (comboBox2.Text == "大發")
                {

                    WriteExcelDAFA(FileName);


                    dtCost = MakeTableF();

                    
                    if (FRU != "")
                    {
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0361";
                        dr["廠商名稱"] = "嘉里大榮物流股份有限公司";
                        dr["產品編號"] = "ZA0SZ0400";

                        dr["費用名稱"] = FRU2;
                        dr["輔助金額"] = FRU;
                        dr["EXCEL金額"] = FRU;
                        dr["SAP金額"] = FRU;
                        dr["倉庫"] = "Z0011";
                        dtCost.Rows.Add(dr);
                    }
                    if (FTU != "")
                    {
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0361";
                        dr["廠商名稱"] = "嘉里大榮物流股份有限公司";
                        dr["產品編號"] = "ZA0SZ0400";

                        dr["費用名稱"] = FTU2;
                        dr["輔助金額"] = FTU;
                        dr["EXCEL金額"] = FTU;
                        dr["SAP金額"] = FTU;
                        dr["倉庫"] = "Z0011";
                        dtCost.Rows.Add(dr);
                    }

                    if (FZU != "")
                    {
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0361";
                        dr["廠商名稱"] = "嘉里大榮物流股份有限公司";
                        dr["產品編號"] = "ZA0SF0005";

                        dr["費用名稱"] = FZU2;
                        dr["輔助金額"] = FZU;
                        dr["EXCEL金額"] = FZU;
                        dr["SAP金額"] = FZU;
                        dr["倉庫"] = "Z0009";
                        dtCost.Rows.Add(dr);
                    }


                    if (FTAI!= "")
                    {
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0361";
                        dr["廠商名稱"] = "嘉里大榮物流股份有限公司";
                        dr["產品編號"] = "ZA0SZ0400";

                        dr["費用名稱"] = FTAI2;
                        dr["輔助金額"] = FTAI;
                        dr["EXCEL金額"] = FTAI;
                        dr["SAP金額"] = FTAI;
                        dr["倉庫"] = "Z0011";
                        dtCost.Rows.Add(dr);
                    }
                }

                if (comboBox2.Text == "新得利")
                {

                    WriteExcelProduct5(FileName);

                    System.Data.DataTable GF1 = Getdata7();
                    dtCost = MakeTableF();
                    if (GF1.Rows.Count > 0)
                    {
                        string 產品編號 = GF1.Rows[0]["產品編號"].ToString();
                        string 產品名稱 = GF1.Rows[0]["產品名稱"].ToString();
                        string 費用 = GF1.Rows[0]["費用"].ToString();
                        string 倉庫 = GF1.Rows[0]["倉庫"].ToString();
           
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0447";
                        dr["廠商名稱"] = "新得利倉儲股份有限公司";
                        dr["產品編號"] = 產品編號;
                        dr["費用名稱"] = 產品名稱;
                        dr["輔助金額"] = 費用;
                        dr["EXCEL金額"] = TOTALA;
                        dr["SAP金額"] = 費用;
                        dr["倉庫"] = 倉庫;

                        dtCost.Rows.Add(dr);
                    }
                    //                    dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 倉租費-" + 聯倉坪+ "坪 TW LOCAL";
                    dr = dtCost.NewRow();
                    dr["廠商編號"] = "U0447";
                    dr["廠商名稱"] = "新得利倉儲股份有限公司";
                    dr["產品編號"] = "ZA0SZ0400";
                    //TFT-109年3月 進出倉理貨費
                    dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月 進出倉理貨費";
                    dr["輔助金額"] = 新得利進出費用;
                    dr["EXCEL金額"] = 新得利進出費用;
                    dr["SAP金額"] = 新得利進出費用;
                    dr["倉庫"] = "Z0011";

                    dtCost.Rows.Add(dr);

                    System.Data.DataTable GF2 = Getdata8();
                    if (GF2.Rows.Count > 0)
                    {

                        string 產品編號 = GF2.Rows[0]["產品編號"].ToString();
                        string 產品名稱 = GF2.Rows[0]["產品名稱"].ToString();
                        string 費用 = GF2.Rows[0]["費用"].ToString();
                        string 倉庫 = GF2.Rows[0]["倉庫"].ToString();
                
         
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0447";
                        dr["廠商名稱"] = "新得利倉儲股份有限公司";
                        dr["產品編號"] = 產品編號;
                        dr["費用名稱"] = 產品名稱;
                        dr["輔助金額"] = TOTALB;
                        dr["EXCEL金額"] = TOTALB;
                        dr["SAP金額"] = TOTALB;
                        dr["倉庫"] = 倉庫;

                        dtCost.Rows.Add(dr);
                    }

                    System.Data.DataTable GF3 = Getdata9();
                    if (GF3.Rows.Count > 0)
                    {

                        string 產品編號 = GF3.Rows[0]["產品編號"].ToString();
                        string 產品名稱 = GF3.Rows[0]["產品名稱"].ToString();
                        string 費用 = GF3.Rows[0]["費用"].ToString();
                        string 倉庫 = GF3.Rows[0]["倉庫"].ToString();

                        if (費用 != "0")
                        {
                            dr = dtCost.NewRow();
                            dr["廠商編號"] = "U0447";
                            dr["廠商名稱"] = "新得利倉儲股份有限公司";
                            dr["產品編號"] = 產品編號;
                            dr["費用名稱"] = 產品名稱;
                            dr["輔助金額"] = TOTALC;
                            dr["EXCEL金額"] = TOTALC;
                            dr["SAP金額"] = TOTALC;
                            dr["倉庫"] = 倉庫;

                            dtCost.Rows.Add(dr);
                        }


                        textBox5.Text = "TFT-" + CMON + "月份倉儲理貨費" +
          Environment.NewLine + "發票號碼:";
                    }

                    if (新得利加班費用 != "")
                    {
                        dr = dtCost.NewRow();
                        dr["廠商編號"] = "U0447";
                        dr["廠商名稱"] = "新得利倉儲股份有限公司";
                        dr["產品編號"] = "ZA0SZ0701";
                        dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月加班/急單費用";
                        dr["輔助金額"] = 新得利加班費用;
                        dr["EXCEL金額"] = 新得利加班費用;
                        dr["SAP金額"] = 新得利加班費用;
                        dr["倉庫"] = "Z0011";
                        dtCost.Rows.Add(dr);
                    }

                 

                }
                if (comboBox2.Text == "航通快遞")
                {
                    DELFEE3();
                    WriteExcelProduct6(FileName);

                    System.Data.DataTable GF1 = Getdata11("新得利");
         
                    if (GF1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= GF1.Rows.Count - 1; i++)
                        {
                            string QTY = GF1.Rows[i]["QTY"].ToString();
                            string PRICE = GF1.Rows[i]["PRICE"].ToString();
                            string AMT = GF1.Rows[i]["AMT"].ToString();


                            dr = dtCostF.NewRow();
                            dr["廠商編號"] = "U0224";
                            dr["廠商名稱"] = "航通興業有限公司";
                            dr["產品編號"] = "ZA0TE0105";
                            dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月份 新得利TFT-快遞費-貨件內銷";
                            dr["輔助金額"] = AMT;
                            dr["EXCEL金額"] = AMT;
                            dr["SAP金額"] = AMT;
                            dr["數量"] = QTY;
                            dr["單價"] = PRICE;
                            dr["倉庫"] = "Z0013";

                            dtCostF.Rows.Add(dr);
                        }
       
                    }


                    System.Data.DataTable GF11 = Getdata11("聯揚");

                    if (GF11.Rows.Count > 0)
                    {
                        for (int i = 0; i <= GF11.Rows.Count - 1; i++)
                        {
                            string QTY = GF11.Rows[i]["QTY"].ToString();
                            string PRICE = GF11.Rows[i]["PRICE"].ToString();
                            string AMT = GF11.Rows[i]["AMT"].ToString();


                            dr = dtCostF.NewRow();
                            dr["廠商編號"] = "U0224";
                            dr["廠商名稱"] = "航通興業有限公司";
                            dr["產品編號"] = "ZA0TE0105";
                            dr["費用名稱"] = "TFT-" + CYEAR + "年" + CMON + "月份 聯揚TFT-快遞費-貨件內銷";
                            dr["輔助金額"] = AMT;
                            dr["EXCEL金額"] = AMT;
                            dr["SAP金額"] = AMT;
                            dr["數量"] = QTY;
                            dr["單價"] = PRICE;
                            dr["倉庫"] = "Z0013";

                            dtCostF.Rows.Add(dr);
                        }

                    }
                    System.Data.DataTable GF2 = Getdata12("新得利");
                    System.Data.DataTable GF3 = Getdata12("聯揚");

                    textBox5.Text = CMON + "月份新倉" + GF2.Rows[0][0].ToString() + "件" +
      Environment.NewLine + "發票號碼:" +
            Environment.NewLine + CMON + "月份聯倉" + GF3.Rows[0][0].ToString() + "件" +
      Environment.NewLine + "發票號碼:";

                }
                if (comboBox2.Text == "航通快遞")
                {
                    dataGridView2.DataSource = dtCostF;
                }
                else
                {
                    dataGridView2.DataSource = dtCost;
                }

            }
        }

        public void TRUNTABLE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table WH_TEMP2 ", connection);
            command.CommandType = CommandType.Text;
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
        public void AddTEMP(string SHIPPINGCODE, string AMT, string DOCDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_TEMP2(SHIPPINGCODE,AMT,DOCDATE) values(@SHIPPINGCODE,@AMT,@DOCDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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
        private System.Data.DataTable MakeTableF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("費用名稱", typeof(string));
            dt.Columns.Add("輔助金額", typeof(string));
            dt.Columns.Add("EXCEL金額", typeof(string));
            dt.Columns.Add("SAP金額", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableF2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("費用名稱", typeof(string));
            dt.Columns.Add("輔助金額", typeof(string));
            dt.Columns.Add("EXCEL金額", typeof(string));
            dt.Columns.Add("SAP金額", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            return dt;
        }
        private void WriteExcelProduct4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string DOCDATE;

                string WHNO;
                string FEE;

                string FEE5;
                string FEE2;
                string FEE3;
                string FEE4;
                string FEE6;
                string FEE7;
                string TS;

                string FW1 = "0";
                string FW2 = "0";
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    FEE = range.Text.ToString().Trim();

                    System.Data.DataTable GG1 = Getdata3(WHNO);
                    if (GG1.Rows.Count > 0)
                    {

                        string FEES = GG1.Rows[0]["FEE"].ToString();
                        DateTime D1 = Convert.ToDateTime(DOCDATE);
                        DateTime D2 = Convert.ToDateTime(GG1.Rows[0]["DOCDATE"]);
                        if (FEE != FEES)
                        {
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            range.ClearComments();
                            string MM = "輔助金額 : " + FEES;
                            range.AddComment(MM);

                            int wCount = CountText(MM, '\n');
                            range.Comment.Shape.Height = wCount * 20;
                        }

                        if (D1 != D2)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            range.ClearComments();
                            string MM = "日期 " + D2.ToString("yyyyMMdd");
                            range.AddComment(MM);

                            int wCount = CountText(MM, '\n');
                            range.Comment.Shape.Height = wCount * 20;
                        }



                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    TS = range.Text.ToString().Trim();
                    int T1 = TS.IndexOf("總計");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    TS = range.Text.ToString().Trim();
                    int T2 = TS.IndexOf("總計");
                    if (T1 != -1 || T2 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        range.Select();
                        TOTALC = range.Text.ToString().Trim().Replace(",", "");


                    }

                }

                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                excelSheet2.Activate();

                int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
                int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;
                string FEET = "";
                string DUP = "";
                string CAR = "";
                for (int iRecord = 2; iRecord <= iRowCnt2; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();
                    //if (WHNO == "WH20200520014X")
                    //{
                    //    MessageBox.Show("a");
                    //}


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    CAR = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    FEE = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    FEE2 = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    FEE3 = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    FEE4 = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    FEE5 = range.Text.ToString().Trim().Replace(",", "");


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    FEE7 = range.Text.ToString().Trim().Replace(",", "");

                    if (DOCDATE == "")
                    {
                        DOCDATE = DUP;
                    }
                    if (CAR.IndexOf("折") != -1 )
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 11]);
                        range.Select();
                        FEE = range.Text.ToString().Trim();
                        int J1 = FEE.LastIndexOf("=");
                        聯出倉明細 = FEE.Substring(J1 + 1, FEE.Length - J1 - 1);
                       
                    }
                    if (FEE.IndexOf("折") != -1)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 12]);
                        range.Select();
                        FEE = range.Text.ToString().Trim();
                        int J1 = FEE.LastIndexOf("=");
                        聯出倉明細 = FEE.Substring(J1 + 1, FEE.Length - J1 - 1).Replace(",", "");

                    }
                    if (CAR.IndexOf("月運費") != -1)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 18]);
                        range.Select();
                        FEE6 = range.Text.ToString().Trim();

                        聯加班費 = FEE6;

                    }
                    if (CAR.IndexOf("月運費") != -1)
                    {
                        if (String.IsNullOrEmpty(FEE2))
                        {
                            FEE2 = "0";
                        }
                        SIN = Convert.ToString(Convert.ToInt32(FEE) + Convert.ToInt32(FEE2) + Convert.ToInt32(FEE7));
                        if (!String.IsNullOrEmpty(FEE3))
                        {
                            出倉理貨費 = Convert.ToString(Convert.ToInt32(FEE3) + Convert.ToInt32(FEE4) + Convert.ToInt32(FEE5));
                        }
                    }
         

                    if (!String.IsNullOrEmpty(FEE) && !String.IsNullOrEmpty(WHNO))
                    {
                        System.Data.DataTable GG1 = Getdata4(WHNO);
                        if (GG1.Rows.Count > 0)
                        {

                            DateTime D1 = Convert.ToDateTime(DOCDATE);
                            DateTime D2 = Convert.ToDateTime(GG1.Rows[0]["DOCDATE"]);
                            string FEES = GG1.Rows[0]["FEE"].ToString();
                            if (String.IsNullOrEmpty(FEE2))
                            {
                                FEE2 = "0";
                            }

                            string FT = Convert.ToString(Convert.ToInt16(FEE) + Convert.ToInt16(FEE2));
                            FW1 = FT;
                            if (FT != FEES)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 12]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : " + FEES;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }

                            if (D1 != D2)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 1]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "日期 " + D2.ToString("yyyyMMdd");
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }

                            DUP = DOCDATE;




                        }
                    }

      
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    TS = range.Text.ToString().Trim();
                    int T1 = TS.IndexOf("運費");
                    if (T1 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 12]);
                        range.Select();
                        TOTALB = range.Text.ToString().Trim().Replace(",", "");


                    }

                    if (!String.IsNullOrEmpty(WHNO))
                    {
                        if (String.IsNullOrEmpty(FEE3))
                        {
                            FEE3 = "0";
                        }
                        if (String.IsNullOrEmpty(FEE4))
                        {
                            FEE4 = "0";
                        }
                        if (String.IsNullOrEmpty(FEE5))
                        {
                            FEE5 = "0";
                        }

                        string FT = Convert.ToString(Convert.ToInt16(FEE3) + Convert.ToInt16(FEE4) + Convert.ToInt16(FEE5));
                        FW2 = FT;
                        System.Data.DataTable GG1 = Getdata5(WHNO);
                        if (GG1.Rows.Count > 0)
                        {
                            string FEES = GG1.Rows[0]["FEE"].ToString();
                            DateTime D1 = Convert.ToDateTime(DOCDATE);
                            DateTime D2 = Convert.ToDateTime(GG1.Rows[0]["DOCDATE"]);

                            if (FT != FEES)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 14]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : " + FEES;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }
                            if (D1 != D2)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 1]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "日期 " + D2.ToString("yyyyMMdd");
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }

                        }
                        else
                        {

                            if (FT != "0")
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 13]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : ";
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }
                        }


                        if (WHNO.IndexOf("WH") != -1)
                        {
                            System.Data.DataTable FE = null;

                            if (FW1 != "0")
                            {
                                FE = Getdata4WH(WHNO);
                            }

                            if (FW2 != "0")
                            {
                                FE = Getdata4WH2(WHNO);
                            }

                            if (FE.Rows.Count == 0)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 2]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "倉庫不是聯倉";
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;


                            }

                        }
                    }

                }

                Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
                excelSheet3.Activate();
                int M1 = 0;
                int M2 = 0;
                int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
                int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;
                for (int b = 1; b <= 700; b++)
                {
                    for (int jj = 1; jj <= 5; jj++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[jj, b]);
                        range.Select();
                        string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

                        string FS = CYEAR + "." + CMON + "月";
                        int G1 = id.IndexOf(FS);
                        if (G1 != -1)
                        {

                            M2 = jj + 2;
                            M1 = b + 2;
                            break;
                        }

                    }

                    if (M1 > 0)
                    {
                        break;
                    }
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[37, M1]);
                range.Select();
                聯倉坪 = range.Text.ToString().Trim().ToUpper().Replace(" ", "").Replace(",", "");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[38, M1]);
                range.Select();
                聯倉租 = range.Text.ToString().Trim().ToUpper().Replace(" ", "").Replace(",", "");
                //for (int iRecord = M2; iRecord <= iRowCnt3; iRecord++)
                //{

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[38, M1]);
                //    range.Select();
                //    string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");



                //}

            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


                System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelNAN(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string LOC;
                string CAR;
                string WHNO;
                string FEE;
                int J1 = 0;
                StringBuilder sb = new StringBuilder();
                string M2 = "TFT-" + CYEAR + "年-" + CMON + "月份聯倉台南出貨卡車費";
                sb.Append(M2);
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {
                    J1++;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();


                    if (!String.IsNullOrEmpty(WHNO))
                    {

                        System.Data.DataTable GG1 = GetdataNAN1(WHNO);
                        if (GG1.Rows.Count > 0)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                            range.Select();
                            FEE = range.Text.ToString().Trim().Replace(",", "");


                            NANFF1 += Convert.ToInt32(FEE);
                            NANFF2 += Convert.ToInt32(GG1.Rows[0][0]);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            range.Select();
                            CAR = range.Text.ToString().Trim();
                            int GR = CAR.ToUpper().IndexOf("T");
                            if (GR != -1)
                            {
                                CAR = CAR.Substring(0, GR + 1);
                            }


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                            range.Select();
                            LOC = range.Text.ToString().Trim();

                            string M1 = J1.ToString() + "." + CAR + "*1台(" + FEE.ToString() + ")" + LOC;
                            sb.Append(Environment.NewLine + M1);
                        }
                    }


                }
                sb.Append(Environment.NewLine + "發票號碼:");

                textBox5.Text = sb.ToString();


            }
            finally
            {



                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();




            }



        }
        private void WriteExcelDAFA(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;
            Microsoft.Office.Interop.Excel.Range range = null;
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            foreach (Microsoft.Office.Interop.Excel.Worksheet excelSheet in excelBook.Worksheets)
            {
                excelSheet.Activate();
                string NAME = excelSheet.Name.ToString().Trim();
                if (NAME == "請款總明細")
                {
                    int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                    int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                    string DNAME;

                    int J1 = 0;
                    decimal FT = 0;
                    StringBuilder sb = new StringBuilder();
                    string FUZF4 = "";
                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                        range.Select();
                        DNAME = range.Text.ToString().Trim();
                        if (DNAME.IndexOf("夏月電價政策調漲倉租") != -1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            FUZF4 = range.Text.ToString().Trim();
                        }
                    }
                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                        J1++;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                        range.Select();
                        DNAME = range.Text.ToString().Trim();
                        //夏月電價政策調漲倉租2%

              
                        if (DNAME.IndexOf("月倉租") != -1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            string FUZF1 = range.Text.ToString().Trim();

                            int pin2 = 0;

                            int pinf = 0;
                            if (FUZF4 != "")
                            {
                                pinf = Convert.ToInt32(FUZF4);
                                 pin2 = Convert.ToInt32(FUZF4) / Convert.ToInt32(FUZF1);
                            }
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            range.Select();
                            string FUZF2 = range.Text.ToString().Trim();
                            int pin3 = pin2 + Convert.ToInt32(FUZF2);
                            //TFT-109/4月 倉租費-350*4坪 TW LOCA 
                            FZU2 = "TFT-" + CYEAR + "/" + CMON + "月 倉租費-" + pin3.ToString() +"*" + FUZF1 + "坪 TW LOCAL";
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            FZU = range.Text.ToString().Trim().Replace(",", "").Replace("-", "");
                            int pinf2 = pinf + Convert.ToInt32(FZU);
                            FZU = pinf2.ToString();
                        }

                        if (DNAME.IndexOf("月入庫費用") != -1)
                        {
                            //TFT-109/4月 進倉理貨費--27板*100
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            string FUZF1 = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            range.Select();
                            string FUZF2 = range.Text.ToString().Trim();

                            FRU2 = "TFT-" + CYEAR + "/" + CMON + "月 進倉理貨費--" + FUZF1 + "板*" + FUZF2;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            FRU = range.Text.ToString().Trim().Replace(",", "").Replace("-", "");
                        }

                        if (DNAME.IndexOf("月出庫費用") != -1)
                        {
                            //TFT-109/4月出倉理貨費--11板*130
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            string FUZF1 = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            range.Select();
                            string FUZF2 = range.Text.ToString().Trim();

                            FTU2 = "TFT-" + CYEAR + "/" + CMON + "月 出倉理貨費--" + FUZF1 + "板*" + FUZF2;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            FTU = range.Text.ToString().Trim().Replace(",", "").Replace("-", "");
                        }

                        if (DNAME.IndexOf("月裝拆櫃費用") != -1)
                        {

                            //TFT-109/4月 卸裝櫃費1*20'*2500
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            string FUZF1 = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            range.Select();
                            string FUZF2 = range.Text.ToString().Trim().Replace(",", "").Replace("-", "");

                            if (FUZF2 == "2500")
                            {
                                FTAI2 = "TFT-" + CYEAR + "/" + CMON + "月 卸裝櫃費" + FUZF1 + "*20'*2500";
                            }



                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            FTAI = range.Text.ToString().Trim().Replace(",", "").Replace("-", "");
                            decimal n;
                            if (decimal.TryParse(FTAI, out n))
                            {

                                FT += Convert.ToDecimal(FTAI);
                            }
                            // if
                        }
                        //月裝拆櫃費用




                    }
               string M1=    "TFT-" + CYEAR + "年/" + CMON + "月倉儲理貨費";
                    sb.Append(M1);
                    sb.Append(Environment.NewLine + "發票號碼:");

                    textBox5.Text = sb.ToString();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                }
                if (NAME == "入庫費")
                {
                    int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                    int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                    string WHNO;

                    int J1 = 0;
                    decimal FT = 0;
                    StringBuilder sb = new StringBuilder();

                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        range.Select();
                        WHNO = range.Text.ToString().Trim();
                      //  lINFEETextBox
                        if (WHNO.IndexOf("WH") != -1)
                        {
                            System.Data.DataTable FF1 = Getdata3(WHNO);
                            if (FF1.Rows.Count > 0)
                            {
                                FRU3 += Convert.ToDecimal(FF1.Rows[0][0]);
                            }


                        }

                
                        //月裝拆櫃費用




                    }
  

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                }

            }
                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
          
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
     

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();







        }
        private void WriteExcelProduct5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
          
            Microsoft.Office.Interop.Excel.Worksheet excelSheet0 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet0.Activate();
            int M1 = 0;
            int M2 = 0;
            int M3 = 0;

            int iRowCnt0 = excelSheet0.UsedRange.Cells.Rows.Count;
            int iColCnt0 = excelSheet0.UsedRange.Cells.Columns.Count;
      
                for (int jj = 1; jj <= 30; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet0.UsedRange.Cells[jj, 1]);
                    range.Select();
                    string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

                    int G1 = id.IndexOf("進貨/出貨費用");
                    if (G1 != -1)
                    {

                        M2 = jj;
                        M1 = 1 + 2;
                        break;
                    }


                    int G2 = id.IndexOf("加班費用");
                    if (G2 != -1)
                    {

                        M3 = jj;
                        break;
                    }
                }

                for (int jj = 1; jj <= 30; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet0.UsedRange.Cells[jj, 1]);
                    range.Select();
                    string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

       

                    int G2 = id.IndexOf("加班費用");
                    if (G2 != -1)
                    {

                        M3 = jj;
                        break;
                    }
                }


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet0.UsedRange.Cells[M2, M1]);
            range.Select();
            新得利進出費用 = range.Text.ToString().Trim().ToUpper().Replace(",", "");

            if (M3 != 0)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet0.UsedRange.Cells[M3, M1]);
                range.Select();
                新得利加班費用 = range.Text.ToString().Trim().ToUpper().Replace(",", "").Replace("-", "");
            }

            
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

       




            try
            {
                string DOCDATE;

                string TOTAL;
                string DUP = "";
                string PLATE;
                string YEAR;
                string MONTH;
                string DAY;
                string WHNO;
                string FEEB;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    TOTAL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    PLATE = range.Text.ToString().Trim();

                    if (DOCDATE.Length > 9)
                    {
                        YEAR = DOCDATE.Substring(0, 4);
                        MONTH = Convert.ToInt16(DOCDATE.Substring(5, 2)).ToString();
                        DAY = Convert.ToInt16(DOCDATE.Substring(8, 2)).ToString();
                        System.Data.DataTable G1 = Getdata6(YEAR, MONTH, DAY);
                        
                        if (G1.Rows.Count > 0)
                        {

                            string QTY = G1.Rows[0][0].ToString();


                            if (QTY != PLATE)
                            {
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助版數 : " + QTY;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }
                        }
                        
                        
                    }

                    int T1 = TOTAL.IndexOf("合計");
                    if (T1 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        range.Select();
                        TOTALA = range.Text.ToString().Trim().Replace(",", "");

                    
                    }
            
           
                 
                }
                //NEWNEW
                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(4);
                excelSheet2.Activate();

                int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
                int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;
                string TS;

                for (int iRecord = 2; iRecord <= iRowCnt2; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    TS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    FEEB = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(FEEB) && !String.IsNullOrEmpty(WHNO))
                    {
                        System.Data.DataTable GG1 = Getdata4(WHNO);
                        if (GG1.Rows.Count > 0)
                        {
                            DateTime D1 = Convert.ToDateTime(DOCDATE);
                            DateTime D2 = Convert.ToDateTime(GG1.Rows[0]["DOCDATE"]);
                            string FEES = GG1.Rows[0]["FEE"].ToString();
                

                            if (FEEB != FEES)
                            {
                    
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : " + FEES;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }
                            if (D1 != D2)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 1]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助日期: " + D2;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }

                            DUP = DOCDATE;

                        }
                    }
                    int T1 = TS.IndexOf("合計");
                    if (T1 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 10]);
                        range.Select();
                        TOTALB = range.Text.ToString().Trim().Replace(",", "");


                    }

                    int T2 = TS.IndexOf("0.9");
                    if (T2 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 10]);
                        range.Select();
                        TOTALB = range.Text.ToString().Trim().Replace(",", "");


                    }
                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(5);
                excelSheet3.Activate();

                int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
                int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;


                for (int iRecord = 2; iRecord <= iRowCnt3; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    TS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    FEEB = range.Text.ToString().Trim();
                    if (!String.IsNullOrEmpty(FEEB) && !String.IsNullOrEmpty(WHNO))
                    {
                        System.Data.DataTable GG1 = Getdata3(WHNO);
                        if (GG1.Rows.Count > 0)
                        {
                            string FEES = GG1.Rows[0]["FEE"].ToString();


                            if (FEEB != FEES)
                            {

                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : " + FEES;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }


                        }
                    }

                    int T1 = TS.IndexOf("合計");
                    if (T1 != -1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord, 5]);
                        range.Select();
                        TOTALC = range.Text.ToString().Trim().Replace(",", "");


                    }
                }

            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


                System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void WriteExcelProduct6(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
         
            try
            {
                 for (int s = 1; s <= 2; s++)
                {
       

    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(s);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


    
                string QTY;
                string AMT;
                string MEMO;
                string DOCTYPE;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                  
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    DOCTYPE = range.Text.ToString().Trim();

                    if (DOCTYPE =="進金生-3")
                    {
                        DOCTYPE = "聯揚";
                    }

                    if (DOCTYPE == "進金生")
                    {
                        DOCTYPE = "新得利";
                    }
                

                    
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    QTY = range.Text.ToString().Trim();



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    AMT = range.Text.ToString().Trim();

                    //if (AMT == "630")
                    //{
                    //    MessageBox.Show("saA");
                    //}
                    int n;
                    if (int.TryParse(QTY, out n) && int.TryParse(AMT, out n))
                    {

                        String[] split = MEMO.Split(' ');
                        int count = 0;
                        StringBuilder sb = new StringBuilder();
                        foreach (String F in split)
                        {
                            string GG = F.ToString();

                            int G2 = GG.ToUpper().IndexOf("WH");
                            if (G2 != -1)
                            {
                                string WH = GG.Substring(G2, 14);
                                sb.Append("'" + WH + "',");
                            }
                        }
                        if (sb.Length > 1)
                        {
                            sb.Remove(sb.Length - 1, 1);
                        }
                        int G1 = MEMO.ToUpper().IndexOf("WH");
                        if (G1 != -1)
                        {
                            string WH = MEMO.Substring(G1, 14);
                            if (WH == "WH20200921017X")
                            {
                                MessageBox.Show("S");
                            }
                            int GQ  = Convert.ToInt32(QTY);
                            int GA = Convert.ToInt32(AMT);
                            int GQ2 = 0;
                            int GA2 = 0;
                            int GQ3 = 0;
                            int GA3 = 0;
                            if (GQ > 10)
                            {
                                GQ2 = GQ - 10;
                                GA2 = 70 * GQ2;
                                GQ3 = 10;
                                GA3 = GA - GA2;
                                INSFEE3(DOCTYPE, Convert.ToInt32(GQ2), Convert.ToInt32(GA2), sb.ToString());
                                INSFEE3(DOCTYPE, Convert.ToInt32(GQ3), Convert.ToInt32(GA3), sb.ToString());
                            }
                            else
                            {
                                INSFEE3(DOCTYPE, Convert.ToInt32(QTY), Convert.ToInt32(AMT), sb.ToString());
                            }
          
                            //if (WH == "WH20200612027X")
                            //{
                            //    MessageBox.Show("a");
                            //}
                            System.Data.DataTable C1 = Getdata10(sb.ToString());
                            string AMT2 = "0";
                            if (C1.Rows.Count > 0)
                            {
                                AMT2 = C1.Rows[0][0].ToString();
                            }

                            if (AMT != AMT2)
                            {
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                range.ClearComments();
                                string MM = "輔助金額 : " + AMT2;
                                range.AddComment(MM);

                                int wCount = CountText(MM, '\n');
                                range.Comment.Shape.Height = wCount * 20;
                            }
                            
                        
                        }

                             //if (AMT != AMT2)
                            //{
                            //    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            //    range.ClearComments();
                            //    string MM = "輔助金額 : " + AMT2;
                            //    range.AddComment(MM);

                            //    int wCount = CountText(MM, '\n');
                            //    range.Comment.Shape.Height = wCount * 20;
                            //}
                    }
                }




                }
  


       
            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


                System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void INSFEE3(string DOCTYPE, int QTY, int AMT, string WHNO)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO  [dbo].[WH_FEE3]");
            sb.Append("            ([DOCTYPE],[QTY],[AMT],[WHNO])");
            sb.Append("      VALUES (@DOCTYPE,@QTY,@AMT,@WHNO)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
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
        private void DELFEE3()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("TRUNCATE TABLE  [dbo].[WH_FEE3]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);




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
        private System.Data.DataTable Getdata3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT CAST(ISNULL(lINFEE,0) AS INT) FEE,CAST(T0.QUANTITY AS DATETIME)  DOCDATE  FROM WH_MAIN T0     ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)    ");
            sb.Append(" WHERE  T0.ShippingCode=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetdataNAN1(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ISNULL(KT2,0) FEE FROM WH_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata4(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT  ISNULL(CAST(ISNULL(KT2,0) AS INT),0) FEE,CAST(T0.QUANTITY AS DATETIME)  DOCDATE  FROM WH_MAIN T0     ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)    ");
           // sb.Append(" WHERE T0.ShippingCode=@SHIPPINGCODE AND T1.ENA <> '1' ");
            sb.Append(" WHERE T0.ShippingCode=@SHIPPINGCODE  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        //LOUC
        private System.Data.DataTable Getdata4WH(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT C1 FROM WH_FEE WHERE ShippingCode=@SHIPPINGCODE AND C1='聯倉' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Getdata4WH2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT C1 FROM WH_FEE WHERE ShippingCode=@SHIPPINGCODE AND LOUC='聯倉' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata5(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            //            sb.Append(" AND  REPLACE( T0.QUANTITY,'/','') BETWEEN @DATE1 AND @DATE2");
            sb.Append(" SELECT (CAST(ISNULL(lOUFEE,0) AS INT))  FEE,CAST(T0.QUANTITY AS DATETIME)  DOCDATE  FROM WH_MAIN T0    ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)   ");
            sb.Append(" WHERE CAST(ISNULL(lOUFEE,0) AS INT)<>0 AND T0.ShippingCode=@SHIPPINGCODE");
        
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Getdata6(string DOCYEAR, string DOCMONTH, string DOCDATE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT TOTALQTY QTY  FROM WH_PLATE WHERE DOCYEAR=@DOCYEAR AND DOCMONTH =@DOCMONTH AND DOCDATE=@DOCDATE AND WHSCODE ='新得利倉' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCYEAR", DOCYEAR));
            command.Parameters.Add(new SqlParameter("@DOCMONTH", DOCMONTH));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Getdata7()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  'ZA0SF0005' 產品編號,'TFT-'+CAST(MAX(DOCYEAR)-1911 AS VARCHAR)+'年'+CAST(MAX(DOCMONTH) AS VARCHAR)+'月 倉租費-10元一板*'+CAST(CAST(SUM(TOTALQTY)*10 AS  INT) AS VARCHAR) 產品名稱,SUM(TOTALQTY)*10 費用,'Z0009' 倉庫");
            sb.Append("   FROM WH_PLATE WHERE  WHSCODE ='新得利倉' ");
            sb.Append("    AND  CAST(DOCYEAR AS VARCHAR) + CASE WHEN DOCMONTH <10 THEN '0'+CAST(DOCMONTH AS VARCHAR) ELSE CAST(DOCMONTH AS VARCHAR) END");
            sb.Append("  +CASE WHEN DOCDATE <10 THEN '0'+CAST(DOCDATE AS VARCHAR) ELSE CAST(DOCDATE AS VARCHAR) END");
            sb.Append("  BETWEEN @DATE1 AND @DATE2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata8()
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'ZA0SB0005' 產品編號,'TFT-'+ CAST(YEAR(CAST(MAX(T0.QUANTITY) AS DATETIME)) -1911 AS VARCHAR) +'年'+CAST(MONTH(CAST(MAX(T0.QUANTITY) AS DATETIME)) AS VARCHAR)+'月 出倉卡車費-TW LOCAL' 產品名稱,ISNULL(SUM(CAST(ISNULL(KT2,0) AS INT)),0) 費用,'Z0014' 倉庫  FROM WH_MAIN T0     ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)    ");
            sb.Append(" WHERE (C1='新得利' OR LOUC='新得利'  OR LINC ='新得利' )  ");
            sb.Append(" AND  REPLACE( T0.QUANTITY,'/','') BETWEEN @DATE1 AND @DATE2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@C1", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata9()
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'ZA0SZ0400' 產品編號,'TFT-'+ CAST(YEAR(CAST(MAX(T0.QUANTITY) AS DATETIME)) -1911 AS VARCHAR) +'年'+CAST(MONTH(CAST(MAX(T0.QUANTITY) AS DATETIME)) AS VARCHAR)+'月 卸裝櫃費' ");
            sb.Append(" +CASE WHEN ISNULL(SUM(CASE LIN20  WHEN 'CHECKED' THEN 1 END),0) =0 THEN '' ELSE CAST(SUM(CASE LIN20  WHEN 'CHECKED' THEN 1 END) AS VARCHAR) +'*20''*1500+' END");
            sb.Append(" +CASE WHEN ISNULL(SUM(CASE LIN40  WHEN 'CHECKED' THEN 1 END),0) =0 THEN '' ELSE CAST(SUM(CASE LIN40  WHEN 'CHECKED' THEN 1 END) AS VARCHAR) +'*40''*3000' END 產品名稱");
            sb.Append(" ,ISNULL(SUM(CAST(ISNULL(LINFEE,0) AS INT)),0) 費用,'Z0011' 倉庫  FROM WH_MAIN T0       ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)  ");
            sb.Append(" WHERE (C1='新得利' OR LOUC='新得利'  OR LINC ='新得利' )   ");
            sb.Append(" AND  REPLACE( T0.QUANTITY,'/','') BETWEEN @DATE1 AND @DATE2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata10(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(CAST(ISNULL(kT,0) AS INT))  kT  FROM WH_FEE WHERE SHIPPINGCODE  IN ( " + SHIPPINGCODE + ") AND k1='航通'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
          //  command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata11(string DOCTYPE)
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT SUM(QTY) QTY,CAST(CAST(AMT AS DECIMAL)/CAST(QTY AS DECIMAL) AS DECIMAL(12,4)) PRICE, SUM(AMT) AMT FROM WH_FEE3 ");
            sb.Append(" WHERE DOCTYPE =@DOCTYPE");
            sb.Append("              GROUP BY CAST(CAST(AMT AS DECIMAL)/CAST(QTY AS DECIMAL) AS DECIMAL(12,4)) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata12(string DOCTYPE)
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(QTY),0) QTY  FROM WH_FEE3 WHERE DOCTYPE =@DOCTYPE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata9B()
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'ZA0SZ0400' 產品編號,'TFT-'+ CAST(YEAR(CAST(MAX(T0.QUANTITY) AS DATETIME)) -1911 AS VARCHAR) +'年'+CAST(MONTH(CAST(MAX(T0.QUANTITY) AS DATETIME)) AS VARCHAR)+'月 進倉理貨費' 產品名稱");
            sb.Append(" ,ISNULL(SUM(CAST(ISNULL(LINFEE,0) AS INT)),0) 費用,'Z0011' 倉庫  FROM WH_MAIN T0      ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)     ");
            sb.Append(" WHERE (C1='聯倉' OR LOUC='聯倉'  OR LINC ='聯倉' )   ");
            sb.Append(" AND  REPLACE( T0.QUANTITY,'/','') BETWEEN @DATE1 AND @DATE2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Getdata9B2()
        {

            SqlConnection connection = globals.Connection;


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'ZA0SB0005' 產品編號,'TFT-'+ CAST(YEAR(CAST(MAX(T0.QUANTITY) AS DATETIME)) -1911 AS VARCHAR) +'年'+CAST(MONTH(CAST(MAX(T0.QUANTITY) AS DATETIME)) AS VARCHAR)+'月 出倉卡車費-TW LOCAL'  產品名稱");
            sb.Append(" ,ISNULL(SUM(CAST(ISNULL(KT2,0) AS INT)),0)-ISNULL(SUM(CAST(ISNULL(CBF,0) AS INT)),0) 費用,'Z0014' 倉庫  FROM WH_MAIN T0       ");
            sb.Append(" LEFT JOIN  WH_FEE T1 ON (T0.ShippingCode=T1.ShippingCode)      ");
            sb.Append(" WHERE (C1='聯倉'  )   ");
            sb.Append(" AND  REPLACE( T0.QUANTITY,'/','') BETWEEN @DATE1 AND @DATE2");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public int CountText(String text, Char w)
        {
            String[] split = text.Split(w);
            int count = 0;
            foreach (String s in split)
            {
                if (s.Length > 0) count++;
            }
            return count;
        }



        private void button4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G2 = dtCost;
            if (comboBox2.Text == "航通快遞")
            {
                G2 = dtCostF;
            }
            if (G2.Rows.Count > 0)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = FA;
                oCompany.UserName = "manager";
                oCompany.Password = "19571215";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                 
                    oPURCH.CardCode = dataGridView2.Rows[0].Cells["廠商編號"].Value.ToString();
                    oPURCH.DocCurrency = "NTD";
                    oPURCH.DocumentsOwner = 8;
                    oPURCH.Comments = textBox5.Text;

                    for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
                    {

                        DataGridViewRow row;

                        row = dataGridView2.Rows[i2];
 
                        string 產品編號 = row.Cells["產品編號"].Value.ToString();
                        string 費用名稱 = row.Cells["費用名稱"].Value.ToString();
   
                        string SAP金額 = row.Cells["SAP金額"].Value.ToString();
                        string 倉庫 = row.Cells["倉庫"].Value.ToString();

                        oPURCH.Lines.WarehouseCode = 倉庫;
                        oPURCH.Lines.ItemCode = 產品編號;
                        oPURCH.Lines.ItemDescription = 費用名稱;
                        oPURCH.Lines.LineTotal = Convert.ToDouble(SAP金額);
                        if (comboBox2.Text == "航通快遞")
                        {
                            oPURCH.Lines.Price = Convert.ToDouble(row.Cells["單價"].Value);
                            oPURCH.Lines.Quantity = Convert.ToDouble(row.Cells["數量"].Value);
                        }
                        oPURCH.Lines.VatGroup = "AP5%";
                        oPURCH.Lines.Currency = "NTD";
                        oPURCH.Lines.CostingCode = "11111";
                        oPURCH.Lines.Add();
                    }


                   


                    int res = oPURCH.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
                        MessageBox.Show("上傳成功 採購單號 : " + OWTR);

                    
                    }




                }
            }
        }

        public System.Data.DataTable GetDI4()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OPOR");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G2 = dtCost;
            if (comboBox2.Text == "航通快遞")
            {
                G2 = dtCostF;
            }
            if (G2.Rows.Count > 0)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = FA;
                oCompany.UserName = "manager";
                oCompany.Password = "0918";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oPURCH.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    oPURCH.CardCode = dataGridView2.Rows[0].Cells["廠商編號"].Value.ToString();
                    oPURCH.DocCurrency = "NTD";
                    oPURCH.DocumentsOwner = 8;
                    oPURCH.Comments = textBox5.Text;

                    for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
                    {

                        DataGridViewRow row;

                        row = dataGridView2.Rows[i2];

                        string 產品編號 = row.Cells["產品編號"].Value.ToString();
                        string 費用名稱 = row.Cells["費用名稱"].Value.ToString();

                        string SAP金額 = row.Cells["SAP金額"].Value.ToString();
                        string 倉庫 = row.Cells["倉庫"].Value.ToString();

                        oPURCH.Lines.WarehouseCode = 倉庫;
                        oPURCH.Lines.ItemCode = 產品編號;
                        oPURCH.Lines.ItemDescription = 費用名稱;
                        oPURCH.Lines.LineTotal = Convert.ToDouble(SAP金額);
                        if (comboBox2.Text == "航通快遞")
                        {
                            oPURCH.Lines.Price = Convert.ToDouble(row.Cells["單價"].Value);
                            oPURCH.Lines.Quantity = Convert.ToDouble(row.Cells["數量"].Value);
                        }
                        oPURCH.Lines.VatGroup = "AP5%";
                        oPURCH.Lines.Currency = "NTD";
                        oPURCH.Lines.CostingCode = "11111";
                        oPURCH.Lines.Add();
                    }





                    int res = oPURCH.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
                        MessageBox.Show("上傳成功 採購單號 : " + OWTR);


                    }




                }
            }
        }
    }
}

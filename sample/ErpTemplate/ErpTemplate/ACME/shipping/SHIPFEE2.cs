using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
using SAPbobsCOM;
namespace ACME
{
    public partial class SHIPFEE2 : Form
    {
        System.Data.DataTable dtCost = null;
        DataRow dr = null;
        string FA = "acmesql02";
        public SHIPFEE2()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable T1 = GetSHIP();
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }
        }
        private System.Data.DataTable GetSHIP()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.SHIPPINGCODE FROM SHIP_FEE T0");
            sb.Append(" LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T1.CloseDay  BETWEEN @AA AND @BB ");
            sb.Append(" AND (T0.CARDNAME LIKE '%" + textBox3.Text + "%' OR T0.CARDNAME2 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" OR T0.CARDNAME3 LIKE '%" + textBox3.Text + "%' OR T0.CARDNAME4 LIKE '%" + textBox3.Text + "%'");

            sb.Append(" OR T0.CARDNAME5 LIKE '%" + textBox3.Text + "%' OR T0.CARDNAME6 LIKE '%" + textBox3.Text + "%')");
            sb.Append("     UNION ALL  ");
            sb.Append("      SELECT T0.SHIPPINGCODE FROM SHIP_FEE T0 ");
            sb.Append("              LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T1.CloseDay  BETWEEN @AA AND @BB ");
            sb.Append("			  AND ISNULL(bT,0) <> 0 AND '財政部台北關務局'  LIKE  '%" + textBox3.Text + "%' ");
            sb.Append("     UNION ALL  ");
            sb.Append("      SELECT T0.SHIPPINGCODE FROM SHIP_FEE T0 ");
            sb.Append("              LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T1.CloseDay  BETWEEN @AA AND @BB ");
            sb.Append("			  AND ISNULL(bT,0) <> 0 AND '財政部基隆關務局'  LIKE  '%" + textBox3.Text + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetOWNER(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SLPCODE SALES,OwnerCode  SA FROM AcmeSql02.DBO.ORDR WHERE DOCENTRY IN (" );
            sb.Append(" SELECT TOP 1 DOCENTRY  FROM Shipping_Item WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCENTRY <>'')");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        private System.Data.DataTable GetSHIP2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SB0005' 產品編號,'卡車費' 費用名稱,CAST(ISNULL(aAMT,0) AS INT) 金額,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(aAMT,0) AS INT)<>0 AND CARDNAME2 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SA0004' ITEMCODE,'報關費' 費用名稱,CAST(ISNULL(bAF,0) AS INT)+CAST(ISNULL(bSF,0) AS INT) 報關費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(bAF,0) AS INT)+CAST(ISNULL(bSF,0) AS INT) <>0 AND (CARDNAME LIKE '%" + textBox3.Text + "%' OR CARDNAME4 LIKE '%" + textBox3.Text + "%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SZ0701' ITEMCODE,'加班費' 費用名稱,CAST(ISNULL(bAE,0) AS INT) 加班費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(bAE,0) AS INT) <>0 AND (CARDNAME LIKE '%" + textBox3.Text + "%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SZ1100' ITEMCODE,'逾時費' 費用名稱,CAST(ISNULL(aTime,0) AS INT) 逾時費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(aTime,0) AS INT) <>0 AND (CARDNAME2 LIKE '%" + textBox3.Text + "%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE4,CARDNAME4,'ZA0SZ0601' ITEMCODE,'機械使用費' 費用名稱,CAST(ISNULL(cSG,0) AS INT) 機械使用費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSG,0) AS INT) <> 0 AND CARDNAME4 LIKE '%" + textBox3.Text + "%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT 'U0133' CARDCODE,'財政部台北關務局' CARDNAME,'ZA0SD0100' ITEMCODE,'推貿費' 費用名稱,CAST(ISNULL(BT,0) AS INT) 推貿費,'Z0002' 倉庫  FROM SHIP_FEE WHERE CAST  (ISNULL(BT,0) AS INT)  <> 0 AND SHIPPINGCODE=@SHIPPINGCODE AND '財政部台北關務局'  LIKE '%" + textBox3.Text + "%'   ");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT 'U0134' CARDCODE,'財政部基隆關務局' CARDNAME,'ZA0SD0100' ITEMCODE,'推貿費' 費用名稱,CAST(ISNULL(CT,0) AS INT) 推貿費,'Z0002' 倉庫  FROM SHIP_FEE WHERE CAST  (ISNULL(CT,0) AS INT)  <> 0 AND SHIPPINGCODE=@SHIPPINGCODE AND '財政部基隆關務局'  LIKE '%" + textBox3.Text + "%' ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SE0001' ITEMCODE,'提單文件費' 費用名稱,CAST(ISNULL(cAT,0) AS INT)+CAST(ISNULL(cST,0) AS INT) 提單文件費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cAT,0) AS INT)+CAST(ISNULL(cST,0) AS INT)   <> 0 AND (CARDNAME5 LIKE '%" + textBox3.Text + "%' OR CARDNAME6 LIKE '%" + textBox3.Text + "%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE5,CARDNAME5,'ZA0SF0005' ITEMCODE,'倉租費' 費用名稱,CAST(ISNULL(cAZ,0) AS INT) 倉租費,'Z0009' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cAZ,0) AS INT) <> 0 AND CARDNAME5 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE5,CARDNAME5,'ZA0SF0005' ITEMCODE,'X光過檢費' 費用名稱,CAST(ISNULL(cAX,0) AS INT) 倉租費,'Z0009' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cAX,0) AS INT) <> 0 AND CARDNAME5 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE,CARDNAME,'ZA0SZ1500' ITEMCODE,'艙單申報' 費用名稱,CAST(ISNULL(cAS,0) AS INT)+CAST(ISNULL(cSS,0) AS INT) 艙單申報,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cAS,0) AS INT)+CAST(ISNULL(cSS,0) AS INT)   <> 0 AND (CARDNAME5 LIKE '%" + textBox3.Text + "%' OR CARDNAME6 LIKE '%" + textBox3.Text + "%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SC0301' ITEMCODE,'併櫃費' 費用名稱,CAST(ISNULL(cSB,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSB,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SZ1100' ITEMCODE,'VGM (其他)' 費用名稱,CAST(ISNULL(cSV,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSV,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SE0201' ITEMCODE,'提單電放費' 費用名稱,CAST(ISNULL(cSD,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSD,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SZ0001' ITEMCODE,'操作手續費' 費用名稱,CAST(ISNULL(cSH,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSH,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SZ1403' ITEMCODE,'低硫附加費' 費用名稱,CAST(ISNULL(cSLIU,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSLIU,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE3,CARDNAME3,'ZA0TE0104' ITEMCODE,'出口運費' 費用名稱,CAST(ISNULL(dHL2,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(dHL2,0) AS INT) <> 0 AND CARDNAME3 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SZ1201 ' ITEMCODE,'封條費' 費用名稱,CAST(ISNULL(cSS2,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSS2,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CARDCODE6,CARDNAME6,'ZA0SC0204' ITEMCODE,'吊櫃費' 費用名稱,CAST(ISNULL(cSS3,0) AS INT) 倉租費,'Z0002' 倉庫  FROM SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE AND CAST(ISNULL(cSS3,0) AS INT) <> 0 AND CARDNAME6 LIKE '%" + textBox3.Text + "%'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private void SHIPFEE2_Load(object sender, EventArgs e)
        {

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.Day();
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("費用名稱", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            return dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
       
            StringBuilder sb = new StringBuilder();
            try
            {
               dtCost = MakeTable();
                ArrayList al = new ArrayList();

                for (int s = 0; s <= listBox2.Items.Count - 1; s++)
                {
                    string SHIPNO = listBox2.Items[s].ToString();
                    System.Data.DataTable dt = GetSHIP2(SHIPNO);
                    if (dt.Rows.Count > 0)
                    {
                        
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {
                            dr = dtCost.NewRow();
                            dr["廠商編號"] = textBox4.Text;
                            dr["廠商名稱"] = textBox3.Text;
                            dr["工單號碼"] = SHIPNO;
                            dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                            dr["費用名稱"] = dt.Rows[i]["費用名稱"].ToString();
                            dr["金額"] = dt.Rows[i]["金額"].ToString();
                            dr["倉庫"] = dt.Rows[i]["倉庫"].ToString();


                     
                            dtCost.Rows.Add(dr);
                        }

                    }
                }

                dataGridView1.DataSource = dtCost;



            }
            catch { }
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
        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G2 = dtCost;

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
                oCompany.UserName = "A01";
                oCompany.Password = "89206602";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);


                    oPURCH.CardCode = G2.Rows[0]["廠商編號"].ToString();
                    oPURCH.DocCurrency = "NTD";
                    oPURCH.Comments = textBox5.Text;
                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {

                            string ITEMCODE = G2.Rows[n]["產品編號"].ToString();
                            string 倉庫 = G2.Rows[n]["倉庫"].ToString();
                                               string 工單號碼 = G2.Rows[n]["工單號碼"].ToString();
                                               System.Data.DataTable G3 = GetOWNER(工單號碼);
                                               if (G3.Rows.Count > 0)
                                               {
                                                   oPURCH.SalesPersonCode = Convert.ToInt16(G3.Rows[0]["SALES"]);
                                                   oPURCH.DocumentsOwner = Convert.ToInt16(G3.Rows[0]["SA"]);
                                               }
                            oPURCH.Lines.WarehouseCode = 倉庫;
                                    oPURCH.Lines.ItemCode = ITEMCODE;
                                    oPURCH.Lines.LineTotal = Convert.ToDouble(G2.Rows[n]["金額"]);
                     
                                    oPURCH.Lines.VatGroup = "AP5%";
                                    oPURCH.Lines.Currency = "NTD";
                                    oPURCH.Lines.CostingCode = "11111";
                                  //  oPURCH.Lines.UserFields.Fields.Item("U_ACME_WhsName").Value = "出口";
                                    oPURCH.Lines.UserFields.Fields.Item("U_Shipping_no").Value = 工單號碼;
                 
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

                                    //System.Data.DataTable GG3 = GetOPOR5(T1, fmLogin.LoginID.ToString());

                                    //if (GG3.Rows.Count > 0)
                                    //{
                                    //    for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                    //    {

                                    //        string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                    //        string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                    //        //UPDATEOPOR(OWTR, LINENUM, REMARK);

                                    //        //UPDATEAPOPOR(T1, LINENUM, OWTR);
                                    //    }

                                    //}
                                }
                            
                        

                    
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                textBox4.Text = Convert.ToString(LookupValues[0]);
                textBox3.Text = Convert.ToString(LookupValues[1]);
            }
        }
    }
    }


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
//SELECT * FROM JDT1 WHERE TransId =425958 AND Account ='22610102'
//UPDATE JDT1 SET U_remark1 ='已沖帳' WHERE TransId =425958 AND Account ='22610102'

//SELECT * FROM SATT5 WHERE TRANSID=104227
//UPDATE SATT5 SET CHECKED ='Y',CHECKDATE ='20191226' WHERE TRANSID=124751
namespace ACME
{
    public partial class TTACC : Form
    {
        System.Data.DataTable dtCost = null;
        public TTACC()
        {
            InitializeComponent();
        }
        private void ACC()
        {
            System.Data.DataTable dt = Get1();
            System.Data.DataTable dt2 = Get2();

            dtCost = MakeTableCombine();

            DataRow dr = null;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                dr["群組"] = dt.Rows[i]["BU"].ToString();
                string CARDCODE = dt.Rows[i]["CARDCODE"].ToString();
                dr["客戶代碼"] = CARDCODE;
                dr["客戶名稱"] = dt.Rows[i]["CARDNAME"].ToString();

                System.Data.DataTable G7 = Get7(CARDCODE);
                if (G7.Rows.Count > 0)
                {
                    dr["客戶名稱"] = G7.Rows[0][0].ToString();
                }
                string TRANSID = dt.Rows[i]["TRANSID"].ToString();
                dr["傳票號碼"] = TRANSID;
                dr["傳票日期"] = dt.Rows[i]["TRANSDATE"].ToString();
                string ORDR = dt.Rows[i]["ORDR"].ToString();
                dr["銷售單號"] = ORDR;
                System.Data.DataTable SA1 = GetSA1(ORDR);
                if (SA1.Rows.Count > 0)
                {
                    dr["訂單交期"] = SA1.Rows[0][0].ToString();
                }
                dr["AR單號"] = dt.Rows[i]["OINV"].ToString();
                string MEMO = dt.Rows[i]["JRNLMEMO"].ToString();
                dr["摘要"] = MEMO;
              
                dr["NTD"] = dt.Rows[i]["AMT"].ToString();
                dr["匯率"] = dt.Rows[i]["CURRENCY"].ToString();
                dr["USD"] = dt.Rows[i]["USD"].ToString();
                dr["離倉日期"] = dt.Rows[i]["SHIPDATE"].ToString();
                System.Data.DataTable SA = GetSA(ORDR);
                if (SA.Rows.Count > 0)
                {
                    dr["業管"] = SA.Rows[0]["SA"].ToString();
                    dr["業務"] = SA.Rows[0]["SALES"].ToString();

                    System.Data.DataTable G2 = Get5(ORDR);
                    if (G2.Rows.Count > 0)
                    {
                        dr["離倉日期"] = G2.Rows[0][0].ToString();
                    }
                }
                else
                {
                    dr["業管"] = dt.Rows[i]["SA"].ToString();
                    dr["業務"] = dt.Rows[i]["SALES"].ToString();
                }
           
                dr["備註"] = dt.Rows[i]["MEMO"].ToString();
                dr["STYPE"] = "";
                if (MEMO.Substring(0, 1) == "#")
                {

                    dr["LORDER"] = MEMO.Replace("#", "");
                }
                else
                {
                    dr["LORDER"] = TRANSID;
                }
                dtCost.Rows.Add(dr);
            }

            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                dr["群組"] = dt2.Rows[i]["BU"].ToString();
                string CARDCODE = dt2.Rows[i]["CARDCODE"].ToString();
                System.Data.DataTable G6 = Get6(CARDCODE);

                if (G6.Rows.Count > 0)
                {
                    dr["群組"] = G6.Rows[0]["BU"].ToString();
                }
                dr["客戶代碼"] = CARDCODE;
                dr["客戶名稱"] = dt2.Rows[i]["CARDNAME"].ToString();
                string TRANSID = dt2.Rows[i]["TRANSID"].ToString();
                dr["傳票號碼"] = TRANSID;
                dr["傳票日期"] = dt2.Rows[i]["TRANSDATE"].ToString();
                string U_SATT = dt2.Rows[i]["U_SATT"].ToString();
                //if (TRANSID == "404526")
                //{
                //    MessageBox.Show("a");
                //}
                System.Data.DataTable G1 = Get4(U_SATT);

                if (G1.Rows.Count > 0)
                {
                    string ORDR = G1.Rows[0]["ORDR"].ToString();
                    dr["銷售單號"] = ORDR;
                    dr["AR單號"] = G1.Rows[0]["OINV"].ToString();
                    dr["業管"] = G1.Rows[0]["SA"].ToString();
                    dr["業務"] = G1.Rows[0]["SALES"].ToString();
                    System.Data.DataTable SA1 = GetSA1(ORDR);
                    if (SA1.Rows.Count > 0)
                    {
                        dr["訂單交期"] = SA1.Rows[0][0].ToString();
                    }
                    System.Data.DataTable G2 = Get5(ORDR);
                    if (G2.Rows.Count > 0)
                    {
                        dr["離倉日期"] = G2.Rows[0][0].ToString();
                    }
                }
                string LINE = dt2.Rows[i]["LineMemo"].ToString();
                dr["摘要"] = LINE;
                decimal NTD = Convert.ToDecimal(dt2.Rows[i]["AMT"]);
                dr["NTD"] = dt2.Rows[i]["AMT"].ToString();
                dr["匯率"] = dt2.Rows[i]["RATE"].ToString();
                string USD = "";
                int D1 = LINE.ToUpper().LastIndexOf("US");
                int D2 = LINE.ToUpper().LastIndexOf("*");
                if (D1 != -1 && D2 != -1)
                {
                    USD = LINE.Substring(D1 + 2, D2-D1-2);
                }
                if (String.IsNullOrEmpty(USD))
                {
                    USD = "0";
                }
                if (NTD < 0)
                {
                    USD = "-" + USD; 
                }
                dr["USD"] = USD;
     

                dr["備註"] = "";
                dr["STYPE"] = dt2.Rows[i]["STYPE"].ToString();
                if (LINE.Substring(0, 1) == "#")
                {

                    dr["LORDER"] = LINE.Replace("#", "");
                }
                else
                {
                    dr["LORDER"] = TRANSID;
                }
                dtCost.Rows.Add(dr);
            }
            System.Data.DataTable dt21 = Get21();

            for (int i = 0; i <= dt21.Rows.Count - 1; i++)
            {
                string TRANSID = dt21.Rows[i][0].ToString();
                System.Data.DataTable dt212 = Get212(TRANSID);
                if (dt212.Rows.Count > 0)
                {
                    for (int i2 = 0; i2 <= dt212.Rows.Count - 1; i2++)
                    {
                        string TRANSID2 = dt212.Rows[i2][0].ToString();
                        string Line_ID = dt212.Rows[i2][1].ToString();
                        System.Data.DataTable dt213 = Get213(TRANSID2, Line_ID);

                                      dr = dtCost.NewRow();

                                      dr["群組"] = dt213.Rows[0]["BU"].ToString();
                                      string CARDCODE = dt213.Rows[0]["CARDCODE"].ToString();
                                      System.Data.DataTable G6 = Get6(CARDCODE);

                                      if (G6.Rows.Count > 0)
                                      {
                                          dr["群組"] = G6.Rows[0]["BU"].ToString();
                                      }
                                      dr["客戶代碼"] = CARDCODE;
                                      dr["客戶名稱"] = dt213.Rows[0]["CARDNAME"].ToString();
                                      string TRANSID1 = dt213.Rows[0]["TRANSID"].ToString();
                                      dr["傳票號碼"] = TRANSID1;
                                      dr["傳票日期"] = dt213.Rows[0]["TRANSDATE"].ToString();
                                      string U_SATT = dt213.Rows[0]["U_SATT"].ToString();

                  
                                      System.Data.DataTable G1 = Get4(U_SATT);

                                      if (G1.Rows.Count > 0)
                                      {
                                          string ORDR = G1.Rows[0]["ORDR"].ToString();
                                          dr["銷售單號"] = ORDR;
                                          System.Data.DataTable SA1 = GetSA1(ORDR);
                                          if (SA1.Rows.Count > 0)
                                          {
                                              dr["訂單交期"] = SA1.Rows[0][0].ToString();
                                          }
                                          dr["AR單號"] = G1.Rows[0]["OINV"].ToString();
                                          dr["業管"] = G1.Rows[0]["SA"].ToString();
                                          dr["業務"] = G1.Rows[0]["SALES"].ToString();

                                          System.Data.DataTable G2 = Get5(ORDR);
                                          if (G2.Rows.Count > 0)
                                          {
                                              dr["離倉日期"] = G2.Rows[0][0].ToString();
                                          }
                                      }
                                      string LINE = dt213.Rows[0]["LineMemo"].ToString();
                                      dr["摘要"] = LINE;
                                      dr["NTD"] = dt213.Rows[0]["AMT"].ToString();
                                      dr["匯率"] = dt213.Rows[0]["RATE"].ToString();
                                      decimal NTD = Convert.ToDecimal(dt213.Rows[0]["AMT"]);
                                      string USD = "";
                                      int D1 = LINE.ToUpper().LastIndexOf("US");
                                      int D2 = LINE.ToUpper().LastIndexOf("*");
                                      if (D1 != -1 && D2 != -1)
                                      {
                                          USD = LINE.Substring(D1 + 2, D2 - D1 - 2);
                                      }
                                      if (String.IsNullOrEmpty(USD))
                                      {
                                          USD = "0";
                                      }
                                      if (NTD < 0)
                                      {
                                          USD = "-" + USD;
                                      }
                                      dr["USD"] = USD;


                                      dr["備註"] = "";
                                      dr["STYPE"] = dt213.Rows[0]["STYPE"].ToString();
                                      if (LINE.Substring(0, 1) == "#")
                                      {
                                          dr["LORDER"] = LINE.Replace("#", "");
                                      }
                                      else
                                      {
                                          dr["LORDER"] = TRANSID1;
                                      }
                                      dtCost.Rows.Add(dr);
                            
                    }
                
                }

            }
            dtCost.DefaultView.Sort = "客戶代碼,LORDER";
            dataGridView1.DataSource = dtCost;
            if (comboBox1.Text != "Please-Select")
            {
                dtCost.DefaultView.RowFilter = "業務= '" + comboBox1.Text + "' ";
            }
            if (comboBox2.Text != "Please-Select")
            {
                dtCost.DefaultView.RowFilter = "業管= '" + comboBox2.Text + "' ";
            }

            if (artextBox12.Text != "")
            {
                dtCost.DefaultView.RowFilter = "客戶名稱 like '%" + artextBox12.Text.ToString() + "%'";

            }
            string g = dtCost.Compute("Sum(NTD)", null).ToString();
            string g2 = Get3().Rows[0][0].ToString();
            decimal sh = Convert.ToDecimal(g);
            decimal sh2 = Convert.ToDecimal(g2);
            label3.Text = "台幣合計:" + sh.ToString("#,##0");
            label1.Text = "SAP預收款科目金額:" + sh2.ToString("#,##0");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ACC();
          
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("傳票日期", typeof(string));
            dt.Columns.Add("傳票號碼", typeof(string));
            dt.Columns.Add("銷售單號", typeof(string));
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("NTD", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("USD", typeof(Decimal));
            dt.Columns.Add("業管", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("離倉日期", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("訂單交期", typeof(string));
            dt.Columns.Add("STYPE", typeof(string));
            dt.Columns.Add("LORDER", typeof(string));
            
            return dt;
        }
        private System.Data.DataTable Get1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT *  FROM  SATT5 WHERE ID NOT IN (SELECT ID  FROM  SATT5 WHERE ISNULL(CHECKED,'') = 'Y' AND ISNULL(CHECKDATE,'') BETWEEN '20171231' AND  @DocDate2)　");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get2()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'TFT' BU,U_CUSTCODE CARDCODE,T2.CARDNAME,T0.TRANSID,T1.U_SATT,T0.U_OTRANSID,Convert(varchar(8),T0.TAXDATE,112) TRANSDATE ");
            sb.Append(" ,T0.LineMemo,CAST((T0.Credit-T0.Debit) AS INT)  AMT,");
            sb.Append(" CASE WHEN Credit <> 0 THEN '立帳' ELSE '沖帳' END STYPE,");
            sb.Append(" REPLACE(ltrim(substring(LineMemo,CHARINDEX('US', LineMemo)+2,CASE CHARINDEX('*', LineMemo) WHEN 0 THEN 0 ELSE CHARINDEX('*', LineMemo)-2-CHARINDEX('US', LineMemo) END )),',','') USD,");
            sb.Append(" CASE CHARINDEX('*', LineMemo) WHEN 0 THEN '' ELSE ltrim(substring(LineMemo,CHARINDEX('*', LineMemo)+1,10 ))  END RATE");
            sb.Append(" FROM JDT1  T0  ");
            sb.Append(" LEFT JOIN OJDT T1 ON  (T0.TRANSID=T1.TRANSID)  ");
            sb.Append(" LEFT JOIN OCRD T2 ON  (T0.U_CUSTCODE=T2.CARDCODE)  ");
            sb.Append(" LEFT JOIN (SELECT SUM(Credit-Debit) AMT,U_OTRANSID FROM JDT1      WHERE ACCOUNT=22610102 AND TAXDATE >'2018-05-22 00:00:00.000'   ");
            sb.Append(" AND Convert(varchar(8),TAXDATE,112)  between '20071231' and @DocDate2  ");
            sb.Append(" AND ISNULL(U_OTRANSID,'') <> '' ");
            sb.Append(" GROUP BY U_OTRANSID) T3 ON  (T0.TRANSID=T3.U_OTRANSID)  ");
            sb.Append(" WHERE ACCOUNT=22610102 AND T0.TAXDATE >'2018-05-22 00:00:00.000'  AND (T0.Credit-T0.Debit)+ISNULL(T3.AMT,0) <>0 ");
            sb.Append(" AND LineMemo NOT LIKE  ('%迴轉%') AND ISNULL(T0.U_OTRANSID,'') = ''  AND ISNULL(T0.U_remark1,'') NOT IN ('已立帳','已沖帳')  ");
            sb.Append(" AND Convert(varchar(8),T0.TAXDATE,112)  between '20071231' and @DocDate2 and t0.TransId <>404526 ");   
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT 'TFT' BU,U_CUSTCODE CARDCODE,T2.CARDNAME,T0.TRANSID,T1.U_SATT,T0.U_OTRANSID,Convert(varchar(8),T0.TAXDATE,112) TRANSDATE  ");
            sb.Append("               ,T0.LineMemo,CAST((T0.Credit-T0.Debit) AS INT),CASE WHEN Credit <> 0 THEN '立帳' ELSE '沖帳' END STYPE  ");
            sb.Append("               ,REPLACE(ltrim(substring(LineMemo,CHARINDEX('US', LineMemo)+2,CASE CHARINDEX('*', LineMemo) WHEN 0 THEN 0 ELSE CHARINDEX('*', LineMemo)-2-CHARINDEX('US', LineMemo) END )),',','') USD, ");
            sb.Append("               CASE CHARINDEX('*', LineMemo) WHEN 0 THEN '' ELSE ltrim(substring(LineMemo,CHARINDEX('*', LineMemo)+1,10 ))  END RATE ");
            sb.Append("               FROM JDT1  T0  ");
            sb.Append("               LEFT JOIN OJDT T1 ON  (T0.TRANSID=T1.TRANSID)   ");
            sb.Append("               LEFT JOIN OCRD T2 ON  (T0.U_CUSTCODE=T2.CARDCODE)   ");
            sb.Append("               WHERE ACCOUNT=22610102 AND T0.TAXDATE >'2018-05-22 00:00:00.000'    ");
            sb.Append("               AND ISNULL(U_OTRANSID,'') <> '' AND U_OTRANSID <363261  ");
            sb.Append("               AND Convert(varchar(8),T0.TAXDATE,112)  between '20071231' and @DocDate2 ");
            sb.Append(" 						  AND T0.TransId+' '+T0.Line_ID  NOT IN (	SELECT T0.TransId+' '+T0.Line_ID     FROM JDT1 T0                 ");
            sb.Append(" 			WHERE   SUBSTRING(ISNULL(U_remark1,''),1,3)  IN ('已立帳','已沖帳')           ");
            sb.Append(" 						    AND CASE SUBSTRING(ISNULL(U_remark1,''),4,8) WHEN '' THEN  '20180501'  ELSE  SUBSTRING(ISNULL(U_remark1,''),4,8) END  between '20071231' and @DocDate2 )");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            command.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Get21()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         select  DISTINCT T0.TRANSID FROM JDT1  T0  ");
            sb.Append(" LEFT JOIN (SELECT SUM(Credit-Debit) AMT,U_OTRANSID FROM JDT1      WHERE ACCOUNT=22610102 AND TAXDATE >'2018-05-22 00:00:00.000'   ");
            sb.Append(" AND Convert(varchar(8),TAXDATE,112)  between '20071231' and @DocDate2  ");
            sb.Append(" AND ISNULL(U_OTRANSID,'') <> '' ");
            sb.Append(" GROUP BY U_OTRANSID) T3 ON  (T0.TRANSID=T3.U_OTRANSID)  ");
            sb.Append(" WHERE ACCOUNT=22610102 AND T0.TAXDATE >'2018-05-22 00:00:00.000'  AND (T0.Credit-T0.Debit)+ISNULL(T3.AMT,0) <>0 ");
            sb.Append(" AND LineMemo NOT LIKE  ('%迴轉%') AND ISNULL(T0.U_OTRANSID,'') = ''  AND ISNULL(T0.U_remark1,'') NOT IN ('已立帳','已沖帳')  ");
            sb.Append(" AND Convert(varchar(8),T0.TAXDATE,112)  between '20071231' and @DocDate2  ");
            sb.Append(" AND  ISNULL(T3.AMT,0) <>0 AND T0.TRANSID <> 404526 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get212(string U_OTRANSID)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("       SELECT TRANSID,Line_ID LINE FROM JDT1 WHERE U_OTRANSID=@U_OTRANSID  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_OTRANSID", U_OTRANSID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get213(string TransId, string Line_ID)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" 			                select 'TFT' BU,U_CUSTCODE CARDCODE,T2.CARDNAME,T0.TRANSID,T1.U_SATT,T0.U_OTRANSID,Convert(varchar(8),T0.TAXDATE,112) TRANSDATE  ");
            sb.Append("               ,T0.LineMemo,CAST((T0.Credit-T0.Debit) AS INT)  AMT, ");
            sb.Append("               CASE WHEN Credit <> 0 THEN '立帳' ELSE '沖帳' END STYPE, ");
            sb.Append("               REPLACE(ltrim(substring(LineMemo,CHARINDEX('US', LineMemo)+2,CASE CHARINDEX('*', LineMemo) WHEN 0 THEN 0 ELSE CHARINDEX('*', LineMemo)-2-CHARINDEX('US', LineMemo) END )),',','') USD, ");
            sb.Append("               CASE CHARINDEX('*', LineMemo) WHEN 0 THEN '' ELSE ltrim(substring(LineMemo,CHARINDEX('*', LineMemo)+1,10 ))  END RATE ");
            sb.Append("               FROM JDT1  T0   ");
            sb.Append("               LEFT JOIN OJDT T1 ON  (T0.TRANSID=T1.TRANSID)   ");
            sb.Append("               LEFT JOIN OCRD T2 ON  (T0.U_CUSTCODE=T2.CARDCODE)   ");
            sb.Append("               WHERE ACCOUNT=22610102 AND  T0.TransId =@TransId  AND  T0.Line_ID =@Line_ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TransId", TransId));
            command.Parameters.Add(new SqlParameter("@Line_ID", Line_ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT SUM(Credit-Debit) DEB  FROM JDT1 WHERE ACCOUNT ='22610102' AND Convert(varchar(19),TaxDate,112)  <= @TaxDate ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TaxDate", textBox2.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get4(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.Docentry ORDR,T1.MEMO OINV,T5.lastName +T5.firstName SA,(T6.[SlpName]) SALES FROM SATT1 T0");
            sb.Append(" LEFT JOIN SATT2 T1 ON (T0.TTCode =T1.TTCode AND T0.Seqno =T1.ID)");
            sb.Append("  left join ACMESQL02.DBO.ORDR T4 on (T1.DOCENTRY=T4.DOCENTRY )");
            sb.Append(" left join ACMESQL02.DBO.OHEM T5 on (T4.OwnerCode =T5.EMPID )");
            sb.Append("  left JOIN ACMESQL02.DBO.OSLP T6 ON T4.SlpCode = T6.SlpCode");

                sb.Append("  WHERE T0.ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get5(string DOCENTRY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT Convert(varchar(19),U_ACME_SHIPDAY,111)  SHIPDAY  FROM RDR1 WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Get6(string CARDCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUBSTRING(GROUPNAME,4,6) BU  FROM ocrd T9 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OCRG T10 ON (T9.GROUPCODE=T10.GROUPCODE)");
            sb.Append(" WHERE T9.CARDCODE=@CARDCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Get7(string CARDCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT　CARDNAME  FROM OCRD WHERE CARDCODE=@CARDCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GetSA(string PINO)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (T2.[SlpName]) SALES,(T3.[lastName]+T3.[firstName]) SA");
            sb.Append(" FROM ORDR T0 ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode");
            sb.Append(" INNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" WHERE    CAST(T0.DOCENTRY AS VARCHAR)=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", PINO));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetSA1(string DOCENTRY)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 Convert(varchar(10),SHIPDATE,111) DDATE FROM RDR1 WHERE CAST(DOCENTRY AS VARCHAR)=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        private void TTACC_Load(object sender, EventArgs e)
        {
            textBox2.Text = GetMenu.Day();
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetOslp1(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetOhem(), "DataValue", "DataValue");
          }

        private void button3_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\預收款.xlsx";


            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
   
            //產生 Excel Report
            ExcelReport.ExcelReportOutput(dtCost.DefaultView.ToTable(), ExcelTemplate, OutPutFile, "N");
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               string  FileName = openFileDialog1.FileName;


                WriteExcelProduct4(FileName);

            }
        }

        private void WriteExcelProduct4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string CUR;

                string AMT;

                string CARD;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    CUR = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    AMT = range.Text.ToString().Replace(",", "").Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    CARD = range.Text.ToString().Trim();

                    if (CUR != "" && AMT != "" && CARD != "")
                    {
                        AddTEMPG1(CARD,CUR,AMT);

                    }
                }




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

        public void AddTEMPG1(string CARDCODE, string U_CURRENCY,string U_AMOUNT)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OCRD SET U_CURRENCY=@U_CURRENCY,U_AMOUNT=@U_AMOUNT WHERE CARDCODE=@CARDCODE ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@U_CURRENCY", U_CURRENCY));
            command.Parameters.Add(new SqlParameter("@U_AMOUNT", U_AMOUNT));
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
    }
}

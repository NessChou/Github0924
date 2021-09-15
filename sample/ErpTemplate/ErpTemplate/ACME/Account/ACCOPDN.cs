using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Management;
using System.Diagnostics;
using System.Text.RegularExpressions;
namespace ACME
{
    public partial class ACCOPDN : Form
    {
        System.Data.DataTable T1 = null;
        System.Data.DataTable T2 = null;
        System.Data.DataTable T3 = null;
        System.Data.DataTable T4 = null;
        DataRow dr = null;
        public ACCOPDN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            F1("已結","1");
            F1("未結", "1");

            F1("已結", "2");
            F1("未結", "2");
            dataGridView1.DataSource = T1;
            dataGridView2.DataSource = T2;
            dataGridView3.DataSource = GetDORPC();
            dataGridView4.DataSource = T3;
            dataGridView5.DataSource = T4;
            for (int i = 8; i <= 13; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";

            }
            for (int i = 22; i <= 23; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";

            }
            for (int i = 8; i <= 13; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

            for (int i = 22; i <= 23; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

            for (int i = 9; i <= 13; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }


            for (int i = 6; i <= 9; i++)
            {
                DataGridViewColumn col = dataGridView4.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

  
                DataGridViewColumn col3 = dataGridView4.Columns[16];
            col3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col3.DefaultCellStyle.Format = "#,##0.00";






            for (int i = 6; i <= 9; i++)
            {
                DataGridViewColumn col = dataGridView5.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

            DataGridViewColumn col1 = dataGridView5.Columns[16];
            col1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col1.DefaultCellStyle.Format = "#,##0.00";




            DataGridViewColumn colS = dataGridView3.Columns[19];


            colS.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            colS.DefaultCellStyle.Format = "#,##0";
        }

        private void F1(string DOCTYPE,string DTYPE)
        {
            System.Data.DataTable dt = null;
            if (DOCTYPE == "已結")
            {
                if (DTYPE == "1")
                {
                    dt = GetDATA(DOCTYPE);
                }
                if (DTYPE == "2")
                {
                    dt = GetDATAS(DOCTYPE);
                }
            }
            if (DOCTYPE == "未結")
            {

                if (DTYPE == "1")
                {
                    dt = GetDATA(DOCTYPE);
                }
                if (DTYPE == "2")
                {
                    dt = GetDATAS(DOCTYPE);
                }
            }

            System.Data.DataTable dtCost = null;

            if (DTYPE == "1")
            {
                dtCost = MakeTable();
            }

            if (DTYPE == "2")
            {
                dtCost = MakeTable2();
            }
            string DUP = "";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string DOC = dt.Rows[i]["收貨採購單單號"].ToString();
                if (DTYPE == "1")
                {
                    if (DUP == "DOC")
                    {
                        dr["折扣"] = "";
                    }
                    else
                    {
                        dr["折扣"] = Convert.ToInt32(dt.Rows[i]["折扣"]);
                        DUP = "";
                    }
                    dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                    dr["品名描述"] = dt.Rows[i]["品名描述"].ToString();
                    dr["借項"] = dt.Rows[i]["借項"].ToString();
                    dr["貸項"] = dt.Rows[i]["貸項"].ToString();
                    dr["單價"] = dt.Rows[i]["單價"].ToString();
                }
                if (DTYPE == "2")
                {
                    dr["LC"] = dt.Rows[i]["LC"].ToString();
                }
                    dr["收貨採購單單號"] = DOC;
                dr["收貨採購單傳票"] = dt.Rows[i]["收貨採購單傳票"].ToString();
                dr["AP傳票"] = dt.Rows[i]["AP傳票"].ToString();
                dr["過帳日期"] = dt.Rows[i]["過帳日期"].ToString();
        
                dr["廠商編號"] = dt.Rows[i]["廠商編號"].ToString();
                dr["廠商名稱"] = dt.Rows[i]["廠商名稱"].ToString();
                dr["數量"] = dt.Rows[i]["數量"].ToString();
                dr["未稅總計"] = Convert.ToInt32(dt.Rows[i]["未稅總計"]);
                dr["稅額"] = Convert.ToInt32(dt.Rows[i]["稅額"]);
                dr["總計"] = Convert.ToInt32(dt.Rows[i]["總計"]);
                dr["原幣金額"] = Convert.ToDecimal(dt.Rows[i]["原幣金額"]);
                dr["倉庫名稱"] = dt.Rows[i]["倉庫名稱"].ToString();

                dr["InvoiceNo"] = dt.Rows[i]["InvoiceNo"].ToString();
                dr["日期"] = dt.Rows[i]["日期"].ToString();
                dr["發票號碼"] = dt.Rows[i]["發票號碼"].ToString();
                dr["發票日期"] = dt.Rows[i]["發票日期"].ToString();

                if (String.IsNullOrEmpty(dr["發票日期"].ToString()))
                {

                    dr["發票日期"] = dt.Rows[i]["日期"].ToString();

                }

                System.Data.DataTable gg1 = GetOPTW2(DOC);
                if (gg1.Rows.Count > 0)
                {
                    //__________
                    if (String.IsNullOrEmpty(dr["發票號碼"].ToString()))
                    {
                        dr["發票號碼"] = gg1.Rows[0]["檔案名稱"].ToString();

                    }

                }

       
                UPOPDN(dr["發票號碼"].ToString(),DOC);
                dr["收採發票號碼"] = GetOPDNINV(DOC).Rows[0][0].ToString();
                dr["原始幣別"] = dt.Rows[i]["原始幣別"].ToString();
                dr["匯率"] = dt.Rows[i]["匯率"].ToString();
                string SHIPNO = dt.Rows[i]["SHIPNO"].ToString();
                System.Data.DataTable FF1 = GetDATAF(SHIPNO);
                if (FF1.Rows.Count > 0)
                {
                    dr["貿易條件"] = FF1.Rows[0]["貿易條件"].ToString();
                }
                DUP = DOC;
                dtCost.Rows.Add(dr);
            }

            if (DOCTYPE == "已結")
            {

                if (DTYPE == "1")
                {
                    T1 = dtCost;
                }
                if (DTYPE == "2")
                {
                    T3 = dtCost;
                }
          
            }
            if (DOCTYPE == "未結")
            {
                if (DTYPE == "1")
                {
                    T2 = dtCost;
                }
                if (DTYPE == "2")
                {
                    T4 = dtCost;
                }
            }
        }

    

        private System.Data.DataTable GetDORPC()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE T0.DOCTYPE WHEN 'I' THEN '項目' ELSE '服務' END 文件類型,T0.[DocEntry] AP貸項通知單,T0.transid AP貸項通知單傳票 ");
            sb.Append(" ,(SELECT TOP 1 U_BSRN1 FROM [@CADMEN_PMD] J1 LEFT JOIN [@CADMEN_PMD1]  J2 ON (J1.DOCENTRY=J2.DOCENTRY) WHERE (T0.DOCENTRY=J1.U_BSREN )) 原AP傳票 ");
            sb.Append(" ,convert(varchar,T0.docdate, 102) 過帳日期,T1.[ItemCode] 產品編號, T1.[Dscription] 品名描述,T0.cardcode 廠商編號,T0.cardname 廠商名稱,T1.[Quantity] 數量, T1.[Price] 單價,T1.[LineTotal] 未稅總計,t1.vatsum 稅額,T1.[LineTotal]+t1.vatsum 總計, T1.[U_ACME_WhsName] 倉庫名稱 ");
            sb.Append(" ,(select TOP 1 ACCTNAME from JDT1 J1 LEFT JOIN OACT J2 ON (J1.ACCOUNT=J2.ACCTCODE) WHERE (T0.TRANSID=J1.TRANSID AND DEBIT <> 0)) 借項 ");
            sb.Append(" ,(select TOP 1 ACCTNAME from JDT1 J1 LEFT JOIN OACT J2 ON (J1.ACCOUNT=J2.ACCTCODE) WHERE (T0.TRANSID=J1.TRANSID AND CREDIT <> 0)) 貸項, ");
            sb.Append(" CASE WHEN U_RP_BSTY1 = 0 THEN '退貨'  WHEN U_RP_BSTY1 = 1  THEN '折讓' END 退折類別, ");
            sb.Append(" CAST(ISNULL(T1.U_ACME_INV,0) AS DECIMAL(10,2)) 美金金額, ");
            sb.Append(" CASE WHEN T0.DocType='S' THEN ROUND(CASE WHEN T1.U_ACME_INV='0' THEN '0' ELSE  CASE WHEN ISNUMERIC(T1.U_ACME_INV)=1 THEN T1.[LineTotal]/CAST(ISNULL(T1.U_ACME_INV,0) AS decimal(18,4)) ELSE 0 END END,2) ");
            sb.Append(" ELSE  ROUND(CASE WHEN T1.U_ACME_INV='0' THEN '0' ELSE  CASE WHEN ISNUMERIC(T1.U_ACME_INV)=1 THEN T1.[LineTotal]/((CAST(ISNULL(T1.U_ACME_INV,0) AS decimal(18,4)))*T1.[Quantity])  ELSE 0 END END,2) END 匯率 ");
            sb.Append(" FROM ORPC T0 ");
            sb.Append(" left join RPC1 T1 on t0.docentry=t1.docentry ");
            sb.Append(" left join oitm t3 on t3.itemcode=t1.itemcode ");
            sb.Append(" where SUBSTRING(T0.CARDCODE,1,1)='S'  ");
            sb.Append(" AND Convert(varchar(10),T0.docdate,112) between @t1 and @t2");
            if (textBox7.Text != "")
            {
                sb.Append(" and T0.CARDNAME like '%" + textBox7.Text + "%'  ");
            }

            sb.Append(" order by T0.[DocEntry]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOPCH()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT Convert(varchar(10),CASE ISNULL(T1.[U_PC_BSAPP],'') WHEN '' THEN  P.[U_PC_BSAPP] ELSE T1.[U_PC_BSAPP] END,111)  as 申報年月");
            sb.Append("	,Convert(varchar(10),j.refdate ,111)過帳日期,CASE J.TRANSTYPE WHEN 18 THEN 'PU' WHEN 30 THEN 'JE' END 來源代碼, ");
            sb.Append("              P1.DOCENTRY 單號,P1.[CardCode] 廠商代號,P1.[CardName] 廠商名稱, ");
            sb.Append("              J.TransID 傳票號碼,''''+ CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.[U_PC_BSNOT]  ELSE T1.[U_PC_BSNOT] END  as 統一編號,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY1  ");
            sb.Append("              WHEN 0 THEN '三聯式發票/電子計算機發票'  ");
            sb.Append("              WHEN 1 THEN '三聯式收銀機發票'  ");
            sb.Append("              WHEN 2 THEN '有稅憑證'  ");
            sb.Append("              WHEN 3 THEN '海關代徵稅'  ");
            sb.Append("              WHEN 4 THEN '免用統一發票/收據'  ");
            sb.Append("              WHEN 8 THEN '一般稅額計算之電子發票'  ");
            sb.Append("              END ELSE CASE T1.U_PC_BSTY1  ");
            sb.Append("              WHEN 0 THEN '三聯式發票/電子計算機發票'  ");
            sb.Append("              WHEN 1 THEN '三聯式收銀機發票'  ");
            sb.Append("              WHEN 2 THEN '有稅憑證'  ");
            sb.Append("              WHEN 3 THEN '海關代徵稅'  ");
            sb.Append("              WHEN 4 THEN '免用統一發票/收據'  ");
            sb.Append("              WHEN 8 THEN '一般稅額計算之電子發票'  ");
            sb.Append("              END END 憑證類別,  ");
            sb.Append("              Convert(varchar(10),CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSDAT ELSE T1.U_PC_BSDAT END,111)  發票日期, ");
            sb.Append("              CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSINV ELSE T1.U_PC_BSINV END as 發票號碼, ");
            sb.Append("              cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int)  未稅金額, ");
            sb.Append("              cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int)   稅額, ");
            sb.Append("              cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)  含稅總額, ");
            sb.Append("              CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4  ");
            sb.Append("              WHEN 0  THEN '進貨' ");
            sb.Append("              WHEN 1  THEN '費用' ");
            sb.Append("              WHEN 2  THEN '固定資產' ");
            sb.Append("              END ELSE CASE T1.U_PC_BSTY4  ");
            sb.Append("              WHEN 0  THEN '進貨' ");
            sb.Append("              WHEN 1  THEN '費用' ");
            sb.Append("              WHEN 2  THEN '固定資產' ");
            sb.Append("              END END 進項科目, ");
            sb.Append("              CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSCUS ELSE T1.U_PC_BSCUS END  海關代徵號碼,J.u_acme_user 製單人,P2.BASE 收貨採購單 FROM OJDT  J ");
            sb.Append("              LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN] ");
            sb.Append("              LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND T1.U_PC_BSTY4='0' AND Convert(varchar(8),T1.[U_PC_BSAPP],112) =@t1)  ");
            sb.Append("              LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID AND P.U_PC_BSTY4='0' AND Convert(varchar(8),P.[U_PC_BSAPP],112) =@t1) ");
            sb.Append("              LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID) ");
            sb.Append("              LEFT JOIN (SELECT MAX(BaseEntry) BASE,DOCENTRY   FROM PCH1 WHERE BaseType =20 GROUP BY DOCENTRY ");
            sb.Append("              ) P2 ON (P1.DOCENTRY=P2.DOCENTRY) ");
            sb.Append("              WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0 AND J.TRANSTYPE IN (18,30) ");
            sb.Append("               ORDER BY CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTY1 ELSE T1.U_PC_BSTY1 END ,J.TransID, ");
            sb.Append("               CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSDAT ELSE T1.U_PC_BSDAT END,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox3.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDATA(string DOCTYPE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[DocEntry] 收貨採購單單號,t1.transid 收貨採購單傳票,t5.transid AP傳票,Convert(varchar(8),t1.docdate,112)  過帳日期,T1.DISCSUM 折扣,  ");
            sb.Append(" T0.[ItemCode] 產品編號, T0.[Dscription] 品名描述,t1.cardcode 廠商編號,t1.cardname 廠商名稱  ");
            sb.Append(" ,T0.[Quantity] 數量, T0.[Price] 單價,T0.[LineTotal] 未稅總計,t0.vatsum 稅額,T0.[LineTotal]+t0.vatsum 總計 ");
            sb.Append("  ,ISNULL(T8.PRICE*T0.[Quantity],0)  原幣金額,");
            sb.Append(" T0.[U_ACME_WhsName]  倉庫名稱  ");
            sb.Append(" ,(select TOP 1 ACCTNAME from JDT1 J1 LEFT JOIN OACT J2 ON (J1.ACCOUNT=J2.ACCTCODE) WHERE (T1.TRANSID=J1.TRANSID AND DEBIT <> 0)) 借項  ");
            sb.Append(" ,(select TOP 1 ACCTNAME from JDT1 J1 LEFT JOIN OACT J2 ON (J1.ACCOUNT=J2.ACCTCODE) WHERE (T1.TRANSID=J1.TRANSID AND CREDIT <> 0)) 貸項 ");
            sb.Append(" ,T5.U_ACME_INV InvoiceNo,Convert(varchar(8),T5.u_acme_invoice,112) 日期,T5.U_PC_BSINV 發票號碼 ");
            sb.Append(" ,Convert(varchar(8),T5.U_PC_BSDAT,112)  發票日期,T8.Currency 原始幣別,CASE WHEN ISNULL(T8.PRICE*T0.[Quantity],0)=0 THEN 0 ELSE  ROUND((T0.[LineTotal]/ISNULL(T8.PRICE*T0.[Quantity],0)),2) END 匯率,T1.U_Shipping_no SHIPNO   FROM Pdn1 T0  ");
            sb.Append(" left join opdn t1 on t0.docentry=t1.docentry ");
            sb.Append(" left join ocrd t2 on t1.cardcode=t2.cardcode  ");
            sb.Append(" left join oitm t3 on t3.itemcode=t0.itemcode ");
            sb.Append(" left join pch1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum  and t4.basetype='20' ) ");
            sb.Append(" left join opch t5 on (t4.docentry=T5.docentry ) ");
            sb.Append(" LEFT JOIN POR1 T8 ON (T8.docentry=T0.baseentry AND T8.linenum=T0.baseline)   ");
            sb.Append(" where substring(t1.cardcode,1,1) in ('S','U') ");
            sb.Append(" And t0.itemcode not in (select itemcode from oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z')) ");
            sb.Append(" And substring(t0.itemcode,1,2) <> 'ZR' ");
            if (DOCTYPE == "已結")
            {
                sb.Append(" AND T1.DOCSTATUS='C'");
            }
            if (DOCTYPE == "未結")
            {
                sb.Append(" AND T1.DOCSTATUS='O'");
            }
            sb.Append(" AND Convert(varchar(10),t1.docdate,112) between @t1 and @t2");
            if (textBox7.Text != "")
            {
                sb.Append(" and T1.CARDNAME like '%" + textBox7.Text + "%'  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public void UPOPDN(string U_PC_BSINV, string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE OPDN SET U_PC_BSINV=@U_PC_BSINV WHERE DOCENTRY=@DOCENTRY ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_PC_BSINV", U_PC_BSINV));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));



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
        private System.Data.DataTable GetDATAS(string DOCTYPE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                                     SELECT  MAX(T0.[DocEntry]) 收貨採購單單號,MAX(t1.transid) 收貨採購單傳票,MAX(t5.transid) AP傳票,MAX(Convert(varchar(8),t1.docdate,112))  過帳日期,   ");
            sb.Append("                                                    MAX(t1.cardcode) 廠商編號,MAX(t1.cardname) 廠商名稱      ");
            sb.Append("                                                     ,SUM(T0.[Quantity]) 數量, MAX(T1.DOCTOTAL-T1.VATSUM) 未稅總計,MAX(T1.VATSUM) 稅額,MAX(T1.DOCTOTAL) 總計     ");
            sb.Append("                                                      ,CAST(round(max(t1.doctotal/t1.u_acme_rate1),2) AS decimal(14,2))  原幣金額,    ");
            sb.Append("                                                     MAX(T0.[U_ACME_WhsName])  倉庫名稱      ");
            sb.Append("                                                     ,MAX(T1.U_ACME_INV) InvoiceNo,MAX(Convert(varchar(8),T1.u_acme_invoice,112)) 日期,MAX(T5.U_PC_BSINV) 發票號碼     ");
            sb.Append("                                                     ,MAX(Convert(varchar(8),T5.U_PC_BSDAT,112))  發票日期,MAX(T8.Currency) 原始幣別");
            sb.Append("		             				 ,MAX(t1.u_acme_rate1) 匯率  ");
            //sb.Append("		             													 ,''''+CAST(CAST(ROUND(MAX(CASE WHEN ISNULL(T8.PRICE*T0.[Quantity],0)=0 THEN 0 ELSE  ROUND((T0.[LineTotal]/ISNULL(T8.PRICE*T0.[Quantity],0)),2) END),5) AS DECIMAL(18,3)) AS VARCHAR) 匯率 ");
            sb.Append("													 ,MAX(T1.U_Shipping_no) SHIPNO,MAX(T1.U_ACME_LC) LC   FROM Pdn1 T0      ");
            sb.Append("                                                     left join opdn t1 on t0.docentry=t1.docentry     ");
            sb.Append("                                                     left join ocrd t2 on t1.cardcode=t2.cardcode      ");
            sb.Append("                                                     left join oitm t3 on t3.itemcode=t0.itemcode     ");
            sb.Append("                                                     left join pch1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum  and t4.basetype='20' )     ");
            sb.Append("                                                     left join opch t5 on (t4.docentry=T5.docentry )     ");
            sb.Append("                                                     LEFT JOIN POR1 T8 ON (T8.docentry=T0.baseentry AND T8.linenum=T0.baseline)       ");
            sb.Append("                                                     where substring(t1.cardcode,1,1) in ('S','U')     ");
            sb.Append("                                                     And t0.itemcode not in (select itemcode from oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))     ");
            sb.Append("                                                     And substring(t0.itemcode,1,2) <> 'ZR'     ");



            if (DOCTYPE == "已結")
            {
                sb.Append(" AND T1.DOCSTATUS='C'");
            }
            if (DOCTYPE == "未結")
            {
                sb.Append(" AND T1.DOCSTATUS='O'");
            }
            sb.Append(" AND Convert(varchar(10),t1.docdate,112) between @t1 and @t2");
            if (textBox7.Text != "")
            {
                sb.Append(" and T1.CARDNAME like '%" + textBox7.Text + "%'  ");
            }
            sb.Append("   GROUP BY T0.DOCENTRY  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDATAF(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT TradeCondition 貿易條件 FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE  ");
  

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("收貨採購單單號", typeof(string));
            dt.Columns.Add("收貨採購單傳票", typeof(string));
            dt.Columns.Add("AP傳票", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名描述", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("數量", typeof(decimal ));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("未稅總計", typeof(int));
            dt.Columns.Add("稅額", typeof(int));
            dt.Columns.Add("總計", typeof(int));
            dt.Columns.Add("折扣", typeof(int));
            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("倉庫名稱", typeof(string));
            dt.Columns.Add("借項", typeof(string));
            dt.Columns.Add("貸項", typeof(string));
            dt.Columns.Add("原始幣別", typeof(string));
            dt.Columns.Add("原幣金額", typeof(decimal));
            dt.Columns.Add("匯率", typeof(decimal));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("收採發票號碼", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("收貨採購單單號", typeof(string));
            dt.Columns.Add("收貨採購單傳票", typeof(string));
            dt.Columns.Add("AP傳票", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));

            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));

            dt.Columns.Add("未稅總計", typeof(int));
            dt.Columns.Add("稅額", typeof(int));
            dt.Columns.Add("總計", typeof(int));

            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("倉庫名稱", typeof(string));

            dt.Columns.Add("原始幣別", typeof(string));
            dt.Columns.Add("原幣金額", typeof(decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("LC", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("收採發票號碼", typeof(string));
            return dt;
        }
        private void ACCOPDN_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            textBox3.Text = DateTime.Now.ToString("yyyyMM") + "15";
        }

        private void button2_Click(object sender, EventArgs e)
        {
                     if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
               else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToExcelSHARON2(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToExcelSHARON2(dataGridView5);
            }
        }
        public System.Data.DataTable GetOPTW(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱 from oclg t2     ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)     ");
            sb.Append(" where  T2.DOCTYPE='20'  ");
            sb.Append(" and   t2.docentry=@docentry and  T3.[FILENAME]  LIKE '%INV%'    ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetOPTWS(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("			        select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱 from OPDN t2      ");
            sb.Append("              LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)      ");
            sb.Append("              where  t2.docentry=@docentry and  T3.[FILENAME]  LIKE '%INV%'    ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetOUT(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("			        select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱2 from OPDN t2      ");
            sb.Append("              LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)      ");
            sb.Append("              where  t2.docentry=@docentry    ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOPTW2(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select REPLACE(REPLACE(REPLACE(T3.FILENAME,'進',''),T2.DOCENTRY,''),'_','') 檔案名稱,cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱2 from oclg t2     ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)     ");
            sb.Append(" where  T2.DOCTYPE='20'  ");
            sb.Append("              and   t2.docentry=@docentry  and  (T3.[FILENAME] NOT LIKE '%PK%'  and  T3.[FILENAME] NOT LIKE '%INV%')     ");
            sb.Append("			  and  LEN(REPLACE(REPLACE(REPLACE(T3.FILENAME,'進',''),T2.DOCENTRY,''),'_',''))=10");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetOPDNINV(string DOCENTRY)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
          
            sb.Append(" SELECT U_PC_BSINV FROM OPDN WHERE DOCENTRY=@DOCENTRY    ");
           

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOPTW3(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select REPLACE(REPLACE(REPLACE(T3.FILENAME,'進',''),T2.DOCENTRY,''),'_','') 檔案名稱 from oclg t2     ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)     ");
            sb.Append(" where  T2.DOCTYPE='20'  ");
            sb.Append("              and   t2.docentry=@docentry  and  (T3.[FILENAME] NOT LIKE '%PK%'  and  T3.[FILENAME] NOT LIKE '%INV%')     ");
            sb.Append("			  and  LEN(REPLACE(REPLACE(REPLACE(T3.FILENAME,'進',''),T2.DOCENTRY,''),'_',''))=10");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void DELETEFILE()
        {
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        public static void Print(string filePath)
        {
            // Status = PrintJobStatus.Printing;
            // Message = string.Empty;
            try
            {
                //     logger.Debug($"Printing... {filePath}");
                ProcessStartInfo info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = filePath;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;

                Process p = new Process();
                p.StartInfo = info;
                p.Start();

                p.WaitForInputIdle();
                //以下邏輯克服無法得知Acrobat Reader或Foxit Reader是否列印完成的問題
                //最多等待180秒（假設所有檔案可在3分鐘內印完）
                // var timeOut = DateTime.Now.AddSeconds(180);
                var timeOut = DateTime.Now.AddSeconds(5);
                bool printing = false; //是否開始列印
                bool done = false; //是否列印完成
                //取純檔名部分，跟PrintQueue進行比對
                string pureFileName = Path.GetFileName(filePath);
                //限定最大等待時間
                while (DateTime.Now.CompareTo(timeOut) < 0)
                {
                    if (!printing)
                    {
                        //未開始列印前發現檔名相同的列印工作
                        if (CheckPrintQueue(pureFileName))
                        {
                            printing = true;
                            //        Console.WriteLine($"[{pureFileName}]列印中...");
                        }
                    }
                    else
                    {
                        //已開始列印後，同檔名列印工作消失表示列印完成
                        if (!CheckPrintQueue(pureFileName))
                        {
                            done = true;
                            //     Console.WriteLine($"[{pureFileName}]列印完成");
                            break;
                        }
                    }
                    System.Threading.Thread.Sleep(100);
                }
                try
                {
                    //若程序尚未關閉，強制關閉之
                    if (false == p.CloseMainWindow())
                        p.Kill();
                }
                catch
                {
                }
                if (!done)
                {
                    //    Console.WriteLine($"無法確認報表[{pureFileName}]列印狀態！");
                }
            }
            catch (Exception ex)
            {
                //        Console.WriteLine($"Error: {DateTime.Now:HH:mm:ss} {ex.Message}");
            }
        }

        private static bool CheckPrintQueue(string file)
        {
            string Job_Name = "";
            string searchQuery =
                "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher printJobs = new ManagementObjectSearcher(searchQuery);
            foreach (ManagementObject mo in printJobs.Get())
            {
                Job_Name = mo.Properties["Document"].Value.ToString();
            }

            if (Job_Name == file)
            { return true; }
            else
                return false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("請確認否要列印", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                DELETEFILE();
                for (int i = dataGridView4.SelectedRows.Count - 1; i >= 0; i--)
                {

                    DataGridViewRow row;

                    row = dataGridView4.SelectedRows[i];
                    string 收貨採購單單號 = row.Cells["收貨採購單單號"].Value.ToString();
                    System.Data.DataTable gg1 = null;
                    System.Data.DataTable gg2 = null;

                    gg1 = GetOPTW(收貨採購單單號);
                    if (gg1.Rows.Count == 0)
                    {
                        gg1 = GetOPTWS(收貨採購單單號);
                    }
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        Print(NewFileName);
                    }


                    gg2 = GetOPTW2(收貨採購單單號);
                    if (gg2.Rows.Count > 0)
                    {
                        string path = gg2.Rows[0]["path"].ToString();
                        string 路徑 = gg2.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        Print(NewFileName);
                    }

             

                }

                for (int i = dataGridView5.SelectedRows.Count - 1; i >= 0; i--)
                {

                    DataGridViewRow row;

                    row = dataGridView5.SelectedRows[i];
                    string 收貨採購單單號 = row.Cells["收貨採購單單號"].Value.ToString();

                    System.Data.DataTable gg1 = null;
                    System.Data.DataTable gg2 = null;


                    gg1 = GetOPTW(收貨採購單單號);
                    if (gg1.Rows.Count == 0)
                    {
                        gg1 = GetOPTWS(收貨採購單單號);
                    }
                    if (gg1.Rows.Count > 0)
                    {
                        for (int i2 = 0; i2 <= gg1.Rows.Count - 1; i2++)
                        {
                            string path = gg1.Rows[i2]["path"].ToString();
                            string 路徑 = gg1.Rows[i2]["路徑"].ToString();
                            string 檔案名稱 = gg1.Rows[i2]["檔案名稱"].ToString();

                            string aa = path + "\\" + 路徑;


                            string filename = 檔案名稱;
                            string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                            System.IO.File.Copy(aa, NewFileName, true);

                            Print(NewFileName);
                        }

                    }



                    gg2 = GetOPTW2(收貨採購單單號);
                    if (gg2.Rows.Count > 0)
                    {
                        string path = gg2.Rows[0]["path"].ToString();
                        string 路徑 = gg2.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        Print(NewFileName);
                    }

                }
            }
        }

        private void dataGridView5_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "InvoiceNo")
                {
                    string 收貨採購單單號 = dataGridView5.CurrentRow.Cells["收貨採購單單號"].Value.ToString();

                    System.Data.DataTable gg1 = null;
                    System.Data.DataTable gg2 = null;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


                    gg1 = GetOPTW(收貨採購單單號);
                    if (gg1.Rows.Count == 0)
                    {
                        gg1 = GetOPTWS(收貨採購單單號);
                    }
                    if (gg1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= gg1.Rows.Count - 1; i++)
                        {
                            string path = gg1.Rows[i]["path"].ToString();
                            string 路徑 = gg1.Rows[i]["路徑"].ToString();
                            string 檔案名稱 = gg1.Rows[i]["檔案名稱"].ToString();

                            string aa = path + "\\" + 路徑;

                       
                            string filename = 檔案名稱;
                            string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                            System.IO.File.Copy(aa, NewFileName, true);

                            System.Diagnostics.Process.Start(NewFileName);
                        }

                    }


                    gg2 = GetOPTW2(收貨採購單單號);
                    if (gg2.Rows.Count > 0)
                    {
                        string path = gg2.Rows[0]["path"].ToString();
                        string 路徑 = gg2.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);
                    }


                }
            }
            catch { }
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "InvoiceNo")
                {
                    string 收貨採購單單號 = dataGridView4.CurrentRow.Cells["收貨採購單單號"].Value.ToString();
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    System.Data.DataTable gg1 = null;
                    System.Data.DataTable gg2 = null;

                  

                    gg1 = GetOPTW(收貨採購單單號);
                    if (gg1.Rows.Count == 0)
                    {
                        gg1 = GetOPTWS(收貨採購單單號);
                    }
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();

                        string aa = path + "\\" + 路徑;

  
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);

                    }

                    gg2 = GetOPTW2(收貨採購單單號);
                    if (gg2.Rows.Count > 0)
                    {
                        string path = gg2.Rows[0]["path"].ToString();
                        string 路徑 = gg2.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);
                    }

                }
            }
            catch { }
        }

        private void dataGridView4_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }



        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void dataGridView5_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
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

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView6.DataSource = GetOPCH();
        }

        private void dataGridView6_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "發票號碼")
                {
                    string 收貨採購單 = dataGridView6.CurrentRow.Cells["收貨採購單"].Value.ToString();
                    string 廠商代號 = dataGridView6.CurrentRow.Cells["廠商代號"].Value.ToString();

                    System.Data.DataTable gg2 = null;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    //S0001 S0623-GD
                    string CARD = 廠商代號.Substring(0, 5);

                    if (CARD == "S0001" || CARD == "S0623")
                    {
                        gg2 = GetOPTW2(收貨採購單);
                    }
                    else
                    {
                        
                           gg2 = GetOUT(收貨採購單);
                    }
           
                    if (gg2.Rows.Count > 0)
                    {
                        string path = gg2.Rows[0]["path"].ToString();
                        string 路徑 = gg2.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                        string aa = path + "\\" + 路徑;
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);
                    }

          



                }
            }
            catch { }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView6);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result;
                result = MessageBox.Show("請確認否要列印", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {

                    DELETEFILE();


                    for (int i = dataGridView6.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        DataGridViewRow row;
                        row = dataGridView6.SelectedRows[i];
                        string 收貨採購單 = row.Cells["收貨採購單"].Value.ToString();
                        string 廠商代號 = row.Cells["廠商代號"].Value.ToString();


                        System.Data.DataTable gg2 = null;
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        //S0001 S0623-GD
                        string CARD = 廠商代號.Substring(0, 5);

                        if (CARD == "S0001" || CARD == "S0623")
                        {
                            gg2 = GetOPTW2(收貨採購單);
                        }
                        else
                        {

                            gg2 = GetOUT(收貨採購單);
                        }

                        if (gg2.Rows.Count > 0)
                        {
                            string path = gg2.Rows[0]["path"].ToString();
                            string 路徑 = gg2.Rows[0]["路徑"].ToString();
                            string 檔案名稱 = gg2.Rows[0]["檔案名稱2"].ToString();
                            string aa = path + "\\" + 路徑;
                            string filename = 檔案名稱;
                            string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
                            System.IO.File.Copy(aa, NewFileName, true);
                            Print(NewFileName);
                        }





                    }
                }
            }
            catch { }
        }
    }
}

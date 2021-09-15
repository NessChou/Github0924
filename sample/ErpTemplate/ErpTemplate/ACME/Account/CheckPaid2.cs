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
    public partial class CheckPaid2 : Form
    {
        string ssd;
        private decimal sd;
        private decimal sdf;
        private decimal sd2;
        private decimal sc;
        private decimal sk;
        public CheckPaid2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GETGRIDVIEW(dataGridView1);
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
        
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("應收帳款2", typeof(Decimal));
            dt.Columns.Add("美金應收帳款2", typeof(Decimal));
            dt.Columns.Add("收款條件", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("數量加總", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("美金平均單價", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("業管", typeof(string));
            dt.Columns.Add("訂單號碼", typeof(string));
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("發票總類", typeof(string));
            dt.Columns.Add("invoice", typeof(string));
            dt.Columns.Add("最終客戶", typeof(string));
            dt.Columns.Add("逾期日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(string));
            dt.Columns.Add("國家", typeof(string));
            dt.Columns.Add("出口報單類別", typeof(string));
            dt.Columns.Add("出口證明文件號碼", typeof(string));
            dt.Columns.Add("SHIPTO", typeof(string));
            dt.Columns.Add("SAP1", typeof(string));
  
            return dt;
        }
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("SHIPTO", typeof(string));
            dt.Columns.Add("發票總類", typeof(string));
            return dt;
        }

        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("            select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,CAST(T0.DOCTOTAL AS INT) 應收帳款,CAST(T0.DOCTOTAL-T0.VATSUM AS INT) 應收帳款2  from oinv t0");
            sb.Append("            left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("            left JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append("             where  Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2 ");

            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
      
            sb.Append(" union all");
            sb.Append("            select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,CAST(T0.DOCTOTAL AS INT)*-1 應收帳款,CAST(T0.DOCTOTAL-T0.VATSUM AS INT)*-1 應收帳款2  from orin t0");
            sb.Append("                           left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("                           left JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append("             where  Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2 ");
            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

//        private System.Data.DataTable GetOrderDataAPF()
//        {
//            SqlConnection connection = globals.shipConnection;
//            StringBuilder sb = new StringBuilder();


//    sb.Append(" select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,CAST(T0.DOCTOTAL AS INT) 應收帳款,CAST(T0.DOCTOTAL-T0.VATSUM AS INT) 應收帳款2  from oinv t0 ");
//sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode  ");
//sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
//sb.Append(" LEFT JOIN RCT2 T4 ON (T0.DOCENTRY=T4.DOCENTRY AND T4.InvType =13)");
//sb.Append(" INNER  JOIN  ORCT T5 ON (T4.DOCNUM=T5.DOCENTRY)");
//sb.Append(" where  Convert(varchar(8),T5.docdate,112)  =@DocDate1 ");

//            sb.Append(" union all");
//            sb.Append("            select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,CAST(T0.DOCTOTAL AS INT)*-1 應收帳款,CAST(T0.DOCTOTAL-T0.VATSUM AS INT) 應收帳款2  from orin t0");
//            sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode  ");
//            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
//            sb.Append(" LEFT JOIN RCT2 T4 ON (T0.DOCENTRY=T4.DOCENTRY AND T4.InvType =14)");
//            sb.Append(" INNER  JOIN  ORCT T5 ON (T4.DOCNUM=T5.DOCENTRY)");
//            sb.Append(" where  Convert(varchar(8),T5.docdate,112)  =@DocDate1 ");


//            SqlCommand command = new SqlCommand(sb.ToString(), connection);
//            command.CommandType = CommandType.Text;
//            command.Parameters.Add(new SqlParameter("@DocDate1", GetMenu.Day() ));

//            SqlDataAdapter da = new SqlDataAdapter(command);

//            DataSet ds = new DataSet();
//            try
//            {
//                connection.Open();
//                da.Fill(ds, "oinv");
//            }
//            finally
//            {
//                connection.Close();
//            }

//            return ds.Tables[0];

//        }
        private System.Data.DataTable GetOrderDataAP1(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append("  dbo.fun_CreditDate(CASE WHEN T0.DOCENTRY=49839 AND T1.PRICE=0 THEN 'OA 30 days' WHEN T0.DOCENTRY=50079 AND T1.Quantity =299 THEN 'OA 30 days' ELSE t9.u_acme_pay END,T0.CardCode,T0.DocDate) 逾期日期,T0.u_in_bscls 出口報單類別,T0.u_in_bsren 出口證明文件號碼,T0.u_acme_shipto1 SHIPTO");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='15' ");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("             ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類, T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append(" dbo.fun_CreditDate(T0.u_acme_pay,T0.CardCode,T0.DocDate) 逾期日期,T0.u_in_bscls 出口報單類別,T0.u_in_bsren 出口證明文件號碼,T0.u_acme_shipto1 SHIPTO");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='17' ");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類, T0.COMMENTS 備註,'' 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv   美金單價,T0.JRNLMEMO 應收總計,'' 訂單號碼,'' 最終客戶, ");
            sb.Append("                       t0.docdate 逾期日期,T0.u_in_bscls 出口報單類別,T0.u_in_bsren 出口證明文件號碼,T0.u_acme_shipto1 SHIPTO");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype =-1 ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                           ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                            T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv 美金單價,T0.JRNLMEMO 摘要,cast(T0.u_acme_arap as varchar) 訂單號碼,'' 最終客戶,t0.docdate 逾期日期,T0.u_in_bscls 出口報單類別,T0.u_in_bsren 出口證明文件號碼,T0.u_acme_shipto1 SHIPTO  FROM Orin T0  ");
            sb.Append("                           LEFT JOIN rin1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("                         where  t0.docentry=@docentry and t0.objtype=@bb");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("             ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類, T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append(" dbo.fun_CreditDate(T0.u_acme_pay,T0.CardCode,T0.DocDate) 逾期日期,T0.u_in_bscls 出口報單類別,T0.u_in_bsren 出口證明文件號碼,T0.u_acme_shipto1 SHIPTO");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN QUT1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN OQUT T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='23' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
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
  
        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex ==1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }
  
        private void CheckPaid_Load(object sender, EventArgs e)
        {
            label6.Text = "";
            label3.Text = "";

            textBox4.Text = GetMenu.DFirst();

            textBox2.Text = GetMenu.DLast();
        }
 
        private void button3_Click_1(object sender, EventArgs e)
        {
             if (tabControl1.SelectedIndex == 0)
            {
                   CalcTotals2();
             }
          
        }
        private void CalcTotals2()
        {


            Int32 iTotal = 0;
            decimal iVatSum = 0;


            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["應收帳款"].Value);
                iVatSum += Convert.ToDecimal(dataGridView1.SelectedRows[iRecs].Cells["美金應收帳款"].Value);

            }


            textBox1.Text = iTotal.ToString("#,##0");

            textBox3.Text = iVatSum.ToString("#,##0.0000");


        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView1.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
            try
            {
                if (!String.IsNullOrEmpty(dgr.Cells["逾期天數"].Value.ToString()))
                {

                    if (Convert.ToInt32(dgr.Cells["逾期天數"].Value.ToString()) >= 0)
                    {

                        dgr.DefaultCellStyle.BackColor = Color.Pink;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GetCountry(string cardcode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   select distinct cardcode 客戶編號,u_territory 國家 from crd1 where u_territory is not null and cardcode=@cardcode ");
        

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetORCT4(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT Convert(varchar(10),T0.DOCDATE,111) DOCDATE   FROM  [dbo].[ORCT] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N' ");
            sb.Append(" AND	invtype='13'  AND T1.DOCENTRY =@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public  void GETGRIDVIEW(DataGridView dgv)
        {
            string 單號;
            string 總類;
            decimal 台幣金額;

            string 美金單價;
            string 文件類型;
            string 客戶代碼;
            string usd;
            string usdf;
            int MA = 0;
            DateTime 逾期日期;
            System.Data.DataTable dt = GetOrderDataAP();


            System.Data.DataTable dtCost = MakeTableCombine();

            System.Data.DataTable dt1 = null;

            System.Data.DataTable dt2 = null;
            DataRow dr = null;
            DataRow dr22 = null;
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("查無資料");
                return;
            }
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                decimal MQTY = 0;
                decimal MQTYT = 0;
                decimal MPRICE = 0;
                decimal MAMT = 0;


                單號 = dt.Rows[i]["docentry"].ToString();
                文件類型 = dt.Rows[i]["文件類型"].ToString();
                dt1 = GetOrderDataAP1(單號, 文件類型);

                dr = dtCost.NewRow();
                總類 = dt1.Rows[0]["總類"].ToString();
                台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                美金單價 = dt1.Rows[0]["美金單價"].ToString();

                dr["摘要"] = dt1.Rows[0]["摘要"].ToString();
                dr["過帳日期"] = dt1.Rows[0]["過帳日期"].ToString();
                dr["應收帳款"] = dt.Rows[i]["應收帳款"].ToString();
                dr["應收帳款2"] = dt.Rows[i]["應收帳款2"].ToString();
                dr["客戶名稱"] = dt1.Rows[0]["客戶名稱"].ToString();
                dr["收款條件"] = dt1.Rows[0]["收款條件"].ToString();
                dr["AR單號"] = 單號;
                客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                dr["客戶代碼"] = "'" + 客戶代碼;
                usd = "0";
                usdf = "0";

                dr["業務"] = dt.Rows[i]["業務"].ToString();
                dr["業管"] = dt.Rows[i]["業管"].ToString();
                dr["最終客戶"] = dt1.Rows[0]["最終客戶"].ToString();
                dr["invoice"] = dt.Rows[i]["發票號碼"].ToString();

                dr["逾期日期"] = Convert.ToDateTime(dt1.Rows[0]["逾期日期"]).ToString("yyyyMMdd");
                dr["出口報單類別"] = dt1.Rows[0]["出口報單類別"].ToString();
                dr["出口證明文件號碼"] = dt1.Rows[0]["出口證明文件號碼"].ToString();
                dr["SHIPTO"] = dt1.Rows[0]["SHIPTO"].ToString();

                if (總類 == "AR")
                {
                    sc = Convert.ToDecimal(台幣金額) - Convert.ToDecimal(sk);
                }
                else
                {
                    sc = (Convert.ToDecimal(台幣金額) - Convert.ToDecimal(sk)) * -1;
                }
                sd = 0;
                sdf = 0;
                sd2 = 0;
                for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                {
                    DataRow dd = dt1.Rows[j];
                    string hg = dd["美金單價"].ToString();

                    if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                    {

                        sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);
                        sdf += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]);

                        if (總類 == "AR")
                        {

                            usd = sd.ToString("#,##0.0000");
                            usdf = sdf.ToString("#,##0.0000");
                        }
                        else
                        {
                            usd = (sd * -1).ToString("#,##0.0000");
                            usdf = (sdf * -1).ToString("#,##0.0000");
                        }


                    }

                    dr["美金應收帳款"] = usd;
                    dr["美金應收帳款2"] = usdf;

                    sd2 += Convert.ToDecimal(dd["數量"].ToString());
                    dr["數量加總"] = sd2.ToString();
                    if (單號 == "53099")
                    {
                        MessageBox.Show("A");
                    }
                    if (dt1.Rows.Count == 1)
                    {
                        dr["品名"] = dd["品名"].ToString();
                        dr["數量"] = dd["數量"].ToString();
                        dr["訂單號碼"] = dd["訂單號碼"].ToString();
                        dr["發票總類"] = dd["發票總類"].ToString();
                        if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                        {
                            MA = 2;
                            decimal sr = Convert.ToDecimal(dd["美金單價"]);
                            dr["美金單價"] = sr.ToString("#,##0.00");
                        }
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                        {
                            MQTY = Convert.ToDecimal(dd["數量"]);
                            MPRICE = Convert.ToDecimal(dd["美金單價"]);
                            if (MQTY != 0)
                            {
                                if (MPRICE != 0)
                                {
                                    MA = 1;



                                    MQTYT += MQTY;
                                    MAMT += (MQTY * MPRICE);
                                }
                            }
                        }

                        if (j == dt1.Rows.Count - 1)
                        {
                     
                            dr["品名"] += dd["品名"].ToString();
                            dr["數量"] += dd["數量"].ToString();

                            dr["訂單號碼"] += dd["訂單號碼"].ToString();
                            if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                dr["美金單價"] += sr.ToString("#,##0.00");
                            }
                            dr["發票總類"] = dd["發票總類"].ToString();
                        }
                        else
                        {

                            dr["訂單號碼"] += dd["訂單號碼"].ToString() + "/";
                            dr["發票總類"] = dd["發票總類"].ToString() + "/";
                            dr["品名"] += dd["品名"].ToString() + "/";
                            dr["數量"] += dd["數量"].ToString() + "&";
                            if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                dr["美金單價"] += sr.ToString("#,##0.00") + "/";
                            }
                            ssd = dd["訂單號碼"].ToString();

                        }
                    }

              
                }

                if (MA == 1)
                {
                    if (MQTYT != 0)
                    {
                        decimal SM = (MAMT / MQTYT);
                        dr["美金平均單價"] = SM.ToString("#,##0.00");
                    }
                }
     
                else
                {
                   // string ss = dr["美金單價"].ToString();
                    dr["美金平均單價"] = dr["美金單價"];
                }


                if (!String.IsNullOrEmpty(dr["應收帳款"].ToString()) && !String.IsNullOrEmpty((dr["美金應收帳款"].ToString())))
                {

                    decimal s = Convert.ToDecimal(dr["應收帳款"]);
                    decimal v = Convert.ToDecimal(dr["美金應收帳款"]);

                    try
                    {
                        decimal dsz = Convert.ToDecimal(dr["應收帳款"]) / Convert.ToDecimal(dr["美金應收帳款"]);
                        dr["匯率"] = dsz.ToString("#,##0.0000");
                    }
                    catch
                    {
                        dr["匯率"] = "";
                    }
                }


                dt2 = GetCountry(客戶代碼);
                StringBuilder sb6 = new StringBuilder();
                for (int k = 0; k <= dt2.Rows.Count - 1; k++)
                {
                    DataRow dk = dt2.Rows[k];

                    sb6.Append(dk["國家"].ToString() + "/");
                }
                if (String.IsNullOrEmpty(sb6.ToString()))
                {
                    dr["國家"] = "";
                }
                else
                {
                    sb6.Remove(sb6.Length - 1, 1);
                    dr["國家"] = sb6.ToString();
                }
                StringBuilder sb1 = new StringBuilder();

                System.Data.DataTable GORCT4 = GetORCT4(單號);
                for (int j = 0; j <= GORCT4.Rows.Count - 1; j++)
                {

                    string DOCDATE = GORCT4.Rows[j]["DOCDATE"].ToString();

                    sb1.Append(DOCDATE + "/");


                }

                if (!String.IsNullOrEmpty(sb1.ToString()))
                {
                    sb1.Remove(sb1.Length - 1, 1);
                    dr["SAP1"] = sb1.ToString();
                }
                int overday = GetMenu.DaySpan(Convert.ToDateTime(dt1.Rows[0]["逾期日期"]).ToString("yyyyMMdd"), textBox2.Text);
                dr["逾期天數"] = overday;
                dtCost.Rows.Add(dr);


            }

            if (dtCost.Rows.Count > 0)
            {
                dtCost.DefaultView.Sort = "客戶代碼";
                dgv.DataSource = dtCost;
                string g = dtCost.Compute("Sum(應收帳款)", null).ToString();
                string gk = dtCost.Compute("Sum(美金應收帳款)", null).ToString();

                decimal sh = Convert.ToDecimal(g);
                decimal shk = Convert.ToDecimal(gk);
                label6.Text = "美金合計:" + shk.ToString("#,##0.0000");
                label3.Text = "台幣合計:" + sh.ToString("#,##0");

            }


            System.Data.DataTable dtCost2 = MakeTableCombine2();
            System.Data.DataTable G2 = dtCost;

            for (int B = 0; B <= G2.Rows.Count - 1; B++)
            {


                DataRow dz = G2.Rows[B];
                dr22 = dtCost2.NewRow();
                dr22["過帳日期"] = dz["過帳日期"].ToString();
                dr22["客戶名稱"] = dz["客戶名稱"].ToString();
                dr22["客戶代碼"] = dz["客戶代碼"].ToString();
                dr22["美金應收帳款"] = dz["美金應收帳款"].ToString();
                dr22["發票總類"] = dz["發票總類"].ToString();
                dr22["SHIPTO"] = dz["SHIPTO"].ToString();
                dtCost2.Rows.Add(dr22);
            }

            if (dtCost2.Rows.Count > 0)
            {
                dtCost2.DefaultView.RowFilter = " 發票總類  in ('41100105 - 銷貨收入-境外','41100102 - 銷貨收入-經海關') ";
                string das1 = dtCost.Compute("Sum(美金應收帳款)", "發票總類  = ('41100105 - 銷貨收入-境外')").ToString();
                string das2 = dtCost.Compute("Sum(美金應收帳款)", "發票總類  in ('41100102 - 銷貨收入-經海關','41100105 - 銷貨收入-境外')").ToString();
                dtCost2.DefaultView.Sort = "客戶代碼";
                dataGridView2.DataSource = dtCost2;
                label4.Text = "三角比重 : " + (Convert.ToDecimal(das1) / Convert.ToDecimal(das2)).ToString("#,##0.0000");
            }
        }


       
   

        

      

      







    }
}
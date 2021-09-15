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
    public partial class OPCH : Form
    {
    
        private decimal sd;
        private decimal se;

        public OPCH()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            decimal 台幣金額;
            string 美金單價;            
            string 文件類型;
            string 客戶代碼;
            string 傳票備註;
            string 傳票號碼;
            string 採購單;
            string AP發票;
            string 收貨採購單;
            string INVOICENO;
            string 日期;
            string LC;
            System.Data.DataTable dt = GetOrderDataAP();
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("無資料");
                return;
            }
            else
            {
                System.Data.DataTable dtCost = MakeTableCombine();
                System.Data.DataTable dt1 = null;

                DataRow dr = null;


                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {



                    單號 = dt.Rows[i]["docentry"].ToString();
                    文件類型 = dt.Rows[i]["文件類型"].ToString();
                    dt1 = GetOrderDataAP1(單號, 文件類型);

                    dr = dtCost.NewRow();
                    總類 = dt1.Rows[0]["總類"].ToString();
                    過帳日期 = dt1.Rows[0]["過帳日期"].ToString();
                    台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                    美金單價 = dt1.Rows[0]["美金單價"].ToString();
                    客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                    採購單 = dt1.Rows[0]["採購單"].ToString();
                    AP發票 = dt1.Rows[0]["AP發票"].ToString();
                    收貨採購單 = dt1.Rows[0]["收貨採購單"].ToString();
                    傳票備註 = dt1.Rows[0]["傳票備註"].ToString();
                    傳票號碼 = dt1.Rows[0]["傳票號碼"].ToString();
                    客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                    美金單價 = dt1.Rows[0]["美金單價"].ToString();

                    INVOICENO = dt1.Rows[0]["INVOICENO"].ToString();
                    日期 = dt1.Rows[0]["日期"].ToString();
                    LC = dt1.Rows[0]["LC"].ToString();
                    dr["台幣金額"] = 台幣金額.ToString("#,##0");
                    dr["客戶代碼"] = 客戶代碼;
                    dr["傳票備註"] = 傳票備註;
                    dr["傳票NO"] = 過帳日期 + "-" + 傳票號碼;
                    dr["採購單"] = 採購單;
                    dr["AP發票"] = AP發票;
                    dr["收貨採購單"] = 收貨採購單;
                    dr["INVOICENO"] = INVOICENO;
                    dr["日期"] = 日期;
                    dr["LC"] = LC;
                    sd = 0;
                    decimal sh = 0;
                    decimal sv = 0;
                    for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                    {
                        DataRow dd = dt1.Rows[j];
                        string sa = dd["數量"].ToString();
                        string sds = dd["美金單價"].ToString();
                        string sae = dd["稅率"].ToString();

                        if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                        {

                            sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);
                            sh += Convert.ToDecimal(dd["數量"]);
                            sv += Convert.ToDecimal(dd["美金單價"]);
                            se = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]) / sd;

                            if (總類 == "AR")
                            {
                                dr["美金金額"] = sd.ToString("#,##0.00");

                            }
                            else
                            {
                                dr["美金金額"] = (sd * -1).ToString("#,##0.00");

                            }
                            dr["總數量"] = sh.ToString("#,##0");
                            dr["匯率"] = se.ToString("#,##0.000");
                            dr["美金單價2"] = sv.ToString("#,##0");

                        }
                        else
                        {
                            sd = 0;
                            dr["美金金額"] = 0;

                            dr["匯率"] = 0;
                        }





                    }


                    dtCost.Rows.Add(dr);


                }

                if (dtCost.Rows.Count > 0)
                {
                    dtCost.DefaultView.Sort = "客戶代碼";
                    bindingSource1.DataSource = dtCost;
                    dataGridView1.DataSource = bindingSource1.DataSource;
                    string g = dtCost.Compute("Sum(台幣金額)", null).ToString();
                    string gk = dtCost.Compute("Sum(美金金額)", null).ToString();

                    decimal sh = Convert.ToDecimal(g);
                    decimal shk = Convert.ToDecimal(gk);
                    label3.Text = "台幣合計"+sh.ToString("#,##0");
                    label6.Text = "美金合計"+shk.ToString("#,##0.0000");
                }
            }
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("傳票NO", typeof(string));
            dt.Columns.Add("AP發票", typeof(string));
            dt.Columns.Add("採購單", typeof(string));
            dt.Columns.Add("收貨採購單", typeof(string));

            dt.Columns.Add("傳票備註", typeof(string));
            dt.Columns.Add("總數量", typeof(string));
            dt.Columns.Add("美金單價2", typeof(string));
            dt.Columns.Add("台幣金額", typeof(Decimal));
            dt.Columns.Add("美金金額", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("INVOICENO", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("LC", typeof(string));

            return dt;
        }
        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'AR' 總類,t0.docentry,t0.objtype 文件類型  from OPCH t0");
            sb.Append(" left join ORPC t1 on(cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar))");
            sb.Append(" LEFT JOIN dbo.ocrd T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append(" left JOIN dbo.VPM2 T2 ON (T2.DOCENTRY=T0.DOCENTRY) ");
            sb.Append(" where 1=1 ");
         
              sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
            
        
                sb.Append(" and year(t0.docdate)='" + comboBox2.SelectedValue.ToString() + "'  ");
            
            if (comboBox3.SelectedValue.ToString() != "ALL")
            {
                sb.Append(" and month(t0.docdate)='" + comboBox3.SelectedValue.ToString() + "'  ");
            }
            sb.Append(" union all");
            sb.Append(" select 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型  from ORPC t0");
            sb.Append(" left join OPCH t1 on(cast(t1.docentry as varchar)=cast(t0.u_acme_arap as varchar))");
            sb.Append(" LEFT JOIN dbo.ocrd  T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append(" left JOIN dbo.VPM2 T2 ON (T2.DOCENTRY=T0.DOCENTRY) ");
            sb.Append(" where  t0.doctype <> 'S'    ");
           
                sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
            
          
                sb.Append(" and year(t0.docdate)='" + comboBox2.SelectedValue.ToString() + "'  ");

                if (comboBox3.SelectedValue.ToString() != "ALL")
            {
                sb.Append(" and month(t0.docdate)='" + comboBox3.SelectedValue.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

  

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
        private System.Data.DataTable GetOrderDataAP1(string aa,string bb)
        {
            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CAST(T8.PRICE AS VARCHAR) 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,substring(Convert(varchar(10),T0.u_acme_invoice,111),6,6) 日期,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,T7.docentry 收貨採購單,t0.docentry AP發票,T8.docentry 採購單 FROM OPCH T0  ");
            sb.Append("              LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN POR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              where t1.basetype='20'");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("              union all");
            sb.Append("              SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CAST(T8.PRICE AS VARCHAR)  美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,substring(Convert(varchar(10),T0.u_acme_invoice,111),6,6) 日期,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,T8.docentry 採購單 FROM OPCH T0  ");
            sb.Append("              LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN POR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              where t1.basetype='22'");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                           SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                           T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                         ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                         T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,substring(Convert(varchar(10),T0.u_acme_invoice,111),6,6) 日期,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,'' 採購單 FROM OPCH T0  ");
            sb.Append("                         LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t1.basetype =-1 and  t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT)*-1 台幣金額 ");
            sb.Append("                           ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                           T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,substring(Convert(varchar(10),T0.u_acme_invoice,111),6,6) 日期,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,'' 採購單 FROM ORPC T0  ");
            sb.Append("                           LEFT JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t0.docentry=@docentry and t0.objtype=@bb");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderDataAP2(string cc,string dd)
        {
            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT isnull(T0.U_USD,0) 金額,T0.DOCENTRY 來源,T0.DOCNUM 單號 FROM RCT2  T0");
            sb.Append(" inner join orct t1 on (t0.docnum=t1.docnum) ");
            sb.Append(" WHERE  t1.canceled <> 'Y' AND T0.DOCENTRY=@cc and t0.invtype=@dd ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@dd", dd));
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


        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void CheckPaid_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.MoneyBU("checkmoney"), "DataTEXT", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, Getship(), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, Getmonth(), "DataText", "DataValue");
            label6.Text = "";
            label3.Text = "";
        }
        System.Data.DataTable Getship()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='shipyear' order by Datatext desc ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }
        System.Data.DataTable Getmonth()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='shipmonth' union all SELECT TOP 1 'ALL' AS DataValue,'ALL' as DataText FROM RMA_PARAMS ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }
     

    }
}
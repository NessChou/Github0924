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
    public partial class CheckMoney3 : Form
    {
        System.Data.DataTable H1 = null;
        System.Data.DataTable H2 = null;
        private decimal sd;
        private decimal se;
        private decimal sc;
        private decimal sdk;
        private decimal USD;
        public CheckMoney3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataRow dr22 = null;
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            string 客戶名稱;
            decimal 台幣金額;
            decimal 稅額;
            decimal 未稅金額;
            decimal 未支付;
            string 美金單價;
            string 收款條件;
            string 文件類型;
            string 客戶代碼;
            string 應收總計;
            string 傳票備註;
            string 傳票號碼;
            string 採購單;
            string 日期;
            string INVOICENO;
            string LC;
            string 收貨採購單;
            string AP發票;
            string fh = "";
            string 到期日期;
            string 原始幣別;
            string 立帳匯率;
            string 對應科目;
            decimal RATE = 0;
            if (textBox1.Text == "")
            {
                fh = "1";
            }
            else
            {
                fh = "2";
            }


            System.Data.DataTable dt = GetOrderDataAP(fh);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
   
            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
            DataRow dr = null;
          

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {



                單號 = dt.Rows[i]["docentry"].ToString();
                文件類型 = dt.Rows[i]["文件類型"].ToString();
                dt1 = GetOrderDataAP1(單號,文件類型);
             
                dr = dtCost.NewRow();
                總類 = dt1.Rows[0]["總類"].ToString();
                過帳日期 = dt1.Rows[0]["過帳日期"].ToString();
                台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                稅額 = Convert.ToDecimal(dt1.Rows[0]["稅額"]);
                未稅金額 = Convert.ToDecimal(dt1.Rows[0]["未稅金額"]);
                未支付 = Convert.ToDecimal(dt1.Rows[0]["未支付"]);
                採購單= dt1.Rows[0]["採購單"].ToString();
                收款條件 = dt1.Rows[0]["收款條件"].ToString();
                美金單價 = dt1.Rows[0]["美金單價"].ToString();
                客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                客戶名稱 = dt1.Rows[0]["客戶名稱"].ToString();
                應收總計 = dt1.Rows[0]["應收總計"].ToString();
                傳票備註 = dt1.Rows[0]["傳票備註"].ToString();
                傳票號碼 = dt1.Rows[0]["傳票號碼"].ToString();
                到期日期 = dt1.Rows[0]["到期日期"].ToString();
                LC = dt1.Rows[0]["LC"].ToString();
                客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                美金單價 = dt1.Rows[0]["美金單價"].ToString();
                日期 = dt1.Rows[0]["日期"].ToString();
                INVOICENO = dt1.Rows[0]["INVOICENO"].ToString();
                AP發票 = dt1.Rows[0]["AP發票"].ToString();
                收貨採購單 = dt1.Rows[0]["收貨採購單"].ToString();
                原始幣別 = dt1.Rows[0]["原始幣別"].ToString();
                立帳匯率 = dt1.Rows[0]["立帳匯率"].ToString();
                收款條件 = dt1.Rows[0]["收款條件"].ToString();
                對應科目 = dt1.Rows[0]["ACCTCODE"].ToString();
                dr["AP發票"] = AP發票;
                dr["採購單"] = 採購單;
                dr["收貨採購單"] = 收貨採購單;

                System.Data.DataTable G4 = GetF4(對應科目);
                if (G4.Rows.Count > 0)
                {

                    dr["對應科目"] = G4.Rows[0][0].ToString();

                }
                string usd = "";
                dr["客戶代碼"] = 客戶代碼;
                dr["客戶名稱"] = 客戶名稱;
                dr["傳票備註"] = 傳票備註;
                dr["傳票NO"] = 過帳日期 + "-" + 傳票號碼;
                dr["日期"] = 日期;
                dr["INVOICENO"] = INVOICENO;
                string INVNO = dt1.Rows[0]["發票號碼"].ToString();
                if (INVNO == "__________"|| String.IsNullOrEmpty(INVNO))
                {
                    System.Data.DataTable G1 = GetF1(傳票號碼);
                    if (G1.Rows.Count > 0)
                    {

                        dr["發票號碼"] = G1.Rows[0][0].ToString();

                    }
                }
                else
                {

                    dr["發票號碼"] = INVNO;
                }


                //__________
                dr["發票日期"] = dt1.Rows[0]["發票日期"].ToString();
                dr["項目成本差異原因"] = dt1.Rows[0]["項目成本差異原因"].ToString();
                dr["LC"] = LC;
                dr["到期日期"] = 到期日期;
                dr["原始幣別"] = 原始幣別;
                dr["立帳匯率"] = 立帳匯率;
                H1 = GetCMONEY(收貨採購單);
                H2 = GetCMONEY2(單號);
                if (總類=="AR")
                {
                    sc = Convert.ToDecimal(台幣金額) ;

                }
                else 

                {
                    sc = (Convert.ToDecimal(台幣金額)) * -1;

                }
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
                        try
                        {
                            sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);
                        }
                        catch 
                        {
                            sd += Convert.ToDecimal(dd["數量"]) * 53 * Convert.ToDecimal(dd["稅率"]);
                        }
                          
                           sh += Convert.ToDecimal(dd["數量"]);
                           
                        try
                        {
                            sv = Convert.ToDecimal(dd["美金單價"]);
                        }
                        catch
                        {
                            sv = Convert.ToDecimal(53);
                        }

                        Double x = Convert.ToDouble(sd);
                           sd = Convert.ToDecimal(C1Round(x, 2).ToString());
                           if (sd != 0)
                           {
                               se = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]) / sd;
                           }

                           if (總類 == "AR")
                           {
                               usd = sd.ToString();
                           }
                           else
                           {
                               usd = (sd * -1).ToString();
                           }

                           dr["總數量"] = sh.ToString("#,##0");
                           dr["美金單價2"] = sv.ToString("#,##0");
                           dr["匯率"] = se.ToString("#,##0.000");

                       if (H1.Rows.Count > 0)
                       {
                            RATE = Convert.ToDecimal(H1.Rows[0]["RATE"]);
                           dr["匯率"] = RATE.ToString("#,##0.000");
                       }

                   }
                   else
                   {
                       sd = 0;
                       usd = "0";
                       dr["匯率"] = 0;
                   }
       
                  
                    dr["應收帳款2"] = sc.ToString("#,##0");
                    dr["稅額"] = 稅額.ToString("#,##0");
                    dr["未稅金額"] = 未稅金額.ToString("#,##0");
                    dr["美金應收帳款2"] = Convert.ToDecimal(usd).ToString("#,##0.00");


                    if (H1.Rows.Count > 0)
                    {
                        decimal USD2 = Convert.ToDecimal(H1.Rows[0]["USD"]);
                        dr["美金應收帳款2"] = Convert.ToDecimal(USD2).ToString("#,##0.00");
                    }

                    if (H2.Rows.Count > 0)
                    {
                        decimal USD2 = Convert.ToDecimal(H2.Rows[0]["AMOUNT"]);
                        dr["美金應收帳款2"] = Convert.ToDecimal(USD2).ToString("#,##0.00");
                    }
                    sdk = 0;
                    USD = 0;
                    dt2 = GetOrderDataAP2(單號, 文件類型);
                    for (int k = 0; k <= dt2.Rows.Count - 1; k++)
                    {
                        dr["付款日期"] = dt2.Rows[0]["DOCDATE"].ToString();
                        dr["付款傳票號碼"] = dt2.Rows[0]["TRANSID"].ToString();

                        DataRow ddk = dt2.Rows[k];
                        if ((!String.IsNullOrEmpty(ddk["金額"].ToString())))
                        {
                            sdk += Convert.ToDecimal(ddk["金額"]);

                            if (總類 == "AR")
                            {
                                USD = sdk;
                            }
                            else
                            {
                                USD = (sdk * -1);
                            }
                     
                        }
                        else
                        {
                            sdk = 0;
                            USD = 0;
  
                        }

                        decimal da = Convert.ToDecimal(usd) - USD;
                            dr["美金應收帳款2"] = da.ToString();

                    }


                    string F1 = dr["匯率"].ToString();

                    if (F1 == "1.000")
                    {
                        dr["美金應收帳款2"] = "0";
                    }
                    dr["餘額TWD"] = 未支付;
                    if (Convert.ToDecimal(dr["匯率"]) != 0)
                    {
                        dr["餘額USD"] = (未支付 / Convert.ToDecimal(dr["匯率"])).ToString("#,##0.00");
                    }
                }
  
                   
                dtCost.Rows.Add(dr);


            }
      
            if (dtCost.Rows.Count > 0)
            {
                dtCost.DefaultView.Sort = "客戶代碼";
       
                dataGridView1.DataSource = dtCost;
                decimal F1 = 0;
                decimal F2 = 0;
                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {

                    DataGridViewRow row;

                    row = dataGridView1.Rows[i];
                    F1+= Convert.ToDecimal(row.Cells["應收帳款2"].Value);
                    F2 += Convert.ToDecimal(row.Cells["美金應收帳款2"].Value);
                }

                label3.Text = "台幣合計:" + F1.ToString("#,##0");
                label6.Text = "美金合計:" + F2.ToString("#,##0.0000");


            }

        }

        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
  

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["客戶代碼"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("傳票NO", typeof(string));
            dt.Columns.Add("AP發票", typeof(string));
            dt.Columns.Add("採購單", typeof(string));
            dt.Columns.Add("收貨採購單", typeof(string));

            dt.Columns.Add("總數量", typeof(string));
            dt.Columns.Add("美金單價2", typeof(string));
        
            dt.Columns.Add("傳票備註", typeof(string));
 

            dt.Columns.Add("未稅金額", typeof(Decimal));
            dt.Columns.Add("稅額", typeof(Decimal));

            dt.Columns.Add("應收帳款2", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("美金應收帳款2", typeof(Decimal));
            dt.Columns.Add("餘額TWD", typeof(Decimal));
            dt.Columns.Add("餘額USD", typeof(Decimal));
            dt.Columns.Add("INVOICENO", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("到期日期", typeof(string));

            dt.Columns.Add("項目成本差異原因", typeof(string));
            dt.Columns.Add("LC", typeof(string));
            dt.Columns.Add("付款日期", typeof(string));
            dt.Columns.Add("付款傳票號碼", typeof(string));
            dt.Columns.Add("對應科目", typeof(string));
            dt.Columns.Add("原始幣別", typeof(string));
            dt.Columns.Add("立帳匯率", typeof(string));
            return dt;
        }
        private System.Data.DataTable GetOrderDataAP(string jh)
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型");
            sb.Append("             from OPCH t0");
            sb.Append("             left join ORPC t1 on(cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate4   )");
            sb.Append("             LEFT JOIN ocrd T9 ON (T0.cardcode=T9.cardcode)    ");
            sb.Append("             left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("             [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4 GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='18' )");
            sb.Append("             left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied,MAX(T0.DOCDATE) DD FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("             [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4  GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='18' )");
            sb.Append("             LEFT JOIN (SELECT DISTINCT T1.BASETYPE,T1.BASEENTRY,T0.DOCTOTAL FROM ORPC T0");
            sb.Append("             LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("             WHERE BASETYPE=18 and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4 ) T6 ON (T0.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              where  t0.transid not in (select je from ACMESQLSP.DBO.account_je)   ");
            if (textBox2.Text != "" && textBox3.Text != "")
            {
                FD = textBox3.Text;
                sb.Append(" AND   Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2    ");
            }
            else
            {
                FD =DateTime.Now.AddYears(1).ToString("yyyyMMdd");
           
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) <>  0 ");
                sb.Append("  and ((isnull(t0.doctotal,0)-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0)) - isnull(t1.doctotal,0)) <> 0   ");
            }
            if (jh == "1")
            {
                if (comboBox1.SelectedValue.ToString() == "NAUO")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='S'  AND SUBSTRING(T0.CARDCODE,1,5) <> 'S0001'   ");

                }
                else if (comboBox1.SelectedValue.ToString() == "UCARD")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='U'     ");

                }
                else
                {
                    sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
                }
            }
            else
            {
                sb.Append(" and T0.CARDCODE in (" + textBox1.Text.ToString() + ")  ");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" AND  Convert(varchar(8),T5.DD,112)   =@DocDate3 ");
            }
       
            sb.Append(" union all");
            sb.Append("               select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型 from ORPC t0");
            sb.Append("               left join OPCH t1 on(cast(t1.docentry as varchar)=cast(t0.u_acme_arap as varchar) and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate4 ) ");
            sb.Append("               LEFT JOIN ocrd  T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("               [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4    GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='19'  )");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied,MAX(T0.DOCDATE) DD FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("               [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4 GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='19'  )");
            sb.Append("               left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0  ");
            sb.Append("               WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4   ) t6 on (t0.docentry=t6.docentry)");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("               [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4    GROUP BY T1.[DocEntry],invtype )  t7 on (cast(t0.u_acme_arap as varchar)=cast(t7.docentry as varchar)  and t7.invtype='18'  )");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("               [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4  GROUP BY T1.[DocEntry],invtype  ) t8 on (cast(t0.u_acme_arap as varchar)=cast(t8.docentry as varchar) and t8.invtype='18'  )");
            sb.Append("   where   t0.transid not in (select je from ACMESQLSP.DBO.account_je)  ");
            if (textBox2.Text != "" && textBox3.Text != "")
            {
                sb.Append(" AND   Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2    ");
            }
        
            if (checkBox1.Checked)
            {
                sb.Append("   and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0) <>  0    and ((isnull(t0.doctotal,0)-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0)) - (isnull(t1.doctotal,0)-isnull(t7.sumapplied,0)-isnull(t8.sumapplied,0))) <> 0  AND (CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT)) <> 0  ");
                sb.Append("  AND T0.DOCENTRY NOT IN (select DISTINCT T0.DOCENTRY from RPC1 T0");
                sb.Append("   LEFT JOIN PCH1 T1 ON (T0.BASEENTRY=T1.DOCENTRY AND T0.LINENUM=T1.BASELINE ) WHERE T0.BASETYPE=18 AND Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate4) ");
            }
            if (jh == "1")
            {
                if (comboBox1.SelectedValue.ToString() == "NAUO")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='S'  AND SUBSTRING(T0.CARDCODE,1,5) <> 'S0001'   ");

                }
                else if (comboBox1.SelectedValue.ToString() == "UCARD")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='U'     ");

                }
                else
                {
                    sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
                }
            }
            else
            {
                sb.Append(" and T0.CARDCODE in (" + textBox1.Text.ToString() + ")  ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" AND  Convert(varchar(8),T5.DD,112)   =@DocDate3 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
    
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DocDate3", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@DocDate4", FD));
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

        private System.Data.DataTable GetF1(string TRANSID)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("          select U_PC_BSINV   from [@CADMEN_FMD] T0 LEFT JOIN  [@CADMEN_FMD1] T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE U_BSREN =@U_BSREN");
 
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_BSREN", TRANSID));
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
        private System.Data.DataTable GetF4(string ACCTCODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT ACCTNAME   FROM OACT WHERE ACCTCODE=@ACCTCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ACCTCODE", ACCTCODE));
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
        private System.Data.DataTable GetOrderDataAP1(string aa,string bb)
        {
            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR) 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,T7.docentry 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111)  發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,CASE U_ACME_SHIPMENT WHEN '1' THEN Convert(varchar(10),T0.DOCDUEDATE,111) ELSE '' END 到期日期,T0.U_ACME_RATE1 立帳匯率,T8.CURRENCY 原始幣別,CAST(T0.VATSUM AS INT) 稅額,CAST(T0.doctotal-T0.VATSUM  AS INT) 未稅金額,T1.ACCTCODE  FROM OPCH T0  ");
            sb.Append("              LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN POR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              where t1.basetype='20' and  T0.CARDCODE LIKE '%S0001%' ");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("              union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR) 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,T7.docentry 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111)  發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因, Convert(varchar(10),T0.DOCDUEDATE,111)  到期日期,T0.U_ACME_RATE1 DOCCUR,T8.CURRENCY,CAST(T0.VATSUM AS INT)  稅額,CAST(T0.doctotal-T0.VATSUM  AS INT) 未稅金額,T1.ACCTCODE   FROM OPCH T0  ");
            sb.Append("              LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN POR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              where t1.basetype='20' and  T0.CARDCODE NOT LIKE '%S0001%' ");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("              union all");
            sb.Append("              SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR)  美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111) 發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_ACME_RATE1 DOCCUR,T8.CURRENCY,CAST(T0.VATSUM AS INT)  稅額,CAST(T0.doctotal-T0.VATSUM  AS INT) 未稅金額,T1.ACCTCODE   FROM OPCH T0  ");
            sb.Append("              LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN POR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              where t1.basetype='22'");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                           SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                           T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                         ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,");
            sb.Append("                         T0.COMMENTS 備註,t0.u_acme_pay 收款條件,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,'' 採購單,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111) 發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_ACME_RATE1 DOCCUR,T1.CURRENCY,CAST(T0.VATSUM AS INT)  稅額,CAST(T0.doctotal-T0.VATSUM  AS INT) 未稅金額,T1.ACCTCODE   FROM OPCH T0  ");
            sb.Append("                         LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t1.basetype =-1 and  t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                           ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,");
            sb.Append("                           T0.COMMENTS 備註,t0.u_acme_pay 收款條件,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111)日期 ,'' 採購單,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111) 發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_ACME_RATE1 DOCCUR,T1.CURRENCY,CAST(T0.VATSUM AS INT)*-1  稅額,CAST(T0.doctotal-T0.VATSUM  AS INT)*-1 未稅金額,T1.ACCTCODE   FROM ORPC T0  ");
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
            sb.Append(" SELECT isnull(T0.U_USD,0) 金額,T0.DOCENTRY 來源,T0.DOCNUM 單號,Convert(varchar(10),T1.[DocDate],111) DOCDATE,T1.TRANSID  FROM VPM2  T0");
            sb.Append(" inner join OVPM t1 on (t0.docnum=t1.docnum) ");
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

        private System.Data.DataTable GetCMONEY(string MEMO)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PARAM_NO USD,PARAM_DESC RATE FROM PARAMS WHERE PARAM_KIND='CMONEY' AND MEMO=@MEMO");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
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

        private System.Data.DataTable GetCMONEY2(string DOCENTRY)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT AMOUNT FROM Account_CHECKMONEY WHERE DOCENTRY=@DOCENTRY ");


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
        private void button2_Click(object sender, EventArgs e)
        {

            ExcelReport.GridViewToExcel(dataGridView1);
          
        }

        private void CheckPaid_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.MoneyBU("checkmoney3"), "DataTEXT", "DataValue");
            textBox2.Text = GetMenu.DFirst();
            textBox3.Text = GetMenu.DLast();
            label6.Text = "";
            label3.Text = "";
            
        }
        public double C1Round(double value, int digit)
        {
            double vt = Math.Pow(10, digit);
            double vx = value * vt;

            vx += 0.5;
            return (Math.Floor(vx) / vt);
        }
        private void button3_Click(object sender, EventArgs e)
        {
          
            string cs;
            APS3 frm1 = new APS3();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                cs = frm1.q;
                if (!String.IsNullOrEmpty(cs))
                {

                    textBox1.Text = cs;


                }
            }
        }

 







        
    }
}
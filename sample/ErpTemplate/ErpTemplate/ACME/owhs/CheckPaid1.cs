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
    public partial class CheckPaid1 : Form
    {
        int JG = 1 ;
        string ssd;
        private decimal sd,sd2,sd3;
        private decimal se;
        private decimal sc;
        private decimal sdk;
        private decimal sk;
        Int32 iTotal = 0;
        decimal iVatSum = 0;

        public CheckPaid1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
                

                if (artextBox12.Text == "" )
                {
                    MessageBox.Show("請輸入客戶");
                    return;
                }

            
                string 單號;
                string 客戶;
                string 總類;
                string 過帳日期;
                string 工作天數;
                string 發票號碼;
                decimal 台幣金額;

                string 美金單價;
                string 收款條件;
                string 發票金額;
                string 文件類型;
                string 客戶代碼;
                string 應收帳款;
                string 業務;
                string 業管;
                string 摘要;
                string 憑證類別;
                string 通關方式;
                string 外銷方式;
                string 最終客戶;
                string usd;
                string payusd;
                string sdf="";
           
                DateTime 逾期日期;
                System.Data.DataTable dt = GetOrderDataAP();
                System.Data.DataTable dtt = GetOrderDataAP3();

                System.Data.DataTable dtCost = MakeTableCombine();
         
                System.Data.DataTable dt1 = null;
                System.Data.DataTable dt2 = null;
                System.Data.DataTable dt3 = null;
                System.Data.DataTable dt4 = null;
  
        
                DataRow dr = null;




                if (dt.Rows.Count > 0)
                {



                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {



                        單號 = dt.Rows[i]["docentry"].ToString();
                        文件類型 = dt.Rows[i]["文件類型"].ToString();
                        dt1 = GetOrderDataAP1(單號, 文件類型);

                        dr = dtCost.NewRow();
                        總類 = dt1.Rows[0]["總類"].ToString();
                        過帳日期 = dt1.Rows[0]["過帳日期"].ToString();
                        工作天數 = dt1.Rows[0]["工作天數"].ToString();
                        發票號碼 = dt1.Rows[0]["發票號碼"].ToString();
                        逾期日期 = Convert.ToDateTime(dt1.Rows[0]["逾期日期"]);
                        台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                        收款條件 = dt1.Rows[0]["收款條件"].ToString();
                        美金單價 = dt1.Rows[0]["美金單價"].ToString();
                        發票金額 = dt1.Rows[0]["發票金額"].ToString();
                        客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                        應收帳款 = dt.Rows[i]["應收帳款"].ToString();
                        憑證類別 = dt.Rows[i]["憑證類別"].ToString();
                        通關方式 = dt.Rows[i]["通關方式"].ToString();
                        外銷方式 = dt.Rows[i]["外銷方式"].ToString();
                        發票號碼 = dt.Rows[i]["發票號碼"].ToString();
                        客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                        業務 = dt.Rows[i]["業務"].ToString();
                        摘要 = dt1.Rows[0]["摘要"].ToString();
                        業管 = dt.Rows[i]["業管"].ToString();
                        最終客戶 = dt1.Rows[0]["最終客戶"].ToString();

                        dr["摘要"] = 摘要;
                        dr["過帳日期"] = 過帳日期;
                        dr["應收帳款"] = 應收帳款;
                        dr["客戶名稱"] = 客戶;
                        dr["收款條件"] = 收款條件;
                        dr["客戶代碼"] = 客戶代碼;
                        usd = "0";
                        payusd = "0";
                        dr["業務"] = 業務;
                        dr["業管"] = 業管;
                        dr["最終客戶"] = 最終客戶;
                        dr["AR單號"] = 單號;

                        dr["逾期日期"] = 逾期日期.ToString("yyyyMMdd");




                        dt4 = GetWHNO();
                        for (int j = 0; j <= dt4.Rows.Count - 1; j++)
                        {
                            DataRow dd = dt4.Rows[j];
                            sdf = dd["單號"].ToString();


                            if (sdf == 單號)
                            {
                                dr["WHNO"] = "1234";

                            }


                        }
                        if (總類 == "AR")
                        {
                            sc = Convert.ToDecimal(台幣金額) - Convert.ToDecimal(sk);
                        }
                        else
                        {
                            sc = (Convert.ToDecimal(台幣金額) - Convert.ToDecimal(sk)) * -1;
                        }
                        sd = 0;
                        for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                        {
                            DataRow dd = dt1.Rows[j];
                            string hg = dd["美金單價"].ToString();

                            if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                            {

                                sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);


                                if (總類 == "AR")
                                {

                                    usd = sd.ToString("#,##0.0000");
                                }
                                else
                                {
                                    usd = (sd * -1).ToString("#,##0.0000");
                                }


                            }

                            dr["美金應收帳款"] = usd;



                            if (dt1.Rows.Count == 1)
                            {
                                dr["品名"] = dd["品名"].ToString();
                                dr["數量"] = dd["數量"].ToString();
                                dr["訂單號碼"] = dd["訂單號碼"].ToString();
                                if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                                {
                                    decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                    dr["美金單價"] = sr.ToString("#,##0.0000");
                                }
                            }
                            else
                            {

                                if (j == dt1.Rows.Count - 1)
                                {
                                    dr["品名"] += dd["品名"].ToString();
                                    dr["數量"] += dd["數量"].ToString();
                                    dr["訂單號碼"] += dd["訂單號碼"].ToString();
                                    if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                                    {
                                        decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                        dr["美金單價"] += sr.ToString("#,##0.0000");
                                    }
                                }
                                else
                                {

                                    dr["訂單號碼"] += dd["訂單號碼"].ToString() + "/";

                                    dr["品名"] += dd["品名"].ToString() + "/";
                                    dr["數量"] += dd["數量"].ToString() + "&";
                                    if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                                    {
                                        decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                        dr["美金單價"] += sr.ToString("#,##0.0000") + "/";
                                    }
                                    ssd = dd["訂單號碼"].ToString();

                                }
                            }
                            sdk = 0;
                            dt2 = GetOrderDataAP2(單號, 文件類型);

                            for (int k = 0; k <= dt2.Rows.Count - 1; k++)
                            {
                                DataRow ddk = dt2.Rows[k];
                                if ((!String.IsNullOrEmpty(ddk["金額"].ToString())))
                                {
                                    sdk += Convert.ToDecimal(ddk["金額"]);

                                    if (總類 == "AR")
                                    {
                                        payusd = sdk.ToString("#,##0.0000");
                                    }
                                    else
                                    {
                                        payusd = (sdk * -1).ToString("#,##0.0000");
                                    }

                                }
                                else
                                {
                                    sdk = 0;
                                    payusd = "0";

                                }



                                dr["美金應收帳款"] = (Convert.ToDecimal(usd) - Convert.ToDecimal(payusd)).ToString();

                            }

                        }
                        dt3 = GetOrderinv(單號, 文件類型);

                        if (通關方式 == "0" || (通關方式 == "1" && 外銷方式 == "1"))
                        {
                            dr["發票總類"] = "國外";
                            for (int p = 0; p <= dt3.Rows.Count - 1; p++)
                            {
                                DataRow ddp = dt3.Rows[p];


                                if (dt3.Rows.Count == 1)
                                {
                                    dr["invoice"] = ddp["invoice"].ToString();

                                }
                                else
                                {

                                    if (p == dt3.Rows.Count - 1)
                                    {
                                        dr["invoice"] += ddp["invoice"].ToString();

                                    }
                                    else
                                    {

                                        dr["invoice"] += ddp["invoice"].ToString() + "/";


                                    }
                                }
                            }
                        }
                        else if ((通關方式 == "1" && 外銷方式 == "0") || (通關方式 == "1" && 外銷方式 == "4"))
                        {
                            dr["發票總類"] = "國內";
                            dr["invoice"] = 發票號碼;

                            if (收款條件 == "LC at sight")
                            {


                                逾期日期 = 逾期日期.AddDays(-13);
                                dr["逾期日期"] = 逾期日期.ToString("yyyyMMdd");
                            }
                        }
                        else if (憑證類別 == "5")
                        {
                            dr["發票總類"] = "免用";
                        }
                        TimeSpan ts = DateTime.Today - 逾期日期;

                        dr["逾期天數"] = ts.TotalDays;

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


                        dtCost.Rows.Add(dr);



                    }
                    for (int m = 0; m <= dtt.Rows.Count - 1; m++)
                    {
                        dr = dtCost.NewRow();
                        DataRow dy = dtt.Rows[m];

                        dr["客戶代碼"] = dy["客戶代碼"].ToString();
                        dr["客戶名稱"] = dy["客戶名稱"].ToString();
                        dr["應收帳款"] = dy["台幣金額"].ToString();
                        dr["美金應收帳款"] = dy["美金金額"].ToString();
                        dr["摘要"] = dy["摘要"].ToString();
                        dr["業務"] = dy["業務"].ToString();
                        dr["匯率"] = dy["匯率"].ToString();
                        dtCost.Rows.Add(dr);

                    }

                    if (dtCost.Rows.Count > 0)
                    {

                        dtCost.DefaultView.Sort = "客戶代碼";


                        dataGridView1.DataSource = dtCost;

                    }
                }
                else
                {

                    MessageBox.Show("應收帳款沒有資料");
                }
                    CalcTotals3();
                    CalcTotals2();

                    CalLC();
                    CalLC2();
                


   
            //}
            //catch (Exception ex1)
            //{

            //    MessageBox.Show(ex1.Message);

            //}

        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("收款條件", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("業管", typeof(string));
            dt.Columns.Add("訂單號碼", typeof(string));
            dt.Columns.Add("發票總類", typeof(string));
            dt.Columns.Add("invoice", typeof(string));
            dt.Columns.Add("最終客戶", typeof(string));
            dt.Columns.Add("逾期日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(string));
            dt.Columns.Add("WHNO", typeof(string));
  
            return dt;
        }

 
        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) AS INT) 應收帳款,T0.U_Delivery_date  預收日期,SUBSTRING(GROUPNAME,4,10) 群組  from oinv t0");
            sb.Append(" left join orin t1 on(cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2   )");
            sb.Append(" LEFT JOIN ocrd T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" LEFT JOIN OCRG T10 ON (T9.GROUPCODE=T10.GROUPCODE) ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append(" [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='13' )");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append(" [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='13' )");
            sb.Append(" LEFT JOIN (SELECT DISTINCT T1.BASETYPE,T1.BASEENTRY,T0.DOCTOTAL FROM ORIN T0");
            sb.Append(" LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE BASETYPE=13 and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 ) T6 ON (T0.DOCENTRY=T6.BASEENTRY)");
            sb.Append("  where  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) <>  0  and ((isnull(t0.doctotal,0)-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0)) - isnull(t1.doctotal,0)) <> 0  and t0.docentry not in (SELECT PARAM_NO FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='SAPAR'  AND PARAM_DESC  between '20071231' and @DocDate2) AND T0.CARDCODE <> 'R0001'  ");
            if (checkBox1.Checked)
            {
                sb.Append(" AND T0.[CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T0.[CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }
            sb.Append(" union all");
            sb.Append("              select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,(CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT))*-1 應收帳款,T0.U_Delivery_date 預收日期,SUBSTRING(GROUPNAME,4,10) 群組  from orin t0");
            sb.Append("              left join oinv t1 on(cast(t1.docentry as varchar)=cast(t0.u_acme_arap as varchar) and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2  ) ");
            sb.Append("              LEFT JOIN ocrd  T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append("              left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("              left JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" LEFT JOIN OCRG T10 ON (T9.GROUPCODE=T10.GROUPCODE) ");
            sb.Append("              left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("              [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='14'  )");
            sb.Append("              left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("              [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='14'  )");
            sb.Append("              left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0  ");
            sb.Append("              WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    ) t6 on (t0.docentry=t6.docentry)");
            sb.Append("              left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("              [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    GROUP BY T1.[DocEntry],invtype )  t7 on (cast(t0.u_acme_arap as varchar)=cast(t7.docentry as varchar)  and t7.invtype='13'  )");
            sb.Append("              left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("              [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype  ) t8 on (cast(t0.u_acme_arap as varchar)=cast(t8.docentry as varchar) and t8.invtype='13'  )");
            sb.Append("              left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0  ");
            sb.Append("              WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2   ) t11 on (cast(t0.u_acme_arap as varchar)=t11.docentry)");
            sb.Append("  where   Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0) <>  0    and ((isnull(t0.doctotal,0)-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0)) - (isnull(t1.doctotal,0)-isnull(t7.sumapplied,0)-isnull(t8.sumapplied,0)-isnull(t11.sumapplied,0))) <> 0  and t0.docentry not in (SELECT PARAM_NO FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='SAPAD'  AND PARAM_DESC  between '20071231' and @DocDate2) ");
            sb.Append(" AND (CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT)) <> 0  AND T0.DOCENTRY NOT IN (select DISTINCT T0.DOCENTRY from RIN1 T0 ");
            sb.Append(" LEFT JOIN INV1 T1 ON (T0.BASEENTRY=T1.DOCENTRY AND T0.LINENUM=T1.BASELINE ) WHERE T0.BASETYPE=13) ");
            if (checkBox1.Checked)
            {
                sb.Append(" AND T0.[CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T0.[CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

        private System.Data.DataTable GetOrderDataAP1(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,case when t0.cardcode ='0511-00' then t9.u_beneficiary when t0.cardcode ='0257-00' then t9.u_beneficiary else T0.[CardName] end  客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append("   dbo.fun_CreditDate(T9.u_acme_pay,T0.CardCode,T0.DocDate) 逾期日期");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='15' ");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,case when t0.cardcode ='0511-00' then t9.u_beneficiary when t0.cardcode ='0257-00' then t9.u_beneficiary else T0.[CardName] end 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("             ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append("   dbo.fun_CreditDate(T9.u_acme_pay,T0.CardCode,T0.DocDate) 逾期日期");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='17'");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,'' 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv   美金單價,T0.JRNLMEMO 應收總計,'' 訂單號碼,'' 最終客戶, ");
            sb.Append("                       t0.docdate 逾期日期");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype =-1  ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                           ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                           T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv 美金單價,T0.JRNLMEMO 摘要,cast(T0.u_acme_arap as varchar) 訂單號碼,'' 最終客戶,t0.docdate 逾期日期  FROM Orin T0  ");
            sb.Append("                           LEFT JOIN rin1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t0.docentry=@docentry and t0.objtype=@bb  ");
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

 
        private System.Data.DataTable GetOrderDataAP2(string cc, string dd)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT isnull(T0.U_USD,0) 金額,T0.DOCENTRY 來源,T0.DOCNUM 單號 FROM RCT2  T0");
            sb.Append(" inner join orct t1 on (t0.docnum=t1.docnum) ");
            sb.Append(" WHERE Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2 and t1.canceled <> 'Y' AND T0.DOCENTRY=@cc and t0.invtype=@dd ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@dd", dd));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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

      

        private System.Data.DataTable Get2DO()
        {
                  SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select owneridnum 額度期限,cast(creditline-balance as int) 額度餘額,cast(creditline as int) 信用額度 from ocrd T0 where isnull(creditline,0) <> 0 ");
            if (checkBox1.Checked)
            {
                sb.Append(" AND T0.[CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T0.[CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
 
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
        private System.Data.DataTable GetOrderDataAP3()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CAST(SUBSTRING(u_remark1,6,12) AS DECIMAL)*-1 美金金額, shortname 客戶代碼,cardname 客戶名稱,linememo 摘要,cast(t0.credit as int)*-1 台幣金額,t4.slpname 業務,cast(cast(t0.credit as int)/cast(SUBSTRING(u_remark1,6,12) as decimal) as decimal(5,2)) 匯率 from jdt1 t0");
            sb.Append(" inner join ocrd t1 on (t0.shortname=t1.cardcode)");
            sb.Append(" INNER join oslp t4 on (t1.slpCODE=t4.slpCODE) ");
            sb.Append(" where SUBSTRING(u_remark1,1,5)='AR-AP'  and Convert(varchar(8),t0.refdate,112)  between '20071231' and @DocDate2 ");



            if (checkBox1.Checked)
            {
                sb.Append(" AND T1.[CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T1.[CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }

          //      sb.Append(" and T1.[cardname] like '%" + artextBox12.Text.ToString() + "%' ");


            
     
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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

        private System.Data.DataTable GetOrderinv(string aa, string bb)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct t0.objtype,t0.docentry 單號,INVOICENO+INVOICENO_SEQ invoice  from acmesql02.dbo.inv1 t0");
            sb.Append(" left join acmesql02.dbo.dln1 t1 on (t0.baseentry=T1.docentry and  t0.baseline=t1.linenum  )");
            sb.Append(" left join acmesql02.dbo.rdr1 t2 on (t1.baseentry=T2.docentry and  t1.baseline=t2.linenum  )");
            sb.Append(" left join  dbo.TRADE_ORDER T3 on (T2.docentry=T3.PROD_NO AND T2.linenum=T3.PI_NO)");
            sb.Append(" left join  DBO.INVOICEM t4 on (t3.SHIPNO=t4.shippingcode)");
            sb.Append(" where INVOICENO is not null  and t0.docentry=@docentry and t0.objtype=@bb");
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

        private System.Data.DataTable GetWHNO()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select distinct case  when  isnull(t0.memo,'') = '' then t5.docentry else t0.memo end 單號  from acmesqlsp.dbo.satt2 t0  ");
            sb.Append("               inner join acmesqlsp.dbo.satt1 t2 on (t0.ttcode=T2.ttcode  )");
            sb.Append("              inner join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum )");
            sb.Append("              inner join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )");
            sb.Append("          where (isnull(t2.whno,'') <> '' or joy ='1')  ");

            if (checkBox1.Checked)
            {
                sb.Append(" AND T0.[CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T0.[CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
               ExcelReport.GridViewToExcel(dataGridView5);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
         
        }

        private void CheckPaid_Load(object sender, EventArgs e)
        {
            label17.Text = "";
            label18.Text = "";
            label19.Text = "";
    
            label6.Text = "";
            label3.Text = "";

            label7.Text = "";
            label4.Text = "";

            label9.Text = "";
            label5.Text = "";
            label11.Text = "";
            label2.Text = "";
       
            textBox2.Text = GetMenu.Day();
        }
        System.Data.DataTable GetOslp()
        {

            SqlConnection con = globals.shipConnection;
            string sql = "select slpname as DataValue  from oslp where memo like '%業務%'  UNION ALL SELECT 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM' ORDER BY DataValue";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["oslp"];
        }

        System.Data.DataTable GetOhem()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "select [lastName]+[firstName] as DataValue  from ohem where jobtitle like '%業助%'  UNION ALL SELECT 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM' ORDER BY DataValue ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ohem");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ohem"];
        }


        private System.Data.DataTable Gettt1()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct T2.TTDATE,Seqno,t1.CardName,Company,WHNO");
            sb.Append(" ,Bank,PAYCHECK,Currency,Amount,CurrencyRate,Fee,");
            sb.Append(" TotalAmount,case when TTTotal2=0 then TTTotal when isnull(TTTotal2,0)=0 then TTTotal else TTTotal2 end TTTotal,Detail,t0.ttcode");
            sb.Append(" ,case Currency when 'USD' THEN TotalAmount END USD ");
            sb.Append(" ,case Currency when 'NTD' THEN TotalAmount END NTD,T0.ID,t0.REMARK 備註  ");
            sb.Append(" from satt1 t0 ");
            sb.Append(" left join satt2 t1 on(t0.ttcode=t1.ttcode and t0.seqno=t1.id)");
            sb.Append(" left join satt t2 on(t0.ttcode=t2.ttcode) ");

            sb.Append(" WHERE  isnull(WHNO,'') = '' ");

            if (checkBox1.Checked)
            {
                sb.Append(" AND T1.[CardName]  like  N'%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T1.[CARDCODE2]  = '" + textBox1.Text.ToString() + "' ");
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
        private System.Data.DataTable Gettt1F()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct T2.TTDATE,Seqno,t1.CardName,Company,WHNO");
            sb.Append(" ,Bank,PAYCHECK,Currency,Amount,CurrencyRate,Fee,");
            sb.Append(" TotalAmount,case when TTTotal2=0 then TTTotal when isnull(TTTotal2,0)=0 then TTTotal else TTTotal2 end TTTotal,Detail,t0.ttcode");
            sb.Append(" ,case Currency when 'USD' THEN TotalAmount END USD ");
            sb.Append(" ,case Currency when 'NTD' THEN TotalAmount END NTD,T0.ID,t0.REMARK 備註,T1.LINENUM,T1.DOCENTRY   ");
            sb.Append(" from satt1 t0 ");
            sb.Append(" left join satt2 t1 on(t0.ttcode=t1.ttcode and t0.seqno=t1.id)");
            sb.Append(" left join satt t2 on(t0.ttcode=t2.ttcode) ");

            sb.Append(" WHERE  isnull(WHNO,'') = '' ");

            if (checkBox1.Checked)
            {
                sb.Append(" AND T1.[CardName]  like  N'%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND T1.[CARDCODE2]  = '" + textBox1.Text.ToString() + "' ");
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
        private System.Data.DataTable GetCARD()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDCODE,CARDNAME  FROM ORDR WHERE CAST(DOCENTRY AS VARCHAR)=@DOCENTRY");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox3.Text));
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
        private System.Data.DataTable GetS1(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY FROM ORDR WHERE U_ACME_PAY LIKE '%OA%'  AND DOCENTRY=@DOCENTRY");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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
        private System.Data.DataTable GetS2(string DOCENTRY, string LINENUM)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM WH_ITEM4 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        private System.Data.DataTable Gettt2(string ttcode,string id)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select t0.ID,CardName,t0.Docentry,t0.ItemCode,t0.ShipDate,t0.Quantity,t0.Price,Tax,USDAmount,TTRate,");
            sb.Append(" t0.NTDAmount,case  when  isnull(t0.memo,'') = '' then t5.docentry else t0.memo end memo,t0.ID1,isnull(t0.Joy,0) Joy from satt2 t0  ");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum )");
            sb.Append(" left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )");
            sb.Append(" WHERE  ttcode=@ttcode and id=@id ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
            command.Parameters.Add(new SqlParameter("@id", id));
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


        private void CalcTotals2()
        {


            decimal NTD = 0;
            decimal USD = 0;
           

            int i = this.dataGridView5.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                string SUSD = dataGridView5.Rows[iRecs].Cells["USD"].Value.ToString();
                string SNTD = dataGridView5.Rows[iRecs].Cells["NTD"].Value.ToString();
                if (!String.IsNullOrEmpty(SUSD))
                {
                    USD += Convert.ToDecimal(SUSD);
                }
                else
                {
                    USD += 0;
                }

                if (!String.IsNullOrEmpty(SNTD))
                {
                    NTD += Convert.ToDecimal(SNTD);
                }
                else
                {
                    NTD += 0;
                }
          
            }

            label7.Text = "USD: " + USD.ToString("#,##0.00");
            label4.Text = "NTD: " + NTD.ToString("#,##0"); 
  
        }
        private void CalcTotals3()
        {


            Int32 iTotal = 0;
            decimal iVatSum = 0;


            int i = this.dataGridView1.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.Rows[iRecs].Cells["應收帳款"].Value);
                iVatSum += Convert.ToDecimal(dataGridView1.Rows[iRecs].Cells["美金應收帳款"].Value);

            }

            label6.Text = "USD: " + iVatSum.ToString("#,##0.0000");
            label3.Text = "NTD: " + iTotal.ToString("#,##0");

            System.Data.DataTable Z1 = Gettt1F();
            if (Z1.Rows.Count > 0)
            {
                for (int s = 0; s <= Z1.Rows.Count - 1; s++)
                {
                    string DOC = Z1.Rows[s]["DOCENTRY"].ToString();
                    string LINENUM = Z1.Rows[s]["LINENUM"].ToString();
                    string ID = Z1.Rows[s]["ID"].ToString();
                    System.Data.DataTable Z2 = GetS1(DOC);

                    if (Z2.Rows.Count > 0)
                    {
                        System.Data.DataTable Z3 = GetS2(DOC, LINENUM);
                        if (Z3.Rows.Count > 0)
                        {
                            string WHNO = Z3.Rows[0][0].ToString();

                            UpdateSQL2(WHNO, ID);
                        }
                    }

                }
            }
                dataGridView5.DataSource = Gettt1();
             
                dataGridView2.DataSource = GetLC1();
                System.Data.DataTable DT = Get2DO();
                if (DT.Rows.Count > 0)
                {

                    string g2 = "額度期限:" + DT.Rows[0]["額度期限"].ToString();
                    string g3 = DT.Rows[0]["信用額度"].ToString();

                    decimal sh3 = Convert.ToDecimal(g3);
                    decimal sh4 = Convert.ToDecimal(g3) - iTotal;
                    label19.Text = "信用額度:" + sh3.ToString("#,##0");
                    label17.Text = "額度餘額:" + sh4.ToString("#,##0");
                    label18.Text = g2;
                }
                else
                {
                    label19.Text = "";
                    label17.Text = "";
                    label18.Text = "";
                
                }
                     
        }

        private void CalcTotals4()
        {


            Int32 iTotal = 0;
            decimal iVatSum = 0;


            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["應收帳款"].Value);
                iVatSum += Convert.ToDecimal(dataGridView1.SelectedRows[iRecs].Cells["美金應收帳款"].Value);

            }

            label11.Text = "USD: " + iVatSum.ToString("#,##0.00");
            label2.Text = "NTD: " + iTotal.ToString("#,##0");

        }

        private void CalLC()
        {

            decimal USD = 0;
      

            int i = this.dataGridView2.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
   
                if (!String.IsNullOrEmpty(dataGridView2.Rows[iRecs].Cells["幣別"].Value.ToString()))
                {
                    if (dataGridView2.Rows[iRecs].Cells["幣別"].Value.ToString() == "USD")
                    {
                        USD += Convert.ToDecimal(dataGridView2.Rows[iRecs].Cells["餘額"].Value);
                    }
                  
                }
        
            }

            label9.Text = "USD: " + USD.ToString("#,##0.00");
     
        }

        private void CalLC2()
        {
            decimal NTD = 0;     
            int i = this.dataGridView2.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                if (!String.IsNullOrEmpty(dataGridView2.Rows[iRecs].Cells["幣別"].Value.ToString()))
                {
                    if (dataGridView2.Rows[iRecs].Cells["幣別"].Value.ToString() == "NTD")
                    {
                        NTD += Convert.ToDecimal(dataGridView2.Rows[iRecs].Cells["餘額"].Value);
                    }       
                }
         
            }

            label5.Text = "NTD: " + NTD.ToString("#,##0"); 
        
        }
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView1.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
        
            try
            {
                if (!String.IsNullOrEmpty(dgr.Cells["Column3"].Value.ToString()))
                {
                    if (dgr.Cells["Column3"].Value.ToString() == "1234")
                    {

                        dgr.DefaultCellStyle.BackColor = Color.SkyBlue;
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(dgr.Cells["逾期天數"].Value.ToString()))
                    {

                        if (Convert.ToInt32(dgr.Cells["逾期天數"].Value.ToString()) >= 0)
                        {

                            dgr.DefaultCellStyle.BackColor = Color.Pink;
                        }
                    }
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GetLC1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("          select seqno,lcno LC,MODEL 品名,QUANTITY 數量,AMOUNT 金額,QUANTITY1 餘額數量,AMOUNT1 餘額,OCCUR 幣別,EXPIRY,shipdate from Account_LC a  ");
            sb.Append("              left join Account_LC1 b on (a.LCCODE=b.LCCODE) where AMOUNT1 > 0 ");
            if (checkBox1.Checked)
            {
                sb.Append(" AND [CardName]  like '%" + artextBox12.Text.ToString() + "%' ");
            }
            else
            {
                sb.Append(" AND [CARDCODE]  = '" + textBox1.Text.ToString() + "' ");
            }

              //  sb.Append("  and  [cardname] like '%" + artextBox12.Text.ToString() + "%' ");
        

            sb.Append(" order by lcno ");
     SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

   

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

        private void dataGridView5_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView5.SelectedRows.Count > 0)
            {
                string da = dataGridView5.SelectedRows[0].Cells["Seqno"].Value.ToString();
                string da1 = dataGridView5.SelectedRows[0].Cells["ttcode"].Value.ToString();
                dataGridView6.DataSource = Gettt2(da1, da);
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView5.Rows.Count - 1; i++)
            {
                DataGridViewRow row;

                row = dataGridView5.Rows[i];

                try
                {
                    string a0 = row.Cells["ID"].Value.ToString();
                    string a1 = row.Cells["WHNO"].Value.ToString();

                    UpdateSQL(a1, a0);

                    MessageBox.Show("資料已更新");


                }
                catch (Exception ex1)
                {



                }
            }
           
        }
        private void UpdateSQL(string whno, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("update satt1 set whno=@whno where ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@whno", whno));
            command.Parameters.Add(new SqlParameter("@ID", ID));
 
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
        private void UpdateSQL2(string whno, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("update satt1 set whno=@whno where ID=@ID AND ISNULL(whno,'') ='' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@whno", whno));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        private void UpdateJoy(string Joy, string ID1)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("update satt2 set Joy=@Joy where ID1=@ID1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@Joy", Joy));
            command.Parameters.Add(new SqlParameter("@ID1", ID1));

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
        private void button3_Click(object sender, EventArgs e)
        {
             iTotal = 0;
             iVatSum = 0;
             label11.Text = "美金合計: 0" ;
             label2.Text = "台幣合計: 0" ;
        }

     

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
              
                object[] LookupValues = GetMenu.GetMenuList();

                if (LookupValues != null)
                {
                    string Y1=Convert.ToString(LookupValues[1]);

                    int J1 = Y1.IndexOf("'");

                    if (J1 > 0)
                    {
                        artextBox12.Text = Y1.Substring(0, J1);
                    }
                    else
                    {

                        artextBox12.Text = Y1;
                    }
                    textBox1.Text = Convert.ToString(LookupValues[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView6.Rows.Count - 1; i++)
            {
                DataGridViewRow row;

                row = dataGridView6.Rows[i];

                try
                {
                    string a0 = row.Cells["Joy1"].Value.ToString();
                    string a1 = row.Cells["Column6"].Value.ToString();

                    UpdateJoy(a0, a1);


                    MessageBox.Show("資料已更新");

                }
                catch (Exception ex1)
                {



                }
            }
        }

  

        private void dataGridView6_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView6.SelectedRows.Count > 0)
                {

                    for (int i = dataGridView6.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        iTotal += Convert.ToInt32(dataGridView6.SelectedRows[i].Cells["NTDAmount"].Value);
                        iVatSum += Convert.ToDecimal(dataGridView6.SelectedRows[i].Cells["USDAmount"].Value);

                    }
                    label11.Text = "美金合計: " + iVatSum.ToString("#,##0.0000");
                    label2.Text = "台幣合計: " + iTotal.ToString("#,##0");
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

        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                

                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        iTotal += Convert.ToInt32(dataGridView1.SelectedRows[i].Cells["應收帳款"].Value);
                        iVatSum += Convert.ToDecimal(dataGridView1.SelectedRows[i].Cells["美金應收帳款"].Value);

                    }
                    label11.Text = "美金合計: " + iVatSum.ToString("#,##0.0000");
                    label2.Text = "台幣合計: " + iTotal.ToString("#,##0");
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

        private void dataGridView2_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {


                    for (int i = dataGridView2.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        iTotal += Convert.ToInt32(dataGridView2.SelectedRows[i].Cells["餘額"].Value);

                    }
      
                    label2.Text = "餘額: " + iTotal.ToString("#,##0");
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            System.Data.DataTable GTCARD = GetCARD();
            if (GTCARD.Rows.Count > 0)
            {
                textBox1.Text = GTCARD.Rows[0][0].ToString();
                artextBox12.Text = GTCARD.Rows[0][1].ToString();
            }
            else 
            {
                textBox1.Text = "";
                artextBox12.Text = "";
            }
        }

     
    










    }
}
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
using System.Text.RegularExpressions;
namespace ACME
{
    
    public partial class CheckPaid : Form
    {
        string ssd;
        private decimal sd;
        private decimal sc;
        private decimal sdk;
        private decimal sk;


        string 發票總類1;
        string CARDGROUP;
        System.Data.DataTable dtCost = null;
        System.Data.DataTable dtCost32 = null;

        public CheckPaid()
        {
            InitializeComponent();
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            EXEC(dataGridView1,"0");
            EXEC2(dataGridView11);
        }


    
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("預收日期", typeof(string));
            dt.Columns.Add("預計客戶付款日", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("評價匯率", typeof(string));
            dt.Columns.Add("評價後金額", typeof(Decimal));
            dt.Columns.Add("評價損益", typeof(Decimal));
            dt.Columns.Add("收款條件", typeof(string));
            dt.Columns.Add("逾期天數", typeof(Int32));
            dt.Columns.Add("修改業管", typeof(string));
            dt.Columns.Add("業管", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("訂單號碼", typeof(string));
            dt.Columns.Add("發票總類", typeof(string));
            dt.Columns.Add("invoice", typeof(string));
            dt.Columns.Add("最終客戶", typeof(string));
            dt.Columns.Add("逾期日期", typeof(string));
            dt.Columns.Add("合併日期日", typeof(string));
            dt.Columns.Add("合併日期日2", typeof(string));
            dt.Columns.Add("SAP1", typeof(string));
            dt.Columns.Add("SAP2", typeof(string));
            dt.Columns.Add("SAP3", typeof(string));
            dt.Columns.Add("會計科目", typeof(string));
            return dt;
        }
 
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("<0", typeof(string));
            dt.Columns.Add("0~30", typeof(string));
            dt.Columns.Add("31~60", typeof(string));
            dt.Columns.Add("61~90", typeof(string));
            dt.Columns.Add("91~120", typeof(string));
            dt.Columns.Add("121~150", typeof(string));
            dt.Columns.Add("151~180", typeof(string));
            dt.Columns.Add(">180", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("上市櫃代碼", typeof(string));
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["客戶代碼"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private System.Data.DataTable MakeTableCombineSALES()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("<0", typeof(string));
            dt.Columns.Add("0~30", typeof(string));
            dt.Columns.Add("31~60", typeof(string));
            dt.Columns.Add("61~90", typeof(string));
            dt.Columns.Add("91~180", typeof(string));
            dt.Columns.Add(">180", typeof(string));
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["業務"];
            dt.PrimaryKey = colPk;

            return dt;
        }

        private System.Data.DataTable MakeTableCombineNANCY()
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            //dt.Columns.Add("客戶代碼", typeof(string));
            //dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("美金", typeof(Decimal));
            dt.Columns.Add("台幣", typeof(Decimal));
            dt.Columns.Add("未逾期", typeof(string));
            dt.Columns.Add("逾期1~30天", typeof(string));
            dt.Columns.Add("逾期31~60天", typeof(string));
            dt.Columns.Add("逾期61~90天", typeof(string));
            dt.Columns.Add("逾期91~120天", typeof(string));
            dt.Columns.Add("逾期121~150天", typeof(string));
            dt.Columns.Add("逾期151天以上", typeof(string));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["客戶代碼"];
            //dt.PrimaryKey = colPk;
            return dt;
        }
        private System.Data.DataTable MakeTableCombine32()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("<0", typeof(string));
            dt.Columns.Add("0~30", typeof(string));
            dt.Columns.Add("31~60", typeof(string));
            dt.Columns.Add("61~90", typeof(string));
            dt.Columns.Add("91~120", typeof(string));
            dt.Columns.Add("121~150", typeof(string));
            dt.Columns.Add("151~180", typeof(string));
            dt.Columns.Add(">180", typeof(string));
            dt.Columns.Add("發票總類", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("上市櫃代碼", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableCombine3()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("摘要", typeof(string));
            dt.Columns.Add("美金應收帳款", typeof(Decimal));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("逾期日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(Int32));

            return dt;
        }
 
        private System.Data.DataTable MakeTableCombine4()
        {
            System.Data.DataTable dt = new System.Data.DataTable();



            dt.Columns.Add("應收帳款", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(string));
            dt.Columns.Add("逾期天數1", typeof(Int32));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["逾期天數"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private System.Data.DataTable GetOrderDataAPDRS()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) AS INT) 應收帳款,T0.U_Delivery_date  預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管  from oinv t0");
            sb.Append("  left join (SELECT CAST(u_acme_arap AS VARCHAR) u_acme_arap,T0.docdate,doctotal  FROM ORIN T0 LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T1.BASETYPE=-1) T1 ");
            sb.Append("  on (cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and '20130703'   )  ");
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

            sb.Append(" and  T0.[cardname]  LIKE '%達睿生%'");
            sb.Append(" union all");
            sb.Append("              select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,(CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT))*-1 應收帳款,T0.U_Delivery_date 預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管  from orin t0");
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


            sb.Append(" and  T0.[cardname]  LIKE '%達睿生%'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
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
    
        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) AS INT) 應收帳款,T0.U_Delivery_date  預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管  from oinv t0");
            sb.Append("  left join (SELECT CAST(u_acme_arap AS VARCHAR) u_acme_arap,T0.docdate,doctotal  FROM ORIN T0 LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T1.BASETYPE=-1) T1 ");
            sb.Append("  on (cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and '20130703'   )  ");
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

            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and T2.SlpName =@SlpName  ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and T3.[lastName]+T3.[firstName]=@lastName ");
            }
            sb.Append(" union all");
            sb.Append("              select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,(CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT))*-1 應收帳款,T0.U_Delivery_date 預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管  from orin t0");
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

            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and T2.SlpName =@SlpName  ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and T3.[lastName]+T3.[firstName]=@lastName ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@lastName", comboBox2.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@SlpName", comboBox1.SelectedValue.ToString()));
            
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
        private System.Data.DataTable GetOrderDataAPF()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.U_IN_BSinv 發票號碼,t0.U_IN_BSTY1 憑證類別,t0.U_IN_BSTY3 通關方式,t0.U_IN_BSTY7 外銷方式,AppliedSys 應收帳款,T0.U_Delivery_date  預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管,ST4.U_USD 美金應收帳款  from oinv t0 ");
            sb.Append(" left join (SELECT CAST(u_acme_arap AS VARCHAR) u_acme_arap,T0.docdate,doctotal  FROM ORIN T0 LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T1.BASETYPE=-1) T1  ");
            sb.Append(" on (cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and '20130703'   )   ");
            sb.Append(" LEFT JOIN ocrd T9 ON (T0.cardcode=T9.cardcode)  ");
            sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode  ");
            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append(" LEFT JOIN OCRG T10 ON (T9.GROUPCODE=T10.GROUPCODE)  ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1 GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='13' ) ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1  GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='13' ) ");
            sb.Append(" LEFT JOIN (SELECT DISTINCT T1.BASETYPE,T1.BASEENTRY,T0.DOCTOTAL FROM ORIN T0 ");
            sb.Append(" LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE BASETYPE=13 and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1) T6 ON (T0.DOCENTRY=T6.BASEENTRY) ");
            sb.Append(" LEFT JOIN RCT2 ST4 ON (T0.DOCENTRY=ST4.DOCENTRY AND ST4.InvType =13)");
            sb.Append(" INNER  JOIN  ORCT ST5 ON (ST4.DOCNUM=ST5.DOCENTRY) ");
            sb.Append(" where  Convert(varchar(8),ST5.docdate,112)  =@DocDate1");
            sb.Append(" union all");
            sb.Append(" select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,'' 發票號碼,'' 憑證類別,'' 通關方式,'' 外銷方式,AppliedSys*-1 應收帳款,T0.U_Delivery_date 預收日期,T0.U_Shipping_unit 預計客戶付款日,SUBSTRING(GROUPNAME,4,10) 群組,T0.u_acme_rma_no 修改業管,ST4.U_USD 美金應收帳款  from orin t0 ");
            sb.Append(" left join oinv t1 on(cast(t1.docentry as varchar)=cast(t0.u_acme_arap as varchar) and  Convert(varchar(8),t1.docdate,112)  between '20071231' and '20171130'  )  ");
            sb.Append(" LEFT JOIN ocrd  T9 ON (T0.cardcode=T9.cardcode)  ");
            sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode  ");
            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append(" LEFT JOIN OCRG T10 ON (T9.GROUPCODE=T10.GROUPCODE)  ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1   GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='14'  ) ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1 GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='14'  ) ");
            sb.Append(" left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0   ");
            sb.Append(" WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1   ) t6 on (t0.docentry=t6.docentry) ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1   GROUP BY T1.[DocEntry],invtype )  t7 on (cast(t0.u_acme_arap as varchar)=cast(t7.docentry as varchar)  and t7.invtype='13'  ) ");
            sb.Append(" left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN  ");
            sb.Append(" [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1 GROUP BY T1.[DocEntry],invtype  ) t8 on (cast(t0.u_acme_arap as varchar)=cast(t8.docentry as varchar) and t8.invtype='13'  ) ");
            sb.Append(" left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0   ");
            sb.Append(" WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate1   ) t11 on (cast(t0.u_acme_arap as varchar)=t11.docentry) ");
            sb.Append(" LEFT JOIN RCT2 ST4 ON (T0.DOCENTRY=ST4.DOCENTRY AND ST4.InvType =14)");
            sb.Append(" INNER  JOIN  ORCT ST5 ON (ST4.DOCNUM=ST5.DOCENTRY) ");
            sb.Append(" where  Convert(varchar(8),ST5.docdate,112)  = @DocDate1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox5.Text));

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
        private System.Data.DataTable GetSACHECK(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_Delivery_date 預收日期,U_Shipping_unit 預計客戶付款日 FROM OINV WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable GetMEMO(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TRANSID FROM OINV  WHERE  (JRNLMEMO LIKE '%A/R 發票%'  OR  JRNLMEMO LIKE '%应收发票%'　OR 　ISNULL(JRNLMEMO,'')='')　AND DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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

        private System.Data.DataTable GetMEMOS(string TRANSID)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TRANSID FROM JDT1 WHERE ACCOUNT='41100102' AND TRANSID=@TRANSID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));

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
        private System.Data.DataTable GetORDR()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT (T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業助,(T0.[Cardcode]) 客戶代碼,(T0.[CardName]) 客戶名稱,T0.DOCENTRY 訂單號碼,T1.CURRENCY 幣別 ");
            sb.Append("               ,SUM(T1.GtotalFC) 金額,Convert(varchar(10),T1.U_U_PAYDATE,111)  預計付款日,T0.U_ACME_PAY 付款方式  ");
            sb.Append("               FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("               iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("               iNNER JOIN OWHS T4 ON T4.whsCode = T1.whscode  ");
            sb.Append("               WHERE    T1.[LINESTATUS] ='O' AND  T1.LINETOTAL <> 0 ");
            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and T2.SlpName = '" + comboBox1.SelectedValue.ToString() + "'  ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and T3.[lastName]+T3.[firstName]='" + comboBox2.SelectedValue.ToString() + "' ");
            }
            sb.Append(" GROUP BY (T2.[SlpName]) ,(T3.[lastName]+T3.[firstName]) ,(T0.[Cardcode]) ,(T0.[CardName]) ,T0.DOCENTRY ,T1.CURRENCY ");
            sb.Append(" ,T1.U_U_PAYDATE ,T0.U_ACME_PAY  ");
            sb.Append(" ORDER BY T0.[Cardcode]");
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
        private System.Data.DataTable GetSACHECKP(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT u_acme_rma_no 修改業管 FROM OINV WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable GetSACHECK2(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_Delivery_date 預收日期,U_Shipping_unit 預計客戶付款日 FROM ORIN WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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

        private System.Data.DataTable GetSACHECK2P(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT u_acme_rma_no 修改業管 FROM ORIN WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable Getsheet()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BU,ACTION,docmonth+docdate 日期,amt 金額USD FROM Account_Sheet where docdate <> '期初餘額'");



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

        private System.Data.DataTable Getsheet2()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select LcNo LC號碼,OPENDATE 開狀日,DUEDATE 到期日,DOCINT 利率,");
            sb.Append(" DOCRATE 匯率,CAST(ISNULL(AmtUSD,0)-ISNULL(PAIDUSD,0) AS INT) 金額USD,");
            sb.Append(" CAST(ISNULL(AmtNTD,0)-ISNULL(PAIDNTD,0) AS INT) 金額NTD,DXDATE 到單日,");
            sb.Append(" StartDate 起息日,ExpireDate 貸款到期日,BU,CRDate,Remark from Account_Bank2");


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
        private System.Data.DataTable GetOrderDataAP1(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,CASE WHEN T0.DOCENTRY=49839 AND T1.PRICE=0 THEN 'OA 30 days' WHEN T0.DOCENTRY=50079 AND T1.Quantity =299 THEN 'OA 30 days' ELSE t9.u_acme_pay END 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append("   dbo.fun_CreditDate(CASE WHEN T0.DOCENTRY=49839 AND T1.PRICE=0 THEN 'OA 30 days' WHEN T0.DOCENTRY=50079 AND T1.Quantity =299 THEN 'OA 30 days' ELSE t9.u_acme_pay END,T0.CardCode,T0.DocDate) 逾期日期,T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,'' PAYDATE");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("    LEFT JOIN RIN1 T11 ON (T1.DOCENTRY=T11.BASEENTRY AND T1.LINENUM=T11.BASELINE  )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='15'  ");
            //       sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype='15'        AND ISNULL(T11.DOCENTRY,'') ='' ");
            sb.Append("                            union all");
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("             ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t9.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   美金單價,T0.JRNLMEMO 摘要,cast(T8.docentry as varchar) 訂單號碼,t9.u_beneficiary 最終客戶, ");
            sb.Append("  dbo.fun_CreditDate(T9.u_acme_pay,T0.CardCode,T0.DocDate) 逾期日期,T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,Convert(varchar(10),T8.U_U_PAYDATE  ,111)  PAYDATE");
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
            sb.Append("              T0.COMMENTS 備註,'' 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv   美金單價,T0.JRNLMEMO 應收總計,'' 訂單號碼,'' 最終客戶, ");
            sb.Append("                       t0.docdate 逾期日期,T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,'' PAYDATE");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry=@docentry and t0.objtype=@bb and t1.basetype =-1 ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,Convert(varchar(10),t0.docdate,112) 過帳日期,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                           ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                           T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,T1.u_acme_inv 美金單價,T0.JRNLMEMO 摘要,cast(T0.u_acme_arap as varchar) 訂單號碼,'' 最終客戶,t0.docdate 逾期日期,T10.ACCTCODE+' - '+T10.ACCTNAME 發票總類,'' PAYDATE  FROM Orin T0  ");
            sb.Append("                           LEFT JOIN rin1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
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
            sb.Append("  UNION ALL  ");
            sb.Append(" SELECT isnull(T0.U_USD,0) 金額,T0.DOCENTRY 來源,T0.DOCNUM 單號 FROM VPM2  T0");
            sb.Append(" inner join OVPM t1 on (t0.docnum=t1.docnum) ");
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
        private System.Data.DataTable GetORCT(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT T0.TRANSID TRANSID   FROM  [dbo].[ORCT] T0  INNER  JOIN  ");
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
        private System.Data.DataTable GetORCT2(string TRANSID)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(DEBIT AS INT) DEBIT, Convert(varchar(10),REFDATE,111) REFDATE FROM JDT1 WHERE TRANSID=@TRANSID AND ACCOUNT=12980103  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));

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

        private System.Data.DataTable GetORCT3(string REFDATE, string CREDIT)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TRANSID  FROM JDT1 WHERE  ACCOUNT=12980103 AND Convert(varchar(10),REFDATE,111)  =@REFDATE AND CREDIT=@CREDIT ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@REFDATE", REFDATE));
            command.Parameters.Add(new SqlParameter("@CREDIT", CREDIT));
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


        private System.Data.DataTable GETORTT1()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT CAST(ROUND((CAST(BPRICE AS DECIMAL(10,3))+CAST(SPRICE AS DECIMAL(10,3)))/2,3) AS decimal(10,3)) 平均 FROM WH_HAIGUAN3 T0");
            sb.Append("   INNER JOIN (  SELECT   MAX(  Convert(varchar(10),DATE_TIME,112))  TDATE FROM    acmesqlsp.dbo.Y_2004  WHERE  (IsRestDay   =   0 OR WD = 'Y') ");
            sb.Append("  GROUP BY YEAR(DATE_TIME),MONTH(DATE_TIME) ) T1 ON (T0.HDAY =T1.TDATE)");
            sb.Append("  WHERE SUBSTRING(HDAY,1,6) =@HDAY ");

            string HD = textBox2.Text.Substring(0, 6);
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HDAY", HD));

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
        private System.Data.DataTable GetOrderDataAP2P(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(ISNULL(cast(isnull(case U_ACME_INV when '' then 0 end,0) as decimal(12,4))*QUANTITY*(1+vatprcnt/100),0)),0)  FROM RIN1  WHERE  BASETYPE=13 AND BASEENTRY=@DOCENTRY ");

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
  
        private System.Data.DataTable GetOrderDataAP3()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDCODE 客戶代碼,CARDNAME 客戶名稱,SUMAPPLIED 台幣金額,U_USD 美金金額,");
            sb.Append(" cast(cast(t1.SUMAPPLIED as int)/cast(U_USD as decimal) as decimal(5,2)) 匯率,t0.DOCDATE 逾期日期 ");
            sb.Append(" FROM ORCT T0 LEFT JOIN RCT2 T1 ON(T0.DOCNUM=T1.DOCNUM)  ");
            sb.Append(" WHERE T0.DOCNUM='4835' AND INVOICEID='33'  and Convert(varchar(8),t0.DOCDATE,112) > @DocDate2 ");
            if (artextBox12.Text.ToString() != "")
            {
                sb.Append
                    (" and T0.[cardname] like'%" + artextBox12.Text.ToString() + "%' ");
            }
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
        private void EXEC(DataGridView dgv, string F)
        {
           
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            string 工作天數;
            string 發票號碼;
            decimal 台幣金額;
            decimal 美金應收帳款 = 0;
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
            string 預收日期;
            string 預計客戶付款日;
            string 修改業管;
            string 逾期日期2;
            DateTime 逾期日期;
            System.Data.DataTable dt = null;
            System.Data.DataTable dtt = GetOrderDataAP3();

            if (F == "0")
            {
                dt = GetOrderDataAP();
            }
            if (F == "1")
            {
                dt = GetOrderDataAPF();
            }


            dtCost = MakeTableCombine();
            dtCost32 = MakeTableCombine32();
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
            //System.Data.DataTable dt3 = null;

            DataRow dr = null;
            DataRow dr22 = null;
            DataRow dr32 = null;
            DataRow dr222 = null;
            DataRow dr2222 = null;

            DataRow dr22SALES = null;
            DataRow drNANCY = null;

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("查無資料");
                return;
            }
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
                if (F == "1")
                {
                    美金應收帳款 = Convert.ToDecimal(dt.Rows[i]["美金應收帳款"]);
                }
                收款條件 = dt1.Rows[0]["收款條件"].ToString();
                美金單價 = dt1.Rows[0]["美金單價"].ToString();
                發票金額 = dt1.Rows[0]["發票金額"].ToString();
                客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                應收帳款 = dt.Rows[i]["應收帳款"].ToString();
                憑證類別 = dt.Rows[i]["憑證類別"].ToString();
                通關方式 = dt.Rows[i]["通關方式"].ToString();
                外銷方式 = dt.Rows[i]["外銷方式"].ToString();
                發票號碼 = dt.Rows[i]["發票號碼"].ToString();
                預收日期 = dt.Rows[i]["預收日期"].ToString();
                預計客戶付款日 = dt.Rows[i]["預計客戶付款日"].ToString().Replace("''", "");
                發票總類1 = dt1.Rows[0]["發票總類"].ToString();
                客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                業務 = dt.Rows[i]["業務"].ToString();
                摘要 = dt1.Rows[0]["摘要"].ToString();
                業管 = dt.Rows[i]["業管"].ToString();
                最終客戶 = dt1.Rows[0]["最終客戶"].ToString();
                CARDGROUP = dt.Rows[i]["群組"].ToString();
                修改業管 = dt.Rows[i]["修改業管"].ToString();


                dr["群組"] = CARDGROUP;
                dr["AR單號"] = 單號;
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
                dr["預收日期"] = 預收日期;

                dr["逾期日期"] = "'" + 逾期日期.ToString("yyyy") + "/" + 逾期日期.ToString("MM") + "/" + 逾期日期.ToString("dd");
                逾期日期2 = 逾期日期.ToString("yyyy") + "/" + 逾期日期.ToString("MM") + "/" + 逾期日期.ToString("dd");
                dr["修改業管"] = 修改業管;


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
                    string hg2 = dd["稅率"].ToString();
                    string QTY = dd["數量"].ToString();
                    if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                    {

                        sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);


                        if (總類 == "AR")
                        {

                            usd = sd.ToString("#,##0.0000");


                            System.Data.DataTable INV1 = GETOINV(單號);
                            if (INV1.Rows.Count > 0)
                            {
                                UpINVOICE2(單號);
                                string DOCENTRY = INV1.Rows[0]["DOCENTRY"].ToString();
                                string NUMATCARD = INV1.Rows[0]["NUMATCARD"].ToString();
                                string U_ACME_PAYGUI = INV1.Rows[0]["U_ACME_PAYGUI"].ToString();
                                string U_PN = INV1.Rows[0]["U_PN"].ToString();
         
                                string U_PN2 = "";
                                if (客戶代碼 == "0342-00")
                                {
                                    if (!String.IsNullOrEmpty(U_PN))
                                    {
                                        U_PN2 = U_PN;
                                    }
                                }

                                usd = sd.ToString("#,##0.00");
                                string ES = "更正發票請於次月7日前提出,逾期恕不受理,謝謝合作!  AR NO " + DOCENTRY + "        " + U_PN2 + NUMATCARD + " " + U_ACME_PAYGUI + " USD" + usd;
                                int k1 = ES.Length;
                                if (k1 > 200)
                                {
                                    ES = ES.Substring(0, 200);
                                }

                                UpINVOICE(DOCENTRY, ES, sd.ToString("0.00"));
                            }
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

                    //12345
                    if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日 == "'")
                    {
                        if (CARDGROUP == "ESCO")
                        {
                            string ORDR = dd["PAYDATE"].ToString();
                            預計客戶付款日 = ORDR;
                        }

                    }

                    string U1 = "'" + 預計客戶付款日;
                    string U2 = 預計客戶付款日;
                    if (U1 == "''" || U1 == "'")
                    {
                        U1 = "";
                    }
                    if (U2 == "''" || U2 == "'")
                    {
                        U2 = "";
                    }
                    U1 = U1.Replace("''", "'");
                    U2 = U2.Replace("''", "");
                    U2 = U2.Replace("'", "");
                    dr["預計客戶付款日"] = U1;
                    if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日 == "'")
                    {

                        dr["合併日期日"] = dr["逾期日期"];
                    }
                    else
                    {
                        dr["合併日期日"] = U1;
                    }

                    if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日 == "'")
                    {

                        dr["合併日期日2"] = 逾期日期2;
                    }
                    else
                    {
                        dr["合併日期日2"] = U2;
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

                    }


                }

                System.Data.DataTable dt2P = GetOrderDataAP2P(單號);
                decimal USDP = 0;
                if (dt2P.Rows.Count > 0)
                {
                    USDP = Convert.ToDecimal(dt2P.Rows[0][0].ToString());
                }
                decimal AA = Convert.ToDecimal(usd);
                string FF = AA.ToString("#,##0.0000");
                if (Convert.ToDecimal(payusd) - USDP != 0)
                {
                    dr["美金應收帳款"] = (Convert.ToDecimal(FF) - Convert.ToDecimal(payusd) - USDP).ToString();
                }

                if (F == "1")
                {
                    if (美金應收帳款 > 0)
                    {
                        dr["美金應收帳款"] = 美金應收帳款.ToString("#,##0.00");
                    }
                }

    
                if ((通關方式 == "1" && 外銷方式 == "0") || (通關方式 == "1" && 外銷方式 == "4"))
                {
                    dr["發票總類"] = "國內";
                    dr["invoice"] = 發票號碼;


                }
                else if (憑證類別 == "5")
                {
                    dr["發票總類"] = "免用";
                }
                decimal DDSZ = 0;
                int overday = GetMenu.DaySpan(逾期日期.ToString("yyyyMMdd"), textBox2.Text);
                dr["逾期天數"] = overday;
                if (!String.IsNullOrEmpty(dr["應收帳款"].ToString()) && !String.IsNullOrEmpty((dr["美金應收帳款"].ToString())))
                {

                    decimal s = Convert.ToDecimal(dr["應收帳款"]);
                    decimal v = Convert.ToDecimal(dr["美金應收帳款"]);

                    try
                    {
                        decimal dsz = Convert.ToDecimal(dr["應收帳款"]) / Convert.ToDecimal(dr["美金應收帳款"]);
                        DDSZ = dsz;
                        dr["匯率"] = dsz.ToString("#,##0.0000");
                        if (dsz < 5)
                        {
                            dr["匯率"] = "";
                            dr["美金應收帳款"] = 0;
                        }
                    }
                    catch
                    {
                        dr["匯率"] = "";
                    }
                }


                System.Data.DataTable OR1 = GETORTT1();
                if (OR1.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dr["匯率"].ToString()))
                    {
                        dr["評價匯率"] = "";
                        dr["評價後金額"] = dr["應收帳款"];
                    }
                    else
                    {
                        dr["評價匯率"] = OR1.Rows[0][0].ToString();
                        dr["評價後金額"] = (Convert.ToDecimal(dr["美金應收帳款"]) * Convert.ToDecimal(OR1.Rows[0][0])).ToString("#,##0.00");
                    }
                    dr["評價損益"] = (Convert.ToDecimal(dr["評價後金額"]) - Convert.ToDecimal(dr["應收帳款"])).ToString("#,##0.00");
                }
                StringBuilder sb1 = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();
                StringBuilder sb3 = new StringBuilder();
                System.Data.DataTable GORCT = GetORCT(單號);
                if (GORCT.Rows.Count > 0)
                {
                    for (int j = 0; j <= GORCT.Rows.Count - 1; j++)
                    {


                        string TRANSID = GORCT.Rows[j]["TRANSID"].ToString();
                    
                        sb2.Append(TRANSID + "/");

                        System.Data.DataTable GORCT2 = GetORCT2(TRANSID);
                        if (GORCT2.Rows.Count > 0)
                        {
                            string DEBIT = GORCT2.Rows[0]["DEBIT"].ToString();
                            string REFDATE = GORCT2.Rows[0]["REFDATE"].ToString();
                            System.Data.DataTable GORCT3 = GetORCT3(REFDATE, DEBIT);
                            if (GORCT3.Rows.Count > 0)
                            {
                                string TRANSID2 = GORCT3.Rows[0]["TRANSID"].ToString();
                                sb3.Append(TRANSID2 + "/");
                            }
                        }
                    }

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
                    if (!String.IsNullOrEmpty(sb2.ToString()))
                    {
                        sb2.Remove(sb2.Length - 1, 1);
                        dr["SAP2"] = sb2.ToString();
                    }
                    if (!String.IsNullOrEmpty(sb3.ToString()))
                    {
                        sb3.Remove(sb3.Length - 1, 1);
                        dr["SAP3"] = sb3.ToString();
                    }

                }

                System.Data.DataTable D1 = null;
                if (總類 == "AR")
                {
                    D1 = GetACCT(單號);

                }
                else
                {
                    D1 = GetACCT2(單號);
                }
                if (D1.Rows.Count > 0)
                {

                    dr["會計科目"] = D1.Rows[0][0].ToString();
                }
                dtCost.Rows.Add(dr);

                if (文件類型 == "13")
                {
                    string CARDNAME = 客戶;
                    string ORDER = dr["訂單號碼"].ToString();
                    int IDX1 = ORDER.IndexOf("/");
                    if (IDX1 != -1)
                    {
                        ORDER = ORDER.Substring(0, IDX1);
                    }
                    if (CARDNAME != "")
                    {
                        Regex rex = new Regex(@"^[A-Za-z0-9]+$");
                        string ENG = CARDNAME.Substring(0, 1);
                        Match ma = rex.Match(ENG);
                        if (ma.Success)
                        {
                            int t1 = CARDNAME.IndexOf(" ");
                            if (t1 != -1)
                            {
                                CARDNAME = CARDNAME.Substring(0, t1);
                            }
                        }
                        else
                        {
                            if (CARDNAME.Length > 4)
                            {
                                CARDNAME = CARDNAME.Substring(0, 4);
                            }
                        }
                    }
                    string DDDTIIME = Convert.ToInt16(過帳日期.Substring(4, 2)).ToString() + "/" + Convert.ToInt16(過帳日期.Substring(6, 2)).ToString();
                    string INVNO = 發票號碼;
                    string INVNO2 = 發票號碼;
                    if (INVNO == "__________")
                    {
                        INVNO = "三角";
                        INVNO2 = "";
                    }
                    decimal usds = Convert.ToDecimal(dr["美金應收帳款"]);

                    System.Data.DataTable GH1 = GetMEMO(單號);
                    if (GH1.Rows.Count > 0)
                    {
                        string MEMO = "";
                        string TRANSID = GH1.Rows[0][0].ToString();
                        if (CARDGROUP == "TFT")
                        {
                            decimal FF1 = Math.Round(usds, 2, MidpointRounding.AwayFromZero);
                            decimal FF2 = Math.Round(DDSZ, 3, MidpointRounding.AwayFromZero);
                            MEMO = ORDER + "/" + 單號 + CARDNAME + DDDTIIME + INVNO + "US" + FF1.ToString("G29") + "*" + FF2.ToString("G29");
                        }
                        if (CARDGROUP == "ESCO")
                        {
                            MEMO = ORDER + "/" + 單號 + CARDNAME + DDDTIIME + INVNO2;
                        }
                        if (MEMO.Length > 49)
                        {
                            MEMO = MEMO.Substring(0, 50);
                        }

                        System.Data.DataTable  K1 = GetMEMOS(TRANSID);
                        if (K1.Rows.Count > 0)
                        {
                            MEMO = MEMO.Replace("三角", "零稅");
                        }


                        UpdateSQLMEMO(單號, MEMO);
                        if (!String.IsNullOrEmpty(TRANSID))
                        {
                            UpdateTRANSID(TRANSID, MEMO);
                        }
                    }
                }

                //客戶總計明細
                dr32 = dtCost32.NewRow();
                dr32["群組"] = CARDGROUP;
                string CARDCODE = dr["客戶代碼"].ToString();
                dr32["客戶代碼"] = CARDCODE;
                dr32["客戶名稱"] = dr["客戶名稱"];
                dr32["應收帳款"] = dr["應收帳款"];
                dr32["美金應收帳款"] = dr["美金應收帳款"];



                if (overday < 0)
                {
                    dr32["<0"] = dr["應收帳款"];
                }
                if (overday >= 0 && overday <= 30)
                {
                    dr32["0~30"] = dr["應收帳款"];
                }
                if (overday > 30 && overday <= 60)
                {
                    dr32["31~60"] = dr["應收帳款"];
                }
                if (overday > 60 && overday <= 90)
                {
                    dr32["61~90"] = dr["應收帳款"];
                }
                if (overday > 90 && overday <= 120)
                {
                    dr32["91~120"] = dr["應收帳款"];
                }
                if (overday > 120 && overday <= 150)
                {
                    dr32["121~150"] = dr["應收帳款"];
                }
                if (overday > 150 && overday <= 180)
                {
                    dr32["151~180"] = dr["應收帳款"];
                }
                if (overday > 180)
                {
                    dr32[">180"] = dr["應收帳款"];
                }
                dr32["付款條件"] = 收款條件;
                dr32["發票總類"] = 發票總類1;
                System.Data.DataTable J1 = GetSTOCK(CARDCODE);
                if (J1.Rows.Count > 0)
                {
                    dr32["上市櫃代碼"] = J1.Rows[0][0].ToString();

                }
                dtCost32.Rows.Add(dr32);
            }


            for (int m = 0; m <= dtt.Rows.Count - 1; m++)
            {
                dr = dtCost.NewRow();
                DataRow dy = dtt.Rows[m];
                DateTime 逾期日期y = Convert.ToDateTime(dy["逾期日期"]);
                string exp = Convert.ToDateTime(dy["逾期日期"]).ToString("yyyyMMdd");
                dr["逾期日期"] = exp;
                dr["逾期天數"] = GetMenu.DaySpan(exp, textBox2.Text);
                dr["客戶代碼"] = dy["客戶代碼"].ToString();
                dr["客戶名稱"] = dy["客戶名稱"].ToString();
                dr["應收帳款"] = dy["台幣金額"].ToString();
                dr["美金應收帳款"] = dy["美金金額"].ToString();
                dr["匯率"] = dy["匯率"].ToString();
                dtCost.Rows.Add(dr);
     
            }
            if (dtCost.Rows.Count > 0)
            {
                dtCost.DefaultView.Sort = "客戶代碼";
                dgv.DataSource = dtCost;

                if (F == "0")
                {
                    string g = dtCost.Compute("Sum(應收帳款)", null).ToString();
                    string gk = dtCost.Compute("Sum(美金應收帳款)", null).ToString();

                    decimal sh = Convert.ToDecimal(g);
                    decimal shk = Convert.ToDecimal(gk);
                    label6.Text = "美金合計:" + shk.ToString("#,##0.0000");
                    label3.Text = "台幣合計:" + sh.ToString("#,##0");
                }
            }

            if (F == "0")
            {
                if (artextBox12.Text == "")
                {
                    DELETEPAY();
                }
                System.Data.DataTable dtCost2 = MakeTableCombine2();

                System.Data.DataTable dtSALES = MakeTableCombineSALES();
                System.Data.DataTable dtNANCY = MakeTableCombineNANCY();
                System.Data.DataTable dtCost3 = MakeTableCombine3();
                System.Data.DataTable dtCost4 = MakeTableCombine4();

                string 客戶1;
                string 客戶名稱;
                string 業務1;
                string 收款條件1;
                string 逾期天數1;

                string CARDGROUP1;
                for (int l = 0; l <= dtCost.Rows.Count - 1; l++)
                {

                    DataRow drFind;
                    DataRow drFind2;
                    DataRow drFindSALES;
                    DataRow drFindNANCY;
                    DataRow dz = dtCost.Rows[l];
                    客戶1 = dz["客戶代碼"].ToString();
                    業務1 = dz["業務"].ToString();
                    逾期天數1 = dz["逾期天數"].ToString();
                    客戶名稱 = dz["客戶名稱"].ToString();
                    發票總類1 = dz["發票總類"].ToString();
                    收款條件1 = dz["收款條件"].ToString();
                    逾期天數1 = dz["逾期天數"].ToString();
                    CARDGROUP1 = dz["群組"].ToString();
                    drFind = dtCost2.Rows.Find(客戶1);
                    drFind2 = dtCost4.Rows.Find(逾期天數1);
                    drFindSALES = dtSALES.Rows.Find(業務1);
                    //drFindNANCY = dtNANCY.Rows.Find(客戶1);
                    //if (客戶1 == "0511-00" || 客戶1 == "1349-00" || 客戶1 == "1030-00")
                    //{
                    if (drFind == null)
                    {
                        dr22 = dtCost2.NewRow();
                        string das = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "'").ToString();
                        string das1 = dtCost.Compute("Sum(美金應收帳款)", "客戶代碼='" + 客戶1 + "'").ToString();
                        string das11 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 < 0").ToString();
                        string das2 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 >= 0 and 逾期天數 <=30").ToString();
                        string das3 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 30 and 逾期天數 <=60").ToString();
                        string das4 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 60 and 逾期天數 <=90").ToString();
                        string das5 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 90 and 逾期天數 <=120").ToString();
                        string das6 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 120 and 逾期天數 <=150").ToString();
                        string das7 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 150 and 逾期天數 <=180").ToString();
                        string das8 = dtCost.Compute("Sum(應收帳款)", "客戶代碼='" + 客戶1 + "' and 逾期天數 > 180").ToString();

                        dr22["群組"] = CARDGROUP1;
                        dr22["客戶代碼"] = 客戶1;
                        dr22["客戶名稱"] = 客戶名稱;
                        System.Data.DataTable J1 = GetSTOCK(客戶1);
                        dr22["應收帳款"] = Convert.ToDecimal(das);
                        dr22["美金應收帳款"] = Convert.ToDecimal(das1);
                        string fs = Convert.ToDecimal(das1).ToString("###0");
                        dr22["<0"] = das11;
                        dr22["0~30"] = das2;
                        dr22["31~60"] = das3;
                        dr22["61~90"] = das4;
                        dr22["91~120"] = das5;
                        dr22["121~150"] = das6;
                        dr22["151~180"] = das7;
                        dr22[">180"] = das8;

                        dr22["付款條件"] = 收款條件1;
                        if (J1.Rows.Count > 0)
                        {
                            dr22["上市櫃代碼"] = J1.Rows[0][0].ToString();

                        }
                        dtCost2.Rows.Add(dr22);

                        if (artextBox12.Text == "")
                        {
                            ADDCUSTPAY(客戶1, 客戶名稱, Convert.ToInt32(das), Convert.ToInt32(fs));
                        }
                    }

                    if (drFindSALES == null)
                    {
                        dr22SALES = dtSALES.NewRow();
                        string das = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "'").ToString();
                        string das1 = dtCost.Compute("Sum(美金應收帳款)", "業務='" + 業務1 + "'").ToString();
                        string das11 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 < 0").ToString();
                        string das2 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 >= 0 and 逾期天數 <=30").ToString();
                        string das3 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 > 30 and 逾期天數 <=60").ToString();
                        string das4 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 > 60 and 逾期天數 <=90").ToString();
                        string das6 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 > 90 and 逾期天數 <=180").ToString();
                        string das5 = dtCost.Compute("Sum(應收帳款)", "業務='" + 業務1 + "' and 逾期天數 > 180").ToString();

                        dr22SALES["業務"] = 業務1;
                        dr22SALES["應收帳款"] = Convert.ToDecimal(das);
                        dr22SALES["美金應收帳款"] = Convert.ToDecimal(das1);
                        dr22SALES["<0"] = das11;
                        dr22SALES["0~30"] = das2;
                        dr22SALES["31~60"] = das3;
                        dr22SALES["61~90"] = das4;
                        dr22SALES["91~180"] = das6;
                        dr22SALES[">180"] = das5;
                        dtSALES.Rows.Add(dr22SALES);
                    }


                    dr222 = dtCost3.NewRow();

                    dr222["客戶名稱"] = dz["客戶名稱"].ToString();
                    dr222["應收帳款"] = dz["應收帳款"].ToString();
                    dr222["匯率"] = dz["匯率"].ToString();
                    dr222["美金應收帳款"] = dz["美金應收帳款"];
                    dr222["過帳日期"] = dz["過帳日期"].ToString();
                    dr222["摘要"] = dz["摘要"].ToString();
                    dr222["付款條件"] = dz["收款條件"].ToString();
                    dr222["逾期日期"] = dz["逾期日期"].ToString();
                    dr222["逾期天數"] = dz["逾期天數"];
                    dtCost3.Rows.Add(dr222);

                    if (drFind2 == null)
                    {
                        dr2222 = dtCost4.NewRow();
                        string dvs = dtCost.Compute("Sum(應收帳款)", "逾期天數 ='" + 逾期天數1 + "'").ToString();
                        dr2222["應收帳款"] = dvs;
                        dr2222["過帳日期"] = dz["過帳日期"].ToString();
                        dr2222["逾期天數"] = 逾期天數1;
                        dr2222["逾期天數1"] = dz["逾期天數"];
                        dtCost4.Rows.Add(dr2222);
                    }

                    //if (drFindNANCY == null)
                    //{


                    // }
                    //   }

                }


                drNANCY = dtNANCY.NewRow();
                string Ndas = dtCost.Compute("Sum(應收帳款)", null).ToString();
                string Ndas1 = dtCost.Compute("Sum(美金應收帳款)", null).ToString();
                string Ndas11 = dtCost.Compute("Sum(應收帳款)", "逾期天數 <= 0").ToString();
                string Ndas2 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 0 and 逾期天數 <=30").ToString();
                string Ndas3 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 30 and 逾期天數 <=60").ToString();
                string Ndas4 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 60 and 逾期天數 <=90").ToString();
                string Ndas5 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 90 and 逾期天數 <=120").ToString();
                string Ndas6 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 120 and 逾期天數 <=150").ToString();
                string Ndas7 = dtCost.Compute("Sum(應收帳款)", "逾期天數 > 150").ToString();


                drNANCY["美金"] = Convert.ToDecimal(Ndas1);
                drNANCY["台幣"] = Convert.ToDecimal(Ndas);
                drNANCY["未逾期"] = Ndas11;
                drNANCY["逾期1~30天"] = Ndas2;
                drNANCY["逾期31~60天"] = Ndas3;
                drNANCY["逾期61~90天"] = Ndas4;

                drNANCY["逾期91~120天"] = Ndas5;
                drNANCY["逾期121~150天"] = Ndas6;
                drNANCY["逾期151天以上"] = Ndas7;
                dtNANCY.Rows.Add(drNANCY);

                dataGridView10.DataSource = dtNANCY;
                dtCost2.DefaultView.Sort = "客戶代碼";
                dataGridView2.DataSource = dtCost2;

                dtCost32.DefaultView.Sort = "客戶代碼";
                dataGridView7.DataSource = dtCost32;

                dataGridView8.DataSource = dtSALES;

                dtCost3.DefaultView.RowFilter = " 逾期天數 > 0 ";
                dtCost3.DefaultView.Sort = "逾期天數 desc ";
                string g1 = dtCost3.Compute("Sum(美金應收帳款)", "逾期天數 > 0").ToString();
                string gk1 = dtCost3.Compute("Sum(應收帳款)", "逾期天數 > 0 ").ToString();
                dataGridView3.DataSource = dtCost3;
                decimal sh1 = 0;
                decimal shk1 = 0;
                if (!String.IsNullOrEmpty(g1))
                {
                    sh1 = Convert.ToDecimal(g1);
                }
                if (!String.IsNullOrEmpty(gk1))
                {
                    shk1 = Convert.ToDecimal(gk1);
                }

                label12.Text = sh1.ToString("#,##0.0000");
                label14.Text = shk1.ToString("#,##0");

                label16.Text = shk1.ToString("#,##0");
                dtCost4.DefaultView.RowFilter = " 逾期天數1 > 0 ";
                dtCost4.DefaultView.Sort = "逾期天數1 desc ";
                dataGridView4.DataSource = dtCost4;

                System.Data.DataTable O1 = GetORDR();
                if (O1.Rows.Count > 0)
                {
                    dataGridView6.DataSource = O1;

                    DataGridViewColumn col2 = dataGridView6.Columns[6];


                    col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col2.DefaultCellStyle.Format = "#,##0.0000";
                }
            }
        }
        public void DELETEPAY()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE Account_CUSTPAY2", connection);
            command.CommandType = CommandType.Text;

          
            //
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
        public void ADDCUSTPAY(string CARDCODE,string CARDNAME,int  NTD,int USD)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Account_CUSTPAY2(CARDCODE,CARDNAME,NTD,USD,DOCDATE,USERS) values(@CARDCODE,@CARDNAME,@NTD,@USD,@DOCDATE,@USERS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@NTD", NTD));
            command.Parameters.Add(new SqlParameter("@USD", USD));
            command.Parameters.Add(new SqlParameter("@DOCDATE", GetMenu.Day()));
           
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            //
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
        private void EXEC2(DataGridView dgv)
        {
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            string 工作天數;
            string 發票號碼;
            decimal 台幣金額;
            decimal 美金應收帳款=0;
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
            string 預收日期;
            string 預計客戶付款日;
            string 修改業管;
            string 逾期日期2;
            DateTime 逾期日期;
            System.Data.DataTable dt = null;


            dt = GetOrderDataAPDRS();
 

            dtCost = MakeTableCombine();

            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
         //   System.Data.DataTable dt3 = null;

            DataRow dr = null;
 
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("查無資料");
                return;
            }
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
                預收日期 = dt.Rows[i]["預收日期"].ToString();
                預計客戶付款日 = dt.Rows[i]["預計客戶付款日"].ToString().Replace("''", "");
                發票總類1 = dt1.Rows[0]["發票總類"].ToString();
                客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                業務 = dt.Rows[i]["業務"].ToString();
                摘要 = dt1.Rows[0]["摘要"].ToString();
                業管 = dt.Rows[i]["業管"].ToString();
                最終客戶 = dt1.Rows[0]["最終客戶"].ToString();
                CARDGROUP = dt.Rows[i]["群組"].ToString();
                修改業管 = dt.Rows[i]["修改業管"].ToString();

         
                dr["群組"] = CARDGROUP;
                dr["AR單號"] = 單號;
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
                dr["預收日期"] = 預收日期;

                dr["逾期日期"] = "'" + 逾期日期.ToString("yyyy") + "/" + 逾期日期.ToString("MM") + "/" + 逾期日期.ToString("dd");
                逾期日期2 = 逾期日期.ToString("yyyy") + "/" + 逾期日期.ToString("MM") + "/" + 逾期日期.ToString("dd");
                dr["修改業管"] = 修改業管;
            

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
                    string hg2 = dd["稅率"].ToString();
                    string QTY = dd["數量"].ToString();
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

                    //12345
                    if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日=="'")
                    {
                        if (CARDGROUP == "ESCO")
                        {
                            string ORDR = dd["PAYDATE"].ToString();
                            預計客戶付款日 = ORDR;
                        }
                    
                    }

                 string U1 = "'"+預計客戶付款日;
                 string U2 = 預計客戶付款日;
                 if (U1 == "''" || U1 == "'")
                 {
                     U1 = "";
                 }
                 if (U2 == "''" || U2 == "'")
                 {
                     U2 = "";
                 }
              U1 = U1.Replace("''", "'");
              U2 = U2.Replace("''", "");
              U2 = U2.Replace("'", "");
              dr["預計客戶付款日"] = U1;
                 if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日 == "'")
                    {

                        dr["合併日期日"] = dr["逾期日期"];
                    }
                    else
                    {
                        dr["合併日期日"] = U1;
                    }

                 if (String.IsNullOrEmpty(預計客戶付款日) || 預計客戶付款日 == "'")
                    {

                        dr["合併日期日2"] = 逾期日期2;
                    }
                    else
                    {
                        dr["合併日期日2"] = U2;
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

                    }


                }

                System.Data.DataTable dt2P = GetOrderDataAP2P(單號);
                decimal USDP = 0;
                if (dt2P.Rows.Count > 0)
                {
                    USDP = Convert.ToDecimal(dt2P.Rows[0][0].ToString());
                }
                decimal AA = Convert.ToDecimal(usd);
                string FF = AA.ToString("#,##0.0000");
                if (Convert.ToDecimal(payusd) - USDP != 0)
                {
                    dr["美金應收帳款"] = (Convert.ToDecimal(FF) - Convert.ToDecimal(payusd) - USDP).ToString();
                }

          

                //dt3 = GetOrderinv(單號, 文件類型);

                //if (通關方式 == "0" || (通關方式 == "1" && 外銷方式 == "1"))
                //{
                //    dr["發票總類"] = "國外";
                //    for (int p = 0; p <= dt3.Rows.Count - 1; p++)
                //    {
                //        DataRow ddp = dt3.Rows[p];


                //        if (dt3.Rows.Count == 1)
                //        {
                //            dr["invoice"] = ddp["invoice"].ToString();

                //        }
                //        else
                //        {

                //            if (p == dt3.Rows.Count - 1)
                //            {
                //                dr["invoice"] += ddp["invoice"].ToString();

                //            }
                //            else
                //            {

                //                dr["invoice"] += ddp["invoice"].ToString() + "/";


                //            }
                //        }
                //    }
                //}
                //else 
                if ((通關方式 == "1" && 外銷方式 == "0") || (通關方式 == "1" && 外銷方式 == "4"))
                {
                    dr["發票總類"] = "國內";
                    dr["invoice"] = 發票號碼;


                }
                else if (憑證類別 == "5")
                {
                    dr["發票總類"] = "免用";
                }
                decimal DDSZ=0;
                int overday = GetMenu.DaySpan(逾期日期.ToString("yyyyMMdd"), textBox2.Text);
                dr["逾期天數"] = overday;
                if (!String.IsNullOrEmpty(dr["應收帳款"].ToString()) && !String.IsNullOrEmpty((dr["美金應收帳款"].ToString())))
                {

                    decimal s = Convert.ToDecimal(dr["應收帳款"]);
                    decimal v = Convert.ToDecimal(dr["美金應收帳款"]);

                    try
                    {
                        decimal dsz = Convert.ToDecimal(dr["應收帳款"]) / Convert.ToDecimal(dr["美金應收帳款"]);
                        DDSZ=dsz;
                        dr["匯率"] = dsz.ToString("#,##0.0000");
                        if (dsz < 5)
                        {
                            dr["匯率"] = "";
                            dr["美金應收帳款"] = 0;
                        }
                    }
                    catch
                    {
                        dr["匯率"] = "";
                    }
                }

     
                dtCost.Rows.Add(dr);



            }

            if (dtCost.Rows.Count > 0)
            {
                dtCost.DefaultView.Sort = "客戶代碼";
                dgv.DataSource = dtCost;
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
 
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\應收帳款.xlsx";


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                EXEC(dataGridView1, "0");
                //產生 Excel Report
                ExcelReport.ExcelReportOutput(dtCost, ExcelTemplate, OutPutFile, "N");
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView9);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView6);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToExcel(dataGridView7);
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 7)
            {
                ExcelReport.GridViewToExcel(dataGridView8);
            }
            else if (tabControl1.SelectedIndex == 8)
            {
                ExcelReport.GridViewToExcel(dataGridView5);
            }
            else if (tabControl1.SelectedIndex == 10)
            {
                ExcelReport.GridViewToExcel(dataGridView11);
            }
        }
        private System.Data.DataTable GetACCT(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(T1.Account+'-'+T2.AcctName) 會計科目   FROM OINV T0");
            sb.Append(" LEFT JOIN JDT1 T1 ON (T0.TransId =T1.TransId)");
            sb.Append(" LEFT JOIN OACT T2 ON (T1.Account =T2.AcctCode)");
            sb.Append("  WHERE T0.DOCENTRY=@DOCENTRY AND T1.Account IN (11420101,11430101)");

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


        private System.Data.DataTable GetACCT2(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(T1.Account+'-'+T2.AcctName) 會計科目   FROM ORIN T0");
            sb.Append(" LEFT JOIN JDT1 T1 ON (T0.TransId =T1.TransId)");
            sb.Append(" LEFT JOIN OACT T2 ON (T1.Account =T2.AcctCode)");
            sb.Append("  WHERE T0.DOCENTRY=@DOCENTRY AND T1.Account IN (11420101,11430101)");

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
        private System.Data.DataTable GetSCE()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT sum(T1.[Credit]-T1.[Debit]) 金額,CAST(sum(T1.[Credit]-T1.[Debit])/32 AS DECIMAL(10,2)) 美金金額,shortname 客戶編號,T2.CARDNAME 客戶名稱,T3.PARAM_DESC 逾期日期 FROM  [dbo].[JDT1] T1  ");
            sb.Append("LEFT JOIN OCRD T2 ON(T1.SHORTNAME=T2.CARDCODE )  ");
            sb.Append("INNER JOIN ACMESQLSP.DBO.PARAMS T3 ON(T1.SHORTNAME=T3.PARAM_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS  AND T3.PARAM_KIND='SCE' )  ");
             sb.Append("group by shortname,T2.CARDNAME,T3.PARAM_DESC");

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
        private System.Data.DataTable GetSTOCK(string CARDCODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT [PASSWORD] FROM OCRD WHERE CARDCODE=@CARDCODE AND ISNULL(PASSWORD,'') <> ''  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
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
        private void CheckPaid_Load(object sender, EventArgs e)
        {
            if (globals.GroupID.ToString().Trim() == "ACCS" )
            {
                button1.Visible = false;
                button2.Visible = false;
                return;
            }
            label6.Text = "";
            label3.Text = "";

            label12.Text = "";
            label14.Text = "";
            label16.Text = "";


     
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetOslp1(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetOhem(), "DataValue", "DataValue");

            textBox2.Text = GetMenu.Day();
            textBox5.Text = GetMenu.Day();
            dataGridView1.ReadOnly = false;
            dataGridView1.Enabled = true;
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



        private void GridViewToExcel(DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;
            string  DuplicateKey = "";
            int n = 0;
            try
            {

           
                for (int i = 0; i < dgv.Columns.Count; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;



                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];
                        DataGridViewCell cell2 = row.Cells[1];
                        try
                        {


                            if (j == 0)
                            {
                                string dd = cell2.Value.ToString();
                                if (dd != DuplicateKey && !String.IsNullOrEmpty(DuplicateKey))
                                {

                                    n++;
                                }
                             
                                DuplicateKey = dd;
                            }

                            wsheet.Cells[i + 2 + n, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                        }

                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);

                        }

                    }

                }

                wapp.Visible = true;


            }

            catch (Exception ex1)
            {

                MessageBox.Show(ex1.Message);

            }

            wapp.UserControl = true;


        }


        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                DataGridViewRow row;

                row = dataGridView1.Rows[i];

                 string a0 = row.Cells["AR單號"].Value.ToString();
                    string a1 = row.Cells["預收日期"].Value.ToString();
                    string a2 = row.Cells["預計客戶付款日"].Value.ToString();
                
                    int ff = Convert.ToInt32(row.Cells["應收帳款"].Value.ToString());
                    if (ff < 0)
                    {
                        System.Data.DataTable bb = GetSACHECK2(a0);
                        if (bb.Rows.Count > 0)
                        {
                            DataRow drw1 = bb.Rows[0];
                            string a4 = drw1["預收日期"].ToString();
                            string a5 = drw1["預計客戶付款日"].ToString();
                            if (a1 != a4)
                            {

                                UpdateSQL2(a0, a1);
                            }
                            if (a2 != a5)
                            {

                                UpdateSQL2S(a0, a2);
                            }
                        }
                    }
                    else
                    {
                        System.Data.DataTable bb = GetSACHECK(a0);
                        if (bb.Rows.Count > 0)
                        {
                            DataRow drw1 = bb.Rows[0];
                            string a4 = drw1["預收日期"].ToString();
                            string a5 = drw1["預計客戶付款日"].ToString();
                            if (a1 != a4)
                            {

                                UpdateSQL(a0, a1);
                            }
                            if (a2 != a5)
                            {

                                UpdateSQLS(a0, a2);
                            }
                        }
                    }
            }

            MessageBox.Show("修改成功");

          }


        private void UpdateSQL(string DOCENTRY, string U_Delivery_date)
        {

  

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OINV SET U_Delivery_date=@U_Delivery_date WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_Delivery_date", U_Delivery_date));
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
        private void UpdateSQLS(string DOCENTRY, string U_Shipping_unit)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OINV SET U_Shipping_unit=@U_Shipping_unit WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_Shipping_unit", U_Shipping_unit));
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
        private void UpdateSQLP(string DOCENTRY, string u_acme_rma_no)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OINV SET u_acme_rma_no=@u_acme_rma_no WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@u_acme_rma_no", u_acme_rma_no));
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
        private void UpdateSQLMEMO(string DOCENTRY, string JRNLMEMO)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OINV SET JRNLMEMO=@JRNLMEMO WHERE  DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@JRNLMEMO", JRNLMEMO));
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
        private void UpINVOICE(string DOCENTRY, string U_ACME_UNIT1, string U_ACME_Price1)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   update  OINV SET U_ACME_UNIT1=@U_ACME_UNIT1,U_ACME_Price1=@U_ACME_Price1 WHERE DOCENTRY =@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_ACME_UNIT1", U_ACME_UNIT1));
            command.Parameters.Add(new SqlParameter("@U_ACME_Price1", U_ACME_Price1));
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
        private void UpINVOICE2(string DOCENTRY)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("				Declare @name varchar(100) ");
            sb.Append("				select @name =SUBSTRING(COALESCE(@name + '/',''),0,99) + U_CUSTDOCENTRY ");
            sb.Append("				from   (SELECT   DISTINCT U_CUSTDOCENTRY    FROM INV1 WHERE DOCENTRY=@DOCENTRY) pc");
            sb.Append("				iF (ISNULL(@name,'') <>'')");
            sb.Append("				BEGIN");
            sb.Append("				update OINV sET numatcard=@name");
            sb.Append("				 WHERE DOCENTRY =@DOCENTRY AND UPDINVNT='I' AND substring(numatcard,len(numatcard),1)<>'~'");
            sb.Append("				 END");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
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
        private void UpdateTRANSID(string TRANSID, string LineMemo)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE JDT1 SET LineMemo=@LineMemo WHERE  TRANSID=@TRANSID");
         

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));
            command.Parameters.Add(new SqlParameter("@LineMemo", LineMemo));
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
        private void UpdateSQL2(string DOCENTRY, string U_Delivery_date)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE ORIN SET U_Delivery_date=@U_Delivery_date WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_Delivery_date", U_Delivery_date));
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
        private void UpdateSQL2S(string DOCENTRY, string U_Shipping_unit)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE ORIN SET U_Shipping_unit=@U_Shipping_unit WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_Shipping_unit", U_Shipping_unit));
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
        private void UpdateSQL2P(string DOCENTRY, string u_acme_rma_no)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE ORIN SET u_acme_rma_no=@u_acme_rma_no WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@u_acme_rma_no", u_acme_rma_no));
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
    

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                DataGridViewRow row;

                row = dataGridView1.Rows[i];

                string a0 = row.Cells["AR單號"].Value.ToString();
                string a1 = row.Cells["修改業管"].Value.ToString();
                int ff = Convert.ToInt32(row.Cells["應收帳款"].Value.ToString());
                if (ff < 0)
                {
                    System.Data.DataTable bb = GetSACHECK2P(a0);
                    if (bb.Rows.Count > 0)
                    {
                        DataRow drw1 = bb.Rows[0];
                        string a4 = drw1["修改業管"].ToString();

                        if (a1 != a4)
                        {

                            UpdateSQL2P(a0, a1);
                        }

                    }
                }
                else
                {
                    System.Data.DataTable bb = GetSACHECKP(a0);
                    if (bb.Rows.Count > 0)
                    {
                        DataRow drw1 = bb.Rows[0];
                        string a4 = drw1["修改業管"].ToString();

                        if (a1 != a4)
                        {

                            UpdateSQLP(a0, a1);
                        }

                    }
                }
            }

            MessageBox.Show("修改成功");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = Getprepareend();
        }

        private System.Data.DataTable Getprepareend()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (T0.[CardName]) 客戶名稱,(T8.NUMATCARD) PONO,T0.[docDate] 過帳日期,T5.[docentry] SO ,t0.docentry AR,");
            sb.Append(" CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append(" Substring (T1.[ItemCode],2,8) END Model,");
            sb.Append(" CASE (Substring(T1.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append("  when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append("  when '4' then 'U' when '5' then 'NN' ELSE 'X'");
            sb.Append("  END 等級,");
            sb.Append(" ''''+(Substring(T1.[ItemCode],12,3)) 版本,CAST(t1.quantity AS INT) 數量,");
            sb.Append(" T5.CURRENCY+cast(cast(T5.[Price] as numeric(16,2)) as varchar) 美金單價,cast(T5.[GtotalFC] as varchar) 美金總金額");
            sb.Append(" ,(T1.[Gtotal]) 台幣總金額,T0.U_IN_BSINV 發票號碼, Convert(varchar(8),T0.U_IN_BSDAT,112)  發票日期,  Convert(varchar(8),dbo.fun_CreditDate(T8.u_acme_pay,T0.CardCode,T0.DocDate),112)  到帳日期,");
            sb.Append(" DATEDIFF ( D , dbo.fun_CreditDate(T0.u_acme_pay,T0.CardCode,T0.DocDate) , cast(@DocDate2 as datetime) ) 逾期天數,t0.u_acme_pay 付款條件,");
            sb.Append("  (T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管");
            sb.Append(" FROM OINV T0 ");
            sb.Append(" INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode");
            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
            sb.Append(" left join odln t9 on (t4.docentry=T9.docentry )");
            sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
            sb.Append(" left JOIN Owhs T7 ON T7.whsCode = T1.whscode ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE ");
            sb.Append(" left JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod");
            sb.Append(" WHERE ");
            sb.Append(" T0.DOCENTRY IN (      select distinct t0.docentry  from oinv t0");
            sb.Append("       left join orin t1 on(cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2     )");
            sb.Append("       left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("       [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='13' )");
            sb.Append("       left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("       [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='13' )");
            sb.Append("       LEFT JOIN (SELECT DISTINCT T1.BASETYPE,T1.BASEENTRY,T0.DOCTOTAL FROM ORIN T0");
            sb.Append("       LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("       WHERE BASETYPE=13 and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2) T6 ON (T0.DOCENTRY=T6.BASEENTRY)");
            sb.Append("        where  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) <>  0  and ((isnull(t0.doctotal,0)-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0)) - isnull(t1.doctotal,0)) <> 0  and t0.docentry not in (SELECT PARAM_NO FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='SAPAR'  AND PARAM_DESC  between '20071231' AND @DocDate2) AND T0.CARDCODE <> 'R0001'  )");
            if (textBox4.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + textBox4.Text.ToString() + "%'");
            }
            sb.Append(" order by (t0.docentry) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

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


        private System.Data.DataTable GETOINV(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT DOCENTRY,NUMATCARD,U_ACME_PAYGUI,'P/N:'+U_PN+ DBO.fun_PNLEN(19-LEN('P/N:'+U_PN) % 18)+'PO:' U_PN  FROM   OINV  WHERE DOCENTRY =@DOCENTRY ");
  
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
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView5);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            EXEC(dataGridView9, "1");
        }


    }
}
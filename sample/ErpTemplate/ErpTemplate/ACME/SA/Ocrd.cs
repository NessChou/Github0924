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
    public partial class Ocrd : Form
    {
        public Ocrd()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBatchPayment();

        }
        private void ViewBatchPayment()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked == false)
            {
                sb.Append(" select [CardCode] , [CardName],SUM(總數量) 總數量,AVG(平均單價) 平均單價, SUM(總金額) 總金額, SUM(總成本) 總成本, SUM(總毛利) 總毛利,銷售人員 ");
                sb.Append("  from (sELECT T0.[CardCode], T0.[CardName],Convert(varchar(10),max(t0.docdate),112) 過帳日期, SUM(T1.[Quantity]) 總數量,AVG( T1.[Price]) 平均單價, SUM(T1.[LineTotal]) 總金額, SUM(T1.[Quantity] * T1.[GrossBuyPr] ) 總成本, SUM(T1.[GrssProfit]) 總毛利,T3.[CHINESENAME] 銷售人員  FROM acmesql01.dbo.OINV T0  INNER JOIN acmesql01.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T3 ON (T0.SlpCode=T3.EMPID AND T3.KIND='sales') ");
                sb.Append(" WHERE T0.[DocType] ='I' ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T3.[CHINESENAME] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append(" And t1.itemcode not in (select itemcode from acmesql01.dbo.oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))");
                sb.Append(" GROUP BY T0.[CardCode], T0.[CardName], T3.[CHINESENAME] ");
                sb.Append(" union all");
                sb.Append(" SELECT T0.[CardCode], T0.[CardName], Convert(varchar(10),max(t0.docdate),112) 過帳日期, SUM(T1.[Quantity])* (-1)  總數量,AVG( T1.[Price]) 平均單價, SUM(T1.[LineTotal]) * (-1) 總金額, SUM(T1.[Quantity] * T1.[GrossBuyPr] ) * (-1) 總成本, SUM(T1.[GrssProfit]) * (-1) 總毛利,T3.[CHINESENAME] AS 銷售人員 FROM acmesql01.dbo.ORIN T0 ");
                sb.Append(" INNER JOIN acmesql01.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T3 ON (T0.SlpCode=T3.EMPID AND T3.KIND='sales') ");
                sb.Append(" WHERE T0.[DocType] ='I' ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T3.[CHINESENAME] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("  GROUP BY T0.[CardCode], T0.[CardName],T3.[CHINESENAME] ");
                sb.Append(" union all");
                sb.Append(" sELECT T0.[CardCode], T0.[CardName], Convert(varchar(10),max(t0.docdate),112) 過帳日期,SUM(T1.[Quantity]) 總數量,AVG( T1.[Price]) 平均單價, SUM(T1.[LineTotal]) 總金額, SUM(T1.[Quantity] * T1.[GrossBuyPr] ) 總成本, SUM(T1.[GrssProfit]) 總毛利,T3.[CHINESENAME] 銷售人員  FROM acmesql02.dbo.OINV T0  INNER JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append(" LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T3 ON (T0.SlpCode=T3.EMPID AND T3.KIND='sales') ");
                sb.Append(" WHERE T0.[DocType] ='I' ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T3.[CHINESENAME] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append(" And t1.itemcode not in (select itemcode from acmesql02.dbo.oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                sb.Append(" GROUP BY T0.[CardCode], T0.[CardName],T3.[CHINESENAME] ");
                sb.Append(" union all");
                sb.Append(" SELECT T0.[CardCode], T0.[CardName], Convert(varchar(10),max(t0.docdate),112)  過帳日期, SUM(T1.[Quantity])* (-1)  總數量,AVG( T1.[Price]) 平均單價, SUM(T1.[LineTotal]) * (-1) 總金額, SUM(T1.[Quantity] * T1.[GrossBuyPr] ) * (-1) 總成本, SUM(T1.[GrssProfit]) * (-1) 總毛利,T3.[CHINESENAME] AS 銷售人員 FROM acmesql02.dbo.ORIN T0 ");
                sb.Append(" INNER JOIN acmesql02.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T3 ON (T0.SlpCode=T3.EMPID AND T3.KIND='sales') ");
                sb.Append(" WHERE T0.[DocType] ='I' ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T3.[CHINESENAME] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("  GROUP BY T0.[CardCode], T0.[CardName],T3.[CHINESENAME] ");
                sb.Append(" union all ");
                sb.Append(" SELECT 'CHOICE'  AS CardCode ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS CardName,IsNull(B.日期,'') AS 過帳日期");
                sb.Append("                ,SUM(IsNull(C.數量,0)) AS 總數量,AVG(IsNull(C.單價,0)) AS 平均單價,SUM(IsNull(C.金額,0)) AS 總金額");
                sb.Append("                ,SUM(IsNull(C.平均成本,0)) AS 總成本,ROUND(SUM((IsNull(C.單價,0)-IsNull(C.平均成本,0))*C.數量),0) AS 總毛利");
                sb.Append("              ,A.姓名  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 銷售人員  ");
                sb.Append("              FROM (((OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.comPerson A ");
                sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
                sb.Append("  OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
                sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.comCustomer D ON D.編號=B.廠商編號)");
                sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
                sb.Append("              WHERE   B.旗標>= 5");
                sb.Append("              AND D.旗標= 1");
                sb.Append("              AND B.旗標<= 6");

                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
                sb.Append(" GROUP BY D.編號,D.公司簡稱,A.姓名,B.日期");
                sb.Append(" union all ");
                sb.Append(" SELECT 'TOP'  AS CardCode ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS CardName,IsNull(B.日期,'') AS 過帳日期");
                sb.Append("                ,SUM(IsNull(C.數量,0)) AS 總數量,AVG(IsNull(C.單價,0)) AS 平均單價,SUM(IsNull(C.金額,0)) AS 總金額");
                sb.Append("                ,SUM(IsNull(C.平均成本,0)) AS 總成本,ROUND(SUM((IsNull(C.單價,0)-IsNull(C.平均成本,0))*C.數量),0) AS 總毛利");
                sb.Append("              ,A.姓名  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 銷售人員  ");
                sb.Append("              FROM (((OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.comPerson A ");
                sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
                sb.Append("  OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
                sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.comCustomer D ON D.編號=B.廠商編號)");
                sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
                sb.Append("              WHERE B.旗標>= 5");
                sb.Append("              AND D.旗標= 1");
                sb.Append("              AND B.旗標<= 6");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
                sb.Append(" GROUP BY D.編號,D.公司簡稱,A.姓名,B.日期");
                sb.Append(" )  as aa where 1=1 ");
                sb.Append("group by [CardCode],[CardName],銷售人員 ");
            }
            else
            {
                sb.Append(" SELECT 'CHOICE'  AS CardCode ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS CardName");
                sb.Append("                ,SUM(IsNull(C.數量,0)) AS 總數量,AVG(IsNull(C.單價,0)) AS 平均單價,SUM(IsNull(C.金額,0)) AS 總金額");
                sb.Append("                ,SUM(IsNull(C.平均成本,0)) AS 總成本,ROUND(SUM((IsNull(C.單價,0)-IsNull(C.平均成本,0))*C.數量),0) AS 總毛利");
                sb.Append("              ,A.姓名  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 銷售人員  ");
                sb.Append("              FROM (((OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.comPerson A ");
                sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
                sb.Append("  OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
                sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL21.DBO.comCustomer D ON D.編號=B.廠商編號)");
                sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
                sb.Append("              WHERE   B.旗標>= 5");
                sb.Append("              AND D.旗標= 1");
                sb.Append("              AND B.旗標<= 6");

                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
                sb.Append(" GROUP BY D.編號,D.公司簡稱,A.姓名");
                sb.Append(" union all ");
                sb.Append(" SELECT 'TOP'  AS CardCode ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS CardName");
                sb.Append("                ,SUM(IsNull(C.數量,0)) AS 總數量,AVG(IsNull(C.單價,0)) AS 平均單價,SUM(IsNull(C.金額,0)) AS 總金額");
                sb.Append("                ,SUM(IsNull(C.平均成本,0)) AS 總成本,ROUND(SUM((IsNull(C.單價,0)-IsNull(C.平均成本,0))*C.數量),0) AS 總毛利");
                sb.Append("              ,A.姓名  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 銷售人員  ");
                sb.Append("              FROM (((OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.comPerson A ");
                sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
                sb.Append("  OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
                sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
                sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
                sb.Append("      ).SUNSQL20.DBO.comCustomer D ON D.編號=B.廠商編號)");
                sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
                sb.Append("              WHERE B.旗標>= 5");
                sb.Append("              AND D.旗標= 1");
                sb.Append("              AND B.旗標<= 6");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
                sb.Append(" GROUP BY D.編號,D.公司簡稱,A.姓名");
               
           
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

            label4.Text = ds.Tables[0].Compute("Sum(總數量)", null).ToString();
            label5.Text = ds.Tables[0].Compute("Sum(總金額)", null).ToString();
            label6.Text = ds.Tables[0].Compute("Sum(總成本)", null).ToString();
            label7.Text = ds.Tables[0].Compute("Sum(總毛利)", null).ToString();
        }
        private void ViewBatchPayment2()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked == false)
            {
                sb.Append("    sELECT 'AR發票' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱,Convert(varchar(10),(t0.docdate),112) 過帳日期,T1.ITEMCODE 產品編號,  CAST((T1.[Quantity]) AS INT) 數量,( T1.[Price]) 單價, (T1.[LineTotal]) 金額, (T1.[Quantity] * T1.[GrossBuyPr] ) 成本, (T1.[GrssProfit]) 毛利,T0.COMMENTS 備註 FROM acmesql01.dbo.OINV T0  INNER JOIN acmesql01.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("              left JOIN acmesql01.dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");
                sb.Append("              WHERE T0.[DocType] ='I' ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T2.[SlpName] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("            And t1.itemcode not in (select itemcode from acmesql02.dbo.oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))");
                sb.Append("              union all                       ");
                sb.Append("    sELECT 'AR發票' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱, Convert(varchar(10),(t0.docdate),112) 過帳日期,T1.ITEMCODE 產品編號,CAST((T1.[Quantity]) AS INT) 數量,( T1.[Price]) 單價, (T1.[LineTotal]) 金額, (T1.[Quantity] * T1.[GrossBuyPr] ) 成本,(T1.[GrssProfit]) 毛利,T0.COMMENTS 備註 FROM acmesql02.dbo.OINV T0  INNER JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("              left JOIN acmesql02.dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");
                sb.Append("              WHERE T0.[DocType] ='I' ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T2.[SlpName] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("           And t1.itemcode not in (select itemcode from acmesql01.dbo.oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))");
                sb.Append("              union all");
                sb.Append("    SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱, Convert(varchar(10),(t0.docdate),112) 過帳日期,T1.ITEMCODE 產品編號, CAST((T1.[Quantity]) AS INT) * (-1)  數量,( T1.[Price]) 單價, (T1.[LineTotal]) * (-1) 金額, (T1.[Quantity] * T1.[GrossBuyPr] ) * (-1) 成本, (T1.[GrssProfit]) * (-1) 毛利,T0.COMMENTS 備註 FROM acmesql01.dbo.ORIN T0 ");
                sb.Append("              INNER JOIN acmesql01.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append("              left JOIN acmesql01.dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");
                sb.Append("              WHERE T0.[DocType] ='I'  ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T2.[SlpName] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("      union all     ");
                sb.Append("              SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,T0.[CardCode] 客戶編號, T0.[CardName] 客戶名稱, Convert(varchar(10),(t0.docdate),112)  過帳日期,T1.ITEMCODE 產品編號, CAST((T1.[Quantity]) AS INT)* (-1)  數量,( T1.[Price]) 單價, (T1.[LineTotal]) * (-1) 金額, (T1.[Quantity] * T1.[GrossBuyPr] ) * (-1) 成本, (T1.[GrssProfit]) * (-1) 毛利,T0.COMMENTS 備註 FROM acmesql02.dbo.ORIN T0 ");
                sb.Append("              INNER JOIN acmesql02.dbo.RIN1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append("              left JOIN acmesql02.dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");

                sb.Append("              WHERE T0.[DocType] ='I' ");
                sb.Append(" AND T0.CARDCODE NOT IN ('0511-00','0257-00')");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and Convert(varchar(10),(t0.docdate),112) between  @DocDate1 and @DocDate2 ");
                }
                if (comboBox1.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and T2.[SlpName] = '" + comboBox1.SelectedValue.ToString() + "'  ");
                }
                sb.Append("      union all     ");
            }
            sb.Append(" SELECT '正航' AS '總類',SUBSTRING(B.單號,3,10) AS 單號,'CHOICE'  AS 客戶編號 ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 客戶名稱,IsNull(B.日期,'') AS 過帳日期");
            sb.Append("            ,C.產品編號 COLLATE Chinese_Taiwan_Stroke_CI_AS 產品編號    ,IsNull(C.數量,0) AS 數量,IsNull(C.單價,0) AS 單價,IsNull(C.金額,0) AS 金額");
            sb.Append("                ,IsNull(C.平均成本,0) AS 成本,ROUND(IsNull(C.金額,0)-IsNull(C.平均成本,0),0) AS 毛利,B.備註 COLLATE Chinese_Taiwan_Stroke_CI_AS 備註");
            sb.Append("              FROM ((((OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL21.DBO.comPerson A ");
            sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL21.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
            sb.Append("  OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL21.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
            sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL21.DBO.comCustomer D ON D.編號=B.廠商編號)");
            sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL21.DBO.comProduct F ON F.產品編號=C.產品編號)");
            sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
            sb.Append("              WHERE  B.旗標>= 5");
              if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
            }
            sb.Append("              AND D.旗標= 1");
            sb.Append("              AND B.旗標<= 6");
            sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
            sb.Append("      union all     ");
            sb.Append(" SELECT '正航' AS '總類',SUBSTRING(B.單號,3,10) AS 單號,'TOP'  AS 客戶編號 ,D.公司簡稱  COLLATE Chinese_Taiwan_Stroke_CI_AS AS 客戶名稱,IsNull(B.日期,'') AS 過帳日期");
            sb.Append("            ,C.產品編號 COLLATE Chinese_Taiwan_Stroke_CI_AS 產品編號    ,IsNull(C.數量,0) AS 數量,IsNull(C.單價,0) AS 單價,IsNull(C.金額,0) AS 金額");
            sb.Append("                ,IsNull(C.平均成本,0) AS 成本,ROUND(IsNull(C.金額,0)-IsNull(C.平均成本,0),0) AS 毛利,B.備註 COLLATE Chinese_Taiwan_Stroke_CI_AS 備註");
            sb.Append("              FROM ((((OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL20.DBO.comPerson A ");
            sb.Append("  Inner Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL20.DBO.stkBillMain B ON B.採購員=A.編號) Left Join");
            sb.Append("  OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL20.DBO.stkBillSub C ON C.旗標=B.旗標 AND C.單號=B.單號)");
            sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL20.DBO.comCustomer D ON D.編號=B.廠商編號)");
            sb.Append("              Left Join OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("      ).SUNSQL20.DBO.comProduct F ON F.產品編號=C.產品編號)");
            sb.Append("  LEFT JOIN ACMESQLSP.DBO.EMPLOYEE T0 ON (T0.CHINESENAME=A.姓名)");
            sb.Append("              WHERE  B.旗標>= 5");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and B.日期 between  @DocDate1 and @DocDate2 ");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and A.姓名 = '" + comboBox1.SelectedValue.ToString() + "'  ");
            }
            sb.Append("              AND D.旗標= 1");
            sb.Append("              AND B.旗標<= 6");
            sb.Append("              AND C.產品編號 <> '*' AND  C.產品編號 NOT LIKE 'R%' AND C.產品編號 <> '(*)'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource2.DataSource = ds.Tables[0];
            dataGridView2.DataSource = bindingSource2;


        }
        System.Data.DataTable GetOslp()
        {

            SqlConnection con = globals.Connection;
            string sql = "select chinesename as DataValue from employee where kind='sales'  UNION ALL SELECT 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM' ORDER BY DataValue";


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

        private void Ocrd_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetOslp(), "DataValue", "DataValue");
            textBox1.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox2.Text = DateTime.Now.ToString("yyyyMMdd");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;
            try
            {

                int iX;
                int iY;



                for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
                {


                    wsheet.Cells[1, i + 1] = this.dataGridView1.Columns[i].HeaderText;



                }



                for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                {



                    DataGridViewRow row = this.dataGridView1.Rows[i];



                    for (int j = 0; j < row.Cells.Count; j++)
                    {



                        DataGridViewCell cell = row.Cells[j];



                        try
                        {



                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();



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

        private void button3_Click(object sender, EventArgs e)
        {
            ViewBatchPayment2();
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;
            try
            {

                int iX;
                int iY;



                for (int i = 0; i < this.dataGridView2.Columns.Count; i++)
                {


                    wsheet.Cells[1, i + 1] = this.dataGridView2.Columns[i].HeaderText;



                }



                for (int i = 0; i < this.dataGridView2.Rows.Count; i++)
                {



                    DataGridViewRow row = this.dataGridView2.Rows[i];



                    for (int j = 0; j < row.Cells.Count; j++)
                    {



                        DataGridViewCell cell = row.Cells[j];



                        try
                        {



                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();



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

       
    }
}
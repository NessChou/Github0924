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
namespace ACME
{
    public partial class Statistics : Form
    {

        public Statistics()
        {
            InitializeComponent();
        }



        private void Statistics_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst(); 
            textBox2.Text = GetMenu.DLast();
            textBox5.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox8.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox4.Text = GetMenu.DFirst();
            textBox11.Text = DateTime.Now.ToString("yyyy");
            textBox12.Text = DateTime.Now.ToString("yyyy");
            textBox13.Text = DateTime.Now.ToString("yyyy");
            textBox14.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox15.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox19.Text = DateTime.Now.ToString("yyyy");
            textBox18.Text = DateTime.Now.ToString("yyyy") +"/"+ DateTime.Now.ToString("MM");
            comboBox3.Text = "SAP";
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetBU("StockPapare"), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Getwarehouse1(), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, GetMenu.Getwarehouse1(), "DataText", "DataText");
            UtilSimple.SetLookupBinding(comboBox5, GetMenu.GETEMP(), "DataText", "DataText");
            dataGridView2.Visible = false;


        }


        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
           

        }

        private System.Data.DataTable GetSAPSum1()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT '備貨單' 單據,COUNT(*) 單數,createname 製單人員 FROM acmesqlsp.dbo. WH_MAIN T0 ");
            sb.Append("              INNER JOIN  (select distinct shippingcode from acmesqlsp.dbo.WH_ITEM ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("              WHERE createname IS NOT NULL");
            sb.Append("              and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append("              GROUP BY createname ");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT '放貨單' 單據,COUNT(*) 單數,createname 製單人員 FROM acmesqlsp.dbo.WH_MAIN T0 ");
            sb.Append("              INNER JOIN  (select distinct shippingcode from acmesqlsp.dbo.WH_ITEM2 ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("              WHERE createname IS NOT NULL");
            sb.Append("              and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append("              GROUP BY createname ");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT '收貨單' 單據,COUNT(*) 單數,createname 製單人員 FROM acmesqlsp.dbo.WH_MAIN T0 ");
            sb.Append("              INNER JOIN  (select distinct shippingcode from acmesqlsp.dbo.WH_ITEM3 ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("              WHERE createname IS NOT NULL");
            sb.Append("              and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append("              GROUP BY createname ");
            sb.Append(" union all ");
            sb.Append(" SELECT '交貨單' 單據,SUM(S.單數) 單數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM acmesql02.dbo.ODLN T0");
            sb.Append("            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM acmesql02.dbo.ORDN T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind ) AS S ");
            sb.Append("              GROUP BY S.製單人員 ");
            sb.Append(" union all ");
            sb.Append("  SELECT '收貨採購單' 單據,SUM(S.單數) 單數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM Opdn T0");
            sb.Append("                            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ORPD T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM OPOR T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind )");
            sb.Append(" S GROUP BY S.製單人員");
            sb.Append(" union all ");
            sb.Append("  SELECT '庫存調整調撥' 單據,SUM(單數) 單數,S.製單人員 FROM    ( SELECT 'SAP' 總類,COUNT(*) 單數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM OWTR T0");
            sb.Append("                            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END  ");
            sb.Append("            UNION ALL       ");
            sb.Append("             SELECT 'SAP' 總類,COUNT(*) 單數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM oigN T0");
            sb.Append("                            where (substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') OR substring(t0.REF2,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管'))");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END,u_acme_kind");
            sb.Append("                           union all");
            sb.Append("                       SELECT 'SAP' 總類, COUNT(*)");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人 FROM oigE T0");
            sb.Append("                            where (substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') OR substring(t0.REF2,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管'))");
            sb.Append("                           and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4) END,u_acme_kind ) S ");
            sb.Append(" GROUP BY S.製單人員");
            sb.Append("            UNION ALL       ");
            sb.Append("      SELECT 'AR' 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM Oinv T0");
            sb.Append("                            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4)");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT 'AR貸項' 總類,COUNT(*) 單數 ");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM ORIN T0");
            sb.Append("                            where  Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4)");
            sb.Append("              UNION ALL");
            sb.Append("                        SELECT '採購退貨' 總類,COUNT(*) 單數  ");
            sb.Append("                            ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM ORPD T0 ");
            sb.Append("                               where  Convert(varchar(8),t0.taxdate,112) between @aa AND @bb  ");
            sb.Append("                            GROUP BY substring(T0.U_ACME_USER,0,4) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetOHEM(string pager)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select pager from acmesql02.dbo.ohem where jobtitle='船務倉管' AND pager=@pager ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@pager", pager));


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
        private System.Data.DataTable GetSAPRevenue11()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '備貨單' 單據,'備貨單' 總類,COUNT(*) 單數,substring(createname,1,3) 製單人員 FROM WH_MAIN T0 ");
            sb.Append(" INNER JOIN  (select distinct shippingcode from WH_ITEM ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE createname IS NOT NULL");
            sb.Append("               and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append(" GROUP BY createname ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '備貨單' 單據,'放貨單' 總類,COUNT(*) 單數,substring(createname,1,3) 製單人員 FROM WH_MAIN T0 ");
            sb.Append(" INNER JOIN  (select distinct shippingcode from WH_ITEM2 ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE createname IS NOT NULL");
            sb.Append("               and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append(" GROUP BY createname ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '備貨單' 單據,'收貨單' 總類,COUNT(*) 單數,substring(createname,1,3) 製單人員 FROM WH_MAIN T0 ");
            sb.Append(" INNER JOIN  (select distinct shippingcode from WH_ITEM3 ) T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE createname IS NOT NULL");
            sb.Append("               and substring(ntdollars,0,9) between @aa AND @bb ");
            sb.Append(" GROUP BY createname ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue2()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '交貨單' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ODLN T0");
            sb.Append("                            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ORDN T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind ) AS S ");
            sb.Append("              GROUP BY S.總類,S.製單人員 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue21(string ConnStrChi)
        {
            //合計 AS 銷售金額

            SqlConnection connection = new SqlConnection(ConnStrChi);
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT '交貨單' 單據,總類,SUM(單數) 單數,製單人員 FROM (    Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag=500 AND BillDate ");
            sb.Append("                        between @aa AND @bb AND Maker NOT LIKE '%[A-Z]%'  GROUP BY MAKER");
            sb.Append("                         UNION ALL");
            sb.Append("                        Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag=500 AND BillDate ");
            sb.Append("                        between @aa AND @bb AND Maker NOT LIKE '%[A-Z]%' GROUP BY MAKER ) S GROUP BY  S.總類,S.製單人員");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetChiSum(string ConnStrChi)
        {
            //合計 AS 銷售金額
          
            SqlConnection connection = new SqlConnection(ConnStrChi);
            StringBuilder sb = new StringBuilder();
            //2011
            sb.Append("              SELECT '交貨單正航' 單據,總類,SUM(單數) 單數,CASE 製單人員 WHEN 'ViviWeng' THEN '翁若婷' WHEN 'MillyGeng' THEN '耿玲玲' WHEN 'TONY' THEN '吳昭憲'  ELSE 製單人員 END 製單人員 FROM (  ");
             sb.Append("  Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag=500 AND BillDate ");
             sb.Append("                        between @aa AND @bb AND Maker NOT LIKE '%[A-Z]%' GROUP BY MAKER");
            sb.Append("                         UNION ALL");
            sb.Append("                        Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag=600 AND BillDate ");
            sb.Append("                        between @aa AND @bb AND Maker NOT LIKE '%[A-Z]%' GROUP BY MAKER ");
            sb.Append("                         ) S GROUP BY  S.總類,S.製單人員");
            sb.Append("    union all  ");
            sb.Append("           SELECT '收貨採購單正航' 單據,S.總類,SUM(S.單數) 單數,CASE S.製單人員 WHEN 'ViviWeng' THEN '翁若婷' WHEN 'MillyGeng' THEN '耿玲玲' WHEN 'TONY' THEN '吳昭憲'  ELSE S.製單人員 END 製單人員 FROM   (      ");
            sb.Append("                    Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag IN (100,200) AND BillDate ");
            sb.Append("                        between @aa AND @bb  GROUP BY MAKER            ");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 FROM OrdBillMain ");
            sb.Append("              WHERE FLAG IN (4)");
            sb.Append("           AND BillDate  between @aa AND @bb  GROUP BY MAKER  ");

            sb.Append("           )   S GROUP BY S.總類,S.製單人員");
            sb.Append(" union all ");
            sb.Append("            SELECT '庫存調整調撥正航' 單據,S.總類,SUM(單數) 單數,CASE S.製單人員 WHEN 'ViviWeng' THEN '翁若婷' WHEN 'MillyGeng' THEN '耿玲玲' WHEN 'TONY' THEN '吳昭憲'  ELSE S.製單人員 END 製單人員 FROM    (");
            sb.Append("           SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkAdjustMain");
            sb.Append("           WHERE AdjustDate between @aa AND @bb GROUP BY MAKER");
            sb.Append("           UNION ALL");
            sb.Append("           SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkMoveMain");
            sb.Append("           WHERE MoveDate between @aa AND @bb  GROUP BY MAKER");
            sb.Append("           UNION ALL");
            sb.Append("       SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkBorrowMain");
            sb.Append("           WHERE BorrowDate between @aa AND @bb GROUP BY MAKER ");
            sb.Append("         ) S  GROUP BY S.總類,S.製單人員");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue3()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '收貨採購單' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM Opdn T0");
            sb.Append("     where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') "); ;
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ORPD T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM OPOR T0");
            sb.Append("              where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind )");
            sb.Append(" S GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue31(string ConnStrChi)
        {
           
            SqlConnection connection = new SqlConnection(ConnStrChi);
            StringBuilder sb = new StringBuilder();
            sb.Append("           SELECT '收貨採購單' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM   (      ");
            sb.Append("                    Select CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 From comBillAccounts A   Where A.Flag IN (100,200) AND BillDate ");
            sb.Append("                        between @aa AND @bb  GROUP BY MAKER            ");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員 FROM OrdBillMain ");
            sb.Append("              WHERE FLAG IN (4)");
            sb.Append("           AND BillDate  between @aa AND @bb  GROUP BY MAKER             )");
            sb.Append("              S GROUP BY S.總類,S.製單人員");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("總類", typeof(string));
            dt.Columns.Add("單數", typeof(string));
            dt.Columns.Add("製單人員", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeSTOCK()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("總類", typeof(string));
            dt.Columns.Add("單數", typeof(string));
            dt.Columns.Add("製單人員", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableSUNNY()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("項目號碼", typeof(string));
            dt.Columns.Add("項目說明", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("庫存量", typeof(string));
            dt.Columns.Add("庫存值", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableAPPLE()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("項目群組", typeof(string));
            dt.Columns.Add("項目號碼", typeof(string));
            dt.Columns.Add("項目說明", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("庫存金額", typeof(string));
            //dt.Columns.Add("項目成本", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableS2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("年度", typeof(string));
            dt.Columns.Add("片數", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableS3()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("年度", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableSHIP()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("年", typeof(string));
            dt.Columns.Add("月", typeof(string));
            dt.Columns.Add("出貨筆數", typeof(string));
            dt.Columns.Add("出貨數量", typeof(string));
            dt.Columns.Add("出貨板數", typeof(string));
            return dt;
        }
        private System.Data.DataTable GetSAPRevenue4()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT '庫存調整調撥' 單據,S.總類,SUM(單數) 單數,S.製單人員 FROM    ( SELECT 'SAP' 總類,COUNT(*) 單數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM OWTR T0");
            sb.Append("                            where ( substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') OR substring(t0.REF2,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') )");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END  ");
            sb.Append("            UNION ALL       ");
            sb.Append("             SELECT 'SAP' 總類,COUNT(*) 單數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM oigN T0");
            sb.Append("                            where ( substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') OR substring(t0.REF2,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管'))");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END,u_acme_kind");
            sb.Append("                           union all");
            sb.Append("                       SELECT 'SAP' 總類, COUNT(*)");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人 FROM oigE T0");
            sb.Append("                            where ( substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') OR substring(t0.REF2,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管'))");
            sb.Append("                           and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4) END,u_acme_kind ) S ");
            sb.Append(" GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue41(string ConnStrChi)
        {
            //合計 AS 銷售金額

            SqlConnection connection = new SqlConnection(ConnStrChi);
            StringBuilder sb = new StringBuilder();


            sb.Append("            SELECT '庫存調整調撥' 單據,S.總類,SUM(單數) 單數,S.製單人員 FROM    (");
            sb.Append("           SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkAdjustMain");
            sb.Append("           WHERE AdjustDate between @aa AND @bb GROUP BY MAKER");
            sb.Append("           UNION ALL");
            sb.Append("           SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkMoveMain");
            sb.Append("           WHERE MoveDate between @aa AND @bb  GROUP BY MAKER");
            sb.Append("           UNION ALL");
            sb.Append("       SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 單數,MAKER 製單人員  FROM StkBorrowMain");
            sb.Append("           WHERE BorrowDate between @aa AND @bb GROUP BY MAKER ) S");
            sb.Append("           GROUP BY S.總類,S.製單人員");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue5()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT 'AR' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM (  SELECT 'SAP' 總類,COUNT(*) 單數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM Oinv T0");
            sb.Append("                            where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");

            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4) ) S");
            sb.Append(" GROUP BY S.總類,S.製單人員");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetSAPRevenue6()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("            SELECT 'AR' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM (  ");
            sb.Append("              SELECT 'SAP' 總類,COUNT(*) 單數 ");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM ORIN T0");
            sb.Append("                            where  Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4) ) S");
            sb.Append(" GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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

        private System.Data.DataTable GetSAPRevenue7()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("            SELECT '採購退貨' 單據,S.總類,SUM(S.單數) 單數,S.製單人員 FROM (  ");
            sb.Append("              SELECT 'SAP' 總類,COUNT(*) 單數 ");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM ORPD T0");
            sb.Append("                            where  Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4) ) S");
            sb.Append(" GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private void button3_Click(object sender, EventArgs e)
        {
            CalcTotals2();
        }
        private void CalcTotals2()
        {


            Int32 iTotal = 0;
            decimal iVatSum = 0;


            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["單數"].Value);


            }

            textBox3.Text = iTotal.ToString();




        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = MakeTableCombine();
            if (comboBox1.SelectedValue.ToString() == "1")
            {
                System.Data.DataTable dt = GetSAPRevenue11();
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }


            }
            else if (comboBox1.SelectedValue.ToString() == "2")
            {
                System.Data.DataTable dt = GetSAPRevenue2();
             
                DataRow dr = null;
                DataRow dr1 = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }
      
                       string ConnStrChi = "";
                       for (int s = 0; s <= 3; s++)
                       {
                           if (s == 0)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp20";
                           }
                           if (s == 1)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp16";
                           }
                           if (s == 2)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp21";
                           }
                           if (s == 3)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp22";
                           }
                           System.Data.DataTable dt1 = GetSAPRevenue21(ConnStrChi);
                           for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                           {
                               dr1 = dtCost.NewRow();
                               dr1["總類"] = dt1.Rows[i]["總類"].ToString();
                               dr1["單數"] = dt1.Rows[i]["單數"].ToString();
                               dr1["製單人員"] = dt1.Rows[i]["製單人員"].ToString();
                               dtCost.Rows.Add(dr1);
                           }
                       }

            }
            else if (comboBox1.SelectedValue.ToString() == "3")
            {
                System.Data.DataTable dt = GetSAPRevenue3();
       
                DataRow dr = null;
                DataRow dr1 = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }
                
                       string ConnStrChi = "";
                       for (int s = 0; s <= 3; s++)
                       {
                           if (s == 0)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp20";
                           }
                           if (s == 1)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp16";
                           }
                           if (s == 2)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp21";
                           }
                           if (s == 3)
                           {
                               ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp22";
                           }
                           System.Data.DataTable dt1 = GetSAPRevenue31(ConnStrChi);
                           for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                           {
                               dr1 = dtCost.NewRow();
                               dr1["總類"] = dt1.Rows[i]["總類"].ToString();
                               dr1["單數"] = dt1.Rows[i]["單數"].ToString();
                               dr1["製單人員"] = dt1.Rows[i]["製單人員"].ToString();
                               dtCost.Rows.Add(dr1);
                           }
                       }

            }
            else if (comboBox1.SelectedValue.ToString() == "4")
            {
                System.Data.DataTable dt = GetSAPRevenue4();
              
                DataRow dr = null;
                DataRow dr1 = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }
                string ConnStrChi = "";
                for (int s = 0; s <= 3; s++)
                {
                    if (s == 0)
                    {
                        ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp20";
                    }
                    if (s == 1)
                    {
                        ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp16";
                    }
                    if (s == 2)
                    {
                        ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp21";
                    }
                    if (s == 3)
                    {
                        ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp22";
                    }
                    System.Data.DataTable dt1 = GetSAPRevenue41(ConnStrChi);
                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        dr1 = dtCost.NewRow();
                        dr1["總類"] = dt1.Rows[i]["總類"].ToString();
                        dr1["單數"] = dt1.Rows[i]["單數"].ToString();
                        dr1["製單人員"] = dt1.Rows[i]["製單人員"].ToString();
                        dtCost.Rows.Add(dr1);
                    }
                }

            }
            else if (comboBox1.SelectedValue.ToString() == "5")
            {
                System.Data.DataTable dt = GetSAPRevenue5();
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }

            }
            else if (comboBox1.SelectedValue.ToString() == "6")
            {
                System.Data.DataTable dt = GetSAPRevenue6();
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }

            }
            else if (comboBox1.SelectedValue.ToString() == "7")
            {
                System.Data.DataTable dt = GetSAPRevenue7();
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["總類"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    dr["製單人員"] = dt.Rows[i]["製單人員"].ToString();
                    dtCost.Rows.Add(dr);
                }

            }
            else if (comboBox1.SelectedValue.ToString() == "8")
            {
                System.Data.DataTable dt = GetSAPSum1();
            
                DataRow dr = null;
                DataRow dr1 = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();
                    dr["總類"] = dt.Rows[i]["單據"].ToString();
                    dr["單數"] = dt.Rows[i]["單數"].ToString();
                    string MAKER = dt.Rows[i]["製單人員"].ToString().Trim();
                    dr["製單人員"] = MAKER;

                    System.Data.DataTable GT = GetOHEM(MAKER);
                    if (GT.Rows.Count > 0)
                    {
                        dtCost.Rows.Add(dr);
                    }
   
                }
                          string ConnStrChi = "";
                          for (int s = 0; s <= 3; s++)
                          {

                              if (s == 0)
                              {
                                  ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp20";
                              }
                              if (s == 1)
                              {
                                  ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp16";
                              }
                              if (s == 2)
                              {
                                  ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp21";
                              }
                              if (s == 3)
                              {
                                  ConnStrChi = "server=10.10.1.40;pwd=@cmewebstock;uid=webstock;database=CHIComp22";
                              }

                              System.Data.DataTable dt1 = GetChiSum(ConnStrChi);
                              for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                              {
                                  dr1 = dtCost.NewRow();
                                  dr1["總類"] = dt1.Rows[i]["單據"].ToString();
                                  dr1["單數"] = dt1.Rows[i]["單數"].ToString();
                                  string MAKER = dt1.Rows[i]["製單人員"].ToString().Trim();
                                  dr1["製單人員"] = MAKER;
                                  System.Data.DataTable GT = GetOHEM(MAKER);
                                  if (GT.Rows.Count > 0)
                                  {

                                      dtCost.Rows.Add(dr1);
                                  }
                              }
                          }
            }
            bindingSource1.DataSource = dtCost;
            dataGridView1.DataSource = bindingSource1.DataSource;
        }

    

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    dataGridView2.Visible = true; 
                    string da = dataGridView1.SelectedRows[0].Cells["總類"].Value.ToString();
                    string ds = dataGridView1.SelectedRows[0].Cells["製單人員"].Value.ToString();
                    if (da == "備貨單")
                    {
                        System.Data.DataTable dt = Gettriangle1(ds);
                        dataGridView2.DataSource = dt;
                    }
                    else if (da == "交貨單")
                    {
                        System.Data.DataTable dt = Gettriangle6(ds);
                        dataGridView2.DataSource = dt;
                    }
                  
                    else if (da == "收貨採購單")
                    {
                        System.Data.DataTable dt = Gettriangle4(ds);
                        dataGridView2.DataSource = dt;
                    }
                    else if (da == "AR")
                    {
                        System.Data.DataTable dt = Gettriangle5(ds);
                        dataGridView2.DataSource = dt;
                    }
                    else
                    {
                        dataGridView2.Visible = false;
                    }
                }
               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }
        private System.Data.DataTable Gettriangle1(string cc)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    select '備貨單' 總類, 貿易形式,sum(單數) 單數,substring(createname,0,4) 姓名 from (  select boardcountno 貿易形式,count(*) 單數,createname  from wh_main t0  inner join (select distinct shippingcode from wh_item ) t1 on (t0.shippingcode=t1.shippingcode)");
            sb.Append("                        where  createname like @cc and substring(ntdollars,0,9) between @aa AND @bb  ");
            sb.Append("                 group by boardcountno,createname");
            sb.Append(" union all");
            sb.Append("      select boardcountno 貿易形式,count(*) 單數,createname from wh_main t0  inner join (select distinct shippingcode from wh_item2 ) t1 on (t0.shippingcode=t1.shippingcode)");
            sb.Append("                        where  createname like @cc and substring(ntdollars,0,9) between @aa AND @bb  ");
            sb.Append("                 group by boardcountno,createname");
            sb.Append(" union all");
            sb.Append("      select boardcountno 貿易形式,count(*) 單數,createname from wh_main t0  inner join (select distinct shippingcode from wh_item3 ) t1 on (t0.shippingcode=t1.shippingcode)");
            sb.Append("                        where  createname like @cc and substring(ntdollars,0,9) between @aa AND @bb  ");
            sb.Append("                 group by boardcountno,createname ) as a  group by 貿易形式,createname");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@cc", "%" + cc + "%"));
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
 
        private System.Data.DataTable Gettriangle4(string cc)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                  SELECT '收貨採購單' 總類, boardcountno  貿易形式,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) 姓名");
            sb.Append("                         FROM Opdn T0");
            sb.Append("      left join acmesqlsp.dbo.shipping_main t1 on (t0.u_shipping_no=t1.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                  where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("                and substring(T0.U_ACME_USER,0,4)=@cc    and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY boardcountno,substring(T0.U_ACME_USER,0,4)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@cc", cc));
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
        private System.Data.DataTable Gettriangle5(string cc)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT 'AR' 總類,u_acme_workday 貿易形式,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) 姓名");
            sb.Append("   FROM Oinv T0");
            sb.Append(" left join (select distinct docentry,u_acme_workday from inv1) t1 on (t0.docentry=t1.docentry)");
            sb.Append("                                     where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("                    and substring(T0.U_ACME_USER,0,4)=@cc    and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                          GROUP BY u_acme_workday,substring(T0.U_ACME_USER,0,4)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@cc", cc));
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
        private System.Data.DataTable Gettriangle6(string cc)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT '交貨單' 總類,u_acme_workday 貿易形式,COUNT(*) 單數,substring(T0.U_ACME_USER,0,4) 姓名");
            sb.Append("   FROM odln T0");
            sb.Append(" left join (select distinct docentry,u_acme_workday from dln1) t1 on (t0.docentry=t1.docentry)");
            sb.Append("                                     where substring(t0.U_ACME_USER,1,3) in (select pager from acmesql02.dbo.ohem where jobtitle='船務倉管') ");
            sb.Append("                    and substring(T0.U_ACME_USER,0,4)=@cc    and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                          GROUP BY u_acme_workday,substring(T0.U_ACME_USER,0,4)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@cc", cc));
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

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

  

        private System.Data.DataTable Getwork()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT T0.SHIPPINGCODE JOBNO,boardCountNo 貿易形式,CAST(TT.QTY AS INT) 片數,createname 所有人, ");
            sb.Append("                  SUBSTRING(T0.SHIPPINGCODE,3,4)+'/'+SUBSTRING(T0.SHIPPINGCODE,7,2)+'/'+SUBSTRING(T0.SHIPPINGCODE,9,2)  起始日期, ");
            sb.Append("                CASE buCardcode WHEN 'Checked' then isnull(SUBSTRING(buCardname,1,4)+'/'+SUBSTRING(buCardname,5,2)+'/'+SUBSTRING(buCardname,7,2),'') else '' end 結案日期,shipping_obu 倉庫 FROM WH_MAIN T0 ");
            sb.Append("                  inner JOIN (SELECT SHIPPINGCODE,SUM(CASE QUANTITY WHEN '' THEN 0 ELSE CAST(ISNULL(QUANTITY,0) AS decimal) END ) QTY FROM WH_ITEM4  WHERE QUANTITY <>'一批' GROUP BY SHIPPINGCODE) TT ON (T0.SHIPPINGCODE=TT.SHIPPINGCODE) ");
            sb.Append("    WHERE 1=1 ");
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                sb.Append(" and  substring(T0.shippingcode,3,8) between '" + textBox4.Text.ToString() + "' and '" + textBox5.Text.ToString() + "' ");
            }
            if (textBox6.Text != "" && textBox7.Text != "")
            {
                sb.Append(" and  T0.SHIPPINGCODE between '" + textBox6.Text.ToString() + "' and '" + textBox7.Text.ToString() + "' ");
            }
            sb.Append(" ORDER BY T0.SHIPPINGCODE");
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
      
        private void button8_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\WHWork.xls";


         System.Data.DataTable  OrderData = Getwork();

         if (OrderData.Rows.Count > 0)
         {
             //Excel的樣版檔
             string ExcelTemplate = FileName;

             //輸出檔
             string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                   DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

             //產生 Excel Report
             ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
         }
         else
         {
             MessageBox.Show("無資料");
         }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ViewBatchPayment1();

           ExcelReport.GridViewToExcel(dataGridView3);
        }
        public  System.Data.DataTable GETPACKAGE()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.SHIPPINGCODE JOBNO,CARDCODE,CARDNAME,sayTotal PACKAGE,MEMO INVOICE,tradeCondition 貿易條件,mEMO3 倉管工單 FROM packingListM T0 LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  WHERE SUBSTRING(T0.SHIPPINGCODE,3,8) BETWEEN @DATE1 AND @DATE2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox14.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox15.Text));
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
        public System.Data.DataTable GETPACKAGE2S(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" Declare @name3 varchar(100) ");
            sb.Append(" SELECT  @name3 =SUBSTRING(COALESCE(@name3 + '/',''),0,99) + Shipping_OBU   FROM ");
            sb.Append(" (SELECT DISTINCT Shipping_OBU FROM  WH_MAIN WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ")) A");
            sb.Append(" SELECT @name3 WHNO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
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
        public System.Data.DataTable GETWHUSER()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDCODE 客戶編號,CARDMANE 客戶名稱,P1 內銷,P2 外銷,P3 三角 FROM WH_USER ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
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
        public System.Data.DataTable GETWHUSER2(string CARDCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT COUNT(*) DOC FROM ODLN WHERE CARDCODE=@CARDCODE AND YEAR(DOCDATE) BETWEEN @S1 AND @S2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@S1", textBox16.Text));
            command.Parameters.Add(new SqlParameter("@S2", textBox17.Text));
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
        public System.Data.DataTable GETPACKAGE2()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT U_SHIPPING_NO JOBNO,U_ACME_INV INVOICE,CARDCODE,CARDNAME FROM OPDN WHERE SUBSTRING(U_SHIPPING_NO,3,8) BETWEEN @DATE1 AND @DATE2  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox14.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox15.Text));
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

        public System.Data.DataTable GETPACKAGE3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT tradeCondition FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        private System.Data.DataTable MakeTableCombineF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("PACKAGE", typeof(string));
            dt.Columns.Add("倉管工單", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("倉別", typeof(string));
            dt.Columns.Add("INVOICE", typeof(string));
            return dt;
        }
        private void GerPare()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT T0.CARDNAME 客戶,ITEMCODE 產品編號,T1.QUANTITY 數量,WHNAME 倉庫 FROM WH_MAIN T0 ");
            sb.Append("   LEFT JOIN WH_ITEM T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("    WHERE SUBSTRING(T0.SHIPPINGCODE,3,8)='" + textBox5.Text.ToString() + "' ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView3.DataSource = bindingSource1;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            GerPare();
            ExcelReport.GridViewToExcel(dataGridView3);
        }



        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetSUNNY(comboBox2.SelectedValue.ToString());
            System.Data.DataTable dtCost = MakeTableSUNNY();
          
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                dr["項目號碼"] = Convert.ToString(dt.Rows[i]["項目號碼"]);
                dr["項目說明"] = Convert.ToString(dt.Rows[i]["項目說明"]);
                dr["庫存量"] = Convert.ToString(dt.Rows[i]["庫存量"]);
                dr["庫存值"] = Convert.ToString(dt.Rows[i]["庫存值"]);
                dr["料號"] = Convert.ToString(dt.Rows[i]["料號"]);

                dtCost.Rows.Add(dr);

            }
            dataGridView3.DataSource = dtCost;
            ExcelReport.GridViewToExcel(dataGridView3);
        }
        public static System.Data.DataTable GetSUNNY(string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
                       if (globals.DBNAME == "達睿生")
                {
            
                sb.Append("    SELECT T0.[ItemCode] 項目號碼, T0.[ItemName] 項目說明, cast(T1.[OnHand] as int) 庫存量,t1.avgprice*T1.[OnHand] 庫存值,T0.U_PARTNO 料號    FROM OITM T0  INNER JOIN OITW T1 ON T0.ItemCode = T1.ItemCode WHERE   T1.[OnHand] <>0    and  ISNULL(T0.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            }
            else
            {
                sb.Append("    SELECT T0.[ItemCode] 項目號碼, T0.[ItemName] 項目說明, cast(T1.[OnHand] as int) 庫存量,t0.avgprice*T1.[OnHand] 庫存值,T0.U_PARTNO 料號    FROM OITM T0  INNER JOIN OITW T1 ON T0.ItemCode = T1.ItemCode WHERE   T1.[OnHand] <>0    and  ISNULL(T0.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            }
                if (aa != "全部倉")
            {
                sb.Append(" AND  T1.[WhsCode]=@aa  ");
            }
             sb.Append(" ORDER BY T0.[ItemCode] ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@aa", aa));
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

        public static System.Data.DataTable GetSUNNY2(string aa,string bb)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

sb.Append(" SELECT SUBSTRING(T0.SHIPPINGCODE,7,2) 月,COUNT(*) 張  FROM WH_MAIN T0");
sb.Append(" INNER JOIN WH_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
sb.Append("  WHERE SHIPPING_OBU=@aa and SUBSTRING(T0.SHIPPINGCODE,3,4)=@bb ");

                sb.Append(" GROUP BY  SUBSTRING(T0.SHIPPINGCODE,7,2)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
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
        public static System.Data.DataTable GetAPPLE(string docdate)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MAX(substring(T2.itmsgrpNAM,4,15)) 項目群組,T0.[ItemCode] 項目號碼,MAX(T1.[ItemName]) 項目說明,W.WhsName 倉庫, SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,ROUND(MAX(t1.avgprice)*(SUM(T0.[InQty])-SUM(T0.[OutQty])),0) 庫存金額,");
            sb.Append(" T1.U_PARTNO PARTNO FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  (T1.[ItemCode] = T0.ItemCode )  ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T2  ON  T1.itmsgrpcod = T2.itmsgrpcod   ");
            sb.Append(" inner JOIN OWHS W on (T0.warehouse=W.whscode) ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  (T0.[docdate] >='2007.12.31'  and Convert(varchar(8),T0.[docdate],112)  <=@docdate) ");
            sb.Append("  AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組' ");
            sb.Append(" GROUP BY T0.[ItemCode],W.WhsName,T0.warehouse,T1.U_PARTNO");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0 )");
            sb.Append(" ORDER BY T0.[ItemCode],W.WhsName,T0.warehouse"); 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@docdate", docdate));
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
        public  System.Data.DataTable T2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT * FROM ACMESQLSP.DBO.WH_NOTRETURN2 WHERE TITLE=@TITLE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TITLE", comboBox3.Text));
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
        public System.Data.DataTable GETSHIP()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T0.SHIPPINGCODE 出貨筆數,T1.tradeCondition 貿易條件,");
            sb.Append(" SUBSTRING(T0.SHIPPINGCODE,3,4)年,SUBSTRING(T0.SHIPPINGCODE,7,2)月,");
            sb.Append(" SUM(T0.QUANTITY) 出貨數量,SUM(cast(sayTotal as int)) 出貨板數 FROM PackingListM T0");
            sb.Append(" LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(T0.SHIPPINGCODE,3,4)=@YEAR  AND SUBSTRING(CARDCODE,1,1) <> 'S' ");
            sb.Append("  GROUP BY  T0.SHIPPINGCODE,T1.tradeCondition");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", textBox11.Text));
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

        public System.Data.DataTable GETSHIP2(string SHIPNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPING_OBU FROM WH_MAIN WHERE boardCount LIKE '%" + SHIPNO + "%' ");


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
        public System.Data.DataTable GETEMP()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUBSTRING(convert(varchar,T0.DOCDATE, 112),1,6) 年月,T1.HOMETEL 員工,T0.CARDNAME 廠商,SUM(DOCTOTAL) 金額 FROM OPCH  T0");
            sb.Append(" LEFT JOIN OHEM T1 ON (T0.OWNERCODE=T1.EMPID)");
            sb.Append(" WHERE T1.HOMETEL=@HOMETEL AND YEAR(T0.DOCDATE) =@YEAR1");
            sb.Append(" GROUP BY SUBSTRING(convert(varchar,T0.DOCDATE, 112),1,6),T1.HOMETEL,T0.CARDNAME");
            sb.Append(" ORDER BY SUBSTRING(convert(varchar,T0.DOCDATE, 112),1,6)");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR1", textBox19.Text));
            command.Parameters.Add(new SqlParameter("@HOMETEL", comboBox5.Text));

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
        public System.Data.DataTable GETCAR()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE CARDCODE WHEN 'U0017' THEN '友福'  WHEN 'U0193' THEN '聯倉' WHEN 'U0447' THEN '新得利' END   倉庫,YEAR(DOCDATE) 年度,SUM([InQty])-SUM([OutQty]) 片數,case when SUM([InQty])-SUM([OutQty]) = 0 then 0 else  SUM(TransValue)  end 金額");
            sb.Append("  FROM OINM  WHERE  CARDCODE =('U0447')  AND ITEMCODE IN ('ZA0SZ0400','ZA0SB0005')");
            sb.Append(" AND YEAR(DOCDATE) BETWEEN @YEAR1 AND @YEAR2 ");
            sb.Append("               GROUP BY YEAR(DOCDATE),CARDCODE ");
            sb.Append("    UNION ALL ");
            sb.Append(" SELECT CASE CARDCODE WHEN 'U0017' THEN '友福'  WHEN 'U0193' THEN '聯倉' WHEN 'U0447' THEN '新得利' END   倉庫,YEAR(DOCDATE) 年度,SUM([InQty])-SUM([OutQty]) 片數,case when SUM([InQty])-SUM([OutQty]) = 0 then 0 else  SUM(TransValue)  end 金額 FROM OINM T0");
            sb.Append("  WHERE ITEMCODE='ZA0SB0005 ' AND CARDCODE IN( 'U0017','U0193','U0447') AND YEAR(DOCDATE) BETWEEN @YEAR1 AND @YEAR2");
            sb.Append(" GROUP BY YEAR(DOCDATE),CARDCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@YEAR2", textBox11.Text));
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
        public System.Data.DataTable GETCAR2(string WHSCODE, string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QUANTITY) 片數 FROM INV1 WHERE WHSCODE IN ('TW002','TW012','TW017')");
            sb.Append(" AND (CASE WHSCODE WHEN 'TW002' THEN '友福'  WHEN 'TW012' THEN '聯倉' WHEN 'TW017' THEN '新得利' END)=@WHSCODE  AND YEAR(DOCDATE) BETWEEN @YEAR1 AND @YEAR2");
            sb.Append(" AND YEAR(DOCDATE)=@DOCDATE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@YEAR1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@YEAR2", textBox11.Text));
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
        public void UPT2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE [WH_NOTRETURN2] SET COMMENT=@COMMENT where TITLE=@TITLE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@COMMENT", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@TITLE", comboBox3.Text));
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
        public static System.Data.DataTable GetAPPLE1(string ItemCode, string docdate)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[ItemCode]");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T2  ON  T1.itmsgrpcod = T2.itmsgrpcod   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  (T0.[docdate] >='2007.12.31' and Convert(varchar(8),T0.[docdate],112)  <=@docdate) ");
            sb.Append("  AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.[ItemCode] =@ItemCode ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));

            command.Parameters.Add(new SqlParameter("@docdate", docdate));
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
        public static System.Data.DataTable GetAPPLE2(string itemcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select USERTEXT 主要描述 from oitm where itemcode=@itemcode");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@itemcode", itemcode));

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
            System.Data.DataTable dt = GetAPPLE(textBox8.Text);
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
            System.Data.DataTable dtCost = MakeTableAPPLE();
            string 項目號碼="";
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                項目號碼 = Convert.ToString(dt.Rows[i]["項目號碼"]);
                dr = dtCost.NewRow();
                dt1 = GetAPPLE1(項目號碼,textBox8.Text);
                dt2 = GetAPPLE2(項目號碼);
                if (dt1.Rows.Count > 0)
                {
                    dr["項目群組"] = Convert.ToString(dt.Rows[i]["項目群組"]);
                    dr["項目號碼"] = 項目號碼;
                    dr["項目說明"] = Convert.ToString(dt.Rows[i]["項目說明"]);
                    dr["倉庫"] = Convert.ToString(dt.Rows[i]["倉庫"]);
                    dr["數量"] = Convert.ToString(dt.Rows[i]["數量"]);
                    dr["庫存金額"] = Convert.ToString(dt.Rows[i]["庫存金額"]);
                    dr["庫存金額"] = Convert.ToString(dt.Rows[i]["庫存金額"]);

                    dr["料號"] = Convert.ToString(dt.Rows[i]["PARTNO"]); ;
              
                    dtCost.Rows.Add(dr);
               }
            }
            dataGridView3.DataSource = dtCost;
            ExcelReport.GridViewToExcel(dataGridView3);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox9.Text != "" && comboBox3.Text != "")
            {
                UPT2();
                MessageBox.Show("儲存成功");
            }
            else
            {
                MessageBox.Show("請輸入資料");
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable TS = T2();

            textBox9.Text = TS.Rows[0][0].ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GETCAR();
            
            System.Data.DataTable dtCost = MakeTableS2();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string WHS = Convert.ToString(dt.Rows[i]["倉庫"]);
                string YEAR = Convert.ToString(dt.Rows[i]["年度"]);
                dr = dtCost.NewRow();
                dr["倉庫"] = WHS;
                dr["年度"] = YEAR;
                dr["金額"] = Convert.ToString(dt.Rows[i]["金額"]);

                System.Data.DataTable dt2 = GETCAR2(WHS, YEAR);
                if (dt2.Rows.Count > 0)
                {
                    dr["片數"] = Convert.ToString(dt2.Rows[0]["片數"]);
                }

                dtCost.Rows.Add(dr);
            }
            
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\wh\\卡車費.xls";



            if (dtCost.Rows.Count > 0)
            {
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutputS2(dtCost, ExcelTemplate, OutPutFile, "pivot");
            }
            else
            {
                MessageBox.Show("無資料");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GETSHIP();

            System.Data.DataTable dtCost = MakeTableSHIP();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string JOBNO = Convert.ToString(dt.Rows[i]["出貨筆數"]);
                dr = dtCost.NewRow();
                dr["出貨筆數"] = JOBNO;
                dr["貿易條件"] = Convert.ToString(dt.Rows[i]["貿易條件"]);
                dr["年"] = Convert.ToString(dt.Rows[i]["年"]);
                dr["月"] = Convert.ToString(dt.Rows[i]["月"]);
                dr["出貨數量"] = Convert.ToString(dt.Rows[i]["出貨數量"]);
                dr["出貨板數"] = Convert.ToString(dt.Rows[i]["出貨板數"]);
                System.Data.DataTable dt2 = GETSHIP2(JOBNO);
                if (dt2.Rows.Count > 0)
                {
                    dr["倉庫"] = Convert.ToString(dt2.Rows[0][0]);
                }
                dtCost.Rows.Add(dr);
        
           
            }
            dataGridView3.DataSource = dtCost;
            ExcelReport.GridViewToExcel(dataGridView3);
 
        }

        private void button13_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetSUNNY2(comboBox4.SelectedValue.ToString(),textBox13.Text);
            dataGridView4.DataSource = dt;
            ExcelReport.GridViewToExcel(dataGridView4);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = MakeTableCombineF();
            System.Data.DataTable dt = GETPACKAGE();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["JOBNO"] = dt.Rows[i]["JOBNO"].ToString();
                dr["客戶名稱"] = dt.Rows[i]["CARDNAME"].ToString();
                dr["PACKAGE"] = dt.Rows[i]["PACKAGE"].ToString();
                string WHNO = dt.Rows[i]["倉管工單"].ToString();

                string[] arrurl = WHNO.Split(new Char[] { ',' });
                StringBuilder sb = new StringBuilder();
                foreach (string ESi in arrurl)
                {
                    sb.Append("'" + ESi + "',");
                }
                sb.Remove(sb.Length - 1, 1);

                
                dr["倉管工單"] = WHNO;
                dr["貿易條件"] = dt.Rows[i]["貿易條件"].ToString();
                dr["INVOICE"] = dt.Rows[i]["INVOICE"].ToString();
                System.Data.DataTable G1 = GETPACKAGE2S(sb.ToString());
                if (G1.Rows.Count > 0)
                {

                    dr["倉別"] = G1.Rows[0][0].ToString();
                }
          

                
                dtCost.Rows.Add(dr);
            }

            System.Data.DataTable dt2 = GETPACKAGE2();
            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string JOBNO = dt2.Rows[i]["JOBNO"].ToString();
                dr["JOBNO"] = JOBNO;
                dr["PACKAGE"] = "";
           
                dr["客戶名稱"] = dt2.Rows[i]["CARDNAME"].ToString();
                dr["INVOICE"] = dt2.Rows[i]["INVOICE"].ToString();
                System.Data.DataTable dt3 = GETPACKAGE3(JOBNO);
                if (dt3.Rows.Count > 0)
                {
                    dr["貿易條件"] = dt3.Rows[0][0].ToString();
                }

                dtCost.Rows.Add(dr);
            }
            dataGridView3.DataSource = dtCost;
            ExcelReport.GridViewToExcel(dataGridView3);
        }

        private void ViewBatchPayment1()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.shippingcode 工單號碼,t0.cardname 客戶名稱,t1.dscription 品名規格,t1.quantity 數量,t0.createname 所有人  from wh_main t0");
            sb.Append(" left join wh_item t1 on(t0.shippingcode=t1.shippingcode)");
            sb.Append("  where  t1.shippingcode+isnull(cast(t1.docentry1 as nvarchar),'') not in ");
            sb.Append("  (select shippingcode+isnull(cast(basedoc as nvarchar),'') from wh_item2)");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("請確定SAP A01是否登出？", "YES/NO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Data.DataTable T1 = Getpdn1();
                for (int i = 0; i < T1.Rows.Count - 1; i++)
                {
                    string DOCENTRY = T1.Rows[i][0].ToString();
                    UpdateMasterSQL22("5", "5", DOCENTRY);
                    //UpdateMasterSQL22("5", "10", DOCENTRY);
                    //UpdateMasterSQL22("5", "11", DOCENTRY);
                }
                MessageBox.Show("已更新");
            }
        }
        private void UpdateMasterSQL22(string FID, string TID, string ID2)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @FormID nvarchar(20)");
            sb.Append(" Declare @From nvarchar(20)");
            sb.Append(" Declare @To nvarchar(20)");
            sb.Append(" set @FormID = " + ID2 + "");
            sb.Append(" set @From = " + FID + "");
            sb.Append(" set @To = " + TID + "");
            sb.Append(" if exists(Select 1 from ACMESQL98.DBO.OUSR where UserID=@To)");
            sb.Append(" begin");
            sb.Append(" Delete From  ACMESQL02.DBO.CPRF Where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@To");
            sb.Append(" Insert Into  ACMESQLSP.DBO.CPRF2 Select * From ACMESQLSP.DBO.CPRF where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@From");
            sb.Append(" update ACMESQLSP.DBO.CPRF2 set UserSign=@to");
            sb.Append(" Insert Into  ACMESQL02.DBO.CPRF Select * From ACMESQLSP.DBO.CPRF2");
            sb.Append(" truncate table ACMESQLSP.DBO.CPRF2 ");
            sb.Append(" end");
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

        public System.Data.DataTable Getpdn1()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY, FormName as datatext FROM FORMID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "FORMID");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["FORMID"];
        }


        public System.Data.DataTable GETOPDN(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(10),U_ACME_INVOICE,111) INVDATE ,CARDNAME  FROM OPDN WHERE DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "FORMID");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["FORMID"];
        }

        private void button16_Click(object sender, EventArgs e)
        {
            DELCUSTUSER();
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {

                    GD5(opdf.FileName);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GDOPDN(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);



                for (int i = 2; i <= iRowCnt; i++)
                {
                    string DINV;
                    string AUNINV;
                    string DOCENTRY;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    DINV = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    AUNINV = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    DOCENTRY = range.Text.ToString().Trim();
                    if (DINV.IndexOf("達") != -1 || AUNINV.IndexOf("達") != -1 || DINV.IndexOf("/") != -1 || AUNINV.IndexOf("/") != -1)
                    {
                        MessageBox.Show("格式錯誤" + DINV + " " + AUNINV);
                    }


                    if (DINV != "" && AUNINV != "")
                    {
                        UPOPEN(AUNINV, DINV, DOCENTRY);
                    }

                    System.Data.DataTable GET1 = GETOPDN(DOCENTRY);
                    if (GET1.Rows.Count > 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                        range.Select();
                        range.Value2 = GET1.Rows[0][0].ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                        range.Select();
                        range.Value2 = GET1.Rows[0][1].ToString();
                    }





                }
            }

            finally
            {

                //334499
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);
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
                MessageBox.Show("更新完成");
                System.Diagnostics.Process.Start(NewFileName);
            }


        }
        private void GDOPDN2(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);



                for (int i = 2; i <= iRowCnt; i++)
                {
                    string DINV;
                    string AUNINV;


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    DINV = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    AUNINV = range.Text.ToString().Trim();



                    if (DINV != "" && AUNINV != "")
                    {
                        UPOPEN2(AUNINV, DINV);
                    }

               

   



                }
            }
            
            finally
            {

                //334499
            
                try
                {
                    //excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                MessageBox.Show("更新完成");
                 //  System.Diagnostics.Process.Start(NewFileName);
            }


        }
        private void GD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);



                for (int i = 2; i <= iRowCnt; i++)
                {
                    string CARDCODE;
                    string CARDNAME;
                    string CARDTYPE;
                    string P1;
                    string P2;
                    string P3;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 22]);
                    range.Select();
                    CARDCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 23]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    P1 = range.Text.ToString().Trim().ToUpper();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    P2 = range.Text.ToString().Trim().ToUpper();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    P3 = range.Text.ToString().Trim().ToUpper();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    CARDTYPE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 28]);
                    range.Select();
                    range.Value2 = "交貨單數量";

                    System.Data.DataTable GG1 = GETWHUSER2(CARDCODE);
                    if (GG1.Rows.Count > 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 28]);
                        range.Select();
                        range.Value2 = GG1.Rows[0][0].ToString();
                    }

                    if (!String.IsNullOrEmpty(CARDNAME))
                    {

                        ADDCUSTUSER(CARDCODE, CARDNAME, P1, P2, P3, CARDTYPE);
                    }



                }
            }
            finally
            {
                string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);
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
                MessageBox.Show("更新完成");
             //   System.Diagnostics.Process.Start(NewFileName);
            }
 

        }
        public System.Data.DataTable GETS1(string DOCDATE, string 類別)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM (");
            sb.Append(" SELECT '航通TFT' 類別,CAST(CAST(MAX(T0.DOCTOTAL) AS INT) AS VARCHAR) 金額,CAST(CAST(SUM(T1.QUANTITY) AS INT) AS VARCHAR)  件數 FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE='U0224'");
            sb.Append(" AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0TE0105'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉進倉理貨費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND DSCRIPTION LIKE '%進倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉卡車費(桃園)' ,CAST(SUM(T1.GTOTAL) AS INT) GTOTAL,'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SB0005'");
            sb.Append("  GROUP BY T1.ITEMCODE,T1.DSCRIPTION");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉出倉理貨費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND DSCRIPTION LIKE '%出倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉倉租費' ,CAST((T1.GTOTAL) AS INT),CAST(ltrim(substring(REPLACE(T1.DSCRIPTION,'TFT-',''),CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))+1,CHARINDEX('坪', REPLACE(T1.DSCRIPTION,'TFT-',''))-CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))-1))*2 AS VARCHAR)  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉加班費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0701'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利倉租費' ,CAST((T1.GTOTAL) AS INT),ltrim(substring(REPLACE(T1.DSCRIPTION,'TFT-',''),CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))+1,LEN(REPLACE(T1.DSCRIPTION,'TFT-',''))-CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))-1)) FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利進出倉理貨費' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%進出倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利倉卡車費' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SB0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '卸裝櫃費用' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%櫃費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利加班費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0701'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉進倉理貨費--1板*120' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%120%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉出倉理貨費--1板*150' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%150%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉倉租費-400*1坪' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005' AND T1.DSCRIPTION LIKE '%400%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉倉租費-408*1坪' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(6),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SF0005' AND T1.DSCRIPTION LIKE '%408%'");
            sb.Append("  ) AS A WHERE 類別=@類別");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@類別", 類別));

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
        public System.Data.DataTable GETS1Y(string DOCDATE, string 類別)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM (");
            sb.Append(" SELECT '航通TFT' 類別,CAST(CAST(MAX(T0.DOCTOTAL) AS INT) AS VARCHAR) 金額,CAST(CAST(SUM(T1.QUANTITY) AS INT) AS VARCHAR)  件數 FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE='U0224'");
            sb.Append(" AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0TE0105'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉進倉理貨費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND DSCRIPTION LIKE '%進倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉卡車費(桃園)' ,CAST(SUM(T1.GTOTAL) AS INT) GTOTAL,'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SB0005'");
            sb.Append("  GROUP BY T1.ITEMCODE,T1.DSCRIPTION");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉出倉理貨費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND DSCRIPTION LIKE '%出倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉倉租費' ,CAST((T1.GTOTAL) AS INT),CAST(ltrim(substring(REPLACE(T1.DSCRIPTION,'TFT-',''),CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))+1,CHARINDEX('坪', REPLACE(T1.DSCRIPTION,'TFT-',''))-CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))-1))*2 AS VARCHAR)  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '聯揚倉加班費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0193'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0701'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利倉租費' ,CAST((T1.GTOTAL) AS INT),ltrim(substring(REPLACE(T1.DSCRIPTION,'TFT-',''),CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))+1,LEN(REPLACE(T1.DSCRIPTION,'TFT-',''))-CHARINDEX('-', REPLACE(T1.DSCRIPTION,'TFT-',''))-1)) FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利進出倉理貨費' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%進出倉理貨費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利倉卡車費' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SB0005'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '卸裝櫃費用' ,CAST((T1.GTOTAL) AS INT),''  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%櫃費%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '新得利加班費' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0447'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SZ0701'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉進倉理貨費--1板*120' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%120%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉出倉理貨費--1板*150' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SZ0400' AND T1.DSCRIPTION LIKE '%150%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉倉租費-400*1坪' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE AND ITEMCODE='ZA0SF0005' AND T1.DSCRIPTION LIKE '%400%'");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT '嘉里大榮倉倉租費-408*1坪' ,CAST(T1.GTOTAL AS INT),'' FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.CARDCODE='U0361'");
            sb.Append("  AND Convert(varchar(4),T0.DOCDATE,112) =@DOCDATE  AND ITEMCODE='ZA0SF0005' AND T1.DSCRIPTION LIKE '%408%'");
            sb.Append("  ) AS A WHERE 類別=@類別");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@類別", 類別));

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
        private System.Data.DataTable GGY1(string DOCYEAR,string cardcode) 
        {
            System.Data.DataTable dt;
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'01" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'01" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'01" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'01" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'02" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'02" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'02" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'02" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'03" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'03" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'03" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'03" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005')");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'04" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'04" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'04" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'04" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'05" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'05" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'05" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'05" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'06" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'06" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'06" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'06" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'07" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'07" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'07" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'07" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'08" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'08" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'08" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'08" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'09" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'09" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'09" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'09" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') "); 
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'10" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'10" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'10" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'10" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'11" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'11" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'11" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'11" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'12" + "' AND ITEMCODE='ZA0SZ0400' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'12" + "' AND ITEMCODE='ZA0SF0005' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'12" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105') ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT sum(GTOTAL) as 金額 FROM OPOR T0 LEFT JOIN  POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= @CARDCODE AND Convert(varchar(6),T0.DOCDATE,112) = @DOCDATE" + "+'12" + "' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105' OR ItemCode= 'ZA0SZ0400' OR ItemCode= 'ZA0SF0005') ");


            /*
             * SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201901' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201901' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201901' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201902' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201902' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201902' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201903' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201903' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201903' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201904' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201904' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201904' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201905' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201905' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201905' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201906' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201906' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201906' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201907' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201907' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201907' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201908' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201908' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201908' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201909' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201909' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201909' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201910' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201910' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201910' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201911' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201911' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201911' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201912' AND ITEMCODE='ZA0SZ0400'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201912' AND ITEMCODE='ZA0SF0005'
UNION ALL
SELECT sum(DocTotal) as 金額 FROM  [AcmeSql02].[dbo].OPOR T0 LEFT JOIN  [AcmeSql02].[dbo].POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE  T0.CARDCODE= 'U0193'AND Convert(varchar(6),T0.DOCDATE,112) = '201912' AND (ITEMCODE='ZA0SB0005' OR ItemCode= 'ZA0TE0105')


            */
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", cardcode));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCYEAR));
            

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
        private void GDS1(string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);

                DateTime D1 = Convert.ToDateTime(textBox18.Text + "/01");
                DateTime D2 = D1.AddMonths(-2);
                DateTime D3 = D1.AddMonths(-1);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 2]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + " " + D2.ToString("MM");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 5]);
                range.Select();
                range.Value2 = D3.ToString("yyyy") + " " + D3.ToString("MM");

                for (int i = 4; i <= 18; i++)
                {
                    string DOCTYPE;
         
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    DOCTYPE = range.Text.ToString().Trim();

                    System.Data.DataTable GG1 = GETS1(D2.ToString("yyyyMM"), DOCTYPE);
                    if (GG1.Rows.Count > 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                        range.Select();
                        range.Value2 = GG1.Rows[0][1].ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                        range.Select();
                        range.Value2 = GG1.Rows[0][2].ToString();
                    }


                    System.Data.DataTable GG2 = GETS1(D3.ToString("yyyyMM"), DOCTYPE);
                    if (GG2.Rows.Count > 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                        range.Select();
                        range.Value2 = GG2.Rows[0][1].ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                        range.Select();
                        range.Value2 = GG2.Rows[0][2].ToString();
                    }

                }
            }
            finally
            {

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);
            }

                        
         //   

        }
        private void GDS1Y(string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;

            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);
            Microsoft.Office.Interop.Excel.Range range = null;
            Microsoft.Office.Interop.Excel.Range range2 = null;



            //Open the worksheet file


            try
            {
                object SelectCell = "A1";
                range = excelSheet.get_Range("A1", "M10");
                DateTime D1 = Convert.ToDateTime(textBox18.Text + "/01");
                DateTime D2 = D1.AddYears(-2);
                DateTime D3 = D1.AddYears(-1);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年聯倉費用含快遞運費";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 5]);
                range.Select();
                range.Value2 = D3.ToString("yyyy") + "年聯倉費用含快遞運費";


                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);
                /*
                D1 = Convert.ToDateTime(textBox18.Text + "/01");
                D2 = D1.AddYears(-2);
                D3 = D1.AddYears(-1);
                range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range2.Select();
                range2.Value2 = D2.ToString("yyyy") + "年聯倉費用含快遞運費";

                range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, 5]);
                range2.Select();
                range2.Value2 = D3.ToString("yyyy") + "年聯倉費用含快遞運費";
                */

                System.Data.DataTable GG1 = GGY1(D2.ToString("yyyy"), "U0193");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }

                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i][0].ToString() == "" ? "0" : GG1.Rows[4 * i][0].ToString();
                    //聯倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 1][0].ToString();
                    //聯倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 3][0].ToString();
                }

                System.Data.DataTable GG2 = GGY1(D3.ToString("yyyy"), "U0193");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }

                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[9, i + 2]);
                    range.Select();
                    range.Value2 = GG2.Rows[4 * i][0].ToString() == "" ? "0" : GG2.Rows[4 * i][0].ToString();
                    //聯倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, i + 2]);
                    range.Select();
                    range.Value2 = GG2.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG2.Rows[4 * i + 1][0].ToString();
                    //聯倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, i + 2]);
                    range.Select();
                    range.Value2 = GG2.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG2.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[12, i + 2]);
                    range.Select();
                    range.Value2 = GG2.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG2.Rows[4 * i + 3][0].ToString();
                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年新倉費用含快遞運費";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[20, 5]);
                range.Select();
                range.Value2 = D3.ToString("yyyy") + "年新倉費用含快遞運費";


                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);
                /*
                D1 = Convert.ToDateTime(textBox18.Text + "/01");
                D2 = D1.AddYears(-2);
                D3 = D1.AddYears(-1);
                range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range2.Select();
                range2.Value2 = D2.ToString("yyyy") + "年聯倉費用含快遞運費";

                range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, 5]);
                range2.Select();
                range2.Value2 = D3.ToString("yyyy") + "年聯倉費用含快遞運費";
                */

                System.Data.DataTable GG3 = GGY1(D2.ToString("yyyy"), "U0447");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }
                    excelSheet.Activate();
                    //新倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[16, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i][0].ToString() == "" ? "0" : GG3.Rows[4 * i][0].ToString();
                    //新倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[17, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 1][0].ToString();
                    //新倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[18, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[19, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 3][0].ToString();
                }

                System.Data.DataTable GG4 = GGY1(D3.ToString("yyyy"), "U0447");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }

                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22, i + 2]);
                    range.Select();
                    range.Value2 = GG4.Rows[4 * i][0].ToString() == "" ? "0" : GG4.Rows[4 * i][0].ToString();
                    //聯倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[23, i + 2]);
                    range.Select();
                    range.Value2 = GG4.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG4.Rows[4 * i + 1][0].ToString();
                    //聯倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[24, i + 2]);
                    range.Select();
                    range.Value2 = GG4.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG4.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[25, i + 2]);
                    range.Select();
                    range.Value2 = GG4.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG4.Rows[4 * i + 3][0].ToString();
                }


            }
            catch (Exception ex)
            {

            }
            finally
            {

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);
            }


            //   

        }
        private void GDM12(string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;

            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);
            Microsoft.Office.Interop.Excel.Range range = null;
            Microsoft.Office.Interop.Excel.Range range2 = null;



            //Open the worksheet file


            try
            {
                object SelectCell = "A1";
                range = excelSheet.get_Range("A1", "M10");
                DateTime D1 = Convert.ToDateTime(textBox18.Text + "/01");
                DateTime D2 = D1;
                int Month = DateTime.Now.Month;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年聯倉費用含快遞運費";




                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);

                System.Data.DataTable GG1 = GGY1(D2.ToString("yyyy"), "U0193");
                Decimal Total1 = 0;//聯倉月總
                for (int i = 0; i < Month; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }

                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i][0].ToString() == "" ? "0" : GG1.Rows[4 * i][0].ToString();
                    //聯倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 1][0].ToString();
                    //聯倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 3][0].ToString();
                    Total1 += GG1.Rows[4 * i + 3][0].ToString() == "" ? 0 : decimal.Parse(GG1.Rows[4 * i + 3][0].ToString());
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, 14]);
                range.Select();
                range.Value2 = Total1;//表格6N總數

               
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年新倉費用含快遞運費";



                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);


                System.Data.DataTable GG3 = GGY1(D2.ToString("yyyy"), "U0447");
                Decimal Total2 = 0;//新倉月總
                for (int i = 0; i < Month; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }
                    excelSheet.Activate();
                    //新倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i][0].ToString() == "" ? "0" : GG3.Rows[4 * i][0].ToString();
                    //新倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 1][0].ToString();
                    //新倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[12, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 3][0].ToString();
                    Total2 += GG3.Rows[4 * i + 3][0].ToString() == "" ? 0 : decimal.Parse(GG3.Rows[4 * i + 3][0].ToString());
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 14]);
                range.Select();
                range.Value2 = Total2;//表格13N總數




            }
            catch (Exception ex)
            {

            }
            finally
            {

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);
            }


            //   

        }
        private void GDS12(string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;

            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt); 
            Microsoft.Office.Interop.Excel.Range range = null;
            Microsoft.Office.Interop.Excel.Range range2 = null;



            //Open the worksheet file


            try
            {
                object SelectCell = "A1";
                range = excelSheet.get_Range("A1", "M10");
                DateTime D1 = Convert.ToDateTime(textBox18.Text + "/01");
                DateTime D2 = D1;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年聯倉費用含快遞運費";

  

                
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);
        
                System.Data.DataTable GG1 = GGY1(D2.ToString("yyyy"), "U0193");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }
                    
                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i][0].ToString() == "" ? "0" : GG1.Rows[4 * i][0].ToString();
                    //聯倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 1][0].ToString();
                    //聯倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, i + 2]);
                    range.Select();
                    range.Value2 = GG1.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG1.Rows[4 * i + 3][0].ToString();
                }


                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                excelSheet.Activate();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range.Select();
                range.Value2 = D2.ToString("yyyy") + "年新倉費用含快遞運費";

         

                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);
          

                System.Data.DataTable GG3 = GGY1(D2.ToString("yyyy"), "U0447");
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i).ToString();
                    }
                    else
                    {
                        monthcode = (i).ToString();
                    }
                    excelSheet.Activate();
                    //新倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i][0].ToString() == "" ? "0" : GG3.Rows[4 * i][0].ToString();
                    //新倉i月倉租
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 1][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 1][0].ToString();
                    //新倉i月運費
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 2][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 2][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, i + 2]);
                    range.Select();
                    range.Value2 = GG3.Rows[4 * i + 3][0].ToString() == "" ? "0" : GG3.Rows[4 * i + 3][0].ToString();
                }

   


            }
            catch (Exception ex) 
            {

            }
            finally
            {

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);
            }


            //   

        }
        public void UPOPEN(string U_AUOINV, string U_ACME_INV, string DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString );
            SqlCommand command = new SqlCommand("UPDATE OPDN SET  U_AUOINV=@U_AUOINV,U_ACME_INV=@U_ACME_INV WHERE DOCENTRY=@DOCENTRY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUOINV", U_AUOINV));

            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
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

        public void UPOPEN2(string U_AUOINV, string U_ACME_INV)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OPDN SET  U_AUOINV=@U_AUOINV WHERE U_ACME_INV=@U_ACME_INV", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUOINV", U_AUOINV));

            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));

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
        public void ADDCUSTUSER(string CARDCODE, string CARDMANE, string P1, string P2, string P3, string CARDTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_USER(CARDCODE,CARDMANE,P1,P2,P3,CARDTYPE) values(@CARDCODE,@CARDMANE,@P1,@P2,@P3,@CARDTYPE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDMANE", CARDMANE));
            command.Parameters.Add(new SqlParameter("@P1", P1));
            command.Parameters.Add(new SqlParameter("@P2", P2));
            command.Parameters.Add(new SqlParameter("@P3", P3));
            command.Parameters.Add(new SqlParameter("@CARDTYPE", CARDTYPE));
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
        public void DELCUSTUSER()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE WH_USER ", connection);
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

        private void button17_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = GETEMP();




            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\wh\\費用分析.xls";



            if (dtCost.Rows.Count > 0)
            {
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutputS2(dtCost, ExcelTemplate, OutPutFile, "pivot");
            }
            else
            {
                MessageBox.Show("無資料");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\wh\\TFT後勤維運部倉管組月度費用記錄表.xlsx";


                string OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                GDM12(FileName, OutPutFile);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
//            try
//            {
//                string FileName = string.Empty;
//                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
     
//                FileName = lsAppDir + "\\Excel\\wh\\費用記錄表.xlsx";


//                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
//DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

//                GDS1(FileName,OutPutFile);

                
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.Message);
//            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\wh\\TFT後勤維運部倉管組年度費用記錄表.xlsx";


                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                GDS1Y(FileName, OutPutFile);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
          
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {

                    GDOPDN2(opdf.FileName);


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

     
    }
}
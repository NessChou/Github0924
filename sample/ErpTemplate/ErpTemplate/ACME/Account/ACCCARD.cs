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
    public partial class ACCCARD : Form
    {
        
        System.Data.DataTable dtAD = null;

     
        public ACCCARD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Int64  Q1 = 0;
            Int64 Q2 = 0;
            Int64 Q3 = 0;
            Int64 Q4 = 0;
            Int64 Q5 = 0;
            dtAD = MakeTableCombine();
            System.Data.DataTable F1 = GETOPCH();
            DataRow dr = null;
            for (int i = 0; i <= F1.Rows.Count - 1; i++)
            {

                dr = dtAD.NewRow();
                dr["排行"] = F1.Rows[i]["排行"].ToString();
                dr["廠商編號"] = F1.Rows[i]["廠商編號"].ToString();
                dr["廠商名稱"] = F1.Rows[i]["廠商名稱"].ToString();
                dr["總數量"] = Convert.ToInt32(F1.Rows[i]["總數量"]);
                dr["總進貨金額(未稅)"] = Convert.ToInt64(F1.Rows[i]["總進貨金額"]);
                dr["總退貨金額(未稅)"] = Convert.ToInt64(F1.Rows[i]["總退貨金額"]);

                dr["進貨折讓金額"] = Convert.ToInt64(F1.Rows[i]["進貨折讓金額"]);

                dr["總進貨淨額"] = Convert.ToInt64(F1.Rows[i]["總進貨淨額"]);
                Q1 += Convert.ToInt64(F1.Rows[i]["總數量"]);
                Q2 += Convert.ToInt64(F1.Rows[i]["總進貨金額"]);
                Q3 += Convert.ToInt64(F1.Rows[i]["總退貨金額"]);
                Q4 += Convert.ToInt64(F1.Rows[i]["進貨折讓金額"]);
                Q5 += Convert.ToInt64(F1.Rows[i]["總進貨淨額"]);

                string N2 = GETOPCHT().Rows[0][0].ToString();
                string N1 = F1.Rows[i]["總進貨淨額"].ToString();
                if (N2 == "0")
                {
                    dr["佔比"] = "";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(N1) / Convert.ToDecimal(N2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["佔比"] = G;
                }

                dtAD.Rows.Add(dr);
            }

                //dr = dtAD.NewRow();
                //dr["排行"] = "1";
                //dr["廠商編號"] = "S0001";
                //dr["廠商名稱"] = "友達光電股份有限公司";
                //dr["總數量"] = Q1;
                //dr["總進貨金額(未稅)"] = Q2;
                //dr["總退貨金額(未稅)"] = Q3;

                //dr["進貨折讓金額"] = Q4;

                //dr["總進貨淨額"] = Q5;
                ////string QQ2 = GETOPCHT2().Rows[0][0].ToString();
                ////string QQ1 = F2.Rows[0]["總進貨淨額"].ToString();
                ////if (QQ2 == "0")
                ////{
                ////    dr["佔比"] = "";
                ////}
                ////else
                ////{
                ////    string G = Math.Round((Convert.ToDecimal(QQ1) / Convert.ToDecimal(QQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                ////    dr["佔比"] = G;
                ////}

                //dtAD.Rows.Add(dr);
     



          


            dataGridView1.DataSource = dtAD;

            //for (int i = 5; i <= 12; i++)
            //{
            //    DataGridViewColumn col = dataGridView1.Columns[i];


            //    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            //    col.DefaultCellStyle.Format = "#,##0";


            //}



        }

        private System.Data.DataTable MakeTableCombine()
        {
            //排行	廠商編號	廠商名稱	 總數量 	 總進貨金額(未稅)A 	總退貨金額(未稅)B	進貨折讓金額C	總進貨淨額D	佔比
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("排行", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("總數量", typeof(Int64));
            dt.Columns.Add("總進貨金額(未稅)", typeof(Int64));
            dt.Columns.Add("總退貨金額(未稅)", typeof(Int64));
            dt.Columns.Add("進貨折讓金額", typeof(Int64));
            dt.Columns.Add("總進貨淨額", typeof(Int64));
            dt.Columns.Add("佔比", typeof(string));

            return dt;
        }
        public System.Data.DataTable GETOPCH()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT    ROW_NUMBER() OVER (ORDER BY SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0) DESC ) 排行,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,SUM(T1.QUANTITY) 總數量,SUM(T1.LineTotal) 總進貨金額");
            sb.Append(" ,ISNULL(MAX(T2.LineTotal),0) 總退貨金額,ISNULL(MAX(T3.LineTotal),0) 進貨折讓金額,");
            sb.Append(" SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0) 總進貨淨額 FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='I' GROUP BY T0.CARDCODE) T2 ON (T0.CARDCODE =T2.CARDCODE)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='S' GROUP BY T0.CARDCODE) T3 ON (T0.CARDCODE =T3.CARDCODE)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
          //  sb.Append(" AND T0.CardCode LIKE '%S0001%' GROUP BY T0.CARDCODE,T0.CARDNAME");
            sb.Append(" AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE<>'S0001' GROUP BY T0.CARDCODE,T0.CARDNAME");
            sb.Append(" ORDER BY SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0) DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GETOPCHT()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0),0) 總進貨淨額 FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='I' GROUP BY T0.CARDCODE) T2 ON (T0.CARDCODE =T2.CARDCODE)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='S' GROUP BY T0.CARDCODE) T3 ON (T0.CARDCODE =T3.CARDCODE)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
         //   sb.Append(" AND T0.CardCode LIKE '%S0001%'");
            sb.Append(" AND T0.CardCode LIKE '%S%'");
            sb.Append(" ORDER BY SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0) DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
     
        public System.Data.DataTable GETOPCHT2()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
 
            sb.Append(" SELECT ISNULL(SUM(T1.LineTotal) -ISNULL(MAX(T2.LineTotal),0) -ISNULL(MAX(T3.LineTotal),0),0) 總進貨淨額 FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='I' GROUP BY T0.CARDCODE) T2 ON (T0.CARDCODE =T2.CARDCODE)");
            sb.Append(" LEFT JOIN (  SELECT T0.CARDCODE,SUM(T1.LineTotal) LineTotal FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" AND T0.DocType ='S' GROUP BY T0.CARDCODE) T3 ON (T0.CARDCODE =T3.CARDCODE)");
            sb.Append(" WHERE   Convert(varchar(8),T0.[DOCDate],112) BETWEEN @DATE1 AND @DATE2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
 
     
        private void ACCADLAB_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
     
        }
    }
}

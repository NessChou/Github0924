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

namespace ACME
{
    public partial class APMODEL : Form
    {
        System.Data.DataTable dtCost = null;
        public string cs;
        public APMODEL()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }
        private System.Data.DataTable Get1(string cs)
        {
        
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'ACME' 公司,  T1.ITEMCODE 產品編號,CardName 客戶,T1.Currency 幣別, AVG(T1.PRICE) 平均單價,SUM(T1.QUANTITY) 數量,'' 備註  FROM ACMESQL02.DBO.ORDR T0  ");
            sb.Append(" INNER JOIN ACMESQL02.DBO.RDR1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" left JOIN ACMESQL02.DBO.OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append(" WHERE  T1.PRICE <> 0 ");
            sb.Append(" AND  T1.[ItemCode] in ( " + cs + ")  ");
            sb.Append(" AND  T0.[DOCDATE] BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" GROUP BY T1.ITEMCODE,CardName,T1.Currency");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'DRS' 公司,  T1.ITEMCODE 產品編號,CardName 客戶,T1.Currency 幣別, AVG(T1.PRICE) 平均單價,SUM(T1.QUANTITY) 數量,'' 備註  FROM ACMESQL05.DBO.ORDR T0  ");
            sb.Append(" INNER JOIN ACMESQL05.DBO.RDR1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" left JOIN ACMESQL05.DBO.OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append(" WHERE T1.PRICE <> 0 ");
            sb.Append(" AND  T1.[ItemCode] in ( " + cs + ")  ");
            sb.Append(" AND  T0.[DOCDATE] BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" GROUP BY T1.ITEMCODE,CardName,T1.Currency");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
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

        private System.Data.DataTable Get2(string cs)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select   'CHOICE' 公司,T1.ProdID 產品編號,U.ShortName 客戶,CURRENCYNAME 幣別 ");
            sb.Append(" , AVG(T1.Price) 平均單價,SUM(T1.QUANTITY) 數量,'' 備註  from otherDB.CHIComp21.DBO.ordBillMain T0         ");
            sb.Append(" left join otherDB.CHIComp21.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join otherDB.CHIComp21.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append("  LEFT JOIN otherDB.CHIComp21.DBO.comCurrencySys T2 ON    (T0.CURRID=T2.CurrencyID) ");
            sb.Append(" WHERE  T0.Flag =2  AND T1.PRICE <> 0  ");
            sb.Append(" AND  T1.ProdID in ( " + cs + ")  ");
            sb.Append(" AND  T0.[BILLDATE] BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" GROUP BY T1.ProdID,U.ShortName,CURRENCYNAME");
            sb.Append(" UNION ALL");
            sb.Append(" select   'IPGI' 公司,T1.ProdID 產品編號,U.ShortName 客戶,CURRENCYNAME 幣別 ");
            sb.Append(" , AVG(T1.Price) 平均單價,SUM(T1.QUANTITY) 數量,'' 備註  from otherDB.CHIComp22.DBO.ordBillMain T0         ");
            sb.Append(" left join otherDB.CHIComp22.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append("  LEFT JOIN otherDB.CHIComp22.DBO.comCurrencySys T2 ON    (T0.CURRID=T2.CurrencyID) ");
            sb.Append(" WHERE  T0.Flag =2  AND T1.PRICE <> 0  ");
            sb.Append(" AND  T1.ProdID in ( " + cs + ")  ");
            sb.Append(" AND  T0.[BILLDATE] BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" GROUP BY T1.ProdID,U.ShortName,CURRENCYNAME");
            sb.Append(" UNION ALL");
            sb.Append(" select   '禾中' 公司,T1.ProdID 產品編號,U.ShortName 客戶,CURRENCYNAME 幣別 ");
            sb.Append(" , AVG(T1.Price) 平均單價,SUM(T1.QUANTITY) 數量,'' 備註  from otherDB.CHIComp23.DBO.ordBillMain T0         ");
            sb.Append(" left join otherDB.CHIComp23.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join otherDB.CHIComp23.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append("  LEFT JOIN otherDB.CHIComp23.DBO.comCurrencySys T2 ON    (T0.CURRID=T2.CurrencyID) ");
            sb.Append(" WHERE  T0.Flag =2  AND T1.PRICE <> 0  ");
            sb.Append(" AND  T1.ProdID in ( " + cs + ")  ");
            sb.Append(" AND  T0.[BILLDATE] BETWEEN @DATE1 AND @DATE2 ");
            sb.Append(" GROUP BY T1.ProdID,U.ShortName,CURRENCYNAME");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
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
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("幣別", typeof(string));
            dt.Columns.Add("平均單價", typeof(Decimal));
            dt.Columns.Add("數量", typeof(Int32));
            dt.Columns.Add("備註", typeof(string));


            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {

            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

                cs = frm1.q;

             

                if (!String.IsNullOrEmpty(cs))
                {
                    System.Data.DataTable dt = Get1(cs);
                    System.Data.DataTable dt2 = Get2(cs);
                    dtCost = MakeTableCombine();

                    DataRow dr = null;

                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        dr = dtCost.NewRow();

                        dr["公司"] = dt.Rows[i]["公司"].ToString();
                        dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                        dr["客戶"] = dt.Rows[i]["客戶"].ToString();
                        dr["幣別"] = dt.Rows[i]["幣別"].ToString();
                        dr["平均單價"] = Convert.ToDecimal(dt.Rows[i]["平均單價"]);
                        dr["數量"] = Convert.ToInt32(dt.Rows[i]["數量"]);
                        dr["備註"] = dt.Rows[i]["備註"].ToString();
                        dtCost.Rows.Add(dr);
                    }

                    for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                    {
                        dr = dtCost.NewRow();

                        dr["公司"] = dt2.Rows[i]["公司"].ToString();
                        dr["產品編號"] = dt2.Rows[i]["產品編號"].ToString();
                        dr["客戶"] = dt2.Rows[i]["客戶"].ToString();
                        dr["幣別"] = dt2.Rows[i]["幣別"].ToString();
                        dr["平均單價"] = Convert.ToDecimal(dt2.Rows[i]["平均單價"]);
                        dr["數量"] = Convert.ToInt32(dt2.Rows[i]["數量"]);
                        dr["備註"] = dt2.Rows[i]["備註"].ToString();
                        dtCost.Rows.Add(dr);
                    }
                    dataGridView1.DataSource = dtCost;

                }
            }
        }

        private void APMODEL_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }
    }
}

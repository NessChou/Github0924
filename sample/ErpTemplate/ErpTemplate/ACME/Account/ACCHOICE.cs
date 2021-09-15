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
    public partial class ACCCHOICE : Form
    {
        DataRow dr = null;
        int QQ1 = 0;
        int QQ2 = 0;
        int QQ3 = 0;
        int QQ4 = 0;

        int QQ12 = 0;
        int QQ22 = 0;
        int QQ32 = 0;
        int QQ42 = 0;

        int Q1 = 0;
        int Q2 = 0;
        int Q3 = 0;
        int Q4 = 0;
        int QQQ1 = 0;
        int QQQ2 = 0;
        int QQQ3 = 0;
        int QQQ4 = 0;

        int QQQ1F = 0;
        int QQQ2F = 0;
        int QQQ3F = 0;
        int QQQ4F = 0;
     
        string str16 = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        System.Data.DataTable dtAD = null;
        string strCn = "";
        public ACCCHOICE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            System.Data.DataTable dt = GetCHO4();
            dataGridView1.DataSource = dt;

            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 3; j <= 6; j++)
                {
                    Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);

                }
            }

            DataRow row;

            row = dt.NewRow();

     

            for (int j = 3; j <=6; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);

            for (int i = 3; i <= 7; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                if (i == 7)
                {
                    col.DefaultCellStyle.Format = "#,##0.00";
                }
                else
                {
                    col.DefaultCellStyle.Format = "#,##0";
                }


            }
        }
        public System.Data.DataTable GetCHO4()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 排行,T.CustID 客戶代碼,CAST(U.FullName AS nvarchar)   客戶名稱,SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END)  數量");
            sb.Append(" ,SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額");
            sb.Append(" ,SUM(CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 成本 ,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  毛利, ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(A.MLAmount-A.CostForAcc)/SUM(A.MLAmount) END)*100 毛利率          ");
            sb.Append(" From ComProdRec A           ");
            sb.Append(" left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END");
            sb.Append(" left join comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (comboBox4.Text == "博豐")
            {
                if (comboBox2.Text == "銷售")
                {
                    sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)<>'R0'  ");
                }
                if (comboBox2.Text == "維修")
                {
                    sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)='R0'  ");
                }
            }
            sb.Append("GROUP BY T.CustID,U.FullName ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  DESC");
       

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetCHOD1ARES()
        {
            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號");
            sb.Append(" From DBO.ComProdRec A             ");
            sb.Append(" left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag             ");
            sb.Append(" Where A.Flag IN (500,600)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (comboBox1.Text == "銷售")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)<>'R0'  ");
            }
            if (comboBox1.Text == "維修")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)='R0'  ");
            }
            sb.Append(" GROUP BY T.CustID ");
            sb.Append(" ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
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
        public System.Data.DataTable GetCHOD1(string DTYPE)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A             ");
            sb.Append(" left join otherDB.CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END                        ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");

            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (DTYPE == "PVM")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4)='2101' )  ");
            }
            else if (DTYPE == "PVI")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4)='2102' )  ");
            }
            else if (DTYPE == "OTH")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121') )  ");
            }
            else
            {
                sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE=@DTYPE AND USERS=@USERS )");
            }
            sb.Append(" GROUP BY T.CustID ");
            sb.Append(" ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
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
        public System.Data.DataTable GetCHOD2(string DTYPE, string CUSTID)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號,CAST(U.SHORTNAME AS nvarchar)    客戶簡稱,    ");
            sb.Append(" A.PRODID 產品編號,A.ProdName 品名規格,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END)  總數量,SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN 0 ELSE A.MLAmount*-1 END)/NULLIF(SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0) 平均單價, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 總收入, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  總成本,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 總毛利,       ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount END) END)*100 毛利率       ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A             ");
            sb.Append(" left join otherDB.CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END              ");
            sb.Append(" left join otherDB.CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");

            if (DTYPE == "PVM")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4)='2101' )  ");
            }
            else if (DTYPE == "PVI")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4)='2102' )  ");
            }
            else if (DTYPE == "OTH")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121') )  ");
            }
            else
            {
                sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE=@DTYPE AND USERS=@USERS )");
            }

            sb.Append(" AND T.CUSTID=@CUSTID ");
            sb.Append(" GROUP BY T.CustID,U.ShortName,A.PRODID,A.ProdName");
            sb.Append(" ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@CUSTID", CUSTID));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
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
        public System.Data.DataTable GetCHOD2ARES(string CUSTID)
        {
            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號,CAST(U.SHORTNAME AS nvarchar)    客戶簡稱,    ");
            sb.Append(" ''''+A.PRODID 產品編號,A.ProdName 品名規格,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  總數量,CASE WHEN SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END) =0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  END 平均單價,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) 總收入, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  總成本,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 總毛利,       ");
            sb.Append(" ( CASE WHEN SUM(A.MLAmount)  =0 THEN 0 WHEN 	  SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)=0 THEN 0  ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) END)*100 毛利率         ");
            sb.Append(" From DBO.ComProdRec A             ");
            sb.Append(" left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag             ");
            sb.Append(" left join DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Where A.Flag IN (500,600)    ");
            sb.Append(" AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (comboBox1.Text == "銷售")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)<>'R0'  ");
            }
            if (comboBox1.Text == "維修")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)='R0'  ");
            }
            sb.Append(" AND T.CUSTID=@CUSTID ");
            sb.Append(" GROUP BY T.CustID,U.ShortName,A.PRODID,A.ProdName");

            sb.Append(" ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CUSTID", CUSTID));
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
        public System.Data.DataTable GetARES(decimal F1)
        {
            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號,CAST(U.SHORTNAME AS nvarchar)    客戶簡稱,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  總數量,SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  平均單價, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) 總收入, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  總成本,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 總毛利,     ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) /@F1 END)*100 總收入比率,      ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) END)*100 毛利率       ");
            sb.Append(" From DBO.ComProdRec A             ");
            sb.Append(" left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag             ");
            sb.Append(" left join DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Where A.Flag IN (500,600)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (comboBox1.Text == "銷售")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)<>'R0'  ");
            }
            if (comboBox1.Text == "維修")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)='R0'  ");
            }
            sb.Append(" GROUP BY T.CustID,U.ShortName ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC");

           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
     
            command.Parameters.Add(new SqlParameter("@F1", F1));
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
        public System.Data.DataTable GetCHO(decimal F1,string DTYPE)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號,CAST(U.SHORTNAME AS nvarchar)   客戶簡稱,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END)  總數量,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag= 701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  平均單價, ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 總收入, ");
            sb.Append(" SUM(CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)   總成本,    ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  總毛利,     ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) /@F1 END)*100 總收入比率,      ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END) END)*100 毛利率       ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A             ");
            sb.Append(" left join otherDB.CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END           ");
            sb.Append(" left join otherDB.CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");

            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
          
            if (DTYPE == "AUO")
            {
                sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE IN ('NAUO','AUO') AND USERS=@USERS )");
            }
            else
            {
                sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE = ('PV') AND USERS=@USERS )");
            }
            sb.Append(" GROUP BY T.CustID,U.ShortName ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@F1", F1));
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
        public System.Data.DataTable GetCHOAUO(decimal F1)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            //sb.Append(" Select ROW_NUMBER() OVER( ORDER BY SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC) 客戶排行,T.CustID 客戶編號,CAST(U.SHORTNAME AS nvarchar)   客戶簡稱,    ");
            //sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END)  總數量,");
            //sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag= 701 THEN 0 ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END)  平均單價, ");
            //sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 總收入, ");
            //sb.Append(" SUM(CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)   總成本,    ");
            //sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)  總毛利,     ");
            //sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) /@F1 END)*100 總收入比率,      ");
            //sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST  ELSE A.MLAmount END) END)*100 毛利率       ");
            //sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A             ");
            //sb.Append(" left join otherDB.CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END           ");
            //sb.Append(" left join otherDB.CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            //sb.Append(" Where A.Flag IN (500,600,701)    ");


            sb.Append(" Select 'AUO' 客戶簡稱,");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END),0) 總數量,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag= 701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END),0)  平均單價, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 總收入, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0)  總成本,    ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 總毛利,     ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1   ELSE A.MLAmount*-1 END) /@F1 END),0)*100 總收入比率,      ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) END),0)*100 毛利率       ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A                             ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE = ('AUO') AND USERS=@USERS )");
            sb.Append(" UNION ALL");
            sb.Append(" Select  'NON AUO',");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END),0)  總數量,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag= 701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag= 701 THEN 0 ELSE  A.Quantity*-1 END),0)  平均單價, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 總收入, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0)  總成本,    ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 總毛利,     ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) /@F1 END),0)*100 總收入比率,      ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1  ELSE A.MLAmount*-1 END) END),0)*100 毛利率       ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A                             ");
            sb.Append(" Where A.Flag IN (500,600,701)    ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE = ('NAUO') AND USERS=@USERS )");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@F1", F1));
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
        public System.Data.DataTable GetCHOAUO2(decimal F1)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select 'PV Module' 客戶簡稱, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0) 總數量,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0)  平均單價,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END),0) 總收入,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0)  總成本,     ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 總毛利,      ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) /@F1 END),0)*100 總收入比率,       ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) END),0)*100 毛利率        ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                              ");
            sb.Append(" Where A.Flag IN (500,600)     ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            sb.Append(" AND (SUBSTRING(PRODID,1,4)='2101' )  ");
            sb.Append(" UNION ALL");
            sb.Append(" Select 'PV Inverter' 客戶簡稱, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0) 總數量,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0)  平均單價,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END),0) 總收入,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0)  總成本,     ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 總毛利,      ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) /@F1 END),0)*100 總收入比率,       ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) END),0)*100 毛利率        ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                              ");
            sb.Append(" Where A.Flag IN (500,600)     ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            sb.Append(" AND (SUBSTRING(PRODID,1,4)='2102' )  ");
            sb.Append(" UNION ALL");
            sb.Append(" Select 'Others' 客戶簡稱, ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0) 總數量,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.Quantity ELSE  A.Quantity*-1 END),0)  平均單價,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END),0) 總收入,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0)  總成本,     ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 總毛利,      ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) /@F1 END),0)*100 總收入比率,       ");
            sb.Append(" ISNULL(( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) END),0)*100 毛利率        ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                              ");
            sb.Append(" Where A.Flag IN (500,600)     ");
            sb.Append("  AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            sb.Append(" AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121') )  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@F1", F1));
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
        public System.Data.DataTable GetCHOSUM()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0)   ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A ");
            sb.Append(" Where A.Flag IN (500,600,701)  ");
            sb.Append(" AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE USERS=@USERS )");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetCHOSUMARES()
        {
            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END)   ");
            sb.Append(" From DBO.ComProdRec A ");
            sb.Append(" Where A.Flag IN (500,600)  ");
            sb.Append(" AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            if (comboBox1.Text == "銷售")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)<>'R0'  ");
            }
            if (comboBox1.Text == "維修")
            {
                sb.Append(" AND SUBSTRING(ProdID,1,1)+CAST(CAST(CostForAcc AS decimal(15,0)) AS VARCHAR)='R0'  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetCHOSUMAUO()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
   
            sb.Append(" Select SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END)   ");
            sb.Append(" From otherDB.CHIComp16.DBO.ComProdRec A ");
            sb.Append(" Where A.Flag IN (500,600,701)  ");
            sb.Append(" AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            sb.Append(" AND PRODID IN (SELECT ITEMCODE FROM AD_TYPE WHERE USERS=@USERS AND DTYPE IN ('NAUO','AUO'))");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetCHOSUMAUO2()
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END) ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                              ");
            sb.Append(" Where A.Flag IN (500,600)     ");
            sb.Append(" AND A.BillDate  between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            sb.Append(" AND (SUBSTRING(PRODID,1,4) IN ('2101','2102','2103','21S0','2121')  )  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetS1(string DTYPE)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();


            if (DTYPE == "AUO")
            {
                sb.Append("SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE IN ('NAUO','AUO') AND USERS=@USERS");
            }
            else
            {
                sb.Append("SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE = ('PV') AND USERS=@USERS ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetS12(string DTYPE)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();


            sb.Append("SELECT ITEMCODE FROM AD_TYPE WHERE DTYPE =@DTYPE AND USERS=@USERS");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
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
        public System.Data.DataTable GetTYPES()
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            string GDATE = textBox2.Text;
            //2101
            sb.Append(" Select '當年' DTYPE,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701) AND (SUBSTRING(PRODID,1,4)='2101' )      ");
            sb.Append(" AND A.BillDate  between @M1 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '前月' 客戶簡稱,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 金額,   ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利                      ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701)   AND (SUBSTRING(PRODID,1,4)='2101')      ");
            sb.Append(" AND A.BillDate  between @M3 and @M4 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '當天' 客戶簡稱,  ");
            sb.Append(" SUM(A.MLAmount) 金額,SUM(A.MLAmount - A.CostForAcc) 毛利 From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag = 500 AND (SUBSTRING(PRODID,1,4)='2101' )     ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '減少' 客戶簡稱,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END- A.CostForAcc) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A Where A.Flag IN (600,701) AND (SUBSTRING(PRODID,1,4)='2101')      ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            //
            sb.Append(" Select '當年' DTYPE,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701) AND (SUBSTRING(PRODID,1,4)='2102' )      ");
            sb.Append(" AND A.BillDate  between @M1 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '前月' 客戶簡稱,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 金額,   ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利                      ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701)   AND (SUBSTRING(PRODID,1,4)='2102')      ");
            sb.Append(" AND A.BillDate  between @M3 and @M4 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '當天' 客戶簡稱,  ");
            sb.Append(" SUM(A.MLAmount) 金額,SUM(A.MLAmount - A.CostForAcc) 毛利 From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag = 500 AND (SUBSTRING(PRODID,1,4)='2102' )     ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '減少' 客戶簡稱,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END- A.CostForAcc) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A Where A.Flag IN (600,701) AND (SUBSTRING(PRODID,1,4)='2102')      ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            //2103 21S0
            sb.Append(" Select '當年' DTYPE,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701) AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121') )      ");
            sb.Append(" AND A.BillDate  between @M1 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '前月' 客戶簡稱,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 金額,   ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利                      ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701)   AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121'))      ");
            sb.Append(" AND A.BillDate  between @M3 and @M4 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '當天' 客戶簡稱,  ");
            sb.Append(" SUM(A.MLAmount) 金額,SUM(A.MLAmount - A.CostForAcc) 毛利 From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag = 500 AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121'))     ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '減少' 客戶簡稱,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END- A.CostForAcc) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A Where A.Flag IN (600,701) AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121'))      ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '當年' DTYPE,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701) AND (SUBSTRING(PRODID,1,2)='21' )      ");
            sb.Append(" AND A.BillDate  between @M1 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '前月' 客戶簡稱,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 金額,   ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利                      ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag IN (500,600,701)   AND (SUBSTRING(PRODID,1,2)='21')      ");
            sb.Append(" AND A.BillDate  between @M3 and @M4 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '當天' 客戶簡稱,  ");
            sb.Append(" SUM(A.MLAmount) 金額,SUM(A.MLAmount - A.CostForAcc) 毛利 From CHIComp16.DBO.ComProdRec A                               ");
            sb.Append(" Where A.Flag = 500 AND (SUBSTRING(PRODID,1,2)='21' )     ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            sb.Append(" UNION ALL ");
            sb.Append(" Select '減少' 客戶簡稱,  ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END) 金額,   ");
            sb.Append(" SUM(CASE WHEN A.Flag=701 THEN A.MLDIST ELSE A.MLAmount END- A.CostForAcc) 毛利                       ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A Where A.Flag IN (600,701) AND (SUBSTRING(PRODID,1,2)='21')      ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@M1", DateTime.Now.ToString("yyyy") + "0101"));
            command.Parameters.Add(new SqlParameter("@M2", GDATE));
            command.Parameters.Add(new SqlParameter("@M3", DateTime.Now.ToString("yyyyMM") + "01"));
            command.Parameters.Add(new SqlParameter("@M4", DateTime.Now.AddDays(-1).ToString("yyyyMMdd")));
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
        public System.Data.DataTable GetTYPES2(string TYPE)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            string GDATE = textBox2.Text;
            sb.Append(" Select CAST(U.ShortName AS VARCHAR)    客戶,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END) 金額,");
            sb.Append(" SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) 毛利,                       ");
            sb.Append(" ( CASE SUM(A.MLAmount) WHEN 0 THEN 0 ELSE SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END)/SUM(CASE WHEN A.Flag=500 THEN A.MLAmount ELSE A.MLAmount END) END) 毛利率        ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A        ");
            sb.Append(" left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE WHEN A.Flag=701 THEN T.Flag+3 ELSE  T.Flag END         ");
            sb.Append(" left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID                            ");
            sb.Append(" Where A.Flag IN (500,600,701)      ");
            sb.Append(" AND A.BillDate  between @M2 and @M2 ");
            if (TYPE == "1")
            {
                sb.Append("  AND (SUBSTRING(PRODID,1,4)='2101' )   ");
            }
            if (TYPE == "2")
            {
                sb.Append("  AND (SUBSTRING(PRODID,1,4)='2102' )   ");
            }
            if (TYPE == "3")
            {
                sb.Append(" AND (SUBSTRING(PRODID,1,4) IN ('2103','21S0','2121') )   ");
            }
            sb.Append(" GROUP BY U.ShortName");
            sb.Append(" ORDER BY  SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag =500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@M2", GDATE));

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
        public System.Data.DataTable GETCUST(string ID)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            string GDATE = textBox2.Text;
            sb.Append("    SELECT CAST(SHORTNAME AS NVARCHAR) 客戶  FROM DBO.comCustomer WHERE ID=@ID ");
           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        public System.Data.DataTable GETPVI(string PRODID)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            string GDATE = textBox2.Text;
            sb.Append(" SELECT ltrim(substring(InvoProdName ,0,CHARINDEX(' ', InvoProdName))) PROD, ltrim(substring(InvoProdName ,0,CHARINDEX('-', InvoProdName))) PROD2");
            sb.Append(" ,T1.ClassName   FROM comProduct T0 LEFT JOIN comProductClass T1 ON (T0.ClassID =T1.ClassID)");
            sb.Append(" WHERE T0.PRODID=@PRODID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));

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
        public System.Data.DataTable GetADTYPE(string ITEMCODE)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DTYPE FROM AD_TYPE    WHERE DTYPE='AUO'  AND ITEMCODE=@ITEMCODE AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GETBRAND(string PARAM_NO)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PARAM_DESC  FROM PARAMS   WHERE PARAM_KIND='ADBRAND' AND PARAM_NO=@PARAM_NO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));
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
        public System.Data.DataTable GetTYPE()
        {
         

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select A.PRODID,'NAUO' DTYPE   From otherDB.CHIComp16.DBO.ComProdRec A             ");
            sb.Append(" Where A.Flag IN (500,600,701)  ");
            sb.Append(" AND A.BillDate    between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND PRODID NOT IN ( ");
            sb.Append(" Select  DISTINCT PRODID   From  otherDB.CHIComp16.DBO.ComProdRec A                  ");
            sb.Append(" Where A.Flag IN (500,600,701) ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND ( ");
            sb.Append(" (SUBSTRING(PRODID,1,2)='12' AND SUBSTRING(PRODID,8,4)='-201') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='13' AND SUBSTRING(PRODID,8,4)='-301') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='14' AND SUBSTRING(PRODID,8,4)='-410') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='19' AND SUBSTRING(PRODID,8,4)='-901') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,1)='T') OR  (SUBSTRING(PRODID,1,1)='G')  OR  (PRODID ='1201-190-201012-3F')    ");
            sb.Append(" )");
            sb.Append(" AND PRODID NOT IN ( ");
            sb.Append(" SELECT PRODID FROM otherDB.CHIComp16.DBO.ComProdRec A ");
            sb.Append(" WHERE  A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND SUBSTRING(PRODID,1,4)='1203' AND SUBSTRING(PRODID,8,4)='-201' AND Flag IN (500,600,701) ) AND PRODID <>'T001-001'	 ");
            sb.Append(" UNION ALL");
            sb.Append(" Select  DISTINCT PRODID From otherDB.CHIComp16.DBO.ComProdRec A                  ");
            sb.Append(" Where A.Flag IN (500,600,701) ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            sb.Append(" AND (SUBSTRING(PRODID,1,2)='21') ");
            sb.Append(" )	  		   AND PRODID NOT IN ('(*)','R001-001')  ");
            sb.Append(" UNION ALL");
            sb.Append(" Select A.PRODID,'AUO' DTYPE   From otherDB.CHIComp16.DBO.ComProdRec A              ");
            sb.Append(" Where A.Flag IN (500,600,701)  ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND PRODID  IN ( ");
            sb.Append(" Select  DISTINCT PRODID From otherDB.CHIComp16.DBO.ComProdRec A                  ");
            sb.Append(" Where A.Flag IN (500,600,701) ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND ( ");
            sb.Append(" (SUBSTRING(PRODID,1,2)='12' AND SUBSTRING(PRODID,8,4)='-201') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='13' AND SUBSTRING(PRODID,8,4)='-301') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='14' AND SUBSTRING(PRODID,8,4)='-410') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,2)='19' AND SUBSTRING(PRODID,8,4)='-901') ");
            sb.Append(" OR  (SUBSTRING(PRODID,1,1)='T') OR  (SUBSTRING(PRODID,1,1)='G')  OR  (PRODID ='1201-190-201012-3F')  ");
            sb.Append(" ) ");
            sb.Append(" AND PRODID NOT IN ( ");
            sb.Append(" SELECT PRODID FROM otherDB.CHIComp16.DBO.ComProdRec A ");
            sb.Append(" WHERE  A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND SUBSTRING(PRODID,1,4)='1203' AND SUBSTRING(PRODID,8,4)='-201' AND Flag IN (500,600,701) AND PRODID <>'T001-001' )	  		 ");
            sb.Append(" )	 ");
            sb.Append(" UNION ALL");
            sb.Append(" Select A.PRODID,'PV' DTYPE  From otherDB.CHIComp16.DBO.ComProdRec A              ");
            sb.Append(" Where A.Flag IN (500,600,701)  ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND PRODID  IN ( ");
            sb.Append(" Select  DISTINCT PRODID              From otherDB.CHIComp16.DBO.ComProdRec A                  ");
            sb.Append(" Where A.Flag IN (500,600,701) ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            sb.Append(" AND (SUBSTRING(PRODID,1,2)='21'))");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
       
        private void ACCCHOICE_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            comboBox4.Text = "博豐";
       
            if (globals.UserID.ToUpper() != "NANCYWEI" && globals.GroupID.ToString().Trim() != "EEP" && globals.UserID.ToUpper() != "SHARONHUANG")
            {

                button1.Visible = false;
            }

            if (globals.UserID.ToUpper() == "FIONALAI")
            {

                comboBox4.Items.Clear();

                comboBox4.Items.Add("博豐");
                comboBox4.Items.Add("宇豐");

                button1.Visible = true;
                comboBox4.Text = "宇豐";
             //   comboBox4.Enabled = false;
            }
        }
        public void AddTYPE(string ITEMCODE, string DTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" Insert into AD_TYPE(ITEMCODE,DTYPE,USERS) values(@ITEMCODE,@DTYPE,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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
        public void DELTYPE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete AD_TYPE where users=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
             QQ1 = 0;
             QQ2 = 0;
             QQ3 = 0;
             QQ4 = 0;

             QQ12 = 0;
             QQ22 = 0;
             QQ32 = 0;
             QQ42 = 0;
             Q1 = 0;
             Q2 = 0;
             Q3 = 0;
             Q4 = 0;
            Eun24();
            Eun22();
            System.Data.DataTable H1 = GetS1("AUO");
            if (H1.Rows.Count > 0)
            {
                Eun22AUO();
                Eun22AUO2();
            }
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\AD.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


            string CC = "";
            if (comboBox4.Text == "博豐")
            {
                CC = " 博豐光電股份有限公司";
            }
            if (comboBox4.Text == "聿豐")
            {
                CC = "聿豐實業股份有限公司";
            }
            if (comboBox4.Text == "宇豐")
            {
                CC = "宇豐光電股份有限公司";
            }
            if (comboBox4.Text == "INFINITE")
            {
                CC = "INFINITE";
            }
            if (comboBox4.Text == "CHOICE")
            {
                CC = "CHOICE";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                CC = "TOP GARDEN";
            }
            if (comboBox4.Text == "韋峰")
            {
                CC = "韋峰實業股份有限公司";
            }
            ExcelReport.FIONA(dtAD, ExcelTemplate, OutPutFile,CC);
        }
        private void Eun22()
        {

            dtAD = MakeTableCombine();
            DataRow dr = null;
            decimal F1 = Convert.ToDecimal(GetCHOSUM().Rows[0][0]);
            System.Data.DataTable H1 = GetS1("AUO");
            System.Data.DataTable H2 = GetS1("PV");
            if (H1.Rows.Count > 0)
            {
                Eun23(GetCHO(F1, "AUO"), "TFT  Module");
            }
            if (H2.Rows.Count > 0)
            {
                Eun23(GetCHO(F1, "PV"), "PV Module");
            }
            dr = dtAD.NewRow();
            dr["客戶排行"] = "";
            dr["客戶編號"] = "總計";
            dr["客戶簡稱"] = "";
            dr["總數量"] = QQ1.ToString();
            dr["平均單價"] = "";
            dr["總收入"] = QQ2.ToString();
            dr["總成本"] = QQ3.ToString();
            dr["總毛利"] = QQ4.ToString();
            if (QQ4 == 0 || QQ2 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(QQ4) / Convert.ToDecimal(QQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["毛利率"] = G;
            }
            dr["總收入比率"] = "100.00%";
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);
        }
        private void Eun22ARES()
        {

            dtAD = MakeTableCombine();
            decimal F1 = Convert.ToDecimal(GetCHOSUMARES().Rows[0][0]);
            Eun23ARES(GetARES(F1));
        
        }
        private void Eun22D(string DTYPE,string DTYPE2)
        {

             QQQ1 = 0;
             QQQ2 = 0;
             QQQ3 = 0;
             QQQ4 = 0;
            DataRow dr = null;
            System.Data.DataTable F1 = GetCHOD1(DTYPE);

            if (F1.Rows.Count > 0)
            {
                for (int i = 0; i <= F1.Rows.Count - 1; i++)
                {
                    string CUSTID = F1.Rows[i]["客戶編號"].ToString();
                    Eun23D(GetCHOD2(DTYPE, CUSTID), DTYPE2, i, DTYPE);
                }

                dr = dtAD.NewRow();
                dr["一星"] = "";
                dr["二星"] = "";
                dr["客戶編號"] = "";
                if (DTYPE == "NAUO")
                {
                    DTYPE = "NON AUO";
                }
                else if (DTYPE == "PVM")
                {
                    DTYPE = "PV Module";
                }
                else if (DTYPE == "PVI")
                {
                    DTYPE = "PV Inverter";
                }
                else if (DTYPE == "OTH")
                {
                    DTYPE = "Ohters";
                }
                dr["客戶簡稱"] = DTYPE + "小計";
                dr["廠牌"] = "";
                dr["產品類別"] = "";
                dr["產品編號"] = "";
                dr["品名規格"] = "";
                dr["總數量"] = QQQ1.ToString();
                dr["平均單價"] = "";
                dr["總收入"] = QQQ2.ToString();
                dr["總成本"] = QQQ3.ToString();
                dr["總毛利"] = QQQ4.ToString();
                if (QQQ4 == 0 || QQQ2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    dr["毛利率"] = Math.Round((Convert.ToDecimal(QQQ4) / Convert.ToDecimal(QQQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }

        }

        private void Eun22DARES()
        {

            QQQ1 = 0;
            QQQ2 = 0;
            QQQ3 = 0;
            QQQ4 = 0;
            DataRow dr = null;
            System.Data.DataTable F1 = GetCHOD1ARES();

            for (int i = 0; i <= F1.Rows.Count - 1; i++)
            {
                string CUSTID = F1.Rows[i]["客戶編號"].ToString();
                string CUSTR = F1.Rows[i]["客戶排行"].ToString();
                Eun23DARES(GetCHOD2ARES(CUSTID), CUSTR);
            }

            dr = dtAD.NewRow();
            dr["客戶排行"] = "";
            dr["客戶編號"] = "";
            dr["客戶簡稱"] = "小計";
            dr["產品編號"] = "";
            dr["品名規格"] = "";
            dr["總數量"] = QQQ1.ToString();
            dr["平均單價"] = "";
            dr["總收入"] = QQQ2.ToString();
            dr["總成本"] = QQQ3.ToString();
            dr["總毛利"] = QQQ4.ToString();
            if (QQQ4 == 0 || QQQ2 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                dr["毛利率"] = Math.Round((Convert.ToDecimal(QQQ4) / Convert.ToDecimal(QQQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
            }
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);


        }
        private void Eun22AUO()
        {

            decimal F1 = Convert.ToDecimal(GetCHOSUMAUO().Rows[0][0]);
            Eun23AUO(GetCHOAUO(F1));
     
        }
        private void Eun22AUO2()
        {

            decimal F1 = Convert.ToDecimal(GetCHOSUMAUO2().Rows[0][0]);
            Eun23AUO2(GetCHOAUO2(F1));
            
            dr = dtAD.NewRow();
            dr["客戶排行"] = "";
            dr["客戶編號"] = "總計";
            dr["客戶簡稱"] = "";
            dr["總數量"] = QQ12.ToString();
            dr["平均單價"] = "";
            dr["總收入"] = QQ22.ToString();
            dr["總成本"] = QQ32.ToString();
            dr["總毛利"] = QQ42.ToString();
            if (QQ42 == 0 || QQ22 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(QQ42) / Convert.ToDecimal(QQ22)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["毛利率"] = G;
            }
            dr["總收入比率"] = "100.00%";
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);
        }
        private void Eun23(System.Data.DataTable DT,string DTYPE)
        {
             Q1 = 0;
             Q2 = 0;
             Q3 = 0;
             Q4 = 0;
            System.Data.DataTable dt = DT;
            DataRow dr = null;
            if (dt.Rows.Count > 0)
            {
                dr = dtAD.NewRow();
                dr["客戶排行"] = DTYPE;
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "";
                dr["總數量"] = "";
                dr["平均單價"] = "";
                dr["總收入"] = "";
                dr["總成本"] = "";
                dr["總毛利"] = "";
                dr["毛利率"] = "";
                dr["總收入比率"] = "";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);


                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    dr["客戶排行"] = dd["客戶排行"].ToString();
                    string CUSTID = dd["客戶編號"].ToString();
                    dr["客戶編號"] = CUSTID;
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    System.Data.DataTable CUST1 = GETCUST(CUSTID);
                    if (CUST1.Rows.Count > 0)
                    {
                        dr["客戶簡稱"] = CUST1.Rows[0][0].ToString();
                    }
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    dr["毛利率"] = dd["毛利率"].ToString() + "%";
                    dr["總收入比率"] = dd["總收入比率"].ToString() + "%";
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                    
                    dtAD.Rows.Add(dr);

                }

                QQ1 += Q1;
                QQ2 += Q2;
                QQ3 += Q3;
                QQ4 += Q4;

                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "小計";
                dr["客戶簡稱"] = "";
                dr["總數量"] = Q1.ToString();
                dr["平均單價"] = "";
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
                decimal F1 = Convert.ToDecimal(GetCHOSUM().Rows[0][0]);
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    string G=Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }
                if (Q2 == 0 || F1 == 0)
                {
                    dr["總收入比率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q2) / Convert.ToDecimal(F1)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["總收入比率"] = G;
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void Eun23ARES(System.Data.DataTable DT)
        {
            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            System.Data.DataTable dt = DT;
            DataRow dr = null;
            if (dt.Rows.Count > 0)
            {


                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    dr["客戶排行"] = dd["客戶排行"].ToString();
                    string CUSTID = dd["客戶編號"].ToString();
                    dr["客戶編號"] = CUSTID;
                    System.Data.DataTable CUST1 = GETCUST(CUSTID);
                    if (CUST1.Rows.Count > 0)
                    {
                        dr["客戶簡稱"] = CUST1.Rows[0][0].ToString();
                    }
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    dr["毛利率"] = dd["毛利率"].ToString() + "%";
                    dr["總收入比率"] = dd["總收入比率"].ToString() + "%";
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;

                    dtAD.Rows.Add(dr);

                }

                QQ1 += Q1;
                QQ2 += Q2;
                QQ3 += Q3;
                QQ4 += Q4;

                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "小計";
                dr["客戶簡稱"] = "";
                dr["總數量"] = Q1.ToString();
                dr["平均單價"] = "";
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
             
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }
                if (Q2 == 0)
                {
                    dr["總收入比率"] = "0.00%";
                }
                else
                {
                    dr["總收入比率"] = "100.00%";
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void Eun23D(System.Data.DataTable DT, string DTYPE,int ROW, string DTYPE1)
        {

            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            System.Data.DataTable dt = DT;
            DataRow dr = null;
            if (dt.Rows.Count > 0)
            {
                if (ROW == 0 && DTYPE == "PVM")
                {
                    dr = dtAD.NewRow();
                    dr["一星"] = "";
                    dr["二星"] = "";
                    dr["客戶編號"] = DTYPE;
                    dr["客戶簡稱"] = "";
                    dr["產品編號"] = "";
                    dr["品名規格"] = "";
                    dr["總數量"] = "";
                    dr["平均單價"] = "";
                    dr["總收入"] = "";
                    dr["總成本"] = "";
                    dr["總毛利"] = "";
                    dr["毛利率"] = "";
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                    dtAD.Rows.Add(dr);
                }
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    string ITEM = dd["產品編號"].ToString();
                    string ITEM1 = "";
                    if (ITEM.Length > 11)
                    {
                         ITEM1 = ITEM.Substring(8, 3);
                    }
                    //if (ITEM1 == "B01")
                    //{
                    //    ITEM1 = ITEM.Substring(8, 3);
                    //    MessageBox.Show("A");
                    //}
                    dr["一星"] = "";
                    System.Data.DataTable DD1 = GetADTYPE(ITEM);
                    if (DD1.Rows.Count > 0)
                    {
                        dr["一星"] = "*";
                    }
                    dr["二星"] = "";
                    //if (DTYPE1 == "AUO" || DTYPE1 == "NON AUO" || DTYPE1 == "NAUO")
                    //{
                        System.Data.DataTable PV = GETPVI(ITEM);

                        if (PV.Rows.Count > 0)
                        {
                     
                            string PV1 = PV.Rows[0][2].ToString();


                            DTYPE = PV1;
                        }
                   // }
                    if (DTYPE == "PVM" || DTYPE == "PVI" || DTYPE == "OTH")
                    {
                        dr["二星"] = "**";
                    }
                    string CUSTID = dd["客戶編號"].ToString();

                    dr["客戶編號"] = CUSTID;
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    System.Data.DataTable CUST1 = GETCUST(CUSTID);
                    if (CUST1.Rows.Count > 0)
                    {
                        dr["客戶簡稱"] = CUST1.Rows[0][0].ToString();
                    }
                    if (DTYPE1 == "NAUO")
                    {
                        DTYPE1 = "NON AUO";
                    }
                    else  if (DTYPE1 == "PVM")
                    {
                        DTYPE1 = "AUO";
                    }
                    else if (DTYPE1 == "PVI")
                    {
       
                        if (PV.Rows.Count > 0)
                        {
                            string PV1 = PV.Rows[0][0].ToString();
                            string PV2 = PV.Rows[0][1].ToString();
                            string PTYPE = PV1;
                            if (String.IsNullOrEmpty(PV1))
                            {
                                PTYPE = PV2;
                            }
                            DTYPE1 = PTYPE.ToUpper();
                        }
                    }
                    System.Data.DataTable FD1 = GETBRAND(ITEM1);
                    if (FD1.Rows.Count > 0)
                    {
                        DTYPE1 = FD1.Rows[0][0].ToString();
                    }
                    dr["廠牌"] = DTYPE1;
                    dr["產品類別"] = DTYPE;
                    dr["產品編號"] = ITEM;
                    dr["品名規格"] = dd["品名規格"].ToString();
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    int  T1 = Convert.ToInt32(dd["總收入"]);
                    int T2 = Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    if (T1 == 0 || T2 == 0)
                    {
                        dr["毛利率"] = "0.00%";
                    }
                    else
                    {
                        dr["毛利率"] = Math.Round((Convert.ToDecimal(T2) / Convert.ToDecimal(T1)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    }
                        dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;

                    dtAD.Rows.Add(dr);

                }

                QQ1 += Q1;
                QQ2 += Q2;
                QQ3 += Q3;
                QQ4 += Q4;

                QQQ1 += Q1;
                QQQ2 += Q2;
                QQQ3 += Q3;
                QQQ4 += Q4;

                QQQ1F += Q1;
                QQQ2F += Q2;
                QQQ3F += Q3;
                QQQ4F += Q4;

                dr = dtAD.NewRow();
                dr["一星"] = "";
                dr["二星"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "";
                dr["產品編號"] = "";
                dr["品名規格"] = "本幣合計";
                dr["總數量"] = Q1.ToString();
                if (Q2 == 0 || Q1 == 0)
                {
                    dr["平均單價"] = "0";
                }
                else
                {
                    dr["平均單價"] = Math.Round((Convert.ToDecimal(Q2) / Convert.ToDecimal(Q1)), 2, MidpointRounding.AwayFromZero).ToString();
                }
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00";
                }
                else
                {
                    dr["毛利率"] = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }

        private void Eun23DARES(System.Data.DataTable DT, string CUSTR)
        {

            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            System.Data.DataTable dt = DT;
            DataRow dr = null;
            string CUSTNAME = "";
            if (dt.Rows.Count > 0)
            {
   
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    string ITEM = dd["產品編號"].ToString();
                    if (i == 0)
                    {
                        dr["客戶排行"] = CUSTR;
                    }
                    else
                    {
                        dr["客戶排行"] = "";
                    }
                    string CUSTID = dd["客戶編號"].ToString();
                    dr["客戶編號"] = CUSTID;
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    System.Data.DataTable CUST1 = GETCUST(CUSTID);
                    if (CUST1.Rows.Count > 0)
                    {
                        dr["客戶簡稱"] = CUST1.Rows[0][0].ToString();
                    }
                    dr["產品編號"] = dd["產品編號"].ToString();
                    dr["品名規格"] = dd["品名規格"].ToString();
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    int T1 = Convert.ToInt32(dd["總收入"]);
                    int T2 = Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    if (T1 == 0 || T2 == 0)
                    {
                        dr["毛利率"] = "0.00%";
                    }
                    else
                    {
                        dr["毛利率"] = Math.Round((Convert.ToDecimal(T2) / Convert.ToDecimal(T1)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    }
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;

                    dtAD.Rows.Add(dr);

                }

                QQ1 += Q1;
                QQ2 += Q2;
                QQ3 += Q3;
                QQ4 += Q4;

                QQQ1 += Q1;
                QQQ2 += Q2;
                QQQ3 += Q3;
                QQQ4 += Q4;

                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = CUSTNAME + " 合計";
                dr["產品編號"] = "";
                dr["品名規格"] = "";
                dr["總數量"] = Q1.ToString();
                if (Q2 == 0 || Q1 == 0)
                {
                    dr["平均單價"] = "0";
                }
                else
                {
                    dr["平均單價"] = Math.Round((Convert.ToDecimal(Q2) / Convert.ToDecimal(Q1)), 2, MidpointRounding.AwayFromZero).ToString();
                }
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00";
                }
                else
                {
                    dr["毛利率"] = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void Eun23AUO(System.Data.DataTable DT)
        {
            DataRow dr = null;
            for (int i = 0; i <= 1; i++)
            {
                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "";
                dr["總數量"] = "";
                dr["平均單價"] = "";
                dr["總收入"] = "";
                dr["總成本"] = "";
                dr["總毛利"] = "";
                dr["毛利率"] = "";
                dr["總收入比率"] = "";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }

            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            System.Data.DataTable dt = DT;

            if (dt.Rows.Count > 0)
            {
                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "廠牌";
                dr["總數量"] = "總數量";
                dr["平均單價"] = "平均單價";
                dr["總收入"] = "總收入";
                dr["總成本"] = "總成本";
                dr["總毛利"] = "總毛利";
                dr["毛利率"] = "毛利率";
                dr["總收入比率"] = "總收入比率";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);


                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    string MODULE = "";
                    if (i == 0)
                    {
                        MODULE = "TFT Module";
                    }
                    dr["客戶排行"] = MODULE;
                    dr["客戶編號"] = "";

                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    dr["毛利率"] = dd["毛利率"].ToString() + "%";
                    dr["總收入比率"] = dd["總收入比率"].ToString() + "%";
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;

                    dtAD.Rows.Add(dr);

                }

                QQ12 += Q1;
                QQ22 += Q2;
                QQ32 += Q3;
                QQ42 += Q4;

                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "小計";
                dr["客戶簡稱"] = "";
                dr["總數量"] = Q1.ToString();
                dr["平均單價"] = "";
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
                decimal F1 = Convert.ToDecimal(GetCHOSUMAUO().Rows[0][0]);
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }
                if (Q2 == 0 || F1 == 0)
                {
                    dr["總收入比率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q2) / Convert.ToDecimal(F1)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["總收入比率"] = G;
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void Eun23AUO2(System.Data.DataTable DT)
        {
            DataRow dr = null;
            dr = dtAD.NewRow();
            dr["客戶排行"] = "";
            dr["客戶編號"] = "";
            dr["客戶簡稱"] = "";
            dr["總數量"] = "";
            dr["平均單價"] = "";
            dr["總收入"] = "";
            dr["總成本"] = "";
            dr["總毛利"] = "";
            dr["毛利率"] = "";
            dr["總收入比率"] = "";
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);

            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            System.Data.DataTable dt = DT;
            
            if (dt.Rows.Count > 0)
            {
                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "產品類別";
                dr["總數量"] = "總數量";
                dr["平均單價"] = "平均單價";
                dr["總收入"] = "總收入";
                dr["總成本"] = "總成本";
                dr["總毛利"] = "總毛利";
                dr["毛利率"] = "毛利率";
                dr["總收入比率"] = "總收入比率";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);


                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow dd = dt.Rows[i];
                    dr = dtAD.NewRow();
                    string MODULE = "";
                    if (i == 0)
                    {
                        MODULE = "PV Module";
                    }
                    dr["客戶排行"] = MODULE;
                    dr["客戶編號"] = "";
 
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    Q1 += Convert.ToInt32(dd["總數量"]);
                    Q2 += Convert.ToInt32(dd["總收入"]);
                    Q3 += Convert.ToInt32(dd["總成本"]);
                    Q4 += Convert.ToInt32(dd["總毛利"]);
                    dr["總數量"] = dd["總數量"].ToString();
                    dr["平均單價"] = dd["平均單價"].ToString();
                    dr["總收入"] = dd["總收入"].ToString();
                    dr["總成本"] = dd["總成本"].ToString();
                    dr["總毛利"] = dd["總毛利"].ToString();
                    dr["毛利率"] = dd["毛利率"].ToString() + "%";
                    dr["總收入比率"] = dd["總收入比率"].ToString() + "%";
                    dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;

                    dtAD.Rows.Add(dr);

                }

                QQ12 += Q1;
                QQ22 += Q2;
                QQ32 += Q3;
                QQ42 += Q4;

                dr = dtAD.NewRow();
                dr["客戶排行"] = "";
                dr["客戶編號"] = "小計";
                dr["客戶簡稱"] = "";
                dr["總數量"] = Q1.ToString();
                dr["平均單價"] = "";
                dr["總收入"] = Q2.ToString();
                dr["總成本"] = Q3.ToString();
                dr["總毛利"] = Q4.ToString();
                decimal F1 = Convert.ToDecimal(GetCHOSUMAUO2().Rows[0][0]);
                if (Q4 == 0 || Q2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }
                if (Q2 == 0 || F1 == 0)
                {
                    dr["總收入比率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(Q2) / Convert.ToDecimal(F1)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["總收入比率"] = G;
                }
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void Eun24()
        {
            DELTYPE();
            System.Data.DataTable dt3 = GetTYPE();

            if (dt3.Rows.Count > 0)
            {
                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];

                   string ITEMCODE = drw3["PRODID"].ToString().Trim();
                   string DDTYPE = drw3["DTYPE"].ToString().Trim();
                   AddTYPE(ITEMCODE, DDTYPE);
                }
            }
            
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("客戶排行", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("總數量", typeof(string));
            dt.Columns.Add("平均單價", typeof(string));
            dt.Columns.Add("總收入", typeof(string));
            dt.Columns.Add("總成本", typeof(string));
            dt.Columns.Add("總毛利", typeof(string));
            dt.Columns.Add("毛利率", typeof(string));
            dt.Columns.Add("總收入比率", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            return dt;

        }
        private System.Data.DataTable MakeTableCombineD()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("一星", typeof(string));
            dt.Columns.Add("二星", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("廠牌", typeof(string));
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("總數量", typeof(string));
            dt.Columns.Add("平均單價", typeof(string));
            dt.Columns.Add("總收入", typeof(string));
            dt.Columns.Add("總成本", typeof(string));
            dt.Columns.Add("總毛利", typeof(string));
            dt.Columns.Add("毛利率", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            return dt;

        }
        private System.Data.DataTable MakeTableCombineARES()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("客戶排行", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("總數量", typeof(string));
            dt.Columns.Add("平均單價", typeof(string));
            dt.Columns.Add("總收入", typeof(string));
            dt.Columns.Add("總成本", typeof(string));
            dt.Columns.Add("總毛利", typeof(string));
            dt.Columns.Add("毛利率", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            return dt;

        }
        private void button4_Click(object sender, EventArgs e)
        {
            QQ1 = 0;
            QQ2 = 0;
            QQ3 = 0;
            QQ4 = 0;
            QQQ1F = 0;
            QQQ2F = 0;
            QQQ3F = 0;
            QQQ4F = 0;

            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            Eun24();
            dtAD = MakeTableCombineD();

            System.Data.DataTable F2 = GetS12("PV");
            DataRow dr = null;

            Eun22D("AUO", "LCD PANEL");
            Eun22D("NAUO", "LCD PANEL");


            dr = dtAD.NewRow();
            dr["一星"] = "";
            dr["二星"] = "";
            dr["客戶編號"] = "";
        
            dr["客戶簡稱"] = "總計";
            dr["廠牌"] = "";
            dr["產品類別"] = "";
            dr["產品編號"] = "";
            dr["品名規格"] = "";
            dr["總數量"] = QQQ1F.ToString();
            dr["平均單價"] = "";
            dr["總收入"] = QQQ2F.ToString();
            dr["總成本"] = QQQ3F.ToString();
            dr["總毛利"] = QQQ4F.ToString();
            if (QQQ4F == 0 || QQQ2F == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                dr["毛利率"] = Math.Round((Convert.ToDecimal(QQQ4F) / Convert.ToDecimal(QQQ2F)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
            }
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);


            dr = dtAD.NewRow();

                dr["一星"] = "";
                dr["二星"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "";
                dr["產品編號"] = "";
                dr["品名規格"] = "";
                dr["總數量"] = "";
                dr["平均單價"] = "";
                dr["總收入"] = "";
                dr["總成本"] = "";
                dr["總毛利"] = "";
                dr["毛利率"] = "";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            
    
                    Eun22D("PVM", "PV Module");
                    Eun22D("PVI", "PV Inverter");
                    Eun22D("OTH", "Ohters");
                
            
 
            dr = dtAD.NewRow();
            dr["一星"] = "";
            dr["二星"] = "";
            dr["客戶編號"] = "";
            dr["客戶簡稱"] = "總計";
            dr["產品編號"] = "";
            dr["品名規格"] = "";
            dr["總數量"] = QQ1.ToString();
            dr["平均單價"] = "";
            dr["總收入"] = QQ2.ToString();
            dr["總成本"] = QQ3.ToString();
            dr["總毛利"] = QQ4.ToString();
            if (QQ4 == 0 || QQ2 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                dr["毛利率"] = Math.Round((Convert.ToDecimal(QQ4) / Convert.ToDecimal(QQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
            }
            dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
            dtAD.Rows.Add(dr);
         //   Eun22AUO();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\AD2.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


            string CC = "";
            if (comboBox4.Text == "博豐")
            {
                CC = " 博豐光電股份有限公司";
            }
            if (comboBox4.Text == "聿豐")
            {
                CC = "聿豐實業股份有限公司";
            }
            if (comboBox4.Text == "宇豐")
            {
                CC = "宇豐光電股份有限公司";
            }
            if (comboBox4.Text == "INFINITE")
            {
                CC = "INFINITE";
            }
            if (comboBox4.Text == "CHOICE")
            {
                CC = "CHOICE";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                CC = "TOP GARDEN";
            }
            if (comboBox4.Text == "韋峰")
            {
                CC = "韋峰實業股份有限公司";
            }
            ExcelReport.FIONA(dtAD, ExcelTemplate, OutPutFile, CC);
        }
        private void AUONAUO(string DOCTYPE)
        {
            dtAD = MakeTableCombineD();
            System.Data.DataTable F1 = GetS12(DOCTYPE);
            DataRow dr = null;
            if (F1.Rows.Count > 0)
            {
                Eun22D("AUO", "TFT Module");

                dr = dtAD.NewRow();

                dr["一星"] = "";
                dr["二星"] = "";
                dr["客戶編號"] = "";
                dr["客戶簡稱"] = "";
                dr["產品編號"] = "";
                dr["品名規格"] = "";
                dr["總數量"] = "";
                dr["平均單價"] = "";
                dr["總收入"] = "";
                dr["總成本"] = "";
                dr["總毛利"] = "";
                dr["毛利率"] = "";
                dr["日期"] = "日期區間:" + textBox1.Text + "～" + textBox2.Text;
                dtAD.Rows.Add(dr);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable gg1 = GetTYPES2("1");
            System.Data.DataTable gg3 = GetTYPES2("2");
            System.Data.DataTable gg4 = GetTYPES2("3");
            System.Data.DataTable gg2 = GetTYPES();
            if (gg2.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\AD3.xls";
                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel ReportdataGridView1
                ExcelReport.FIONA2(gg1,gg3,gg4, gg2, ExcelTemplate, OutPutFile);
            }
            else
            {
                MessageBox.Show("今日無金額");
            }
        }
        private void Eun24S()
        {
            DELTYPE();
            System.Data.DataTable dt3 = GetTYPE();

            if (dt3.Rows.Count > 0)
            {
                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];

                    string ITEMCODE = drw3["PRODID"].ToString().Trim();
                    string DDTYPE = drw3["DTYPE"].ToString().Trim();
                    AddTYPE(ITEMCODE, DDTYPE);
                }
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            QQ1 = 0;
            QQ2 = 0;
            QQ3 = 0;
            QQ4 = 0;
            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            Eun22ARES();
        
                 string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\ARES1.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            string CC = "";
            if (comboBox4.Text == "博豐")
            {
                CC = " 博豐光電股份有限公司";
            }
            if (comboBox4.Text == "聿豐")
            {
                CC = "聿豐實業股份有限公司";
            }
            if (comboBox4.Text == "宇豐")
            {
                CC = "宇豐光電股份有限公司";
            }
            if (comboBox4.Text == "INFINITE")
            {
                CC = "INFINITE";
            }
            if (comboBox4.Text == "CHOICE")
            {
                CC = "CHOICE";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                CC = "TOP GARDEN";
            }
            if (comboBox4.Text == "韋峰")
            {
                CC = "韋峰實業股份有限公司";
            }
            ExcelReport.FIONA(dtAD, ExcelTemplate, OutPutFile, CC);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            QQ1 = 0;
            QQ2 = 0;
            QQ3 = 0;
            QQ4 = 0;
            Q1 = 0;
            Q2 = 0;
            Q3 = 0;
            Q4 = 0;
            dtAD = MakeTableCombineARES();
         
                Eun22DARES();


            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\ARES2.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
            string CC = "";
            if (comboBox4.Text == "博豐")
            {
                CC = " 博豐光電股份有限公司";
            }
            if (comboBox4.Text == "聿豐")
            {
                CC = "聿豐實業股份有限公司";
            }
            if (comboBox4.Text == "宇豐")
            {
                CC = "宇豐光電股份有限公司";
            }
            if (comboBox4.Text == "INFINITE")
            {
                CC = "INFINITE";
            }
            if (comboBox4.Text == "CHOICE")
            {
                CC = "CHOICE";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                CC = "TOP GARDEN";
            }
            if (comboBox4.Text == "韋峰")
            {
                CC = "韋峰實業股份有限公司";
            }
            ExcelReport.FIONA(dtAD, ExcelTemplate, OutPutFile,CC);
        }

        
    }
}

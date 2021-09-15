using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
namespace ACME
{
    public partial class STOCK : Form
    {
        System.Data.DataTable dtCost = null;
        public STOCK()
        {
            InitializeComponent();
        }

 
        private void button1_Click(object sender, EventArgs e)
        {
            decimal QTY = 0;
            decimal AMT = 0;
            System.Data.DataTable dt = GetSTOCK(textBox1.Text);
            DataRow dr = null;
            dtCost = MakeTableCombine();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string ITEMCODE = dt.Rows[i]["項目"].ToString();
                dr = dtCost.NewRow();
                dr["群組"] = dt.Rows[i]["群組"].ToString();
                dr["倉庫"] = dt.Rows[i]["倉庫"].ToString();
                dr["倉庫名稱"] = dt.Rows[i]["倉庫名稱"].ToString();
                dr["項目"] = ITEMCODE;
                
                dr["數量"] = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());

                QTY = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());
                System.Data.DataTable dt1 = GetSTOCK2(ITEMCODE,textBox1.Text);
                for (int h = 0; h <= dt1.Rows.Count - 1; h++)
                {
                    decimal AMOUNT = Convert.ToDecimal(dt1.Rows[h]["庫存金額"].ToString());
                    decimal QTY2 = Convert.ToDecimal(dt1.Rows[h]["數量"].ToString());
                    AMT = (QTY / QTY2) * AMOUNT;
                    dr["庫存金額"] = Convert.ToDecimal(AMT.ToString());
                }

                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;
        }
        public static System.Data.DataTable GetSTOCK(string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MAX(substring(T3.itmsgrpNAM,4,15)) 群組,T0.warehouse as 倉庫,W.WhsName 倉庫名稱,");
            sb.Append(" T0.[ItemCode] 項目, SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量 ");
            sb.Append(" FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append(" ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" LEFT JOIN OWHS W on (T0.warehouse=W.whscode) ");
            sb.Append(" INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append(" AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組' ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append(" GROUP BY T0.warehouse,W.WhsName,T0.[ItemCode]");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) order by T0.[ItemCode]");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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

        public static System.Data.DataTable GetSTOCKS2(string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MAX(substring(T3.itmsgrpNAM,4,15)) 群組,T0.warehouse as 倉庫,W.WhsName 倉庫名稱,");
            sb.Append(" T0.[ItemCode] 項目, SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量 ");
            sb.Append(" FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append(" ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" LEFT JOIN OWHS W on (T0.warehouse=W.whscode) ");
            sb.Append(" INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.warehouse='OT001' ");
            sb.Append(" AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append(" GROUP BY T0.warehouse,W.WhsName,T0.[ItemCode]");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) order by T0.[ItemCode]");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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

        public static System.Data.DataTable GetSTOCKS3(string DOCDATE, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BASE_REF");
            sb.Append(" FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append(" ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" LEFT JOIN OWHS W on (T0.warehouse=W.whscode) ");
            sb.Append(" INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.warehouse='OT001' AND T0.ITEMCODE=@ITEMCODE AND TRANSTYPE=20 ");
            sb.Append(" AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組' ' ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append(" GROUP BY T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF ");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) =@QTY) order by  CAST(T0.BASE_REF AS INT)  DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public static System.Data.DataTable GetSTOCKS3G(string DOCDATE, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BASE_REF");
            sb.Append(" FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append(" ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" LEFT JOIN OWHS W on (T0.warehouse=W.whscode) ");
            sb.Append(" INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.warehouse='OT001' AND T0.ITEMCODE=@ITEMCODE AND TRANSTYPE=67 ");
            sb.Append(" AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append(" AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append(" GROUP BY T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF ");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) =@QTY) order by  CAST(T0.BASE_REF AS INT)  DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public static System.Data.DataTable GetSTOCKS4(string DOCDATE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 2 T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF,SUM(T0.[InQty])-SUM(T0.[OutQty]) QTY");
            sb.Append("            FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append("            ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("            INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append("            INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append("            WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND   @DOCDATE ) ");
            sb.Append("            AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.warehouse='OT001'AND T0.ITEMCODE=@ITEMCODE AND TRANSTYPE='20'");
            sb.Append("            AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("             FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("            WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND   @DOCDATE ) ");
            sb.Append("            AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append("            GROUP BY T0.[ItemCode]  ");
            sb.Append("            Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append("            GROUP BY T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF");
            sb.Append("            Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) order by  CAST(T0.BASE_REF AS INT)  DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public static System.Data.DataTable GetSTOCKT4(string DOCDATE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 3 T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF,SUM(T0.[InQty])-SUM(T0.[OutQty]) QTY");
            sb.Append("            FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append("            ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("            INNER  JOIN [dbo].[OITM] T2  ON  T2.[ItemCode] = T0.ItemCode   ");
            sb.Append("            INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod   ");
            sb.Append("            WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND   @DOCDATE ) ");
            sb.Append("            AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.warehouse='OT001'AND T0.ITEMCODE=@ITEMCODE AND TRANSTYPE='20'");
            sb.Append("            AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("             FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("            WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND   @DOCDATE ) ");
            sb.Append("            AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append("            GROUP BY T0.[ItemCode]  ");
            sb.Append("            Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append("            GROUP BY T0.warehouse,T0.[ItemCode],TRANSTYPE,BASE_REF");
            sb.Append("            Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) order by  CAST(T0.BASE_REF AS INT)  DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public static System.Data.DataTable GetSTOCKS5(string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_SHIPPING_NO JOBNO FROM OPDN WHERE DOCENTRY=@DOCDATE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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

        public static System.Data.DataTable GetSTOCKS6(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT tradeCondition FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        public static System.Data.DataTable GetSTOCKA(string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("         SELECT T0.warehouse as 倉庫,W.WhsName 倉庫名稱,");
            sb.Append("          CAST(SUM(T0.[InQty])-SUM(T0.[OutQty]) AS INT) 數量 ");
            sb.Append("              FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ");
            sb.Append("              ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("              LEFT JOIN OWHS W on (T0.warehouse=W.whscode)   ");
            sb.Append("              WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append("              AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            sb.Append("              AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("               FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("              WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append("              AND ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append("              GROUP BY T0.[ItemCode]  ");
            sb.Append("              Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append("              GROUP BY T0.warehouse,W.WhsName");
            sb.Append("              Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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
        public static System.Data.DataTable GetSTOCK2(string ITEMCODE, string DOCDATE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();



            sb.Append("        SELECT T0.[ItemCode] 項目, SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,SUM(TRANSVALUE) 庫存金額");
            sb.Append("              FROM  [dbo].[OINM] T0  ");
            sb.Append("              WHERE  T0.ITEMCODE = @ITEMCODE  and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
            sb.Append("              GROUP BY T0.[ItemCode]");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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
        public static System.Data.DataTable GetSTOCK2A()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT WHCODE 倉庫編號,WHNAME 倉庫名稱,SUM(QTY) 數量,SUM(AMOUNT)  金額 FROM WH_STOCK");
            sb.Append(" GROUP BY WHCODE,WHNAME");
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
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("倉庫名稱", typeof(string));
            dt.Columns.Add("項目", typeof(string));
            dt.Columns.Add("數量", typeof(Decimal));
            dt.Columns.Add("庫存金額", typeof(Decimal));
            
            return dt;
        }
        private System.Data.DataTable MakeTableCombineS2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("倉庫名稱", typeof(string));
            dt.Columns.Add("項目", typeof(string));
            dt.Columns.Add("數量", typeof(Decimal));
            dt.Columns.Add("庫存金額", typeof(Decimal));
            dt.Columns.Add("收貨採購單號碼", typeof(string));
            dt.Columns.Add("shipping工單號碼", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("庫存調撥", typeof(string));
            //shipping工單號碼
            return dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void STOCK_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            AddAUOGD4();
            decimal QTY = 0;
            decimal AMT = 0;
            System.Data.DataTable dt = GetSTOCK(textBox1.Text);
            DataRow dr = null;
            dtCost = MakeTableCombine();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string ITEMCODE = dt.Rows[i]["項目"].ToString();
                dr = dtCost.NewRow();
                dr["群組"] = dt.Rows[i]["群組"].ToString();
                dr["倉庫"] = dt.Rows[i]["倉庫"].ToString();
                dr["倉庫名稱"] = dt.Rows[i]["倉庫名稱"].ToString();
                dr["項目"] = ITEMCODE;

                dr["數量"] = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());

                QTY = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());
                System.Data.DataTable dt1 = GetSTOCK2(ITEMCODE, textBox1.Text);
                for (int h = 0; h <= dt1.Rows.Count - 1; h++)
                {
                    decimal AMOUNT = Convert.ToDecimal(dt1.Rows[h]["庫存金額"].ToString());
                    decimal QTY2 = Convert.ToDecimal(dt1.Rows[h]["數量"].ToString());

                    AMT = (QTY / QTY2) * AMOUNT;
                    dr["庫存金額"] = Convert.ToDecimal(AMT.ToString());
                }
                dtCost.Rows.Add(dr);
                AddAUOGD5(dr["群組"].ToString(), dr["倉庫"].ToString(), dr["倉庫名稱"].ToString(), dr["項目"].ToString(), QTY, Convert.ToDecimal(dr["庫存金額"]));
            }
            dataGridView1.DataSource = GetSTOCK2A();
        }
        public void AddAUOGD5(string BU, string WHCODE, string WHNAME, string ITEM, decimal QTY, decimal AMOUNT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_STOCK(BU,WHCODE,WHNAME,ITEM,QTY,AMOUNT) values(@BU,@WHCODE,@WHNAME,@ITEM,@QTY,@AMOUNT)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@WHCODE", WHCODE));
            command.Parameters.Add(new SqlParameter("@WHNAME", WHNAME));
            command.Parameters.Add(new SqlParameter("@ITEM", ITEM));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@AMOUNT", AMOUNT));

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
        public void AddAUOGD4()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE WH_STOCK", connection);
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

        private void button4_Click(object sender, EventArgs e)
        {
            decimal QTY = 0;
            decimal AMT = 0;
            System.Data.DataTable dt = GetSTOCKS2(textBox1.Text);
            DataRow dr = null;
            dtCost = MakeTableCombineS2();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string ITEMCODE = dt.Rows[i]["項目"].ToString();
                dr = dtCost.NewRow();
                dr["群組"] = dt.Rows[i]["群組"].ToString();
                dr["倉庫"] = dt.Rows[i]["倉庫"].ToString();
                dr["倉庫名稱"] = dt.Rows[i]["倉庫名稱"].ToString();
                dr["項目"] = ITEMCODE;

                if (ITEMCODE == "M215HTN01.01002")
                {
                    MessageBox.Show("A");
                }
                dr["數量"] = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());

                QTY = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());
                System.Data.DataTable dt1 = GetSTOCK2(ITEMCODE, textBox1.Text);
                for (int h = 0; h <= dt1.Rows.Count - 1; h++)
                {
                    decimal AMOUNT = Convert.ToDecimal(dt1.Rows[h]["庫存金額"].ToString());
                    decimal QTY2 = Convert.ToDecimal(dt1.Rows[h]["數量"].ToString());
                    AMT = (QTY / QTY2) * AMOUNT;
                    dr["庫存金額"] = Convert.ToDecimal(AMT.ToString());
                }
                System.Data.DataTable dt2 = GetSTOCKS3(textBox1.Text, ITEMCODE, dt.Rows[i]["數量"].ToString());
         
                if (dt2.Rows.Count > 0)
                {
                    dr["收貨採購單號碼"] = dt2.Rows[0][0].ToString();
                    System.Data.DataTable dt4 = GetSTOCKS5(dt2.Rows[0][0].ToString());

                    if (dt4.Rows.Count > 0)
                    {
                        dr["shipping工單號碼"] = dt4.Rows[0][0].ToString();

                        System.Data.DataTable dtS4 = GetSTOCKS6(dt4.Rows[0][0].ToString());
                        if (dtS4.Rows.Count > 0)

                        {

                            dr["貿易條件"] = dtS4.Rows[0][0].ToString();
                        }
                    }
                
                }

                System.Data.DataTable dt3 = GetSTOCKS4(textBox1.Text, ITEMCODE);
                if (dt3.Rows.Count == 2)
                {
                    if (String.IsNullOrEmpty(dr["收貨採購單號碼"].ToString()))
                    {
                        decimal J1 = 0;
                        string BASE = "";
                        string JOBNO = "";
                        string TRADE = "";
                        for (int h = 0; h <= 1; h++)
                        {
                            decimal H1 = 0;
                            string fj = dt3.Rows[h]["QTY"].ToString();
                            if (!String.IsNullOrEmpty(dt3.Rows[h]["QTY"].ToString()))
                            {
                                 H1 = Convert.ToDecimal(dt3.Rows[h]["QTY"].ToString());
                            }
                            string BASE_REF = dt3.Rows[h]["BASE_REF"].ToString();
                            System.Data.DataTable dt4 = GetSTOCKS5(BASE_REF);
                            if (dt4.Rows.Count > 0)
                            {
                                JOBNO += dt4.Rows[0][0].ToString() + "/";


                                System.Data.DataTable dtS4 = GetSTOCKS6(dt4.Rows[0][0].ToString());
                                if (dtS4.Rows.Count > 0)
                                {

                                    TRADE += dtS4.Rows[0][0].ToString() + "/";
                                }
                            }



                          

                            BASE += BASE_REF + "/";
                           
                            J1 += H1;
                        }
                        string h1 = dt.Rows[i]["數量"].ToString();
                        if (J1.ToString() == dt.Rows[i]["數量"].ToString())
                        {
                            dr["收貨採購單號碼"] = BASE.Remove(BASE.Length - 1, 1);
                            dr["shipping工單號碼"] = JOBNO.Remove(JOBNO.Length - 1, 1);
                            dr["貿易條件"] = TRADE.Remove(TRADE.Length - 1, 1);
                        }


                    }
                }


                System.Data.DataTable dtH3 = GetSTOCKT4(textBox1.Text, ITEMCODE);
                if (dtH3.Rows.Count > 3)
                {
                    if (String.IsNullOrEmpty(dr["收貨採購單號碼"].ToString()))
                    {
                        decimal J1 = 0;
                        string BASE = "";
                        string JOBNO = "";
                        string TRADE = "";
                        for (int h = 0; h <= 2; h++)
                        {
                            decimal H1 = 0;
                            string fj = dtH3.Rows[h]["QTY"].ToString();
                            if (!String.IsNullOrEmpty(dtH3.Rows[h]["QTY"].ToString()))
                            {
                                H1 = Convert.ToDecimal(dtH3.Rows[h]["QTY"].ToString());
                            }
                            string BASE_REF = dtH3.Rows[h]["BASE_REF"].ToString();
                            System.Data.DataTable dt4 = GetSTOCKS5(BASE_REF);
                            if (dt4.Rows.Count > 0)
                            {
                                JOBNO += dt4.Rows[0][0].ToString() + "/";


                                System.Data.DataTable dtS4 = GetSTOCKS6(dt4.Rows[0][0].ToString());
                                if (dtS4.Rows.Count > 0)
                                {

                                    TRADE += dtS4.Rows[0][0].ToString() + "/";
                                }
                            }





                            BASE += BASE_REF + "/";

                            J1 += H1;
                        }
                        string h1 = dt.Rows[i]["數量"].ToString();
                        if (J1.ToString() == dt.Rows[i]["數量"].ToString())
                        {
                            dr["收貨採購單號碼"] = BASE.Remove(BASE.Length - 1, 1);
                            dr["shipping工單號碼"] = JOBNO.Remove(JOBNO.Length - 1, 1);
                            dr["貿易條件"] = TRADE.Remove(TRADE.Length - 1, 1);
                        }


                    }
                }

                //string N1 = textBox1.Text.Substring(0, 4) + "." + textBox1.Text.Substring(4, 2) + "." + textBox1.Text.Substring(6, 2);
                //DateTime P1 = Convert.ToDateTime(N1).AddMonths(1);
                //string m1 = P1.ToString("yyyyMMdd");
                //System.Data.DataTable dt2G = GetSTOCKS3G(m1, ITEMCODE, dt.Rows[i]["數量"].ToString());
                //if (dt2G.Rows.Count > 0)
                //{

                //    dr["庫存調撥"] = dt2G.Rows[0][0].ToString();
                //}

                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;
        }
    }
}
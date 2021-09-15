 using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

/// 期末存貨為 0 不顯示
/// 所有欄位為 0 就不顯示

namespace ACME
{
    public partial class fmStockMoveS : Form
    {
        private string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";

        public fmStockMoveS()
        {
            InitializeComponent();
        }




        //取得某一時點的庫存列表
        //T0.[TransType] = 162  -> Inventory Valuation 
        private System.Data.DataTable GetItemHisList(string DocDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            //            SELECT ITMSGRPCOD,SUBSTRING(ITMSGRPNAM,4,LEN(ITMSGRPNAM)-3) FROM OITB
            //SELECT ITEMCODE,U_GROUP,ltrim(substring(U_GROUP,CHARINDEX('-', U_GROUP)+1,LEN(U_GROUP))) FROM OITM 
            sb.Append("SELECT T0.[ItemCode], T1.[ItemName],SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue,MAX(SUBSTRING(ITMSGRPNAM,4,LEN(ITMSGRPNAM)-3)) BU,MAX(ltrim(substring(U_GROUP,CHARINDEX('-', U_GROUP)+1,LEN(U_GROUP)))) ITEM ");
            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("INNER  JOIN [dbo].[OITB] T2  ON  T2.[ITMSGRPCOD] = T1.ITMSGRPCOD  ");
            sb.Append("WHERE  T0.[DocDate] <= @DocDate ");
            //sb.Append("AND  T0.[ItemCode] >= @ItemCode1  ");
            //sb.Append("AND  T0.[ItemCode] <= @ItemCode2  ");
            sb.Append("and ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
     
         //   sb.Append("AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
            sb.Append("GROUP BY  T0.[ItemCode], T1.[ItemName] ");
           // sb.Append("Having SUM(T0.[InQty] - T0.[OutQty])> 0 ");
            sb.Append("ORDER BY  T0.[ItemCode]");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }


        //取得某一時點的庫存列表
        //T0.[TransType] = 162  -> Inventory Valuation 
        private System.Data.DataTable GetItemHisListByTransType( string DocDate1, string DocDate2)
        {
         
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT T0.[ItemCode], T1.[ItemName],T0.[TransType],SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue,MAX(SUBSTRING(ITMSGRPNAM,4,LEN(ITMSGRPNAM)-3)) BU,MAX(ltrim(substring(U_GROUP,CHARINDEX('-', U_GROUP)+1,LEN(U_GROUP)))) ITEM ");
            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("INNER  JOIN [dbo].[OITB] T2  ON  T2.[ITMSGRPCOD] = T1.ITMSGRPCOD  ");
            sb.Append("WHERE  T0.[DocDate] >= @DocDate1 ");
            sb.Append("And    T0.[DocDate] <= @DocDate2 ");
            sb.Append("and ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  ");
    
          //  sb.Append("AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
            sb.Append("GROUP BY  T0.[ItemCode], T1.[ItemName],T0.[TransType] ");
          //  sb.Append("Having SUM(T0.[InQty] - T0.[OutQty])> 0 ");
            sb.Append("ORDER BY  T0.[ItemCode]");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
  
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }


        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位
            //
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("ITEM", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("期初存貨", typeof(decimal));
            dt.Columns.Add("期初存貨T", typeof(decimal));
            dt.Columns.Add("本期進貨", typeof(decimal));
            dt.Columns.Add("本期進貨T", typeof(decimal));
            dt.Columns.Add("進貨退出折讓", typeof(decimal));
            dt.Columns.Add("進貨退出折讓T", typeof(decimal));
            dt.Columns.Add("本期銷貨", typeof(decimal));
            dt.Columns.Add("本期銷貨T", typeof(decimal));
            dt.Columns.Add("銷貨退回", typeof(decimal));
            dt.Columns.Add("銷貨退回T", typeof(decimal));
            dt.Columns.Add("本期調整", typeof(decimal));
            dt.Columns.Add("本期調整T", typeof(decimal));
            dt.Columns.Add("本期調撥", typeof(decimal));
            dt.Columns.Add("本期調撥T", typeof(decimal));
            dt.Columns.Add("期末存貨", typeof(decimal));
            dt.Columns.Add("期末存貨T", typeof(decimal));
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["產品編號"];
            dt.PrimaryKey = colPk;


            return dt;
        }


        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void GoEdit(DataRow row, string TransType, decimal Qty, decimal TRAN)
        {

            if (TransType == "20" || TransType == "18")
            {
                row["本期進貨"] = Qty + Convert.ToDecimal(row["本期進貨"]);
                row["本期進貨T"] = TRAN + Convert.ToDecimal(row["本期進貨T"]);
            }
            else if (TransType == "21" || TransType == "19")
            {
                row["進貨退出折讓"] = Qty + Convert.ToDecimal(row["進貨退出折讓"]);
                row["進貨退出折讓T"] = TRAN + Convert.ToDecimal(row["進貨退出折讓T"]);

            }
            else if (TransType == "15" || TransType == "13")
            {
                row["本期銷貨"] = Qty + Convert.ToDecimal(row["本期銷貨"]);
                row["本期銷貨T"] = TRAN + Convert.ToDecimal(row["本期銷貨T"]);

            }
            else if (TransType == "16" || TransType == "14")
            {
                row["銷貨退回"] = Qty + Convert.ToDecimal(row["銷貨退回"]);
                row["銷貨退回T"] = TRAN + Convert.ToDecimal(row["銷貨退回T"]);

            }
            else if (TransType == "59" || TransType == "60")
            {

                row["本期調整"] = Qty + Convert.ToDecimal(row["本期調整"]);
                row["本期調整T"] = TRAN + Convert.ToDecimal(row["本期調整T"]);
                
            }
            else if (TransType == "67")
            {
                row["本期調撥"] = Qty + Convert.ToDecimal(row["本期調撥"]);
                row["本期調撥T"] = TRAN + Convert.ToDecimal(row["本期調撥T"]);
            }

            row["期末存貨"] =
                Convert.ToDecimal(row["期初存貨"])
                + Convert.ToDecimal(row["本期進貨"])
                + Convert.ToDecimal(row["進貨退出折讓"])
                + Convert.ToDecimal(row["本期銷貨"])
                + Convert.ToDecimal(row["銷貨退回"])
                + Convert.ToDecimal(row["本期調整"])
                + Convert.ToDecimal(row["本期調撥"]);


            row["期末存貨T"] =
    Convert.ToDecimal(row["期初存貨T"])
    + Convert.ToDecimal(row["本期進貨T"])
    + Convert.ToDecimal(row["進貨退出折讓T"])
    + Convert.ToDecimal(row["本期銷貨T"])
    + Convert.ToDecimal(row["銷貨退回T"])
    + Convert.ToDecimal(row["本期調整T"])
    + Convert.ToDecimal(row["本期調撥T"]);

        }

        private void fmStockMove_Load(object sender, EventArgs e)
        {
            if ( globals.DBNAME != "進金生能源服務")
            {
                button1.Visible = false;
                button9.Visible = false;
                return;
            }
 
     
            //取前一個月
            DateTime PriorMonth = DateTime.Now.AddMonths(-1);


            int year = PriorMonth.Year;
            int month = PriorMonth.Month;

            //取得當月天數
            int days = DateTime.DaysInMonth(year, month);

            string d = DateToStr(PriorMonth);

            textBoxDocDate1.Text = d.Substring(0, 4) + d.Substring(4, 2) + "01";

            textBoxDocDate2.Text = d.Substring(0, 4) + d.Substring(4, 2) + days.ToString("00");

            EXEC();
        }

        //日期處理--------------------------------------------------------------------------------------------
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }


        int CountDays(DateTime dateFrom, DateTime dateTo, bool including)
        {
            return ((System.TimeSpan)(dateTo - dateFrom)).Days * (-1) + (including ? 1 : 0);
        }


        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }
        private System.Data.DataTable GetItemHisListByTransTypeD(string DocDate1, string DocDate2)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                                          SELECT DBO.fun_SAPDOC(TRANSTYPE) 單據總類,Convert(varchar(10),T0.[DocDate],111) 日期,BASE_REF 單號,T0.ITEMCODE 產品編號, ");
            sb.Append("                    ISNULL(CASE WHEN TRANSTYPE IN (18,20) THEN (T0.[InQty] - T0.[OutQty]) END,0) 本期進貨數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (18,20) THEN T0.[TransValue] END,0) 本期進貨金額, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (19,21) THEN (T0.[InQty] - T0.[OutQty]) END,0) 進貨退出折讓數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (19,21) THEN T0.[TransValue] END,0)  進貨退出折讓金額, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (13,15) THEN (T0.[InQty] - T0.[OutQty]) END,0) 本期銷貨數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (13,15) THEN T0.[TransValue] END,0)  本期銷貨金額, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (14,16) THEN (T0.[InQty] - T0.[OutQty]) END,0) 銷貨退回數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (14,16) THEN T0.[TransValue] END,0)  銷貨退回金額, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (59,60) THEN (T0.[InQty] - T0.[OutQty]) END,0) 本期調整數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (59,60) THEN T0.[TransValue] END,0) 本期調整金額, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (67) THEN (T0.[InQty] - T0.[OutQty]) END,0) 本期調撥數量, ");
            sb.Append("               ISNULL(CASE WHEN TRANSTYPE IN (67) THEN T0.[TransValue] END,0) 本期調撥金額 ");
            sb.Append("                                          FROM  [dbo].[OINM] T0     ");
            sb.Append("                                          INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode     ");
            sb.Append("                                          WHERE ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'      ");
            sb.Append("and  T0.[DocDate] >= @DocDate1 ");
            sb.Append("And    T0.[DocDate] <= @DocDate2 ");
            sb.Append(" order by  T0.[DocDate]");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }
        private void EXEC()
        {
 

            listBox1.Items.Clear();

            System.Data.DataTable dt = MakeTable();

            string DocDate1 = textBoxDocDate1.Text;


            string DocDate2 = textBoxDocDate2.Text;



            string PriorDate1 = DateToStr(StrToDate(DocDate1).AddDays(-1));

            System.Data.DataTable dtHist = GetItemHisList(PriorDate1);

            DataRow dr;

            for (int i = 0; i <= dtHist.Rows.Count - 1; i++)
            {
                dr = dt.NewRow();
                dr["BU"] = Convert.ToString(dtHist.Rows[i]["BU"]);
                dr["ITEM"] = Convert.ToString(dtHist.Rows[i]["ITEM"]);
                dr["產品編號"] = Convert.ToString(dtHist.Rows[i]["ItemCode"]);
                dr["期初存貨"] = Convert.ToDecimal(dtHist.Rows[i]["Qty"]);
                dr["期末存貨"] = Convert.ToDecimal(dtHist.Rows[i]["Qty"]);
                dr["本期進貨"] = 0;
                dr["進貨退出折讓"] = 0;
                dr["本期銷貨"] = 0;
                dr["銷貨退回"] = 0;
                dr["本期調整"] = 0;
                dr["本期調撥"] = 0;

                dr["期初存貨T"] = Convert.ToDecimal(dtHist.Rows[i]["TransValue"]);
                dr["期末存貨T"] = Convert.ToDecimal(dtHist.Rows[i]["TransValue"]);
                dr["本期進貨T"] = 0;
                dr["進貨退出折讓T"] = 0;
                dr["本期銷貨T"] = 0;
                dr["銷貨退回T"] = 0;
                dr["本期調整T"] = 0;
                dr["本期調撥T"] = 0;

                dt.Rows.Add(dr);
            }




            System.Data.DataTable dtNow = GetItemHisListByTransType(DocDate1, DocDate2);



            DataRow row;
            DataRow row2;
            string TransType;
            string ItemCode;
            decimal Qty;
            decimal TransValue;


            for (int i = 0; i <= dtNow.Rows.Count - 1; i++)
            {
                TransType = Convert.ToString(dtNow.Rows[i]["TransType"]);
                ItemCode = Convert.ToString(dtNow.Rows[i]["ItemCode"]);
                Qty = Convert.ToDecimal(dtNow.Rows[i]["Qty"]);
                TransValue = Convert.ToDecimal(dtNow.Rows[i]["TransValue"]);
                row = dt.Rows.Find(ItemCode);


                if (row != null)
                {
                    row.BeginEdit();

                    GoEdit(row, TransType, Qty, TransValue);

                    row.EndEdit();
                }
                else
                {
                    dr = dt.NewRow();
                    dr["BU"] = Convert.ToString(dtNow.Rows[i]["BU"]);
                    dr["ITEM"] = Convert.ToString(dtNow.Rows[i]["ITEM"]);
                    dr["產品編號"] = ItemCode;
                    dr["期初存貨"] = 0;
                    dr["期末存貨"] = 0;
                    dr["本期進貨"] = 0;
                    dr["進貨退出折讓"] = 0;
                    dr["本期銷貨"] = 0;
                    dr["銷貨退回"] = 0;
                    dr["本期調整"] = 0;
                    dr["本期調撥"] = 0;

                    dr["期初存貨T"] = 0;
                    dr["期末存貨T"] = 0;
                    dr["本期進貨T"] = 0;
                    dr["進貨退出折讓T"] = 0;
                    dr["本期銷貨T"] = 0;
                    dr["銷貨退回T"] = 0;
                    dr["本期調整T"] = 0;
                    dr["本期調撥T"] = 0;
                    GoEdit(dr, TransType, Qty, TransValue);

                    try
                    {
                        dt.Rows.Add(dr);
                    }
                    catch
                    {

                        listBox1.Items.Add(ItemCode);
                    }

                }
            }

            dt.DefaultView.RowFilter = "期初存貨  <> 0 or 期末存貨 <> 0 or 本期進貨 <> 0 or 進貨退出折讓 <> 0 or 本期銷貨 <> 0 or 銷貨退回 <> 0 or 本期調整 <> 0 or 本期調撥 <> 0 or 期初存貨T  <> 0 or 期末存貨T <> 0 or 本期進貨T <> 0 or 進貨退出折讓T <> 0 or 本期銷貨T <> 0 or 銷貨退回T <> 0 or 本期調整T <> 0 or 本期調撥T <> 0 ";


            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 1; j <= dt.Columns.Count - 1; j++)
                {
                    try
                    {
                        if (Convert.ToDecimal(dt.Rows[i]["期初存貨"]) != 0 || Convert.ToDecimal(dt.Rows[i]["本期進貨"]) != 0 || Convert.ToDecimal(dt.Rows[i]["進貨退出折讓"]) != 0
                            || Convert.ToDecimal(dt.Rows[i]["本期銷貨"]) != 0 || Convert.ToDecimal(dt.Rows[i]["銷貨退回"]) != 0 || Convert.ToDecimal(dt.Rows[i]["本期調整"]) != 0
                            || Convert.ToDecimal(dt.Rows[i]["本期調撥"]) != 0 || Convert.ToDecimal(dt.Rows[i]["期末存貨"]) != 0)
                        {
                            Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);
                        }
                    }
                    catch
                    {
                        Total[j - 1] += 0;
                    }

                }
            }

            row = dt.NewRow();

            row[2] = "合計";
            for (int j = 3; j <= dt.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);




            dataGridView1.DataSource = dt;


            dataGridView1.Columns[0].Width = 120;
            for (int i = 3; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
                col.Width = 110;

                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0.00";
            }

            if (globals.DBNAME == "進金生能源服務")
            {
                System.Data.DataTable T1 = GetItemHisListByTransTypeD(textBoxDocDate1.Text, textBoxDocDate2.Text);

                T1.DefaultView.RowFilter = "本期進貨數量 <>0 or 本期進貨金額 <> 0 or 進貨退出折讓數量 <> 0 or 進貨退出折讓金額 <> 0 or 本期銷貨數量 <> 0 or 本期銷貨金額 <> 0 or 銷貨退回數量 <> 0 or 銷貨退回金額 <> 0 or 本期調整數量  <> 0 or 本期調整金額 <> 0 or 本期調撥數量 <> 0 or 本期調撥金額 <> 0  ";
                //加入一筆合計
                decimal[] TotalD = new decimal[T1.Columns.Count - 1];

                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {

                    for (int j = 4; j <= T1.Columns.Count - 1; j++)
                    {
                        try
                        {
                            if (Convert.ToDecimal(T1.Rows[i]["本期進貨數量"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期進貨金額"]) != 0 || Convert.ToDecimal(T1.Rows[i]["進貨退出折讓數量"]) != 0
                                || Convert.ToDecimal(T1.Rows[i]["進貨退出折讓金額"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期銷貨數量"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期銷貨金額"]) != 0
                                || Convert.ToDecimal(T1.Rows[i]["銷貨退回數量"]) != 0 || Convert.ToDecimal(T1.Rows[i]["銷貨退回金額"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期調整數量"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期調整金額"]) != 0
                                 || Convert.ToDecimal(T1.Rows[i]["本期調撥數量"]) != 0 || Convert.ToDecimal(T1.Rows[i]["本期調撥金額"]) != 0)
                            {
                                TotalD[j - 1] += Convert.ToDecimal(T1.Rows[i][j]);
                            }
                        }
                        catch
                        {
                            TotalD[j - 1] += 0;
                        }

                    }
                }

                row2 = T1.NewRow();

                row2[3] = "合計";
                for (int j = 4; j <= T1.Columns.Count - 1; j++)
                {
                    row2[j] = TotalD[j - 1];

                }
                T1.Rows.Add(row2);

                dataGridView2.DataSource = T1;

                for (int i = 4; i <= dataGridView2.Columns.Count - 1; i++)
                {
                    DataGridViewColumn col = dataGridView2.Columns[i];
                    col.Width = 110;

                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.00";
                }
            }



            dataGridView1.Columns[3].HeaderText = "期初存貨數量";
            dataGridView1.Columns[4].HeaderText = "期初存貨金額";
            dataGridView1.Columns[5].HeaderText = "本期進貨數量";
            dataGridView1.Columns[6].HeaderText = "本期進貨金額";
            dataGridView1.Columns[7].HeaderText = "進貨退出折讓數量";
            dataGridView1.Columns[8].HeaderText = "進貨退出折讓金額";
            dataGridView1.Columns[9].HeaderText = "本期銷貨數量";
            dataGridView1.Columns[10].HeaderText = "本期銷貨金額";
            dataGridView1.Columns[11].HeaderText = "銷貨退回數量";
            dataGridView1.Columns[12].HeaderText = "銷貨退回金額";
            dataGridView1.Columns[13].HeaderText = "本期調整數量";
            dataGridView1.Columns[14].HeaderText = "本期調整金額";
            dataGridView1.Columns[15].HeaderText = "本期調撥數量";
            dataGridView1.Columns[16].HeaderText = "本期調撥金額";
            dataGridView1.Columns[17].HeaderText = "期末存貨數量";
            dataGridView1.Columns[18].HeaderText = "期末存貨金額";

        }
        private void button9_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
               ExcelReport.GridViewToExcel(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }

  



        private void button1_Click_1(object sender, EventArgs e)
        {
            EXEC();
        }

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                string DOCTYPE = comboBox1.Text;
                if (!String.IsNullOrEmpty(DOCTYPE))
                {
                    fmStockMoveD frm = new fmStockMoveD();
                    frm.DATETIME1 = textBoxDocDate1.Text;
                    frm.DATETIME2 = textBoxDocDate2.Text;
                    frm.ITEMCODE = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                    frm.DOCTYPE = DOCTYPE;
                    frm.ShowDialog();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

            }
        }

     
    }
}
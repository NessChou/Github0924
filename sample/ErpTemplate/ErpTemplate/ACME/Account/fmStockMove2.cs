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
    public partial class fmStockMove2 : Form
    {
        private string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";

        public fmStockMove2()
        {
            InitializeComponent();
        }


        private System.Data.DataTable GetItemHisList(string DocDate)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append("                     SELECT U_SIZE 尺寸 ,SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue  ");
            sb.Append("               FROM  [dbo].[OINM] T0   ");
            sb.Append("                           INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("             WHERE  T0.[DocDate] <= @DocDate ");
            sb.Append("               AND  T1.FROZENFOR = 'N' AND  ITMSGRPCOD=1032 AND     T1.U_GROUP='100-Panel' ");
            sb.Append("                           GROUP BY U_SIZE");
            sb.Append("                               ORDER BY  CAST(U_SIZE AS DECIMAL(10,2))  ");

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
        private System.Data.DataTable GetItemHisListByTransType(string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();



            sb.Append("                           SELECT U_SIZE  尺寸,T0.[TransType],SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue  ");
            sb.Append("                           FROM  [dbo].[OINM] T0   ");
            sb.Append("                           INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("                           WHERE  T0.[DocDate] >= @DocDate1  ");
            sb.Append("                           And    T0.[DocDate] <= @DocDate2   ");
            sb.Append(" AND  T1.FROZENFOR = 'N' AND  ");
            sb.Append("                ITMSGRPCOD=1032 AND     T1.U_GROUP='100-Panel' ");
            sb.Append("                           GROUP BY  U_SIZE ,T0.[TransType]  ");
            sb.Append("                            ORDER BY  CAST(U_SIZE AS DECIMAL(10,2))  ");

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

        private System.Data.DataTable GetOPOR(string ITEMCODE, string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

     

            sb.Append("               SELECT U_SIZE 尺寸,CAST(AVG(T1.TOTALFRGN/T1.QUANTITY) AS INT) ");
            sb.Append("               FROM OPOR T0  ");
            sb.Append("               INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                INNER  JOIN [dbo].[OITM] T2  ON  T1.[ItemCode] = T2.ItemCode  ");
            sb.Append("               left join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
            sb.Append("               left join opdn t44 on (t4.docentry=t44.docentry) ");
            sb.Append("               WHERE   T1.[LINESTATUS] ='C' and T1.trgetentry <>''  and  ISNULL(T2.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            sb.Append("               AND  T2.FROZENFOR = 'N' AND T2.ITMSGRPCOD=1032 AND     T2.U_GROUP='100-Panel'AND T1.CURRENCY='USD' ");
            sb.Append("               AND  U_SIZE=@ITEMCODE ");
            sb.Append("                           AND  t44.[DocDate] >= @DocDate1   ");
            sb.Append("                           And    t44.[DocDate] <= @DocDate2   ");
            sb.Append("               GROUP BY U_SIZE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        private System.Data.DataTable GetODLN(string ITEMCODE, string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();



            sb.Append("               SELECT U_SIZE 尺寸,CAST(AVG(T1.TOTALFRGN/T1.QUANTITY) AS INT) ");
            sb.Append("               FROM ORDR T0  ");
            sb.Append("               INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                INNER  JOIN [dbo].[OITM] T2  ON  T1.[ItemCode] = T2.ItemCode  ");
            sb.Append("               left join DLN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
            sb.Append("               left join ODLN t44 on (t4.docentry=t44.docentry) ");
            sb.Append("               WHERE   T1.[LINESTATUS] ='C' and T1.trgetentry <>'' and  ISNULL(T2.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            sb.Append("               AND  T2.FROZENFOR = 'N' AND T2.ITMSGRPCOD=1032 AND     T2.U_GROUP='100-Panel'AND T1.CURRENCY='USD' ");
            sb.Append("               AND U_SIZE=@ITEMCODE ");
            sb.Append("                           AND  t44.[DocDate] >= @DocDate1   ");
            sb.Append("                           And    t44.[DocDate] <= @DocDate2    ");
            sb.Append("               GROUP BY U_SIZE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        private System.Data.DataTable GetORDR()
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (T0.[CardName]) 客戶名稱,T0.[DocNum] 訂單號碼 ,T1.ITEMCODE 項目號碼,T1.DSCRIPTION 項目名稱,T1.PRICE 單價,T1.QUANTITY 數量,T1.TOTALFRGN 金額,Convert(varchar(10),t1.[ShipDate],111)  預計交貨日期");
            sb.Append(" FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE ");
            sb.Append(" WHERE    T1.[LINESTATUS] ='O' AND  T11.FROZENFOR = 'N' AND  T11.ITMSGRPCOD=1032 AND     T11.U_GROUP='100-Panel'");
            sb.Append(" Order by T1.ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

      
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
            dt.Columns.Add("尺寸", typeof(decimal));
            dt.Columns.Add("平均進貨價格", typeof(Int32));
            dt.Columns.Add("平均銷貨價格", typeof(Int32));
            dt.Columns.Add("期初存貨", typeof(Int32));
            dt.Columns.Add("本期進貨", typeof(Int32));
            dt.Columns.Add("進貨退出折讓", typeof(Int32));
            dt.Columns.Add("本期銷貨", typeof(Int32));
            dt.Columns.Add("銷貨退回", typeof(Int32));
            dt.Columns.Add("本期調整", typeof(Int32));
            dt.Columns.Add("本期調撥", typeof(Int32));
            dt.Columns.Add("期末存貨", typeof(Int32));
            dt.Columns.Add("期初存貨T", typeof(Int32));
            dt.Columns.Add("本期進貨T", typeof(Int32));
            dt.Columns.Add("進貨退出折讓T", typeof(Int32));
            dt.Columns.Add("本期銷貨T", typeof(Int32));
            dt.Columns.Add("銷貨退回T", typeof(Int32));
            dt.Columns.Add("本期調整T", typeof(Int32));
            dt.Columns.Add("本期調撥T", typeof(Int32));
            dt.Columns.Add("期末存貨TT", typeof(Int32));
            dt.Columns.Add("期末存貨T", typeof(Int32));
            //平均進貨價格

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["尺寸"];
            dt.PrimaryKey = colPk;


            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void GoEdit(DataRow row, string TransType, Int32 Qty, Int32 TRAN)
        {

            if (TransType == "20" || TransType == "18")
            {
                row["本期進貨"] = Qty +Convert.ToInt32(row["本期進貨"]);
                row["本期進貨T"] = TRAN + Convert.ToInt32(row["本期進貨T"]);

            }
            else if (TransType == "21" || TransType == "19")
            {
                row["進貨退出折讓"] = Qty + Convert.ToInt32(row["進貨退出折讓"]);
                row["進貨退出折讓T"] = TRAN + Convert.ToInt32(row["進貨退出折讓T"]);

            }
            else if (TransType == "15" || TransType == "13")
            {
                row["本期銷貨"] = Qty + Convert.ToInt32(row["本期銷貨"]);
                row["本期銷貨T"] = TRAN + Convert.ToInt32(row["本期銷貨T"]);

            }
            else if (TransType == "16" || TransType == "14")
            {
                row["銷貨退回"] = Qty + Convert.ToInt32(row["銷貨退回"]);
                row["銷貨退回T"] = TRAN + Convert.ToInt32(row["銷貨退回T"]);

            }
            else if (TransType == "59" || TransType == "60")
            {

                row["本期調整"] = Qty + Convert.ToInt32(row["本期調整"]);
                row["本期調整T"] = TRAN + Convert.ToInt32(row["本期調整T"]);
                
            }
            else if (TransType == "67")
            {
                row["本期調撥"] = Qty + Convert.ToInt32(row["本期調撥"]);
                row["本期調撥T"] = TRAN + Convert.ToInt32(row["本期調撥T"]);
            }

            row["期末存貨"] = 
                Convert.ToInt32(row["期初存貨"]) 
                + Convert.ToInt32(row["本期進貨"])
                + Convert.ToInt32(row["進貨退出折讓"])
                +Convert.ToInt32(row["本期銷貨"])
                + Convert.ToInt32(row["銷貨退回"])
                + Convert.ToInt32(row["本期調整"]) 
                + Convert.ToInt32(row["本期調撥"]);

            row["期末存貨TT"] =
Convert.ToInt32(row["期初存貨"])
+ Convert.ToInt32(row["本期進貨"])
+ Convert.ToInt32(row["進貨退出折讓"])
+ Convert.ToInt32(row["本期銷貨"])
+ Convert.ToInt32(row["銷貨退回"])
+ Convert.ToInt32(row["本期調整"])
+ Convert.ToInt32(row["本期調撥"]);

            row["期末存貨T"] =
    Convert.ToInt32(row["期初存貨T"])
    + Convert.ToInt32(row["本期進貨T"])
    + Convert.ToInt32(row["進貨退出折讓T"])
    + Convert.ToInt32(row["本期銷貨T"])
    + Convert.ToInt32(row["銷貨退回T"])
    + Convert.ToInt32(row["本期調整T"])
    + Convert.ToInt32(row["本期調撥T"]);

        }

        private void fmStockMove_Load(object sender, EventArgs e)
        {
            //textBoxDocDate2.Text = DateTime.Now.ToString("yyyyMMdd");

            //取前一個月
            DateTime PriorMonth = DateTime.Now.AddMonths(-1);


            int year = PriorMonth.Year;
            int month = PriorMonth.Month;

            //取得當月天數
            int days = DateTime.DaysInMonth(year, month);

            string d = DateToStr(PriorMonth);

            textBoxDocDate1.Text = d.Substring(0, 4) + d.Substring(4, 2) + "01";

            textBoxDocDate2.Text = d.Substring(0, 4) + d.Substring(4, 2) + days.ToString("00");

            dataGridView2.DataSource = GetORDR();

            dataGridView2.Columns[0].Width = 150;
            dataGridView2.Columns[1].Width = 80;
            dataGridView2.Columns[2].Width = 150;
            dataGridView2.Columns[3].Width = 150;
            for (int i = 4; i <= dataGridView2.Columns.Count - 2; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];
                col.Width = 60;

                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
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

        private void button3_Click(object sender, EventArgs e)
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
                dr["尺寸"] = Convert.ToString(dtHist.Rows[i]["尺寸"]);
                dr["期初存貨"] = Convert.ToInt32(dtHist.Rows[i]["Qty"]);
                dr["期末存貨"] = Convert.ToInt32(dtHist.Rows[i]["Qty"]);
                dr["本期進貨"] = 0;
                dr["進貨退出折讓"] = 0;
                dr["本期銷貨"] = 0;
                dr["銷貨退回"] = 0;
                dr["本期調整"] = 0;
                dr["本期調撥"] = 0;
                dr["期初存貨T"] = Convert.ToInt32(dtHist.Rows[i]["TransValue"]);
                dr["期末存貨TT"] = Convert.ToInt32(dtHist.Rows[i]["Qty"]);
                dr["期末存貨T"] = Convert.ToInt32(dtHist.Rows[i]["TransValue"]);
                dr["本期進貨T"] = 0;
                dr["進貨退出折讓T"] = 0;
                dr["本期銷貨T"] = 0;
                dr["銷貨退回T"] = 0;
                dr["本期調整T"] = 0;
                dr["本期調撥T"] = 0;
                string SIZE = dr["尺寸"].ToString();
                System.Data.DataTable H1 = GetOPOR(SIZE, DocDate1, DocDate2);
                System.Data.DataTable H2 = GetODLN(SIZE, DocDate1, DocDate2);
                string INPRICE = "";
                string OUPRICE = "";
                if (H1.Rows.Count > 0)
                {
                    INPRICE = H1.Rows[0][1].ToString();
                }
                if (String.IsNullOrEmpty(INPRICE))
                {
                    INPRICE = "0";
                }
                if (H2.Rows.Count > 0)
                {
                    OUPRICE = H2.Rows[0][1].ToString();
                }
                if (String.IsNullOrEmpty(OUPRICE))
                {
                    OUPRICE = "0";
                }
                dr["平均進貨價格"] = INPRICE;
                dr["平均銷貨價格"] = OUPRICE;
                dt.Rows.Add(dr);
            }


            System.Data.DataTable dtNow = GetItemHisListByTransType(DocDate1, DocDate2);



            DataRow row;
            string TransType;
            string 尺寸;
            Int32 Qty;
            Int32 TransValue;


            for (int i = 0; i <= dtNow.Rows.Count - 1; i++)
            {
                TransType = Convert.ToString(dtNow.Rows[i]["TransType"]);
                尺寸 = Convert.ToString(dtNow.Rows[i]["尺寸"]);
                Qty = Convert.ToInt32(dtNow.Rows[i]["Qty"]);
                TransValue = Convert.ToInt32(dtNow.Rows[i]["TransValue"]);
                row = dt.Rows.Find(尺寸);


                if (row != null)
                {
                    row.BeginEdit();

                    GoEdit(row, TransType, Qty, TransValue);

                    row.EndEdit();
                }
                else
                {
                    dr = dt.NewRow();
                    dr["尺寸"] = 尺寸;
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
                    dr["期末存貨TT"] = 0;
                    dr["本期進貨T"] = 0;
                    dr["進貨退出折讓T"] = 0;
                    dr["本期銷貨T"] = 0;
                    dr["銷貨退回T"] = 0;
                    dr["本期調整T"] = 0;
                    dr["本期調撥T"] = 0;

                    string SIZE = 尺寸;
                    System.Data.DataTable H1 = GetOPOR(SIZE, DocDate1, DocDate2);
                    System.Data.DataTable H2 = GetODLN(SIZE, DocDate1, DocDate2);
                    string INPRICE = "";
                    string OUPRICE = "";
                    if (H1.Rows.Count > 0)
                    {
                        INPRICE = H1.Rows[0][1].ToString();
                    }
                    if (String.IsNullOrEmpty(INPRICE))
                    {
                        INPRICE = "0";
                    }
                    if (H2.Rows.Count > 0)
                    {
                        OUPRICE = H2.Rows[0][1].ToString();
                    }
                    if (String.IsNullOrEmpty(OUPRICE))
                    {
                        OUPRICE = "0";
                    }
                    dr["平均進貨價格"] = INPRICE;
                    dr["平均銷貨價格"] = OUPRICE;
                    GoEdit(dr, TransType, Qty, TransValue);

                    try
                    {
                        dt.Rows.Add(dr);
                    }
                    catch
                    {

                        listBox1.Items.Add(尺寸);
                    }

                }
            }

            //dt.DefaultView.Sort = "尺寸";

            //加入一筆合計
            Int32[] Total = new Int32[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 3; j <= dt.Columns.Count - 1; j++)
                {
                    try
                    {
                        Total[j - 1] += Convert.ToInt32(dt.Rows[i][j]);
                    }
                    catch
                    {
                        Total[j - 1] += 0;
                    }

                }
            }



            row = dt.NewRow();


            row[0] = "1000";

            for (int j = 3; j <= dt.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);

            // 0 不顯示

            DataView dv = dt.DefaultView;

            dv.RowFilter = "期初存貨 >0 or 本期進貨>0 or 進貨退出折讓>0 or 本期銷貨>0 or 銷貨退回>0 or 本期調整>0 or 本期調撥>0 or 期末存貨>0 ";

            dv.Sort = "尺寸";

            dataGridView1.DataSource = dv;

            dataGridView1.Columns[0].Width = 60;

            for (int i = 1; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
 
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }

            if (radioButton1.Checked)
            {
                dataGridView1.Columns[3].Visible = true;
                dataGridView1.Columns[4].Visible = true;
                dataGridView1.Columns[5].Visible = true;
                dataGridView1.Columns[6].Visible = true;
                dataGridView1.Columns[7].Visible = true;
                dataGridView1.Columns[8].Visible = true;
                dataGridView1.Columns[9].Visible = true;
                dataGridView1.Columns[10].Visible = true;
                dataGridView1.Columns[11].Visible = false;
                dataGridView1.Columns[12].Visible = false;
                dataGridView1.Columns[13].Visible = false;
                dataGridView1.Columns[14].Visible = false;
                dataGridView1.Columns[15].Visible = false;
                dataGridView1.Columns[16].Visible = false;
                dataGridView1.Columns[17].Visible = false;
                dataGridView1.Columns[18].Visible = false;
                dataGridView1.Columns[19].Visible = false;


            }
            else
            {
                dataGridView1.Columns[3].Visible = true;
                dataGridView1.Columns[4].Visible = false;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[6].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[9].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = true;
                dataGridView1.Columns[12].Visible = true;
                dataGridView1.Columns[13].Visible = true;
                dataGridView1.Columns[14].Visible = true;
                dataGridView1.Columns[15].Visible = true;
                dataGridView1.Columns[16].Visible = true;
                dataGridView1.Columns[17].Visible = true;
                dataGridView1.Columns[18].Visible = true;
                dataGridView1.Columns[19].Visible = true;

                dataGridView1.Columns[3].HeaderText = "期初存貨數量";
                dataGridView1.Columns[11].HeaderText = "期初存貨金額";
                dataGridView1.Columns[12].HeaderText = "本期進貨";
                dataGridView1.Columns[13].HeaderText = "進貨退出折讓";
                dataGridView1.Columns[14].HeaderText = "本期銷貨";
                dataGridView1.Columns[15].HeaderText = "銷貨退回";
                dataGridView1.Columns[16].HeaderText = "本期調整";
                dataGridView1.Columns[17].HeaderText = "本期調撥";
                dataGridView1.Columns[18].HeaderText = "期末存貨數量";
                dataGridView1.Columns[19].HeaderText = "期末存貨金額";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                GridViewToExcel(dataGridView1);
            }
            else
            {
                GridViewToExcel2(dataGridView1);
            }
        }
        public static void GridViewToExcel(DataGridView dgv)
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

                for (int i = 0; i <= 10; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j <= 10; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

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
        public static void GridViewToExcel2(DataGridView dgv)
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
                wsheet.Cells[1, 1] = dgv.Columns[0].HeaderText;
                wsheet.Cells[1, 2] = dgv.Columns[1].HeaderText;
                wsheet.Cells[1, 3] = dgv.Columns[2].HeaderText;
                wsheet.Cells[1, 4] = dgv.Columns[3].HeaderText;
                int h = 0;
                for (int i = 11; i <= 19; i++)
                {
                    h++;

                    wsheet.Cells[1, h + 4] = dgv.Columns[i].HeaderText;

                }




                for (int i = 0; i < dgv.Rows.Count; i++)
                {


                    DataGridViewRow row = dgv.Rows[i];

                    DataGridViewCell cell2 = row.Cells[0];
                    DataGridViewCell cell3 = row.Cells[1];
                    DataGridViewCell cell4 = row.Cells[2];
                    DataGridViewCell cell5 = row.Cells[3];
                    wsheet.Cells[i + 2, 1] = (cell2.Value == null) ? "" : cell2.Value.ToString();
                    wsheet.Cells[i + 2, 2] = (cell3.Value == null) ? "" : cell3.Value.ToString();
                    wsheet.Cells[i + 2, 3] = (cell4.Value == null) ? "" : cell4.Value.ToString();
                    wsheet.Cells[i + 2, 4] = (cell5.Value == null) ? "" : cell5.Value.ToString();
                    int h2 = 0;
                    for (int j = 11; j <= 19; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {
                            h2++;
                            wsheet.Cells[i + 2, h2 + 4] = (cell.Value == null) ? "" : cell.Value.ToString();

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
        private void button1_Click_1(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }
    }
}
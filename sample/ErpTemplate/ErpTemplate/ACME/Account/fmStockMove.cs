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
    public partial class fmStockMove : Form
    {
        private string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";

        public fmStockMove()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
         
            string DocDate = textBoxDocDate1.Text;
            string Warehouse1 = textBoxWh1.Text;
            string Warehouse2 = textBoxWh2.Text;

            System.Data.DataTable dt = GetItemHisList(DocDate, Warehouse1, Warehouse2);

            dataGridView1.DataSource = dt;
        }


        //取得某一時點的庫存列表
        //T0.[TransType] = 162  -> Inventory Valuation 
        private System.Data.DataTable GetItemHisList(string DocDate, string Warehouse1, string Warehouse2)
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
            sb.Append("and ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  AND  T0.[ItemCode] NOT LIKE '%-C%'  ");
            if (!checkBox1.Checked)
            {
                sb.Append("AND  T0.[Warehouse] >= @Warehouse1  ");
                sb.Append("AND  T0.[Warehouse] <= @Warehouse2  ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("AND  T0.[Warehouse] NOT IN ('WO401','WO402')  ");
            }
  
            sb.Append("GROUP BY  T0.[ItemCode], T1.[ItemName] ");
           // sb.Append("Having SUM(T0.[InQty] - T0.[OutQty])> 0 ");
            sb.Append("ORDER BY  T0.[ItemCode]");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));
            command.Parameters.Add(new SqlParameter("@Warehouse1", Warehouse1));
            command.Parameters.Add(new SqlParameter("@Warehouse2", Warehouse2));
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
        private System.Data.DataTable GetItemHisListByTransType( string DocDate1, string DocDate2, string Warehouse1, string Warehouse2)
        {
         
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT T0.[ItemCode], T1.[ItemName],T0.[TransType],SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue,MAX(SUBSTRING(ITMSGRPNAM,4,LEN(ITMSGRPNAM)-3)) BU,MAX(ltrim(substring(U_GROUP,CHARINDEX('-', U_GROUP)+1,LEN(U_GROUP)))) ITEM ");
            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("INNER  JOIN [dbo].[OITB] T2  ON  T2.[ITMSGRPCOD] = T1.ITMSGRPCOD  ");
            sb.Append("WHERE  T0.[DocDate] >= @DocDate1 ");
            sb.Append("And    T0.[DocDate] <= @DocDate2 ");
            sb.Append("and ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  AND  T0.[ItemCode] NOT LIKE '%-C%'   ");
            if (!checkBox1.Checked)
            {
                sb.Append("AND  T0.[Warehouse] >= @Warehouse1  ");
                sb.Append("AND  T0.[Warehouse] <= @Warehouse2  ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("AND  T0.[Warehouse] NOT IN ('WO401','WO402')  ");
            }

            sb.Append("GROUP BY  T0.[ItemCode], T1.[ItemName],T0.[TransType] ");
            sb.Append("ORDER BY  T0.[ItemCode]");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
  
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            command.Parameters.Add(new SqlParameter("@Warehouse1", Warehouse1));
            command.Parameters.Add(new SqlParameter("@Warehouse2", Warehouse2));
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
            dt.Columns.Add("本期進貨", typeof(decimal));
            dt.Columns.Add("進貨退出折讓", typeof(Int64));
            dt.Columns.Add("本期銷貨", typeof(decimal));
            dt.Columns.Add("銷貨退回", typeof(decimal));
            dt.Columns.Add("本期調整", typeof(decimal));
            dt.Columns.Add("本期調撥", typeof(decimal));
            dt.Columns.Add("期末存貨", typeof(decimal));
            dt.Columns.Add("期初存貨T", typeof(Int64));
            dt.Columns.Add("本期進貨T", typeof(Int64));
            dt.Columns.Add("進貨退出折讓T", typeof(Int64));
            dt.Columns.Add("本期銷貨T", typeof(Int64));
            dt.Columns.Add("銷貨退回T", typeof(Int64));
            dt.Columns.Add("本期調整T", typeof(Int64));
            dt.Columns.Add("本期調撥T", typeof(Int64));
            dt.Columns.Add("期末存貨TT", typeof(decimal));
            dt.Columns.Add("期末存貨T", typeof(Int64));
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["產品編號"];
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
                row["本期進貨"] = Qty + Convert.ToDecimal(row["本期進貨"]);
                row["本期進貨T"] = TRAN + Convert.ToInt32(row["本期進貨T"]);
            }
            else if (TransType == "21" || TransType == "19")
            {
                row["進貨退出折讓"] = Qty + Convert.ToDecimal(row["進貨退出折讓"]);
                row["進貨退出折讓T"] = TRAN + Convert.ToInt32(row["進貨退出折讓T"]);

            }
            else if (TransType == "15" || TransType == "13")
            {
                row["本期銷貨"] = Qty + Convert.ToDecimal(row["本期銷貨"]);
                row["本期銷貨T"] = TRAN + Convert.ToInt32(row["本期銷貨T"]);

            }
            else if (TransType == "16" || TransType == "14")
            {
                row["銷貨退回"] = Qty + Convert.ToDecimal(row["銷貨退回"]);
                row["銷貨退回T"] = TRAN + Convert.ToInt32(row["銷貨退回T"]);

            }
            else if (TransType == "59" || TransType == "60")
            {

                row["本期調整"] = Qty + Convert.ToDecimal(row["本期調整"]);
                row["本期調整T"] = TRAN + Convert.ToInt32(row["本期調整T"]);
                
            }
            else if (TransType == "67")
            {
                row["本期調撥"] = Qty + Convert.ToDecimal(row["本期調撥"]);
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
    Convert.ToDecimal(row["期初存貨"])
    + Convert.ToDecimal(row["本期進貨"])
    + Convert.ToDecimal(row["進貨退出折讓"])
    + Convert.ToDecimal(row["本期銷貨"])
    + Convert.ToDecimal(row["銷貨退回"])
    + Convert.ToDecimal(row["本期調整"])
    + Convert.ToDecimal(row["本期調撥"]);

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
            if (globals.GroupID.ToString().Trim() == "ACCS" )
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

        private void button2_Click(object sender, EventArgs e)
        {

            string DocDate1 = textBoxDocDate1.Text;
            string DocDate2 = textBoxDocDate2.Text;

            string Warehouse1 = textBoxWh1.Text;
            string Warehouse2 = textBoxWh2.Text;


            System.Data.DataTable dtNow = GetItemHisListByTransType(DocDate1, DocDate2, Warehouse1, Warehouse2);

            dataGridView1.DataSource = dtNow;
        }



        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void EXEC()
        {

            listBox1.Items.Clear();

            System.Data.DataTable dt = MakeTable();

            string DocDate1 = textBoxDocDate1.Text;
            string Warehouse1 = textBoxWh1.Text;
            string Warehouse2 = textBoxWh2.Text;

            string DocDate2 = textBoxDocDate2.Text;



            string PriorDate1 = DateToStr(StrToDate(DocDate1).AddDays(-1));

            System.Data.DataTable dtHist = GetItemHisList(PriorDate1, Warehouse1, Warehouse2);

            DataRow dr;

            for (int i = 0; i <= dtHist.Rows.Count - 1; i++)
            {
                dr = dt.NewRow();
                dr["BU"] = Convert.ToString(dtHist.Rows[i]["BU"]);
                dr["ITEM"] = Convert.ToString(dtHist.Rows[i]["ITEM"]);
                dr["產品編號"] = Convert.ToString(dtHist.Rows[i]["ItemCode"]);
                //if (Convert.ToString(dtHist.Rows[i]["ItemCode"]) == "41609012.601.01")
                //{
                //    MessageBox.Show("A");
                //}
                dr["期初存貨"] = Convert.ToDecimal(dtHist.Rows[i]["Qty"]);
                dr["期末存貨"] = Convert.ToDecimal(dtHist.Rows[i]["Qty"]);
                dr["本期進貨"] = 0;
                dr["進貨退出折讓"] = 0;
                dr["本期銷貨"] = 0;
                dr["銷貨退回"] = 0;
                dr["本期調整"] = 0;
                dr["本期調撥"] = 0;

                dr["期初存貨T"] = Convert.ToInt32(dtHist.Rows[i]["TransValue"]);
                dr["期末存貨TT"] = Convert.ToDecimal(dtHist.Rows[i]["Qty"]);
                dr["期末存貨T"] = Convert.ToInt32(dtHist.Rows[i]["TransValue"]);
                dr["本期進貨T"] = 0;
                dr["進貨退出折讓T"] = 0;
                dr["本期銷貨T"] = 0;
                dr["銷貨退回T"] = 0;
                dr["本期調整T"] = 0;
                dr["本期調撥T"] = 0;

                dt.Rows.Add(dr);
            }




            System.Data.DataTable dtNow = GetItemHisListByTransType(DocDate1, DocDate2, Warehouse1, Warehouse2);



            DataRow row;
            string TransType;
            string ItemCode;
            Int32 Qty;
            Int32 TransValue;


            for (int i = 0; i <= dtNow.Rows.Count - 1; i++)
            {
                TransType = Convert.ToString(dtNow.Rows[i]["TransType"]);
                ItemCode = Convert.ToString(dtNow.Rows[i]["ItemCode"]);
                Qty = Convert.ToInt32(dtNow.Rows[i]["Qty"]);
                TransValue = Convert.ToInt32(dtNow.Rows[i]["TransValue"]);
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
                    dr["期末存貨TT"] = 0;
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

            dt.DefaultView.RowFilter = "期初存貨 >0 or 本期進貨>0 or 進貨退出折讓>0 or 本期銷貨>0 or 銷貨退回>0 or 本期調整>0 or 本期調撥>0 or 期末存貨>0 ";


            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 1; j <= dt.Columns.Count - 1; j++)
                {
                    try
                    {
                        if (Convert.ToDecimal(dt.Rows[i]["期初存貨"]) > 0 || Convert.ToDecimal(dt.Rows[i]["本期進貨"]) > 0 || Convert.ToDecimal(dt.Rows[i]["進貨退出折讓"]) > 0
                            || Convert.ToDecimal(dt.Rows[i]["本期銷貨"]) > 0 || Convert.ToDecimal(dt.Rows[i]["銷貨退回"]) > 0 || Convert.ToDecimal(dt.Rows[i]["本期調整"]) > 0
                            || Convert.ToDecimal(dt.Rows[i]["本期調撥"]) > 0 || Convert.ToDecimal(dt.Rows[i]["期末存貨"]) > 0)
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


            dataGridView1.Columns[0].Width = 150;
            for (int i = 3; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
                col.Width = 110;

                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0.0";
            }


            dataGridView2.DataSource = dt;


            dataGridView2.Columns[0].Width = 150;
            for (int i = 3; i <= dataGridView2.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];
                col.Width = 110;

                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
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
          

            
        
                dataGridView2.Columns[3].Visible = true;
                dataGridView2.Columns[4].Visible = false;
                dataGridView2.Columns[5].Visible = false;
                dataGridView2.Columns[6].Visible = false;
                dataGridView2.Columns[7].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[9].Visible = false;
                dataGridView2.Columns[10].Visible = false;
                dataGridView2.Columns[11].Visible = true;
                dataGridView2.Columns[12].Visible = true;
                dataGridView2.Columns[13].Visible = true;
                dataGridView2.Columns[14].Visible = true;
                dataGridView2.Columns[15].Visible = true;
                dataGridView2.Columns[16].Visible = true;
                dataGridView2.Columns[17].Visible = true;
                dataGridView2.Columns[18].Visible = true;
                dataGridView2.Columns[19].Visible = true;

                dataGridView1.Columns[0].HeaderText = "群組";
                dataGridView1.Columns[1].HeaderText = "產品分類";
                dataGridView2.Columns[0].HeaderText = "群組";
                dataGridView2.Columns[1].HeaderText = "產品分類";
                dataGridView2.Columns[3].HeaderText = "期初存貨數量";
                dataGridView2.Columns[11].HeaderText = "期初存貨金額";
                dataGridView2.Columns[12].HeaderText = "本期進貨";
                dataGridView2.Columns[13].HeaderText = "進貨退出折讓";
                dataGridView2.Columns[14].HeaderText = "本期銷貨";
                dataGridView2.Columns[15].HeaderText = "銷貨退回";
                dataGridView2.Columns[16].HeaderText = "本期調整";
                dataGridView2.Columns[17].HeaderText = "本期調撥";
                dataGridView2.Columns[18].HeaderText = "期末存貨數量";
                dataGridView2.Columns[19].HeaderText = "期末存貨金額";
            
        }
        private void button9_Click(object sender, EventArgs e)
        {
            GridViewToExcel(dataGridView1);
            //if (radioButton1.Checked)
            //{
            //    GridViewToExcel(dataGridView1);
            //}
            //else
            //{
            //    GridViewToExcel2(dataGridView1);
            //}
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
            EXEC();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GridViewToExcel2(dataGridView2);
        }

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView2.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

     
    }
}
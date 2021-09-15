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
namespace ACME
{
    public partial class fmAcmeFIFO : Form
    {
                

        public fmAcmeFIFO()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {


            string DocDate = textBoxDocDate.Text;

            System.Data.DataTable dt = null;
            System.Data.DataTable dt2 = null;

            dt = GetStockListToAge(DocDate, "Qty");

            dt2 = GetStockListToAge(DocDate, "Amt");
            

            dataGridView1.DataSource = dt;

            dataGridView2.DataSource = dt2;


            for (int i = 3; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0.0";
            }


            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 3; j <= dt.Columns.Count - 1; j++)
                {
                    Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);

                }
            }

            DataRow row;

            row = dt.NewRow();

            row[2] = "合計";
            for (int j = 3; j <= dt.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);




            //加入一筆合計
            Int32[] Total2 = new Int32[dt2.Columns.Count - 1];

            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {
               

                for (int j = 3; j <= dt2.Columns.Count - 1; j++)
                {
                    Total2[j - 1] += Convert.ToInt32(dt2.Rows[i][j]);

                }
            }


            decimal h1 = 0;
            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {


                h1 += Convert.ToDecimal(dt2.Rows[i][3]);

                
            }

            DataRow row2;

            row2 = dt2.NewRow();

            row2[2] = "合計";
            for (int j = 3; j <= dt2.Columns.Count - 1; j++)
            {
                if (j == 3)
                {
                    row2[3] = h1;

                }
                else
                {
                    row2[j] = Total2[j - 1];
                }

            }
            dt2.Rows.Add(row2);


            for (int i = 3; i <= dataGridView2.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];
                if (i == 3)
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0.0";
                }
                else
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0";
                }
            }

        }
        private System.Data.DataTable GetStockListToAge(string DocDate, string Mode)
        {




            System.Data.DataTable dt = GetStockList(DocDate);


            System.Data.DataTable dtStock = null;

            if (Mode == "Qty")
            {
                dtStock = MakeTable_Stock2();
            }
            else
            {
                dtStock = MakeTable_Stock();
            }

            DataRow row;
            DataRow rowFind;
            string ItemCode;
            Int32 StockDays;
            decimal  Qty;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                ItemCode = Convert.ToString(dt.Rows[i]["產品編號"]);
      
                StockDays = Convert.ToInt32(dt.Rows[i]["庫存天數"]);
          
                Qty = Convert.ToDecimal(dt.Rows[i]["庫存量"]);

            

                row = dtStock.Rows.Find(ItemCode);


                if (row != null)
                {
                    row.BeginEdit();

                    GoEdit(row, StockDays, Qty, "U");

                    row.EndEdit();
                }
                else
                {

                    row = dtStock.NewRow();
                    row["群組"] = Convert.ToString(dt.Rows[i]["群組"]);
                    string ITEM = Convert.ToString(dt.Rows[i]["群組2"]);
                    row["群組2"] = ITEM.Replace("-", "");
                    row["產品編號"] = Convert.ToString(dt.Rows[i]["產品編號"]);

                    GoEdit(row, StockDays, Qty, "A");


                    dtStock.Rows.Add(row);
                }



            }
         

            //如果是金額 -> 數量 * 項目成本

            if (Mode == "Qty")
            {
                return dtStock;
            }

            decimal StockCost = 0;

            for (int i = 0; i <= dtStock.Rows.Count - 1; i++)
            {
                ItemCode = Convert.ToString(dtStock.Rows[i]["產品編號"]);
                //取得 Cost
           
 
                StockCost = GetItemCost(ItemCode);

                row = dtStock.Rows[i];

                row.BeginEdit();
                row["數量"] = Convert.ToDecimal(row["0-30"]) +
                Convert.ToDecimal(row["31-60"]) +
                Convert.ToDecimal(row["61-90"]) +
                Convert.ToDecimal(row["91-120"]) +
                Convert.ToDecimal(row["121-180"]) +
                Convert.ToDecimal(row["181-360"]) +
                Convert.ToDecimal(row["360以上"]);
         
                row["0-30"] = Convert.ToDecimal(row["0-30"]) * StockCost;
                row["31-60"] = Convert.ToDecimal(row["31-60"]) * StockCost;
                row["61-90"] = Convert.ToDecimal(row["61-90"]) * StockCost;
                row["91-120"] = Convert.ToDecimal(row["91-120"]) * StockCost;
                row["121-180"] = Convert.ToDecimal(row["121-180"]) * StockCost;
                row["181-360"] = Convert.ToDecimal(row["181-360"]) * StockCost;
                row["360以上"] = Convert.ToDecimal(row["360以上"]) * StockCost;

                row["小計"] = Convert.ToDecimal(row["0-30"]) +
                                Convert.ToDecimal(row["31-60"]) +
                                Convert.ToDecimal(row["61-90"]) +
                                Convert.ToDecimal(row["91-120"]) +
                                Convert.ToDecimal(row["121-180"]) +
                                Convert.ToDecimal(row["181-360"]) +
                                Convert.ToDecimal(row["360以上"]);


                row.EndEdit();


            }



            return dtStock;

        }
        private decimal GetItemCost(string ItemCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT cast((cast(((select sum(b.TransValue)  from oinm b where b.itemcode = t0.[ItemCode]  and  Convert(varchar(8),B.docdate,112) <=@DOCDATE and InvntAct is not null and InvntAct <>'')/case (SUM(T0.[InQty])-SUM(T0.[OutQty])) when 0 then 1 else (SUM(T0.[InQty])-SUM(T0.[OutQty])) end) as decimal(23,10))) as varchar) StockCost");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  (T0.[docdate] >='2007.12.31' AND  Convert(varchar(8),T0.docdate,112) <=@DOCDATE) ");
            sb.Append(" and  ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'   AND T0.ITEMCODE=@ItemCode ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBoxDocDate.Text));
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

            if (!Convert.IsDBNull(dt.Rows[0][0]))
            {
                return Convert.ToDecimal(dt.Rows[0][0]);
            }
            else
            {
                return 0;
            }


        }
        private void GoEdit(DataRow row, Int32 StockDays, decimal Qty, string Mode)
        {

            if (Mode == "A")
            {
                row["0-30"] = 0;
                row["31-60"] = 0;
                row["61-90"] = 0;
                row["91-120"] = 0;
                row["121-180"] = 0;
                row["181-360"] = 0;
                row["360以上"] = 0;

            }

            if (StockDays >= 0 && StockDays <= 30)
            {
                if (Mode == "A")
                {
                    row["0-30"] = Qty;


                }
                else
                {
                    row["0-30"] = Qty + Convert.ToDecimal(row["0-30"]);
                }
            }
            else if (StockDays >= 31 && StockDays <= 60)
            {
                if (Mode == "A")
                {
                    row["31-60"] = Qty;


                }
                else
                {
                    row["31-60"] = Qty + Convert.ToDecimal(row["31-60"]);
                }
            }
            else if (StockDays >= 61 && StockDays <= 90)
            {
                if (Mode == "A")
                {
                    row["61-90"] = Qty;


                }
                else
                {
                    row["61-90"] = Qty + Convert.ToDecimal(row["61-90"]);
                }
            }
            else if (StockDays >= 91 && StockDays <= 120)
            {
                if (Mode == "A")
                {
                    row["91-120"] = Qty;


                }
                else
                {
                    row["91-120"] = Qty + Convert.ToDecimal(row["91-120"]);
                }
            }
            else if (StockDays >= 121 && StockDays <= 180)
            {
                if (Mode == "A")
                {
                    row["121-180"] = Qty;


                }
                else
                {
                    row["121-180"] = Qty + Convert.ToDecimal(row["121-180"]);
                }
            }
            else if (StockDays >= 181 && StockDays <= 360)
            {
                if (Mode == "A")
                {
                    row["181-360"] = Qty;


                }
                else
                {
                    row["181-360"] = Qty + Convert.ToDecimal(row["181-360"]);
                }
            }
            else
            {
                if (Mode == "A")
                {
                    row["360以上"] = Qty;


                }
                else
                {
                    row["360以上"] = Qty + Convert.ToDecimal(row["360以上"]);
                }

            }








            if (Mode == "A")
            {
                row["小計"] = Qty;
            }
            else
            {

                row["小計"] = Qty + Convert.ToDecimal(row["小計"]);
            }

        }
        private System.Data.DataTable MakeTable_Stock()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("群組2", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));

            dt.Columns.Add("數量", typeof(decimal));

            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360以上", typeof(decimal));

            dt.Columns.Add("小計", typeof(decimal));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["產品編號"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable MakeTable_Stock2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("群組2", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360以上", typeof(decimal));

            dt.Columns.Add("小計", typeof(decimal));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["產品編號"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable GetStockList(string DocDate)
        {

            System.Data.DataTable dt = GetItemHisList(DocDate);


            System.Data.DataTable dtStock = MakeTable();


            System.Data.DataTable dtTmp;

            DataRow row;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                string ItemCode = Convert.ToString(dt.Rows[i]["ItemCode"]);
               
           
                dtTmp = GetFIFO_Stock(ItemCode, DocDate);

                for (int j = 0; j <= dtTmp.Rows.Count - 1; j++)
                {


                    row = dtStock.NewRow();
                    row["產品編號"] = Convert.ToString(dtTmp.Rows[j]["ItemCode"]);
                    row["日期"] = Convert.ToString(dtTmp.Rows[j]["DocDate"]);
                    row["群組"] = Convert.ToString(dt.Rows[i]["群組"]);
                    row["群組2"] = Convert.ToString(dt.Rows[i]["群組2"]);
                 
                   row["庫存量"] = Convert.ToDecimal(dtTmp.Rows[j]["庫存量"]);
    

                    row["庫存天數"] = CountDays(StrToDate(DocDate), StrToDate(Convert.ToString(dtTmp.Rows[j]["DocDate"])), false);
                    dtStock.Rows.Add(row);

                }
            }

            return dtStock;

        }
        int CountDays(DateTime dateFrom, DateTime dateTo, bool including)
        {
            return ((System.TimeSpan)(dateTo - dateFrom)).Days * (-1) + (including ? 1 : 0);
        }
        private System.Data.DataTable GetItemHisList(string DocDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ItemCode], T1.[ItemName],SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue,MAX(substring(T2.itmsgrpNAM,4,15)) 群組,max(substring(t1.u_group,5,20)) 群組2 ");
            sb.Append(" FROM  [dbo].[OINM] T0  ");
            sb.Append(" INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T2  ON  T1.itmsgrpcod = T2.itmsgrpcod  left JOIN OITM T11 ON T0.ITEMCODE = T11.ITEMCODE ");

            sb.Append("WHERE  T0.[DocDate] <= @DocDate ");
            sb.Append(" and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  AND  T0.[ItemCode] NOT LIKE '%-C%'   ");

            sb.Append("  AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
            sb.Append("     and  T0.TransType <>67 ");
            sb.Append("GROUP BY  T0.[ItemCode], T1.[ItemName] ");
            sb.Append("Having SUM(T0.[InQty] - T0.[OutQty]) <> 0 ");

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
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }
        private System.Data.DataTable GetFIFO_Stock(string ItemCode1, string DocDate)
        {
          
            System.Data.DataTable dt = GetFIFO(ItemCode1, DocDate);


            DataRow row = null;

            decimal CalQty = 0;
            Int32 CalValue = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                row = dt.Rows[i];
                row.BeginEdit();

                CalQty = CalQty + Convert.ToDecimal(row["Qty"]);
                row["累計數量"] = CalQty;


                CalValue = CalValue + Convert.ToInt32(row["TransValue"]);
                row["累計值"] = CalValue;

                row.EndEdit();

            }

            //反推回去,剩下的庫存量,應該是那幾個日期

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {

                string InQty  = row["InQty"].ToString();
                row = dt.Rows[i];

                if (Convert.ToDecimal(row["InQty"]) == 0)
                {
                    continue;
                }

                row.BeginEdit();

                row["累計數量"] = 0;
                row["累計值"] = 0;


                if (CalQty - Convert.ToDecimal(row["InQty"]) <= 0)
                {
                    row["庫存量"] = CalQty;
                    row.EndEdit();



                    break;
                }
                else
                {
                    row["庫存量"] = Convert.ToDecimal(row["InQty"]);
                }

                CalQty = CalQty - Convert.ToDecimal(row["InQty"]);

                row.EndEdit();
            }


            // return dt;

            DataView dv = dt.DefaultView;
            string fd = row["Qty"].ToString();
           // dv.RowFilter = "庫存量 > 0";

            return dv.ToTable();

        }
        private System.Data.DataTable GetFIFO(string ItemCode1, string DocDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Convert(Varchar(8),T0.[DocDate],112) DocDate,T0.[ItemCode], T1.[ItemName],SUM(T0.[InQty]) InQty, SUM(T0.[OutQty]) OutQty,SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue, ");

            //計算欄位
            sb.Append(" 0.0 as 累計數量, 0 as 累計值, 0.0 庫存量 ");

            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("WHERE  T0.[DocDate] <= @DocDate ");
            sb.Append("AND  T0.[ItemCode] = @ItemCode1  ");
            sb.Append("AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
            sb.Append("     and  T0.TransType <>67");
            sb.Append("GROUP BY T0.[DocDate],T0.[ItemCode], T1.[ItemName] ");
            sb.Append("ORDER BY  T0.[DocDate]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode1", ItemCode1));
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

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位
            //
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("群組2", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("日期", typeof(string));
           
            dt.Columns.Add("庫存量", typeof(decimal));
            dt.Columns.Add("庫存天數", typeof(Int32));

            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["產品編號"];
            colPk[1] = dt.Columns["日期"];
            dt.PrimaryKey = colPk;


            return dt;
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }

            if (e.RowIndex >= dataGridView1.Rows.Count)
                return;

            try
            {
                if (!String.IsNullOrEmpty(dgr.Cells["Column9"].Value.ToString()))
                {
                   // string FA = dgr.Cells["Column9"].Value.ToString();
                    if (Convert.ToDecimal(dgr.Cells["Column9"].Value.ToString()) < 0)
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

        private void fmAcmeFIFO_Load(object sender, EventArgs e)
        {
            if (globals.GroupID.ToString().Trim() == "ACCS")
            {
                button8.Visible = false;
                button9.Visible = false;
                return;
            }

            textBoxDocDate.Text = GetMenu.Day();
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

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView2.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }


            try
            {
                if (!String.IsNullOrEmpty(dgr.Cells["AMTSUM"].Value.ToString()))
                {
                    // string FA = dgr.Cells["Column9"].Value.ToString();
                    if (Convert.ToDecimal(dgr.Cells["AMTSUM"].Value.ToString()) < 0)
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

    

    }
}
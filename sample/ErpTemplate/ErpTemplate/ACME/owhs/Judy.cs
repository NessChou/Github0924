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

    public partial class Judy : Form
    {
        string NewFileName;
        string FileName;
        string OutPutFile = "";
        public Judy()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.sALES_DOCCURTableAdapter1.Fill(this.wh.SALES_DOCCUR);
     

            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetBU("Judy"), "DataText", "DataValue");

            comboBox2.Text = "ASP";

            System.Data.DataTable D1 = GetTYPE(comboBox2.Text);
            dataGridView4.DataSource = D1;
            
        }

   

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("算式", typeof(string));
            dt.Columns.Add("結果", typeof(string));



            return dt;
        }






        private System.Data.DataTable Get1()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  select  T0.ITEMCODE ITEMCODE,Substring(T0.[ItemCode],2,8) Model,Substring(T1.[ItemCode],12,1) Version,");
            sb.Append(" CASE (Substring(T1.[ItemCode],11,1)) when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append("  when '1' then 'P' when '2' then 'N' when '3' then 'V' when '4' then 'U' when '5' then 'N' ELSE 'X'END Grade,");
            sb.Append(" CASE WHEN Substring (T0.[ItemCode],15,1)='9' THEN '@' ELSE ''  END Moving, ");
            sb.Append(" CAST(T2.ONHAND AS INT) Qty_Stock, ");
            sb.Append(" T4.PRICE  Cost_Stock,");
            sb.Append(" CONVERT(varchar(12) ,T5.DOCDUEDATE,111) Date_Stock,T6.COST,T6.AddDate,CAST(T2.AVGPRICE/(select top 1 RATE from ortt where currency='USD'  order by ratedate desc) AS DECIMAL(10,2)) Cost_Average,T7.Price,T7.AddDate Date_Price from oinm t0");
            sb.Append(" inner join (select max(transnum) num,itemcode from oinm t0 ");
            sb.Append(" where transtype IN ('20') and CASE INQTY WHEN 0 THEN OUTQTY ELSE INQTY END <> 0  group by itemcode ) t1 on (t0.transnum=t1.num)");
            sb.Append(" LEFT JOIN OITM T2 ON (T0.ITEMCODE=T2.ITEMCODE)");
            sb.Append(" LEFT JOIN PDN1 T3 ON (T0.BASE_REF=T3.docentry and T0.DOCLINENUM=T3.linenum)");
            sb.Append(" LEFT JOIN POR1 t4 on (t3.baseentry=T4.docentry and  t3.baseline=t4.linenum  )");
            sb.Append(" LEFT JOIN OPOR T5 on (T4.DOCENTRY=T5.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.AP_STOCK T6 ON (T0.ITEMCODE=T6.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.AP_STOCK2 T7 ON (T0.ITEMCODE=T7.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE substring(t0.itemcode,1,1)='T'");
            sb.Append(" and (CASE substring(T0.ITEMCODE,2,1)");
            sb.Append(" WHEN 'G' THEN 'G1' WHEN 'A' THEN 'G1'");
            sb.Append(" WHEN 'M' THEN 'G1' WHEN 'T' THEN 'G2'");
            sb.Append(" WHEN 'P' THEN 'G2' WHEN 'B' THEN 'G3' END) = @G");
            sb.Append(" ORDER BY Substring(T0.[ItemCode],2,8) ,Substring(T1.[ItemCode],12,1) ,");
            sb.Append(" CASE (Substring(T1.[ItemCode],11,1)) when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append("  when '1' then 'P' when '2' then 'N' when '3' then 'V' when '4' then 'U' when '5' then 'N' ELSE 'X'END ,");
            sb.Append(" CASE WHEN Substring (T0.[ItemCode],15,1)='9' THEN '@' ELSE ''  END ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@G", comboBox1.SelectedValue.ToString()));
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
        private System.Data.DataTable Get11(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT COST FROM AP_STOCK WHERE ITEMCODE=@ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        private System.Data.DataTable Get2(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Price FROM AP_STOCK2 WHERE ITEMCODE=@ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        private System.Data.DataTable GetTYPE(string TYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT  SALES,[MONTH],AMOUNT FROM SALES_REPORT WHERE TYPE=@TYPE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
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
        private void Delete(string itemcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  AP_Stock where [itemcode]=@itemcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@itemcode", itemcode));
       
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
        private void Delete2(string itemcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  AP_Stock2 where [itemcode]=@itemcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@itemcode", itemcode));

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
        private void Update1(string ItemCode, string Cost, string AddDate)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO AP_Stock(ItemCode,Cost,AddDate) VALUES (@ItemCode,@Cost,@AddDate)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Cost", Cost));
            command.Parameters.Add(new SqlParameter("@AddDate", AddDate));
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

        private void Update2(string ItemCode, string Price, string AddDate)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO AP_Stock2(ItemCode,Price,AddDate) VALUES (@ItemCode,@Price,@AddDate)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Price", Price));
            command.Parameters.Add(new SqlParameter("@AddDate", AddDate));
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
            ExcelReport.GridViewToExcel(dataGridView3);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView3.Rows.Count - 1; i++)
            {

                DataGridViewRow row;

                row = dataGridView3.Rows[i];
                string a0 = row.Cells["ITEMCODE"].Value.ToString();
                string a1 = row.Cells["COST"].Value.ToString();
                string AddDate = row.Cells["AddDate"].Value.ToString();
                System.Data.DataTable bb = Get11(a0);
                if (bb.Rows.Count > 0)
                {
                    Delete(a0);
                    Update1(a0, a1, AddDate);

                }
                else if (bb.Rows.Count.ToString() == "0" && a1 != "")
                {
                    Delete(a0);
                    Update1(a0, a1, AddDate);
                }
         



            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Get1();
        }

        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView3.Rows.Count - 1; i++)
            {

                DataGridViewRow row;

                row = dataGridView3.Rows[i];
                string a0 = row.Cells["ITEMCODE"].Value.ToString();
                string a1 = row.Cells["Price"].Value.ToString();
                string Date_Stock = row.Cells["Date_Stock"].Value.ToString();
                
                System.Data.DataTable bb = Get2(a0);
                if (bb.Rows.Count > 0)
                {
                    Delete2(a0);
                    Update2(a0, a1, Date_Stock);

                }
                else if (bb.Rows.Count.ToString() == "0" && a1 != "")
                {
                    Delete2(a0);
                    Update2(a0, a1, Date_Stock);
                }




            }
        }



        private void button6_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.sALES_DOCCURBindingSource.EndEdit();
            this.sALES_DOCCURTableAdapter1.Update(this.wh.SALES_DOCCUR);

            MessageBox.Show("更新成功");
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridView3.Columns[e.ColumnIndex].Name == "COST" )
            {
                this.dataGridView3.Rows[e.RowIndex].Cells["AddDate"].Value = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");

            }
            if (dataGridView3.Columns[e.ColumnIndex].Name == "Cost_Stock")
            {
                this.dataGridView3.Rows[e.RowIndex].Cells["Date_Stock"].Value = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");

            }
        }

        private void button7_Click(object sender, EventArgs e)
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
                    TruncateTable();
                    GetExcelContentGD44(opdf.FileName);


                    System.Data.DataTable D1 = GetTYPE(comboBox2.Text);
                    dataGridView4.DataSource = D1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
             int i = excelBook.Sheets.Count;

             for (int xi = 1; xi <= i; xi++)
             {
                 Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                 string d = excelSheet.Name.Trim().ToString();
                 int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                 int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                 Hashtable ht = new Hashtable(iRowCnt);



                 Microsoft.Office.Interop.Excel.Range range = null;



                 object SelectCell = "A1";
                 range = excelSheet.get_Range(SelectCell, SelectCell);


                 string id;
                 string id2 = "";

                 for (int b = 1; b <= iRowCnt; b++)
                 {

                     range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 1]);
                     id = range.Text.ToString();
                     for (int jj = 2; jj <= 13; jj++)
                     {

                         range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, jj]);
                         id2 = range.Text.ToString();
                         id2 = id2.Replace(",", "");
                         if (!String.IsNullOrEmpty(id))
                         {
                             AddAUOGD4(d, id, (jj - 1).ToString(), id2.Trim());
                         }

                     }

                 }


                 System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                 System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                 range = null;
                 excelSheet = null;
             }

            //Quit
            excelApp.Quit();
    

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


           
            excelApp = null;
            excelBook = null;
          

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            MessageBox.Show("匯入成功");
        }

        public void AddAUOGD4(string TYPE,string SALES,string MONTH,string AMOUNT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SALES_REPORT(TYPE,SALES,MONTH,AMOUNT) values(@TYPE,@SALES,@MONTH,@AMOUNT)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
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

      
        private void TruncateTable()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("DELETE SALES_REPORT ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


           // command.Parameters.Add(new SqlParameter("@TYPE", TYPE));

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
        private void UPDATE(string TYPE)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("DELETE SALES_REPORT WHERE TYPE=@TYPE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));

            SqlDataAdapter da = new SqlDataAdapter(command);


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    MessageBox.Show("刪除成功");
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



        private void button11_Click(object sender, EventArgs e)
        {
            execuse();


        }

        private System.Data.DataTable Gen_201004()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append(" DROP TABLE ACMESQLSP.DBO.SALESREPORT  ");

            sb.Append("      SELECT ");
            sb.Append("               SUM(總數量) 總數量,");
            sb.Append("               SUM(總金額) 總金額,");
            sb.Append("               SUM(總成本) 總成本,");
            sb.Append("               SUM(總毛利) 總毛利,");
            sb.Append("               SUM(佣金) 佣金,");
            sb.Append("               姓名 業務,MONTH(日期) 日期 INTO ACMESQLSP.DBO.SALESREPORT");
            sb.Append("               FROM ( ");
            sb.Append("                           SELECT 'AR' as 單別,");
            sb.Append("                          ltrim(substring(T3.SLPNAME,CHARINDEX('(',T3.SLPNAME)+1,LEN(T3.SLPNAME)-CHARINDEX('(',T3.SLPNAME)-1)) 姓名,");
            sb.Append("                           sum(T2.LineTotal) 總金額,");
            sb.Append("                           sum(T2.Quantity) 總數量,");
            sb.Append("                           sum(Round(T2.StockPrice*T2.Quantity,0)) 總成本,");
            sb.Append("                           sum(T2.LineTotal) - sum(Round(T2.StockPrice*T2.Quantity,0))  總毛利,");
            sb.Append("                           T0.DocDate 日期,sum(ISNULL(t2.u_commission,0)*(CASE T7.DOCRATE WHEN 0 THEN 1 ELSE T7.DOCRATE END)*T2.QUANTITY) 佣金");
            sb.Append("                           FROM OINV T0 ");
            sb.Append("                           INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T2.ItemCode ");
            sb.Append("                           INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode ");
            sb.Append("                           INNER JOIN OCRD T4 ON T0.CardCode = T4.CardCode ");
            sb.Append("                           left join dln1 t5 on (t2.baseentry=T5.docentry and  t2.baseline=t5.linenum  and t2.basetype='15')");
            sb.Append("                           left join rdr1 t6 on (t5.baseentry=T6.docentry and  t5.baseline=t6.linenum  and t6.targettype='15')");
            sb.Append("                           left join ORDR t7 ON(T6.DOCENTRY=T7.DOCENTRY) ");
            sb.Append("                           WHERE T0.[DocType] ='I' ");
            sb.Append("                           and ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            sb.Append("                           and YEAR(T0.[DocDate])=YEAR(GETDATE()) ");
            sb.Append("                           and T4.CardType='C' AND ltrim(substring(T3.SLPNAME,CHARINDEX('(',T3.SLPNAME)+1,LEN(T3.SLPNAME)-CHARINDEX('(',T3.SLPNAME)-1))  IN (SELECT DISTINCT LTRIM(RTRIM(SALES))  COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.SALES_REPORT WHERE SALES <> 'STONE')");
            sb.Append("                           GROUP BY T0.DocDate,T0.SlpCode , T3.SlpName,T0.CardCode,T4.CardName");
            sb.Append("                           union all");
            sb.Append("                           SELECT '貸項' as 單別,");
            sb.Append("                           ltrim(substring(T3.SLPNAME,CHARINDEX('(',T3.SLPNAME)+1,LEN(T3.SLPNAME)-CHARINDEX('(',T3.SLPNAME)-1)) 姓名,");
            sb.Append("                           sum(T2.LineTotal) * (-1) 總金額,");
            sb.Append("                           sum(T2.Quantity)  * (-1)  總數量,");
            sb.Append("                           sum(Round(T2.StockPrice*T2.Quantity,0)) * (-1) 總成本,");
            sb.Append("                           (sum(T2.LineTotal) - sum(Round(T2.StockPrice*T2.Quantity,0))) * (-1)  總毛利,");
            sb.Append("                           T0.DocDate 日期,0 佣金");
            sb.Append("                           FROM ORIN T0 ");
            sb.Append("                           INNER JOIN RIN1 T2 ON T0.DocEntry = T2.DocEntry  INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T2.ItemCode ");
            sb.Append("                           INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode ");
            sb.Append("                           INNER JOIN OCRD T4 ON T0.CardCode = T4.CardCode ");
            sb.Append("                           WHERE T0.[DocType] ='I' ");
            sb.Append("                           and ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
            sb.Append("                           and YEAR(T0.[DocDate])=YEAR(GETDATE()) ");
            sb.Append("                           and T4.CardType='C' AND ltrim(substring(T3.SLPNAME,CHARINDEX('(',T3.SLPNAME)+1,LEN(T3.SLPNAME)-CHARINDEX('(',T3.SLPNAME)-1))  IN (SELECT DISTINCT LTRIM(RTRIM(SALES))  COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.SALES_REPORT WHERE SALES <> 'STONE')");
            sb.Append("                           GROUP BY T0.DocDate,T0.SlpCode , T3.SlpName,T0.CardCode,T4.CardName");
            sb.Append("               ) T ");
            sb.Append(" GROUP BY MONTH(日期),姓名");
      
            sb.Append(" SELECT * FROM (");
            sb.Append(" SELECT 業務,'Qty' ITEM,T1.AMOUNT TARGET,CAST(cast(總數量 as int) AS VARCHAR) ACT,T0.日期 FROM ACMESQLSP.DBO.SALESREPORT T0");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SALES_REPORT T1 ON (T0.業務=T1.SALES COLLATE  Chinese_Taiwan_Stroke_CI_AS AND T0.日期=T1.[MONTH] COLLATE  Chinese_Taiwan_Stroke_CI_AS and t1.[TYPE]='QTY')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 業務,'ASP' Item,T1.AMOUNT Target,CAST(cast(round(總金額/總數量,0) as int) AS VARCHAR) Actually,T0.日期 FROM ACMESQLSP.DBO.SALESREPORT T0");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SALES_REPORT T1 ON (T0.業務=T1.SALES COLLATE  Chinese_Taiwan_Stroke_CI_AS AND T0.日期=T1.[MONTH] COLLATE  Chinese_Taiwan_Stroke_CI_AS and t1.[TYPE]='ASP')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 業務,'REV' Item,T1.AMOUNT Target,CAST(cast(總金額 as int) AS VARCHAR) Actually,T0.日期 FROM ACMESQLSP.DBO.SALESREPORT T0");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SALES_REPORT T1 ON (T0.業務=T1.SALES COLLATE  Chinese_Taiwan_Stroke_CI_AS AND T0.日期=T1.[MONTH] COLLATE  Chinese_Taiwan_Stroke_CI_AS and t1.[TYPE]='REV')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 業務,'GP%' Item,T1.AMOUNT Target,CAST(CAST(ROUND(((總金額-總成本-佣金)/總金額)*100,2) AS DECIMAL(10,2)) AS VARCHAR)+'%' Actually,T0.日期 FROM ACMESQLSP.DBO.SALESREPORT T0");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SALES_REPORT T1 ON (T0.業務=T1.SALES COLLATE  Chinese_Taiwan_Stroke_CI_AS AND T0.日期=T1.[MONTH] COLLATE  Chinese_Taiwan_Stroke_CI_AS and t1.[TYPE]='GP%')");
            sb.Append(" ) AS A where 日期 <> month(getdate()) ");
            sb.Append(" ORDER BY 業務,日期");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }




            return ds.Tables[0];

        }

        private void execuse()
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\SALESGP.xls";


                System.Data.DataTable OrderData = Gen_201004();


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N","1");


            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message);
            }
        }
        public static void ExcelReportOutput(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag,string TT)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Activate();
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = 5;

            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(OrderData, sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            break;
                        }

                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(OrderData, aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }


            }
            finally
            {

                //  string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
                //HttpContext.Current.User.Identity.Name.ToString() + Util.GetDate14() +
                //Path.GetFileName(ExcelFile);

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                string Msg = string.Empty;
                string Mo;

                System.Diagnostics.Process.Start(OutPutFile);

            }

        }

        //設定明細檔資料
        public static void SetRow(System.Data.DataTable OrderData, int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[iRow][FieldName]);
            }

        }
        public static bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }
            //}
            return false;
        }
        public static bool CheckSerial(System.Data.DataTable OrderData, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {

            System.Data.DataTable D1 = GetTYPE(comboBox2.Text);
            dataGridView4.DataSource = D1;
        }

     


    

   
    }
}
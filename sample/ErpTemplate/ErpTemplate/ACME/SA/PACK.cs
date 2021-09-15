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
    public partial class PACK : Form
    {
        private System.Data.DataTable TempDt;
        private System.Data.DataTable TempDt2;
        System.Data.DataTable dtCost2 = null;
        System.Data.DataTable dtCostF = null;
        private string FileName;
        string FA = "acmesql98";
        private int[] iCount;
        string GS;
        string YM = "";
        public PACK()
        {
            InitializeComponent();
        }

      
        private void GetExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
           
            
            excelApp.Visible = checkBox1.Checked;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
         
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
         
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
              int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
              int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

              progressBar1.Maximum = iRowCnt;

            
            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


          

            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {
                    
                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord,1 ]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();
                    //SERIAL_NO = range.Text.ToString().Trim().Replace("-", "");
                    //SERIAL_NO = SERIAL_NO.Substring(0, 7);
                    range.Select();


                    //如果找不到時才新增
                     drFind =TempDt.Rows.Find(SERIAL_NO);

                     if (drFind == null)
                     {

                         dr = TempDt.NewRow();

                         dr["Mo"] = SERIAL_NO;

                         TempDt.Rows.Add(dr);
                     }

                }

        

            }
            finally
            {

           
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
            
            }
   

            dataGridView1.DataSource = TempDt;
        }

        //動態產生資料結構
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
        
            //第一個固定欄位(工單號碼)
            dt.Columns.Add("Mo", typeof(string));

            //最後一個總計
          //  dt.Columns.Add("Qty", typeof(int));

            
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Mo"];
            dt.PrimaryKey = colPk;
            
            return dt;
        }
        //動態產生資料結構
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("KEY", typeof(string));
    


            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["KEY"];

            dt.PrimaryKey = colPk;

            return dt;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
              TempDt= MakeTable();

              CreateboundColumn("Mo", "工單號碼");
              CreateboundColumn("Qty", "數量");
              dataGridView1.DataSource = TempDt;

              if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
              {
                  button3.Visible = false;

              }

        }

        private void CreateboundColumn(string FieldName,string Caption)
        {
            // Initialize the button column.
            DataGridViewTextBoxColumn TextBoxColumn = new DataGridViewTextBoxColumn();
            TextBoxColumn.Name = FieldName;
            TextBoxColumn.HeaderText = Caption;
            TextBoxColumn.DataPropertyName = FieldName;

            // Add the button column to the control.
            dataGridView1.Columns.Add(TextBoxColumn);
        }

        private void CreateboundColumn2(string FieldName, string Caption)
        {
            // Initialize the button column.
            DataGridViewTextBoxColumn TextBoxColumn = new DataGridViewTextBoxColumn();
            TextBoxColumn.Name = FieldName;
            TextBoxColumn.HeaderText = Caption;
            TextBoxColumn.DataPropertyName = FieldName;

            // Add the button column to the control.
            dataGridView2.Columns.Add(TextBoxColumn);
        }
        private void button1_Click(object sender, EventArgs e)
        {
           
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
    

                iCount = new int[TempDt.Rows.Count];
                GetExcelProduct(FileName);

               
            }
        }


        private bool CheckSerial(string Serial_no)
        {
            string Mo = string.Empty;

            int iChecked=-1;

            for (int i = 0; i <= TempDt.Rows.Count - 1; i++)
            {
                Mo =Convert.ToString(TempDt.Rows[i]["Mo"]);
                Mo = Mo.Substring(3, 7);

                if (Mo == Serial_no)
                {
                    iChecked = i;

                    iCount[i] = iCount[i] + 1;

                    if (iCount[i] > Convert.ToInt32(TempDt.Rows[i]["Qty"]))
                    {
                        return false;
                    }

                    return true;
                }
            
            }
            return false;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                iCount = new int[TempDt.Rows.Count];
                WriteExcelProduct(FileName);


            }
        }
        private void WriteExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = checkBox1.Checked;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            MessageBox.Show(iRowCnt.ToString());

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;

                int Qty = 0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();


                    range.Select();


                    drFind = TempDt.Rows.Find(SERIAL_NO);

                    if (drFind != null)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, Convert.ToInt32(textBox1.Text) + 1]);
                        range.Value2 = "Z";

                        //System.Data.DataTable H1 = GetITEM2(SERIAL_NO);
                        //if (H1.Rows.Count > 0)
                        //{
                        //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, Convert.ToInt32(textBox1.Text) + 2]);
                        //    range.Value2 = H1.Rows[0][0].ToString();
                        //}
                    }

                }


            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
               "Acme_" + Path.GetFileName(FileName);

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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                //MessageBox.Show("產生一個檔案->"+NewFileName);


            }


            dataGridView1.DataSource = TempDt;
        }
        private void WriteExcelDD(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = checkBox1.Checked;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            MessageBox.Show(iRowCnt.ToString());

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;

                int Qty = 0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();


                    range.Select();

                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {
                        drFind = TempDt.Rows.Find(SERIAL_NO);

                        if (drFind != null)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                            range.Value2 = "Z";


                        }
                    }

                }


            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
               "Acme_" + Path.GetFileName(FileName);

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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                //MessageBox.Show("產生一個檔案->"+NewFileName);


            }


            dataGridView1.DataSource = TempDt;
        }

        private void WriteDD(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false ;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            MessageBox.Show(iRowCnt.ToString());

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string CARDNAME;

                int Qty=0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 23]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    
                    range.Select();

                    System.Data.DataTable GG1 = GETCARDNAME(CARDNAME);
                      if(GG1.Rows.Count >0)
                      {
        
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 28]);
                        range.Value2 = GG1.Rows[0][0].ToString();
                      }

                }


            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
               "Acme_" + Path.GetFileName(FileName);

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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                //MessageBox.Show("產生一個檔案->"+NewFileName);

             
            }
    

            dataGridView1.DataSource = TempDt;
        }

        private void WriteExcelProduct2(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true ;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            progressBar2.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string MODEL="";
                string VER="";
                string GRADE="";
                string Qty="";

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar2.Value = iRecord;
                    progressBar2.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                   string  SERIAL_NO = range.Text.ToString().Trim();
                   if (!String.IsNullOrEmpty(SERIAL_NO))
                   {
                       int D = SERIAL_NO.LastIndexOf(".");
                       if (D != -1)
                       {
                           VER = SERIAL_NO.Substring(D+1, 3);


                           range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                           range.Select();
                           string SERIAL_NO2 = range.Text.ToString().Trim();
             

                           if (!String.IsNullOrEmpty(SERIAL_NO2))
                           {
                               MODEL = SERIAL_NO2.Substring(0, 8);

                               int G = SERIAL_NO2.LastIndexOf("-");
                               if (G != -1)
                               {
                                   GRADE = SERIAL_NO2.Substring(G + 1, 1);
                               }
                           }


                           range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                           range.Select();
                           string hj = range.Text.ToString().Trim();
                           Qty = hj.Replace(",", "");

                           string GG = MODEL + GRADE + VER + Qty;
                           drFind = TempDt2.Rows.Find(GG);

                           if (drFind == null)
                           {
                        
                               string ff =SERIAL_NO.Substring(0,2);
                               if (ff == "97")
                               {
                                   object SelectCell_From = "A" + iRecord;
                                   object SelectCell_To = "E" + iRecord;
                                   range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                                   range.Select();

                                   range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                               }
                           }
                       
                       }
                      
                     
                   }
               



                }


            }
            finally
            {
                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
               "Acme_" + Path.GetFileName(FileName);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                MessageBox.Show("產生一個檔案->"+NewFileName);


            }



        }

        private void button4_Click(object sender, EventArgs e)
        {
            TempDt2 = MakeTable2();
            DataRow dr = null;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                System.Data.DataTable dt = GetITEM();
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = TempDt2.NewRow();

                    string MODEL = dt.Rows[i]["MODEL"].ToString().Trim();
                    string GRADE = dt.Rows[i]["GRADE"].ToString().Trim();
                    string VER = dt.Rows[i]["VER"].ToString().Trim();
                    string QTY = dt.Rows[i]["QTY"].ToString().Trim();
                    dr["KEY"] = MODEL + GRADE + VER + QTY;

                    TempDt2.Rows.Add(dr);
                }


                FileName = openFileDialog1.FileName;
       
                iCount = new int[TempDt.Rows.Count];
                WriteExcelProduct2(FileName);


            }
        }

        private System.Data.DataTable GetITEM()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  Substring (T0.[ItemCode],2,8) Model,");
            sb.Append(" CASE (Substring(T0.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append("  when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append("  when '4' then 'U' when '5' then 'N' ELSE 'X'");
            sb.Append("  END GRADE,Substring(T1.[ItemCode],12,3) VER,SUM(cast(T1.[OnHand] as int)) QTY FROM OITM T0 ");
            sb.Append("  INNER JOIN OITW T1 ON T0.ItemCode = T1.ItemCode WHERE T1.[WhsCode] ='TW002' AND  T1.[OnHand] <>0 ");
            sb.Append(" GROUP BY");
            sb.Append(" Substring (T0.[ItemCode],2,8) ,");
            sb.Append(" CASE (Substring(T0.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append("  when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append("  when '4' then 'U' when '5' then 'N' ELSE 'X'");
            sb.Append("  END ,Substring(T1.[ItemCode],12,3) ");
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
        private System.Data.DataTable GetITEM2(string SHIPNO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT COUNT(*) QTY FROM  AA WHERE SHIPNO=@SHIPNO ");
      
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPNO", SHIPNO));
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
        private System.Data.DataTable GetSERF2RMA(string INVOICE, string PARTNO)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE ");
            sb.Append(" AND T2.U_PARTNO =@PARTNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable MakeMain2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("片序", typeof(string));
            dt.Columns.Add("箱序", typeof(string));
            dt.Columns.Add("棧板號", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("收採單號", typeof(string));
            dt.Columns.Add("供應商名稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("INVOICE", typeof(string));
            dt.Columns.Add("INVOICE日期", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("銷貨單號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));

            return dt;
        }
        private void button3_Click(object sender, EventArgs e)
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
                    dtCost2 = MakeMain2();
                    GD4(opdf.FileName);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GD4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string SHIPNO;


            for (int i = 2; i <= iRowCnt; i++)
            {

                string SO;
                string CO;
                string PO;
                string CARTON = "";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                SO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                CO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                PO = range.Text.ToString().Trim();



                try
                {


                    System.Data.DataTable dt = GetScheSap2(SO, CO, PO);
                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = null;




                        dr = dtCost2.NewRow();
                        string INVOICE = dt.Rows[0]["INVOICE"].ToString();
                        string ITEMCODE = dt.Rows[0]["產品編號"].ToString();


                        if (String.IsNullOrEmpty(CO))
                        {
                            if (!String.IsNullOrEmpty(PO))
                            {
                                System.Data.DataTable GETCART = GetCARTON3(PO);
                                if (GETCART.Rows.Count > 0)
                                {

                                    CARTON = GETCART.Rows[0][0].ToString();
                                }
                            }
                            if (!String.IsNullOrEmpty(SO))
                            {
                                System.Data.DataTable GETCART = GetCARTON(SO);
                                System.Data.DataTable GETCART2 = GetCARTON2(SO);
                                if (GETCART.Rows.Count > 0)
                                {

                                    CARTON = GETCART.Rows[0][0].ToString();
                                }
                                else if (GETCART2.Rows.Count > 0)
                                {

                                    CARTON = GETCART2.Rows[0][0].ToString();
                                }

                            }


                        }
                        else
                        {
                            CARTON = CO;
                        }


                        System.Data.DataTable TODLN = GetODLN(INVOICE, ITEMCODE, CARTON);
                        if (TODLN.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                            range.Select();
                            range.Value2 = TODLN.Rows[0]["出貨日期"].ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                            range.Select();
                            range.Value2 = TODLN.Rows[0]["銷貨單號"].ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                            range.Select();
                            range.Value2 = TODLN.Rows[0]["客戶名稱"].ToString();

                        }


                    }


                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
            MessageBox.Show("匯出成功");
            dataGridView3.DataSource = dtCost2;
        }
        private void GD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string SHIPNO;
  

            for (int i = 2; i <= iRowCnt; i++)
            {

                string USER;
                string P1;
                string P2;
                string P3;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                USER = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                P1 = range.Text.ToString().Trim().ToUpper();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                P2 = range.Text.ToString().Trim().ToUpper();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                P3 = range.Text.ToString().Trim().ToUpper();


                try
                {
                    if (!String.IsNullOrEmpty(USER))
                    {

                        if (P1 == "V")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                            range.Select();
                            range.Value2 = USER;
                        }

                        if (P2 == "V")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                            range.Select();
                            range.Value2 = USER;
                        }

                        if (P3 == "V")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                            range.Select();
                            range.Value2 = USER;
                        }
                        
                    }


                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
            MessageBox.Show("匯出成功");
            dataGridView3.DataSource = dtCost2;
        }
        private System.Data.DataTable GetODLN(string INVOICE_NO, string ITEMCODE, string CARTON)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT DISTINCT T9.CARDNAME 客戶名稱,Convert(varchar(10),(t9.docdate),112) 出貨日期,T4.DOCENTRY 銷貨單號,CAST(T4.Quantity AS INT) 數量   FROM  ACMESQLSP.DBO.AP_INVOICEIN T1");
            sb.Append("           LEFT JOIN ACMESQLSP.DBO.WH_Item4 T2 ON (T1.WHNO =T2.ShippingCode AND T1.ITEMCODE =T2.ItemCode) ");
            sb.Append("           left join ACMESQL02.DBO.dln1 t4 on (t4.baseentry=T2.DOCENTRY and  t4.baseline=t2.linenum  and t4.basetype='17') ");
            sb.Append("           left join ACMESQL02.DBO.odln t9 on (t4.docentry=T9.docentry ) ");
            sb.Append("           WHERE T1.INV=@INVOICE_NO   AND T1.ITEMCODE=@ITEMCODE ");
            if (!String.IsNullOrEmpty(CARTON))
            {
                sb.Append("         AND T1.CARTON =@CARTON ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE_NO", INVOICE_NO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetScheSap2(string SO, string CO, string PO)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  Convert(varchar(10),(t0.docdate),112) 過帳日期,T5.DOCENTRY 採購單號,t0.docentry 收採單號 ,t0.cardname 供應商名稱");
            sb.Append(" ,t1.ItemCode 產品編號,t1.dscription 品名,cast(t1.quantity as int) 數量, ");
            sb.Append(" LTRIM(RTRIM(T0.U_ACME_INV)) INVOICE,Convert(varchar(10),T0.U_ACME_Invoice ,111) INVOICE日期,T0.Comments 備註  from opdn t0 ");
            sb.Append(" LEFT JOIN PDN1 T1 ON(T0.docentry=t1.docentry) ");
            sb.Append(" LEFT JOIN POR1 T5 ON (T5.docentry=T1.baseentry AND T5.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN OITM T10 ON T1.ITEMCODE = T10.ITEMCODE    ");
            sb.Append(" WHERE  ISNULL(T10.U_GROUP,'') <> 'Z&R-費用類群組' ");



            string STYPE = "";
            string SER = "";
            if (SO != "")
            {
                SER = SO;
                STYPE = "A";
            }
            else if (CO != "")
            {
                SER = CO;
                STYPE = "C";
            }
            else if (PO != "")
            {
                SER = PO;
                STYPE = "B";
            }
            System.Data.DataTable DTSFRMA = GetSERFRMA(SER.Trim(), STYPE);
            System.Data.DataTable DTSF = GetSERF(SER.Trim(), STYPE);
            if (DTSF.Rows.Count > 0)
            {
                string INVOICE = DTSF.Rows[0]["INVOICE"].ToString();
                string MODEL = DTSF.Rows[0]["MODEL"].ToString();
                string GRADE = DTSF.Rows[0]["GRADE"].ToString();
                string PARTNO = DTSF.Rows[0]["PARTNO"].ToString();
                string ITEM = "";
                if (String.IsNullOrEmpty(MODEL))
                {
                    System.Data.DataTable DTSF2 = GetSERF2(INVOICE, PARTNO, GRADE);
                    if (DTSF2.Rows.Count > 0)
                    {
                        ITEM = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND  t0.u_acme_inv = '" + INVOICE + "' AND T1.ITEMCODE = '" + ITEM + "' ");
                    }

                }
                else
                {

                    System.Data.DataTable DTS = GetSER(SER.Trim(), STYPE);
                    ITEM = DTS.Rows[0][1].ToString();
                    System.Data.DataTable DTSF2 = GetSERF3(INVOICE, ITEM);
                    if (DTSF2.Rows.Count > 0)
                    {
                        string ITEMCODE = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND  t0.u_acme_inv = '" + INVOICE + "' AND T1.ITEMCODE = '" + ITEMCODE + "' ");
                    }
                }

            }
            else if (DTSFRMA.Rows.Count > 0)
            {

                string INVOICE = DTSFRMA.Rows[0]["INVOICE"].ToString();
                string MODEL = DTSFRMA.Rows[0]["MODEL"].ToString();
                string PARTNO = DTSFRMA.Rows[0]["PARTNO"].ToString();
                string MODEL1 = DTSFRMA.Rows[0]["MODEL1"].ToString();
                string VER = DTSFRMA.Rows[0]["VER"].ToString();
                string CARTON = DTSFRMA.Rows[0]["CARTON"].ToString();
                string ITEM = "";
                if (String.IsNullOrEmpty(MODEL))
                {
                    System.Data.DataTable DTSF2 = GetSERF2RMA(INVOICE, PARTNO);
                    if (DTSF2.Rows.Count > 0)
                    {
                        ITEM = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND  t0.u_acme_inv = '" + INVOICE + "' AND T1.ITEMCODE = '" + ITEM + "' ");

                       
                    }

                }
                else
                {

                    System.Data.DataTable DTSF2 = GetSERF2RMA2(INVOICE, MODEL1, VER);
                    if (DTSF2.Rows.Count > 0)
                    {
                        string ITEMCODE = "";
                        if (DTSF2.Rows.Count > 1)
                        {
                            System.Data.DataTable DTSF3 = GetSERF2RMA3(INVOICE, CARTON);
                            ITEMCODE = DTSF3.Rows[0][0].ToString();
                        }
                        else
                        {
                            ITEMCODE = DTSF2.Rows[0][0].ToString();
                        }
                        sb.Append("  AND  t0.u_acme_inv = '" + INVOICE + "' AND T1.ITEMCODE = '" + ITEMCODE + "' ");
                    }
                }

            }
            else
            {
                sb.Append("and 1=0  ");
            }




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ODLN");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ODLN"];
        }
        private System.Data.DataTable GetSERF(string SHIPPING_NO, string TYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT INVOICE_NO INVOICE,MODEL_NO MODEL,FINAL_GRADE GRADE,PART_NO PARTNO  FROM WH_AUO T0");

            if (TYPE == "A")
            {
                sb.Append(" WHERE T0.SHIPPING_NO = @SHIPPING_NO ");
            }

            if (TYPE == "B")
            {
                sb.Append(" WHERE T0.PALLET_NO = @SHIPPING_NO ");
            }

            if (TYPE == "C")
            {
                sb.Append(" WHERE T0.CARTON_NO = @SHIPPING_NO ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING_NO", SHIPPING_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF2(string INVOICE, string PARTNO, string GRADE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE ");
            sb.Append(" AND T2.U_PARTNO =@PARTNO AND T2.U_GRADE =@GRADE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF2RMA3(string INVOICE, string CARTON)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE FROM AP_INVOICEIN WHERE INV=@INVOICE AND CARTON=@CARTON");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GETCARDNAME(string CARDNAME)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT COUNT(*) 比數 FROM ODLN WHERE CARDNAME=@CARDNAME");
            //sb.Append(" SELECT t0.CardName  FROM OCRD T0");
            //sb.Append("  LEFT JOIN OSLP T2 ON (T0.SlpCode = T2.SlpCode)");
            //sb.Append("  LEFT JOIN OCRG T3 ON (T0.GROUPCODE = T3.GROUPCODE)");
            //sb.Append(" LEFT JOIN (SELECT CARDCODE,MAX(OWNERCODE) OWNERCODE FROM ORDR  where OWNERCODE in (select empid from ohem where isnull(termdate,'') =  '' )");
            //sb.Append("   GROUP BY CARDCODE) T4 ON (T0.CARDCODE=T4.CARDCODE)");
            //sb.Append(" LEFT JOIN OHEM T1 ON (T4.OWNERCODE=T1.EMPID)");
            //sb.Append(" where cardtype='c'  and t0.CARDCODE in (SELECT distinct CARDCODE FROM OINV WHERE YEAR(DOCDATE) between '2016' and '2019')");
            //sb.Append(" and  SUBSTRING(T3.GROUPNAME,4,15)='TFT' AND t0.CardName=@CARDNAME");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF3(string INVOICE, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE  ");
            sb.Append(" AND ACMESQLSP.dbo.fn_RemoveLastChar(T1.ITEMCODE)  COLLATE Chinese_Taiwan_Stroke_CI_AS =@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERFRMA(string SHIPPING, string TYPE)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT INVOICE,CASE SUBSTRING(PART,1,2) WHEN '97' THEN PART END PARTNO,");
            sb.Append("  CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN PART END MODEL,");
            sb.Append("  CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN SUBSTRING(PART,0,CHARINDEX('.', PART)) END MODEL1,");
            sb.Append("   CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN SUBSTRING(PART,CHARINDEX('.', PART)+1,3) END VER,CARTON");
            sb.Append("   FROM RMA_INVOICEOUT");
            sb.Append("   WHERE ISNULL(PART,'') <>'' ");

            if (TYPE == "A")
            {
                sb.Append(" AND SHIPPING = @SHIPPING ");
            }


            if (TYPE == "C")
            {
                sb.Append(" AND CARTON = @SHIPPING ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSER(string SHIPPING_NO, string TYPE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT T0.INVOICE_NO INVOICE,ACMESQLSP.dbo.fn_RemoveLastChar(T1.ITEMCODE) ITEMCODE FROM ACMESQLSP.DBO.WH_AUO  T0   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (CASE SUBSTRING(PART_NO,1,2) WHEN '91' THEN 'O' ELSE SUBSTRING(T0.MODEL_NO,1,1) END+  SUBSTRING(T0.MODEL_NO,2,CHARINDEX('.', T0.MODEL_NO)-2)=T1.U_TMODEL  COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.FINAL_GRADE=T1.U_GRADE  COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.PART_NO =T1.U_PARTNO COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            if (TYPE == "A")
            {
                sb.Append(" WHERE T0.SHIPPING_NO = @SHIPPING_NO ");
            }

            if (TYPE == "B")
            {
                sb.Append(" WHERE T0.PALLET_NO = @SHIPPING_NO ");
            }

            if (TYPE == "C")
            {
                sb.Append(" WHERE T0.CARTON_NO = @SHIPPING_NO ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING_NO", SHIPPING_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON_NO CARTON FROM WH_AUO WHERE SHIPPING_NO=@SHIPPING ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON3(string PALLET_NO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON_NO CARTON FROM WH_AUO WHERE PALLET_NO=@PALLET_NO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@PALLET_NO", PALLET_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON2(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON FROM RMA_INVOICEOUT WHERE SHIPPING=@SHIPPING ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF2RMA2(string INVOICE, string MODEL1, string VER)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE ");
            sb.Append(" AND T2.U_TMODEL  LIKE '%" + MODEL1 + "%' AND T2.U_PARTNO LIKE '%" + VER + "%'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                iCount = new int[TempDt.Rows.Count];
                WriteDD(FileName);


            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                iCount = new int[TempDt.Rows.Count];
                WriteExcelDD(FileName);


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
                    dtCost2 = MakeMain2();
                    GD5(opdf.FileName);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                try
                {

                    textBox2.Text  = openFileDialog1.FileName;

                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                try
                {

                    textBox3.Text = openFileDialog1.FileName;

                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DELSTATUS2();
            DELETEFILE2();
            WriteExcelAP2(textBox2.Text, "1");

            WriteExcelAP2(textBox3.Text, "2");

            WriteExcelAP3(textBox3.Text);
            MessageBox.Show("匯入成功");
        }

        private void WriteExcelAP2(string ExcelFile, string DOCTYPE)
        {

            DataRow drS = null;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string PARTNO;
                string GRADE;
                string TOTAL;
                string DOCDATE;
                string DOCQTY;
                string MODEL;
                string BRAND;
                //Feb TTL
                int M1 = 0;
                int M2 = 0;
                int M3 = 0;
                int GG1 = 0;
                int QTY = 0;
                int YMM = 0;
                for (int b = 1; b <= 17; b++)
                {
                    for (int jj = 1; jj <= 5; jj++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[jj, b]);
                        range.Select();
                        string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

                        int G1 = id.IndexOf("PARTNO");
                        if (G1 != -1)
                        {

                            M2 = jj + 1;
                            M1 = b;
                            break;
                        }

                        int G2 = id.IndexOf("GRADE");

                        if (G2 != -1)
                        {
                            GG1 = b;
                            break;
                        }

                        int G4 = id.IndexOf("BRAND");

                        if (G4 != -1)
                        {
                            M3 = b;
                            break;
                        }

                        int G3 = id.IndexOf("TOTAL");
                        int GS3 = id.IndexOf("CUM");
                        int GS4 = id.IndexOf("TTL");
                        if (G3 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "TOTAL";
                            break;
                        }
                        if (GS3 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "CUM";
                            break;
                        }
                        if (GS4 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "FEBTTL";
                            break;
                        }
                    }
                }
                if (GS == "CUM" || GS == "FEBTTL")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, YMM]);
                    range.Select();
                    YM = range.Text.ToString().Trim().Replace(",", "");
                    YM = YM.Substring(0, YM.IndexOf("/"));
                    if (YM.Length == 1)
                    {
                        YM = "0" + YM;
                    }

                    YM = DateTime.Now.ToString("yyyy") + YM;
                }
                if (GS == "TOTAL")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, YMM]);
                    range.Select();
                    YM = range.Text.ToString().Trim().Replace(",", "");
                }
                int L = 0;
                for (int iRecord = M2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1]);
                    range.Select();
                    PARTNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M3]);
                    range.Select();
                    BRAND = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1 - 1]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GG1]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                    range.Select();
                    TOTAL = range.Text.ToString().Trim().Replace(",", "");
                    //APTYPE

                    if (!String.IsNullOrEmpty(PARTNO))
                    {

                        L++;


            
                        if (GS == "CUM")
                        {
                            if (DOCTYPE == "1")
                            {

                                AddSTATUS3(DOCTYPE, PARTNO, MODEL, GRADE, BRAND);
                            }

                            if (DOCTYPE == "2")
                            {
                                AddSTATUS3(DOCTYPE, PARTNO, MODEL, GRADE, BRAND);
                            }
                            for (int K = QTY + 1; K <= QTY + 31; K++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, K]);
                                range.Select();
                                DOCDATE = range.Text.ToString().Trim().Replace(",", "");

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, K]);
                                range.Select();
                                DOCQTY = range.Text.ToString().Trim().Replace(",", "");


                                if (!String.IsNullOrEmpty(DOCQTY))
                                {
                                    if (DOCTYPE == "1")
                                    {
                                        AddSTATUS2(PARTNO, GRADE, TOTAL, YM, GS, DOCDATE, DOCQTY, MODEL, BRAND);

                                    }

                                }
                            }
                        }
                    }
                }


            }
            finally
            {



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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void WriteExcelAP3(string ExcelFile)
        {

            DataRow drS = null;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string PARTNO;
                string GRADE;
                string TOTAL;
                string DOCDATE;
                string DOCQTY;
                string MODEL;
                string BRAND;
                //Feb TTL
                int M1 = 0;
                int M2 = 0;
                int M3 = 0;
                int GG1 = 0;
                int QTY = 0;
                int YMM = 0;
                for (int b = 1; b <= 17; b++)
                {
                    for (int jj = 1; jj <= 5; jj++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[jj, b]);
                        range.Select();
                        string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

                        int G1 = id.IndexOf("PARTNO");
                        if (G1 != -1)
                        {

                            M2 = jj + 1;
                            M1 = b;
                            break;
                        }


                        int G4 = id.IndexOf("BRAND");

                        if (G4 != -1)
                        {
                            M3 = b;
                            break;
                        }

                        int G2 = id.IndexOf("GRADE");

                        if (G2 != -1)
                        {
                            GG1 = b;
                            break;
                        }

                        int G3 = id.IndexOf("TOTAL");
                        int GS3 = id.IndexOf("CUM");
                        int GS4 = id.IndexOf("TTL");
                        if (G3 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "TOTAL";
                            break;
                        }
                        if (GS3 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "CUM";
                            break;
                        }
                        if (GS4 != -1)
                        {
                            QTY = b;
                            YMM = b + 1;
                            GS = "FEBTTL";
                            break;
                        }
                    }
                }
                if (GS == "CUM" || GS == "FEBTTL")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, YMM]);
                    range.Select();
                    YM = range.Text.ToString().Trim().Replace(",", "");
                    YM = YM.Substring(0, YM.IndexOf("/"));
                    if (YM.Length == 1)
                    {
                        YM = "0" + YM;
                    }

                    YM = DateTime.Now.ToString("yyyy") + YM;
                }
                if (GS == "TOTAL")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, YMM]);
                    range.Select();
                    YM = range.Text.ToString().Trim().Replace(",", "");
                }
                int L = 0;
                for (int iRecord = M2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1]);
                    range.Select();
                    PARTNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1 - 1]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M3]);
                    range.Select();
                    BRAND = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GG1]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                    range.Select();
                    TOTAL = range.Text.ToString().Trim().Replace(",", "");
                    //APTYPE

                    if (!String.IsNullOrEmpty(PARTNO))
                    {

                        L++;

                        if (L == 1)
                        {
                            //           UPSTATUS(GS, YM);
                        }

                        if (GS == "CUM")
                        {
                            for (int K = QTY + 1; K <= QTY + 31; K++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, K]);
                                range.Select();
                                DOCDATE = range.Text.ToString().Trim().Replace(",", "");

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, K]);
                                range.Select();
                                DOCQTY = range.Text.ToString().Trim().Replace(",", "");

                                int YEL = 0;
                                if (!String.IsNullOrEmpty(DOCDATE))
                                {

                                    System.Data.DataTable J2 = GJ2(PARTNO, GRADE, DOCDATE, BRAND);
                                    if (J2.Rows.Count > 0)
                                    {

                                        string DQTY = J2.Rows[0]["QTY"].ToString().Trim();
                                        string DDOCQTY = J2.Rows[0]["DOCQTY"].ToString().Trim();
                                        string DMODEL = J2.Rows[0]["MODEL"].ToString().Trim();

                                        if (DDOCQTY != DOCQTY)
                                        {
                                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                            range.ClearComments();
                                            string MM = "上一版數量 : " + DDOCQTY;
                                            range.AddComment(MM);

                                            int wCount = CountText(MM, '\n');
                                            range.Comment.Shape.Height = wCount * 20;

                                            YEL = 1;

                                        }

                                        if (TOTAL != DQTY)
                                        {

                                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                                            range.Select();

                                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                            range.ClearComments();
                                            string MM = "上一版數量 : " + DQTY;
                                            range.AddComment(MM);

                                            int wCount = CountText(MM, '\n');
                                            range.Comment.Shape.Height = wCount * 20;

                                            YEL = 1;
                                        }

                                        if (DMODEL != MODEL)
                                        {
                                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                                            range.Select();

                                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                            range.ClearComments();
                                            string MM = "上一版MODEL : " + DMODEL;
                                            range.AddComment(MM);

                                            int wCount = CountText(MM, '\n');
                                            range.Comment.Shape.Height = wCount * 20;

                                            YEL = 1;
                                        }
                                    }
                                    else
                                    {

                                        System.Data.DataTable J2S = GJ2S(PARTNO, GRADE, BRAND);
                                        if (J2S.Rows.Count > 0)
                                        {

                                            if (DOCQTY != "")
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                                range.ClearComments();
                                                string MM = "上一版數量 : 0";
                                                range.AddComment(MM);

                                                int wCount = CountText(MM, '\n');
                                                range.Comment.Shape.Height = wCount * 20;

                                                YEL = 1;
                                            }
                                        }
                                        else
                                        {

                                            System.Data.DataTable J2S2 = GJ2S3(PARTNO, GRADE, BRAND);
                                            if (J2S2.Rows.Count == 0)
                                            {
                                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1]);
                                                range.Select();
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                                range.ClearComments();
                                                string MM = "新的型號";
                                                range.AddComment(MM);

                                                int wCount = CountText(MM, '\n');
                                                range.Comment.Shape.Height = wCount * 20;

                                                YEL = 1;
                                            }
                                        }
                                    }

                                }

                                if (YEL == 1)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, M1 - 1]);
                                    range.Select();
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                }
                            }
                        }
                    }
                }

                System.Data.DataTable GGG1 = GJ2S2();
                if (GGG1.Rows.Count > 0)
                {


                    for (int i = 1; i <= GGG1.Rows.Count; i++)
                    {

                        DataRow drw3 = GGG1.Rows[i - 1];
                        string R1 = drw3[0].ToString();
                        string R2 = drw3[1].ToString();
                        string R3 = drw3[2].ToString();
                        string R4 = drw3[3].ToString();
                        string R5 = drw3[4].ToString();
                        string R6 = drw3[5].ToString();


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 1]);
                        range.Select();
                        range.Value2 = R1;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 2]);
                        range.Select();
                        range.Value2 = R2;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 3]);
                        range.Select();
                        range.Value2 = R3;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 4]);
                        range.Select();
                        range.Value2 = R4;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 5]);
                        range.Select();
                        range.Value2 = R5;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + i, 6]);
                        range.Select();
                        range.Value2 = R6;
                    }

                }

            }
            finally
            {
                         string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
               
                string NewFileName =  lsAppDir + "\\Excel\\temp\\" +
         DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

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


                         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void DELETEFILE2()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private System.Data.DataTable GJ2S(string PARTNO, string GRADE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT QTY,DOCQTY,MODEL FROM AP_STATUS2 WHERE ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND  BRAND =@BRAND AND USERS=@USERS");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }

        private System.Data.DataTable GJ2S3(string PARTNO, string GRADE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT MODEL FROM AP_STATUS3 WHERE APTYPE='1' AND  ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND  BRAND =@BRAND AND USERS=@USERS");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private System.Data.DataTable GJ2S2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '消失列',MODEL,PARTNO,GRADE,'',BRAND  FROM [AP_STATUS3] WHERE APTYPE='1' AND USERS=@USERS AND  PARTNO+' '+GRADE+' '+BRAND NOT IN (SELECT PARTNO+' '+GRADE+' '+BRAND FROM [AP_STATUS3] WHERE APTYPE='2' AND USERS=@USERS )");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        public void AddSTATUS2(string PARTNO, string GRADE, string QTY, string YM, string APTYPE,  string DOCDATE, string DOCQTY, string MODEL, string BRAND)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_STATUS2(PARTNO,GRADE,QTY,YM,APTYPE,DOCDATE,DOCQTY,MODEL,BRAND,USERS) values(@PARTNO,@GRADE,@QTY,@YM,@APTYPE,@DOCDATE,@DOCQTY,@MODEL,@BRAND,@USERS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@YM", YM));
            command.Parameters.Add(new SqlParameter("@APTYPE", APTYPE));
            command.Parameters.Add(new SqlParameter("@USERS",  fmLogin.LoginID.ToString().Trim().ToUpper()));

            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@DOCQTY", DOCQTY));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));

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

        public void AddSTATUS3(string APTYPE, string PARTNO, string MODEL, string GRADE, string BRAND)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_STATUS3(APTYPE,PARTNO,MODEL,GRADE,BRAND,USERS) values(@APTYPE,@PARTNO,@MODEL,@GRADE,@BRAND,@USERS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));

            command.Parameters.Add(new SqlParameter("@APTYPE", APTYPE));

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));
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

        public int CountText(String text, Char w)
        {
            String[] split = text.Split(w);
            int count = 0;
            foreach (String s in split)
            {
                if (s.Length > 0) count++;
            }
            return count;
        }
        private System.Data.DataTable GJ2(string PARTNO, string GRADE, string DOCDATE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT QTY,DOCQTY,MODEL FROM AP_STATUS2 WHERE ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND DOCDATE =@DOCDATE AND BRAND =@BRAND AND USERS=@USERS");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        public void DELSTATUS2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" DELETE AP_STATUS2 WHERE USERS=@USERS DELETE AP_STATUS3 WHERE USERS=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim().ToUpper()));

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


  


 
    }
}
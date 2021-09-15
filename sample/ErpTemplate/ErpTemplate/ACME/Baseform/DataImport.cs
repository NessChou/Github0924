using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class DataImport : Form
    {
        private System.Data.DataTable TempDt;
        private string FileName;
        public DataImport()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                GetExcelContentGD4(opdf.FileName);
            }
        }
        private void GetExcelContent1(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
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

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2;
            string id3;
            string id4;
            string id5;
            string id6;
            string id7;
            string id8;
            string id9;
            string id10;
            string id11;
            string id12;
            string id13;
            string id14;
            string id15;
            string id16;

            for (int i = 2; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id4 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                id5 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                id6 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                id7 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                id8 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                range.Select();
                id9 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                range.Select();
                id10 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                range.Select();
                id11 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                range.Select();
                id12 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 13]);
                range.Select();
                id13 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 14]);
                range.Select();
                id14 = range.Text.ToString();
                string ver = id4.Substring(9, 1);


                strText = "";
                try
                {
                    AddAUOGD(id, id2, id3, id4, id5, id6,ver, "200805", id7, id8, id9, id10, id11, id12, id13, id14);

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
        }
        private void GetExcelContentGD4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
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

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string idx;
            string id2;
            string id3;
            string id4;
            string id5;
            string id6;

            for (int i = 3; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                idx = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id2 = range.Text.ToString()+'.'+idx.Substring(9,1);
                
                         
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                id3 = range.Text.ToString();
    
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                id4 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                id5 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                id6 = range.Text.ToString();

                strText = "";
                try
                {
                    AddAUOGD4(id,id2,id3,id4,id5,"","200903");

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
        }
        private void GetExcelContentGD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
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

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string idx;
            string id2;
            string id3;
            string id4;
            string id5;
            string id6;

            for (int i = 3; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id4 = range.Text.ToString();

         

                strText = "";
                try
                {
                    AddAUOGD5(id, id2, id3, id4);

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
        }
        private void GetExcelContent2(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
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

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2;
            string id3;
            string id4;
            string id5;
            string id6;
         
            for (int i = 2; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id4 = range.Text.ToString();


       


                strText = "";
                try
                {
                    AddAUOGD2(id, id2, id3, id4);

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
        }


        private void AddTRACKER_LOG(string KIND_NO, string BOOK_NO, string BOOK_NAME, string AUTHOR, string PRESS, string DONATOR)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO [AcmeSqlSP].[dbo].[AUOGD]");
            sb.Append("            ([DocEntry]");
            sb.Append("            ,[CardCode]");
            sb.Append("            ,[Size]");
            sb.Append("            ,[Model]");
            sb.Append("            ,[PartNo]");
            sb.Append("            ,[Grade]");
            sb.Append("            ,[Site]");
            sb.Append("            ,[Ver]");
            sb.Append("            ,[FocDate]");
            sb.Append("            ,[Qty1]");
            sb.Append("            ,[Qty12]");
            sb.Append("            ,[Qty2]");
            sb.Append("            ,[Qty22]");
            sb.Append("            ,[Qty3]");
            sb.Append("            ,[Qty32]");
            sb.Append("            ,[Qty4]");
            sb.Append("            ,[Qty42])");
            sb.Append("      VALUES");
            sb.Append("            (<DocEntry, int,>");
            sb.Append("            ,<CardCode, nvarchar(50),>");
            sb.Append("            ,<Size, nchar(10),>");
            sb.Append("            ,<Model, nchar(20),>");
            sb.Append("            ,<PartNo, nchar(20),>");
            sb.Append("            ,<Grade, nchar(5),>");
            sb.Append("            ,<Site, nchar(10),>");
            sb.Append("            ,<Ver, nchar(5),>");
            sb.Append("            ,<FocDate, nchar(10),>");
            sb.Append("            ,<Qty1, nchar(10),>");
            sb.Append("            ,<Qty12, nchar(10),>");
            sb.Append("            ,<Qty2, nchar(10),>");
            sb.Append("            ,<Qty22, nchar(10),>");
            sb.Append("            ,<Qty3, nchar(10),>");
            sb.Append("            ,<Qty32, nchar(10),>");
            sb.Append("            ,<Qty4, nchar(10),>");
            sb.Append("            ,<Qty42, nchar(10),>)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@KIND_NO", KIND_NO));
            command.Parameters.Add(new SqlParameter("@BOOK_NO", BOOK_NO));
            command.Parameters.Add(new SqlParameter("@BOOK_NAME", BOOK_NAME));
            command.Parameters.Add(new SqlParameter("@AUTHOR", AUTHOR));
            command.Parameters.Add(new SqlParameter("@PRESS", PRESS));
            command.Parameters.Add(new SqlParameter("@DONATOR", DONATOR));


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
        private void WriteExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();


           // excelApp.Visible = checkBox1.Checked;

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
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();
                    range.Select();
                    DataTable aa = GetOrderData2();
                    drFind = aa.Rows.Find(SERIAL_NO);

                    if (drFind != null)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        range.Value2 = "Y";
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
   
            }

           // dataGridView1.DataSource = TempDt;
        }
        public void AddAUOGD(string CardCode,string Size,string Model,string PartNo,string Grade,string Site,string Ver,string FocDate,string Qty1,string Qty12,string Qty2,string Qty22,string Qty3,string Qty32,string Qty4,string Qty42)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AUOGD(CardCode,Size,Model,PartNo,Grade,Site,Ver,FocDate,Qty1,Qty12,Qty2,Qty22,Qty3,Qty32,Qty4,Qty42) values(@CardCode,@Size,@Model,@PartNo,@Grade,@Site,@Ver,@FocDate,@Qty1,@Qty12,@Qty2,@Qty22,@Qty3,@Qty32,@Qty4,@Qty42)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));      
            command.Parameters.Add(new SqlParameter("@Size", Size));
            command.Parameters.Add(new SqlParameter("@Model", Model));     
            command.Parameters.Add(new SqlParameter("@PartNo", PartNo));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@Site", Site));
            command.Parameters.Add(new SqlParameter("@Ver", Ver));
            command.Parameters.Add(new SqlParameter("@FocDate", FocDate));
            command.Parameters.Add(new SqlParameter("@Qty1", Qty1));
            command.Parameters.Add(new SqlParameter("@Qty12", Qty12));
            command.Parameters.Add(new SqlParameter("@Qty2", Qty2));
            command.Parameters.Add(new SqlParameter("@Qty22", Qty22));
            command.Parameters.Add(new SqlParameter("@Qty3", Qty3));
            command.Parameters.Add(new SqlParameter("@Qty32", Qty32));
            command.Parameters.Add(new SqlParameter("@Qty4", Qty4));
            command.Parameters.Add(new SqlParameter("@Qty42", Qty42));
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

        public void AddAUOGD4(string Size, string Model, string Grade, string CardCode, string ForCast, string Commend,string GD4Month)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AUOGD4(Size,Model,Grade,CardCode,ForCast,Commend,GD4Month) values(@Size,@Model,@Grade,@CardCode,@ForCast,@Commend,@GD4Month)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Size", Size));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@ForCast", ForCast));
            command.Parameters.Add(new SqlParameter("@Commend", Commend));
            command.Parameters.Add(new SqlParameter("@GD4Month", GD4Month));
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

        public void AddAUOGD5(string Size, string Model, string Grade, string CardCode)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AUOGD5(Size,Model,Grade,CardCode) values(@Size,@Model,@Grade,@CardCode)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Size", Size));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
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
        public void AddAUOGD2(string Customer,string Size,string Model,string Grade)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AUOGD3(Customer,Size,Model,Grade) values(@Customer,@Size,@Model,@Grade)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Customer",Customer));
            command.Parameters.Add(new SqlParameter("@Size", Size));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
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

        private void DataImport_Load(object sender, EventArgs e)
        {

        }
        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select Model from auogd4 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                //MessageBox.Show(FileName);



       //         iCount = new int[TempDt.Rows.Count];
                WriteExcelProduct(FileName);


            }
        }
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

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                GetExcelContentGD5(opdf.FileName);
            }
        }
    }
}
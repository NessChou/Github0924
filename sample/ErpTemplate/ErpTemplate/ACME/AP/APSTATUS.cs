using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class APSTATUS : Form
    {
        private string FileName;
        public APSTATUS()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

                     string[] filecSize = Directory.GetFiles("//acmew08r2ap//採購貨狀EXCEL");

                     TRUNSTATUS();
                     foreach (string fie in filecSize)
                     {
                         FileInfo filess = new FileInfo(fie);
                         string dd = filess.Name.ToString();

                         int ad = dd.LastIndexOf(".");

                         string PPATH = "//acmew08r2ap//採購貨狀EXCEL//" + dd;
                         string PanelName = dd.Substring(0, ad).ToString();
                         WriteExcelAP2(fie.ToString(), PanelName, PPATH);
                     }

                     MessageBox.Show("匯入成功");
                     dataGridView1.DataSource = GetG1();
        }
        string GS = "";
        string YM = "";
        private void WriteExcelAP(string ExcelFile, string FILENAME, string PPATH)
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

                        if (L == 1)
                        {
                            UPSTATUS(GS, YM);
                        }
                        //System.Data.DataTable GT = GETSTATUS(PARTNO, GRADE, YM);
                        //if (GT.Rows.Count > 0)
                        //{
                        //    UPSTATUS(PARTNO, GRADE, TOTAL, YM);
                        //}
                        //else
                        //{
                        //    AddSTATUS(PARTNO, GRADE, TOTAL, YM);
                        //}

                        AddSTATUS(PARTNO, GRADE, TOTAL, YM, GS, FILENAME, ExcelFile);
                        //if (GS == "CUM")
                        //{
                        //    for (int K = QTY + 1; K <= QTY + 30; K++)
                        //    {
                        //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, K]);
                        //        range.Select();
                        //        DOCDATE = range.Text.ToString().Trim().Replace(",", "");

                        //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, K]);
                        //        range.Select();
                        //        DOCQTY = range.Text.ToString().Trim().Replace(",", "");

                        //        if (!String.IsNullOrEmpty(DOCQTY))
                        //        {
                        //            AddSTATUS2(PARTNO, GRADE, TOTAL, YM, GS, FILENAME, PPATH, DOCDATE, DOCQTY, MODEL, BRAND);
                        //        }
                        //    }
                        //}
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
        private void WriteExcelAP2(string ExcelFile, string FILENAME, string PPATH)
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

                int L = 0;
                for (int iRecord = 1; iRecord <= iRowCnt-1; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    PARTNO = range.Text.ToString().Trim();

 


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    TOTAL = range.Text.ToString().Trim().Replace(",", "");
                    //APTYPE

                    if (!String.IsNullOrEmpty(PARTNO))
                    {
                        if (PARTNO.ToLower().Trim() != "part no")
                        {
                            if (!String.IsNullOrEmpty(TOTAL))
                            {
                                AddSTATUS(PARTNO, GRADE, TOTAL, YM, GS, FILENAME, ExcelFile);
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

        private DataTable GetG1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT [FILENAME]  檔案已更新 FROM AP_STATUS  WHERE YM=Convert(varchar(6),GETDATE() ,112)   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

        private DataTable GJ1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  FILEPATH FROM AP_STATUS2  WHERE YM=Convert(varchar(6),GETDATE() ,112) GROUP BY FILEPATH");
            sb.Append("   ORDER BY MAX(ID) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

        private DataTable GJ2(string PARTNO, string GRADE, string DOCDATE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT QTY,DOCQTY,MODEL FROM AP_STATUS2 WHERE ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND DOCDATE =@DOCDATE AND BRAND =@BRAND");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
        
                        command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
                        command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
                        command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
                        command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
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
        private DataTable GJ2S(string PARTNO, string GRADE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT QTY,DOCQTY,MODEL FROM AP_STATUS2 WHERE ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND  BRAND =@BRAND");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
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

        private DataTable GJ2S3(string PARTNO, string GRADE, string BRAND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT MODEL FROM AP_STATUS3 WHERE APTYPE='1' AND  ");
            sb.Append("    PARTNO=@PARTNO AND GRADE=@GRADE AND  BRAND =@BRAND");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));
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
        private DataTable GJ2S2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '消失列',MODEL,PARTNO,GRADE,'',BRAND  FROM [AP_STATUS3] WHERE APTYPE='1' AND  PARTNO+' '+GRADE+' '+BRAND NOT IN (SELECT PARTNO+' '+GRADE+' '+BRAND FROM [AP_STATUS3] WHERE APTYPE='2' )");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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
        public void AddSTATUS(string PARTNO, string GRADE, string QTY, string YM, string APTYPE, string FILENAME, string FILEPATH)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_STATUS(PARTNO,GRADE,QTY,YM,APTYPE,FILENAME,FILEPATH) values(@PARTNO,@GRADE,@QTY,@YM,@APTYPE,@FILENAME,@FILEPATH)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@YM", YM));
            command.Parameters.Add(new SqlParameter("@APTYPE", APTYPE));
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
            command.Parameters.Add(new SqlParameter("@FILEPATH", FILEPATH));
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

        public void UPSTATUS(string APTYPE, string YM)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_STATUS WHERE APTYPE=@APTYPE AND YM=@YM  ", connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@APTYPE", APTYPE));
            command.Parameters.Add(new SqlParameter("@YM", YM));
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

        public void TRUNSTATUS()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE AP_STATUS ", connection);
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

 
      
        private void WriteExcelProduct4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string DOCDATE;

                string WHNO;
                string FEE;
                string FEEB;
                string FEE2;
                string FEE3;
                string FEE4;
                string TS;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    WHNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    FEE = range.Text.ToString().Trim();

                    //System.Data.DataTable GG1 = Getdata3(WHNO);
                    //if (GG1.Rows.Count > 0)
                    //{

                    //    string FEES = GG1.Rows[0]["FEE"].ToString();
                    //    DateTime D1 = Convert.ToDateTime(DOCDATE);
                    //    DateTime D2 = Convert.ToDateTime(GG1.Rows[0]["DOCDATE"]);
                    //    if (FEE != FEES)
                    //    {
                    //        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    //        range.ClearComments();
                    //        string MM = "輔助金額 : " + FEES;
                    //        range.AddComment(MM);

                    //        int wCount = CountText(MM, '\n');
                    //        range.Comment.Shape.Height = wCount * 20;
                    //    }

                    //    if (D1 != D2)
                    //    {
                    //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    //        range.Select();
                    //        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    //        range.ClearComments();
                    //        string MM = "輔助日期: " + D2;
                    //        range.AddComment(MM);

                    //        int wCount = CountText(MM, '\n');
                    //        range.Comment.Shape.Height = wCount * 20;
                    //    }



                    //}

               

                }

  


            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


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

        
    }
}

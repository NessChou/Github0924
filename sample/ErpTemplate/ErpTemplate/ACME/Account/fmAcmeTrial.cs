using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Net.Mime;
using System.IO;

namespace ACME
{
    public partial class fmAcmeTrial : Form
    {
        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

        System.Data.DataTable dtc = null;

        string NewFileName;
        string YYYY = "";
        string MONTH = "";
        string COMPANY;
        public fmAcmeTrial()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {
                    string file = openFileDialog1.FileName;
                    string filename = Path.GetFileName(openFileDialog1.FileName);
                    string FILE = DateTime.Now.ToString("yyyyMMddHHmmss") + filename;
                    string ss = lsAppDir + "\\EXCEL\\temp\\" + FILE;
                    System.IO.File.Copy(file, ss, true);
                    string server = "//Acmew08r2ap//EXEXPORT//OUTPUT//";
                    bool F1 = getrma.UploadFile(ss, server, false);

                    if (F1 == false)
                    {
                        return;
                    }

                    GetMenu.InsertEXEXPORT("02試算表", textBox3.Text + "/" + textBox4.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), server + FILE);

                    MessageBox.Show("訊息已送出");

                    //FileName = openFileDialog1.FileName;

                    //GetExcelProduct(FileName);

                    //MessageBox.Show("產生檔案->" + NewFileName);
               
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }

        }

        private void GetExcelProduct(string ExcelFile)
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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            int 資產 = 0;
            int 負債 = 0;
            int 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 5]);
            range.Select();
            資產 = Convert.ToInt32(range.Value2);

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;

                int Line_Liab = 0;

                for (int iRecord = iRowCnt; iRecord >= 1; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    int iPos = sTemp.IndexOf("-");

                    if (iPos > 0)
                    {

                        string s = sTemp.Substring(0, iPos - 1);



                        if (s.Length == 4)
                        {

                            range.Select();


                            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                        }

                        else if (s.Length == 8)
                        {



                            range.InsertIndent(1);

                        }

                    }



                }


                object Cell_From;
                object Cell_To;






                //// //指定AV 的範圍
                object SelectCell_From = "B2";
                object SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";
              
                SelectCell_From = "A1";
                SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();


                //設定 width
                // ((Range)excelSheet.Columns["B", oMissing]).ColumnWidth = 40;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);

                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = "D1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "試算表";
                //加底線
                range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                range.Font.Size = 16;


                SelectCell_From = "A2";
                SelectCell_To = "D2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 1]);
                range.Select();
                range.Value2 = "日期:從 " + textBox1.Text + " 至 " + textBox2.Text + " 止";
                SelectCell_From = "A3";
                SelectCell_To = "D3";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "A4";
                SelectCell_To = "D4";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;




            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
               "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";
                //GetFileName(ExcelFile);
                //   MessageBox.Show(NewFileName);

                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                // MessageBox.Show("產生一個檔案->" + NewFileName);



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

        //日期處理--------------------------------------------------------------------------------------------

        private bool isNumber(string s)
        {
            int Flag = 0;
            char[] str = s.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                if (Char.IsNumber(str[i]))
                {
                    Flag++;
                }
                else
                {
                    Flag = -1;
                    break;
                }
            }
            if (Flag > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void fmAcmeTrial_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
            textBox3.Text = GetMenu.DFirst();
            textBox4.Text = GetMenu.DLast();
            textBox5.Text = GetMenu.Day();
            textBox6.Text = GetMenu.Day();
            textBox10.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox11.Text = DateTime.Now.ToString("yyyyMM");

            string PriorMonth = DateToStr(DateTime.Now.AddMonths(-1)).Substring(0, 6);
            int year = Convert.ToInt32(PriorMonth.Substring(0, 4));
            int month = Convert.ToInt32(PriorMonth.Substring(4, 2));
            //取得當月天數
            int days = DateTime.DaysInMonth(year, month);

            textBox8.Text = PriorMonth.Substring(0, 4) + "." + PriorMonth.Substring(4, 2) + "." + days.ToString();

            COMPANY = GetCOMPANY().Rows[0][0].ToString();

        }

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\試算表後製作業.doc";
            System.Diagnostics.Process.Start(aa);
        }

        private void button2_Click(object sender, EventArgs e)
        {


        }
        private void GetExcelProduct2(string ExcelFile, string NewFileName)
        {

            //Create an Excel App
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



            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            int 負債 = 0;
            int 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;



            try
            {

                string sTemp = string.Empty;
                string sTemp2 = string.Empty;
                string sTemp3 = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;
                int Line_Liab = 0;


                for (int iRecord = iRowCnt; iRecord >= 2; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();


                    
                    int iPos = sTemp.IndexOf("-");

                    if ((iPos < 0) || sTemp == ("保險費-員工") || sTemp == ("保險費-產物") || sTemp == ("折    舊-服務成本"))
                    {


                        range.Select();
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                        range.Font.Size = 12;

                    }
                    else
                    {

                        range.Select();
                        range.Value2 = range.Value2.ToString();

                    }



                    if (sTemp == "銷貨成本")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = "C" + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "銷貨毛利")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = "C" + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }
                    if (sTemp == "營業淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = "C" + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "稅前淨利(損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = "C" + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }
                    if (sTemp == "本期淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = "C" + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    }



                }


                object Cell_From;
                object Cell_To;
                object FixCell;


                //// //指定AV 的範圍
                object SelectCell_From = "B2";
                object SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

                SelectCell_From = "A1";
                SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);

                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = "D1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "損益表";
                //加底線
                range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                range.Font.Size = 16;


                SelectCell_From = "A2";
                SelectCell_To = "D2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 1]);
                range.Select();
                range.Value2 = "日期:從 " + textBox3.Text + " 至 " + textBox4.Text + " 止";
                SelectCell_From = "A3";
                SelectCell_To = "D3";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "A4";
                SelectCell_To = "D4";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                SelectCell_From = "B5";
                SelectCell_To = "C" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;


            }
            finally
            {




                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

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

             //   SENDMAIL();

            }

        }
        private void GetExcelINCOME(string ExcelFile, string NewFileName)
        {

            //Create an Excel App
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



            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            int 負債 = 0;
            int 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;



            try
            {

                string sTemp = string.Empty;
                string sTemp2 = string.Empty;
                string sTemp3 = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;
                int Line_Liab = 0;

                int mm = Convert.ToInt16(MONTH);


                string N = "";
                if (mm == 1)
                {
                    N = "C";
                }
                if (mm == 2)
                {
                    N = "D";
                }
                if (mm == 3)
                {
                    N = "E";
                }
                if (mm == 4)
                {
                    N = "F";
                }
                if (mm == 5)
                {
                    N = "G";
                }
                if (mm == 6)
                {
                    N = "H";
                }
                if (mm == 7)
                {
                    N = "I";
                }
                if (mm == 8)
                {
                    N = "J";
                }
                if (mm == 9)
                {
                    N = "K";
                }
                if (mm == 10)
                {
                    N = "L";
                }
                if (mm == 11)
                {
                    N = "M";

                }
                if (mm == 12)
                {
                    N = "N";

                }
                for (int iRecord = iRowCnt; iRecord >= 2; iRecord--)
                {

                
                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();



                    int iPos = sTemp.IndexOf("-");

                    if ((iPos < 0) || sTemp == ("保險費-員工") || sTemp == ("保險費-產物") || sTemp == ("折    舊-服務成本"))
                    {


                        range.Select();
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        range.Font.Size = 12;
                        range.Font.Bold = true;

                        object AFrom = "B" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Font.Bold = true;
                       
         
                    }
                    else
                    {

                        range.Select();
                        range.Value2 = range.Value2.ToString();
                  
                    }

             
                    if (sTemp == "銷貨成本")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "銷貨毛利")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }
                    if (sTemp == "營業淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "稅前淨利(損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "本期淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    }

                    for (int S = 1; S <= mm+2; S++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, S]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        if (S == mm+2)
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDouble;
                        }
                        else
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        }

                        if ((iPos < 0) || sTemp == ("保險費-員工") || sTemp == ("保險費-產物") || sTemp == ("折    舊-服務成本"))
                        {
                            if (String.IsNullOrEmpty(sTemp))
                            {

                                range.Value2 = "0";
                            }
                        }
                    }

                }


                object Cell_From;
                object Cell_To;
                object FixCell;


                //// //指定AV 的範圍
                object SelectCell_From = "B2";
                object SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

                SelectCell_From = "A1";
                SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();

                SelectCell_From = "A1";
                SelectCell_To = N + "1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                    range.Select();
                    range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
                    range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                    SelectCell_From = "A1";
                    SelectCell_To = "A" + Convert.ToString(iRowCnt);
                    range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                    range.Select();
                    range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;

                    
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);





                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = N + "1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "損益表";
                //加底線
                range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 2]);
                range.Select();
                range.Value2 = YYYY + "年總計";
                range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Bold = true;
                for (int iRecord = 1; iRecord <= mm; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 2 + iRecord]);
                    range.Select();
                    range.Value2 = YYYY + "/" + iRecord.ToString();
                    range.Font.Size = 10;
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Font.Bold = true;
                    if (iRecord == mm)
                    {
                        range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDouble;
                    }
                    else
                    {
                        range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }
                }

                SelectCell_From = "A2";
                SelectCell_To = N + "2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "B5";
                SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;


            }
            finally
            {




                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

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

        private void GetExcelINCOME2(string ExcelFile, string NewFileName)
        {

            //Create an Excel App
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



            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            int 負債 = 0;
            int 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;



            try
            {

                string sTemp = string.Empty;
                string sTemp2 = string.Empty;
                string sTemp3 = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;
                int Line_Liab = 0;

                int mm = Convert.ToInt16(MONTH);
                if (mm >= 1 && mm <= 3)
                {
                    mm = mm + 1;
                }
                else if (mm >= 4 && mm <= 6)
                {
                    mm = mm + 2;
                }
                else if (mm >= 7 && mm <= 9)
                {
                    mm = mm + 3;
                }
                else if (mm >= 10 && mm <= 12)
                {
                    mm = mm + 4;
                }

                string N = "";
                if (mm == 1)
                {
                    N = "C";
                }
                if (mm == 2)
                {
                    N = "D";
                }
                if (mm == 3)
                {
                    N = "E";
                }
                if (mm == 4)
                {
                    N = "F";
                }
                if (mm == 5)
                {
                    N = "G";
                }
                if (mm == 6)
                {
                    N = "H";
                }
                if (mm == 7)
                {
                    N = "I";
                }
                if (mm == 8)
                {
                    N = "J";
                }
                if (mm == 9)
                {
                    N = "K";
                }
                if (mm == 10)
                {
                    N = "L";
                }
                if (mm == 11)
                {
                    N = "M";

                }
                if (mm == 12)
                {
                    N = "N";
                }
                if (mm == 13)
                {
                    N = "O";
                }
                if (mm == 14)
                {
                    N = "P";
                }
                if (mm == 15)
                {
                    N = "Q";
                }
                if (mm == 16)
                {
                    N = "R";
                }
                for (int iRecord = iRowCnt; iRecord >= 2; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();



                    int iPos = sTemp.IndexOf("-");

                    if ((iPos < 0) || sTemp == ("保險費-員工") || sTemp == ("保險費-產物") || sTemp == ("折    舊-服務成本"))
                    {


                        range.Select();
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        range.Font.Size = 12;
                        range.Font.Bold = true;

                        object AFrom = "B" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Font.Bold = true;


                    }
                    else
                    {

                        range.Select();
                        range.Value2 = range.Value2.ToString();

                    }


                    if (sTemp == "銷貨成本")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "銷貨毛利")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }
                    if (sTemp == "營業淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "稅前淨利(損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    }

                    if (sTemp == "本期淨利(淨損)")
                    {

                        object AFrom = "A" + iRecord;
                        object ATo = N + iRecord;
                        range = excelSheet.get_Range(AFrom, ATo);
                        range.Select();
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    }

                    for (int S = 1; S <= mm + 2; S++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, S]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        if (S == mm + 2)
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDouble;
                        }
                        else
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        }

                        if ((iPos < 0) || sTemp == ("保險費-員工") || sTemp == ("保險費-產物") || sTemp == ("折    舊-服務成本"))
                        {
                            if (String.IsNullOrEmpty(sTemp))
                            {

                                range.Value2 = "0";
                            }
                        }
                    }

                }


                object Cell_From;
                object Cell_To;
                object FixCell;


                //// //指定AV 的範圍
                object SelectCell_From = "B2";
                object SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

                SelectCell_From = "A1";
                SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();

                SelectCell_From = "A1";
                SelectCell_To = N + "1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                SelectCell_From = "A1";
                SelectCell_To = "A" + Convert.ToString(iRowCnt);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);





                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = N + "1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "損益表";
                //加底線
                range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 2]);
                range.Select();
                range.Value2 = YYYY + "年總計";
                range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Bold = true;
                System.Data.DataTable K1 = GetINCOME3(YYYY);
                for (int iRecord = 0; iRecord <= K1.Rows.Count-1; iRecord++)
                {
                    string H1 = K1.Rows[iRecord][0].ToString().Trim();
                    int H2 = Convert.ToInt16(H1.Substring(4, 2));
                    string YEAR = textBox11.Text.Substring(0, 4);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 3 + iRecord]);
                    range.Select();
                    if (H1 == YYYY +"031")
                    {
                        range.Value2 = YEAR + "Q1";
                    }
                    else if (H1 == YYYY + "061")
                    {
                        range.Value2 = YEAR + "Q2";
                    }
                    else if (H1 == YYYY + "091")
                    {
                        range.Value2 = YEAR + "Q3";
                    }
                    else if (H1 == YYYY + "121")
                    {
                        range.Value2 = YEAR + "Q4";
                    }
                    else
                    {
                        range.Value2 = YYYY + "/" + H2.ToString();
                    }
                    range.Font.Size = 10;
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Font.Bold = true;
                    if (iRecord == mm)
                    {
                        range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDouble;
                    }
                    else
                    {
                        range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }
                }

                SelectCell_From = "A2";
                SelectCell_To = N + "2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "B5";
                SelectCell_To = N + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;


            }
            finally
            {




                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

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
        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                //try
                //{
                //    FileName = openFileDialog1.FileName;

                //    GetExcelProduct3(FileName);

                //    MessageBox.Show("產生檔案->" + NewFileName);
                //}
                //finally
                //{
                //    Cursor = Cursors.Default;
                //}
                try
                {
 
                    string file = openFileDialog1.FileName;
                    string filename = Path.GetFileName(openFileDialog1.FileName);
                    string FILE = DateTime.Now.ToString("yyyyMMddHHmmss") + filename;
                    string ss = lsAppDir + "\\EXCEL\\temp\\" + FILE;
                    System.IO.File.Copy(file, ss, true);
                    string server = "//Acmew08r2ap//EXEXPORT//OUTPUT//";
                    bool F1 = getrma.UploadFile(ss, server, false);

                    if (F1 == false)
                    {
                        return;
                    }

                    GetMenu.InsertEXEXPORT("03明細分類帳", textBox5.Text + "/" + textBox6.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), server + FILE);

                    MessageBox.Show("訊息已送出");



                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void GetExcelProduct3(string ExcelFile)
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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 4]);
            range.Select();


            try
            {

                string sTemp = string.Empty;
                string sTemp2 = string.Empty;
                string sTemp3 = string.Empty;
                string sTemp4 = string.Empty;
                string FieldValue = string.Empty;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 6]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 7]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 7]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 8]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);
                for (int iRecord = iRowCnt; iRecord >= 2; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp2 = (string)range.Text;
                    sTemp2 = sTemp2.Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    sTemp3 = (string)range.Text;
                    sTemp3 = sTemp3.Trim();


                    if (sTemp.Trim() == "期間小計")
                    {
                        for (int i = 5; i <= 6; i++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, i]);
                            range.Select();
                            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        }
                    }



                    if (sTemp2 == "費用")
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 4]);
                        range.Select();
                        range.Value2 = sTemp3;


                    }


                }



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 2]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 2]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);



                //// //指定AV 的範圍
                object SelectCell_From = "D2";
                object SelectCell_To = "E" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

                SelectCell_From = "A1";
                SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);

                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = "D1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //range.Font.Size = 16;
                range.Font.Bold = true;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "明細分類帳";
                range.Font.Bold = true;


                SelectCell_From = "A2";
                SelectCell_To = "D2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 1]);
                range.Select();
                range.Value2 = textBox5.Text + " ~ " + textBox6.Text;
                SelectCell_From = "A3";
                SelectCell_To = "D3";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.Font.Bold = true;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "A3";
                SelectCell_To = "E3";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;


            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
               "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";
                //GetFileName(ExcelFile);
                //   MessageBox.Show(NewFileName);

                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }




        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\資產負債表後製作業.doc";
            System.Diagnostics.Process.Start(aa);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\損益表後製作業.doc";
            System.Diagnostics.Process.Start(aa);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {
                    //FileName = openFileDialog1.FileName;

                    //GetExcelProduct5(FileName);

                    //MessageBox.Show("產生檔案->" + NewFileName);

                    try
                    {
                        string file = openFileDialog1.FileName;
                        string filename = Path.GetFileName(openFileDialog1.FileName);
                        string FILE = DateTime.Now.ToString("yyyyMMddHHmmss") + filename;
                        string ss = lsAppDir + "\\EXCEL\\temp\\" + FILE;
                        System.IO.File.Copy(file, ss, true);
                        string server = "//Acmew08r2ap//EXEXPORT//OUTPUT//";
                        bool F1 = getrma.UploadFile(ss, server, false);

                        if (F1 == false)
                        {
                            return;
                        }

                        GetMenu.InsertEXEXPORT("04資產負債表", textBox8.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), server + FILE);

                        MessageBox.Show("訊息已送出");



                    }
                    finally
                    {
                        Cursor = Cursors.Default;
                    }
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }


        private void GetExcelProduct5(string ExcelFile)
        {

            //Create an Excel App
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

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Int64 資產 = 0;
            Int64 負債 = 0;
            Int64 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 4]);
            range.Select();
            資產 = Convert.ToInt64(range.Value2);

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;

                int Line_Liab = 0;

                for (int iRecord = iRowCnt; iRecord >= 1; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    //如果欄位值是負債,就要搬移至 右方

                    if (string.IsNullOrEmpty(sTemp))
                    {
                        range.EntireRow.Delete(XlDirection.xlDown);
                    }

                    if (sTemp == "業主權益")
                    {
                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        range.Select();
                        業主權益 = Convert.ToInt64(range.Value2);

                    }



                    if (sTemp == "負債")
                    {
                        Line_Liab = iRecord;
                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        range.Select();
                        負債 = Convert.ToInt64(range.Value2);

                        break;
                    }


                }


                object Cell_From;
                object Cell_To;
                object FixCell;


                // 指定 複製 的範圍
                Cell_From = "A1";
                Cell_To = "D1";
                excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                range.Select();
                FixCell = "E1";
                range = excelSheet.get_Range(FixCell, FixCell);
                range.Select();
                excelSheet.Paste(oMissing, oMissing);

                Cell_From = "A" + Convert.ToString(Line_Liab);
                Cell_To = "D" + Convert.ToString(iRowCnt + 1);
                excelSheet.get_Range(Cell_From, Cell_To).Cut(oMissing);
                range.Select();

                FixCell = "E2";
                range = excelSheet.get_Range(FixCell, FixCell);
                range.Select();
                excelSheet.Paste(oMissing, oMissing);

                object SelectCell_From = "D2";
                object SelectCell_To = "D" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";
    
                SelectCell_From = "H2";
                SelectCell_To = "H" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


                SelectCell_From = "A1";
                SelectCell_To = "H" + Convert.ToString(iRowCnt + 1);
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Columns.AutoFit();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                //插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);


                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);

                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = COMPANY;
                SelectCell_From = "A1";
                SelectCell_To = "H1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Font.Size = 16;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 1]);
                range.Select();
                range.Value2 = "資產負債表";
                //加底線
                range.Font.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
                range.Font.Size = 16;


                SelectCell_From = "A2";
                SelectCell_To = "H2";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 1]);
                range.Select();
                range.Value2 = "日期:" + textBox8.Text + "止";
                SelectCell_From = "A3";
                SelectCell_To = "H3";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                SelectCell_From = "A4";
                SelectCell_To = "H4";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;


                //畫線
                SelectCell_From = "A" + (Line_Liab + 2).ToString();
                SelectCell_To = "H" + (Line_Liab + 2).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;


                SelectCell_From = "A" + (Line_Liab + 3).ToString();
                SelectCell_To = "H" + (Line_Liab + 3).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;



                //補上總計
                SelectCell_From = "A" + (Line_Liab + 3).ToString();
                SelectCell_To = "A" + (Line_Liab + 3).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Value2 = "資產總計";
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                //補上總計
                SelectCell_From = "E" + (Line_Liab + 3).ToString();
                SelectCell_To = "E" + (Line_Liab + 3).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Value2 = "負債及業主權益總計";
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;


                //填入總計
                SelectCell_From = "D" + (Line_Liab + 3).ToString();
                SelectCell_To = "D" + (Line_Liab + 3).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Value2 = 資產;
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                SelectCell_From = "H" + (Line_Liab + 3).ToString();
                SelectCell_To = "H" + (Line_Liab + 3).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Value2 = 負債 + 業主權益;
                range.HorizontalAlignment = XlHAlign.xlHAlignRight;


                //刪除 FGCB
                ////刪除欄
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 7]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 6]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 3]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 2]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlToLeft);


                //刪除列
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 1]);
                range.Select();
                range.EntireRow.Delete(XlDirection.xlUp);

                //劃一中線
                SelectCell_From = "C5";
                SelectCell_To = "C" + (Line_Liab + 2).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;




                string str = "";
                for (int i = 5; i <= Line_Liab + 2; i++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i, 1]);

                    try
                    {
                        str = range.Text.ToString().Substring(0, 1);
                        if (Char.IsNumber(str[0]))
                        {
                            range.InsertIndent(1);
                        }
                        else
                        {
                            range.Font.ColorIndex = 5;
                        }

                    }
                    catch
                    {

                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i, 3]);

                    try
                    {
                        if (!String.IsNullOrEmpty(range.Text.ToString()))
                        {
                            str = range.Text.ToString().Substring(0, 1);
                            if (Char.IsNumber(str[0]))
                            {
                                range.InsertIndent(1);
                            }
                            else
                            {
                                range.Font.ColorIndex = 5;
                            }
                        }
                    }
                    catch
                    {

                    }

                }


                //找到業主權益,在上面畫一條分隔線
                for (int i = 5; i <= Line_Liab + 2; i++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i, 3]);
                    if (range.Text.ToString() == "業主權益")
                    {

                        SelectCell_From = "C" + (i).ToString();
                        SelectCell_To = "D" + (i).ToString();
                        range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                        range.Select();

                        range.Insert(XlDirection.xlDown, oMissing);


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i, 3]);
                        range.Value2 = "負債總計";
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i, 4]);
                        range.Value2 = 負債;
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[i + 1, 4]);
                        range.Value2 = "";

                        break;
                    }
                }

                SelectCell_From = "C" + (Line_Liab + 2).ToString();
                SelectCell_To = "D" + (Line_Liab + 2).ToString();
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Delete(XlDirection.xlUp);

                SelectCell_From = "A1";
                SelectCell_To = "A1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();



            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
               "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";


                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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



            }

        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\總帳後製作業.doc";
            System.Diagnostics.Process.Start(aa);
        }



        public System.Data.DataTable GetPATH()
        {
            SqlConnection connection = globals.Connection;

            string sql = "SELECT PARAM_NO  FROM RMA_PARAMS WHERE PARAM_KIND='EXEXPORT'";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" || textBox4.Text == "" || textBox10.Text == "")
            {
                MessageBox.Show("請輸入期初日期");
                return;

            }

            GetMenu.InsertEXEXPORT("01財務損益表", textBox3.Text + "/" + textBox4.Text + "/" + textBox10.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"),"");
           // System.Diagnostics.Process.Start(OutPutFileS);

            MessageBox.Show("訊息已送出");
     

            //DELETEFILE2("損益表.xls");
            //SD();

            //string FileName = string.Empty;
            //string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            //FileName = lsAppDir + "\\Excel\\損益表.xls";

            ////Excel的樣版檔
            //string ExcelTemplate = FileName;

            ////輸出檔
            //string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
            //      Path.GetFileName(FileName);

            ////產生 Excel Report
            //ExcelReport.ExcelReportOutputANYA(dtc, ExcelTemplate, OutPutFile, "N");

            //string OutPutFile1 = lsAppDir + "\\Excel\\temp\\損益表.xls";
            ////輸出檔
            //string DD = lsAppDir + "\\Excel\\temp\\" +
            //      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
            //GetExcelProduct2(OutPutFile1, DD);
    
        }

        private void SD()
        {
            decimal A1 = 0;
            decimal A2 = 0;
            decimal A3 = 0;
            decimal A4 = 0;
            decimal A5 = 0;
            decimal A6 = 0;
            decimal A7 = 0;
            decimal A8 = 0;
            decimal A9 = 0;
            decimal A10 = 0;
            decimal A11 = 0;
            decimal A12 = 0;
            dtc = MakeTable();
            string a = textBox3.Text;
            string b = textBox4.Text;
            string s = textBox10.Text;
            System.Data.DataTable dt = Get1(a, b, s);
            string f1;
            string f2;
            string f3;
           
            decimal g1 = 0;
            decimal g2 = 0;
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtc.NewRow();
                string 科目代碼 = dt.Rows[i]["科目代碼"].ToString();
                string 科目名稱 = dt.Rows[i]["科目名稱"].ToString();
                string CATID = dt.Rows[i]["CATID"].ToString();
                string 本月 = dt.Rows[i]["本月"].ToString();
                string 本年 = dt.Rows[i]["本年"].ToString();
                string 備註 = dt.Rows[i]["備註"].ToString();
                string 迴轉 = dt.Rows[i]["迴轉"].ToString();
                f1 = 科目名稱;
                f2 = 本月;
                f3 = 本年;
            
                if (科目名稱 == "收入" || 科目名稱 == "銷貨收入淨額")
                {
                    System.Data.DataTable Fdt = Get2(a, b, "4");

                    if (Fdt.Rows.Count > 0)
                    {
                        f2 = Fdt.Rows[0][0].ToString();
                    }
                    else
                    {
                        f2 = "0";
                    }
                    System.Data.DataTable Fdt2 = Get2(s, b, "4");
                    if (Fdt2.Rows.Count > 0)
                    {
                        f3 = Fdt2.Rows[0][0].ToString();
                    }
                    else
                    {
                        f2 = "0";
                    }
                }
                if (科目名稱 == "銷貨成本"&& CATID =="6")
                {
                    System.Data.DataTable Fdt3 = Get4OUT(a, b, s);
                    if (Fdt3.Rows.Count > 0)
                    {
                        string G1 = Fdt3.Rows[0][0].ToString();
                        string G2 = Fdt3.Rows[0][1].ToString();
                        if (String.IsNullOrEmpty(G1))
                        {
                            G1 = "0";
                        }
                        if (String.IsNullOrEmpty(G2))
                        {
                            G2 = "0";
                        }

                        A1 = Convert.ToDecimal(G1);
                        A2 = Convert.ToDecimal(G2);
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();
                    }

                }
                if (科目名稱 == "銷貨毛利")
                {
                    System.Data.DataTable Fdt3 = Get3(a, b, s);
                    string G1 = Fdt3.Rows[0][0].ToString();
                    string G2 = Fdt3.Rows[0][1].ToString();
                    if (String.IsNullOrEmpty(G1))
                    {
                        G1 = "0";
                    }
                    if (String.IsNullOrEmpty(G2))
                    {
                        G2 = "0";
                    }

                    A1 = Convert.ToDecimal(G1);
                    A2 = Convert.ToDecimal(G2);
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();
                }
                if (科目名稱 == "管理及總務費用")
                {
                    System.Data.DataTable Fdt3 = Get4(a, b, s);
                    string G1 = Fdt3.Rows[0][0].ToString();
                    string G2 = Fdt3.Rows[0][1].ToString();
                    if (String.IsNullOrEmpty(G1))
                    {
                        G1 = "0";
                    }
                    if (String.IsNullOrEmpty(G2))
                    {
                        G2 = "0";
                    }

                    A3 = Convert.ToDecimal(G1);
                    A4 = Convert.ToDecimal(G2);
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();
                }
                if (科目名稱 == "營業淨利(淨損)")
                {
                    A5 = A1 + A3;
                    A6 = A2 + A4;
                    f2 = (A1 + A3).ToString();
                    f3 = (A2 + A4).ToString();
                }
                if (科目名稱 == "營業外收入及利益")
                {
                    System.Data.DataTable Fdt3 = Get5(a, b, s, "71", "74");
                    string G1 = Fdt3.Rows[0][0].ToString();
                    string G2 = Fdt3.Rows[0][1].ToString();
                    if (String.IsNullOrEmpty(G1))
                    {
                        G1 = "0";
                    }
                    if (String.IsNullOrEmpty(G2))
                    {
                        G2 = "0";
                    }

                    A7 = Convert.ToDecimal(G1);
                    A8 = Convert.ToDecimal(G2);
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();

                }

                if (科目名稱 == "營業外費用及損失")
                {
                    System.Data.DataTable Fdt3 = Get5(a, b, s, "75", "76");
                    string G1 = Fdt3.Rows[0][0].ToString();
                    string G2 = Fdt3.Rows[0][1].ToString();
                    if (String.IsNullOrEmpty(G1))
                    {
                        G1 = "0";
                    }
                    if (String.IsNullOrEmpty(G2))
                    {
                        G2 = "0";
                    }
                    A9 = Convert.ToDecimal(G1);
                    A10 = Convert.ToDecimal(G2);
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();

                }
                if (科目名稱 == "所得稅(費用)利益")
                {
                    if (f2 == "")
                    {
                        f2 = "0";
                    }
                    if (f3 == "")
                    {
                        f3 = "0";
                    }
                    A11 = Convert.ToDecimal(f2);
                    A12 = Convert.ToDecimal(f3);

                }

                if (迴轉 == "Y")
                {
                    int G1 = f2.IndexOf("-");
                    if (G1 != -1)
                    {
                        f2 = f2.Replace("-", "");
                    }
                    else
                    {
                        f2 = "-" + f2;
                    }

                    int G2 = f3.IndexOf("-");
                    if (G2 != -1)
                    {
                        f3 = f3.Replace("-", "");
                    }
                    else
                    {
                        f3 = "-" + f3;
                    }

                }

                if (科目名稱 == "稅前淨利(損)")
                {
                    f2 = (A5 + A7 + A9 ).ToString();
                    f3 = (A6 + A8 + A10 ).ToString();
                }

                if (科目名稱 == "本期淨利(淨損)")
                {
                    f2 = (A5 + A7 + A9 + A11).ToString();
                    f3 = (A6 + A8 + A10 + A12).ToString();
                }


                if (f2.Trim() == "0")
                {
                    f2 = "";
                }
                if (f3.Trim() == "0")
                {
                    f3 = "";
                }

                if (f2 == "")
                {
                    f2 = "-";
                }
                if (f3 == "")
                {
                    f3 = "-";
                }
                dr["科目名稱"] = f1;
                dr["本月"] = f2;
                dr["本年"] = f3;

                if (!String.IsNullOrEmpty(f2 + f3))
                {
                    if (f2 + f3 != "--")
                    {

                        dtc.Rows.Add(dr);

                    }
                }
            }

        }

        private void    SD2()
        {

            int MM = Convert.ToInt16(MONTH);

            for (int y = 1; y <= MM; y++)
            {
                decimal A1 = 0;
                decimal A2 = 0;
                decimal A3 = 0;
                decimal A4 = 0;
                decimal A5 = 0;
                decimal A6 = 0;
                decimal A7 = 0;
                decimal A8 = 0;
                decimal A9 = 0;
                decimal A10 = 0;
                decimal A11 = 0;
                decimal A12 = 0;
                dtc = MakeTableG();
                string DAY = "";
                if (y < 10)
                {
                    DAY = "0" + y.ToString();

                }
                else
                {
                    DAY = y.ToString();
                }

                string a = "";
                if (YYYY == "2014" && y==1)
                {
                    a = YYYY + DAY + "02";
                }
                else
                {
                    a = YYYY + DAY + "01";
                }


                string b = YYYY + DAY + "31";
                string s = "";
                if (YYYY == "2014")
                {
                    s = YYYY + "0102";
                }
                else
                {
                    s = YYYY + "0101";
                }
                System.Data.DataTable dt = Get1(a, b, s);
                string f1;
                string f2;
                string f3;
                decimal g1 = 0;
                decimal g2 = 0;
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtc.NewRow();
                    string 科目代碼 = dt.Rows[i]["科目代碼"].ToString();
                    string 科目名稱 = dt.Rows[i]["科目名稱"].ToString();
                    string CATID = dt.Rows[i]["CATID"].ToString();
                    string 本月 = dt.Rows[i]["本月"].ToString();
                    string 本年 = dt.Rows[i]["本年"].ToString();
                    string 備註 = dt.Rows[i]["備註"].ToString();
                    string 迴轉 = dt.Rows[i]["迴轉"].ToString();
                    f1 = 科目名稱;
                    f2 = 本月;
                    f3 = 本年;
                    if (科目名稱 == "收入" || 科目名稱 == "銷貨收入淨額")
                    {
                        System.Data.DataTable Fdt = Get2(a, b, "4");

                        f2 = Fdt.Rows[0][0].ToString();

                        System.Data.DataTable Fdt2 = Get2(s, b, "4");
                        f3 = Fdt2.Rows[0][0].ToString();
                    }

                    if (科目名稱 == "銷貨毛利")
                    {
                        System.Data.DataTable Fdt3 = Get3(a, b, s);
                        A1 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                        A2 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();
                    }
                    if (科目名稱 == "管理及總務費用")
                    {
                        System.Data.DataTable Fdt3 = Get4(a, b, s);
                        A3 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                        A4 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();
                    }
                    if (科目名稱 == "銷貨成本" && CATID =="6")
                    {
                        System.Data.DataTable Fdt3 = Get4OUT(a, b, s);
                        A1 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                        A2 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();
                    }
                    if (科目名稱 == "營業淨利(淨損)")
                    {
                        A5 = A1 + A3;
                        A6 = A2 + A4;
                        f2 = (A1 + A3).ToString();
                        f3 = (A2 + A4).ToString();
                    }
                    if (科目名稱 == "營業外收入及利益")
                    {
                        System.Data.DataTable Fdt3 = Get5(a, b, s, "71", "74");
                        A7 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                        A8 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();

                    }

                    if (科目名稱 == "營業外費用及損失")
                    {
                        System.Data.DataTable Fdt3 = Get5(a, b, s, "75", "76");
                        A9 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                        A10 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                        f2 = Fdt3.Rows[0][0].ToString();
                        f3 = Fdt3.Rows[0][1].ToString();

                    }
                    if (科目名稱 == "所得稅(費用)利益")
                    {
                        if (f2 == "")
                        {
                            f2 = "0";
                        }
                        if (f3 == "")
                        {
                            f3 = "0";
                        }
                        A11 = Convert.ToDecimal(f2);
                        A12 = Convert.ToDecimal(f3);

                    }

                    if (迴轉 == "Y")
                    {
                        int G1 = f2.IndexOf("-");
                        if (G1 != -1)
                        {
                            f2 = f2.Replace("-", "");
                        }
                        else
                        {
                            f2 = "-" + f2;
                        }

                        int G2 = f3.IndexOf("-");
                        if (G2 != -1)
                        {
                            f3 = f3.Replace("-", "");
                        }
                        else
                        {
                            f3 = "-" + f3;
                        }

                    }


                    if (科目名稱 == "稅前淨利(損)")
                    {
                        f2 = (A5 + A7 + A9).ToString();
                        f3 = (A6 + A8 + A10).ToString();
                    }


                    if (科目名稱 == "本期淨利(淨損)")
                    {
                        f2 = (A5 + A7 + A9 + A11).ToString();
                        f3 = (A6 + A8 + A10 + A12).ToString();
                    }


                    if (f2.Trim() == "0")
                    {
                        f2 = "";
                    }
                    if (f3.Trim() == "0")
                    {
                        f3 = "";
                    }

                    if (f2 == "")
                    {
                        f2 = "-";
                    }
                    if (f3 == "")
                    {
                        f3 = "-";
                    }
                    dr["科目名稱"] = f1;
                    dr["本月"] = f2;
                    dr["本年"] = f3;
                    dr["DATE"] = YYYY  + DAY;
                    if (!String.IsNullOrEmpty(f2 + f3))
                    {
                        if (f2 != "--" && f2 != "-" && f2 != "-0")
                        {

                            dtc.Rows.Add(dr);
                            AddG(i, f1, f2, dr["DATE"].ToString());
                        }
                    }
                }
            }

        }




        private void SD3()
        {
            decimal A1 = 0;
            decimal A2 = 0;
            decimal A3 = 0;
            decimal A4 = 0;
            decimal A5 = 0;
            decimal A6 = 0;
            decimal A7 = 0;
            decimal A8 = 0;
            decimal A9 = 0;
            decimal A10 = 0;
            decimal A11 = 0;
            decimal A12 = 0;
            dtc = MakeTable();


            string a = textBox11.Text + "01";
            string b = textBox11.Text + "31";
            string s = "";
            if (YYYY == "2014")
            {
                s = YYYY + "0102";
            }
            else
            {
                s = YYYY + "0101";
            }

           
            System.Data.DataTable dt = Get1(a, b, s);
            string f1;
            string f2;
            string f3;
            decimal g1 = 0;
            decimal g2 = 0;
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtc.NewRow();
                string 科目代碼 = dt.Rows[i]["科目代碼"].ToString();
                string 科目名稱 = dt.Rows[i]["科目名稱"].ToString();
                string CATID = dt.Rows[i]["CATID"].ToString();
                string 本月 = dt.Rows[i]["本月"].ToString();
                string 本年 = dt.Rows[i]["本年"].ToString();
                string 備註 = dt.Rows[i]["備註"].ToString();
                string 迴轉 = dt.Rows[i]["迴轉"].ToString();
                f1 = 科目名稱;
                f2 = 本月;
                f3 = 本年;
                if (科目名稱 == "收入" || 科目名稱 == "銷貨收入淨額")
                {
                    System.Data.DataTable Fdt = Get2(a, b, "4");

                    f2 = Fdt.Rows[0][0].ToString();

                    System.Data.DataTable Fdt2 = Get2(s, b, "4");
                    f3 = Fdt2.Rows[0][0].ToString();
                }
                if (科目名稱 == "銷貨成本" && CATID == "6")
                {
                    System.Data.DataTable Fdt3 = Get4OUT(a, b, s);
                    A1 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                    A2 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();
                }
                if (科目名稱 == "銷貨毛利")
                {
                    System.Data.DataTable Fdt3 = Get3(a, b, s);
                    A1 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                    A2 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();
                }
                if (科目名稱 == "管理及總務費用")
                {
                    System.Data.DataTable Fdt3 = Get4(a, b, s);
                    A3 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                    A4 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();
                }
                if (科目名稱 == "營業淨利(淨損)")
                {
                    A5 = A1 + A3;
                    A6 = A2 + A4;
                    f2 = (A1 + A3).ToString();
                    f3 = (A2 + A4).ToString();
                }
                if (科目名稱 == "營業外收入及利益")
                {
                    System.Data.DataTable Fdt3 = Get5(a, b, s, "71", "74");
                    A7 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                    A8 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();

                }

                if (科目名稱 == "營業外費用及損失")
                {
                    System.Data.DataTable Fdt3 = Get5(a, b, s, "75", "76");
                    A9 = Convert.ToDecimal(Fdt3.Rows[0][0].ToString());
                    A10 = Convert.ToDecimal(Fdt3.Rows[0][1].ToString());
                    f2 = Fdt3.Rows[0][0].ToString();
                    f3 = Fdt3.Rows[0][1].ToString();

                }
                if (科目名稱 == "所得稅(費用)利益")
                {
                    if (f2 == "")
                    {
                        f2 = "0";
                    }
                    if (f3 == "")
                    {
                        f3 = "0";
                    }
                    A11 = Convert.ToDecimal(f2);
                    A12 = Convert.ToDecimal(f3);

                }

                if (迴轉 == "Y")
                {
                    int G1 = f2.IndexOf("-");
                    if (G1 != -1)
                    {
                        f2 = f2.Replace("-", "");
                    }
                    else
                    {
                        f2 = "-" + f2;
                    }

                    int G2 = f3.IndexOf("-");
                    if (G2 != -1)
                    {
                        f3 = f3.Replace("-", "");
                    }
                    else
                    {
                        f3 = "-" + f3;
                    }

                }


                if (科目名稱 == "稅前淨利(損)")
                {
                    f2 = (A5 + A7 + A9).ToString();
                    f3 = (A6 + A8 + A10).ToString();
                }


                if (科目名稱 == "本期淨利(淨損)")
                {
                    f2 = (A5 + A7 + A9 + A11).ToString();
                    f3 = (A6 + A8 + A10 + A12).ToString();
                }


                if (f2.Trim() == "0")
                {
                    f2 = "";
                }
                if (f3.Trim() == "0")
                {
                    f3 = "";
                }

                if (f2 == "")
                {
                    f2 = "-";
                }
                if (f3 == "")
                {
                    f3 = "-";
                }
                dr["科目名稱"] = f1;
                dr["本月"] = f2;
                dr["本年"] = f3;

                if (!String.IsNullOrEmpty(f2 + f3))
                {
                    if (f2 + f3 != "--")
                    {

                        dtc.Rows.Add(dr);
                     
                        AddG(i, f1, f3, YYYY);
                    }
                }
            }

        }
        private System.Data.DataTable Get1(string a, string b, string c)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T1.ACCTCODE 科目代碼,T0.CATID,T0.FRGNNAME 備註,CASE T0.LEVELS WHEN 4 THEN T1.ACCTCODE+' - '+T4.ACCTNAME ELSE T0.[NAME] END 科目名稱");
            sb.Append(" ,CASE T0.LEVELS WHEN 4 THEN T2.AMOUNT ELSE T5.AMOUNT END 本月");
            sb.Append(" ,CASE T0.LEVELS WHEN 4 THEN T3.AMOUNT ELSE T6.AMOUNT END 本年,T0.REVERSAL 迴轉");
            sb.Append(" FROM OFRC T0");
            sb.Append(" LEFT JOIN FRC1 T1 ON(T0.CATID=T1.CATID AND T0.TEMPLATEID=T1.TEMPLATEID)");
            sb.Append(" LEFT JOIN OACT T4 ON(T1.ACCTCODE=T4.ACCTCODE )");
            sb.Append(" LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE ");
            sb.Append("  Convert(varchar(8),REFDATE,112) BETWEEN @aa AND @bb GROUP BY ACCOUNT) T2 ON(T1.ACCTCODE=T2.ACCOUNT)");
            sb.Append(" LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE ");
            sb.Append("  Convert(varchar(8),REFDATE,112) BETWEEN @cc AND @bb GROUP BY ACCOUNT) T3 ON(T1.ACCTCODE=T3.ACCOUNT)");
            sb.Append(" LEFT JOIN (SELECT SUM(AMOUNT) AMOUNT,FATHERNUM FROM OFRC T0");
            sb.Append(" LEFT JOIN FRC1 T1 ON(T0.CATID=T1.CATID AND T0.TEMPLATEID=T1.TEMPLATEID)");
            sb.Append(" LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE ");
            sb.Append("  Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb GROUP BY ACCOUNT) T2 ON(T1.ACCTCODE=T2.ACCOUNT)");
            sb.Append(" WHERE T0.TEMPLATEID=23 GROUP BY FATHERNUM) T5 ON (T0.CATID=T5.FATHERNUM)");
            sb.Append(" LEFT JOIN (SELECT SUM(AMOUNT) AMOUNT,FATHERNUM FROM OFRC T0");
            sb.Append(" LEFT JOIN FRC1 T1 ON(T0.CATID=T1.CATID AND T0.TEMPLATEID=T1.TEMPLATEID)");
            sb.Append(" LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE ");
            sb.Append("  Convert(varchar(8),REFDATE,112) BETWEEN @cc AND @bb GROUP BY ACCOUNT) T2 ON(T1.ACCTCODE=T2.ACCOUNT)");
            sb.Append(" WHERE T0.TEMPLATEID=23 GROUP BY FATHERNUM) T6 ON (T0.CATID=T6.FATHERNUM)");
            sb.Append(" WHERE  T0.TEMPLATEID=23 AND [NAME] NOT IN ('費用') ORDER BY T0.VISORDER,T1.ACCTCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable GetINCOME(string DD)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ACCOUNT 科目名稱,[" + DD + "] '2012',[" + DD + "01] '201201',[" + DD + "02] '201202',[" + DD + "03] '201203',[" + DD + "04] '201204',[" + DD + "05] '201205',[" + DD + "06] '201206',[" + DD + "07] '201207',[" + DD + "08] '201208',[" + DD + "09] '201209',[" + DD + "10] '201210',[" + DD + "11] '201211',[" + DD + "12] '201212'");
            sb.Append(" from (");
            sb.Append(" SELECT IDINCOME,ACCOUNT,CAST(INYEAR AS FLOAT) INYEAR,INDATE FROM Account_Income  where INYEAR <> '-' ");
            sb.Append(" )");
            sb.Append(" T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(INYEAR)");
            sb.Append(" FOR INDATE IN");
            sb.Append(" ( [" + DD + "],[" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" order by IDINCOME");
        
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable GetINCOME2(string DD)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" Select ACCOUNT 科目名稱,[" + DD + "] '2012',[" + DD + "01] '201201',[" + DD + "02] '201202',[" + DD + "03] '201203',[" + DD + "Q1] '2012Q1',[" + DD + "04] '201204',[" + DD + "05] '201205',[" + DD + "06] '201206',[" + DD + "Q2] '2012Q2',[" + DD + "07] '201207',[" + DD + "08] '201208',[" + DD + "09] '201209',[" + DD + "Q3] '2012Q3',[" + DD + "10] '201210',[" + DD + "11] '201211',[" + DD + "12] '201212',[" + DD + "Q4] '2012Q4'");
            sb.Append(" from (");
            sb.Append("              SELECT IDINCOME,ACCOUNT,CAST(INYEAR AS   FLOAT) INYEAR,INDATE FROM Account_Income where INYEAR <> '-'");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'Q1' FROM Account_Income");
            sb.Append("              WHERE SUBSTRING((INDATE),5,2) BETWEEN '01' AND '03'");
            sb.Append("              GROUP BY IDINCOME");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'Q2' FROM Account_Income");
            sb.Append("              WHERE SUBSTRING((INDATE),5,2) BETWEEN '04' AND '06'");
            sb.Append("              GROUP BY IDINCOME");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'Q3' FROM Account_Income");
            sb.Append("              WHERE SUBSTRING((INDATE),5,2) BETWEEN '07' AND '09'");
            sb.Append("              GROUP BY IDINCOME");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'Q4' FROM Account_Income");
            sb.Append("              WHERE SUBSTRING((INDATE),5,2) BETWEEN '10' AND '12'");
            sb.Append("              GROUP BY IDINCOME");
            sb.Append(" )");
            sb.Append(" T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(INYEAR)");
            sb.Append(" FOR INDATE IN");
            sb.Append(" ( [" + DD + "],[" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "Q1],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "Q2],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "Q3],[" + DD + "10],[" + DD + "11],[" + DD + "12],[" + DD + "Q4] )");
            sb.Append(" ) AS pvt");
            sb.Append(" order by IDINCOME");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", DD));
            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable GetINCOME3(string a)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT INDATE FROM ( SELECT IDINCOME,ACCOUNT,CAST(INYEAR AS   FLOAT) INYEAR,INDATE FROM Account_Income");
            sb.Append("                           UNION ALL");
            sb.Append("                           SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'031' FROM Account_Income");
            sb.Append("                           WHERE SUBSTRING((INDATE),5,2) BETWEEN '01' AND '03'");
            sb.Append("                           GROUP BY IDINCOME");
            sb.Append("                           UNION ALL");
            sb.Append("                           SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'061' FROM Account_Income");
            sb.Append("                           WHERE SUBSTRING((INDATE),5,2) BETWEEN '04' AND '06'");
            sb.Append("                           GROUP BY IDINCOME");
            sb.Append("                           UNION ALL");
            sb.Append("                           SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'091' FROM Account_Income");
            sb.Append("                           WHERE SUBSTRING((INDATE),5,2) BETWEEN '07' AND '09'");
            sb.Append("                           GROUP BY IDINCOME");
            sb.Append("                           UNION ALL");
            sb.Append("                           SELECT IDINCOME,MAX(ACCOUNT) 科目,SUM(CAST(INYEAR AS   FLOAT)) INYEAR,@aa+'121' FROM Account_Income");
            sb.Append("                           WHERE SUBSTRING((INDATE),5,2) BETWEEN '10' AND '12'");
            sb.Append("                           GROUP BY IDINCOME ) AS A WHERE INDATE <> @aa ORDER BY INDATE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            SqlDataAdapter da = new SqlDataAdapter(command);

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
        private System.Data.DataTable Get1H(string a, string b)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT T1.ACCTCODE 科目代碼,T0.CATID");
            sb.Append("             ,CASE T0.LEVELS WHEN 4 THEN T2.AMOUNT ELSE T5.AMOUNT END 每月 ");
            sb.Append("             FROM OFRC T0");
            sb.Append("             LEFT JOIN FRC1 T1 ON(T0.CATID=T1.CATID AND T0.TEMPLATEID=T1.TEMPLATEID)");
            sb.Append("             LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE");
            sb.Append("              Convert(varchar(8),REFDATE,112) BETWEEN @bb+@aa+'01' AND @bb+@aa+'31' GROUP BY ACCOUNT) T2 ON(T1.ACCTCODE=T2.ACCOUNT)");
            sb.Append("             LEFT JOIN (SELECT SUM(AMOUNT) AMOUNT,FATHERNUM FROM OFRC T0");
            sb.Append("             LEFT JOIN FRC1 T1 ON(T0.CATID=T1.CATID AND T0.TEMPLATEID=T1.TEMPLATEID)");
            sb.Append("             LEFT JOIN (SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,ACCOUNT FROM JDT1 WHERE");
            sb.Append("              Convert(varchar(8),REFDATE,112)  BETWEEN @bb+@aa+'01' AND @bb+@aa+'31' GROUP BY ACCOUNT) T2 ON(T1.ACCTCODE=T2.ACCOUNT)");
            sb.Append("             WHERE T0.TEMPLATEID=23 GROUP BY FATHERNUM) T5 ON (T0.CATID=T5.FATHERNUM)");
            sb.Append("             WHERE  T0.TEMPLATEID=23 AND [NAME] NOT IN ('費用') ORDER BY T0.VISORDER,T1.ACCTCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable GetCOMPANY()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("        SELECT COMPNYNAME FROM OADM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

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
        private void DELETEFILE2(string aa)
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp\\";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                FileInfo filess = new FileInfo(file);
                string fd = filess.Name.ToString();
                if (fd == aa)
                {
                    File.Delete(file);
                }
            }
        }


        private System.Data.DataTable Get2(string a, string b, string c)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CAST(SUM(CREDIT)-SUM(DEBIT) AS float ) AMOUNT,SUBSTRING(ACCOUNT,1,1) ACCOUNT FROM JDT1 WHERE ");
            sb.Append("  Convert(varchar(8),REFDATE,112) BETWEEN @aa AND @bb and SUBSTRING(ACCOUNT,1,1)=@cc GROUP BY SUBSTRING(ACCOUNT,1,1)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable Get3(string a, string b, string c)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (SUM(CREDIT)-SUM(DEBIT))+ISNULL((SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=5");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb ");
            sb.Append(" GROUP BY SUBSTRING(ACCOUNT,1,1)),0) 銷貨毛利A,ISNULL((SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=4");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @cc AND @bb ");
            sb.Append(" )+(SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=5");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @cc AND @bb ");
            sb.Append(" ),0) 銷貨毛利B");
            sb.Append("  FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=4");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            SqlDataAdapter da = new SqlDataAdapter(command);

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

        private System.Data.DataTable Get4(string a, string b, string c)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (SUM(CREDIT)-SUM(DEBIT)) 管理A,(SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1) =6");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @cc AND @bb");
            sb.Append(" ) 管理B FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=6");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            SqlDataAdapter da = new SqlDataAdapter(command);

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
        private System.Data.DataTable Get4OUT(string a, string b, string c)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (SUM(CREDIT)-SUM(DEBIT)) 管理A,(SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1) =5");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @cc AND @bb");
            sb.Append(" ) 管理B FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,1)=5");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            SqlDataAdapter da = new SqlDataAdapter(command);

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
        private System.Data.DataTable Get5(string a, string b, string c, string d, string e)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (SUM(CREDIT)-SUM(DEBIT)) 管理A,(SELECT SUM(CREDIT)-SUM(DEBIT) AMOUNT FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,2) between @d and @e");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @cc AND @bb");
            sb.Append(" ) 管理B FROM JDT1 ");
            sb.Append(" WHERE SUBSTRING(ACCOUNT,1,2) between @d and @e");
            sb.Append(" AND Convert(varchar(8),REFDATE,112)  BETWEEN @aa AND @bb");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", a));
            command.Parameters.Add(new SqlParameter("@bb", b));
            command.Parameters.Add(new SqlParameter("@cc", c));
            command.Parameters.Add(new SqlParameter("@d", d));
            command.Parameters.Add(new SqlParameter("@e", e));
            SqlDataAdapter da = new SqlDataAdapter(command);

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


        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("科目名稱", typeof(string));
            dt.Columns.Add("本月", typeof(string));
            dt.Columns.Add("本年", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("科目名稱", typeof(string));
            dt.Columns.Add("本月", typeof(string));
            dt.Columns.Add("本年", typeof(string));
            dt.Columns.Add("DATE", typeof(string));
            return dt;
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (textBox11.Text == "")
            {
                MessageBox.Show("請輸入年月");
                return;

            }
            GetMenu.InsertEXEXPORT("05損益表每月", textBox11.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), "");
            MessageBox.Show("訊息已送出");
            //YYYY = textBox11.Text.Substring(0, 4);
            //MONTH = textBox11.Text.Substring(4, 2);
            //TRUNG();
            //DELETEFILE2("損益表每月.xls");
            //SD3();
            //SD2();


            //System.Data.DataTable T1 = GetINCOME(YYYY);

            //string FileName = string.Empty;
            //string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            //FileName = lsAppDir + "\\Excel\\ACC\\損益表每月.xls";

            ////Excel的樣版檔
            //string ExcelTemplate = FileName;

            ////輸出檔
            //string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
            //      Path.GetFileName(FileName);

            ////產生 Excel Report
            //ExcelReport.ExcelReportOutputANYA(T1, ExcelTemplate, OutPutFile, "N");

            //string OutPutFile1 = lsAppDir + "\\Excel\\temp\\損益表每月.xls";
            ////輸出檔
            //string DD = lsAppDir + "\\Excel\\temp\\" +
            //      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
            //GetExcelINCOME(OutPutFile1, DD);
        }

        public void AddG(int IDINCOME, string ACCOUNT, string INYEAR, string INDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("  Insert into Account_Income(IDINCOME,ACCOUNT,INYEAR,INDATE) values(@IDINCOME,@ACCOUNT,@INYEAR,@INDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@IDINCOME", IDINCOME));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@INYEAR", INYEAR));
            command.Parameters.Add(new SqlParameter("@INDATE", INDATE));

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

        public void TRUNG()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" TRUNCATE TABLE Account_Income ", connection);
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

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox11.Text == "")
            {
                MessageBox.Show("請輸入年月");
                return;

            }
            GetMenu.InsertEXEXPORT("06損益表每季", textBox11.Text, fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), "");
            MessageBox.Show("訊息已送出");
           // YYYY = textBox11.Text.Substring(0, 4);
           // MONTH = textBox11.Text.Substring(4, 2);
           // TRUNG();
           // DELETEFILE2("損益表每季.xls");
           // SD3();
           // SD2();


           // System.Data.DataTable T1 = GetINCOME2(YYYY);

           // string FileName = string.Empty;
           // string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

           // FileName = lsAppDir + "\\Excel\\ACC\\損益表每季.xls";

           // //Excel的樣版檔
           // string ExcelTemplate = FileName;

           // //輸出檔
           // string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
           //       Path.GetFileName(FileName);

           // //產生 Excel Report
           // ExcelReport.ExcelReportOutputANYA2(T1, ExcelTemplate, OutPutFile);

           // string OutPutFile1 = lsAppDir + "\\Excel\\temp\\損益表每季.xls";
           // //輸出檔
           // string DD = lsAppDir + "\\Excel\\temp\\" +
           //       DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
           //GetExcelINCOME2(OutPutFile1, DD);
       }


        private void DELETEFILE()
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
 
    
    }
}
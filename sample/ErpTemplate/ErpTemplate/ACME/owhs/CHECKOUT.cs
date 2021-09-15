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
using System.Globalization;
namespace ACME
{
    public partial class CHECKOUT : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn2 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=sa;Password=m@ggie";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCneep = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        private System.Data.DataTable TempDt;
        private System.Data.DataTable TempDt2;
        private string FileName;
        string dd = "";
        private int[] iCount;
        public CHECKOUT()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

       

                if (comboBox1.Text == "新得利")
                {
                    WriteExcelProduct(FileName, 1, 4, 5, "TW017");
                }
                if (comboBox1.Text == "聯揚倉")
                {
                    if (checkBox1.Checked)
                    {
                        WriteExcelProduct(FileName, 1, 2, 4, "TW012");
                    }
                    else
                    {
                        WriteExcelProduct(FileName, 3, 7, 15, "TW012");
                    }
                }
            }
        }
        private void WriteExcelProduct(string ExcelFile, int a, int b, int c, string WHSCODE)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
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

            if (!checkBox5.Checked)
            {
                if (iRowCnt > 2000)
                {
                    iRowCnt = 2000;
                }

            }

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;
                string QTY;
                int Qty = 0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, a]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, b]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    System.Data.DataTable T1 = GETOITW(SERIAL_NO, WHSCODE);
                    System.Data.DataTable T2 = GETOITM(SERIAL_NO, WHSCODE);
                    if (T2.Rows.Count == 0 && QTY != "")
                    {

                        if (SERIAL_NO != "產品編號")
                        {
                            int G1 = SERIAL_NO.IndexOf("總");
                            if (G1 == -1)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, a]);
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
                    }
                    else
                    {
                        if (T1.Rows.Count > 0)
                        {
                            string G1 = T1.Rows[0][0].ToString();
                            if (G1 != QTY)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, b]);
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c]);
                                range.Value2 = G1;
                            }
                        }
                    }


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

        private void WriteExcelProduct2(string ExcelFile, int SERNO, int PARTNO3, int INV, int c, string WHSCODE, int GRA,string WAREHOUSE)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
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

            if (!checkBox5.Checked)
            {
                if (iRowCnt > 2000)
                {
                    iRowCnt = 2000;
                }

            }

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;
                string PARTNO;
                string GRADE;
                string INV2;
                string DOCDATE;
                string NAME = "";
                string sTemp = string.Empty;
                int u = 0;
                int v = 0;

                if (checkBox2.Checked)
                {
                    for (int b = 1; b <= 10; b++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                        range.Select();
                        NAME = range.Text.ToString().Trim();


                        if (NAME == "產品編號")
                        {
                            SERNO = b;
                        }
                        if (NAME == "料號")
                        {
                            PARTNO3 = b;
                        }
                        if (NAME == "等級")
                        {
                            GRA = b;
                        }
                    }
                }
                else
                {
                    GRA = 1;
                }
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, SERNO]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GRA]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, PARTNO3]);
                    range.Select();
                    PARTNO = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, INV]);
                    range.Select();
                    INV2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    if (SERIAL_NO != "")
                    {
                        System.Data.DataTable T2 = GETOITM2(SERIAL_NO);
                        if (T2.Rows.Count == 0 && PARTNO != "")
                        {
                            if (SERIAL_NO.IndexOf("產品") == -1)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, SERNO]);
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
                        else
                        {

                            if (T2.Rows.Count > 0)
                            {
                                string PARTNO2 = T2.Rows[0][0].ToString().Trim();
                                int T1 = PARTNO.IndexOf(PARTNO2);

                                if (PARTNO == "9AQ23807-011")
                                {
                                    T1 = 1;
                                }
                                if (T1 == -1)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, PARTNO3]);
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                }



                                if (checkBox2.Checked)
                                {
                                    string GRADES = T2.Rows[0][1].ToString().Trim().ToUpper();
                                    int J1 = GRADE.IndexOf(GRADES);
                                    int J2 = GRADES.IndexOf(GRADE);
                                    if (J1 == -1 && J2 == -1)
                                    {
                                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GRA]);
                                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BlueViolet);
                                    }

                                }
                            }
                        }

                    }

                    if (!checkBox2.Checked)
                    {
                        if (!String.IsNullOrEmpty(INV2))
                        {
                            System.Data.DataTable P1 = GETINV(INV2);
                            if (P1.Rows.Count > 0)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c]);
                                range.Select();
                                range.Value2 = P1.Rows[0][0].ToString();

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c + 1]);
                                range.Select();
                                range.Value2 = P1.Rows[0][1].ToString();
                            }
                            else
                            {
                                if (WAREHOUSE == "新得利")
                                {
                                    DateTime DT = Convert.ToDateTime(DOCDATE);
                                    System.Data.DataTable P2 = GETINV2(DT);
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c]);
                                    range.Select();
                                    range.Value2 = P2.Rows[0][0].ToString();

                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c + 1]);
                                    range.Select();
                                    range.Value2 = P2.Rows[0][1].ToString();
                                    
                                }
                            
                            
                            
                            
                            }
                        }
                    }
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

        private void WriteExcelProduct3(string ExcelFile, int ITEMCODE, int QTY, int c, int MODEL, int GRADE, int VER, int INV, string WHSCODE)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (!checkBox5.Checked)
            {
                if (iRowCnt > 2000)
                {
                    iRowCnt = 2000;
                }

            }

            if (iRowCnt > 700 && fmLogin.LoginID.ToString().ToUpper()=="TONYWU")
            {
                iRowCnt = Convert.ToInt32(textBox1.Text);
            }

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;
                string TQTY;
                string TMODEL;
                string TGRADE;
                string TVER;
                string INV2;

                string CONN = "";
                string WHCHCODE = "";
                if (comboBox4.Text == "CHOICE")
                {
                    CONN = strCn20;
                }
                else if (comboBox4.Text == "INFINITE")
                {
                    CONN = strCn21;
                }
                if (comboBox4.Text != "進金生")
                {
                    System.Data.DataTable GR = GETCHOW(CONN);
                    if (GR.Rows.Count > 0)
                    {
                        WHCHCODE = GR.Rows[0][0].ToString();
                    }
                    else
                    {
                        MessageBox.Show("正航無此倉庫");
                        return;
                    }
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, ITEMCODE]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                    range.Select();
                    TQTY = range.Text.ToString().Trim().Replace(",", "").Replace(".00", "");
                     int n;
                     if (int.TryParse(TQTY, out n))
                     {

                         if (TQTY != "")
                         {
                             int Ts = SERIAL_NO.IndexOf("产品编号");
                             int Ts1 = TQTY.ToUpper().IndexOf("QTY");
                             if (Ts1 == -1)
                             {
                                 if (SERIAL_NO != "產品編號")
                                 {
                                     if (!String.IsNullOrEmpty(SERIAL_NO))
                                     {
                                         if (Ts == -1)
                                         {
                                             AddTEMP(SERIAL_NO, TQTY);
                                         }
                                     }
                                 }
                             }
                         }
                     }
                }

                int J = 0;

                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();

            
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, ITEMCODE]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();
            
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                    range.Select();
                    TQTY = range.Text.ToString().Trim().Replace(",", "");
              
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, INV]);
                    range.Select();
                    INV2 = range.Text.ToString().Trim();

            
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, MODEL]);
                    range.Select();
                    TMODEL = range.Text.ToString().Trim();
         
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GRADE]);
                    range.Select();
                    TGRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, VER]);
                    range.Select();
                    TVER = range.Text.ToString().Trim();

                    System.Data.DataTable T1 = null;
                    System.Data.DataTable T2 = null;
   

                    if (comboBox4.Text == "進金生")
                    {
                        T1 = GETOITW(SERIAL_NO, WHSCODE);
                        T2 = GETOITM(SERIAL_NO, WHSCODE);
                    }
                    else
                    {
               
                            T1 = GETCHOW2(CONN, SERIAL_NO, WHCHCODE);
                            T2 = GETOITMCH(CONN, SERIAL_NO, WHCHCODE);
                  
                    }

                    System.Data.DataTable T3 = GETOITM3(SERIAL_NO);
                    System.Data.DataTable T4 = GETOITM3M(SERIAL_NO);
                    if (!String.IsNullOrEmpty(INV2))
                    {
                        System.Data.DataTable P1 = GETINV(INV2);
                        if (P1.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c + 1]);
                            range.Select();
                            range.Value2 = P1.Rows[0][0].ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c + 2]);
                            range.Select();
                            range.Value2 = P1.Rows[0][1].ToString();
                        }
                    }

                    int H1 = SERIAL_NO.IndexOf("請參考工單");
                    int Ts1 = TQTY.ToUpper().IndexOf("QTY");
                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {
                        if (Ts1 == -1)
                        {
                            if (H1 == -1)
                            {

                                if (T2.Rows.Count == 0 && TQTY != "")
                                {
                                    int Ts = SERIAL_NO.IndexOf("产品编号");
                                    if (SERIAL_NO != "產品編號")
                                    {
                                        if (Ts == -1)
                                        {
                                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, ITEMCODE]);
                                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                        }
                                    }
                                }
                                else
                                {
                                    if (T1.Rows.Count > 0)
                                    {
                                        if (T4.Rows.Count > 0)
                                        {
                                            string G1 = T1.Rows[0][0].ToString();
                                            string G2 = T4.Rows[0][0].ToString();
                                            if (G1 != G2)
                                            {


                                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, QTY]);
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

                                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, c]);
                                                range.Value2 = G1;

                                                System.Data.DataTable S1 = AddTEMPS(SERIAL_NO);

                                                if (S1.Rows.Count == 0)
                                                {
                                                    J++;
                                                    string ONHAND = (Convert.ToInt32(G1) - Convert.ToInt32(G2)).ToString();

                                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2 + J, ITEMCODE]);
                                                    range.Select();
                                                    range.Value2 = SERIAL_NO;

                                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2 + J, ITEMCODE + 1]);
                                                    range.Select();
                                                    range.Value2 = ONHAND;

                                          //          AddTEMPU(SERIAL_NO);
                                                }
                                            }
                                        }
                                    }
                                    if (T3.Rows.Count > 0)
                                    {
                                    string SMODEL = T3.Rows[0]["MODEL"].ToString();
                                    int k1 = TMODEL.IndexOf(SMODEL);
                                    int k2 = SMODEL.IndexOf(TMODEL);
                                    if (k1 == -1)
                                    {
                                        if (k2 == -1)
                                        {
                                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, MODEL]);
                                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                        }

                                    }

                             
                                        string f1 = SERIAL_NO.Substring(0, 1).ToUpper();

                                        string SGRADE = T3.Rows[0]["GRADE"].ToString().Trim();
                                        if (SGRADE != TGRADE)
                                        {
                                            if (f1 == "K")
                                            {
                                                if (TGRADE.ToUpper() != "N/A")
                                                {
                                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GRADE]);
                                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                                }
                                            }
                                            else
                                            {

                                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, GRADE]);
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                            }

                                        }
                                        string SVER = T3.Rows[0]["VER"].ToString();
                                        if (SVER != TVER)
                                        {
                                            if (f1 == "K")
                                            {
                                                if (TVER.ToUpper() != "N/A")
                                                {
                                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, VER]);
                                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                                }
                                            }
                                            else
                                            {

                                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, VER]);
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
                                            }


                                        }

                                    }

                                }

                            }
                        }
                    }
                    else
                    {
                        if (TQTY != "")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, ITEMCODE]);
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    
                    }
                }


                System.Data.DataTable GG1 = null;
                string VV = "";

                StringBuilder sb = new StringBuilder();

                System.Data.DataTable GETEMP = GETEMP1();

                if (GETEMP.Rows.Count > 0)
                {
                    for (int i = 0; i <= GETEMP.Rows.Count - 1; i++)
                    {

                        sb.Append("'" + GETEMP.Rows[i][0].ToString() + "',");

                    }
                    sb.Remove(sb.Length - 1, 1);
                }
                if (comboBox4.Text == "進金生")
                {
                    GG1=GETOITWS(WHSCODE);
                    VV = "外倉庫存沒顯示但SAP系統有";
                }
                else
                {
                    GG1 = GETOITWSCH(CONN, WHCHCODE, sb.ToString());
                    VV = "外倉庫存沒顯示但正航系統有";
                }
                if (GG1.Rows.Count > 0 || J !=0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 1, ITEMCODE]);
                    range.Select();
                    range.Value2 = VV;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2, ITEMCODE]);
                    range.Select();
                    range.Value2 = "料號";
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2, ITEMCODE + 1]);
                    range.Select();
                    range.Value2 = "數量";
                }
                if (GG1.Rows.Count > 0)
                {


                    for (int i = 1; i <= GG1.Rows.Count ; i++)
                    {

                        DataRow drw3 = GG1.Rows[i - 1];
                        string ITEM = drw3["ITEMCODE"].ToString();
                        string ONHAND = drw3["ONHAND"].ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2 + i + J, ITEMCODE]);
                        range.Select();
                        range.Value2 = ITEM;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRowCnt + 2 + i + J, ITEMCODE + 1]);
                        range.Select();
                        range.Value2 = ONHAND;
                    }

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

        private void WNANCY(string ExcelFile)
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string ITEMCODE;



                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    if (ITEMCODE != "")
                    {
                        System.Data.DataTable G1 = GNANCY(ITEMCODE);
                        if (G1.Rows.Count > 0)
                        {
                            System.Data.DataTable G2 = GNANCY2(ITEMCODE);
                            // sb.Append(" SELECT TOP 1 Convert(varchar(10),DDATE,111) 日期,DOCENTRY 單號,GQTY 數量,GTOTAL 金額 FROM Account_Temp612020  ");
                            string 日期 = G1.Rows[0]["日期"].ToString();
                            string 單號 = G1.Rows[0]["單號"].ToString();
                            string 數量 = G1.Rows[0]["數量"].ToString();
                            string 金額 = G1.Rows[0]["金額"].ToString();

                            string 金額2 = G2.Rows[0]["金額"].ToString();
                            string 單價 = G2.Rows[0]["單價"].ToString();
                            string 數量2 = G2.Rows[0]["數量"].ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 6]);
                            range.Select();
                            range.Value2 = 數量2;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 7]);
                            range.Select();
                            range.Value2 = 金額2;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 8]);
                            range.Select();
                            range.Value2 = 單價;


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 9]);
                            range.Select();
                            range.Value2 = 日期;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 10]);
                            range.Select();
                            range.Value2 = 單號;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 11]);
                            range.Select();
                            range.Value2 = 數量;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[iRecord, 12]);
                            range.Select();
                            range.Value2 = 金額;

                        }

                    }



                    //     AddTEMPG1(ACC,CARDCODE);


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
        private void APPLE1(string ExcelFile)
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string WARR;

                string WARR2;
             
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    WARR = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    WARR2 = range.Text.ToString().Trim();

                    if (WARR != "")
                    {
                        UPAPPLE(WARR2, WARR);
                    }


        
               //     AddTEMPG1(ACC,CARDCODE);
                            
                    
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
        private void WriteExcelProduct4T(string ExcelFile)
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string ITEMCODE;
                string BRAND;
                string ITEMNAME;
                string MODEL;
                string DESC;
                string U_SIZE;
                string U_TMODEL;
                string U_GRADE;
                string U_VERSION;
                string U_PARTNO;
                string FRGNNAME;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();




                    if (ITEMCODE != "")
                    {

                        if (ITEMCODE != "項目號碼")
                        {
                            //AddTEMPG1(ITEMCODE, BRAND);
                        }
                    }
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
        private void WriteExcelProduct20151207(string ExcelFile)
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string G1;
                string G2;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                 
                    G1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);

                    G2 = range.Text.ToString().Trim();

                    G2 = G2.Replace(",", "");

                    if (G2 != "")
                    {

                        Add20151207(G1,G2);
                        
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




            }



        }
        private void WriteCHI(string ExcelFile)
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string M1;
                string M2;
                string M3;
                string M4;
                string M5;
                string M6;
                string M7;
      
              
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    M1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    M2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    M3 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    M4 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    M5 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    M6 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    M7 = range.Text.ToString().Trim();



                    AddM(M1, M2, M3, M4, M5, M6, M7,"");
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
        private void WriteExcelAP(string ExcelFile)
        {
          //  AddAP
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string ITEMNAME = "";
                string ITEMCODE = "";

                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    ITEMNAME = range.Text.ToString().Trim();

                               range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                               range.Select();
                               ITEMCODE = range.Text.ToString().Trim();

                               if (!String.IsNullOrEmpty(ITEMCODE))
                                {
                                    UPPCARD(ITEMCODE, ITEMNAME);
                                }
                   // AddAP(DOCENTRY);
                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                try
                {
                   // excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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


       //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelGBPICK3(string ExcelFile)
        {
            //  AddAP
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string SA;
                string CARDCODE;
    
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {
                                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    CARDCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    SA = range.Text.ToString().Trim();

                    System.Data.DataTable FS1 = GETES1(SA);
                    if (FS1.Rows.Count > 0)
                    {
                        string EMPID = FS1.Rows[0][0].ToString();
   
                            UPOCRD(EMPID, CARDCODE);
                        
                    }


               //     AddAPP("聯倉", "15T", LOCATION, AMT);
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

        private void WriteExcelAP2(string ExcelFile)
        {
            //  AddAP
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string BU;
                string CARDCODE;
                string CARDNAME;
                string TRANSID;
                string ORDR;
                string OINV;
                string JRNLMEMO;
                decimal AMT;
                decimal CURRENCY;
                decimal USD;
                string SA;
                string SALES;
                string SHIPDATE;
                string MEMO;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    BU = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CARDCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    TRANSID = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    ORDR = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    OINV = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    JRNLMEMO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    string A1 = range.Text.ToString();
                    string F1 = A1.Replace(",", "").Replace("(", "").Replace(")", "");
                    int I1 = A1.IndexOf("(");
                    if (String.IsNullOrEmpty(F1))
                    {
                        AMT = 0;
                    }
                    else
                    {
                        if (I1 != -1)
                        {
                            AMT = Convert.ToDecimal(F1) * -1;
                        }
                        else
                        {
                            AMT = Convert.ToDecimal(F1);
                        }
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    string F2 = range.Text.ToString().Replace(",", "");
                                     if (String.IsNullOrEmpty(F2))
                                     {
                                         CURRENCY = 0;
                                     }
                                     else
                                     {
                                         CURRENCY = Convert.ToDecimal(F2);
                                     }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    string A3 = range.Text.ToString();
                    string F3 = A3.Replace(",", "").Replace("(", "").Replace(")", "");
                    int I3 = A3.IndexOf("(");
                    if (String.IsNullOrEmpty(F3))
                    {
                        USD = 0;
                    }
                    else
                    {
                        if (I3 != -1)
                        {
                            USD = Convert.ToDecimal(F3) * -1;
                        }
                        else
                        {
                            USD = Convert.ToDecimal(F3);
                        }
                    }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    SA = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    SALES = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    SHIPDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();
                    if (!String.IsNullOrEmpty(BU))
                    {
                        AddAP2(BU,CARDCODE, CARDNAME, TRANSID, ORDR, OINV, JRNLMEMO, AMT, CURRENCY, USD, SA, SALES, SHIPDATE, MEMO);
                    }
                }




            }
            finally
            {

           //     string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           //DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


           //     try
           //     {
           //         excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
           //     }
           //     catch
           //     {
           //     }
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


              //  System.Diagnostics.Process.Start(NewFileName);


            }



        }
        public void TRUNTABLE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table WH_TEMP1 ", connection);
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

        public void AddTEMP(string ITEMCODE, string QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_TEMP1(ITEMCODE,QTY) values(@ITEMCODE,@QTY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));

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
        public void UPORDR(string U_BASE_DOC, string DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE ACMESQL02.DBO.RDR1 SET U_BASE_DOC=@U_BASE_DOC  WHERE DOCENTRY=@DOCENTRY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_BASE_DOC", U_BASE_DOC));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        public void UPOCRD(string DfTcnician, string CARDCODE)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE ACMESQL02.DBO.OCRD SET DfTcnician=@DfTcnician  WHERE CARDCODE=@CARDCODE UPDATE ACMESQL98.DBO.OCRD SET DfTcnician=@DfTcnician  WHERE CARDCODE=@CARDCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@DfTcnician", DfTcnician));

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
        public void UPPCARD(string ITEMCODE, string U_ITEMNAME)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET U_ITEMNAME=@U_ITEMNAME  WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@U_ITEMNAME", U_ITEMNAME));

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


        public System.Data.DataTable AddTEMPS(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT FLAG FROM  WH_TEMP1  WHERE ITEMCODE=@ITEMCODE AND FLAG='Y'  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public void AddAP(string DOCTYPE2)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_STOCK(DOCTYPE2) values(@DOCTYPE2)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCTYPE2", DOCTYPE2));


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
        public void Add20151207(string G1, string G2)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AB(G1,G2) values(@G1,@G2)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@G1", G1));
            command.Parameters.Add(new SqlParameter("@G2", G2));

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

        public void AddGB_PACK3(string PRODNAME, string SuggestPrice, string UDef1, string ProdID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand("UPDATE comproduct  SET PRODNAME=@PRODNAME,SuggestPrice=@SuggestPrice,UDef1=@UDef1 WHERE ProdID=@ProdID", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PRODNAME", PRODNAME));
            command.Parameters.Add(new SqlParameter("@SuggestPrice", SuggestPrice));
            command.Parameters.Add(new SqlParameter("@UDef1", UDef1));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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

        public void AddAPP(string CARNAME, string WEIGHT, string LOCATION, string AMT)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_CARFEE(CARNAME,WEIGHT,LOCATION,AMT) values(@CARNAME,@WEIGHT,@LOCATION,@AMT)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARNAME", CARNAME));
            command.Parameters.Add(new SqlParameter("@WEIGHT", WEIGHT));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));

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

        public void EE(string PARAM_NO, string PARAM_DESC)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_PARAMS(PARAM_KIND,PARAM_NO,PARAM_DESC) values(@PARAM_KIND,@PARAM_NO,@PARAM_DESC)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_KIND", "ASHARON"));
            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));
            command.Parameters.Add(new SqlParameter("@PARAM_DESC", PARAM_DESC));



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

        public void EE2()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE RMA_PARAMS WHERE PARAM_KIND='ESCOT'", connection);
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
        public void AddAP2(string BU, string CARDCODE, string CARDNAME, string TRANSID, string ORDR, string OINV, string JRNLMEMO, decimal AMT, decimal CURRENCY, decimal USD, string SA, string SALES, string SHIPDATE, string MEMO)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SATT5(BU,CARDCODE,CARDNAME,TRANSID,ORDR,OINV,JRNLMEMO,AMT,CURRENCY,USD,SA,SALES,SHIPDATE,MEMO) values(@BU,@CARDCODE,@CARDNAME,@TRANSID,@ORDR,@OINV,@JRNLMEMO,@AMT,@CURRENCY,@USD,@SA,@SALES,@SHIPDATE,@MEMO)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));
            command.Parameters.Add(new SqlParameter("@ORDR", ORDR));
            command.Parameters.Add(new SqlParameter("@OINV", OINV));
            command.Parameters.Add(new SqlParameter("@JRNLMEMO", JRNLMEMO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@CURRENCY", CURRENCY));
            command.Parameters.Add(new SqlParameter("@USD", USD));
            command.Parameters.Add(new SqlParameter("@SA", SA));

            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@SHIPDATE", SHIPDATE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
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

        public void UPAPPLE(string U_ACME_WARRANTY, string U_ACME_WARRANTY2)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE ODLN SET U_ACME_WARRANTY=@U_ACME_WARRANTY WHERE U_ACME_WARRANTY=@U_ACME_WARRANTY2 ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_WARRANTY", U_ACME_WARRANTY));

            command.Parameters.Add(new SqlParameter("@U_ACME_WARRANTY2", U_ACME_WARRANTY2));
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

        public void AddTEMPG1(string DebPayAcct, string CardCode)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OCRD SET DebPayAcct=@DebPayAcct WHERE CardCode=@CardCode", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@DebPayAcct", DebPayAcct));
            
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


        public void UPDATECHI(string ProjectID, string FundBillNo)
        {
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand("UPDATE comBillAccounts SET ProjectID=@ProjectID WHERE FundBillNo=@FundBillNo", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));
            command.Parameters.Add(new SqlParameter("@FundBillNo", FundBillNo));


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
        public void AddTEMP21(string ITEMCODE, string U_BRAND)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET U_BRAND=@U_BRAND WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@U_BRAND", U_BRAND));
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
        public void AddTEMP22(string ITEMCODE, string U_ITEMNAME)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET U_ITEMNAME=@U_ITEMNAME WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@U_ITEMNAME", U_ITEMNAME));

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
        public void AddTEMP23(string ITEMCODE,string U_MODEL)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET U_MODEL=@U_MODEL WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@U_MODEL", U_MODEL));
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
        public void AddTEMP24(string ITEMCODE, string U_LOCATION)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET U_LOCATION=@U_LOCATION WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@U_LOCATION", U_LOCATION));
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
        public void AddTEMP25(string ITEMCODE, string ITEMNAME)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OITM SET ITEMNAME=@ITEMNAME WHERE ITEMCODE=@ITEMCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
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
        public System.Data.DataTable GETOITMCH(string conn, string ITEMCODE, string WHSCODE)
        {
            SqlConnection connection = new SqlConnection(conn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  WareID  FROM comWareAmount  WHERE WareID=@WHSCODE AND ProdID=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public System.Data.DataTable GETOITM(string ITEMCODE, string WHSCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  WHSCODE  FROM OITW  WHERE WHSCODE=@WHSCODE AND ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public System.Data.DataTable GETOITM3(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_TMODEL MODEL,U_GRADE GRADE,U_VERSION VER FROM OITM  WHERE  ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public System.Data.DataTable GETOITM3T(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_TMODEL MODEL,U_GRADE GRADE,U_VERSION VER FROM OITM  WHERE  ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public System.Data.DataTable GETOITM3M(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(QTY AS INT)) QTY FROM WH_TEMP1 WHERE  ITEMCODE=@ITEMCODE GROUP BY ITEMCODE");
       

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public System.Data.DataTable GETOITM2(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_PARTNO 料號,U_GRADE GRADE FROM OITM T0 WHERE ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        public System.Data.DataTable GETOITW(string ITEMCODE, string WHSCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(CAST(ONHAND AS INT))    FROM OITW  WHERE WHSCODE=@WHSCODE AND ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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


        public System.Data.DataTable GETORDR(string ITEMCODE, string WHSCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(CAST(ONHAND AS INT))    FROM OITW  WHERE WHSCODE=@WHSCODE AND ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public System.Data.DataTable GETCHOW(string conn)
        {

            SqlConnection connection = new SqlConnection(conn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT WareHouseID  FROM comWareHouse WHERE ENGNAME='V' AND SHORTNAME=@SHORTNAME");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHORTNAME", comboBox3.Text));
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
        public System.Data.DataTable GETCHOW2(string conn, string ITEMCODE, string WHSCODE)
        {

            SqlConnection connection = new SqlConnection(conn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(Quantity-LendQuan AS INT))  FROM comWareAmount WHERE WareID=@WHSCODE AND ProdID =@ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public System.Data.DataTable GETOITWS( string WHSCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT ITEMCODE,CAST(ONHAND AS INT) ONHAND FROM OITW WHERE WHSCODE=@WHSCODE AND ONHAND > 0 ");
           sb.Append(" AND ITEMCODE NOT IN (SELECT ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.WH_TEMP1)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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

        public System.Data.DataTable GETOITWSCH(string conn, string WHSCODE, string ProdID)
        {

            SqlConnection connection = new SqlConnection(conn);

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT ProdID ITEMCODE,CAST(Quantity AS INT) ONHAND FROM comWareAmount WHERE WareID=@WHSCODE AND (Quantity-LendQuan) > 0 ");
            sb.Append(" AND ProdID NOT in (" + ProdID + ")");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public System.Data.DataTable GETEMP1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ITEMCODE  FROM WH_TEMP1 ");
        
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
        public System.Data.DataTable GNANCY(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT TOP 1 Convert(varchar(10),DDATE,111) 日期,DOCENTRY 單號,GQTY 數量,GTOTAL 金額 FROM Account_Temp612020  ");
            sb.Append(" WHERE ITEMCODE=@ITEMCODE AND DOCENTRY>2153        ORDER BY DDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        public System.Data.DataTable GNANCY2(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT SUM(GQTY) 數量,SUM(GTOTAL) 金額,AVG(GTOTAL/GQTY) 單價 FROM Account_Temp612020  WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GETINV(string INV)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,U_ACME_INVOICE)-1911), 3) +'.'+");
            sb.Append(" RIGHT('0'+CONVERT(VARCHAR(2), DATEPART(month,U_ACME_INVOICE)), 2) +'.'+");
            sb.Append(" RIGHT('0'+CONVERT(VARCHAR(2), DATEPART(day,U_ACME_INVOICE)), 2) INV日期,datediff(d,U_ACME_INVOICE ,getdate()) 天數 FROM OPDN WHERE U_ACME_INV LIKE '%" + INV + "%'");
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
        public System.Data.DataTable GETINV2(DateTime  INV)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,@INV)-1911), 3) +'.'+ ");
            sb.Append("               RIGHT('0'+CONVERT(VARCHAR(2), DATEPART(month,@INV)), 2) +'.'+ ");
            sb.Append("               RIGHT('0'+CONVERT(VARCHAR(2), DATEPART(day,@INV)), 2) INV日期,");
            sb.Append(" datediff(d,@INV ,getdate()) 天數 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@INV", INV));
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
        public System.Data.DataTable GETES1(string EMP)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            //   sb.Append(" SELECT *   FROM AP_INVOICEIN WHERE CARTON=@CARTON AND ITEMCODE IN ('O270HTN02.12002','O270HTN02.52002')");
            sb.Append(" SELECT EMPID FROM OHEM WHERE  LASTNAME+FIRSTNAME=@EMP");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@EMP", EMP));
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

        public System.Data.DataTable GETES2(string U_ACME_INV)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select Convert(varchar(10),T0.DOCDATE,111) 出貨日期,T6.DOCENTRY 訂單單號");
            sb.Append("      from ACMESQL02.DBO.ODLN T0 ");
            sb.Append("         LEFT JOIN ACMESQL02.DBO.DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("          left join acmesql02.dbo.RDR1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum ) ");
            sb.Append("          left join acmesql02.dbo.ORDR t6 on (T5.DOCENTRY=T6.DOCENTRY ) ");
            sb.Append("         LEFT JOIN acmesql02.dbo.OITM T10 ON T1.ITEMCODE = T10.ITEMCODE   ");
            sb.Append("         where T0.U_ACME_INV like '%" + U_ACME_INV + "%'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
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
        public System.Data.DataTable GETWH_TEMP20151207(string CARTON)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


         //   sb.Append(" SELECT *   FROM AP_INVOICEIN WHERE CARTON=@CARTON AND ITEMCODE IN ('O270HTN02.12002','O270HTN02.52002')");
            sb.Append(" SELECT CARTON   FROM AP_INVOICEIN WHERE CARTON=@CARTON AND ITEMCODE IN ('O270HTN02.12002','O270HTN02.52002')");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
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

        public System.Data.DataTable GETWH_TEMP201512072(string PIC)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT *   FROM AP_INVOICEIN WHERE PIC=@PIC AND ITEMCODE='O315DVR01.53001'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@PIC", PIC));
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
        public static bool IsNumber(string strNumber)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"^\d+(\.)?\d*$");
            return r.IsMatch(strNumber);
        }
        private void CHECKOUT_Load(object sender, EventArgs e)
        {
            comboBox4.Text = "進金生";
            comboBox3.Text = "蘇州偉創";

            if (fmLogin.LoginID.ToString().ToUpper() == "TONYWU" || fmLogin.LoginID.ToString().ToUpper() == "LLEYTONCHEN")
            {
                label2.Visible = true;
                textBox1.Visible = true;

            }
            if (globals.GroupID.ToString().Trim() != "EEP" )
            {
                button4.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button7.Visible = false;
                button9.Visible = false;
                button8.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                if (comboBox2.Text == "新得利")
                {
                    WriteExcelProduct2(FileName, 5, 6, 9, 17, "TW017", 0, comboBox2.Text);
                }
                if (comboBox2.Text == "聯揚倉")
                {
                   // WriteExcelProduct2(FileName, 2, 6, "TW012");
                    WriteExcelProduct2(FileName, 2, 4, 3, 10, "TW012", 0, comboBox2.Text);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (IsNumber(textBox1.Text) == false)
            {
                MessageBox.Show("請輸入數字");
                return;
            }
            //深圳宏高
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                TRUNTABLE();
                if (comboBox3.Text == "香港宇思")
                {
                    //int ITEMCODE, int QTY, int c, int MODEL, int GRADE, int VER, int INV, 
                    WriteExcelProduct3(FileName, 8, 11, 20, 8, 12, 9, 5, "HK007");
                }
                if (comboBox3.Text == "蘇州偉創")
                {
                    WriteExcelProduct3(FileName, 8, 11, 15, 5, 9, 6, 4, "CN006");
                }
                if (comboBox3.Text == "廈門宏高")
                {
                    WriteExcelProduct3(FileName, 8, 11,16, 5, 9, 6, 4, "CN004");
                }
                if (comboBox3.Text == "深圳宏高")
                {
                    WriteExcelProduct3(FileName, 8, 11, 20, 5, 9, 6, 4, "CN05");
                }
                if (comboBox3.Text == "蘇州宏高")
                {
                    WriteExcelProduct3(FileName, 8, 11, 15, 5, 9, 6,4, "CN001");
                }
                if (comboBox3.Text == "香港宏高")
                {
                    WriteExcelProduct3(FileName, 8, 11, 14, 5, 9, 6, 3, "HK002");
                }
                if (comboBox3.Text == "聯揚倉")
                {
                    WriteExcelProduct3(FileName, 2, 5, 0, 5, 9, 6, 4, "TW012");
                }
                if (comboBox3.Text == "巨航機保" )
                {
                    WriteExcelProduct3(FileName, 8, 11, 14, 5, 9, 6, 4, "CN009");
                }
                if ( comboBox3.Text == "巨航坪山")
                {
                    WriteExcelProduct3(FileName, 8, 11, 14, 5, 9, 6, 4, "CN010");

                }
                if (comboBox3.Text == "武漢巨航")
                {
                    WriteExcelProduct3(FileName, 8, 11, 14, 5, 9, 6, 4, "CN011");

                }
                if (comboBox3.Text == "香港巨航")
                {
                    WriteExcelProduct3(FileName, 8, 11, 14, 5, 9, 6, 4, "HK006");

                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                APPLE1(FileName);
                   // WNANCY(FileName);
                
            }
        }

        private void linkLabel4_Click(object sender, EventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\比對庫存儲位.pdf";
            System.Diagnostics.Process.Start(aa);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP(FileName);

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP2(FileName);

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                WriteCHI(FileName);

            }
        }
        public void AddM(string M1, string M2, string M3, string M4, string M5, string M6, string M7, string M8)
        {
            SqlConnection Connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPORM(M1,M2,M3,M4,M5,M6,M7,M8) values(@M1,@M2,@M3,@M4,@M5,@M6,@M7,@M8)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@M1", M1));
            command.Parameters.Add(new SqlParameter("@M2", M2));
            command.Parameters.Add(new SqlParameter("@M3", M3));
            command.Parameters.Add(new SqlParameter("@M4", M4));
            command.Parameters.Add(new SqlParameter("@M5", M5));
            command.Parameters.Add(new SqlParameter("@M6", M6));
            command.Parameters.Add(new SqlParameter("@M7", M7));
            command.Parameters.Add(new SqlParameter("@M8", M8));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        private void button8_Click(object sender, EventArgs e)
        {
        
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                WriteExcelProduct20151207(FileName);


            }
        }

        private void WriteExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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

  
            if (iRowCnt > 5000)
            {
                MessageBox.Show("超過五千筆無法上傳");
                return;
            }

            progressBar1.Maximum = iRowCnt;

            
            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SR;
                string SH;
                int Qty = 0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    SR = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    SH = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(SR))
                    {
                        System.Data.DataTable T1 = GETWH_TEMP20151207(SR);
                        if (T1.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                            range.Value2 = "V";

                        }
                        System.Data.DataTable T2 = GETWH_TEMP201512072(SR);
                        if (T2.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                            range.Value2 = "V";

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
                System.Diagnostics.Process.Start(NewFileName);

            }


   
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                WriteExcelProduct(FileName);


            }
        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            if (comboBox3.Text == "香港宇思")
            {
                label1.Text = "產品編號: " + util.EXCEL(8) + " 數量: " + util.EXCEL(11) + " MODEL : " + util.EXCEL(8) + " 等級 : " + util.EXCEL(12) + " 版本 : " + util.EXCEL(9) + " 友達發票號碼 : " + util.EXCEL(5);
            }
            if (comboBox3.Text == "蘇州偉創" || comboBox3.Text == "廈門宏高" || comboBox3.Text == "深圳宏高" || comboBox3.Text == "蘇州宏高" || comboBox3.Text == "巨航機保" || comboBox3.Text == "巨航坪山" || comboBox3.Text == "香港巨航" || comboBox3.Text == "武漢巨航")
            {
                label1.Text = "產品編號: " + util.EXCEL(8) + " 數量: " + util.EXCEL(11) + " MODEL : " + util.EXCEL(5) + " 等級 : " + util.EXCEL(9) + " 版本 : " + util.EXCEL(6) + " 友達發票號碼 : " + util.EXCEL(4);
            }

            if (comboBox3.Text == "香港宏高")
            {
                label1.Text = "產品編號: " + util.EXCEL(8) + " 數量: " + util.EXCEL(11) + " MODEL : " + util.EXCEL(5) + " 等級 : " + util.EXCEL(9) + " 版本 : " + util.EXCEL(6) + " 友達發票號碼 : " + util.EXCEL(3);
            }

            if (comboBox3.Text == "聯揚倉")
            {
                label1.Text = "產品編號: " + util.EXCEL(2) + " 數量: " + util.EXCEL(5) + " MODEL : " + util.EXCEL(5) + " 等級 : " + util.EXCEL(9) + " 版本 : " + util.EXCEL(6) + " 友達發票號碼 : " + util.EXCEL(4);

            }
    
        }

       

    
    }
}

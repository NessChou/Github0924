using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Collections;
namespace ACME
{
    public partial class SHIPMULTI : Form
    {
        string USER = fmLogin.LoginID.ToString();
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public SHIPMULTI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
           
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al.Add(listBox2.Items[i].ToString());
                }

                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }
                sb.Remove(sb.Length - 1, 1);
            }
            catch { }
    
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            decimal AMT = 0;
            string AMT2 = "";
                System.Data.DataTable J1 = GetINVOAMOUNT(sb.ToString());
                if (J1.Rows.Count > 0)
                {
                    AMT = Convert.ToDecimal(J1.Rows[0][0].ToString());
                }
                int G = Convert.ToInt32(AMT.ToString("###0"));
                double AMT1 = Convert.ToDouble(AMT);
                if (G != 0)
                {
                    AMT2 = "SAY TOTAL : US DOLLARS " + new Class1().NumberToString(AMT1);
                }
                else
                {
                    AMT2 = "";
                }

                System.Data.DataTable H1 = GetOrderData2(sb.ToString(), AMT2);
                if (H1.Rows.Count > 0)
                {
                    if (globals.DBNAME == "禾中")
                    {
                        FileName = lsAppDir + "\\Excel\\AT\\INVODRSACMEM.xls";
                        GetExcelProduct2(FileName, H1, sb.ToString(), "Y");
                    }
                    else
                    {
                        if (USER.ToUpper() == "JOYCHEN")
                        {
                            FileName = lsAppDir + "\\Excel\\INVODRSACMEMJ.xls";
                            GetExcelProduct2(FileName, H1, sb.ToString(), "N");
                        }
                        else
                        {
                            FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACMEM.xls";
                            GetExcelProduct2(FileName, H1, sb.ToString(), "Y");
                        }
                    }

                         //System.Data.DataTable O2 = GetSHIPOHEM(fmLogin.LoginID.ToString());
                         //if (O2.Rows.Count > 0)
                         //{
                         //    FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACMEM.xls";
                         //    GetExcelProduct2(FileName, H1, sb.ToString(), "Y");
                         //}
                         //else
                         //{
                         //    GetExcelProduct2(FileName, H1, sb.ToString(), "N");
                         //}
                }
                else
                {
                    MessageBox.Show("沒有資料");
                }

        }
        private System.Data.DataTable GetSHIPOHEM(string USER)
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT HOMETEL FROM OHEM WHERE WORKCOUNTR='CN' AND HOMETEL=@HOMETEL AND ISNULL(TERMDATE,'') =''   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", USER));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
        private void GetExcelProduct2(string ExcelFile, System.Data.DataTable dt, string SB, string FLAG)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false  ;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            //excelSheet.Name = textBox7.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                string B2 = "//acmew08r2ap//table//SIGN//USER//";
                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                if (FLAG == "Y")
                {

                    excelSheet.Shapes.AddPicture(B2 + fmLogin.LoginID.ToString().Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoTrue, 410, 682, 200, 80);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue, dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 23;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue, dt);


                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }
              


                }



                //增加另一talbe處理

                System.Data.DataTable dtmark = Getmark(SB);
                if (dtmark.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }

                }



                //增加另一talbe處理


                string ID1 = "";
                int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
                for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 12]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    ID1 = sTemp;

                    if (sTemp == "Y")
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow+1, 1]);
            
                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);

                    }
                }

                System.Data.DataTable dt4 = GetWHITEM2(SB);
                if (dt4.Rows.Count > 0)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 26, 1]);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);
                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 27 + i, 2]);
                        range.Value2 = dt4.Rows[i][0].ToString();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    }
                }

                System.Data.DataTable dtmark2 = Getmark2(SB);
                if (dtmark2.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark2.Rows.Count - 1; a1Row++)
                    {
                        int F1 = 13;
                        //if (FLAG == "Y")
                        //{
                        //    F1 = 14;
                        //}
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[F1 + a1Row, 8]);

                        string FieldName = "shippingcode";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark2.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;


                    }

                }

                System.Data.DataTable dtmark3 = Getmark3(SB);
                if (dtmark3.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark3.Rows.Count - 1; a1Row++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8 + a1Row, 8]);

                        string FieldName = "INVOICENO";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark3.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;


                    }

                }

                //if (FLAG == "Y")
                //{
                    System.Data.DataTable dtmark4 = Getmark4(SB);
                    if (dtmark4.Rows.Count != 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, 8]);

                        StringBuilder sb = new StringBuilder();

                        for (int i = 0; i <= dtmark4.Rows.Count - 1; i++)
                        {

                            DataRow d = dtmark4.Rows[i];


                            sb.Append(d["DOCENTRY"].ToString() + "/");


                        }

                        sb.Remove(sb.Length - 1, 1);
                        range.Value2 = sb.ToString();
                    }


                    System.Data.DataTable dtmark5 = Getmark5(SB);
                    if (dtmark5.Rows.Count != 0)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[12, 8]);

                        StringBuilder sb = new StringBuilder();

                        for (int i = 0; i <= dtmark5.Rows.Count - 1; i++)
                        {
                            DataRow d = dtmark5.Rows[i];

                            sb.Append(d["INVOICENO"].ToString() + "/");

                        }

                        sb.Remove(sb.Length - 1, 1);
                        range.Value2 = sb.ToString();
                    }

                //}



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 12]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlDown);
            }
            finally
            {

                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);

                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                System.Diagnostics.Process.Start(NewFileName);

                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);

            }
        }

        private void GetExcelProduct22(string ExcelFile, System.Data.DataTable dt, string SB,string FLAG)
        {
            string flag = "Y";
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
            //excelSheet.Name = textBox7.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                string B2 = "//acmew08r2ap//table//SIGN//USER//";
                if (FLAG == "Y")
                {
                    excelSheet.Shapes.AddPicture(B2 + fmLogin.LoginID.ToString().Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
    Microsoft.Office.Core.MsoTriState.msoTrue, 350, 640, 200, 80);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue, dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 10;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue, dt);


                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }



                }



        
                //增加另一talbe處理
                if (FLAG == "N")
                {
                    System.Data.DataTable dtmark = Getmark(SB);
                    if (dtmark.Rows.Count != 0)
                    {
                        for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 6]);
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();
                            string FieldName = "mark";

                            FieldValue1 = "";
                            FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                            range.Value2 = FieldValue1;

                            DetailRow1++;
                        }

                    }
                }

                //if (dOCTYPETextBox.Text == "銷售")
                //{
                //    if (boardCountNoTextBox.Text == "三角" || boardCountNoTextBox.Text == "出口")
                //    {
                System.Data.DataTable dt4 = GetWHITEMLOC(SB);
                if (dt4.Rows.Count > 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 30, 1]);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {
                        string LOC = dt4.Rows[i][0].ToString();
                        System.Data.DataTable dt42 = GetWHITEMLOC2(SB, LOC);
                        if (dt42.Rows.Count > 0)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29 + i, 3]);
                            range.Value2 = "MADE IN " + dt42.Rows[0][0].ToString();

                            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                        }
                    }
                    //if (dt4.Rows.Count == 1)
                    //{
                
                    //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29, 3]);
                    //    range.Value2 = "MADE IN " + dt4.Rows[0][0].ToString();

                    //    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //}
                    //else
                    //{
                    //    //for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    //    //{
                    //    //    string LOC = dt4.Rows[i][0].ToString();
                    //    //    System.Data.DataTable dt42 = GetWHITEMLOC2(SB, LOC);
                    //    //    if (dt42.Rows.Count > 0)
                    //    //    {
                    //    //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29 + i, 3]);
                    //    //        range.Value2 = "MADE IN " + dt4.Rows[0][0].ToString();

                    //    //        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    //    //    }
                    //    //}
                    //}
                }
                            //System.Data.DataTable dt4 = GetWHITEM(SB);
                            //if (dt4.Rows.Count > 0)
                            //{

                            //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 30, 1]);
                            //    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);
                            //    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                            //    {
                            //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29 + i, 3]);
                            //        range.Value2 = dt4.Rows[i][0].ToString();

                            //        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                            //    }
                            //}

                System.Data.DataTable dtF = GetCTN(SB);
                if (dtF.Rows.Count > 0)
                {
                    for (int i = 0; i <= dtF.Rows.Count - 1; i++)
                    {
                        string PACKAGENO = dtF.Rows[i][0].ToString();
                        int SEQNO = Convert.ToInt16(dtF.Rows[i][1]);
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[SEQNO + 26 + i, 1]);
                        // range.Select();
                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[SEQNO + 26 + i, 3]);
                        range.Select();
                        range.Value2 = "(" + PACKAGENO + "PCS/ CTN)";
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.VerticalAlignment = XlVAlign.xlVAlignBottom;
                        range.Font.Bold = true;
                        range.Font.Size = 8;
                    }
                }

                        
                //    }
                //}


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 8]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlDown);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 8]);
                range.Select();
                range.EntireColumn.Delete(XlDirection.xlDown);
            }
            finally
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);
             //   lsAppDir + "\\Excel\\AT\\PACK2M.xls";
                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                System.Diagnostics.Process.Start(NewFileName);

                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);

            }
        }

       
        public System.Data.DataTable GetWHITEMLOC(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

        
            sb.Append("   select distinct LOCATION from PackingListD where shippingcode IN (" + SHIPPINGCODE + "  ) AND ISNULL(LOCATION,'') <> '' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetWHITEMLOC2(string SHIPPINGCODE, string LOCATION)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'ITEM'+CAST(MIN(cast(SEQ as int)) AS VARCHAR)+'~'+ CAST(MAX(cast(SEQ as int)) AS VARCHAR)+') MADE IN '+UPPER(MAX(LOCATION))");
            sb.Append("                                            FROM [PackingListM] as a   ");
            sb.Append("                                           left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)   ");
            sb.Append("                             LEFT JOIN (select shippingcode,PLNo,ID, ");
            sb.Append("               cast( RANK() OVER (ORDER BY AMTF,CAST(seqno AS INT)) as varchar) SEQ  from PackingListD  ");
            sb.Append("               where ISNULL(PACKMARK,'') = 'True'  ");
            sb.Append("                             and  shippingcode IN (" + SHIPPINGCODE + "  ) ) d  ");
            sb.Append("                             on (a.shippingcode=d.shippingcode and a.PLNo=d.PLNo  and b.ID=d.ID)   ");
            sb.Append("                             where  A.shippingcode IN (" + SHIPPINGCODE + "  )  AND ISNULL(LOCATION,'')=@LOCATION ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetWHITEM2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT 'ITEM'+CASE WHEN MIN(CAST(SEQ AS INT))=MAX(CAST(SEQ AS INT)) THEN CAST(MAX(CAST(SEQ AS INT)) AS VARCHAR) ELSE  ");
            sb.Append("                 CAST(MIN(CAST(SEQ AS INT)) AS VARCHAR)+'~'+CAST(MAX(CAST(SEQ AS INT)) AS VARCHAR)  END  + ')MADE IN '+LOCATION  PLATENO FROM InvoiceM A");
            sb.Append(" 				                            left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)  ");
            sb.Append(" 			 LEFT JOIN (select shippingcode,InvoiceNo,InvoiceNo_seq,docentry,cast(  RANK() OVER (ORDER BY shippingcode,seqno) as varchar) SEQ  from InvoiceD where");
            sb.Append("                 SHIPPINGCODE  IN (" + SHIPPINGCODE + "  )) d ");
            sb.Append("               on (a.shippingcode=d.shippingcode and a.InvoiceNo=d.InvoiceNo and a.InvoiceNo_seq=d.InvoiceNo_seq and b.docentry=d.docentry)  ");
            sb.Append("                    WHERE A.SHIPPINGCODE  IN (" + SHIPPINGCODE + "  )  AND ISNULL(SEQ,'') <> ''  AND ISNULL(LOCATION,'') <> ''  ");
            sb.Append("                 GROUP BY B.LOCATION ORDER BY MIN(CAST(SEQ AS INT)) ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetCTN(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                           SELECT  PALQTY PACKAGENO,");
            sb.Append("   SEQ-1 SEQNO   ");
            sb.Append(" FROM [PackingListM] as a   ");
            sb.Append("                                           left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)   ");
            sb.Append("                                           left join shipping_main as c on (a.shippingcode=c.shippingcode)   ");
            sb.Append("                             LEFT JOIN (select shippingcode,PLNo,ID, ");
            sb.Append("               cast( RANK() OVER (ORDER BY AMTF,CAST(seqno AS INT)) as varchar) SEQ  from PackingListD  ");
            sb.Append("                             where  shippingcode IN (" + SHIPPINGCODE + " )) d  ");
            sb.Append("                             on (a.shippingcode=d.shippingcode and a.PLNo=d.PLNo  and b.ID=d.ID)   ");
            sb.Append("                LEFT JOIN (select MAX(SEQNO) SEQNO,SHIPPINGCODE,'Y' TT from PackingListD where   shippingcode IN ( " + SHIPPINGCODE + " )   ");
            sb.Append("                             GROUP BY shippingcode) E ON (a.shippingcode=E.shippingcode AND B.SEQNO=E.SEQNO)  ");
            sb.Append("                             where  A.shippingcode IN ( " + SHIPPINGCODE + "  ) AND   ISNULL(PALQTY,'') <> ''     ORDER BY AMTF,cast(b.seqno as int)      ");

          
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        private System.Data.DataTable Getmark(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select mark from mark where shippingcode IN (" + AA + "  ) order by shippingcode,cast(Seq as int)  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        private System.Data.DataTable Getmark2(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select shippingcode from SHIPPING_MAIN where shippingcode IN (" + AA + "  )  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        private System.Data.DataTable Getmark3(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT INVOICENO+'-'+INVOICENO_SEQ INVOICENO FROM INVOICEM where shippingcode IN (" + AA + "  )  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable Getmark4(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT DOCENTRY FROM SHIPPING_ITEM WHERE SHIPPINGCODE IN (" + AA + "  )   ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable Getmark5(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT PONO INVOICENO FROM INVOICEM where shippingcode IN (" + AA + "  )  AND  ISNULL(PONO,'') <>''  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetOrderData3(string AA, string TOTAL, string QTY, string GROSS, string NET, string cc)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                 SELECT a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期  ");
            sb.Append("                             ,a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]  ");
            sb.Append("                             ,a.[Net],a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo]  ");
            sb.Append("                             ,@QTY as '總數',@GROSS as '螺絲',@NET as '耐特',@TOTAL as '欄位統計',@cc as cc, ");
            sb.Append("                   CASE WHEN ISNULL(PACKMARK,'') <> 'True' THEN '' ELSE SEQ+')' END+CASE ISNULL(TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods]  END DescGoods  ");
            sb.Append("                             ,b.[Quantity] as Quantity ,b.[Net] as Ne ,cast(b.[Gross] as varchar) as Go ,b.[MeasurmentCM],TT,A.SHIPPINGCODE SHIP FROM [PackingListM] as a  ");
            sb.Append("                             left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)  ");
            sb.Append("                             left join shipping_main as c on (a.shippingcode=c.shippingcode)  ");
            sb.Append("               LEFT JOIN (select shippingcode,PLNo,ID,");
            sb.Append(" cast( RANK() OVER (ORDER BY AMTF,CAST(seqno AS INT)) as varchar) SEQ  from PackingListD ");
            sb.Append(" where ISNULL(PACKMARK,'') = 'True' ");
            sb.Append("               and  shippingcode IN (" + AA + " )) d ");
            sb.Append("               on (a.shippingcode=d.shippingcode and a.PLNo=d.PLNo  and b.ID=d.ID)  ");
            sb.Append("  LEFT JOIN (select MAX(SEQNO) SEQNO,SHIPPINGCODE,'Y' TT from PackingListD where   shippingcode IN (" + AA + "   )  ");
            sb.Append("               GROUP BY shippingcode) E ON (a.shippingcode=E.shippingcode AND B.SEQNO=E.SEQNO) ");
            sb.Append("               where  A.shippingcode IN (" + AA + "     )      ORDER BY AMTF,cast(b.seqno as int)          ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@TOTAL", TOTAL));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@GROSS", GROSS));
            command.Parameters.Add(new SqlParameter("@NET", NET));
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.CommandType = CommandType.Text;
    

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData2(string AA, string AmountTotalEng)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("                             SELECT a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期  ");
            sb.Append("                             ,a.[InvoiceTo],a.[AmountTotal],@AmountTotalEng  AmountTotalEng,b.[SeqNo],b.[MarkNos],  ");
            sb.Append("                              SEQ+')'+b.[INDescription]  INDescription   ");
            sb.Append("                             ,b.[InQty] ,b.[UnitPrice]  ,b.[Amount],c.brand +' BRAND' as BRAND,c.TradeCondition as Trade,TT FROM [InvoiceM] as a  ");
            sb.Append("                             left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)  ");
            sb.Append("                             left join shipping_main as c on (a.shippingcode=c.shippingcode)   ");
            sb.Append("               LEFT JOIN (select shippingcode,InvoiceNo,InvoiceNo_seq,docentry,cast(  RANK() OVER (ORDER BY shippingcode,seqno) as varchar) SEQ  from InvoiceD where");
            sb.Append("                 shippingcode IN (" + AA + "  )) d ");
            sb.Append("               on (a.shippingcode=d.shippingcode and a.InvoiceNo=d.InvoiceNo and a.InvoiceNo_seq=d.InvoiceNo_seq and b.docentry=d.docentry)  ");
            sb.Append("               LEFT JOIN (select MAX(SEQNO) SEQNO,SHIPPINGCODE,'Y' TT from InvoiceD where   shippingcode IN (" + AA + "  )  ");
            sb.Append("               GROUP BY shippingcode) E ON (a.shippingcode=E.shippingcode AND B.SEQNO=E.SEQNO) ");
            sb.Append("               where  a.shippingcode IN (" + AA + "  )");
            sb.Append("               ORDER BY A.shippingcode,seqno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@AmountTotalEng", AmountTotalEng));
            command.CommandType = CommandType.Text;
          

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private bool CheckSerial(string sData, ref string FieldValue, System.Data.DataTable dt)
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
                FieldValue = Convert.ToString(dt.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }
        private bool IsDetailRow(string sData)
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
        private void SetRow(int iRow, string sData, ref string FieldValue, System.Data.DataTable dt)
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
                FieldValue = Convert.ToString(dt.Rows[iRow][FieldName]);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable T1 = GetSHIP();
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }
        }

        private System.Data.DataTable GetSHIP()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SHIPPINGCODE FROM SHIPPING_MAIN WHERE SUBSTRING(SHIPPINGCODE,3,8) BETWEEN @AA AND @BB AND SUBSTRING(SHIPPINGCODE,1,2) <> 'SI'  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        private System.Data.DataTable GetPACK2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "              SELECT SUM(ISNULL(CAST(sayTotal AS INT),0)) PACK,　SUM(CASE WHEN ColumnTotal LIKE '%ONE (1) CTNS ONLY%' THEN 1 ELSE 　ltrim(substring(ColumnTotal,CHARINDEX('=', ColumnTotal)+1,CHARINDEX('CTNS', ColumnTotal)-CHARINDEX('=', ColumnTotal)-1)) END)　CTNS　 FROM PackingListM  WHERE SHIPPINGCODE IN  (" + SHIPPINGCODE + ")  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetNET(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SUM([Net]) NET,SUM([Gross]) GROSS,SUM([Quantity]) QTY,CAST(SUM(CAST(SayTotal AS INT)) AS VARCHAR)+' PLTS'  CC  FROM [PackingListM] WHERE  shippingcode IN (" + AA + ")";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetEMP(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "	SELECT ISNULL(SUM(CAST(ISNULL(userName,0) AS INT)),0)　EMP  FROM  PackingListM    WHERE SHIPPINGCODE IN  (" + AA + ")";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetPACKACCOUNT(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT COUNT(*) FROM [PackingListD] WHERE SHIPPINGCODE=@SHIPPINGCODE ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetMARKC(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT TOP 1 MARK  FROM MARK WHERE SHIPPINGCODE=@SHIPPINGCODE ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private System.Data.DataTable GetINVOAMOUNT(string AA)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SUM(AMOUNT) AMOUNT FROM InvoiceD WHERE  shippingcode IN (" + AA + ")";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }
        private void SHIPMULTI_Load(object sender, EventArgs e)
        {

            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            ArrayList al = new ArrayList();
            try
            {
                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    string SHIP = listBox2.Items[i].ToString();
                    al.Add(SHIP);
                    UPDATEADD9(i, SHIP);
                }
                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }
                sb.Remove(sb.Length - 1, 1);
                System.Data.DataTable JJ1 = GetEMP(sb.ToString());

                System.Data.DataTable JJ2 = GetPACK2(sb.ToString());
                if (JJ2.Rows.Count > 0)
                {
                    string f = JJ2.Rows[0][0].ToString();
                    string f2 = JJ2.Rows[0][1].ToString();
                    string EMP = JJ1.Rows[0][0].ToString();
                    string f3 = "";
                    if (EMP != "0")
                    {
                        f3 = " + " + EMP + " EMPTY CTNS";
                    }
                   
                    // new Class1().NumberToString2(amountText, s, f2) + f3 + " ONLY.";
                    int amountText = Convert.ToInt32(f);
                    string s = f;
                    textBox3.Text = new Class1().NumberToString2(amountText, s, f2) + f3 + " ONLY.";
                }
            }
            catch { }


            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            if (globals.DBNAME == "禾中")
            {
                FileName = lsAppDir + "\\Excel\\AT\\PACK2M.xls";

            }
            else
            {
                if (USER.ToUpper() == "JOYCHEN")
                {
                    FileName = lsAppDir + "\\Excel\\PACK2MJ.xls";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\PACK2M.xls";
                }
            }

            System.Data.DataTable J1 = GetNET(sb.ToString());
            string QTY = "";
            string GROSS = "";
            string NET = "";
            string CC = "";

            if (J1.Rows.Count > 0)
            {
                DataRow drw = J1.Rows[0];
                GROSS = drw["GROSS"].ToString();
                QTY = drw["QTY"].ToString();
                NET = drw["NET"].ToString();
                CC = drw["CC"].ToString();
            }

            System.Data.DataTable H1 = GetOrderData3(sb.ToString(), textBox3.Text, QTY, GROSS, NET, CC);
            if (H1.Rows.Count > 0)
            {
                if (USER.ToUpper() == "JOYCHEN")
                {
                    GetExcelProduct22(FileName, H1, sb.ToString(),"N");
                }
                else
                {
                    GetExcelProduct22(FileName, H1, sb.ToString(), "Y");
                }
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private void UPDATEADD9(int  AMTF, string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE PackingListD SET AMTF =@AMTF WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@AMTF", AMTF));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));



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

using System;
using System.Windows.Forms;
using System.Text;
using System.IO;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
/// <summary>
/// ExcelReport 的摘要描述
/// </summary>
public class ExcelReport
{
    public ExcelReport()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    //刪除暫存報表
    //請使用相對路徑
    //Server.MapPath(@"~\ExcelSample");
    //固定路徑不建議使用 
    //"C:\\Inetpub\\wwwroot\\rma\\ExcelExport"
    public static void DeleteReport(string ReportPath, string DeleteAll)
    {
        string[] filenames = Directory.GetFiles(ReportPath);

        //檔案數
        if (filenames.Length > 100)
        {
            foreach (string file in filenames)
            {
                //字元數
                //if (file.Length > 40)
                //{
                //}

                if (DeleteAll == "Y")
                {
                    //刪除
                    System.IO.File.Delete(file);
                }
                else
                {
                    //檔案名稱大於 50 者刪除
                    if (file.ToString().Length > 50)
                    {
                        //刪除
                        System.IO.File.Delete(file);
                    }
                }

            }
        }
    }


    //明細檔使用的符號為 [[欄位]]
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
    public static bool IsDetailRow2(string sData)
    {

        if (sData.Length < 2)
        {
            return false;
        }
        if (sData.Substring(0, 2) == "**")
        {

            return true;
        }
        //}
        return false;
    }

    //主檔使用的符號為 <<欄位>>
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
    public static void SetRow2(System.Data.DataTable OrderData, int iRow, string sData, ref string FieldValue)
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



    public static void ExcelReportOutput2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
        Microsoft.Office.Interop.Excel.Range range = null;


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
                    //   range.Select();
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
                        //range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    DetailRow++;
                }

            }


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表2");

            pivotTable.RefreshTable();



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
            System.GC.WaitForPendingFinalizers();
            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelReportPOS(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
        Microsoft.Office.Interop.Excel.Range range = null;


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
                    //   range.Select();
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
                        //range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    DetailRow++;
                }

            }


            //



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable2 = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("POS2");

            pivotTable2.RefreshTable();

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable3 = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("POS3");

            pivotTable3.RefreshTable();

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(4);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable4 = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("POS4");

            pivotTable4.RefreshTable();

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(5);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable5 = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("POS5");

            pivotTable5.RefreshTable();

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
            System.GC.WaitForPendingFinalizers();
            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelReportOutputPLATE(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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
                    //   range.Select();
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
                        //range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    DetailRow++;
                }

            }





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表2");

            pivotTable.RefreshTable();



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
    public static void ExcelReportOutputANITA(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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
    public static void ExcelReportOutputLEMON(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string PRINT)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        //取得  Worksheet
        //Microsoft.Office.Interop.Excel.Range range1 = null;



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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

            if (PRINT != "N")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }


        }

    }
    public static void ExcelReportOutputLEMONFIT(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string PRINT)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                        range.Rows.AutoFit();
                    }

                    DetailRow++;
                }

            }


            //



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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

            if (PRINT != "N")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }


        }

    }
    public static void ExcelAD(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string PRINT, System.Data.DataTable OrderData2)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        //取得  Worksheet
        //Microsoft.Office.Interop.Excel.Range range1 = null;



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();

        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
        Microsoft.Office.Interop.Excel.Range range = null;

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

                        range.Rows.AutoFit();
                    }

                    DetailRow++;
                }

            }



            Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet2.Activate();

            int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
            int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;
            Microsoft.Office.Interop.Excel.Range range2 = null;

            string sTemp2 = string.Empty;
            string FieldValue2 = string.Empty;
            bool IsDetail2 = false;
            int DetailRow2 = 0;

            for (int iRecord = 1; iRecord <= iRowCnt2; iRecord++)
            {

                for (int iField = 1; iField <= iColCnt2; iField++)
                {
                    range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, iField]);
                    range2.Select();
                    sTemp2 = (string)range2.Text;
                    sTemp2 = sTemp2.Trim();

                    if (CheckSerial(OrderData2, sTemp2, ref FieldValue2))
                    {
                        range2.Value2 = FieldValue2;
                    }
                    if (IsDetailRow(sTemp2))
                    {
                        IsDetail2 = true;
                        DetailRow2 = iRecord;
                        break;
                    }

                }

            }
            if (DetailRow2 != 0)
            {

                for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData2.Rows.Count - 1)
                    {

                        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, 1]);
                        range2.EntireRow.Copy(oMissing);

                        range2.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt2; iField++)
                    {
                        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, iField]);
                        range2.Select();
                        sTemp2 = (string)range2.Text;
                        sTemp2 = sTemp2.Trim();

                        FieldValue2 = "";
                        SetRow(OrderData2, aRow, sTemp2, ref FieldValue2);

                        range2.Value2 = FieldValue2;


                    }

                    DetailRow2++;
                }

            }

            iRowCnt = excelSheet2.UsedRange.CurrentRegion.Cells.Rows.Count;
            iColCnt = excelSheet2.UsedRange.Cells.Columns.Count;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.CurrentRegion);
            range.Copy(oMissing);


            Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet3.Activate();

            SelectCell = "A" + (iRowCnt + 20).ToString();
            range = excelSheet3.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);

        }
        finally
        {


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

            if (PRINT != "N")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }


        }

    }
    public static void ExcelReportOutputPOTATO(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string DOCNAME)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        excelSheet.Name = DOCNAME;
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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
    public static void ExcelReportOutput(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }

            if (flag == "pivot")
            {
                //固定在第二頁
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表1");

                pivotTable.RefreshTable();
            }
            else
            {
                SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);
                range.Select();
            }

        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelNANCY(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, System.Data.DataTable OrderData2)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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

        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        Microsoft.Office.Interop.Excel.Range range = null;
        string sTemp = string.Empty;
        string FieldValue = string.Empty;
        bool IsDetail = false;
        int DetailRow = 0;

        for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
        {
            for (int iField = 1; iField <= 6; iField++)
            {
                string V = Convert.ToString(OrderData2.Rows[aRow][iField - 1]);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow + 2, iField]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                range.Value2 = V;
            }
        }

        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = 7;

        // progressBar1.Maximum = iRowCnt;



        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {



            for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
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


                        if (iField == 7)
                        {
                            int g1 = OrderData.Rows.Count - 1;
                            double TOTAL = Convert.ToDouble(OrderData.Rows[g1][5]);
                            double AMT = Convert.ToDouble(OrderData.Rows[aRow][5]);
                            double f1 = AMT / TOTAL;
                            range.Value2 = f1.ToString();
                        }
                    }

                    DetailRow++;
                }

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelDAVID(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 16;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {


            System.Data.DataTable G1 = OrderData;

            for (int i = 0; i <= G1.Rows.Count - 2; i++)
            {

                object sALL = "J" + (16).ToString();
                range = excelSheet.get_Range("A1", sALL);
                range.Copy(oMissing);



                if (i == 0)
                {
                    SelectCell = "A" + ((iRowCnt) + (3)).ToString();

                }
                else
                {
                    SelectCell = "A" + ((iRowCnt * (i + 1)) + (2 * (i + 1)) + 1).ToString();
                }
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);
            }

            int F = 0;
            for (int i = 0; i <= G1.Rows.Count - 1; i++)
            {
                string INVOICE = G1.Rows[i]["INVOICE"].ToString();
                string DDATE = G1.Rows[i]["DDATE"].ToString();
                string ITEMCODE = G1.Rows[i]["ITEMCODE"].ToString();
                string DSCRIPTION = G1.Rows[i]["DSCRIPTION"].ToString();
                string QTY = G1.Rows[i]["QTY"].ToString();

                F = 18 * i;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2 + F, 2]);
                range.Select();
                range.Value2 = INVOICE;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4 + F, 1]);
                range.Select();
                range.Value2 = ITEMCODE;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4 + F, 6]);
                range.Select();
                range.Value2 = DSCRIPTION;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6 + F, 1]);
                range.Select();
                range.Value2 = DDATE;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6 + F, 2]);
                range.Select();
                range.Value2 = QTY;

            }
            int aa = excelSheet.UsedRange.Cells.Rows.Count;
            excelSheet.PageSetup.PrintArea = "$A$1:$J$" + aa.ToString();
            //SelectCell = "A1";
            //range = excelSheet.get_Range(SelectCell, SelectCell);
            //range.Select();


        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelGBCHOICE(System.Data.DataTable OrderData, System.Data.DataTable OrderData12, string D3, string D4, string ExcelFile, string OutPutFile, System.Data.DataTable OrderData2, System.Data.DataTable OrderData3, string D1, string D5, string D6)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                        string AA1 = FieldValue.Trim();
                        int GG1 = AA1.IndexOf("合計");

                        if (GG1 != -1)
                        {
                            for (int L = 9; L <= 12; L++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, L]);
                                range.Select();
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            }
                        }
                    }

                    DetailRow++;
                }

            }

            System.Data.DataTable dt5 = OrderData12;
            if (dt5.Rows.Count > 0)
            {
                for (int i = 0; i <= dt5.Rows.Count - 1; i++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 1]);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 5]);
                    range.Value2 = dt5.Rows[i][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 6]);
                    range.Value2 = dt5.Rows[i][1].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 7]);
                    range.Value2 = dt5.Rows[i][2].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 8]);
                    range.Value2 = dt5.Rows[i][3].ToString();
                }
            }

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
            range.Value2 = D4;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 1]);
            range.Value2 = D3;




            //Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            //excelSheet2.Activate();

            //int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
            //int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;

            //Microsoft.Office.Interop.Excel.Range range2 = null;

            //string sTemp2 = string.Empty;
            //string FieldValue2 = string.Empty;
            //bool IsDetail2 = false;
            //int DetailRow2 = 0;

            //for (int iRecord2 = 1; iRecord2 <= iRowCnt2; iRecord2++)
            //{
            //    for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
            //    {
            //        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord2, iField2]);
            //        range2.Select();
            //        sTemp2 = (string)range2.Text;
            //        sTemp2 = sTemp2.Trim();

            //        if (CheckSerial(OrderData2, sTemp2, ref FieldValue2))
            //        {
            //            range2.Value2 = FieldValue2;
            //        }

            //        if (IsDetailRow(sTemp2))
            //        {
            //            IsDetail2 = true;
            //            DetailRow2 = iRecord2;
            //            break;
            //        }

            //    }

            //}

            //if (DetailRow2 != 0)
            //{

            //    for (int aRow2 = 0; aRow2 <= OrderData2.Rows.Count - 1; aRow2++)
            //    {

            //        //最後一筆不作
            //        if (aRow2 != OrderData2.Rows.Count - 1)
            //        {

            //            range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, 1]);
            //            range2.EntireRow.Copy(oMissing);

            //            range2.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
            //                oMissing);
            //        }


            //        for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
            //        {
            //            range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, iField2]);
            //            range2.Select();
            //            sTemp2 = (string)range2.Text;
            //            sTemp2 = sTemp2.Trim();

            //            FieldValue2 = "";
            //            SetRow(OrderData2, aRow2, sTemp2, ref FieldValue2);

            //            range2.Value2 = FieldValue2;


            //        }

            //        DetailRow2++;
            //    }

            //}



            Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
            excelSheet3.Activate();

            int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
            int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range3 = null;

            string sTemp3 = string.Empty;
            string FieldValue3 = string.Empty;
            bool IsDetail3 = false;
            int DetailRow3 = 0;

            for (int iRecord3 = 1; iRecord3 <= iRowCnt3; iRecord3++)
            {

                for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
                {
                    range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord3, iField3]);
                    range3.Select();
                    sTemp3 = (string)range3.Text;
                    sTemp3 = sTemp3.Trim();

                    if (CheckSerial(OrderData3, sTemp3, ref FieldValue3))
                    {
                        range3.Value2 = FieldValue3;
                    }

                    if (IsDetailRow(sTemp3))
                    {
                        IsDetail3 = true;
                        DetailRow3 = iRecord3;
                        break;
                    }

                }

            }

            if (DetailRow3 != 0)
            {

                for (int aRow3 = 0; aRow3 <= OrderData3.Rows.Count - 1; aRow3++)
                {


                    if (aRow3 != OrderData3.Rows.Count - 1)
                    {

                        range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, 1]);
                        range3.EntireRow.Copy(oMissing);

                        range3.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
                    {
                        range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, iField3]);
                        range3.Select();
                        sTemp3 = (string)range3.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData3, aRow3, sTemp3, ref FieldValue3);
                        string AA = FieldValue3.Trim();

                        range3.Value2 = FieldValue3;

                        int G1 = AA.IndexOf("合計");

                        if (G1 != -1)
                        {
                            for (int L = 1; L <= 5; L++)
                            {
                                range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, L]);
                                range3.Select();
                                range3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                            }
                        }
                    }

                    DetailRow3++;
                }

            }

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[1, 2]);
            range3.Value2 = D5;

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[2, 2]);
            range3.Value2 = D1;

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[3, 2]);
            range3.Value2 = D6;


        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelGBCHOICE1(System.Data.DataTable OrderData, System.Data.DataTable OrderData12, string D3, string D4, string ExcelFile, string OutPutFile)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

            System.Data.DataTable dt5 = OrderData12;
            if (dt5.Rows.Count > 0)
            {
                for (int i = 0; i <= dt5.Rows.Count - 1; i++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 1]);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 5]);
                    range.Value2 = dt5.Rows[i][0].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 6]);
                    range.Value2 = dt5.Rows[i][1].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 7]);
                    range.Value2 = dt5.Rows[i][2].ToString();
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 7, 8]);
                    range.Value2 = dt5.Rows[i][3].ToString();
                }
            }

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
            range.Value2 = D4;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 1]);
            range.Value2 = D3;



            //Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            //excelSheet2.Activate();

            //int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
            //int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;

            //Microsoft.Office.Interop.Excel.Range range2 = null;

            //string sTemp2 = string.Empty;
            //string FieldValue2 = string.Empty;
            //bool IsDetail2 = false;
            //int DetailRow2 = 0;

            //for (int iRecord2 = 1; iRecord2 <= iRowCnt2; iRecord2++)
            //{
            //    for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
            //    {
            //        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord2, iField2]);
            //        range2.Select();
            //        sTemp2 = (string)range2.Text;
            //        sTemp2 = sTemp2.Trim();

            //        if (CheckSerial(OrderData2, sTemp2, ref FieldValue2))
            //        {
            //            range2.Value2 = FieldValue2;
            //        }

            //        if (IsDetailRow(sTemp2))
            //        {
            //            IsDetail2 = true;
            //            DetailRow2 = iRecord2;
            //            break;
            //        }

            //    }

            //}

            //if (DetailRow2 != 0)
            //{

            //    for (int aRow2 = 0; aRow2 <= OrderData2.Rows.Count - 1; aRow2++)
            //    {

            //        //最後一筆不作
            //        if (aRow2 != OrderData2.Rows.Count - 1)
            //        {

            //            range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, 1]);
            //            range2.EntireRow.Copy(oMissing);

            //            range2.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
            //                oMissing);
            //        }


            //        for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
            //        {
            //            range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, iField2]);
            //            range2.Select();
            //            sTemp2 = (string)range2.Text;
            //            sTemp2 = sTemp2.Trim();

            //            FieldValue2 = "";
            //            SetRow(OrderData2, aRow2, sTemp2, ref FieldValue2);

            //            range2.Value2 = FieldValue2;


            //        }

            //        DetailRow2++;
            //    }

            //}



            //Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
            //excelSheet3.Activate();

            //int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
            //int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;

            //Microsoft.Office.Interop.Excel.Range range3 = null;

            //string sTemp3 = string.Empty;
            //string FieldValue3 = string.Empty;
            //bool IsDetail3 = false;
            //int DetailRow3 = 0;

            //for (int iRecord3 = 1; iRecord3 <= iRowCnt3; iRecord3++)
            //{

            //    for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
            //    {
            //        range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord3, iField3]);
            //        range3.Select();
            //        sTemp3 = (string)range3.Text;
            //        sTemp3 = sTemp3.Trim();

            //        if (CheckSerial(OrderData3, sTemp3, ref FieldValue3))
            //        {
            //            range3.Value2 = FieldValue3;
            //        }

            //        if (IsDetailRow(sTemp3))
            //        {
            //            IsDetail3 = true;
            //            DetailRow3 = iRecord3;
            //            break;
            //        }

            //    }

            //}

            //if (DetailRow3 != 0)
            //{

            //    for (int aRow3 = 0; aRow3 <= OrderData3.Rows.Count - 1; aRow3++)
            //    {


            //        if (aRow3 != OrderData3.Rows.Count - 1)
            //        {

            //            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, 1]);
            //            range3.EntireRow.Copy(oMissing);

            //            range3.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
            //                oMissing);
            //        }


            //        for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
            //        {
            //            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, iField3]);
            //            range3.Select();
            //            sTemp3 = (string)range3.Text;
            //            sTemp3 = sTemp3.Trim();

            //            FieldValue3 = "";
            //            SetRow(OrderData3, aRow3, sTemp3, ref FieldValue3);
            //            string AA = FieldValue3.Trim();

            //            range3.Value2 = FieldValue3;

            //            int G1 = AA.IndexOf("合計");

            //            if (G1 != -1)
            //            {
            //                for (int L = 1; L <= 5; L++)
            //                {
            //                    range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, L]);
            //                    range3.Select();
            //                    range3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //                }
            //            }
            //        }

            //        DetailRow3++;
            //    }

            //}

            //range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[1, 2]);
            //range3.Value2 = D5;

            //range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[2, 2]);
            //range3.Value2 = D1;

            //range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[3, 2]);
            //range3.Value2 = D6;
            //SelectCell = "A1";
            //range = excelSheet.get_Range(SelectCell, SelectCell);
            //range.Select();

        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelGBCHOICE2(string ExcelFile, string OutPutFile, System.Data.DataTable OrderData3, string D1, string D5, string D6)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {







            Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
            int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range3 = null;

            string sTemp3 = string.Empty;
            string FieldValue3 = string.Empty;
            bool IsDetail3 = false;
            int DetailRow3 = 0;

            for (int iRecord3 = 1; iRecord3 <= iRowCnt3; iRecord3++)
            {

                for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
                {
                    range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[iRecord3, iField3]);
                    range3.Select();
                    sTemp3 = (string)range3.Text;
                    sTemp3 = sTemp3.Trim();

                    if (CheckSerial(OrderData3, sTemp3, ref FieldValue3))
                    {
                        range3.Value2 = FieldValue3;
                    }

                    if (IsDetailRow(sTemp3))
                    {
                        IsDetail3 = true;
                        DetailRow3 = iRecord3;
                        break;
                    }

                }

            }

            if (DetailRow3 != 0)
            {

                for (int aRow3 = 0; aRow3 <= OrderData3.Rows.Count - 1; aRow3++)
                {


                    if (aRow3 != OrderData3.Rows.Count - 1)
                    {

                        range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, 1]);
                        range3.EntireRow.Copy(oMissing);

                        range3.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField3 = 1; iField3 <= iColCnt3; iField3++)
                    {
                        range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, iField3]);
                        range3.Select();
                        sTemp3 = (string)range3.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData3, aRow3, sTemp3, ref FieldValue3);
                        string AA = FieldValue3.Trim();

                        range3.Value2 = FieldValue3;

                        int G1 = AA.IndexOf("合計");

                        if (G1 != -1)
                        {
                            for (int L = 1; L <= 5; L++)
                            {
                                range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3, L]);
                                range3.Select();
                                range3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                            }
                        }
                    }

                    DetailRow3++;
                }

            }

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[1, 2]);
            range3.Value2 = D5;

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[2, 2]);
            range3.Value2 = D1;

            range3 = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[3, 2]);
            range3.Value2 = D6;
            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelReportOutputS2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }

            if (flag == "pivot")
            {
                //固定在第二頁
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表1");

                pivotTable.RefreshTable();
            }

        }
        finally
        {


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

    public static void ExcelReportOutputSUNNY(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string FLAG, string DOCTYPE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 17 + OrderData.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

            if (FLAG == "Y")
            {
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.CurrentRegion);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 1).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);
            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();


            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2 * 2) - 4); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }
                int FG2 = sTemp.IndexOf("日      期:");
                int FG3 = sTemp.IndexOf("客戶編號:");
                int FG4 = sTemp.IndexOf("連絡人員:");
                int FG5 = sTemp.IndexOf("聯絡電話:");
                if (FG2 != -1 || FG3 != -1 || FG4 != -1 || FG5 != -1)
                {
                    range.Select();
                    range.RowHeight = 16.50;
                }
                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (FLAG == "Y")
                {
                    if (sTemp == "核准:")
                    {
                        range.Select();
                        range.RowHeight = 41.25;

                    }
                }
                int N1 = OrderData.Rows.Count;
                if (N1 <= 3)
                {
                    int D1 = sTemp.IndexOf("***********");
                    if (D1 != -1)
                    {
                        if (D == 0)
                        {
                            range.Select();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                            range.Select();
                            int N2 = 0;
                            if (N1 == 1)
                            {
                                N2 = 3;
                            }
                            if (N1 == 2)
                            {
                                N2 = 2;
                            }
                            if (N1 == 3)
                            {
                                N2 = 1;
                            }
                            for (int S = 0; S < N2; S++)
                            {
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                               oMissing);
                            }
                        }

                        D = 1;
                    }
                }


            }

        }
        finally
        {


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
    public static void ExcelReportOutputRMAWH(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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



            //取得
            //取得 Excel 的使用區域
            iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 2).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

            for (int iField = iRowCnt; iField <= (iRowCnt * 2); iField++)
            {
                //簽名視同收貨數量及貨物狀況皆無異常
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();
                int num1;
                if (int.TryParse(sTemp, out num1) == true)
                {
                    range.RowHeight = 21;
                }
                int H3 = sTemp.IndexOf("總");
                if (H3 != -1)
                {
                    range.RowHeight = 23.25;
                }
                int H2 = sTemp.IndexOf("貨物異常備註");
                if (H2 != -1)
                {
                    range.RowHeight = 90.75;
                }

                int H5 = sTemp.IndexOf("領貨人");
                if (H5 != -1)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                    range.Select();
                    range.Value2 = "收貨人：";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 4]);
                    range.Select();
                    range.Value2 = "領貨人：";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 5]);
                    range.Select();
                    range.Value2 = "";


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 7]);
                    range.Select();
                    range.Value2 = "放貨人：";

                }
                //領貨人：
                int H4 = sTemp.IndexOf("簽名視同收貨數量及貨物狀況皆無異常");
                if (H4 != -1)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 2, 1]);
                    range.Select();
                    range.EntireRow.Delete(Excel.XlDirection.xlToLeft);
                }

                //取出欄位值 - 科目
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 11]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();
                int H1 = sTemp.IndexOf("聯");
                if (H1 != -1)
                {
                    range.Value2 = "一\n式\n兩\n聯\n     (二)\n領\n貨\n人\n\n繳\n回\n聯";
                }




            }



            excelSheet.PageSetup.PrintArea = "A1:K" + iRowCnt * 2;


        }
        finally
        {


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


    public static void ExcelReportOutputCOMP(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();


        }
        finally
        {


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
    public static void ExcelReportOutputJ2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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




            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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




        }

    }
    public static void ExcelReportTONY(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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
                        int d = FieldValue.Length;
                        if (d > 1045)
                        {
                            range.Value2 = FieldValue.Trim().Substring(0, 1020);
                        }
                        else
                        {
                            range.Value2 = FieldValue.Trim();
                        }

                    }

                    DetailRow++;
                }

            }


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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
            System.GC.WaitForPendingFinalizers();
            // MessageBox.Show("產生一個檔案->" + NewFileName);

            string Msg = string.Empty;



        }

    }
    public static void ExcelReportOutputJOCELIN(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag, string T2)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 1).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();


            for (int iField = 4; iField <= ((iRowCnt * 2) - 4); iField++)
            {


                //取出欄位值 - 科目
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();


                if (sTemp == "第一聯-ACME 存根聯")
                {
                    range.Select();
                    range.Value2 = T2;
                }



            }
        }
        finally
        {


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


    public static void ExcelReportOutputJOCELIN2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag, string T2)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.CurrentRegion.Cells.Rows.Count;
        int iColCnt = 9;

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

            //


            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.CurrentRegion.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.CurrentRegion);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + +2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }

            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();


            for (int iField = 4; iField <= ((iRowCnt2 * 2) - 4); iField++)
            {


                //取出欄位值 - 科目
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();


                if (sTemp == "第一聯-ACME 存根聯")
                {
                    range.Select();
                    range.Value2 = T2;
                }

                int H1 = sTemp.IndexOf("派送司機簽名");
                if (H1 != -1)
                {
                    range.Select();
                    range.RowHeight = 30;

                }


            }
        }
        finally
        {


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

            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelReportOutputLA(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string P1, string P2, string DOCTYPE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 17 + OrderData.Rows.Count;

        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;



            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 30, 300, 50, 40);
            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 150, 305, 30, 30);

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




            iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.CurrentRegion);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 1).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);




            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();
            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2)); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }
                int FG2 = sTemp.IndexOf("日      期:");
                int FG3 = sTemp.IndexOf("客戶編號:");
                int FG4 = sTemp.IndexOf("連絡人員:");
                int FG5 = sTemp.IndexOf("聯絡電話:");
                if (FG2 != -1 || FG3 != -1 || FG4 != -1 || FG5 != -1)
                {
                    range.Select();
                    range.RowHeight = 16.50;
                }

                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (sTemp == "核准:")
                {
                    range.Select();
                    range.RowHeight = 45.75;

                }
                int N1 = OrderData.Rows.Count;
                if (N1 <= 3)
                {
                    int D1 = sTemp.IndexOf("***********");
                    if (D1 != -1)
                    {
                        if (D == 0)
                        {
                            range.Select();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                            range.Select();
                            int N2 = 0;
                            if (N1 == 1)
                            {
                                N2 = 3;
                            }
                            if (N1 == 2)
                            {
                                N2 = 2;
                            }
                            if (N1 == 3)
                            {
                                N2 = 1;
                            }
                            for (int S = 0; S < N2; S++)
                            {
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                               oMissing);
                            }
                        }

                        D = 1;
                    }
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
                OutPutFile = OutPutFile.Replace("“", "").Replace("”", "");
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
            if (DOCTYPE == "Y")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }
        }

    }
    public static void ExcelFUNHOUR2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string P1, string SS, string DOCTYPE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
        excelApp.Visible = false;

        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();

        int iRowCnt = 20 + OrderData.Rows.Count;

        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;



        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;



            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 30, 300, 50, 40);

            string[] arrurl = SS.Split(new Char[] { '/' });

            int O = 0;

            if (SS != "")
            {
                foreach (string i in arrurl)
                {
                    O++;
                    string B2 = "//acmew08r2ap//table//放貨單//" + i.ToString() + ".jpg";
                    if (O == 1)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 370, 60, 25);
                    }

                    if (O == 2)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 80, 370, 60, 25);
                    }

                    if (O == 3)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 150, 370, 60, 25);
                    }

                    if (O == 4)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 220, 370, 60, 25);
                    }

                    if (O == 5)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 395, 60, 25);
                    }

                    if (O == 6)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 80, 395, 60, 25);
                    }

                    if (O == 7)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 150, 395, 60, 25);
                    }

                    if (O == 8)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 220, 395, 60, 25);
                    }

                    if (O == 9)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 190, 395, 60, 25);
                    }
                }
            }

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



            SelectCell = "E" + (iRowCnt).ToString();
            iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            range = excelSheet.get_Range("A1", SelectCell);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 1).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);




            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();
            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2)); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }
                int FG2 = sTemp.IndexOf("日      期:");
                int FG3 = sTemp.IndexOf("客戶編號:");
                int FG4 = sTemp.IndexOf("連絡人員:");
                int FG5 = sTemp.IndexOf("聯絡電話:");
                if (FG2 != -1 || FG3 != -1 || FG4 != -1 || FG5 != -1)
                {
                    range.Select();
                    range.RowHeight = 16.50;
                }

                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (sTemp == "核准:")
                {
                    range.Select();
                    range.RowHeight = 45.75;

                }
                //int N1 = OrderData.Rows.Count;
                //if (N1 <= 3)
                //{
                //    int D1 = sTemp.IndexOf("***********");
                //    if (D1 != -1)
                //    {
                //        if (D == 0)
                //        {
                //            range.Select();

                //            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                //            range.Select();
                //            int N2 = 0;
                //            if (N1 == 1)
                //            {
                //                N2 = 3;
                //            }
                //            if (N1 == 2)
                //            {
                //                N2 = 2;
                //            }
                //            if (N1 == 3)
                //            {
                //                N2 = 1;
                //            }
                //            for (int S = 0; S < N2; S++)
                //            {
                //                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                //                               oMissing);
                //            }
                //        }

                //        D = 1;
                //    }
                //}


            }
        }
        finally
        {

            //  string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
            //HttpContext.Current.User.Identity.Name.ToString() + Util.GetDate14() +
            //Path.GetFileName(ExcelFile);

            try
            {
                OutPutFile = OutPutFile.Replace("“", "").Replace("”", "");
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

            if (DOCTYPE == "Y")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }
        }

    }

    public static void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent, string DATA)
    {
        MailMessage message = new MailMessage();

        message.From = new MailAddress("LleytonChen@acmepoint.com", "系統發送");
        message.To.Add(new MailAddress(MailAddress));

        string template;
        StreamReader objReader;

        objReader = new StreamReader(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\MailTemplates\\RMALEMON.htm");

        template = objReader.ReadToEnd();
        objReader.Close();
        template = template.Replace("##FirstName##", SlpName);
        template = template.Replace("##Content##", MailContent);
        template = template.Replace("##DATA##", DATA);
        message.Subject = strSubject;
        message.Body = template;
        message.IsBodyHtml = true;
        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
        string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
        string[] filenames = Directory.GetFiles(OutPutFile);
        foreach (string file in filenames)
        {


            message.Attachments.Add(new Attachment(file));

        }

        SmtpClient client = new SmtpClient();
        client.Host = "ms.mailcloud.com.tw";
        client.UseDefaultCredentials = true;

        string pwd = "@cmeworkflow";
        client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);

        try
        {
            client.Send(message);
            foreach (Attachment item in message.Attachments)
            {
                item.Dispose();   //一定要释放该对象,否则无法删除附件
            }

            MessageBox.Show("信件已寄出");
        }
        catch (SmtpFailedRecipientsException ex)
        {
            for (int i = 0; i < ex.InnerExceptions.Length; i++)
            {
                SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                if (status == SmtpStatusCode.MailboxBusy ||
                    status == SmtpStatusCode.MailboxUnavailable)
                {
                    //  SetMsg("Delivery failed - retrying in 5 seconds.");
                    System.Threading.Thread.Sleep(5000);
                    client.Send(message);
                }
                else
                {
                    // SetMsg(String.Format("Failed to deliver message to {0}",
                    // ex.InnerExceptions[i].FailedRecipient));
                }
            }
        }
        catch (Exception ex)
        {
            //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
            //        ex.ToString()));
        }

    }
    public static void ExcelReportOutputLA2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string P2)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;

            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 0, 470, 515, 100);

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





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {



            try
            {
                OutPutFile = OutPutFile.Replace("“", "").Replace("”", "");
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

            System.Diagnostics.Process.Start(OutPutFile);
        }

    }

    public static void ExcelFUNHOUR(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string P2, string S)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;

            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 0, 480, 517, 85);

            string[] arrurl = S.Split(new Char[] { '/' });

            int O = 0;

            if (S != "")
            {
                foreach (string i in arrurl)
                {
                    O++;
                    string B2 = "//acmew08r2ap//table//放貨單//" + i.ToString() + ".jpg";
                    if (O == 1)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 575, 70, 30);
                    }

                    if (O == 2)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 100, 575, 70, 30);
                    }

                    if (O == 3)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 190, 575, 70, 30);
                    }

                    if (O == 4)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 610, 70, 30);
                    }

                    if (O == 5)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 100, 610, 70, 30);
                    }

                    if (O == 6)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 190, 610, 70, 30);
                    }

                    if (O == 7)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 645, 70, 20);
                    }

                    if (O == 8)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 100, 645, 70, 30);
                    }

                    if (O == 9)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 190, 645, 70, 30);
                    }
                }
            }

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





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {



            try
            {
                OutPutFile = OutPutFile.Replace("“", "").Replace("”", "");
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

            System.Diagnostics.Process.Start(OutPutFile);
        }

    }
    public static void ExcelReportOutputANYA(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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



        }

    }

    public static void ExcelReportOutputANYA2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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
            //delete
            int J = 0;
            for (int S = 1; S <= 18; S++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, S]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();


                if (String.IsNullOrEmpty(sTemp))
                {
                    J = J + 1;
                    if (J <= 2)
                    {
                        range.EntireColumn.Delete(Excel.XlDirection.xlToLeft);
                        S = S - 1;
                    }

                }
            }




            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {


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



        }

    }
    public static void ExcelHelen(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string A1, string TYPE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 17 + OrderData.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                if (TYPE == "A")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 3]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }
                if (TYPE == "B")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }
                if (TYPE == "C")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";
                }
            }




            iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.CurrentRegion);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 1).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();
            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2 * 2) - 4); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }
                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (sTemp == "核准:")
                {
                    range.Select();
                    range.RowHeight = 41.25;

                }
                int N1 = OrderData.Rows.Count;
                if (N1 <= 3)
                {
                    int D1 = sTemp.IndexOf("***********");
                    if (D1 != -1)
                    {
                        if (D == 0)
                        {
                            range.Select();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                            range.Select();
                            int N2 = 0;
                            if (N1 == 1)
                            {
                                N2 = 3;
                            }
                            if (N1 == 2)
                            {
                                N2 = 2;
                            }
                            if (N1 == 3)
                            {
                                N2 = 1;
                            }
                            for (int S = 0; S < N2; S++)
                            {
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                               oMissing);
                            }
                        }

                        D = 1;
                    }
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
    public static void ExcelHelenPIC(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string A1, string P1, string P2, string TYPE, string DOCTYPE, string START)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 17 + OrderData.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
        //int iColCnt =4;
        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;


            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 30, 300, 50, 40);
            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 150, 305, 30, 30);
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


                if (TYPE == "A")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 3]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }
                if (TYPE == "B")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }

                if (TYPE == "C")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";
                }
            }


            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.CurrentRegion);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 1).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);


            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2)); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();
                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }

                int FG2 = sTemp.IndexOf("日      期:");
                int FG3 = sTemp.IndexOf("客戶編號:");
                int FG4 = sTemp.IndexOf("連絡人員:");
                int FG5 = sTemp.IndexOf("聯絡電話:");
                if (FG2 != -1 || FG3 != -1 || FG4 != -1 || FG5 != -1)
                {
                    range.Select();
                    range.RowHeight = 16.50;
                }
                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (sTemp == "核准:")
                {
                    range.Select();
                    range.RowHeight = 45.75;

                }
                int N1 = OrderData.Rows.Count;
                if (N1 <= 3)
                {
                    int D1 = sTemp.IndexOf("***********");
                    if (D1 != -1)
                    {
                        if (D == 0)
                        {
                            range.Select();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                            range.Select();
                            int N2 = 0;
                            if (N1 == 1)
                            {
                                N2 = 3;
                            }
                            if (N1 == 2)
                            {
                                N2 = 2;
                            }
                            if (N1 == 3)
                            {
                                N2 = 1;
                            }
                            for (int S = 0; S < N2; S++)
                            {
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                               oMissing);
                            }
                        }

                        D = 1;
                    }
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

            if (START == "Y")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }

        }

    }
    public static void ExcelFUNHOUR3(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string A1, string P1, string TYPE, string SS, string DOCTYPE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = 20 + OrderData.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
        //int iColCnt =4;
        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;


            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 30, 300, 50, 40);

            string[] arrurl = SS.Split(new Char[] { '/' });

            int O = 0;

            if (SS != "")
            {
                foreach (string i in arrurl)
                {
                    O++;
                    string B2 = "//acmew08r2ap//table//放貨單//" + i.ToString() + ".jpg";
                    if (O == 1)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 370, 60, 25);
                    }

                    if (O == 2)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 80, 370, 60, 25);
                    }

                    if (O == 3)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 150, 370, 60, 25);
                    }

                    if (O == 4)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 220, 370, 60, 25);
                    }

                    if (O == 5)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 10, 395, 60, 25);
                    }

                    if (O == 6)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 80, 395, 60, 25);
                    }

                    if (O == 7)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 150, 395, 60, 25);
                    }

                    if (O == 8)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 220, 395, 60, 25);
                    }

                    if (O == 9)
                    {
                        excelSheet.Shapes.AddPicture(B2, Microsoft.Office.Core.MsoTriState.msoFalse,
          Microsoft.Office.Core.MsoTriState.msoTrue, 190, 395, 60, 25);
                    }
                }
            }
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


                if (TYPE == "A")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 3]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }
                if (TYPE == "B")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "PO";
                }

                if (TYPE == "C")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.Value2 = A1 + "料號";
                }
            }


            SelectCell = "E" + (iRowCnt).ToString();
            iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            range = excelSheet.get_Range("A1", SelectCell);
            range.Copy(oMissing);

            SelectCell = "A" + (iRowCnt + 1).ToString();
            range = excelSheet.get_Range(SelectCell, SelectCell);

            range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);

            int D = 0;
            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iField = 1; iField <= ((iRowCnt2)); iField++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();
                if (sTemp == "進金生實業股份有限公司")
                {
                    range.Select();
                    range.RowHeight = 32.25;

                }
                int FG1 = sTemp.IndexOf("放貨單WH");
                if (FG1 != -1)
                {
                    range.Select();
                    range.RowHeight = 26.15;
                }

                int FG2 = sTemp.IndexOf("日      期:");
                int FG3 = sTemp.IndexOf("客戶編號:");
                int FG4 = sTemp.IndexOf("連絡人員:");
                int FG5 = sTemp.IndexOf("聯絡電話:");
                if (FG2 != -1 || FG3 != -1 || FG4 != -1 || FG5 != -1)
                {
                    range.Select();
                    range.RowHeight = 16.50;
                }
                if (sTemp == "送貨地址:")
                {
                    range.Select();
                    range.RowHeight = 39;

                }

                if (sTemp == "核准:")
                {
                    range.Select();
                    range.RowHeight = 45.75;

                }
                //int N1 = OrderData.Rows.Count;
                //if (N1 <= 3)
                //{
                //    int D1 = sTemp.IndexOf("***********");
                //    if (D1 != -1)
                //    {
                //        if (D == 0)
                //        {
                //            range.Select();

                //            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iField + 1, 1]);
                //            range.Select();
                //            int N2 = 0;
                //            if (N1 == 1)
                //            {
                //                N2 = 3;
                //            }
                //            if (N1 == 2)
                //            {
                //                N2 = 2;
                //            }
                //            if (N1 == 3)
                //            {
                //                N2 = 1;
                //            }
                //            for (int S = 0; S < N2; S++)
                //            {
                //                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                //                               oMissing);
                //            }
                //        }

                //        D = 1;
                //    }
                //}


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


            if (DOCTYPE == "Y")
            {
                System.Diagnostics.Process.Start(OutPutFile);
            }

        }

    }
    public static void APPLE(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string FILENAME, string USER, string COMPANY, int IF, string EXP, string OHEM)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


                        if (iField == IF)
                        {
                            if (FieldValue == "S")
                            {
                                for (int L = 1; L <= IF - 2; L++)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, L]);
                                    range.Select();
                                    range.Font.Bold = true;

                                }
                            }
                            if (FieldValue == "I")
                            {
                                for (int L = 1; L <= IF - 2; L++)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, L]);
                                    range.Select();
                                    range.Interior.ColorIndex = 35;
                                }
                            }
                        }


                    }

                    DetailRow++;
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, IF]);
                range.Select();
                range.EntireColumn.Delete(Excel.XlDirection.xlToLeft);
            }





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {



            try
            {
                excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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

            if (EXP == "Y")
            {
                string Msg = string.Empty;
                string Mo;
                //
                int G1 = OutPutFile.LastIndexOf("//");
                string OUTPUT2 = @"D:\Users\\" + USER + "\\Desktop\\" + FILENAME;
                if (OHEM != "SUNNYWANG")
                //      if (COMPANY == "禾中" || OHEM == "NANCYTSAI")
                {
                    System.Diagnostics.Process.Start(OutPutFile);
                }
                else
                {
                    File.Copy(OutPutFile, OUTPUT2, true);
                    System.Diagnostics.Process.Start(OUTPUT2);
                }
            }




        }

    }
    public static void ACC(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string FIRST, string LAST, System.Data.DataTable EUNICE)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
        excelSheet.Name = "明細表";

        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;
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

                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 4]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (sTemp == "3rd parties sub total" || sTemp == "Related Parties sub total")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;


                        }
                        if (sTemp == "業務")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 5]);
                            range.Select();
                            range.Value2 = "總數量";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 6]);
                            range.Select();
                            range.Value2 = "總收入";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 7]);
                            range.Select();
                            range.Value2 = "總成本";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 8]);
                            range.Select();
                            range.Value2 = "總毛利";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 9]);
                            range.Select();
                            range.Value2 = "總毛利率";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 10]);
                            range.Select();
                            range.Value2 = "Sales%";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                        }


                    }


                    DetailRow++;


                }

            }






            object SelectCell_From = "E3";
            object SelectCell_To = "H" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "L3";
            SelectCell_To = "O" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "Q3";
            SelectCell_To = "T" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "V3";
            SelectCell_To = "Y" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "AA3";
            SelectCell_To = "AD" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "AF3";
            SelectCell_To = "AI" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            string YEAR = FIRST.Substring(0, 4);
            DateTime df = Convert.ToDateTime(FIRST);
            string Month = df.ToString("MMMM", new System.Globalization.DateTimeFormatInfo()).Substring(0, 3);

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 2]);
            range.Select();
            range.Value2 = YEAR + "-" + Month;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 4]);
            range.Select();
            range.Value2 = "Report date:" + FIRST + "~" + LAST;

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();



            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Name = "總表";
            excelSheet.Activate();

            for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
            {



                for (int iField = 1; iField <= iColCnt; iField++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    if (CheckSerial(EUNICE, sTemp, ref FieldValue))
                    {
                        range.Value2 = FieldValue;
                    }


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

                for (int aRow = 0; aRow <= EUNICE.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != EUNICE.Rows.Count - 1)
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
                        SetRow(EUNICE, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }


                    DetailRow++;


                }

            }



            excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        }
        finally
        {

            try
            {
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
    public static void MARK(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string Y1)
    {
        int gg = 0;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
        excelSheet.Name = "嘜頭";

        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;
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


            for (int aRow = 1; aRow <= OrderData.Rows.Count; aRow++)
            {


                //取出欄位值 - 科目
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp != "")
                {
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                }


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 2]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();

                if (sTemp != "")
                {
                    range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }

            }


            excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        }
        finally
        {

            try
            {
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
    public static void NANCY(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
    {

        int gg = 0;
        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
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

                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 4]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();




                    }


                    DetailRow++;


                }

            }







            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表8");

            pivotTable.RefreshTable();


        }
        finally
        {

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
    public static void EUN(string ExcelFile, string OutPutFile, System.Data.DataTable DF, string g2, string g3, string gh, string gh2)
    {
        int gg = 0;
        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = true;

        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        //取得  Worksheet
        //Microsoft.Office.Interop.Excel.Range range1 = null;



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
        excelSheet.Name = "明細表";

        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;
        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;



            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[19, 3]);
            range.Select();
            range.Value2 = g3;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[19, 2]);
            range.Select();
            range.Value2 = g2;



            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[18, 1]);
            range.Select();
            range.Value2 = gh;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[19, 4]);
            range.Select();
            range.Value2 = gh2;

            for (int i = 0; i <= 4; i++)
            {
                string A = DF.Rows[i]["A"].ToString();
                string B = DF.Rows[i]["B"].ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[20 + i, 2]);
                range.Select();
                range.Value2 = A;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[20 + i, 3]);
                range.Select();
                range.Value2 = B;
            }


            for (int i = 5; i <= 9; i++)
            {
                string A = DF.Rows[i]["A"].ToString();
                string B = DF.Rows[i]["B"].ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[21 + i, 2]);
                range.Select();
                range.Value2 = A;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[21 + i, 3]);
                range.Select();
                range.Value2 = B;
            }

            for (int i = 10; i <= 14; i++)
            {
                string A = DF.Rows[i]["A"].ToString();
                string B = DF.Rows[i]["B"].ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[22 + i, 2]);
                range.Select();
                range.Value2 = A;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[22 + i, 3]);
                range.Select();
                range.Value2 = B;
            }







            excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        }
        finally
        {

            try
            {
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
    public static void ACC2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string YEAR)
    {
        int gg = 0;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 4]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (sTemp == "3rd parties sub total" || sTemp == "Related Parties sub total")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;


                        }
                        if (sTemp == "業務")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 5]);
                            range.Select();
                            range.Value2 = "總數量";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 6]);
                            range.Select();
                            range.Value2 = "總收入";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 7]);
                            range.Select();
                            range.Value2 = "總成本";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 8]);
                            range.Select();
                            range.Value2 = "總毛利";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 9]);
                            range.Select();
                            range.Value2 = "總毛利率";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 10]);
                            range.Select();
                            range.Value2 = "Sales%";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                        }


                    }


                    DetailRow++;


                }

            }




            //// //指定 CUT 的範圍
            //Cell_From = "A" + Convert.ToString(Line_Liab);
            //Cell_To = "D" + Convert.ToString(iRowCnt + 1);
            //excelSheet.get_Range(Cell_From, Cell_To).Cut(oMissing);
            //range.Select();


            object SelectCell_From = "H3";
            object SelectCell_To = "K" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "M3";
            SelectCell_To = "P" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "R3";
            SelectCell_To = "U" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "W3";
            SelectCell_To = "Z" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "AB3";
            SelectCell_To = "AE" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "AG3";
            SelectCell_To = "AJ" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "AL3";
            SelectCell_To = "AO" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "AQ3";
            SelectCell_To = "AT" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "AV3";
            SelectCell_To = "AY" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "BA3";
            SelectCell_To = "BD" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "BF3";
            SelectCell_To = "BI" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "BK3";
            SelectCell_To = "BN" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "BP3";
            SelectCell_To = "BS" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            DateTime df = DateTime.Now;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 2]);
            range.Select();
            range.Value2 = YEAR;

            string last;
            if (df.ToString("yyyy") == YEAR)
            {
                last = df.ToString("MMMM", new System.Globalization.DateTimeFormatInfo()).Substring(0, 3);
            }
            else
            {
                last = "Dec";
            }
            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 4]);
            range.Select();
            range.Value2 = "Jan~" + last;

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {

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
    public static void TEMP61(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string YEAR)
    {
        int gg = 0;
        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = true;

        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        //取得  Worksheet
        //Microsoft.Office.Interop.Excel.Range range1 = null;



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 6]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                    }


                    DetailRow++;


                }

            }

            object SelectCell_From = "G3";
            object SelectCell_To = "J" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "L3";
            SelectCell_To = "O" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "Q3";
            SelectCell_To = "T" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";


            SelectCell_From = "V3";
            SelectCell_To = "Y" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";

            SelectCell_From = "AA3";
            SelectCell_To = "AD" + Convert.ToString(DetailRow + 1);
            range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
            range.Select();
            range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";






        }
        finally
        {

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


    public static void DELETEFILE()
    {

        try
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp\\";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {


                File.Delete(file);

            }
        }
        catch { }
    }

    public static void DELETEFOLDER()
    {

        try
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp\\";
            string[] filenames = Directory.GetDirectories(OutPutFile);
            foreach (string file in filenames)
            {
                int G1 = file.IndexOf("rma");
                int G2 = file.IndexOf("rmar");
                int G3 = file.IndexOf("wh");

                if (G1 == -1 && G2 == -1 && G3 == -1)
                {
                    DirectoryInfo DIFO = new DirectoryInfo(file);

                    DIFO.Delete(true);
                }


            }
        }
        catch { }
    }
    private static void OutputData(System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet wsheet, int h1, int h2)
    {

        try
        {



            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 0; j <= dt.Columns.Count - 2; j++)
                {

                    try
                    {

                        if (dt.Columns[j].DataType == System.Type.GetType("System.String"))
                        {
                            wsheet.Cells[i + h1, j + h2] = (dt.Rows[i][j] == null) ? "" : "'" + dt.Rows[i][j];
                        }
                        else
                        {
                            wsheet.Cells[i + h1, j + h2] = (dt.Rows[i][j] == null) ? 0 : dt.Rows[i][j];
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);

                    }

                }

            }

        }

        catch (Exception ex1)
        {
            MessageBox.Show(ex1.Message);
        }

    }


    private static void OutputData1(System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet wsheet, int h1, int h2)
    {

        try
        {



            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    try
                    {

                        if (dt.Columns[j].DataType == System.Type.GetType("System.String"))
                        {
                            wsheet.Cells[i + h1, j + h2] = (dt.Rows[i][j] == null) ? "" : "'" + dt.Rows[i][j];
                        }
                        else
                        {
                            wsheet.Cells[i + h1, j + h2] = (dt.Rows[i][j] == null) ? 0 : dt.Rows[i][j];
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);

                    }

                }

            }

        }

        catch (Exception ex1)
        {
            MessageBox.Show(ex1.Message);
        }

    }

    public static void ExcelReportOutputODLN(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag, string P1, string P2, string P3)
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


        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;
            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 30, 368, 60, 50);
            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 193, 368, 60, 50);
            excelSheet.Shapes.AddPicture(P3, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 420, 368, 60, 50);
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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

            System.GC.WaitForPendingFinalizers();


            System.Diagnostics.Process.Start(OutPutFile);


        }

    }

    public static void ExcelReportOutpuwh(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag, System.Data.DataTable OrderData2, string SHIPPINGCODE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            if (flag == "Y")
            {
                System.Data.DataTable T1 = OrderData2;

                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {

                    string ITEMCODE = T1.Rows[i]["產品編號"].ToString();
                    string QTY = T1.Rows[i]["出貨數量"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11 + i, 4]);
                    range.Select();
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                  oMissing);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[11 + i, 4]);
                    range.Select();
                    range.Value2 = ITEMCODE + "共" + QTY + "片-" + SHIPPINGCODE;
                    range.Font.Size = 22;

                    object SelectCell_From = "D" + (11 + i);
                    object SelectCell_To = "J" + (11 + i);
                    range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                    range.Select();
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                }


            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

        }
        finally
        {
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

            //System.Diagnostics.Process.Start(OutPutFile);

        }

    }

    public static void GridViewToExcel(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();
        wapp.DisplayAlerts = false;
        wapp.Visible = true;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }
            Excel.Range range = wsheet.get_Range("A1", "Z1");
            range.Interior.ColorIndex = 6;
            range.Font.Bold = true;

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        
                        wsheet.Cells[i + 2, j + 1] = (cell.FormattedValue == null) ? "" : cell.FormattedValue.ToString();

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
    public static void GridViewToExcelSHARONS(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();
        wapp.DisplayAlerts = false;
        wapp.Visible = true;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.Worksheets[wbook.Sheets.Count];

        wsheet.Name = "工作表1";
        //decimal ss = Convert.ToDecimal(cell.Value.ToString());
        //wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : ss.ToString("#,##0");
        try
        {
            DataGridViewRow row2 = dgv.Rows[0];

            for (int j = 0; j < row2.Cells.Count; j++)
            {

                DataGridViewCell cell = row2.Cells[j];

                try
                {

                    wsheet.Cells[1, j + 1] = (cell.FormattedValue == null) ? "" : cell.FormattedValue.ToString();

                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);

                }

            }
            Excel.Range range = wsheet.get_Range("G1", "M1");
            range.Interior.ColorIndex = 6;
            range = wsheet.get_Range("A2", "W2");
            range.Interior.ColorIndex = 15;
            range.Font.Bold = true;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[2, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 1; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {

                        wsheet.Cells[i + 2, j + 1] = (cell.FormattedValue == null) ? "" : cell.FormattedValue.ToString();

                    }

                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);

                    }

                }
                if (String.IsNullOrEmpty(row.Cells[0].Value.ToString()))
                {
                    Excel.Range r = wsheet.get_Range("A" + (i + 2), "W" + (i + 2));
                    r.Interior.ColorIndex = 34;
                    r.Font.Bold = true;
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
    public static void GridViewToExcelES(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();
        wapp.DisplayAlerts = false;
        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            for (int i = 2; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1 - 2] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 2; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {

                        wsheet.Cells[i + 2, j + 1 - 2] = (cell.Value == null) ? "" : cell.Value.ToString();

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
    public static void GridViewToExcelNOTHEAD(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();
        wapp.DisplayAlerts = false;
        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {


            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
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
    public static void GridViewToExcelFAN(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();
        wapp.DisplayAlerts = false;
        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                string F1 = dgv.Columns[i].HeaderText;
                if (F1 != "品名" && F1 != "總計")
                {
                    F1 = dgv.Columns[i].HeaderText.Substring(4, 2);

                    wsheet.Cells[1, i + 1] = F1;
                }

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
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
    public static void GridViewToExcelDOUBLE(DataGridView dgv, DataGridView dgv2)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            for (int i = 0; i < dgv.Columns.Count - 2; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count - 2; j++)
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

            for (int i = 0; i < dgv2.Rows.Count; i++)
            {

                DataGridViewRow row2 = dgv2.Rows[i];

                for (int j = 0; j < row2.Cells.Count - 2; j++)
                {

                    DataGridViewCell cell2 = row2.Cells[j];

                    try
                    {
                        string F1 = (cell2.Value == null) ? "" : cell2.Value.ToString();
                        wsheet.Cells[i + 2 + dgv.Rows.Count, j + 1] = (cell2.Value == null) ? "" : cell2.Value.ToString();

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
    public static void GridViewToExceljo(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        if (j == 1)
                        {
                            decimal ss = Convert.ToDecimal(cell.Value.ToString());
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : ss.ToString("#,##0");
                        }
                        else
                        {
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                        }

                    }

                    catch (Exception ex)
                    {


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

    public static void GridViewToExcelSHARON2(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {
                    //if (i == 38)
                    //{
                    //    MessageBox.Show("AA");
                    //}

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        if (j > 5 && j < 10)
                        {
                            string dd = (cell.Value == null) ? "" : cell.Value.ToString();
                            if (!String.IsNullOrEmpty(dd))
                            {
                                decimal ss = Convert.ToDecimal(cell.Value.ToString());
                                wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : ss.ToString("#,##0");
                            }
                        }
                        else if (j == 16)
                        {
                            string dd = (cell.Value == null) ? "" : cell.Value.ToString();
                            if (!String.IsNullOrEmpty(dd))
                            {
                                decimal ss = Convert.ToDecimal(cell.Value.ToString());
                                wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : ss.ToString("#,##0.00");
                            }
                            else
                            {
                                wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                            }
                        }
                        //else if (j == 17)
                        //{
                        //    string dd = (cell.Value == null) ? "" : cell.Value.ToString();
                        //    if (!String.IsNullOrEmpty(dd))
                        //    {
                        //        decimal ss = Convert.ToDecimal(cell.Value.ToString());
                        //        wsheet.Cells[i + 2, j + 1] =  ss.ToString("#,##0.000");
                        //    }
                        //}
                        else
                        {
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                        }

                    }

                    catch (Exception ex)
                    {


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

    public static void GridViewToExcelPotato(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        if (j == 2 || j == 6)
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();
                        }
                        else
                        {
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                        }

                    }

                    catch (Exception ex)
                    {


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



    public static void GridViewToExcelSHARON(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        if (j == 0 || j == 4 || j == 6)
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();
                        }
                        else
                        {
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                        }

                    }

                    catch (Exception ex)
                    {


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
    public static void GridViewToExcelAP(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                DataGridViewRow row = dgv.Rows[i];

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    try
                    {
                        if (j == 10)
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();
                        }
                        else
                        {
                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();
                        }

                    }

                    catch (Exception ex)
                    {


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
    public static void GridViewToExcelSelect(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.SelectedRows.Count; i++)
            {

                DataGridViewRow row = dgv.SelectedRows[i];

                for (int j = 0; j < row.Cells.Count; j++)
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
    public static void GridViewToExcelSelectJOY(DataGridView dgv)
    {
        Microsoft.Office.Interop.Excel.Application wapp;

        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        Microsoft.Office.Interop.Excel.Workbook wbook;

        wapp = new Microsoft.Office.Interop.Excel.Application();

        wapp.Visible = false;

        wbook = wapp.Workbooks.Add(true);

        wsheet = (Excel.Worksheet)wbook.ActiveSheet;

        try
        {

            int iX;

            int iY;

            for (int i = 0; i < dgv.Columns.Count - 3; i++)
            {

                wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

            }

            for (int i = 0; i < dgv.SelectedRows.Count; i++)
            {

                DataGridViewRow row = dgv.SelectedRows[i];

                for (int j = 0; j < row.Cells.Count - 3; j++)
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

    public static void GridViewToCSV(DataGridView dgv, string FileName)
    {



        StringBuilder sbCSV = new StringBuilder();

        int intColCount = dgv.Columns.Count;





        //表頭

        for (int i = 0; i < dgv.Columns.Count; i++)
        {

            sbCSV.Append(dgv.Columns[i].HeaderText);



            if ((i + 1) != intColCount)
            {

                sbCSV.Append(",");

                //tab

                // sbCSV.Append("\t");

            }



        }

        sbCSV.Append("\n");



        foreach (DataGridViewRow dr in dgv.Rows)
        {



            //資料內容

            for (int x = 0; x < intColCount; x++)
            {



                if (dr.Cells[x].Value != null)
                {



                    sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));

                }

                else
                {

                    sbCSV.Append("");

                }





                if ((x + 1) != intColCount)
                {

                    sbCSV.Append(",");

                    // sbCSV.Append("\t");

                }

            }

            sbCSV.Append("\n");

        }

        using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.Default))
        {

            sw.Write(sbCSV.ToString());

        }



        System.Diagnostics.Process.Start(FileName);



    }
    public static void GridViewToCSVCAT(DataGridView dgv, string FileName)
    {



        StringBuilder sbCSV = new StringBuilder();

        int intColCount = dgv.Columns.Count;

        foreach (DataGridViewRow dr in dgv.Rows)
        {



            //資料內容

            for (int x = 0; x < intColCount; x++)
            {



                if (dr.Cells[x].Value != null)
                {



                    sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));

                }

                else
                {

                    sbCSV.Append("");

                }





                if ((x + 1) != intColCount)
                {

                    sbCSV.Append(",");


                }

            }

            sbCSV.Append("\n");

        }

        using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.UTF8))
        {

            sw.Write(sbCSV.ToString());

        }



        System.Diagnostics.Process.Start(FileName);



    }
    public static void GridViewToCSVCATPOTATO(DataGridView dgv, string FileName)
    {



        StringBuilder sbCSV = new StringBuilder();

        int intColCount = dgv.Columns.Count;

        foreach (DataGridViewRow dr in dgv.Rows)
        {



            //資料內容

            for (int x = 0; x < intColCount; x++)
            {



                if (dr.Cells[x].Value != null)
                {



                    sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));

                }

                else
                {

                    sbCSV.Append("");

                }





                if ((x + 1) != intColCount)
                {

                    sbCSV.Append(",");


                }

            }

            sbCSV.Append("\n");

        }

        using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.Default))
        {

            sw.Write(sbCSV.ToString());

        }



        System.Diagnostics.Process.Start(FileName);



    }
    public static void GridViewToCSVCATPOTATO2(DataGridView dgv, string FileName)
    {



        StringBuilder sbCSV = new StringBuilder();

        int intColCount = dgv.Columns.Count;

        foreach (DataGridViewRow dr in dgv.Rows)
        {

            for (int x = 0; x < intColCount; x++)
            {

                if (dr.Cells[x].Value != null)
                {

                    sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));

                }
                else
                {

                    sbCSV.Append("");

                }

                if ((x + 1) != intColCount)
                {

                    sbCSV.Append("#|#");
                }
                else
                {
                    sbCSV.Append("\r\n");
                }


            }



        }

        using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.UTF8))
        {

            sw.Write(sbCSV.ToString());

        }



        System.Diagnostics.Process.Start(FileName);



    }
    public static void GridViewToCSV2(DataGridView dgv, string FileName)
    {



        StringBuilder sbCSV = new StringBuilder();

        int intColCount = dgv.Columns.Count;





        //表頭

        for (int i = 0; i < dgv.Columns.Count; i++)
        {

            sbCSV.Append(dgv.Columns[i].HeaderText);



            if ((i + 1) != intColCount)
            {

                sbCSV.Append(",");

                //tab

                // sbCSV.Append("\t");

            }



        }

        sbCSV.Append("\n");



        foreach (DataGridViewRow dr in dgv.Rows)
        {



            //資料內容

            for (int x = 0; x < intColCount; x++)
            {



                if (dr.Cells[x].Value != null)
                {



                    sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));

                }

                else
                {

                    sbCSV.Append("");

                }





                if ((x + 1) != intColCount)
                {

                    sbCSV.Append(",");

                    // sbCSV.Append("\t");

                }

            }

            sbCSV.Append("\n");

        }

        using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.Default))
        {

            sw.Write(sbCSV.ToString());

        }



        System.Diagnostics.Process.Start(FileName);



    }
    public static void ODLNN(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, System.Data.DataTable EUNICE)
    {
        int gg = 0;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
        excelSheet.Name = "年度異常";

        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;
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

                        //取出欄位值 - 科目
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (sTemp == "小計")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                            range.Select();
                            range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                        }
                        if (sTemp == "總計")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                            range.Select();
                            range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 4]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";
                        range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    }


                    DetailRow++;


                }

            }

            DateTime before2month = DateTime.Now.AddMonths(-1);
            string dd = before2month.ToString("yyyy");
            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
            range.Select();
            range.Value2 = dd + "年";

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Name = "當月異常";
            excelSheet.Activate();

            for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
            {



                for (int iField = 1; iField <= iColCnt; iField++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    if (CheckSerial(EUNICE, sTemp, ref FieldValue))
                    {
                        range.Value2 = FieldValue;
                    }


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

                for (int aRow = 0; aRow <= EUNICE.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != EUNICE.Rows.Count - 1)
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
                        SetRow(EUNICE, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        if (sTemp == "小計")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                            range.Select();
                            range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                        }
                        if (sTemp == "總計")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.Select();
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 2]);
                            range.Select();
                            range.HorizontalAlignment = XlHAlign.xlHAlignRight;

                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 4]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        range.NumberFormatLocal = "#,##0_);[紅色](#,##0)";
                        range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    }


                    DetailRow++;


                }

            }


            int G1 = Convert.ToInt16(before2month.ToString("MM"));
            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
            range.Select();
            range.Value2 = dd + "--" + G1.ToString() + "月";

            excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        }
        finally
        {

            try
            {
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
    public static void ExcelReportOutputss(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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

    public static void ExcelReportOutputHR(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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



        }

    }

    public static void ExcelReportOutputHR104(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string DATE)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            for (int iRecord = iRowCnt2; iRecord >= 1; iRecord--)
            {


                //取出欄位值 - 科目
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                range.Select();
                sTemp = (string)range.Text;
                sTemp = sTemp.Trim();


                int TT = sTemp.IndexOf("公司別");
                int TT2 = sTemp.IndexOf("部門");
                if (TT != -1)
                {

                    object SelectCell_From = "A" + iRecord;
                    object SelectCell_To = "I" + iRecord;
                    range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                    range.Select();

                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    SelectCell_From = "A" + iRecord;
                    SelectCell_To = "B" + iRecord;
                    range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                    range.Select();
                    range.Merge(true);
                }

                if (TT2 != -1)
                {

                    range.Select();
                    sTemp = (string)range.Text;
                    range.Value2 = sTemp.ToString().Replace("部門", "");

                    object A_From = "A" + iRecord;
                    object A_To = "I" + iRecord;
                    range = excelSheet.get_Range(A_From, A_To);
                    range.Select();

                    range.Merge(true);
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                }
            }

            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                          oMissing);



            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
            range.Select();
            range.Value2 = DATE + "出勤異常人員";
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            range.Font.Size = 14;

            object SelectCell_From2 = "A1";
            object SelectCell_To2 = "D1";
            range = excelSheet.get_Range(SelectCell_From2, SelectCell_To2);
            range.Select();
            range.Merge(true);
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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


            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void ExcelReportOutputMAYTO(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, System.Data.DataTable W1)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


        excelApp.Visible = true;

        //Interop params
        object oMissing = System.Reflection.Missing.Value;

        //The Excel doc paths

        string excelFile = ExcelFile;

        object SelectCell = null;

        //Open the worksheet file
        Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        //取得  Worksheet
        //Microsoft.Office.Interop.Excel.Range range1 = null;



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

            for (int aRow = 0; aRow <= W1.Rows.Count - 1; aRow++)
            {
                if (aRow != W1.Rows.Count - 1)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22 + aRow + OrderData.Rows.Count, 1]);
                    range.EntireRow.Copy(oMissing);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
        oMissing);
                }

                for (int iField = 1; iField <= 4; iField++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22 + aRow + OrderData.Rows.Count, iField + 1]);
                    range.Select();

                    range.Value2 = W1.Rows[aRow][iField - 1];
                }

            }


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();
        }
        finally
        {


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
    public static void ODLNN2(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, System.Data.DataTable EUNICE, System.Data.DataTable EUNICE2, string T1)
    {
        int gg = 0;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;
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



            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Activate();
            System.Data.DataTable TH = EUNICE;

            for (int X = 2; X <= 13; X++)
            {

                string G1 = Convert.ToString(X - 1);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, X]);
                range.Select();
                range.Value2 = T1 + "/" + G1 + "/1";

                string J1 = TH.Rows[0][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, X]);
                range.Select();
                range.Value2 = J1;

                string J2 = TH.Rows[1][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[4, X]);
                range.Select();
                range.Value2 = J2;

                string J3 = TH.Rows[2][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[5, X]);
                range.Select();
                range.Value2 = J3;

                string J4 = TH.Rows[3][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[6, X]);
                range.Select();
                range.Value2 = J4;

                string J5 = TH.Rows[4][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[7, X]);
                range.Select();
                range.Value2 = J5;

                string J6 = TH.Rows[5][X - 2].ToString();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[8, X]);
                range.Select();
                range.Value2 = J6;

            }


            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
            excelSheet.Activate();
            System.Data.DataTable TH2 = EUNICE2;

            int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt2 = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range2 = null;

            string sTemp2 = string.Empty;
            string FieldValue2 = string.Empty;
            bool IsDetail2 = false;
            int DetailRow2 = 0;

            for (int iRecord2 = 1; iRecord2 <= iRowCnt2; iRecord2++)
            {

                for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
                {
                    range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord2, iField2]);
                    range2.Select();
                    sTemp2 = (string)range2.Text;
                    sTemp2 = sTemp2.Trim();

                    if (CheckSerial(TH2, sTemp2, ref FieldValue2))
                    {
                        range2.Value2 = FieldValue2;
                    }


                    if (IsDetailRow(sTemp2))
                    {
                        IsDetail2 = true;
                        DetailRow2 = iRecord2;
                        break;
                    }

                }

            }

            if (DetailRow2 != 0)
            {

                for (int aRow2 = 0; aRow2 <= TH2.Rows.Count - 1; aRow2++)
                {

                    //最後一筆不作
                    if (aRow2 != TH2.Rows.Count - 1)
                    {

                        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, 1]);
                        range2.EntireRow.Copy(oMissing);
                        range2.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField2 = 1; iField2 <= iColCnt2; iField2++)
                    {
                        range2 = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, iField2]);
                        range2.Select();
                        sTemp2 = (string)range2.Text;
                        sTemp2 = sTemp2.Trim();


                        FieldValue2 = "";
                        SetRow(TH2, aRow2, sTemp2, ref FieldValue2);

                        range2.Value2 = FieldValue2;

                    }


                    DetailRow2++;


                }

            }


            excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

        }
        finally
        {

            try
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
    public static void ExcelReportOutputFONTAI(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile)
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

                string BILL = "";
                int G1 = 0;
                for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow + 2, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    if (sTemp.Length == 10)
                    {

                        if (BILL != sTemp)
                        {

                            G1 = 1;
                        }
                        else
                        {
                            G1++;
                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow + 2, 3]);
                        range.Select();
                        range.Value2 = G1;

                        BILL = sTemp;
                    }
                }
            }





            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();


        }
        finally
        {


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
    public static void ExcelReportE1(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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


            //
            if (flag == "Y")
            {

                //取得
                //取得 Excel 的使用區域
                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange);
                range.Copy(oMissing);

                SelectCell = "A" + (iRowCnt + 2).ToString();
                range = excelSheet.get_Range(SelectCell, SelectCell);

                range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                    oMissing);

            }

            if (flag == "pivot")
            {
                //固定在第二頁
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                Microsoft.Office.Interop.Excel.PivotTable pivotTable = (Microsoft.Office.Interop.Excel.PivotTable)excelSheet.PivotTables("樞紐分析表3");

                pivotTable.RefreshTable();
            }
            else
            {
                SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);
                range.Select();
            }

        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void FIONAT(System.Data.DataTable OrderData, System.Data.DataTable OrderData2, string ExcelFile, string OutPutFile, string DDATE)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;


            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 1]);
            range.Select();

            range.Value2 = "終止日期：" + DDATE;

            for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
            {

                //最後一筆不作
                if (aRow != OrderData.Rows.Count - 1)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow + 6, 1]);
                    range.EntireRow.Copy(oMissing);

                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                        oMissing);
                }


                for (int iField = 1; iField <= iColCnt; iField++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow + 6, iField]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    FieldValue = "";
                    SetRow(OrderData, aRow, sTemp, ref FieldValue);

                    range.Value2 = FieldValue;


                }

                DetailRow++;
            }


            for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
            {

                //最後一筆不作
                if (aRow != OrderData2.Rows.Count - 1)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow + 8, 1]);
                    range.EntireRow.Copy(oMissing);

                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                        oMissing);
                }


                for (int iField = 1; iField <= iColCnt; iField++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow + 8, iField]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

                    FieldValue = "";
                    SetRow(OrderData2, aRow, sTemp, ref FieldValue);

                    range.Value2 = FieldValue;


                }

                DetailRow++;
            }

            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();



        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }



    public static void FIONA(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string COMPANY)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            int DetailRow = 0;
            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
            range.Select();
            range.Value2 = COMPANY;
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


            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();



        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }
    public static void FIONA2(System.Data.DataTable OrderData, System.Data.DataTable OrderData3, System.Data.DataTable OrderData4, System.Data.DataTable OrderData2, string ExcelFile, string OutPutFile)
    {

        //Create an Excel App
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        excelApp.DisplayAlerts = false;
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



        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

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

            int D1 = 19;

            if (DetailRow != 0)
            {

                for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    D1++;
                }





                for (int aRow = 0; aRow <= OrderData3.Rows.Count - 1; aRow++)
                {
                    if (OrderData.Rows.Count == 0)
                    {
                        D1++;
                    }

                    //最後一筆不作
                    if (aRow != OrderData3.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1 + 4, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1 + 4, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData3, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    D1++;
                }


                for (int aRow = 0; aRow <= OrderData4.Rows.Count - 1; aRow++)
                {
                    if (OrderData3.Rows.Count == 0)
                    {
                        D1++;
                    }
                    //最後一筆不作
                    if (aRow != OrderData4.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1 + 8, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[D1 + 8, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(OrderData4, aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    D1++;
                }
            }


            if (OrderData2.Rows.Count > 0)
            {
                string AAMT = OrderData2.Rows[0][1].ToString();
                string AREV = OrderData2.Rows[0][2].ToString();
                string BAMT = OrderData2.Rows[1][1].ToString();
                string BREV = OrderData2.Rows[1][2].ToString();
                string CAMT = OrderData2.Rows[2][1].ToString();
                string CREV = OrderData2.Rows[2][2].ToString();
                string DAMT = OrderData2.Rows[3][1].ToString();
                string DREV = OrderData2.Rows[3][2].ToString();

                string AAMT2 = OrderData2.Rows[4][1].ToString();
                string AREV2 = OrderData2.Rows[4][2].ToString();
                string BAMT2 = OrderData2.Rows[5][1].ToString();
                string BREV2 = OrderData2.Rows[5][2].ToString();
                string CAMT2 = OrderData2.Rows[6][1].ToString();
                string CREV2 = OrderData2.Rows[6][2].ToString();
                string DAMT2 = OrderData2.Rows[7][1].ToString();
                string DREV2 = OrderData2.Rows[7][2].ToString();

                string AAMT3 = OrderData2.Rows[8][1].ToString();
                string AREV3 = OrderData2.Rows[8][2].ToString();
                string BAMT3 = OrderData2.Rows[9][1].ToString();
                string BREV3 = OrderData2.Rows[9][2].ToString();
                string CAMT3 = OrderData2.Rows[10][1].ToString();
                string CREV3 = OrderData2.Rows[10][2].ToString();
                string DAMT3 = OrderData2.Rows[11][1].ToString();
                string DREV3 = OrderData2.Rows[11][2].ToString();

                string AAMT4 = OrderData2.Rows[12][1].ToString();
                string AREV4 = OrderData2.Rows[12][2].ToString();
                string BAMT4 = OrderData2.Rows[13][1].ToString();
                string BREV4 = OrderData2.Rows[13][2].ToString();
                string CAMT4 = OrderData2.Rows[14][1].ToString();
                string CREV4 = OrderData2.Rows[14][2].ToString();
                string DAMT4 = OrderData2.Rows[15][1].ToString();
                string DREV4 = OrderData2.Rows[15][2].ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 5]);
                range.Select();
                range.Value2 = AAMT;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 5]);
                range.Select();
                range.Value2 = AREV;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 6]);
                range.Select();
                range.Value2 = BAMT;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 6]);
                range.Select();
                range.Value2 = BREV;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 7]);
                range.Select();
                range.Value2 = CAMT;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 7]);
                range.Select();
                range.Value2 = CREV;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 8]);
                range.Select();
                range.Value2 = DAMT;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 8]);
                range.Select();
                range.Value2 = DREV;



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 5]);
                range.Select();
                range.Value2 = AAMT2;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 5]);
                range.Select();
                range.Value2 = AREV2;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 6]);
                range.Select();
                range.Value2 = BAMT2;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 6]);
                range.Select();
                range.Value2 = BREV2;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 7]);
                range.Select();
                range.Value2 = CAMT2;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 7]);
                range.Select();
                range.Value2 = CREV2;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 8]);
                range.Select();
                range.Value2 = DAMT2;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 8]);
                range.Select();
                range.Value2 = DREV2;



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 5]);
                range.Select();
                range.Value2 = AAMT3;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, 5]);
                range.Select();
                range.Value2 = AREV3;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 6]);
                range.Select();
                range.Value2 = BAMT3;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, 6]);
                range.Select();
                range.Value2 = BREV3;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 7]);
                range.Select();
                range.Value2 = CAMT3;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, 7]);
                range.Select();
                range.Value2 = CREV3;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[10, 8]);
                range.Select();
                range.Value2 = DAMT3;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[11, 8]);
                range.Select();
                range.Value2 = DREV3;



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 5]);
                range.Select();
                range.Value2 = AAMT4;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, 5]);
                range.Select();
                range.Value2 = AREV4;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 6]);
                range.Select();
                range.Value2 = BAMT4;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, 6]);
                range.Select();
                range.Value2 = BREV4;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 7]);
                range.Select();
                range.Value2 = CAMT4;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, 7]);
                range.Select();
                range.Value2 = CREV4;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 8]);
                range.Select();
                range.Value2 = DAMT4;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, 8]);
                range.Select();
                range.Value2 = DREV4;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 8]);
                range.Select();
                range.Value2 = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            }



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();



        }
        finally
        {


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



            System.Diagnostics.Process.Start(OutPutFile);

        }

    }


    public static void ExcelReportOutputP(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string P1, string P2, int X1, int X2, int Y1, int Y2, System.Data.DataTable AA)
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


        Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
        excelSheet.Activate();
        //  object SelectCell = "B10";
        //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


        //取得 Excel 的使用區域
        int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
        int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        // progressBar1.Maximum = iRowCnt;
        Microsoft.Office.Interop.Excel.Range range = null;


        //Microsoft.Office.Interop.Excel.Range FixedRange = null;


        try
        {

            string sTemp = string.Empty;
            string FieldValue = string.Empty;
            bool IsDetail = false;
            string FieldValue1 = string.Empty;
            int DetailRow = 0;
            //int r1 = Convert.ToInt16(textBox1.Text);
            //int r2 = Convert.ToInt16(textBox2.Text);
            excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 700, 115, 60, 50);
            excelSheet.Shapes.AddPicture(P2, Microsoft.Office.Core.MsoTriState.msoFalse,
     Microsoft.Office.Core.MsoTriState.msoTrue, 700, 350, 60, 50);
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

            //增加另一talbe處理

            System.Data.DataTable dtmark = AA;
            if (dtmark.Rows.Count != 0)
            {



                for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                {
                    //   sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) 未稅金額,ROUND(SUM(t1.linevat),0) 稅額,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCT
                    int DetailRow1 = DetailRow + 9 + a1Row;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                    // range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row]["科目"]);
                    range.Value2 = FieldValue1;
                    range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 2]);
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row]["科目名稱"]);
                    range.Value2 = FieldValue1;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 3]);
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row]["未稅金額"]);
                    range.Value2 = FieldValue1;

                    if (a1Row == dtmark.Rows.Count - 1)
                    {
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    }



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 4]);
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row]["稅額"]);
                    range.Value2 = FieldValue1;


                    if (a1Row == dtmark.Rows.Count - 1)
                    {
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 5]);
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row]["加總"]);
                    range.Value2 = FieldValue1;
                    range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                    if (a1Row == dtmark.Rows.Count - 1)
                    {
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    DetailRow1++;

                    if (a1Row == dtmark.Rows.Count - 1)
                    {
                        object SelectCell_From = "A" + (DetailRow1 - 1).ToString();
                        object SelectCell_To = "B" + (DetailRow1 - 1).ToString();
                        range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                        range.Select();
                        range.Merge(true);
                        range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                        range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                }

            }
            //



            SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            range.Select();

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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

            System.GC.WaitForPendingFinalizers();


            //     System.Diagnostics.Process.Start(OutPutFile);




        }

    }

}



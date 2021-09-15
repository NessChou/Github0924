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
    public partial class PROFIT : Form
    {
        string OutPutFile = "";
        public PROFIT()
        {
            InitializeComponent();
        }

        private void PROFIT_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            execuse();
        }

        private System.Data.DataTable Gen_201004()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT 'Q'+datename(quarter,T0.DOCDATE) Q,CASE WHEN  ");
            sb.Append("                    SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                           AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                    Substring (T1.[ItemCode],2,8) END +'.'+(Substring(T1.[ItemCode],12,1)) Model,CASE (Substring(T1.[ItemCode],11,1)) when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
            sb.Append("                     when '1' then 'P' when '2' then 'N' when '3' then 'V' when '4' then 'U' when '5' then 'N' ELSE 'X'END 等級,SUM(T1.QUANTITY) 數量,AVG(T1.PRICE) 進貨平均成本,MAX(T2.銷售單價) 銷貨平均成本,(MAX(T2.銷售單價)-AVG(T1.PRICE))/AVG(T1.PRICE) 毛利率  ");
            sb.Append("                    FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("                    LEFT JOIN  ");
            sb.Append("                    (");
            sb.Append("               SELECT Q,MODEL,VER,等級,AVG(銷售單價) 銷售單價 FROM ( ");
            sb.Append("               SELECT datename(quarter,T0.DOCDATE) Q,CASE WHEN  ");
            sb.Append("                    SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                           AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                    Substring (T1.[ItemCode],2,8) END Model ,(Substring(T1.[ItemCode],12,1)) VER,Substring(T1.[ItemCode],11,1) 等級,CASE DOCCUR WHEN 'USD' THEN  AVG(T1.PRICE) ELSE AVG(T1.PRICE)/30 END 銷售單價,DOCCUR ");
            sb.Append("                    FROM ORDR T0 LEFT JOIN RDR1 T1  ON (T0.DOCENTRY=T1.DOCENTRY) left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
            sb.Append("                     WHERE Convert(varchar(4),T0.[DocDate],112) = @YEAR  ");
            sb.Append("                    AND SUBSTRING(T1.ITEMCODE,15,1) <> '9' ");
            sb.Append("                    AND ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  AND T1.PRICE <> 0  ");
            sb.Append("                    and Substring(T1.[ItemCode],11,1) in ('0','1','5') ");
            sb.Append("               AND T1.LINESTATUS='C' AND T0.CANCELED = 'N' ");
            sb.Append("                    GROUP BY ");
            sb.Append("                     CASE WHEN  ");
            sb.Append("                    SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                           AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                    Substring (T1.[ItemCode],2,8) END ,(Substring(T1.[ItemCode],12,1)),Substring(T1.[ItemCode],11,1), ");
            sb.Append("                    datename(quarter,T0.DOCDATE),DOCCUR ) AS A ");
            sb.Append("               GROUP BY Q,MODEL,VER,等級 ");
            sb.Append("               ) T2 ON (datename(quarter,T0.DOCDATE)=T2.Q AND CASE WHEN  ");
            sb.Append("                    SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                           AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                    Substring (T1.[ItemCode],2,8) END =T2.MODEL AND (Substring(T1.[ItemCode],12,1))=T2.VER AND  ");
            sb.Append("                  Substring(T1.[ItemCode],11,1)=T2.等級) ");
            sb.Append("                    WHERE   ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            sb.Append("                    AND SUBSTRING(T0.CARDCODE,1,5)='S0001' ");
            sb.Append("                    and Substring(T1.[ItemCode],11,1) in('0','1','5') ");
            sb.Append("                    AND SUBSTRING(T0.CARDCODE,7,2) IN ('GD','TV','PID','DD','NB')  ");
            sb.Append("                    AND Convert(varchar(4),T0.[DocDate],112) = @YEAR  ");
            sb.Append("                    AND SUBSTRING(T1.ITEMCODE,15,1) <> 9  ");
            sb.Append("                   AND T1.PRICE <> 0 ");
            sb.Append("                    GROUP BY CASE WHEN  ");
            sb.Append("                    SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                            SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                           AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                    Substring (T1.[ItemCode],2,8) END ,(Substring(T1.[ItemCode],12,1)),Substring(T1.[ItemCode],11,1), ");
            sb.Append("                    datename(quarter,T0.DOCDATE) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.SelectedValue.ToString()));
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

                FileName = lsAppDir + "\\Excel\\OPCH\\PROFIT.xls";


                System.Data.DataTable OrderData = Gen_201004();


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N", "1");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void ExcelReportOutput(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string flag, string TT)
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



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Activate();
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = 7;

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
    }
}
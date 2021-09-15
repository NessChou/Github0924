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
    public partial class CHECKSHIP : Form
    {
        private System.Data.DataTable TempDt;
        private string FileName;
        public CHECKSHIP()
        {
            InitializeComponent();
        }
        private void WriteExcelProduct(string ExcelFile)
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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (!checkBox5.Checked)
            {
                if (iRowCnt > 1000)
                {
                    iRowCnt = 1000;
                }

            }


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string DATE;
                string SHIPCODE;
                int Qty = 0;

                DataRow dr;
                DataRow drFind;


                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

             
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    string G1 = range.Text.ToString().Trim();
                    if (!String.IsNullOrEmpty(G1))
                    {
                         DateTime  n;

                         if (DateTime.TryParse(G1, out n))
                         {


                             DATE = Convert.ToDateTime(range.Text.ToString().Trim()).ToString("yyyyMMdd");

                             range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                             range.Select();
                             SHIPCODE = range.Text.ToString().Trim();

                             if (!String.IsNullOrEmpty(SHIPCODE))
                             {
                                 string S1 = SHIPCODE.Substring(0, 2);
                                 if (S1 == "SH")
                                 {
                                     dr = TempDt.NewRow();
                                     dr["DATE"] = DATE;
                                     dr["船務工單"] = SHIPCODE;
                                     TempDt.Rows.Add(dr);
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
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("DATE", typeof(string));
            dt.Columns.Add("船務工單", typeof(string));

            return dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            TempDt = MakeTable();


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                WriteExcelProduct(FileName);
                dataGridView1.DataSource = TempDt;
            }
        }

        public void UPDATESHIP(string DATE, string SHIPPINGCODE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE SHIPPING_MAIN SET  buCardcode='Checked',buCardname=@DATE FROM SHIPPING_MAIN WHERE buCardcode <> 'Checked' AND SHIPPINGCODE=@SHIPPINGCODE  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("沒有資料匯入");
                    return;
                }

                DataGridViewRow row;
                for (int h = 0; h <= dataGridView1.Rows.Count - 1; h++)
                {

                    row = dataGridView1.Rows[h];



                    string DATE = row.Cells["DATE"].Value.ToString();
                    string SHIPCODE = row.Cells["船務工單"].Value.ToString();

                    UPDATESHIP(DATE, SHIPCODE);
                    
                }
            }
            catch { }
        }
    }
}

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
    public partial class ACCPAYACCOUNT : Form
    {
        private string FileName;
        public ACCPAYACCOUNT()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP(FileName,"支付通知單");

                dataGridView1.DataSource = GETACC("支付通知單");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP(FileName, "零用金");
                dataGridView2.DataSource = GETACC("零用金");
            }
        }
        public System.Data.DataTable GETACC(string DOCTYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  ACCTNAME 名稱,ACCTCODE 總帳科目  FROM ACME_OITTACC  WHERE DOCTYPE=@DOCTYPE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));

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
        public void DELTEMP(string DOCTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ACME_OITTACC WHERE  DOCTYPE=@DOCTYPE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));


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
        public void AddTEMP(string DOCTYPE, string ACCTCODE, string ACCTNAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ACME_OITTACC(DOCTYPE,ACCTCODE,ACCTNAME) values(@DOCTYPE,@ACCTCODE,@ACCTNAME)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@ACCTCODE", ACCTCODE));
            command.Parameters.Add(new SqlParameter("@ACCTNAME", ACCTNAME));

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

        private void WriteExcelAP(string ExcelFile, string DOCTYPE)
        {
          //  AddAP
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

     
            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
                range.Select();
                string S1 = range.Text.ToString().Trim();

                if (!String.IsNullOrEmpty(S1))
                {

                    DELTEMP(DOCTYPE);
                }


                string ACCOUNT = "";
                string ACCNAME = "";
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    ACCNAME = range.Text.ToString().Trim();

                               range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                               range.Select();
                               ACCOUNT = range.Text.ToString().Trim();

                               if (!String.IsNullOrEmpty(ACCOUNT))
                                {
                                    AddTEMP(DOCTYPE, ACCOUNT, ACCNAME);
                                }
                   // AddAP(DOCENTRY);
                }




            }
            finally
            {



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

        private void ACCPAYACCOUNT_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GETACC("支付通知單");
            dataGridView2.DataSource = GETACC("零用金");
            dataGridView3.DataSource = GETACC("國內外勤");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP(FileName, "國內外勤");
                dataGridView3.DataSource = GETACC("國內外勤");
            }
        }
    }
}

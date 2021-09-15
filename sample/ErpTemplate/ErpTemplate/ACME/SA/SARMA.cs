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
    public partial class SARMA : Form
    {
        public SARMA()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("請選擇檔案");
                return;
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("請選擇SHEET");
                return;
            }

            DELAP();
            WriteExcelAP(textBox1.Text);
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\SA\\J1.xlsx";

            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            System.Data.DataTable dt = ExecuteQuery();
            System.Data.DataTable dtCost = MakeTable();
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            DataRow dr = null;
            dr = dtCost.NewRow();
            dr["RMA"] = Convert.ToString(dt.Rows[0]["RMA"]);
            dr["LOCATION"] = Convert.ToString(dt.Rows[0]["LOCATION"]);
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string G = (i + 1).ToString() + ".";
                DataRow dd = dt.Rows[i];
                var c = (char)10;
                sb.Append(G + dd["NG"].ToString() + c);
                sb2.Append(G + dd["MEMO"].ToString() + c);

            }

            dr["NG"] = sb.ToString();
            dr["MEMO"] = sb2.ToString();
            dtCost.Rows.Add(dr);

            ExcelReport.ExcelReportOutput(dtCost, ExcelTemplate, OutPutFile, "N");
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("RMA", typeof(string));
            dt.Columns.Add("LOCATION", typeof(string));
            dt.Columns.Add("NG", typeof(string));
            dt.Columns.Add("MEMO", typeof(string));
            return dt;
        }
        private System.Data.DataTable ExecuteQuery()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT  RMA,LOCATION,NG,JUDGE+MEMO　MEMO FROM SA_J　　WHERE USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));


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
        private void WriteExcelAP(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);
            excelSheet.Activate();
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

                string LOCATION;
                string NG;
                string JUDGE;
                string MEMO;
                string NO;
                for (int iRecord = 3; iRecord <= iRowCnt; iRecord++)
                {

                            
                string sTemp = string.Empty;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);

                    sTemp = (string)range.Text;
                    string RMA = sTemp.Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    NO = range.Text.ToString().Trim().ToUpper();



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    LOCATION = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    NG = range.Text.ToString().Trim();

        
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    JUDGE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(LOCATION)&& NO !="NO.")
                    {
                        AddAP(RMA, LOCATION, NG, JUDGE, MEMO);
                    }
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
        public void AddAP(string RMA, string LOCATION, string NG, string JUDGE, string MEMO)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SA_J(RMA,LOCATION,NG,JUDGE,MEMO,USERS) values(@RMA,@LOCATION,@NG,@JUDGE,@MEMO,@USERS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RMA", RMA));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@NG", NG));
            command.Parameters.Add(new SqlParameter("@JUDGE", JUDGE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

        public void DELAP()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE SA_J WHERE USERS=@USERS", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        private void button1_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                comboBox1.SelectedIndex = -1;
                comboBox1.Items.Clear();
                FileName = openFileDialog1.FileName;
                this.textBox1.Text = FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                object oMissing = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                string Count_Sheet = excelBook.Sheets.Count.ToString();
                int i = excelBook.Sheets.Count;

                for (int xi = 1; xi <= i; xi++)
                {

                    Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                    string X1 = xi.ToString();
                    string X2 = excelsheet.Name.ToString();
                    string name_sheet = X1 + ":" + X2;
                    comboBox1.Items.Add(name_sheet);

                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                excelApp = null;
                excelBook = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


            }
        }
    }
}

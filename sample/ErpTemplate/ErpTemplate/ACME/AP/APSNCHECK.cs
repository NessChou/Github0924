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
    public partial class APSNCHECK : Form
    {
        public APSNCHECK()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                GD6(opdf.FileName);



     
            }
        }

        public void ADDCHECK(string LC, decimal AMT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_LCCHECK(LC,AMT,USERS) values(@LC,@AMT,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LC", LC));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
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
        public void TRUNCHECK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_LCCHECK WHERE USERS=@USERS", connection);
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

        private void GD6(string ExcelFile)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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

            progressBar1.Maximum = iRowCnt;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            System.Data.DataTable dtCost = MakeTableCombine();
            string F1 = "";
            string DF1 = "";
            string F2 = "";
            string F3 = "";
            string F4 = "";
            System.Data.DataTable DT = MakeTableCombine();
       
            for (int i = 2; i <= iRowCnt; i++)
            {

                progressBar1.Value = i;
                progressBar1.PerformStep();
      

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                F1 = range.Text.ToString().Trim();

                if (F1 != "")
                {
                    DF1 = F1;
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                F2 = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                F3 = range.Text.ToString().Trim();

                //try
                //{
                if (!String.IsNullOrEmpty(F3) || !String.IsNullOrEmpty(F2))
                {



                    DataRow dr = null;
                    //   //序號

                           
                            if (F1 == "")
                            {
                                F1 = DF1;
                            }

                            if (String.IsNullOrEmpty(F2))
                            {
                                F4 = "NG";
                            }
                            else if (String.IsNullOrEmpty(F3))
                            {
                                F4 = "NG";
                            }
                            else
                            {
                                string FF1 = F1.Substring(3, 1);
                                string FF2 = F2.Substring(14, 1);
                                string FF3 = F3.Substring(12, 1);

                                if (FF1 == FF2 && FF1 == FF3)
                                {
                                    F4 = "OK";
                                }
                                else
                                {
                                    F4 = "NG";
                                }
                            }

                            if (F4 == "NG")
                            {
                                dr = DT.NewRow();
                                dr["外箱號"] = F1;
                                dr["模組ID"] = F2;
                                dr["O/CID"] = F3;
                                dr["比對結果"] = F4;
                                DT.Rows.Add(dr);
                            }
                  
                }


                //}

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}

            }

            //if (RE == 0)
            //{
            //    MessageBox.Show("全部比較OK");
            //}
            dataGridView1.DataSource = DT;
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
        private System.Data.DataTable MakeTableCombine()
        {
         
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("外箱號", typeof(string));
            dt.Columns.Add("模組ID", typeof(string));
            dt.Columns.Add("O/CID", typeof(string));
            dt.Columns.Add("比對結果", typeof(string));
            return dt;
        }



    }
}

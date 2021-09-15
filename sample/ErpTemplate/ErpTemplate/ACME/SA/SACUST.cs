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
    public partial class SACUST : Form
    {
        public SACUST()
        {
            InitializeComponent();
        }

    


        public static System.Data.DataTable GetSACUST2(string CARDNAME)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM SA_CUST WHERE CARDNAME=@CARDNAME");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void GD5(string ExcelFile)
        {

            int h = 0;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

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



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string CARDNAME;
            string ADDRESS;
            string TEL;
            string PERSON;
            string EMAIL;
            string PAYMENT;
            string MARK;
            string MEMO;

            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                CARDNAME = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                ADDRESS = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                TEL = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                PERSON = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                EMAIL = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                PAYMENT = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                MARK = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                range.Select();
                MEMO = range.Text.ToString().Trim();

                //try
                //{
                    if (!String.IsNullOrEmpty(CARDNAME))
                    {

                        System.Data.DataTable GG1 = GetSACUST2(CARDNAME);
                        if (GG1.Rows.Count > 0)
                        {
                            UPDATESACUST(CARDNAME, ADDRESS, TEL, PERSON, EMAIL, PAYMENT, MARK, MEMO);
                        }
                        else
                        {
                            ADDSACUST(CARDNAME, ADDRESS, TEL, PERSON, EMAIL, PAYMENT, MARK, MEMO);

                        }

             
                    }


//                }

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}

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
        public void ADDSACUST(string CARDNAME, string ADDRESS, string TEL, string PERSON, string EMAIL, string PAYMENT, string MARK, string MEMO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SA_CUST(CARDNAME,ADDRESS,TEL,PERSON,EMAIL,PAYMENT,MARK,MEMO) values(@CARDNAME,@ADDRESS,@TEL,@PERSON,@EMAIL,@PAYMENT,@MARK,@MEMO)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@ADDRESS", ADDRESS));
            command.Parameters.Add(new SqlParameter("@TEL", TEL));
            command.Parameters.Add(new SqlParameter("@PERSON", PERSON));
            command.Parameters.Add(new SqlParameter("@EMAIL", EMAIL));
            command.Parameters.Add(new SqlParameter("@PAYMENT", PAYMENT));
            command.Parameters.Add(new SqlParameter("@MARK", MARK));
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
        public void UPDATESACUST(string CARDNAME, string ADDRESS, string TEL, string PERSON, string EMAIL, string PAYMENT, string MARK, string MEMO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE SA_CUST SET ADDRESS=@ADDRESS,TEL=TEL,PERSON=@PERSON,EMAIL=@EMAIL,PAYMENT=@PAYMENT,MARK=@MARK,MEMO=@MEMO WHERE CARDNAME=@CARDNAME ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@ADDRESS", ADDRESS));
            command.Parameters.Add(new SqlParameter("@TEL", TEL));
            command.Parameters.Add(new SqlParameter("@PERSON", PERSON));
            command.Parameters.Add(new SqlParameter("@EMAIL", EMAIL));
            command.Parameters.Add(new SqlParameter("@PAYMENT", PAYMENT));
            command.Parameters.Add(new SqlParameter("@MARK", MARK));
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

        private void SACUST_Load(object sender, EventArgs e)
        {

            this.sA_CUSTTableAdapter.Fill(this.sa.SA_CUST);

        }

 

        private void sA_CUSTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sA_CUSTBindingSource.EndEdit();
            this.sA_CUSTTableAdapter.Update(this.sa.SA_CUST);

            MessageBox.Show("更新成功");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            sA_CUSTTableAdapter.FillBy(sa.SA_CUST, textBox1.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
                                       OpenFileDialog opdf = new OpenFileDialog();
                           DialogResult result = opdf.ShowDialog();
                           if (opdf.FileName.ToString() == "")
                           {
                               MessageBox.Show("請選擇檔案");
                           }
                           else
                           {

                               GD5(opdf.FileName);


                           }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(sA_CUSTDataGridView);
        }
    }
}

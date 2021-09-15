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
using CarlosAg.ExcelXmlWriter;

namespace ACME
{
    public partial class GB_OITM2 : Form
    {
        string FileName;
        public GB_OITM2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                try
                {
                    TruncateTable();
                    FileName = openFileDialog1.FileName;
                    GetINVOICE(FileName);

                    dataGridView1.DataSource = GetOITM();
                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }


        private System.Data.DataTable GetOITM()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1 產品編碼,T2 產品名稱,T3 小分切 FROM GB_OITM2  ");

    
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


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
        private void TruncateTable()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("truncate table ACMESQLSP.DBO.GB_OITM2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


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
        private void GetINVOICE(string ExcelFile)
        {

            int N1 = 0;
            //Create an Excel App
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

            string NAME = excelSheet.Name;

        


            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id1;
            string id2;
            string id3;
            string id4;

            for (int i = 2; i <= iRowCnt; i++)
            {

           

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString().Trim();



                AddGBOITM2(id1, id2, id3);
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

        }


        public void AddGBOITM2(string T1, string T2, string T3)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_OITM2(T1,T2,T3) values(@T1,@T2,@T3)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@T1", T1));
            command.Parameters.Add(new SqlParameter("@T2", T2));
            command.Parameters.Add(new SqlParameter("@T3", T3));

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

        private void GB_OITM2_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOITM();
        }


   
    }
}

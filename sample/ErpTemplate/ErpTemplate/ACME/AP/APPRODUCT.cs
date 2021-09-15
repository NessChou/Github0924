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
    public partial class APPRODUCT : Form
    {
        public APPRODUCT()
        {
            InitializeComponent();
        }

        private void aP_PRODUCTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_PRODUCTBindingSource.EndEdit();
            this.aP_PRODUCTTableAdapter.Update(this.lC.AP_PRODUCT);

            MessageBox.Show("儲存成功");
        }

        private void APPRODUCT_Load(object sender, EventArgs e)
        {
            this.aP_PRODUCTTableAdapter.Fill(this.lC.AP_PRODUCT);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                    return;
                }
                else
                {

                    GetExcelContentGD44(opdf.FileName);

                    
                    this.aP_PRODUCTTableAdapter.Fill(this.lC.AP_PRODUCT);
                }

                MessageBox.Show("匯入成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            int i = excelBook.Sheets.Count;

            for (int xi = 1; xi <= i; xi++)
            {
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                string d = excelSheet.Name.Trim().ToString();
                int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                Hashtable ht = new Hashtable(iRowCnt);



                Microsoft.Office.Interop.Excel.Range range = null;



                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);


                string id;
                string id2;
                string id3;
                string id4;
                string id5;
                string id6;
                for (int b = 2; b <= iRowCnt; b++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 1]);
                    id = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 2]);
                    id2 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 3]);
                    id3 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 4]);
                    id4 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 5]);
                    id5 = range.Text.ToString();
                    
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[b, 6]);
                    id6 = range.Text.ToString();

                    AddAUOGD(id, id2, id3, id4, id5, id6);
              

                }


                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                range = null;
                excelSheet = null;
            }

            //Quit
            excelApp.Quit();


            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);



            excelApp = null;
            excelBook = null;


            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            MessageBox.Show("匯入成功");
        }
        public void AddAUOGD(string BU, string MODEL, string VER, string SITE, string PHASE, string STATUS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_PRODUCT(BU,MODEL,VER,SITE,PHASE,STATUS) values(@BU,@MODEL,@VER,@SITE,@PHASE,@STATUS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@SITE", SITE));
            command.Parameters.Add(new SqlParameter("@PHASE", PHASE));
            command.Parameters.Add(new SqlParameter("@STATUS", STATUS));

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

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(aP_PRODUCTDataGridView);
        }
    }
}
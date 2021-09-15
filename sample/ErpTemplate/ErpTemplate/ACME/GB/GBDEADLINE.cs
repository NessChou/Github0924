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
    public partial class GBDEADLINE : Form
    {
        string FileName = "";
        public GBDEADLINE()
        {
            InitializeComponent();
        }



        private void gB_DEADLINEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_DEADLINEBindingSource.EndEdit();
            this.gB_DEADLINETableAdapter.Update(this.pOTATO.GB_DEADLINE);
            pOTATO.GB_DEADLINE.AcceptChanges();
            MessageBox.Show("更新成功");

        }

        private void GBDEADLINE_Load(object sender, EventArgs e)
        {

            this.gB_DEADLINETableAdapter.Fill(this.pOTATO.GB_DEADLINE);

        }

        private void button1_Click(object sender, EventArgs e)
        {
                                    DialogResult result;
            result = MessageBox.Show("匯入時舊的資料會刪除，請確認是否要重新匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                TRUNG();
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    FileName = openFileDialog2.FileName;

                    GetINVOICE(FileName);
                }

                this.gB_DEADLINETableAdapter.Fill(this.pOTATO.GB_DEADLINE);

                MessageBox.Show("上傳完成");
            }
        }
        private void GetINVOICE(string ExcelFile)
        {


            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCntS = 2;
            int iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                for (int iRecord = iRowCntS; iRecord <= iRowCntE; iRecord++)
                {

                    string ITEMCODE;
                    string ITEMNAME;
                    decimal  QTY;
                    DateTime DEADLINE;


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    ITEMCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    ITEMNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    string DEAR = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    QTY = Convert.ToDecimal(range.Text.ToString().Trim());


                    if (IsDate(DEAR))
                    {
                        DEADLINE = Convert.ToDateTime(DEAR);
                        AddDEADLINE(ITEMCODE, ITEMNAME, DEADLINE,QTY);
                    }

                }
  




            }
            finally
            {


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



            }


            //  dataGridView1.DataSource = TempDt;

        }
        public bool IsDate(string DateString)
        {
            try
            {
                DateTime.Parse(DateString);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public void AddDEADLINE(string ITEMCODE, string ITEMNAME, DateTime DEADLINE,decimal  QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_DEADLINE(ITEMCODE,ITEMNAME,DEADLINE,QTY) values(@ITEMCODE,@ITEMNAME,@DEADLINE,@QTY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@DEADLINE", DEADLINE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public void TRUNG()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" TRUNCATE TABLE GB_DEADLINE ", connection);
            command.CommandType = CommandType.Text;


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

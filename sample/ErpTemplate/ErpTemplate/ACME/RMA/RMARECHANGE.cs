using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class RMARECHANGE : Form
    {
        private System.Data.DataTable TempDt;
        private string FileName;
        public RMARECHANGE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                TempDt = MakeTable();
                GetExcelProduct(FileName);
            }
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NO", typeof(string));
            dt.Columns.Add("倉別", typeof(string));


            return dt;
        }
        private void GetExcelProduct(string ExcelFile)
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


            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //第一行要
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();
                    //SERIAL_NO = SERIAL_NO.Substring(0, 7);
                    range.Select();

                    if (String.IsNullOrEmpty(SERIAL_NO) && iRecord == 1)
                    {

                        dr = TempDt.NewRow();
                        dr["NO"] = iRecord - 1;
                        dr["倉別"] = SERIAL_NO;

                        TempDt.Rows.Add(dr);

                    }

                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {

                        dr = TempDt.NewRow();
                        dr["NO"] = iRecord-1;
                        dr["倉別"] = SERIAL_NO;

                        TempDt.Rows.Add(dr);
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


            dataGridView1.DataSource = TempDt;
        }

        public void AddWH(int FIELDID,string INDEXID, string FLDVALUE, string DESCR)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO ACMESQL02.DBO.[UFD1]");
            sb.Append("            (TABLEID,FIELDID,INDEXID,FLDVALUE,DESCR)");
            sb.Append("      VALUES(@TABLEID,@FIELDID,@INDEXID,@FLDVALUE,@DESCR)");

            sb.Append(" INSERT INTO ACMESQL05.DBO.[UFD1]");
            sb.Append("            (TABLEID,FIELDID,INDEXID,FLDVALUE,DESCR)");
            sb.Append("      VALUES(@TABLEID,@FIELDID,@INDEXID,@FLDVALUE,@DESCR)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TABLEID", "OCTR"));
            command.Parameters.Add(new SqlParameter("@FIELDID", FIELDID));
            command.Parameters.Add(new SqlParameter("@INDEXID", INDEXID));
            command.Parameters.Add(new SqlParameter("@FLDVALUE", FLDVALUE));
            command.Parameters.Add(new SqlParameter("@DESCR", DESCR));
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
        public void DELWH(int  FIELDID)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE ACMESQL02.DBO.[UFD1] WHERE TABLEID='OCTR' AND FIELDID=@FIELDID ");
            sb.Append(" DELETE ACMESQL05.DBO.[UFD1] WHERE TABLEID='OCTR' AND FIELDID=@FIELDID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FIELDID", FIELDID));

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

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }

            DialogResult result;
            result = MessageBox.Show("請確定認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                //Repair Center
                DataGridViewRow row;
                int FIELD = 0;
                if (comboBox1.Text == "退運倉別")
                {
                    FIELD = 9;
                }
                if (comboBox1.Text == "Repair Center")
                {
                    FIELD = 17;
                }
                if (comboBox1.Text == "Vender")
                {
                    FIELD = 3;
                }
                if (comboBox1.Text == "單據總類")
                {
                    FIELD = 21;
                }
                DELWH(FIELD);

                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {
                    row = dataGridView1.Rows[i];
                    string a0 = row.Cells["NO"].Value.ToString();
                    string a1 = row.Cells["倉別"].Value.ToString();

                    AddWH(FIELD, a0, a1, a1);
                }

                MessageBox.Show("匯入成功，請重新開啟SAP");
            }
        }

        private void RMARECHANGE_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "退運倉別";
        }
    }
}

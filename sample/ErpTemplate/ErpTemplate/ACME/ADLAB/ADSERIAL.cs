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
    public partial class ADSERIAL : Form
    {
        string FileName;
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public ADSERIAL()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {



            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                try
                {

                    FileName = openFileDialog1.FileName;
                    GetINVOICE(FileName);

                    MessageBox.Show("上傳完成");
                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }


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

                string DOCDATE;
                string BILLNO;
                string PONO;
                string MODEL;
                string PRODNAME;
                string INVOICE;
                string DEADLINE;
                string SERNO;
                string CARDNAME;
                //第一行要
                for (int iRecord = iRowCntS; iRecord <= iRowCntE; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    DOCDATE = range.Text.ToString().Trim().Replace("/", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    BILLNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    PONO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    PRODNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    INVOICE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    DEADLINE = range.Text.ToString().Trim().Replace("/", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    SERNO = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(SERNO))
                    {
                        AddINVIN(DOCDATE, BILLNO, PONO, MODEL, PRODNAME, INVOICE, DEADLINE, SERNO, globals.UserID.ToString().Trim(), CARDNAME);
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

        public  System.Data.DataTable GetINVO()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ID,DOCDATE  出貨日期,BILLNO 正航單號,CARDNAME 客戶,PONO,MODEL,PRODNAME,INVOICE,Deadline,SERNO FROM ACMESQLSP.DBO.AD_SERNO ");
            sb.Append(" WHERE  1=1 ");
            if (textBox5.Text == "")
            {
                sb.Append(" AND  (DOCDATE BETWEEN @DATE1 AND @DATE2) ");
            }
            if (textBox1.Text != "")
            {
                sb.Append(" AND BILLNO=@BILLNO ");
            }
            if (textBox2.Text != "")
            {
                sb.Append(" AND MODEL=@MODEL ");
            }
            if (textBox5.Text != "")
            {
                sb.Append(" AND SERNO=@SERNO ");
            }
            if (textBox6.Text.ToString() != "")
            {
                sb.Append(" AND CARDNAME  like '%" + textBox6.Text.Trim() + "%'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox3.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox4.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@BILLNO", textBox1.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@MODEL", textBox2.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@SERNO", textBox5.Text.Trim()));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public System.Data.DataTable GetCUST(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn16);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select U.FullName  from ordBillMain T left join comCustomer U On  (U.ID=T.CustomerID AND U.Flag =1)   WHERE T.BillNO=@BillNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public void AddINVIN(string DOCDATE, string BILLNO, string PONO, string MODEL, string PRODNAME, string INVOICE, string DEADLINE, string SERNO, string NO, string CARDNAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ACMESQLSP.DBO.AD_SERNO(DOCDATE,BILLNO,PONO,MODEL,PRODNAME,INVOICE,DEADLINE,SERNO,NO,CARDNAME) values(@DOCDATE,@BILLNO,@PONO,@MODEL,@PRODNAME,@INVOICE,@DEADLINE,@SERNO,@NO,@CARDNAME)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@PONO", PONO));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@PRODNAME", PRODNAME));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@DEADLINE", DEADLINE));
            command.Parameters.Add(new SqlParameter("@SERNO", SERNO));
            command.Parameters.Add(new SqlParameter("@NO", NO));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            
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
            F1();
        }
        private void F1()
        {
        

            dataGridView1.DataSource = GetINVO();
        }
        private void ADSERIAL_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox4.Text = GetMenu.DLast();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇要刪除的列");
                return;
            }
                        DialogResult result;
            result = MessageBox.Show("請確認是否要刪除", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                int i = this.dataGridView1.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string ID = dataGridView1.SelectedRows[iRecs].Cells["ID"].Value.ToString();

                    DeletePacking(ID);

                }

                F1();
            }

        }
        private void DeletePacking(string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE ACMESQLSP.DBO.AD_SERNO WHERE ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ID", ID));



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

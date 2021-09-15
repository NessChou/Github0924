using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class GBFISH : Form
    {
        string FileName = "";
        public GBFISH()
        {
            InitializeComponent();
        }

        private void gB_FISHBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_FISHBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);

        }

        private void GBFISH_Load(object sender, EventArgs e)
        {
            toolStripTextBox1.Text = GetMenu.DFirst();
            toolStripTextBox2.Text = GetMenu.DLast();
            toolStripTextBox3.Text = GetMenu.DFirst();
            toolStripTextBox4.Text = GetMenu.DLast();
            toolStripTextBox5.Text = GetMenu.DFirst();
            toolStripTextBox6.Text = GetMenu.DLast();
            toolStripComboBox1.ComboBox.DataSource = GetOslp1();
            toolStripComboBox1.ComboBox.ValueMember = "DataValue";
            toolStripComboBox1.ComboBox.DisplayMember = "DataValue";
            toolStripComboBox1.Text = "";
            this.gB_FISHTableAdapter.FillBy(this.pOTATO.GB_FISH, toolStripTextBox1.Text, toolStripTextBox2.Text,toolStripComboBox1.Text);

        }
        public void AddGB_POTATOIEMAIN(string ShippingCode, string ORDDATE, string OrdName, string OrdTel, string OrdCom, string OrdEMail, string DelMan, string DelTel, string ProdName, string Qty, string Price, string FEE, string Amount, string DelAddr, string ARRDATE, string DELTIME, string REALDATE, string MAILNO, string TERM, string PAYMAN, string PAYDATE, string Remark)
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_FISH(ShippingCode,ORDDATE,OrdName,OrdTel,OrdCom,OrdEMail,DelMan,DelTel,ProdName,Qty,Price,FEE,Amount,DelAddr,ARRDATE,DELTIME,REALDATE,MAILNO,TERM,PAYMAN,PAYDATE,Remark) values (@ShippingCode,@ORDDATE,@OrdName,@OrdTel,@OrdCom,@OrdEMail,@DelMan,@DelTel,@ProdName,@Qty,@Price,@FEE,@Amount,@DelAddr,@ARRDATE,@DELTIME,@REALDATE,@MAILNO,@TERM,@PAYMAN,@PAYDATE,@Remark)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ORDDATE", ORDDATE));
            command.Parameters.Add(new SqlParameter("@OrdName", OrdName));
            command.Parameters.Add(new SqlParameter("@OrdTel", OrdTel));
            command.Parameters.Add(new SqlParameter("@OrdCom", OrdCom));
            command.Parameters.Add(new SqlParameter("@OrdEMail", OrdEMail));
            command.Parameters.Add(new SqlParameter("@DelMan", DelMan));
            command.Parameters.Add(new SqlParameter("@DelTel", DelTel));
            command.Parameters.Add(new SqlParameter("@ProdName", ProdName));
            command.Parameters.Add(new SqlParameter("@Qty", Qty));
            command.Parameters.Add(new SqlParameter("@Price", Price));
            command.Parameters.Add(new SqlParameter("@FEE", FEE));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@DelAddr", DelAddr));
            command.Parameters.Add(new SqlParameter("@ARRDATE", ARRDATE));
            command.Parameters.Add(new SqlParameter("@DELTIME", DELTIME));
            command.Parameters.Add(new SqlParameter("@REALDATE", REALDATE));
            command.Parameters.Add(new SqlParameter("@MAILNO", MAILNO));
            command.Parameters.Add(new SqlParameter("@TERM", TERM));
            command.Parameters.Add(new SqlParameter("@PAYMAN", PAYMAN));
            command.Parameters.Add(new SqlParameter("@PAYDATE", PAYDATE));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));


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
        public static string DDATE(string ORDDATE)
        {
            try
            {
                string YEAR = ORDDATE.Substring(0, 4);
                string MON = "";
                string DAY = "";
                int T1 = ORDDATE.IndexOf("/");
                int T2 = ORDDATE.LastIndexOf("/");
                if (T1 != -1)
                {
                    MON = ORDDATE.Substring(T1 + 1, T2 - T1 - 1);
                    DAY = ORDDATE.Substring(T2 + 1, ORDDATE.Length - T2 - 1);

                    if (MON.Length == 1)
                    {
                        MON = "0" + MON;
                    }

                    if (DAY.Length == 1)
                    {
                        DAY = "0" + DAY;
                    }
                }

                return YEAR + MON + DAY;
            }
            catch { return ""; }
        }

        public static System.Data.DataTable GetOslp1()
        {

            SqlConnection con = globals.Connection;
            string sql = "SELECT distinct rtrim(isnull(TERM,'')) DataValue FROM dbo.GB_FISH  ORDER BY rtrim(isnull(TERM,''))";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["oslp"];
        }
        private void GetExcelProduct3(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();




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
                string ShippingCode = "";
                string ORDDATE;
                string OrdName;
                string OrdTel;
                string OrdCom;
                string OrdEMail;
                string DelMan = "";
                string DelTel;
                string ProdName;
                string Qty;
                string Price;
                string FEE;
                string Amount;
                string DelAddr;
                string ARRDATE;
                string DELTIME;
                string REALDATE;
                string MAILNO;
                string TERM;
                string PAYMAN;
                string PAYDATE;
                string Remark;
                string DelMan2 = "";
            
           
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                
    
                    
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    ORDDATE = range.Text.ToString().Trim();
                    if (ORDDATE.Length > 7)
                    {
                        ORDDATE = DDATE(ORDDATE);
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    OrdName = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    OrdTel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    OrdCom = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    OrdEMail = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    DelMan = range.Text.ToString().Trim();

                    if (DelMan2 != DelMan || DelMan2 == "")
                    {

                        string F1 = Convert.ToString(Convert.ToInt16(DateTime.Now.Year) - 1911);
                        string NumberName = "H" + F1 + DateTime.Now.ToString("MMdd");
                        string AutoNum = util.GetAutoNumber(globals.Connection, NumberName);
                        ShippingCode = NumberName + AutoNum;
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    DelTel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    ProdName = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    Qty = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    Price = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    FEE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    Amount = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    DelAddr = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    ARRDATE = range.Text.ToString().Trim();
                    if (ARRDATE.Length > 7)
                    {
                        ARRDATE = DDATE(ARRDATE);
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    DELTIME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    REALDATE = range.Text.ToString().Trim();
                    if (REALDATE.Length > 7)
                    {
                        REALDATE = DDATE(REALDATE);
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    MAILNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 19]);
                    range.Select();
                    TERM = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 20]);
                    range.Select();
                    PAYMAN = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 21]);
                    range.Select();
                    PAYDATE = range.Text.ToString().Trim();
                    if (PAYDATE.Length > 7)
                    {
                        PAYDATE = DDATE(PAYDATE);
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 22]);
                    range.Select();
                    Remark = range.Text.ToString().Trim();

                   

                    if (!String.IsNullOrEmpty(ORDDATE))
                    {

                        DelMan2 = DelMan;
                        AddGB_POTATOIEMAIN(ShippingCode, ORDDATE, OrdName, OrdTel, OrdCom, OrdEMail, DelMan, DelTel, ProdName, Qty, Price, FEE, Amount, DelAddr, ARRDATE, DELTIME, REALDATE, MAILNO, TERM, PAYMAN, PAYDATE, Remark);

                         
               
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


        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                GetExcelProduct3(FileName);

                this.gB_FISHTableAdapter.FillBy(this.pOTATO.GB_FISH, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox1.Text);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.gB_FISHTableAdapter.FillBy(this.pOTATO.GB_FISH, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox1.Text);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.gB_FISHTableAdapter.FillBy1(this.pOTATO.GB_FISH, toolStripTextBox3.Text, toolStripTextBox4.Text, toolStripComboBox1.Text);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            this.gB_FISHTableAdapter.FillBy2(this.pOTATO.GB_FISH, toolStripTextBox5.Text, toolStripTextBox6.Text, toolStripComboBox1.Text);
        }
    }
}

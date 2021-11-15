using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAPbobsCOM;
using System.Collections;

namespace ACME
{
    public partial class Journal : Form
    {
        string FA = "acmesql98";
        public Journal()
        {
            InitializeComponent();
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("SAP").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "SAP");
                }
                else
                {
                    GetMenu.UP(t1, "SAP");
                }

                txbFilePath.Text = t1;
            }
        }
        private void ReadExcel(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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



            for (int i = 2; i <= iRowCnt; i++)
            {

                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string AcctCode;
                string LINE = "1";
                string AU97;
                string GRADE;
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK1;
                string REMARK2 = "";
                string P1;

                string ModelName;
                string AuoPN;
                string Grade;
                string FAB;
                string Price;
                string Remarks;
                string Apr;
                string Tcon;
                string Tqty;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                ModelName = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                Tcon = range.Text.ToString().Trim();

                if (ModelName != "" && ModelName != "MODEL_NAME")
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    AuoPN = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    Grade = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    FAB = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    Price = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                    range.Select();
                    Remarks = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    Apr = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    Tqty = range.Text.ToString().Trim();

                   
                    try
                    {

                        //AddOACT();

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
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
        private void btnImport_Click(object sender, EventArgs e)
        {
           

               
        }
        private bool AddOACT(System.Data.DataTable dt)

        {
            if (globals.UserID == "nesschou")
            {
                MessageBox.Show("確認是否為測試區");
            }

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

            SAPbobsCOM.Documents oPURCH = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);
            oPURCH.DocObjectCode = SAPbobsCOM.BoObjectTypes.oJournalVouchers;

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

            int i = 0; //  to be used as an index

            oCompany.CompanyDB = FA;
            oCompany.UserName = "A02";
            oCompany.Password = "6500";
            int result = oCompany.Connect();
            if (result == 0)
            {
                

                foreach (DataRow row in dt.Rows) 
                {
                    oPURCH.Lines.AccountCode = row["Account"].ToString();
                }
            }

            return true;
        }
    }
   
}

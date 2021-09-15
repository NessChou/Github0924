using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Globalization;
namespace ACME
{

    public partial class APOPEN : Form
    {
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn05 = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn23 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public APOPEN()
        {
            InitializeComponent();
        }

        public static System.Data.DataTable GetOPEN(string INVOICE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'ACME' 公司,DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSP.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'CHOICE',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPCHOICE.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'INFINITE',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPINFINITE.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'TOPGARDEN',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPTOPGARDEN.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'DRS',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPDRS.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '禾中',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPALL.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '宇豐',DOCENTRY,LINENUM,ITEMREMARK,ITEMCODE,Dscription,QUANTITY,INVOICE FROM AcmeSqlSPAD.DBO.WH_ITEM WHERE INVOICE=@INVOICE AND ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            //ITEMCODE
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

        public System.Data.DataTable GetOPEN2(string DOCENTRY, string ITEMCODE, string CHO)
        {
            SqlConnection connection = null;
            if (CHO == "ACME")
            {
                connection = new SqlConnection(strCn02);
            }
            if (CHO == "DRS")
            {
                connection = new SqlConnection(strCn05);
            }
     

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PRICE  FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetOPEN3(string BillNO, string PRODID, string CHO)
        {
            SqlConnection connection = null;
            if (CHO == "CHOICE")
            {
                connection = new SqlConnection(strCn21);
            }
            if (CHO == "INFINITE")
            {
                connection = new SqlConnection(strCn22);
            }
            if (CHO == "TOPGARDEN")
            {
                connection = new SqlConnection(strCn20);
            }

            if (CHO == "禾中")
            {
                connection = new SqlConnection(strCn23);
            }
            if (CHO == "宇豐")
            {
                connection = new SqlConnection(strCn16);
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PRICE FROM ordBillSub WHERE BillNO =@BillNO AND PRODID=@PRODID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("INVOICE", typeof(string));
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("單號", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(string));


            return dt;
        }
        private void WriteExcelAP(string ExcelFile)
        {
            System.Data.DataTable TempDt = MakeTable();

            DataRow drS = null;
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


            try
            {

                string INVOICE;
                string ITEMCODE;
                string DSCRIPTION;
                string QUANTITY;

                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    INVOICE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    DSCRIPTION = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    QUANTITY = range.Text.ToString().Trim();
                    drS = TempDt.NewRow();
                    drS["INVOICE"] = INVOICE;
                    drS["產品編號"] = ITEMCODE;
                    drS["品名規格"] = DSCRIPTION;
                    drS["數量"] = QUANTITY;
                    System.Data.DataTable G1 = GetOPEN(INVOICE, ITEMCODE);
                    if (G1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= G1.Rows.Count - 1; i++)
                        {
                            DataRow dd = G1.Rows[i];
                            string COMPANY = dd["公司"].ToString();
                            drS["公司"] = COMPANY;
                            string DOCENTRY = dd["DOCENTRY"].ToString();
                            drS["單號"] = DOCENTRY;
                            System.Data.DataTable G2 = null;
                            if (COMPANY == "ACME" || COMPANY == "DRS")
                            {
                                G2 = GetOPEN2(DOCENTRY, ITEMCODE, COMPANY);
                            }
                            else
                            {
                                G2 = GetOPEN3(DOCENTRY, ITEMCODE, COMPANY);
                            }
                            if (G2.Rows.Count > 0)
                            {

                                drS["單價"] = G2.Rows[0][0].ToString();
                            }
                        }
                    }

                    TempDt.Rows.Add(drS);

               

          
                }
                dataGridView1.DataSource = TempDt;



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

        private void button1_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;

                WriteExcelAP(FileName);
            }
        }
    }
}

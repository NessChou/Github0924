using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Transactions;
using System.Configuration;
using System.Net;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
namespace ACME
{
    public partial class RMASEARSZ : Form
    {
        public RMASEARSZ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\RMA\\深圳RMA費用統計表.xls";


            System.Data.DataTable OrderData = GetOrderData3INV();

            if (OrderData.Rows.Count > 0)
            {
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\深圳RMA費用統計表.xls";

                //產生 Excel Report
                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_深圳RMA費用統計表";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_深圳RMA費用統計表! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
             
            }
            else
            {
                MessageBox.Show("無資料");
            }


        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }

        private System.Data.DataTable GetOrderData3INV()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                         SELECT DISTINCT add10 SHIPTO,T1.RMANO,InvoiceNo_seq VER,T1.InQty QTY,''''+T1.VENDERNO VENDERNO,T1.RMANO,T0.add6 TOTAL, ");
            sb.Append("                                       createDate 發貨日期,receivePlace 起運地,T0.SHIPPINGCODE JOBNO,CodeName 客戶簡稱,pINO 運送方式, ");
            sb.Append("                                       shipment 運單號碼,buCntctPrsn 運費,cFS 付費方式,boatCompany Consignee,nTDollars 抵達日,''''+shipment 快遞單號,buCntctPrsn 費用,boardDeliver 計費重量  ");
            sb.Append("                                        ,ADD3 Gross,(SELECT TOP 1 CNO FROM Rma_PackingListDSZ TT WHERE TT.SHIPPINGCODE=T0.SHIPPINGCODE ORDER BY CNO DESC) 箱數");
            if (globals.DBNAME == "達睿生")
            {
                sb.Append("                                        ,(SELECT TOP 1 U_RMODEL FROM ACMESQL05.DBO.OCTR TS WHERE ISNULL(U_RMODEL,'') <> '' AND  TS.U_RMA_NO=T1.RMANO COLLATE  Chinese_Taiwan_Stroke_CI_AS) MODEL FROM dbo.Rma_mainSZ T0 ");
            }
            else
            {
                sb.Append("                                        ,(SELECT TOP 1 U_RMODEL FROM ACMESQL02.DBO.OCTR TS WHERE ISNULL(U_RMODEL,'') <> '' AND  TS.U_RMA_NO=T1.RMANO COLLATE  Chinese_Taiwan_Stroke_CI_AS) MODEL FROM dbo.Rma_mainSZ T0 ");
            }
                sb.Append("                                        LEFT JOIN dbo.Rma_INVOICEDSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE ) ");
            sb.Append(" WHERE SUBSTRING(T0.SHIPPINGCODE,4,8) BETWEEN @D1 AND @D2 AND ISNULL(pINO,'') <> '自提' ORDER BY T0.SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@D1", textBox1.Text));

            command.Parameters.Add(new SqlParameter("@D2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void RMASEARSZ_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                GetExcelContentGD44(opdf.FileName);
            }
        }
        private void GetExcelContentGD44(string ExcelFile)
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






            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);




            string id1;
            string id2;
            string id3;
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
                id3.Replace("￥", "");
                try
                {
                    if (!String.IsNullOrEmpty(id1))
                    {

                        UPDATE(id3, id2, id1);

                    }


                }

                catch (Exception ex)
                {
                    // MessageBox.Show(ex.Message);
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
            MessageBox.Show("匯出成功");
        }
        public void UPDATE(string buCntctPrsn, string boardDeliver, string shipment)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("update Rma_mainSZ set buCntctPrsn=@buCntctPrsn,boardDeliver=@boardDeliver where shipment=@shipment", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@buCntctPrsn", buCntctPrsn));
            command.Parameters.Add(new SqlParameter("@boardDeliver", boardDeliver));
            command.Parameters.Add(new SqlParameter("@shipment", shipment));
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
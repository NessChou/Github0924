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
    public partial class AUOENDDATE : Form
    {
        public AUOENDDATE()
        {
            InitializeComponent();
        }
        private void GetExcelContentGD5(string ExcelFile, string OutPutFile)
        {

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

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
       



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string id;

            string x1;
            string x2;
   
            int u1 = 0;
            int u2 = 0;

            for (int b = 1; b <= iColCnt; b++)
            {
              

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                    range.Select();
                    range.Columns.AutoFit();
                    id = range.Text.ToString();

                    if (id.ToUpper() == "INVOICE NO")
                    {
                        u1 = b;

                    }


                    if (id.ToUpper() == "DUE DATE")
                    {
                        u2 = b;

                    }
                    //PO No

            }

            if (u1.ToString() == "0" || u2.ToString() == "0")
            {
                MessageBox.Show("沒有Due Date/Invoice No欄位");
                return;
            }


            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, u1]);
                x1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, u2]);
                x2 = range.Text.ToString().Trim();

        

                string[] s = x1.Split('-');


                string H1 = s[0];
                try
                {
                    System.Data.DataTable T1 = GetData(x1);
                    int G1 = T1.Rows.Count;
                    if (G1 == 0)
                    {
                        MsgDocEntry.Text += string.Format("{0}--資料不存在" + "\r", x1);
                    }
                    //else if (G1 > 1)
                    //{
                    //    MsgDocEntry.Text += string.Format("{0}--資料有多筆" + "\r", x1);
                    //}
                    else
                    {
                        string DOCENTRY = T1.Rows[0][0].ToString();
                        if (!String.IsNullOrEmpty(x2))
                        {
                            UPDATE(x2, H1);
                        }

                    }

                }

                catch (Exception ex)
                {
                    MsgDocEntry.Text += string.Format("{0}--" + ex.Message + "\r", x1);
                   // MessageBox.Show(ex.Message);
                }
            }

            try
            {

                excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            }
            catch
            {
            }
            excelBook.Close(oMissing, oMissing, oMissing);
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
        public void UPDATE(string DOCDUEDATE, string U_ACME_INV)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE OPCH SET DOCDUEDATE=@DOCDUEDATE,U_ACME_SHIPMENT='1' WHERE (CARDCODE LIKE '%S0001%'  OR CARDCODE LIKE '%S0623%') AND U_ACME_INV like '%" + U_ACME_INV + "%' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDUEDATE", DOCDUEDATE));
            if (String.IsNullOrEmpty(DOCDUEDATE))
            {
                command.Parameters["@DOCDUEDATE"].Value = "";
            }


 

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
        //private System.Data.DataTable GetData(string aa, string bb, string cc)
        //{

        //    SqlConnection connection = globals.shipConnection;

        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(" SELECT t0.docentry FROM OPCH T0 ");
        //    sb.Append(" INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry ");
        //    sb.Append(" left join PDN1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='20')");
        //    sb.Append(" left join POR1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='20')");
        //    sb.Append(" WHERE T5.DOCENTRY=@aa");
        //    sb.Append(" AND CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
        //    sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
        //    sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
        //    sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
        //    sb.Append(" Substring (T1.[ItemCode],2,8) END  =@bb and t1.quantity=@cc");

        //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
        //    command.CommandType = CommandType.Text;

        //    command.Parameters.Add(new SqlParameter("@aa", aa));
        //    command.Parameters.Add(new SqlParameter("@bb", bb));
        //    command.Parameters.Add(new SqlParameter("@cc", cc));

        //    SqlDataAdapter da = new SqlDataAdapter(command);

        //    DataSet ds = new DataSet();
        //    try
        //    {
        //        connection.Open();
        //        da.Fill(ds, "auogd4");
        //    }
        //    finally
        //    {
        //        connection.Close();
        //    }

        //    return ds.Tables[0];

        //}
        private System.Data.DataTable GetData(string U_ACME_INV)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT t0.docentry FROM OPCH T0 WHERE  U_ACME_INV like '%" + U_ACME_INV + "%'  ");
          

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


     
         
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    string FILENAME = opdf.FileName;
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FILENAME);
                    MsgDocEntry.Text = "";
                    GetExcelContentGD5(FILENAME, OutPutFile);
                    MessageBox.Show("匯入成功");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
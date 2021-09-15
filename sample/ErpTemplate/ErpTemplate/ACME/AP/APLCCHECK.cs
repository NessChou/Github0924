using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class APLCCHECK : Form
    {
        public APLCCHECK()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                GD5(opdf.FileName);

            }
        }

        public void ADDCHECK(string LC, decimal AMT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_LCCHECK(LC,AMT,USERS) values(@LC,@AMT,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LC", LC));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public void TRUNCHECK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_LCCHECK WHERE USERS=@USERS", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        private void GD5(string ExcelFile)
        {
            TRUNCHECK();

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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
            if (iRowCnt > 1000)
            {
                iRowCnt = 1000;
            }
            progressBar1.Maximum = iRowCnt;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            string INV = "";
            string INVDATE = "";
            string INVAMT = "";
            string PAYMENT = "";
           System.Data.DataTable  DT = MakeTableCombine();
           System.Data.DataTable  DT2 = MakeTableCombine2();
           int F1 = 0;
           int F2 = 0;
           int F3 = 0;
           int F4 = 0;
           int RE = 0;
           int RE2 = 0;
           for (int i = 1; i <= iColCnt; i++)
           {
               range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, i]);
               range.Select();
               INV = range.Text.ToString().Trim().Replace(" ", "").ToUpper();
               if (INV == "INVOICENO")
               {
                   F1 = i;
               }
               if (INV == "INVOICEDATE")
               {
                   F2 = i;
               }

               if (INV == "TRXAMT")
               {
                   F3 = i;
               }

               if (INV == "PAYMENT")
               {
                   F4 = i;
               }
               //Payment
           }

           for (int i = 2; i <= iRowCnt; i++)
           {

               range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F1]);
               range.Select();
               INV = range.Text.ToString().Trim();
               int g = INV.IndexOf("-");
               if (g != -1)
               {
                   INV = INV.Substring(0, g);

               }


               if (!String.IsNullOrEmpty(INV))
               {

                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F3]);
                   range.Select();
                   INVAMT = range.Text.ToString().Trim().Replace(",", "");

                   System.Data.DataTable GINV = GetOPOR(INV);
         
                   if (GINV.Rows.Count > 0)
                   {
                       decimal n;
                       if (decimal.TryParse(INVAMT, out n))
                       {
                           ADDCHECK(INV, Convert.ToDecimal(INVAMT));

                       }
                   }

               }

           }


            for (int i = 2; i <= iRowCnt; i++)
            {

                progressBar1.Value = i;
                progressBar1.PerformStep();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F1]);
                range.Select();
                 INV  = range.Text.ToString().Trim();
                 if (INV == "X131671413")
                 {
                     MessageBox.Show("SA");
                 }
                 int g = INV.IndexOf("-");
                 if (g  != -1)
                 {
                     INV = INV.Substring(0, g);

                 }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F2]);
                range.Select();
             



                //try
                //{
                    if (!String.IsNullOrEmpty(INV))
                    {
                        DateTime G1 = Convert.ToDateTime(range.Text.ToString().Trim());
                        INVDATE = G1.ToString("yyyyMMdd");

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F3]);
                        range.Select();
                        INVAMT = range.Text.ToString().Trim().Replace(",", "");

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, F4]);
                        range.Select();
                        PAYMENT = range.Text.ToString().Trim();
                     
                        int PAY1 = PAYMENT.IndexOf("OA");
                        if (PAY1 == -1)
                        {
                            System.Data.DataTable GINV = GetOPOR(INV);
                            DataRow dr = null;
                            string INVD2 = "";
                            string INVAMT2 = "";
                            if (GINV.Rows.Count > 0)
                            {
                                decimal n;
                                if (decimal.TryParse(INVAMT, out n))
                                {

                                    INVD2 = GINV.Rows[0]["CargoDate"].ToString().Trim();
                                    INVAMT2 = GINV.Rows[0]["AMT"].ToString().Trim();
      
                                    dr = DT.NewRow();

                                    dr["InvoiceNo"] = INV;
                                    dr["Invoicedate"] = INVDATE;
                                    dr["TrxAmt"] = INVAMT;
                                    dr["LCNO"] = GINV.Rows[0]["LC"].ToString().Trim();
                                    dr["DOC"] = GINV.Rows[0]["DOC"].ToString().Trim();
                                    if (!decimal.TryParse(INVAMT2, out n))
                                    {
                                        dr["1比對結果"] = "NG";
                                        dr["2比對結果"] = "NG";
                                        dr["3比對結果"] = "NG";
                                    }
                                    else
                                    {
                                        dr["1比對結果"] = "OK";
                                        if (INVDATE != INVD2)
                                        {
                                            dr["2比對結果"] = "NG";
                                        }
                                        else
                                        {
                                            dr["2比對結果"] = "OK";
                                        }

                                        System.Data.DataTable GINV2 = GetOPOR2(INV, Convert.ToDecimal(INVAMT));
                                        System.Data.DataTable GINV3 = GetOPOR3(INV, Convert.ToDecimal(INVAMT));
                                        System.Data.DataTable GINV4 = GetOPOR4(INV, Convert.ToDecimal(INVAMT));
                                        System.Data.DataTable GINV33 = GetOPOR3(INV, Convert.ToDecimal(GetINV(INV).Rows[0][0].ToString()));
                                        if (GINV2.Rows.Count > 0 || GINV3.Rows.Count > 0 || GINV4.Rows.Count > 0 || GINV33.Rows.Count > 0)
                                        {
                                            dr["3比對結果"] = "OK";
                                        }
                                        else
                                        {
                                            dr["3比對結果"] = "NG";
                                        }

                                    }

                                    string R1 = dr["1比對結果"].ToString();
                                    string R2 = dr["2比對結果"].ToString();
                                    string R3 = dr["3比對結果"].ToString();

                                    if (R1 == "NG" || R2 == "NG" || R3 == "NG")
                                    {
                                        RE = 1;
                                        DT.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                        else
                        {

                            System.Data.DataTable GINV = GetOPOR5(INV);
                            DataRow dr = null;
             
                            if (GINV.Rows.Count > 0)
                            {
                                RE2 = 1;
                                    dr = DT2.NewRow();

                                    dr["InvoiceNo"] = INV;
                                    dr["LCNO"] = GINV.Rows[0]["LC"].ToString().Trim();
                                    dr["DOC"] = GINV.Rows[0]["DOC"].ToString().Trim();
                                    dr["1比對結果"] = "NG";

                                    DT2.Rows.Add(dr);
                            }



                        
                        }

                    }


                //}

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}

            }

            if (RE == 0)
            {
                MessageBox.Show("LC比較OK");
            }
            if (RE2 == 0)
            {
                MessageBox.Show("OA比較OK");
            }
            dataGridView1.DataSource = DT;
            dataGridView2.DataSource = DT2;
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


        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("1比對結果", typeof(string));
            dt.Columns.Add("Invoicedate", typeof(string));
            dt.Columns.Add("2比對結果", typeof(string));
            dt.Columns.Add("TrxAmt", typeof(string));
            dt.Columns.Add("3比對結果", typeof(string));
            dt.Columns.Add("LCNO", typeof(string));
            dt.Columns.Add("DOC", typeof(string));
            return dt;
        }


        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("1比對結果", typeof(string));
            dt.Columns.Add("LCNO", typeof(string));
            dt.Columns.Add("DOC", typeof(string));
            return dt;
        }
        public static System.Data.DataTable GetOPOR(string INVOCENO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT MAX(T1.LCNO) LC,MAX(T0.DocNum) DOC,MAX(CargoDate) CargoDate,SUM(AMT) AMT  FROM PLC1 T0  LEFT JOIN APLC T1 ON (T0.DocNum  =T1.DocNum) WHERE INVOCENO=@INVOCENO ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOCENO", INVOCENO));
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
        public static System.Data.DataTable GetOPOR5(string INVOCENO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT  (T1.LCNO) LC,(T0.DocNum) DOC   FROM PLC1 T0  LEFT JOIN APLC T1 ON (T0.DocNum  =T1.DocNum) WHERE INVOCENO=@INVOCENO AND  memo not like '%OA還款%' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOCENO", INVOCENO));
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
        public static System.Data.DataTable GetOPOR2(string INVOCENO, decimal AMT)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT (AMT) AMT  FROM PLC1 T0  LEFT JOIN APLC T1 ON (T0.DocNum  =T1.DocNum) WHERE INVOCENO=@INVOCENO  AND AMT=@AMT ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOCENO", INVOCENO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
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
        public static System.Data.DataTable GetOPOR3(string INVOCENO, decimal AMT)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT MAX(AMT)  FROM PLC1 T0  LEFT JOIN APLC T1 ON (T0.DocNum  =T1.DocNum) WHERE INVOCENO=@INVOCENO  HAVING SUM(AMT)=@AMT ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOCENO", INVOCENO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
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

        public static System.Data.DataTable GetINV(string INV)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT  SUM(AMT) AMT FROM AP_LCCHECK WHERE LC=@INV AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public static System.Data.DataTable GetOPOR4(string INVOCENO, decimal AMT)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT MAX(AMT)  FROM PLC1 T0  LEFT JOIN APLC T1 ON (T0.DocNum  =T1.DocNum) WHERE INVOCENO=@INVOCENO  GROUP BY DONNO HAVING SUM(AMT)=@AMT ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOCENO", INVOCENO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
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
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "LCNO")
            {
                string DOC = dataGridView1.CurrentRow.Cells["DOC"].Value.ToString();


                APLC a = new APLC();
                a.PublicString = DOC;
                a.Show();
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.Columns[e.ColumnIndex].Name == "LC")
            {
                string DOC = dataGridView2.CurrentRow.Cells["DOC2"].Value.ToString();


                APLC a = new APLC();
                a.PublicString = DOC;
                a.Show();
            }
        }

    }
}

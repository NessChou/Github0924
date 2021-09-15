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

    public partial class SERIAL : Form
    {
        int H = 0;
        int H2 = 0;
        int S1 = 0;
        string DOC = "";
        string CARD = "";
        private System.Data.DataTable TempDt2;
        public SERIAL()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{

             

            //    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            //    string OutPutFile = lsAppDir + "\\EXCEL2\\倉庫匯入\\";
            //    string[] filenames = Directory.GetFiles(OutPutFile);
            //    foreach (string file in filenames)
            //    {
            //        FileInfo info = new FileInfo(file);
            //        string NAME = info.Name;
            //        DOC = NAME.Substring(1, 5);

            //        int A1 = NAME.Length;
            //        CARD = NAME.Substring(6, A1 - 10);
            //        System.Data.DataTable t1 = GetOrderData22(DOC);
            
            //        //if (t1.Rows.Count > 0)
            //        //{
            //        //    MessageBox.Show(NAME + "資料已匯入過");
            //        //}
            //        //else
            //        //{
            //            GetINVTEST(file);

            //            MessageBox.Show(NAME + "匯入成功");
            //            File.Delete(file);
            //       // }
  
            //    }
            //    dataGridView1.DataSource = GetOrderData19();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void GetINVOICE(string ExcelFile)
        {
            S1 = 0;

            int N1 = 0;
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

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id1;
            string id2;
            string id3;
            string id4;
            string id5 = "";
            string id6 = "";
            string id7 = "";
            string id8 = "";
            string id9 = "";
            for (int i = 2; i <= iRowCnt; i++)
            {

                if (S1 == 1)
                {
                    return;
                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString().Trim();

                string A = id2.Substring(9, 1);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id4 = range.Text.ToString().Trim();

                string ITEMCODE = id1;
                System.Data.DataTable T1 = GetMODEL(id1);
                System.Data.DataTable T2 = GetINV(id3);
                if (T1.Rows.Count > 0)
                {

                    string MODEL = T1.Rows[0]["MODEL"].ToString();
                    string VER = T1.Rows[0]["VER"].ToString();
                    string ItemCode = T1.Rows[0]["ItemCode"].ToString();
                    id5 = MODEL;
                    id6 = VER;
                    id8 = ItemCode;
                }

                if (T2.Rows.Count > 0)
                {
                    string INV = T2.Rows[0]["INVDATE"].ToString();
                    string CARD = T2.Rows[0]["客戶編號"].ToString();
                    id7 = INV;
                    id9 = CARD;
                }

                try
                {
                    //if (A == "B" || A == "C" || A == "D")
                    //{
                        AddINVOUT(id2, id1, id3, id4, id5, id6, id7, id8, id9);
                   // }
     
                }

                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                 
                }
            }

            //if (N1 == 1)
            //{
            //    NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
            //  "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";

                //try
                //{
                //    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //}
                //catch
                //{
                //}
            //}
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
        private void GetINVTEST(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

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


            string id;
            string idx = "";
            string id1x = "";
            string id2x = "";
            string id3x = "";
            string id2 = "";
            string id3 = "";
            string id4 = "";
            string id5 = "";
            string id6 = "";
            string id7 = "";
            string id8 = "";

            string A = "";
            int N1 = 0;
            StringBuilder sb = new StringBuilder();
        

            for (int b = 1; b <= iColCnt; b++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                range.Select();
                id = range.Text.ToString();

                if (id.ToUpper() == "數量")
                {
                

                    if (b == 6)
                    {
                        A = "1";
                    }
                    else if (b == 10)
                    {
                        A = "2";
                    }
                    else
                    {
                        MessageBox.Show("格式不符");
                        N1 = 1;
                        return;
                    }
                }

                //PO No

            }
            if (N1 != 1)
            {
                if (A == "1")
                {
                    for (int c = 2; c <= iRowCnt; c++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
                        range.Select();
                        id8 = range.Text.ToString().Trim();
                        id8 = id8.Replace(".", "");
                        id8 = id8.Replace("/", "");


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 2]);
                        range.Select();
                        id = range.Text.ToString().Trim();
                        if (id != "")
                        {
                            id1x = id;
                        }
                        if (id == "")
                        {
                            id = id1x;
                        }

                        string ITEMNAME = "";
                        System.Data.DataTable k1 = GetITEMNAME(id);
                        if (k1.Rows.Count > 0)
                        {
                            ITEMNAME = k1.Rows[0][0].ToString();
                        }
                        //4
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 3]);
                        range.Select();
                        id2 = range.Text.ToString().Trim();
                        if (id2 != "")
                        {
                            id2x = id2;
                        }
                        if (id2 == "")
                        {
                            id2 = id2x;
                        }

                        //6
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 4]);
                        range.Select();
                        id3 = range.Text.ToString().Trim();
                        if (id3 != "")
                        {
                            id3x = id3;
                        }
                        if (id3 == "")
                        {
                            id3 = id3x;
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 5]);
                        range.Select();
                        id4 = range.Text.ToString().Trim();
                        id4 = id4.Replace(",", "");
                        id4 = id4.Replace("*", "");
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 6]);
                        range.Select();
                        id5 = range.Text.ToString().Trim();


                        int n;




                        if (!String.IsNullOrEmpty(id4))
                        {
                            if (int.TryParse(id5, out n))
                            {
                                AddINVIN(id, ITEMNAME, id2, id3, id4, "", id5, CARD, DOC, id8, DateTime.Now.ToString("yyyyMMdd"));
                            }
                        }
                    }


                }



                if (A == "2")
                {
                    for (int c = 2; c <= iRowCnt; c++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
                        range.Select();
                        id8 = range.Text.ToString().Trim();
                        id8 = id8.Replace(".", "");
                        id8 = id8.Replace("/", "");
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 2]);
                        range.Select();
                        id = range.Text.ToString().Trim();
                        if (id != "")
                        {
                            id1x = id;
                        }
                        if (id == "")
                        {
                            id = id1x;
                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 3]);
                        range.Select();
                        id2 = range.Text.ToString().Trim();
                        if (id2 != "")
                        {
                            id2x = id2;
                        }
                        if (id2 == "")
                        {
                            id2 = id2x;
                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 4]);
                        range.Select();
                        id3 = range.Text.ToString().Trim();
                        if (id3 != "")
                        {
                            id3x = id3;
                        }
                        if (id3 == "")
                        {
                            id3 = id3x;
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 6]);
                        range.Select();
                        id4 = range.Text.ToString().Trim();
                        if (id4 != "")
                        {
                            idx = id4;
                        }
                        if (id4 == "")
                        {
                            id4 = idx;
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 8]);
                        range.Select();
                        id5 = range.Text.ToString().Trim();
                        id5 = id5.Replace("*", "");

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 9]);
                        range.Select();
                        id6 = range.Text.ToString().Trim();


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[c, 10]);
                        range.Select();
                        id7 = range.Text.ToString().Trim();

                        int n;




                        if (!String.IsNullOrEmpty(id5) || !String.IsNullOrEmpty(id6))
                        {
                            if (int.TryParse(id7, out n))
                            {
                                AddINVIN(id, id2, id3, id4, id5, id6, id7, CARD, DOC, id8, DateTime.Now.ToString("yyyyMMdd"));
                            }
                        }
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
        public void AddINVIN(string ITEMCODE, string ITEMNAME, string PARTNO, string INV, string CARTON, string PIC, string QTY, string CARD, string SAPDOC, string DOCDATE, string INSERTDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_INVOICEIN(ITEMCODE,ITEMNAME,PARTNO,INV,CARTON,PIC,QTY,CARD,SAPDOC,DOCDATE,INSERTDATE) values(@ITEMCODE,@ITEMNAME,@PARTNO,@INV,@CARTON,@PIC,@QTY,@CARD,@SAPDOC,@DOCDATE,@INSERTDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@PIC", PIC));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@INSERTDATE", INSERTDATE));
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

        public void AddINVOUT(string SHIPPING, string PART, string INVOICE, string CARTON, string MODEL, string VER, string INVDATE,string ITEMCODE,string CARDCODE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_INVOICEOUT(SHIPPING,PART,INVOICE,CARTON,INSERTDATE,MODEL,VER,INVDATE,ITEMCODE,CARDCODE) values(@SHIPPING,@PART,@INVOICE,@CARTON,@INSERTDATE,@MODEL,@VER,@INVDATE,@ITEMCODE,@CARDCODE)", connection);
            command.CommandType = CommandType.Text;
           
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));
            command.Parameters.Add(new SqlParameter("@PART", PART));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@INSERTDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@INVDATE", INVDATE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void UPINVOUT(string SHIPPING, string PART, string INVOICE, string CARTON, string MODEL, string VER, string INVDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_INVOICEOUT SET PART=@PART,INVOICE=@INVOICE,CARTON=@CARTON,MODEL=@MODEL,VER=@VER,INVDATE=@INVDATE WHERE SHIPPING=@SHIPPING ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));
            command.Parameters.Add(new SqlParameter("@PART", PART));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@INVDATE", INVDATE));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private System.Data.DataTable GetOrderData22(string doc)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   select * from AP_INVOICEIN where sapdoc=@doc");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@doc", doc));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOrderData20();
            dataGridView2.DataSource = GetOrderData20AP();
            dataGridView3.DataSource = GetOrderData21AP();
        }

        private System.Data.DataTable GetOrderData19()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select ITEMCODE,ITEMNAME,PARTNO,INV,CARTON,PIC,QTY,CARD,SAPDOC,DOCDATE from ap_invoicein where insertdate=@insertdate ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData19AP()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select INVDATE,SHIPPING,PART,INVOICE,CARTON,MODEL,VER from ap_invoiceOUT where insertdate=@insertdate ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public void AddINVIN(string ITEMCODE, string ITEMNAME, string PARTNO, string INV, string CARTON, string PIC, string QTY, string CARD, string SAPDOC, string DOCDATE, string INSERTDATE, string WAREHOUSE, string WHNO, string DDATE, string INOUT, string FILE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into AP_INVOICEIN(ITEMCODE,ITEMNAME,PARTNO,INV,CARTON,PIC,QTY,CARD,SAPDOC,DOCDATE,INSERTDATE,WAREHOUSE,WHNO,DDATE,INOUT) values(@ITEMCODE,@ITEMNAME,@PARTNO,@INV,@CARTON,@PIC,@QTY,@CARD,@SAPDOC,@DOCDATE,@INSERTDATE,@WAREHOUSE,@WHNO,@DDATE,@INOUT)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@PIC", PIC));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@INSERTDATE", INSERTDATE));
            command.Parameters.Add(new SqlParameter("@WAREHOUSE", WAREHOUSE));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@DDATE", DDATE));
            command.Parameters.Add(new SqlParameter("@INOUT", INOUT));
            //INOUT
            //DDATE
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + FILE);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private System.Data.DataTable GetAP_INVOICEIN(string ITEMNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_INVOICEIN WHERE ITEMNAME=@ITEMNAME");
   


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSERREPORT()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select INVOICE,SUM(片數) 片數,客戶,產出年,產出月,出貨年月 from (");
            sb.Append(" SELECT T0.INVOICE,COUNT(*) 片數,CASE ISNULL(MAX(T1.CARD),'') WHEN '' THEN MAX(T2.CARD) ELSE MAX(T1.CARD) END 客戶");
            sb.Append(" ,'20'+substring(T0.carton,7,2) 產出年 ,case substring(T0.carton,9,1)  ");
            sb.Append("             when 'A' THEN 10 WHEN 'B' THEN 11 WHEN 'C' THEN 12 ELSE  substring(T0.carton,9,1) END 產出月");
            sb.Append(" ,SUBSTRING(CASE ISNULL(MAX(T1.DOCDATE),'') WHEN '' THEN MAX(T2.DOCDATE) ELSE MAX(T1.DOCDATE) END,1,6) 出貨年月  FROM AP_INVOICEOUT T0");
            sb.Append(" LEFT JOIN (SELECT DISTINCT CARTON,CARD,DOCDATE FROM AP_INVOICEIN ) T1 ON ( T0.CARTON=T1.CARTON)");
            sb.Append(" LEFT JOIN (SELECT DISTINCT PIC,CARD,DOCDATE FROM AP_INVOICEIN ");
            sb.Append(" ) T2 ON ( T0.SHIPPING=T2.PIC)");
            sb.Append("  WHERE INSERTDATE between @AA and @BB and  part='97.15G05.300' ");
            sb.Append(" GROUP BY T0.INVOICE,T0.CARTON ) as a");
            sb.Append(" GROUP BY INVOICE,客戶,產出年,產出月,出貨年月");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox8.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSERREPORT2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT 客戶,SUM(片數) 片數 FROM ( select INVOICE,SUM(片數) 片數,客戶,產出年,產出月,出貨年月 from (");
            sb.Append("           SELECT T0.INVOICE,COUNT(*) 片數,CASE ISNULL(MAX(T1.CARD),'') WHEN '' THEN MAX(T2.CARD) ELSE MAX(T1.CARD) END 客戶");
            sb.Append("           ,'20'+substring(T0.carton,7,2) 產出年 ,case substring(T0.carton,9,1)  ");
            sb.Append("                       when 'A' THEN 10 WHEN 'B' THEN 11 WHEN 'C' THEN 12 ELSE  substring(T0.carton,9,1) END 產出月");
            sb.Append("           ,SUBSTRING(CASE ISNULL(MAX(T1.DOCDATE),'') WHEN '' THEN MAX(T2.DOCDATE) ELSE MAX(T1.DOCDATE) END,1,6) 出貨年月  FROM AP_INVOICEOUT T0");
            sb.Append("           LEFT JOIN (SELECT DISTINCT CARTON,CARD,DOCDATE FROM AP_INVOICEIN ) T1 ON ( T0.CARTON=T1.CARTON)");
            sb.Append("           LEFT JOIN (SELECT DISTINCT PIC,CARD,DOCDATE FROM AP_INVOICEIN ");
            sb.Append("           ) T2 ON ( T0.SHIPPING=T2.PIC)");
            sb.Append("            WHERE INSERTDATE between @AA and @BB and  part='97.15G05.300' ");
            sb.Append("           GROUP BY T0.INVOICE,T0.CARTON ) as a");
            sb.Append("           GROUP BY INVOICE,客戶,產出年,產出月,出貨年月) AS A");
            sb.Append("     GROUP BY 客戶");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox8.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData20()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select '出貨' INOUT,DOCDATE,CARD,ITEMCODE,QTY,PARTNO,INV,CARTON,PIC,WAREHOUSE from ap_invoicein where 1=1 ");

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("  and DOCDATE between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (comboBox1.SelectedValue.ToString() != "")
            {
                sb.Append(" and  ITEMCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append(" and  CARD like '%" + textBox3.Text.ToString() + "%'  ");
            }
            if (comboBox2.SelectedValue.ToString() != "")
            {
                sb.Append(" and  Substring(ITEMCODE,12,1)  like '%" + comboBox2.SelectedValue.ToString() + "%'  ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" and  INV like '%" + textBox4.Text.ToString() + "%'  ");
            }

            if (textBox5.Text != "")
            {
                sb.Append(" and  PIC = '" + textBox5.Text.ToString() + "'  ");
            }

            if (textBox6.Text != "")
            {
                sb.Append(" and  carton = '" + textBox6.Text.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData20AP()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select invDATE,SHIPPING,PART,INVOICE,CARTON,MODEL,VER,CARDCODE from ap_invoiceOUT where 1=1 ");

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("  and INSERTDATE between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (comboBox1.SelectedValue.ToString() != "")
            {
                sb.Append(" and  MODEL = '" + comboBox1.SelectedValue.ToString() + "'  ");
            }
  
            if (comboBox2.SelectedValue.ToString() != "")
            {
                sb.Append(" and VER   = '" + comboBox2.SelectedValue.ToString() + "'   ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" and  INVOICE like '%" + textBox4.Text.ToString() + "%'  ");
            }
            if (textBox5.Text != "")
            {
                sb.Append(" and  SHIPPING = '" + textBox5.Text.ToString() + "'  ");
            }

            if (textBox6.Text != "")
            {
                sb.Append(" and  carton = '" + textBox6.Text.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }



        private System.Data.DataTable GetOrderData21AP()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.DOCDATE 出貨日期,SAPDOC 交貨單號 ,CARD 客戶,T0.ITEMCODE 產品編號,T0.ITEMNAME 品名規格");
            sb.Append("     ,T0.QTY 數量,t1.cardcode 供應商,INVDATE,T1.INVOICE '原廠inv#',PARTNO 料號,CASE ISNULL(T0.CARTON,'') WHEN '' THEN T1.CARTON ELSE T0.CARTON END 箱序,");
            sb.Append("      CASE ISNULL(T0.PIC,'') WHEN '' THEN T1.SHIPPING ELSE T0.PIC END 片序 FROM ap_invoiceIN T0 ");
            sb.Append("     LEFT JOIN ap_invoiceOUT T1 ON (T0.PIC=T1.SHIPPING)");
            sb.Append("     WHERE PIC <> ''");


            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("  and T0.DOCDATE between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (comboBox1.SelectedValue.ToString() != "")
            {
                sb.Append(" and  T0.ITEMCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append(" and  T0.CARD like '%" + textBox3.Text.ToString() + "%'  ");
            }
            if (comboBox2.SelectedValue.ToString() != "")
            {
                sb.Append(" and  Substring(T0.ITEMCODE,12,1)  like '%" + comboBox2.SelectedValue.ToString() + "%'  ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" and  T0.INV like '%" + textBox4.Text.ToString() + "%'  ");
            }

            if (textBox5.Text != "")
            {
                sb.Append(" and   CASE ISNULL(T0.PIC,'') WHEN '' THEN T1.SHIPPING ELSE T0.PIC END = '" + textBox5.Text.ToString() + "'  ");
            }
            if (textBox6.Text != "")
            {
                sb.Append(" and  CASE ISNULL(T0.CARTON,'') WHEN '' THEN T1.CARTON ELSE T0.CARTON END = '" + textBox6.Text.ToString() + "'  ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" select T0.DOCDATE 進貨日期,SAPDOC 交貨單號 ,CARD 客戶,T0.ITEMCODE 產品編號,T0.ITEMNAME 品名規格");
            sb.Append("     ,T0.QTY 數量,t1.cardcode 供應商,INVDATE,T1.INVOICE '原廠inv#',PARTNO 料號,CASE ISNULL(T0.CARTON,'') WHEN '' THEN T1.CARTON ELSE T0.CARTON END 箱序,");
            sb.Append("      CASE ISNULL(T0.PIC,'') WHEN '' THEN T1.SHIPPING ELSE T0.PIC END 片序 FROM ap_invoiceIN T0 ");
            sb.Append(" LEFT JOIN ap_invoiceOUT T1 ON (T0.CARTON=T1.CARTON)");
            sb.Append(" WHERE T0.CARTON <> ''");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("  and T0.DOCDATE between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (comboBox1.SelectedValue.ToString() != "")
            {
                sb.Append(" and  T0.ITEMCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append(" and  T0.CARD like '%" + textBox3.Text.ToString() + "%'  ");
            }
            if (comboBox2.SelectedValue.ToString() != "")
            {
                sb.Append(" and  Substring(T0.ITEMCODE,12,1)  like '%" + comboBox2.SelectedValue.ToString() + "%'  ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" and  T0.INV like '%" + textBox4.Text.ToString() + "%'  ");
            }

            if (textBox5.Text != "")
            {
                sb.Append(" and   CASE ISNULL(T0.PIC,'') WHEN '' THEN T1.SHIPPING ELSE T0.PIC END = '" + textBox5.Text.ToString() + "'  ");
            }
            if (textBox6.Text != "")
            {
                sb.Append(" and  CASE ISNULL(T0.CARTON,'') WHEN '' THEN T1.CARTON ELSE T0.CARTON END = '" + textBox6.Text.ToString() + "'  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMODEL()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append(" Substring (T1.[ItemCode],2,8) END Model");
            sb.Append(" FROM  RDR1 T1 ");
            sb.Append(" left join oitm t2 on (t1.itemcode=t2.itemcode)");
            sb.Append(" WHERE   T1.ITEMCODE <> '' and t2.itmsgrpcod=1032");
            sb.Append(" UNION ALL SELECT '' ");
            sb.Append(" order by CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append(" Substring (T1.[ItemCode],2,8) END");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetITEMNAME(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMNAME FROM OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetVER()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct Substring(T1.[ItemCode],12,1)  VER");
            sb.Append(" FROM  RDR1 T1 ");
            sb.Append(" left join oitm t2 on (t1.itemcode=t2.itemcode)");
            sb.Append(" WHERE   Substring(T1.[ItemCode],12,1) NOT IN ('','.') and t2.itmsgrpcod=1032");
            sb.Append(" UNION ALL SELECT '' ");
            sb.Append(" ORDER BY Substring(T1.[ItemCode],12,1)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
 


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetMODEL(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("                 AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append("          Substring (T1.[ItemCode],2,8) END Model, Substring(T1.[ItemCode],12,1)  VER,T1.[ItemCode] ItemCode");
            sb.Append("       FROM OITM T1 WHERE USERTEXT  like '%" + ITEMCODE + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetINV(string INV)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select Convert(varchar(8),U_ACME_INVOICE,112) INVDATE,CARDCODE 客戶編號 from OPDN T1 ");
            sb.Append("   WHERE U_ACME_INV  like '%" + INV + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void SERIAL_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            textBox7.Text = GetMenu.DFirst();
            textBox8.Text = GetMenu.DLast();
            UtilSimple.SetLookupBinding(comboBox1, GetMODEL(), "MODEL", "MODEL");
            UtilSimple.SetLookupBinding(comboBox2, GetVER(), "VER", "VER");
            textBox9.Text = GetMenu.Day();
            comboBox3.Text = "國內";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\EXCEL2\\採購匯入\\TFT\\";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {

                        GetINVOICE(file);                    
                }

                dataGridView2.DataSource = GetOrderData19AP();
            }
            catch (Exception ex)
            {
              
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                for (int i = 0; i <= dataGridView2.SelectedRows.Count - 1; i++)
                {

                    DataGridViewRow row;
                    row = dataGridView2.SelectedRows[0];
                    string INVDATE = row.Cells["INVDATE"].Value.ToString();
                    string PART = row.Cells["PART"].Value.ToString();
                    string SHIPPING = row.Cells["SHIPPING"].Value.ToString();
                    string INVOICE = row.Cells["INVOICE"].Value.ToString();
                    string CARTON = row.Cells["CARTON"].Value.ToString();
                    string MODEL = row.Cells["MODEL"].Value.ToString();
                    string VER = row.Cells["VER"].Value.ToString();
                    UPINVOUT(SHIPPING, PART, INVOICE, CARTON, MODEL, VER, INVDATE);

                    MessageBox.Show("資料已更新");
                }

            }
            else
            {
                MessageBox.Show("請點選要更新的列");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
                    if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToCSV2(dataGridView1, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToCSV2(dataGridView2, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToCSV2(dataGridView3, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable J1 = GetSERREPORT();
            dataGridView4.DataSource = J1;
            System.Data.DataTable J2 = GetSERREPORT2();
            dataGridView5.DataSource = J2;



                CarlosAg.ExcelXmlWriter.Workbook book = new CarlosAg.ExcelXmlWriter.Workbook();
                WorksheetStyle headerStyle = book.Styles.Add("headerStyleID");
                headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
                headerStyle.Alignment.WrapText = true;
                headerStyle.Interior.Color = "#284775";
                headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;
                headerStyle.Font.Color = "white";
                headerStyle.Font.Bold = true;

                WorksheetStyle defaultStyle = book.Styles.Add("workbookStyleID");
                defaultStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
                defaultStyle.Alignment.WrapText = true;
                defaultStyle.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
                defaultStyle.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1, "#000000");
                defaultStyle.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
                defaultStyle.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");
                WH(book, dataGridView5, "加總");

                WH(book, dataGridView4, "明細");
  
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
        DateTime.Now.ToString("yyyyMMddHHmmss") + "序號管理.xls";
                book.Save(OutPutFile);
                System.Diagnostics.Process.Start(OutPutFile);
            
        }
        private void WH(CarlosAg.ExcelXmlWriter.Workbook book, DataGridView DGV, string DD)
        {



           CarlosAg.ExcelXmlWriter.Worksheet sheet = book.Worksheets.Add(DD);
            WorksheetRow headerRow = sheet.Table.Rows.Add();
            for (int i = 0; i < DGV.Columns.Count ; i++)
            {
                headerRow.Cells.Add(DGV.Columns[i].HeaderText, DataType.String, "headerStyleID");
            }

            for (int i = 0; i < DGV.Rows.Count-1; i++)
            {

                DataGridViewRow row = DGV.Rows[i];
                WorksheetRow rowS = sheet.Table.Rows.Add();

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    //if (j == 0 || j == 1)
                    //{
                        rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                   // }
                    //else
                    //{
                    //    rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                    //}
                    rowS.AutoFitHeight = true;
                    rowS.Table.DefaultColumnWidth = 100;

                }

            }
        }
        public void DD2(string PATH)
        {


            string[] filebBrand = Directory.GetDirectories(PATH);
            foreach (string fileabBrand in filebBrand)
            {
                DirectoryInfo DIRINFO = new DirectoryInfo(fileabBrand);

                string DIRNAME = DIRINFO.Name.ToString();
                string IN = DIRNAME.Substring(0, 1);
                string IN2 = comboBox3.Text.Substring(1, 1);
                if (IN == IN2)
                {
                    string YEAR = DIRNAME.Substring(1, 4);
                    string YEAR2 = textBox9.Text.Substring(0, 4);
                    if (YEAR == YEAR2)
                    {
                        string[] fileccVer = Directory.GetDirectories(fileabBrand);
                        foreach (string filee in fileccVer)
                        {
                            DirectoryInfo DIRINFO2 = new DirectoryInfo(filee);
                            string DIRNAME2 = DIRINFO2.Name.ToString();
                            int G1 = DIRNAME2.IndexOf("月");
                            string MONTH = DIRNAME2.Substring(0, G1);
                            string MONTH2 = Convert.ToInt16(textBox9.Text.Substring(4, 2)).ToString();
                            if (MONTH == MONTH2)
                            {
                                string[] filecSize = Directory.GetFiles(filee);
                                foreach (string fie in filecSize)
                                {
                                    int aa = fie.LastIndexOf(".");
                                    string Type;
                                    Type = fileabBrand.Replace(PATH, "");
                                    DataRow dr;
                                    FileInfo filess = new FileInfo(fie);
                                    string dd = filess.Name.ToString();

                                    int ad = dd.LastIndexOf(".");

                                    string size = filess.Length.ToString();
                                    string FileDate = "";
                                    int P1 = dd.IndexOf("序");
                                    if (ad != -1)
                                    {
                                        if (P1 != -1)
                                        {

                                            if (IN == "內")
                                            {
                                                int F1 = Convert.ToInt16(dd.Substring(P1 + 1, 3)) + 1911;
                                                FileDate = F1.ToString() + dd.Substring(P1 + 4, 4);
                                            }

                                            if (IN == "外")
                                            {
                                                FileDate = dd.Substring(P1 + 1, 8);
                                            }
                                            //if (FileDate == textBox9.Text)
                                            //{
                                                string PanelName = dd.Substring(0, ad).ToString();

                                                System.Data.DataTable GF1 = GetAP_INVOICEIN(PanelName);
                                                if (GF1.Rows.Count == 0)
                                               {
                                                   if (PanelName != "Thumbs")
                                                   {
                                                       if (PanelName != "複本 序1051129鈺緯---74")
                                                       {
                                                           dr = TempDt2.NewRow();
                                                           dr["fie"] = fie;
                                                           dr["PanelName"] = PanelName;
                                                           dr["IN"] = IN;
                                                           dr["DIRNAME"] = DIRNAME;

                                                           TempDt2.Rows.Add(dr);
                                                       }
                                                   }
                                               }
                                          //  }
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
            }


        }
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("fie", typeof(string));
            dt.Columns.Add("PanelName", typeof(string));
            dt.Columns.Add("IN", typeof(string));
            dt.Columns.Add("DIRNAME", typeof(string));


            return dt;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            TempDt2 = MakeTable2();
            string ACME = "//acmesrv01//Public//進出貨序號//出貨序號//";

            DD2(ACME);


            dataGridView6.DataSource = TempDt2;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (dataGridView6.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇要匯入檔案");
                return;
            }


            for (int h = dataGridView6.SelectedRows.Count - 1; h >= 0; h--)
            {
                DataGridViewRow row;
                row = dataGridView6.SelectedRows[h];


                string fie = row.Cells["fie"].Value.ToString();
                string PanelName = row.Cells["PanelName"].Value.ToString();
                string IN = row.Cells["IN"].Value.ToString();
                string DIRNAME = row.Cells["DIRNAME"].Value.ToString();
                System.Data.DataTable G1 = GetAP_INVOICEIN(DIRNAME);
                if (G1.Rows.Count == 0)
                {
                    if (IN == "內")
                    {
                        H2 = 0;
                        GetExcelProduct2(fie, PanelName, IN, DIRNAME);
                        if (H2 > 1)
                        {
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 1);
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 2);
                        }
                        else
                        {
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 0);
                        }
                    }
                    else
                    {
                        GetExcelProduct(fie, PanelName, IN, DIRNAME, 0);
                    }
                }

            }

            MessageBox.Show("匯入成功");
            
        }
        private void GetExcelProduct(string ExcelFile, string FILE, string TYPE, string DIRNAME, int FLAG)
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

            int iRowCntS = 0;
            int iRowCntE = 0;
            if (FLAG == 0)
            {
                iRowCntS = 2;
                iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;
            }
            if (FLAG == 1)
            {
                iRowCntS = 2;
                iRowCntE = H - 1;
            }
            if (FLAG == 2)
            {
                iRowCntS = H;
                iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;
            }
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string WHNO = "";
                string ITEMCODE;
                string DOCDATE = "";
                string ITEMCODE2 = "";
                string ITEMCODE3 = "";
                string PACKNO = "";
                string PACKNO2 = "";
                string SERNO = "";
                string PARTNO = "";
                string QTY = "";
                string QTYOUT = "";
                string INV = "";
                string YEAR = "";
                string MON = "";
                int CHECK1 = 0;
                DataRow dr;
                int CHECK = 0;
                int O1 = 0;
                int O2 = 0;
                string PACKNO3 = "";
                int n;
                string id1x = "";
                string id2x = "";
                string id3x = "";

                if (TYPE == "內")
                {
                    DOCDATE = FILE.Substring(1, 7);
                    YEAR = FILE.Substring(1, 3);
                    MON = FILE.Substring(4, 4);
                    DOCDATE = (Convert.ToInt16(YEAR) + 1911).ToString() + MON;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 11]);
                    if (FLAG == 2)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[H, 11]);
                    }
                    // range.Select();
                    WHNO = range.Text.ToString().Trim();
                }
                if (TYPE == "外")
                {
                    DOCDATE = FILE.Substring(1, 8);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 7]);
                    // range.Select();
                    WHNO = range.Text.ToString().Trim();
                }
                if (!String.IsNullOrEmpty(WHNO))
                {
                    //第一行要
                    for (int iRecord = iRowCntS; iRecord <= iRowCntE; iRecord++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        // range.Select();
                        ITEMCODE = range.Text.ToString().Trim();




                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            PARTNO = range.Text.ToString().Trim();


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                            QTY = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            INV = range.Text.ToString().Trim();

                        }

                        if (TYPE == "外")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            PARTNO = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            INV = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            QTY = range.Text.ToString().Trim();

                        }

                        if (ITEMCODE != "")
                        {
                            id1x = ITEMCODE;
                        }
                        if (ITEMCODE == "")
                        {
                            ITEMCODE = id1x;
                        }

                        if (PARTNO != "")
                        {
                            id2x = PARTNO;
                        }
                        if (PARTNO == "")
                        {
                            PARTNO = id2x;
                        }


                        if (INV != "")
                        {
                            id3x = INV;
                        }
                        if (INV == "")
                        {
                            INV = id3x;
                        }


                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                            PACKNO = range.Text.ToString().Trim().Replace("*", "").ToUpper();
                        }
                        if (TYPE == "外")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                            PACKNO = range.Text.ToString().Trim().Replace("*", "").ToUpper();
                        }
                        if (!String.IsNullOrEmpty(PACKNO))
                        {
                            if (PACKNO.Length > 6)
                            {
                                PACKNO2 = PACKNO.Substring(4, 2);

                            }

                        }


                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                            // range.Select();
                            SERNO = range.Text.ToString().Trim();
                        }
                        string CARDNAME = "";
                        string WAREHOUSE = "";
                        System.Data.DataTable t1 = Getdata2(WHNO);
                        if (t1.Rows.Count > 0)
                        {

                            CARDNAME = t1.Rows[0]["CARDNAME"].ToString();
                            WAREHOUSE = t1.Rows[0]["WAREHOUSE"].ToString();
                        }

                        if (!String.IsNullOrEmpty(QTY))
                        {
                            if (PACKNO.Trim() != "片數總計" && SERNO.Trim() != "片數總計" && PACKNO.Trim() != "片數總計:")
                            {
                                string DDATE = textBox1.Text.Substring(0, 6);
                          
                                    AddINVIN(ITEMCODE, FILE, PARTNO, INV, PACKNO, SERNO, QTY, CARDNAME, "", DOCDATE, DateTime.Now.ToString("yyyyMMdd"), WAREHOUSE, WHNO, DDATE, TYPE, fmLogin.LoginID.ToString());
                                
                            }
                        }


                    }

                }
                else
                {
                    CHECK = 2;
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

        private void GetExcelProduct2(string ExcelFile, string FILE, string TYPE, string DIRNAME)
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


            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string WHNO = "";



                if (TYPE == "內")
                {

                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                        WHNO = range.Text.ToString().Trim();


                        int F1 = WHNO.ToUpper().IndexOf("WH");
                        if (F1 != -1)
                        {

                            WHNO = range.Text.ToString().Trim();
                            if (WHNO.Length == 14)
                            {
                                H2++;
                                if (H2 == 2)
                                {
                                    H = iRecord;
                                }


                            }
                        }
                    }
                    // range.Select();

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
        private System.Data.DataTable Getdata2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME,SHIPPING_OBU WAREHOUSE FROM wh_main WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
    }
}
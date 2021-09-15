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
using SAPbobsCOM;
namespace ACME
{
    public partial class APOPOR2 : Form
    {
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string FA = "acmesql98";
        public APOPOR2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
                       //try
                       //{
                           OpenFileDialog opdf = new OpenFileDialog();
                           DialogResult result = opdf.ShowDialog();
                           if (opdf.FileName.ToString() == "")
                           {
                               MessageBox.Show("請選擇檔案");
                           }
                           else
                           {
                               DELOPOR();
                               GD5(opdf.FileName);
                             

                                   System.Data.DataTable G1 = GetOPOR();
                                   dataGridView1.DataSource = G1;
                                                     System.Data.DataTable GG2 = GetOPOR6();
                                                     label1.Text = "數量 : "+GG2.Rows[0]["QTY"].ToString();
                                                     label2.Text = "金額 : "+GG2.Rows[0]["AMT"].ToString();
                               
                               
                           }
                       //}
                       //catch (Exception ex)
                       //{
                       //    MessageBox.Show(ex.Message);
                       //}
        }
        private void GD5(string ExcelFile)
        {

            int h = 0;
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

            string DUP = "";
            int D = 0;
            int L = 0;

            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }


                decimal PRICE = 0;

                string REMARK2;
                string REMARK3;
                string P1;

                string SHIPNO;


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    P1 = range.Text.ToString().Trim();

                    decimal n;
                    if (decimal.TryParse(P1, out n))
                    {
                        PRICE = Convert.ToDecimal(P1);
                    }
    



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    REMARK2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    REMARK3 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    SHIPNO = range.Text.ToString().Trim();

                    try
                    {
                        if (!String.IsNullOrEmpty(SHIPNO))
                        {
                            //if (decimal.TryParse(P1, out n))
                            //{

                            //    ADDOPOR("ZA0SB0005", 1, PRICE, PRICE, REMARK2 + "-" + REMARK3, SHIPNO, 1, h, "U0470", "Z0001", "卡車費");

                            //    h++;
                            //}
                            if (!String.IsNullOrEmpty(SHIPNO))
                            {
                                if (decimal.TryParse(P1, out n))
                                {
                                    if (DUP != REMARK2)
                                    {
                                        D++;
                                        L = 0;
                                    }

                                    int S1 = SHIPNO.IndexOf("+");
                                    if (S1 != -1)
                                    {
                                        string[] arrurl = SHIPNO.Split(new Char[] { '+' });

                                        int L1 = 0;
                                        foreach (string ESi in arrurl)
                                        {
                                            L1++;
                                        }
                                        foreach (string ESi in arrurl)
                                        {
                                            ADDOPOR("ZA0SB0005", 1, PRICE / L1, PRICE / L1, REMARK2 + "-" + REMARK3, ESi, 1, h, "U0470", "Z0001", "卡車費");
                                            //       ADDOPOR(ITEMCODE, 1, Convert.ToDecimal(PRICE) / L1, Convert.ToDecimal(PRICE) / L1, REMARK, ESi, D, L, "U0019", WHSCODE, ITEMNAME);
                                            h++;
                                            L++;
                                        }
                                    }
                                    else
                                    {
                                        ADDOPOR("ZA0SB0005", 1, PRICE, PRICE, REMARK2 + "-" + REMARK3, SHIPNO, 1, h, "U0470", "Z0001", "卡車費");
                                        // ADDOPOR(ITEMCODE, 1, Convert.ToDecimal(PRICE), Convert.ToDecimal(PRICE), REMARK, SHIPNO, D, L, "U0019", WHSCODE, ITEMNAME);
                                        h++;
                                        L++;
                                    }

                                }

                            }

                            DUP = REMARK2;
                        }


                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
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
        private void GD6(string ExcelFile)
        {

            int h = 0;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
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


            string DUP = "";
            int D = 0;
            int L = 0;
            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }


                decimal PRICE = 0;

                string REMARK = "";
                string REMARK1;
                string REMARK2;
                string REMARK3;
                string P1;
                string JITEMNAME;
                string SHIPNO;

                string JIN;
                string ITEMCODE = "";
                string WHSCODE = "";
                string ITEMNAME = "";
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                range.Columns.AutoFit();
                REMARK2 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                REMARK3 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                JITEMNAME = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                SHIPNO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                range.Select();
                P1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                range.Select();
                REMARK1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 16]);
                range.Select();
                JIN = range.Text.ToString().Trim();



                if (JIN == "進口")
                {
                    System.Data.DataTable G1 = GETOPORJ(JITEMNAME);
                    if (G1.Rows.Count > 0)
                    {
                        ITEMCODE = G1.Rows[0]["ITEMCODE"].ToString();
                        WHSCODE = G1.Rows[0]["WHSCODE"].ToString();
                        ITEMNAME = G1.Rows[0]["ITEMNAME"].ToString();
                    }

                    decimal n;
                    if (decimal.TryParse(P1, out n))
                    {
                        PRICE = Convert.ToDecimal(P1);
                    }


                    try
                    {
                        if (!String.IsNullOrEmpty(REMARK1))
                        {
                            //發票#RR71388959, 建新#11190010429, 報單#AT  08G40H1149
                            REMARK = "發票#" + REMARK1 + ", 建新#" + REMARK2 + ", 報單#" + REMARK3;
                     
                        }
                        if (!String.IsNullOrEmpty(SHIPNO))
                        {
                            if (DUP != REMARK2)
                            {
                                D++;
                                L = 0;
                            }
                      
                            int S1 = SHIPNO.IndexOf("+");
                            if (S1 != -1)
                            {
                                string[] arrurl = SHIPNO.Split(new Char[] { '+' });

                                int L1 = 0;
                                foreach (string ESi in arrurl)
                                {
                                    L1++;
                                }
                                foreach (string ESi in arrurl)
                                {
                                    ADDOPOR(ITEMCODE, 1, Convert.ToDecimal(PRICE) / L1, Convert.ToDecimal(PRICE) / L1, REMARK, ESi, D, L, "U0019", WHSCODE,ITEMNAME);
                                    h++;
                                    L++;
                                }
                            }
                            else
                            {
                                ADDOPOR(ITEMCODE, 1, Convert.ToDecimal(PRICE), Convert.ToDecimal(PRICE), REMARK, SHIPNO, D, L, "U0019", WHSCODE, ITEMNAME);
                                h++;
                                L++;
                            }


                           
                        }

                        DUP = REMARK2;
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

        public static System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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

        public static System.Data.DataTable GetOPOR()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR2 WHERE USERS=@USERS ORDER BY ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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

        public static System.Data.DataTable GetOPOR5(string DOC)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR2 WHERE USERS=@USERS AND DOC=@DOC ORDER BY ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
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

        public static System.Data.DataTable GetOPOR55(string DOC)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  REMARK FROM AP_OPOR2 WHERE  USERS=@USERS AND DOC=@DOC AND  ISNULL(REMARK,'') <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
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
        public static System.Data.DataTable GetOPOR6()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR2 WHERE USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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



        public void DELOPOR()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_OPOR2 WHERE USERS=@USERS  ", connection);
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

        public void ADDOPOR(string ITEMCODE, decimal QTY, decimal PRICE, decimal AMT, string REMARK, string SHIPNO, int DOC, int LINENUM, string CARDCODE, string WHSCODE, string ITEMNAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPOR2(ITEMCODE,QTY,PRICE,AMT,REMARK,SHIPNO,DOC,LINENUM,CARDCODE,WHSCODE,ITEMNAME,USERS) values(@ITEMCODE,@QTY,@PRICE,@AMT,@REMARK,@SHIPNO,@DOC,@LINENUM,@CARDCODE,@WHSCODE,@ITEMNAME,@USERS)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@SHIPNO", SHIPNO));
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            //
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
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

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

                
                    System.Data.DataTable G2 = GetOPOR4();

                    if (G2.Rows.Count > 0)
                    {

                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {

                    
                                SAPbobsCOM.Documents oPURCH = null;
                                oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                string T1 = G2.Rows[n][0].ToString();
                                System.Data.DataTable G3 = GetOPOR5(T1);

                                    if (G3.Rows.Count > 0)
                                    {
                                        string CARDCODE = G3.Rows[0]["CARDCODE"].ToString();
                                        oPURCH.CardCode = CARDCODE;
                                        oPURCH.VatPercent = 5;
                                        if (CARDCODE == "U0019")
                                        {
                                            System.Data.DataTable GG3 = GetOPOR55(T1);
                                            if (GG3.Rows.Count > 0)
                                            {

                                                oPURCH.Comments = GG3.Rows[0][0].ToString();
                                            }
                                        }
                                        System.Data.DataTable G7 = GetOPOR7();
                                        if (G7.Rows.Count > 0)
                                        {
                                            oPURCH.DocumentsOwner = Convert.ToInt32(G7.Rows[0][0]);
                                        }
                                        oPURCH.SalesPersonCode = 65;
                                        for (int s = 0; s <= G3.Rows.Count - 1; s++)
                                        {


                                            string ITEMCODE = G3.Rows[s]["ITEMCODE"].ToString();
                                            string WHSCODE = G3.Rows[s]["WHSCODE"].ToString();
                                            string SHIPNO = G3.Rows[s]["SHIPNO"].ToString();
                                            string REMARK = G3.Rows[s]["REMARK"].ToString();
                                            double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                                            double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                                            oPURCH.Lines.WarehouseCode = WHSCODE;
                                            oPURCH.Lines.ItemCode = ITEMCODE;
                                            oPURCH.Lines.Quantity = QTY;
                                            oPURCH.Lines.Price = PRICE;
                                            oPURCH.Lines.VatGroup = "AP5%";
                                            oPURCH.Lines.Currency = "NTD";
                                            oPURCH.Lines.CostingCode = "11111";
                                            oPURCH.Lines.UserFields.Fields.Item("U_Shipping_no").Value = SHIPNO;
                                            if (CARDCODE != "U0019")
                                            {
                                                oPURCH.Lines.UserFields.Fields.Item("U_MEMO").Value = REMARK;
                                            }
                                            oPURCH.Lines.Add();
                          
                                        }
                                    }

                                int res = oPURCH.Add();
                                if (res != 0)
                                {
                                    MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                                }
                                else
                                {
                                    System.Data.DataTable G4 = GetDI4();
                                    string OWTR = G4.Rows[0][0].ToString();
                                    MessageBox.Show("上傳成功 採購單號 : " + OWTR);

                           

                                    //if (GG3.Rows.Count > 0)
                                    //{
                                    //    string CARDCODE = GG3.Rows[0]["CARDCODE"].ToString();

                                    //    if (CARDCODE == "U0019")
                                    //    {
                                    //        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                    //        {
                                    //            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();

                                    //            string SHIPNO = GG3.Rows[j]["SHIPNO"].ToString();
                                    //            UPDATEOPOR2(OWTR, LINENUM, SHIPNO);
                                    //        }

                                    //        System.Data.DataTable GG4 = GetOPOR55();
                                    //        if (GG4.Rows.Count > 0)
                                    //        {
                                    //            UPDATEOPOR3(OWTR, GG4.Rows[0]["REMARK"].ToString());
                                    //        }
                                    //    }
                                    //    else
                                    //    {
                                    //        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                    //        {

                                    //            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                    //            string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                    //            string SHIPNO = GG3.Rows[j]["SHIPNO"].ToString();
                                    //            UPDATEOPOR(OWTR, LINENUM, REMARK, SHIPNO);
                                    //        }
                                    //    }

                                 //  }
                                }
                            
                        }
                    }
                        

                    
                }
           


        

            
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }
        }

        public static System.Data.DataTable GetOPOR4()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR2 WHERE USERS=@USERS ORDER BY DOC ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        public static System.Data.DataTable GetOPOR7()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT EMPID FROM OHEM WHERE homeTel =@homeTel ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@homeTel", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GETOPORJ(string JITEMNAME)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMCODE,WHSCODE,JITEMNAME ITEMNAME FROM AP_OPOR2J WHERE JITEMNAME =@JITEMNAME ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JITEMNAME", JITEMNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["wh_main"];
        }


        public System.Data.DataTable GetDI4()
        {
            SqlConnection connection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OPOR");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void UPDATEOPOR(string DOCENTRY, string LINENUM, string U_MEMO, string U_Shipping_no)
        {

            SqlConnection connection = new SqlConnection(strCn98);

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET U_MEMO=@U_MEMO,U_Shipping_no=@U_Shipping_no WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@U_Shipping_no", U_Shipping_no));
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

        private void UPDATEOPOR2(string DOCENTRY, string LINENUM, string U_Shipping_no)
        {

            SqlConnection connection = new SqlConnection(strCn98);
            //SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET U_Shipping_no=@U_Shipping_no WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@U_Shipping_no", U_Shipping_no));
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

        private void UPDATEOPOR3(string DOCENTRY, string Comments)
        {

            SqlConnection connection = new SqlConnection(strCn98);
            //SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OPOR SET Comments=@Comments WHERE DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@Comments", Comments));
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
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }


        private void button3_Click(object sender, EventArgs e)
        {
             OpenFileDialog opdf = new OpenFileDialog();
                           DialogResult result = opdf.ShowDialog();
                           if (opdf.FileName.ToString() == "")
                           {
                               MessageBox.Show("請選擇檔案");
                           }
                           else
                           {
                               //opdf.FileName
                               DELOPOR();
                               DataRow dr;

                               //string FileName = GetExePath() + "\\Rtf\\" + "rtf01.txt";


                               //string FileText = System.IO.File.ReadAllText(FileName);

                               richTextBox1.LoadFile(opdf.FileName);
                               //textBox6.Text = richTextBox1.Text;

                               string FileText = richTextBox1.Text;
                               //string[] lines = System.IO.File.ReadAllLines(@"C:\Users\Public\TestFolder\WriteLines2.txt");


                               txtRtf.Text = FileText;

                               //string[] stringSeparators = new string[] { "單號：", "備註：", "合計：" };
                               string[] stringSeparators = new string[] { "單號：", "備註：" };

                               string[] stringSeparators2 = new string[] { "合計：", "車型：" };
                               string[] Lines = FileText.Split(stringSeparators, StringSplitOptions.None);

                               //MessageBox.Show(Lines.Length.ToString());

                               string REMARK = "";
                               string 單號 = "";
                               string 合計 = "";
                               string 備註 = "";
                               string 工單號碼 = "";
                               int h = 0;
                               for (int i = 1; i <= Lines.Length - 1; i++)
                               {
                                   // textBox6.Text += i.ToString() + "-->" + Lines[i] + "\r\n";

                                   if (i % 2 == 1)
                                   {
                                       try
                                       {
                                           單號 = Lines[i].Substring(0, 11);

                                           string[] Text2 = Lines[i].Split(stringSeparators2, StringSplitOptions.None);

                                           string[] Text3 = Text2[1].Split('\n');

                                           合計 = Text3[1];

                                       }
                                       catch
                                       { }
                                   }


                                   if (i % 2 == 0)
                                   {

                                       try
                                       {
                                           string Text = Lines[i].Split('\n')[1];
                                           string[] Text5 = Text.Split('-');

                                           備註 = Text5[0];
                                           工單號碼 = Text5[1];
                                           //WH20190528022X, 報單#AW  08G40G9221
                                           if (工單號碼.Substring(0, 2) != "SH")
                                           {
                                               int x = 工單號碼.IndexOf("SH");
                                               工單號碼 = 工單號碼.Substring(x, 14);
                                           }

  
                                       }
                                       catch
                                       {
                                       }
                                       int S1 = 工單號碼.IndexOf("+");
                                       if (S1 != -1)
                                       {
                                           string[] arrurl = 工單號碼.Split(new Char[] { '+' });


                                           foreach (string ESi in arrurl)
                                           {
                                               REMARK = ESi + ", 報單#" + 備註;
                                               ADDOPOR("ZA0SB0005", 1, Convert.ToDecimal(合計) / 2, Convert.ToDecimal(合計) / 2, REMARK, ESi, 1, h, "U0017", "Z0001","卡車費");
                                               h++;
                                           }
                                       }
                                       else
                                       {
                                           REMARK = 工單號碼 + ", 報單#" + 備註;
                                           ADDOPOR("ZA0SB0005", 1, Convert.ToDecimal(合計), Convert.ToDecimal(合計), REMARK, 工單號碼, 1, h, "U0017", "Z0001", "卡車費");
                                           h++;
                                       }

                                   }
                               }


                               System.Data.DataTable G1 = GetOPOR();
                               dataGridView1.DataSource = G1;
                               System.Data.DataTable GG2 = GetOPOR6();
                               label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                               label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

                           }

          

        }

        private void button7_Click(object sender, EventArgs e)
        {
            //try
            //{
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR();
                GD6(opdf.FileName);


                System.Data.DataTable G1 = GetOPOR();
                if (G1.Rows.Count > 0)
                {
                    dataGridView1.DataSource = G1;
                    System.Data.DataTable GG2 = GetOPOR6();
                    label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                    label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();
                }
                else
                {
                    MessageBox.Show("EXCEL 格式不符");
                }


            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CalcTotals1();
        }

        private void CalcTotals1()
        {

            Int32 iTotal = 0;


            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["AMT"].Value);

            }
            textBox11.Text = iTotal.ToString("#,##0");

        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                //opdf.FileName
                DELOPOR();
                DataRow dr;

                //string FileName = GetExePath() + "\\Rtf\\" + "rtf01.txt";


                //string FileText = System.IO.File.ReadAllText(FileName);

                richTextBox1.LoadFile(opdf.FileName);
                //textBox6.Text = richTextBox1.Text;

                string FileText = richTextBox1.Text;
                //string[] lines = System.IO.File.ReadAllLines(@"C:\Users\Public\TestFolder\WriteLines2.txt");


                txtRtf.Text = FileText;

                //string[] stringSeparators = new string[] { "單號：", "備註：", "合計：" };
                string[] stringSeparators = new string[] { "單號：", "備註：" };

                string[] stringSeparators2 = new string[] { "合計：", "車型：" };
                string[] Lines = FileText.Split(stringSeparators, StringSplitOptions.None);

                //MessageBox.Show(Lines.Length.ToString());

                string REMARK = "";
                string 單號 = "";
                string 合計 = "";
                string 備註 = "";
                string 工單號碼 = "";
                int h = 0;
                for (int i = 1; i <= Lines.Length - 1; i++)
                {
                    // textBox6.Text += i.ToString() + "-->" + Lines[i] + "\r\n";

                    if (i % 2 == 1)
                    {
                        try
                        {
                            單號 = Lines[i].Substring(0, 11);

                            string[] Text2 = Lines[i].Split(stringSeparators2, StringSplitOptions.None);

                            string[] Text3 = Text2[1].Split('\n');

                            合計 = Text3[1];

                        }
                        catch
                        { }
                    }

                                               StringBuilder sb = new StringBuilder();
                    if (i % 2 == 0)
                    {

                        try
                        {
                            string Text = Lines[i].Split('\n')[1];
                            string[] Text5 = Text.Split('-');

                            備註 = Text5[0];
                            //工單號碼 = Text5[1];
                            ////WH20190528022X, 報單#AW  08G40G9221
                            //if (工單號碼.Substring(0, 2) != "WH")
                            //{
                            //    int x = 工單號碼.IndexOf("WH");
                            //    工單號碼 = 工單號碼.Substring(x, 14);
                            //}
 
                            int GG1 = Text.IndexOf("SH20");

                            if (GG1 != -1)
                            {
                                string H1 = Text.Substring(GG1, Text.Length - GG1);
                                string[] arrurl = H1.Split(new Char[] { ',' });
                                          int KK2 = H1.IndexOf(",");
                                          int K1 = 0;
                                          if (KK2 != -1)
                                          {

                                              foreach (string i2 in arrurl)
                                              {
                                                  K1++;
                                                  string MM = i2.Substring(0, 14);

                                                  int T1 = MM.IndexOf("SH");

                                                  if (T1 != -1)
                                                  {
                                                      sb.Append(MM.Trim() + "/");
                                                  }
                                              }

                                              if (K1 != 0)
                                              {
                                                  sb.Remove(sb.Length - 1, 1);
                                              }
                                          }

                                if (K1 > 0)
                                {
                              
                                }
                                else
                                {
                                    int K2 = 0;
                                    string[] arrurl2 = H1.Split(new Char[] { '/' });


                                    foreach (string i2 in arrurl2)
                                    {

                                        if (i2.Length > 13)
                                        {
                                            string MM = i2.Trim().Substring(0, 14);

                                            System.Data.DataTable GSHIP = GetSHIP(MM);

                                            if (GSHIP.Rows.Count > 0)
                                            {
                                                K2++;
                                                sb.Append(MM.Trim() + "/");
                                            }
                                            else
                                            {
                                                int T1 = MM.IndexOf("WH");

                                                if (T1 != -1)
                                                {
                                                    K2++;
                                                    sb.Append(MM.Trim() + "/");
                                                }

                                            }

                                        }

           
                                    }
                                    if (K2 != 0)
                                    {
                                        sb.Remove(sb.Length - 1, 1);
                                    }

                                    
                                }

                            }
                        }
                        catch
                        {
                        }
                       // REMARK = sb.ToString() + ", 報單#" + 備註;
                        ADDOPOR("ZA0SB0005", 1, Convert.ToDecimal(合計), Convert.ToDecimal(合計), "", sb.ToString(), 1, h, "U0017", "Z0002", "卡車費");
                        h++;

                    }
                }


                System.Data.DataTable G1 = GetOPOR();
                dataGridView1.DataSource = G1;
                System.Data.DataTable GG2 = GetOPOR6();
                label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

            }
        }

        private void APOPOR2_Load(object sender, EventArgs e)
        {
            if (globals.DBNAME == "進金生")
            {
                strCn98 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                FA = "acmesql02";
            }
        }

    }
}

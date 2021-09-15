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
    public partial class APOPORM : Form
    {
       string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
       string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
       string FA = "acmesql98";
        public APOPORM()
        {
            InitializeComponent();
        }




        private void GD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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



            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string ITEMCODE;
                string LINE;
                string AU97;
                string GRADE;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK;
                string P1;
                string SITE;
                string ARRIVE;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                ITEMCODE = range.Text.ToString().Trim();

                if (ITEMCODE != "")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    LINE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    SITE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    ARRIVE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    AU97 = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                    range.Select();
                    P1 = range.Text.ToString().Trim();

                    decimal n;
                    if (decimal.TryParse(P1, out n))
                    {
                        PRICE = Convert.ToDecimal(P1);
                    }
                    else
                    {
                        PRICE = 0;
                    }

                    AMT = PRICE * QTY;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    REMARK = range.Text.ToString().Trim();

                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {

                            string ITEM = "";
                            System.Data.DataTable F1 = GetITEMCODE(ITEMCODE, GRADE, AU97);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();
                            }
                            else
                            {
                                System.Data.DataTable F2 = GetITEMCODE2(ITEMCODE, GRADE, AU97);
                                if (F2.Rows.Count > 0)
                                {
                                    ITEM = F2.Rows[0][0].ToString();
                                }
                                else
                                {

                                    System.Data.DataTable F3 = GetITEMCODE3(ITEMCODE, GRADE, AU97);
                                    if (F3.Rows.Count > 0)
                                    {
                                        ITEM = F3.Rows[0][0].ToString();
                                    }
                                }
                            }

                            int GA = ARRIVE.IndexOf("回台");
                            if (GA != -1)
                            {
                                ARRIVE = "回台";
                            }

                            if (String.IsNullOrEmpty(ITEM))
                            {
                                ITEM = "ACME00001.00001";
                            }

                          //  ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK, "S0001-DD", fmLogin.LoginID.ToString(),"",SITE,ARRIVE);
                        }


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

        public  System.Data.DataTable GetITEMCODE(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 T0.ITEMCODE  FROM PDN1 T0 LEFT JOIN  OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;

               
                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                         s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (ITEMNAME LIKE '%" + ITEM1 + "%' OR ITEMNAME LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode + "%'   ");
            }


            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }


            sb.Append("  ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
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
        public  System.Data.DataTable GetITEMCODEGD(string MODEL, string U_GRADE, string VERSION)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 T0.ITEMCODE  FROM PDN1 T0 LEFT JOIN  OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append("  WHERE T1.U_TMODEL=@MODEL");

            sb.Append(" AND T1.ITEMNAME  LIKE '%" + VERSION + "%' ");

            if (U_GRADE == "")
            {
                sb.Append(" AND U_GRADE='Z' ");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }


            sb.Append("  ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@VERSION", VERSION));
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
        public System.Data.DataTable GetPP(string DOCENTRY,string ITEMCODE,int QTY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LINENUM FROM PQT1 WHERE DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE AND QUANTITY=@QTY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public System.Data.DataTable GetM1(string M1, string M2, string M3)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPORM WHERE M1  like '%" + M1 + "%'  AND M2 like '%" + M2 + "%' AND M3  like '%" + M3 + "%' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public static System.Data.DataTable GetOPOR(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR3 WHERE USERS=@USERS ORDER BY ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR6(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR WHERE USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR6IN(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
            SqlConnection MyConnection = globals.shipConnection ;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 　TOP 1 RATE　FROM ORTT　WHERE Currency ='USD'　ORDER BY RATEDATE DESC ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public static System.Data.DataTable GetOPOR2(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPOR WHERE USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR2D(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            // sb.Append(" SELECT * FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS ");
            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPOR WHERE  ID IN ( " + cs + ") AND USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR WHERE USERS=@USERS AND DOC >0 ORDER BY DOC ");
            //            sb.Append(" SELECT * FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4D(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS AND DOC >0 ORDER BY DOC ");
            //            sb.Append(" SELECT * FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR3(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT SI FROM [AP_OPOR3] WHERE USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR3D(string cs, string DOCENTRY, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPOR WHERE DOCENTRY=@DOCENTRY AND ID IN ( " + cs + ")  AND DOCENTRY >0 AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5(string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE DOC=@DOC AND DOC >0  AND USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5ID(string cs,string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE  DOC=@DOC AND ID IN ( " + cs + ")  AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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

        public static System.Data.DataTable GetSI1()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT SI FROM [AP_OPOR3] WHERE USERS=@USERS");
   
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
        public static System.Data.DataTable GetSI2(string SI)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT *  FROM [AP_OPOR3] WHERE USERS=@USERS AND SI=@SI");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@SI", SI));
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
        public  System.Data.DataTable GetITEMCODE2(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 T0.ITEMCODE  FROM PDN1 T0 LEFT JOIN  OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            sb.Append(" AND  U_TMODEL +'.'+U_VERSION  like '%" + ItemCode + "%'  ");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;


                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                        s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (   U_TMODEL +'.'+U_VERSION LIKE '%" + ITEM1 + "%' OR U_TMODEL +'.'+U_VERSION LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION  LIKE '%" + ItemCode + "%'   ");
            }
            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }

            sb.Append("  ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));

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

        public  System.Data.DataTable GetITEMCODE3(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 ITEMCODE  FROM   OITM ");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;


                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                        s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (ITEMNAME LIKE '%" + ITEM1 + "%' OR ITEMNAME LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode + "%'   ");
            }


            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
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
        public void DELOPOR(string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_OPOR3 WHERE USERS=@USERS ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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

        public void ADDOPOR(string SI, int DOCENTRY, int LINENUM, string ITEMCODE, decimal QTY, decimal PRICE, decimal AMT, string CUSTOMER, string CLOSEDAY, string TRADE, string P1, string P2, string STATUS,string receivePlace, string goalPlace, string shipment, string unloadCargo)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPOR3(SI,DOCENTRY,LINENUM,ITEMCODE,QTY,PRICE,AMT,CUSTOMER,CLOSEDAY,TRADE,P1,P2,USERS,STATUS,receivePlace,goalPlace,shipment,unloadCargo) values(@SI,@DOCENTRY,@LINENUM,@ITEMCODE,@QTY,@PRICE,@AMT,@CUSTOMER,@CLOSEDAY,@TRADE,@P1,@P2,@USERS,@STATUS,@receivePlace,@goalPlace,@shipment,@unloadCargo)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SI", SI));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@CUSTOMER", CUSTOMER));
            command.Parameters.Add(new SqlParameter("@CLOSEDAY", CLOSEDAY));
            command.Parameters.Add(new SqlParameter("@TRADE", TRADE));
            command.Parameters.Add(new SqlParameter("@P1", P1));
            command.Parameters.Add(new SqlParameter("@P2", P2));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@STATUS", STATUS));
                command.Parameters.Add(new SqlParameter("@receivePlace", receivePlace));
            command.Parameters.Add(new SqlParameter("@goalPlace", goalPlace));
            command.Parameters.Add(new SqlParameter("@shipment", shipment));
            command.Parameters.Add(new SqlParameter("@unloadCargo", unloadCargo));

            //STATUS
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


        public void UPOPOR(int DOC,int LINENUM, int ID,string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPOR SET DOC=@DOC,LINENUM=@LINENUM WHERE ID=@ID AND USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public void UPOPORD(int DOC, int LINENUM, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPOR SET DOC=@DOC,LINENUM=@LINENUM WHERE USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
  
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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


        public System.Data.DataTable GetDI4()
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OPOR");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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

        public System.Data.DataTable GetDESC(string ITEMCODE)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMNAME FROM OITM WHERE ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        private void UPDATEOPOR(string DOCENTRY, string LINENUM, string U_MEMO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET U_MEMO=@U_MEMO,U_ACME_Dscription='OA 45 days'  WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            try
            {

                try
                {
                    MyConnection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                MyConnection.Close();
            }


        }

        private void UPDATEAPOPOR(string DOC, string LINENUM, string SAPDOC)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE AP_OPOR SET SAPDOC=@SAPDOC WHERE DOC=@DOC AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
            try
            {

                try
                {
                    MyConnection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                MyConnection.Close();
            }


        }

        private void UPDATEAPOPORD(string cs,string DOC, string LINENUM, string SAPDOC)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE AP_OPOR SET SAPDOC=@SAPDOC WHERE DOC=@DOC AND ID IN ( " + cs + ")  AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

               command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
            try
            {

                try
                {
                    MyConnection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                MyConnection.Close();
            }


        }
        //ID IN ( " + cs + ") 
   
        private void GD6(string ExcelFile)
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
            excelSheet.Activate();
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
                string ITEMCODE;
                string LINE;
                string CLOSEDAY;
                string CLOSEDAY2="";
                string DOCENTRY;
                string LINENUM = "0";
                string MODEL;
                string GRADE;
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string CUSTOMER;
                string PAYMENT;
                string PAYMENT2 = "";
                string SI;
                string P1;
                string PP1 = "";
                string PP2 = "";
                string STATUS;
                string SITE;
                                string BAN;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                DOCENTRY = range.Text.ToString().Trim();

                if (DOCENTRY != "")
                {
            
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    VER = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    P1 = range.Text.ToString().Trim();

                    decimal n;
                    if (decimal.TryParse(P1, out n))
                    {
                        PRICE = Convert.ToDecimal(P1);
                    }
                    else
                    {
                        PRICE = 0;
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    AMT = Convert.ToDecimal(range.Text.ToString().Trim());


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    CUSTOMER = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                    range.Select();
                    STATUS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    PAYMENT = range.Text.ToString().Trim();
                    int  ff = PAYMENT.IndexOf(",");
                    if (ff != -1)
                    {
                        PAYMENT2 = PAYMENT.Substring(ff + 1, PAYMENT.Length - ff - 1);
                        CLOSEDAY = DateTime.Now.ToString("yyyy") + "/" + PAYMENT.Substring(0, ff);
                        DateTime g1 = Convert.ToDateTime(CLOSEDAY);
                         CLOSEDAY2 = g1.ToString("yyyyMMdd");
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 13]);
                    range.Select();
                    SI = range.Text.ToString().Trim();

                    
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 14]);
                    range.Select();
                    BAN = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 15]);
                    range.Select();
                    SITE = range.Text.ToString().Trim();

                    try
                    {
                        if (!String.IsNullOrEmpty(MODEL))
                        {

                            string ITEM = "";
                            System.Data.DataTable F1 = GetITEMCODEGD(MODEL, GRADE, VER);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();

                                System.Data.DataTable F2 = GetPP(DOCENTRY, ITEM, Convert.ToInt16(QTY));

                                if (F2.Rows.Count > 0)
                                {
                                    LINENUM = F2.Rows[0][0].ToString();
                                }
                            }
                            else
                            {
                                ITEM = "ACME00004.00004";
                            }

                            if(BAN=="CY"||BAN=="CFS")
                            {
                            PP1="SEA";
                            PP2 = "進口";
                            }
                            if (BAN == "TRUCK")
                            {
                                PP1 = "TRUCK";
                                PP2 = "三角";
                            }
                            if (BAN == "DHL")
                            {
                                PP1 = "AIR";
                                PP2 = "進口";
                            }
                            string receivePlace = "";
                            string goalPlace = "";
                            string shipment = "";
                            string unloadCargo = "";
                            System.Data.DataTable T1 = GetM1(PAYMENT2, BAN, SITE);
                            if (T1.Rows.Count > 0)
                            {
                                receivePlace = T1.Rows[0]["M4"].ToString();
                                goalPlace = T1.Rows[0]["M5"].ToString();
                                shipment = T1.Rows[0]["M6"].ToString();
                                unloadCargo = T1.Rows[0]["M7"].ToString(); 
                            
                            }
                            ADDOPOR(SI, Convert.ToInt32(DOCENTRY), Convert.ToInt16(LINENUM), ITEM, QTY, PRICE, AMT, CUSTOMER, CLOSEDAY2, PAYMENT2, PP1, PP2, STATUS, receivePlace, goalPlace, shipment, unloadCargo);
                        }


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
                DELOPOR(fmLogin.LoginID.ToString());
                GD6(opdf.FileName);

                System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                dataGridView1.DataSource = G1;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            System.Data.DataTable SI1 = GetSI1();
           
            if (SI1.Rows.Count > 0)
            {
                for (int i = 0; i <= SI1.Rows.Count - 1; i++)
                {
                    string SI = SI1.Rows[i]["SI"].ToString();
                    System.Data.DataTable SI2 = GetSI2(SI);
                    if (SI2.Rows.Count > 0)
                    {
                 
                            string NumberName = "SI" + DateTime.Now.ToString("yyyyMMdd");
                            SqlConnection Connection = new SqlConnection(strCnSP);
                            string AutoNum = util.GetAutoNumber(Connection, NumberName);

                            string KK = NumberName + AutoNum + "X";

                            string 貿易條件 = SI2.Rows[0]["TRADE"].ToString();

                            string 結關日 = SI2.Rows[0]["CLOSEDAY"].ToString();
                            string 運送方式 = SI2.Rows[0]["P1"].ToString();
                            string 貿易形式 = SI2.Rows[0]["P2"].ToString();
                            string receivePlace = SI2.Rows[0]["receivePlace"].ToString();
                            string goalPlace = SI2.Rows[0]["goalPlace"].ToString();
                            string shipment = SI2.Rows[0]["shipment"].ToString();
                            string unloadCargo = SI2.Rows[0]["unloadCargo"].ToString();
                
                            //KK
                            AddSHIPMAIN(KK, "達擎股份有限公司GD", "S0623-GD", 貿易條件, 結關日, "AUO", 運送方式, 貿易形式, receivePlace, goalPlace, shipment, unloadCargo);
                            MessageBox.Show("上傳成功 SI單號 : " + KK);

     
                            if (SI2.Rows.Count > 0)
                            {
                                for (int i2 = 0; i2 <= SI2.Rows.Count - 1; i2++)
                                {
                                    string ITEMCODE = SI2.Rows[i2]["ITEMCODE"].ToString();
                                    int LINENUM = Convert.ToInt16(SI2.Rows[i2]["LINENUM"]);
                                    int QTY = Convert.ToInt16(SI2.Rows[i2]["QTY"]);
                                    decimal PRICE = Convert.ToDecimal(SI2.Rows[i2]["PRICE"]);
                                    decimal AMT = Convert.ToDecimal(SI2.Rows[i2]["AMT"]);
                                    string REMARK = SI2.Rows[i2]["CUSTOMER"].ToString();
                                    string DOCENTRY = SI2.Rows[i2]["DOCENTRY"].ToString();
                                    string STATUS = SI2.Rows[i2]["STATUS"].ToString();
                                    string DESC = "";

                                    //,receivePlace,goalPlace,shipment,unloadCargo
                                    System.Data.DataTable GD = GetDESC(ITEMCODE);
                                    if (GD.Rows.Count > 0)
                                    {
                                        DESC = GD.Rows[0][0].ToString();
                                    }



                                    AddSHIPITEM(KK, i2, DOCENTRY, LINENUM, "採購報價", ITEMCODE, DESC, QTY, PRICE, AMT, REMARK, STATUS);
                                }
                            }
                        }
                    }
                }
            
            
        }

        public void AddSHIPMAIN(string ShippingCode, string CardName, string CardCode, string TradeCondition, string CLOSEDAY, string BRAND, string receiveDay, string boardCountNo, string receivePlace, string goalPlace, string shipment, string unloadCargo)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_MAIN(ShippingCode,CardName,CardCode,TradeCondition,CLOSEDAY,BRAND,receiveDay,boardCountNo,receivePlace,goalPlace,shipment,unloadCargo) values(@ShippingCode,@CardName,@CardCode,@TradeCondition,@CLOSEDAY,@BRAND,@receiveDay,@boardCountNo,@receivePlace,@goalPlace,@shipment,@unloadCargo)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@TradeCondition", TradeCondition));
            command.Parameters.Add(new SqlParameter("@CLOSEDAY", CLOSEDAY));
            command.Parameters.Add(new SqlParameter("@BRAND", BRAND));

            command.Parameters.Add(new SqlParameter("@receiveDay", receiveDay));
            command.Parameters.Add(new SqlParameter("@boardCountNo", boardCountNo));
            command.Parameters.Add(new SqlParameter("@receivePlace", receivePlace));
            command.Parameters.Add(new SqlParameter("@goalPlace", goalPlace));
            command.Parameters.Add(new SqlParameter("@shipment", shipment));
            command.Parameters.Add(new SqlParameter("@unloadCargo", unloadCargo));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void AddSHIPITEM(string ShippingCode, int SeqNo, string Docentry, int linenum, string ItemRemark, string ItemCode, string Dscription, int Quantity, decimal ItemPrice, decimal ItemAmount, string Remark, string STATUS)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_ITEM(ShippingCode,SeqNo,Docentry,linenum,ItemRemark,ItemCode,Dscription,Quantity,ItemPrice,ItemAmount,Remark,STATUS) values(@ShippingCode,@SeqNo,@Docentry,@linenum,@ItemRemark,@ItemCode,@Dscription,@Quantity,@ItemPrice,@ItemAmount,@Remark,@STATUS)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Dscription", Dscription));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@ItemPrice", ItemPrice));
            command.Parameters.Add(new SqlParameter("@ItemAmount", ItemAmount));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
            command.Parameters.Add(new SqlParameter("@STATUS", STATUS));


            
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void APOPOR_Load(object sender, EventArgs e)
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

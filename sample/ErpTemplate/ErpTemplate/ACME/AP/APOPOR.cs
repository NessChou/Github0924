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
    public partial class APOPOR : Form
    {
       string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
       string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
       string FA = "acmesql98";
        public APOPOR()
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
                               DELOPOR(fmLogin.LoginID.ToString());
                               GD5(opdf.FileName);

                               textBox1.Text = "DD";

                               System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                                   if (G2.Rows.Count > 0)
                                   {
                                       int l = 0;
                                       for (int i = 0; i <= G2.Rows.Count - 1; i++)
                                       {
                                           l++;

                                           string T1 = G2.Rows[i][0].ToString();
                                           System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());
                                        
                                           if (G3.Rows.Count > 0)
                                           {
                                               int h = -1;
                                               for (int s = 0; s <= G3.Rows.Count - 1; s++)
                                               {
                                                   h++;

                                                   if (h >= 20)
                                                   {

                                                       l++;
                                                       h = 0;
                                                   }

                                                   UPOPOR(l,h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                                               }
                                           }
                                       }



                                   System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                                   dataGridView1.DataSource = G1;
                                   System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                                                     label1.Text = "數量 : "+GG2.Rows[0]["QTY"].ToString();
                                                     label2.Text = "金額 : "+GG2.Rows[0]["AMT"].ToString();
                               
                               }
                           }
                       //}
                       //catch (Exception ex)
                       //{
                       //    MessageBox.Show(ex.Message);
                       //}
        }

        private void D1(string CS)
        {
            UPOPORD(0, 0, fmLogin.LoginID.ToString());
            System.Data.DataTable G2 = null;
            int ROW = 0;
            if ((textBox1.Text == "GD" || (listBox1.Items.Count == 0)) )
            {
                G2 = GetOPOR2(fmLogin.LoginID.ToString());
            }

            else
            {
                G2 = GetOPOR2D(CS, fmLogin.LoginID.ToString());
            }
            if (textBox1.Text == "DD" )
            {
                ROW = 20;
            }

            if (textBox1.Text == "GD" )
            {
                ROW = 20;
            }

            if (textBox1.Text == "TV") 
            {
                ROW = Convert.ToInt32(GetOPOR2TV(fmLogin.LoginID.ToString()).Rows[0][0].ToString());//全部同PO
            }

            if (textBox1.Text == "PID")
            {
                ROW = 20;
            }




            if (G2.Rows.Count > 0)
            {
                int l = 0;
                for (int i = 0; i <= G2.Rows.Count - 1; i++)
                {
                    l++;

                    string T1 = G2.Rows[i][0].ToString();
                    System.Data.DataTable G3 = null;
                    if (textBox1.Text == "GD" || (listBox1.Items.Count == 0))
                    {
                        G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());
                    }
                    else
                    {
                        G3 = GetOPOR3D(CS, T1, fmLogin.LoginID.ToString());
                    }

                    if (G3.Rows.Count > 0)
                    {
                        int h = -1;
                        for (int s = 0; s <= G3.Rows.Count - 1; s++)
                        {
                            h++;

                            if (h >= ROW)
                            {

                                l++;
                                h = 0;
                            }

                            UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                        }
                    }
                }


            }
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

                            ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK, "S0001-DD", fmLogin.LoginID.ToString(),"",SITE,ARRIVE);
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
        public static System.Data.DataTable GetOPOR(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR WHERE USERS=@USERS ORDER BY DOCENTRY,DOC,ID");

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
        public System.Data.DataTable GetData(string sql)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(sql);

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
        public static System.Data.DataTable CheckItemCodeExist(string ItemCode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM OITM WHERE ItemCode=@ItemCode");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
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
        public static System.Data.DataTable GetOPOR2TV(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  COUNT(DOCENTRY) FROM AP_OPOR WHERE USERS=@USERS AND DOCENTRY >0  GROUP BY DOCENTRY ORDER BY DOCENTRY ");

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
        public static System.Data.DataTable GetOPOR3(string DOCENTRY, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPOR WHERE DOCENTRY=@DOCENTRY AND DOCENTRY >0 AND USERS=@USERS ");

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

        public static System.Data.DataTable GetSI1(string cs)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [SITE],ARRIVE 到貨地 FROM AP_OPOR 　WHERE USERS=@USERS");
            if (!String.IsNullOrEmpty(cs))
            {
                sb.Append(" AND ID IN ( " + cs + ") ");
            }
            sb.Append("  GROUP BY [SITE],ARRIVE");
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
        public static System.Data.DataTable GetSI2(string SITE, string ARRIVE, string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR 　WHERE USERS=@USERS AND [SITE]=@SITE AND ARRIVE =@ARRIVE ");
            if (!String.IsNullOrEmpty(cs))
            {
                sb.Append(" AND ID IN ( " + cs + ") ");
            }
            sb.Append(" ORDER BY SAPDOC,LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@SITE", SITE));
            command.Parameters.Add(new SqlParameter("@ARRIVE", ARRIVE));
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
        public static System.Data.DataTable GetTCONItemCode(string TCON)
        {
            SqlConnection MyConnection = globals.shipConnection; ;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ItemCode FROM OITM　WHERE ItemName LIKE @TCON AND FROZENFOR = 'N' AND CANCELED = 'N'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TCON", "%" + TCON + "%"));
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
            SqlCommand command = new SqlCommand("DELETE AP_OPOR WHERE USERS=@USERS ", connection);
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


        public void ADDOPOR(int DOCENTRY, string ITEMCODE, decimal QTY, decimal PRICE, decimal AMT, string REMARK, string CARDCODE, string USERS, string OCARDCODE,string SITE,string ARRIVE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPOR(DOCENTRY,ITEMCODE,QTY,PRICE,AMT,REMARK,CARDCODE,USERS,OCARDCODE,SITE,ARRIVE) values(@DOCENTRY,@ITEMCODE,@QTY,@PRICE,@AMT,@REMARK,@CARDCODE,@USERS,@OCARDCODE,@SITE,@ARRIVE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@OCARDCODE", OCARDCODE));
            command.Parameters.Add(new SqlParameter("@SITE", SITE));
            command.Parameters.Add(new SqlParameter("@ARRIVE", ARRIVE));
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
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            if (globals.UserID == "nesschou") 
            {
                MessageBox.Show("確認是否為測試區");
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


                if (listBox1.Items.Count == 0)
                {
                    D1("");

                    System.Data.DataTable G2 = GetOPOR4(fmLogin.LoginID.ToString());
                    if (G2.Rows.Count > 0)
                    {

                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {
                            SAPbobsCOM.Documents oPURCH = null;
                            oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                            string T1 = G2.Rows[n][0].ToString();

                            System.Data.DataTable G3 = GetOPOR5(T1, fmLogin.LoginID.ToString());
                           
                            if (G3.Rows.Count > 0)
                            {
                                oPURCH.CardCode = G3.Rows[0]["CARDCODE"].ToString();
                                oPURCH.DocCurrency = "USD";
                                oPURCH.VatPercent = 0;
                                System.Data.DataTable G1 = GetOPOR7();
                                if (G1.Rows.Count > 0)
                                {
                                    oPURCH.DocRate = Convert.ToDouble(G1.Rows[0][0]);
                                }
                                for (int s = 0; s <= G3.Rows.Count - 1; s++)
                                {
                                    System.Data.DataTable table = CheckItemCodeExist(G3.Rows[s]["ItemCode"].ToString());
                                    System.Data.DataTable dt = GetTCONItemCode(G3.Rows[s]["ItemCode"].ToString());
                                    if (table.Rows.Count > 0 || dt.Rows.Count > 0)
                                    {
                                        string ITEMCODE = G3.Rows[s]["ITEMCODE"].ToString();
                                        string OCARDCODE = G3.Rows[s]["OCARDCODE"].ToString();
                                        double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                                        double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                                        oPURCH.Lines.WarehouseCode = "TW017";
                                        oPURCH.Lines.ItemCode = ITEMCODE;
                                        oPURCH.Lines.Quantity = QTY;
                                        oPURCH.Lines.Price = PRICE;
                                        oPURCH.Lines.VatGroup = "AP0%";
                                        oPURCH.Lines.Currency = "USD";
                                        if (ITEMCODE == "KTCAU43TX.00102")
                                        {
                                            ITEMCODE = "KTCAU43TX.00162";
                                        }
                                        if (ITEMCODE == "ACME00004.00004")
                                        {
                                            oPURCH.Lines.ItemDescription = OCARDCODE;
                                        }
                                        if (textBox1.Text == "PID" || textBox1.Text== "TV")
                                        {
                                            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
                                            oPURCH.DocDate = LastDay;
                                            oPURCH.Lines.ShipDate = LastDay;
                                        }

                                        oPURCH.Lines.Add();
                                    }
                                    else
                                    {
                                        //MessageBox.Show(G3.Rows[s]["ItemCode"].ToString() + "此料號不存在");
                                        //改一次性料號輸入
                                        string ITEMCODE = "ACME00001.00001";
                                        string OCARDCODE = G3.Rows[s]["OCARDCODE"].ToString();
                                        double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                                        double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                                        oPURCH.Lines.WarehouseCode = "TW017";
                                        oPURCH.Lines.ItemCode = "ACME00001.00001";
                                        oPURCH.Lines.Quantity = QTY;
                                        oPURCH.Lines.Price = PRICE;
                                        oPURCH.Lines.VatGroup = "AP0%";
                                        oPURCH.Lines.Currency = "USD";
                                        if (ITEMCODE == "ACME00004.00004")
                                        {
                                            oPURCH.Lines.ItemDescription = OCARDCODE;
                                        }
                                        if (textBox1.Text == "PID" || textBox1.Text == "TV")
                                        {
                                            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
                                            oPURCH.DocDate = LastDay;
                                            oPURCH.Lines.ShipDate = LastDay;
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

                                    System.Data.DataTable GG3 = GetOPOR5(T1, fmLogin.LoginID.ToString());

                                    if (GG3.Rows.Count > 0)
                                    {
                                        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                        {

                                            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                            string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                            UPDATEOPOR(OWTR, LINENUM, REMARK);

                                            UPDATEAPOPOR(T1, LINENUM, OWTR);
                                        }

                                    }
                                }

                            }

                        }

                    }
                }
                else
                {

                    ArrayList al = new ArrayList();

                    for (int i2 = 0; i2 <= listBox1.Items.Count - 1; i2++)
                    {
                        al.Add(listBox1.Items[i2].ToString());
                    }
                    StringBuilder sb = new StringBuilder();



                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);

                    D1(sb.ToString());

                    System.Data.DataTable G2 = GetOPOR4D(sb.ToString(), fmLogin.LoginID.ToString());

                    if (G2.Rows.Count > 0)
                    {

                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {
                            SAPbobsCOM.Documents oPURCH = null;
                            oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                            string T1 = G2.Rows[n][0].ToString();
                            System.Data.DataTable G3 = GetOPOR5ID(sb.ToString(),T1, fmLogin.LoginID.ToString());

                            if (G3.Rows.Count > 0)
                            {
                                oPURCH.CardCode = G3.Rows[0]["CARDCODE"].ToString();
                                oPURCH.DocCurrency = "USD";
                                oPURCH.VatPercent = 0;
                                System.Data.DataTable G1 = GetOPOR7();
                                if (G1.Rows.Count > 0)
                                {
                                    oPURCH.DocRate = Convert.ToDouble(G1.Rows[0][0]);
                                }

                                for (int s = 0; s <= G3.Rows.Count - 1; s++)
                                {
                                    string ITEMCODE = G3.Rows[s]["ITEMCODE"].ToString();
                                    double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                                    double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                                    oPURCH.Lines.WarehouseCode = "TW017";
                                    oPURCH.Lines.ItemCode = ITEMCODE;
                                    oPURCH.Lines.Quantity = QTY;
                                    oPURCH.Lines.Price = PRICE;
                                    oPURCH.Lines.VatGroup = "AP0%";
                                    oPURCH.Lines.Currency = "USD";

                                    oPURCH.Lines.Add();

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

                                    System.Data.DataTable GG3 = GetOPOR5ID(sb.ToString(),T1, fmLogin.LoginID.ToString());

                                    if (GG3.Rows.Count > 0)
                                    {
                                        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                        {

                                            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                            string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                            UPDATEOPOR(OWTR, LINENUM, REMARK);

                                            UPDATEAPOPORD(sb.ToString(), T1,LINENUM, OWTR);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }




            }
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }


            System.Data.DataTable GG1 = GetOPOR(fmLogin.LoginID.ToString());
            dataGridView1.DataSource = GG1;
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
        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    listBox1.Items.Clear();
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];

                        listBox1.Items.Add(row.Cells["ID"].Value.ToString());

                    }


                    ArrayList al = new ArrayList();

                    for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                    {
                        al.Add(listBox1.Items[i].ToString());
                    }
                    StringBuilder sb = new StringBuilder();



                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);

                    System.Data.DataTable GG2 = GetOPOR6IN(sb.ToString(), fmLogin.LoginID.ToString());
                    label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                    label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

                }
                else
                {
                    listBox1.Items.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
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
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK1;
                string REMARK2;
                string P1;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                ITEMCODE = range.Text.ToString().Trim();

                if (ITEMCODE != "")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    LINE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    VER = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
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

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    REMARK1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    REMARK2 = range.Text.ToString().Trim();

                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {

                            string ITEM = "";
                            System.Data.DataTable F1 = GetITEMCODEGD(ITEMCODE, GRADE, VER);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();
                            }
                            else
                            {
                                ITEM = "ACME00004.00004";
                            }


                            ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK1 + " " + REMARK2, "S0623-GD", fmLogin.LoginID.ToString(), ITEMCODE + "." + VER.Substring(0, 1), "", "");
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

                textBox1.Text = "GD";

                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());

                        if (G3.Rows.Count > 0)
                        {
                            int h = -1;
                            for (int s = 0; s <= G3.Rows.Count - 1; s++)
                            {
                                h++;

                                if (h >= 10)
                                {

                                    l++;
                                    h = 0;
                                }

                                UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }



                    System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                    dataGridView1.DataSource = G1;
                    System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                    label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                    label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            string CS = "";
            if (listBox1.Items.Count != 0)
            {
                ArrayList al = new ArrayList();

                for (int i2 = 0; i2 <= listBox1.Items.Count - 1; i2++)
                {
                    al.Add(listBox1.Items[i2].ToString());
                }

                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);

                CS = sb.ToString();
            }

            System.Data.DataTable SI1 = GetSI1(CS);
           
            if (SI1.Rows.Count > 0)
            {
                for (int i = 0; i <= SI1.Rows.Count - 1; i++)
                {
                    string SITE = SI1.Rows[i]["SITE"].ToString();
                    string ARR = SI1.Rows[i]["到貨地"].ToString();
                    string NumberName = "SI" + DateTime.Now.ToString("yyyyMMdd");
                    SqlConnection Connection = new SqlConnection(strCnSP);
                    string AutoNum = util.GetAutoNumber(Connection, NumberName);

                    string KK = NumberName + AutoNum + "X";
                    string 收貨地 = "";
                    string 目的地 = "";
                    string 貿易條件 = "";
                    string 運送方式 = "";
                    string 貿易形式 = "";
                    if (SITE == "S11")
                    {
                        收貨地 = "XIAMEN";
                    }
                    else if (SITE == "S02")
                    {
                        收貨地 = "SUZHUO";
                    }
                    else if (SITE == "Z68")
                    {
                        收貨地 = "HEFEI, CHINA";
                    }
                    else if (SITE == "Z19")
                    {
                        收貨地 = "SUZHUO ";
                    }
                    else if (SITE == "Z0M")
                    {
                        收貨地 = "FUQING, CHINA";
                    }

                    if (ARR == "宏高")
                    {
                        目的地 = "SHENZHEN, CHINA";
                        貿易條件 = "CIP SHENZHEN";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }
                    else if (ARR == "蘇宏高")
                    {
                        目的地 = "SUZHOU, CHINA";
                        貿易條件 = "CIP SUZHOU";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }
                    else if (ARR == "鉅航")
                    {
                        目的地 = "SHENZHEN, CHINA";
                        貿易條件 = "CIP SHENZHEN";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }
                    else if (ARR == "武漢")
                    {
                        目的地 = "WUHAN, CHINA";
                        貿易條件 = "CIP WUHAN";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }
                    else if (ARR == "廈門")
                    {
                        目的地 = "XIAMEN, CHINA";
                        貿易條件 = "CIP XIAMEN";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }
                    else if (ARR == "回台")
                    {
                        目的地 = "TAOYUAN, TAIWAN";
                        貿易條件 = "CIP TAIWAN";
                        運送方式 = "SEA";
                        貿易形式 = "進口";
                    }
                    else if (ARR == "香港")
                    {
                        目的地 = "HONG KONG";
                        貿易條件 = "CIP HONG KONG";
                        運送方式 = "TRUCK";
                        貿易形式 = "三角";
                    }

                    if (SITE == "M02")
                    {
                        收貨地 = "TAOYUAN, TAIWAN";
                        目的地 = "SHENZHEN, CHINA";
                        貿易條件 = "CIP SHENZHEN W/CC";
                        運送方式 = "SEA";
                        貿易形式 = "出口";
                    }
                    if (SITE == "M11")
                    {
                        收貨地 = "TAICHUNG, TAIWAN";
                        目的地 = "SHENZHEN, CHINA";
                        貿易條件 = "CIP SHENZHEN W/CC";
                        運送方式 = "SEA";
                        貿易形式 = "出口";
                    }
                    //KK
                    AddSHIPMAIN(KK, "友達光電股份有限公司DD", "S0001-DD", 貿易條件, 收貨地, 收貨地, 目的地, 目的地, 運送方式, 貿易形式);
                    MessageBox.Show("上傳成功 SI單號 : " + KK);

                    System.Data.DataTable SI2 = GetSI2(SITE,ARR,CS);

                    if (SI2.Rows.Count > 0)
                    {
                        for (int i2 = 0; i2 <= SI2.Rows.Count - 1; i2++)
                        {
                            string SAPDOC = SI2.Rows[i2]["SAPDOC"].ToString();
                            int LINENUM = Convert.ToInt16(SI2.Rows[i2]["LINENUM"]);
                            int QTY = Convert.ToInt16(SI2.Rows[i2]["QTY"]);
                            decimal PRICE = Convert.ToDecimal(SI2.Rows[i2]["PRICE"]);
                            decimal AMT = Convert.ToDecimal(SI2.Rows[i2]["AMT"]);
                            string ITEMCODE = SI2.Rows[i2]["ITEMCODE"].ToString();
                            string REMARK = SI2.Rows[i2]["REMARK"].ToString();
                            string DESC = "";
                            System.Data.DataTable GD = GetDESC(ITEMCODE);
                            if (GD.Rows.Count > 0)
                            {
                                DESC = GD.Rows[0][0].ToString();
                            }



                            AddSHIPITEM(KK, i2, SAPDOC, LINENUM, "採購訂單", ITEMCODE, DESC, QTY, PRICE, AMT, REMARK);
                        }
                    }
                }
            
            }
        }
        public void AddSHIPMAIN(string ShippingCode, string CardName, string CardCode, string TradeCondition, string receivePlace, string shipment, string goalPlace, string unloadCargo, string receiveDay, string boardCountNo)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_MAIN(ShippingCode,CardName,CardCode,TradeCondition,receivePlace,shipment,goalPlace,unloadCargo,receiveDay,boardCountNo) values(@ShippingCode,@CardName,@CardCode,@TradeCondition,@receivePlace,@shipment,@goalPlace,@unloadCargo,@receiveDay,@boardCountNo)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@TradeCondition", TradeCondition));
            command.Parameters.Add(new SqlParameter("@receivePlace", receivePlace));
            command.Parameters.Add(new SqlParameter("@shipment", shipment));
            command.Parameters.Add(new SqlParameter("@goalPlace", goalPlace));
            command.Parameters.Add(new SqlParameter("@unloadCargo", unloadCargo));
            command.Parameters.Add(new SqlParameter("@receiveDay", receiveDay));
            command.Parameters.Add(new SqlParameter("@boardCountNo", boardCountNo));
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
        public void AddSHIPITEM(string ShippingCode, int SeqNo, string Docentry, int linenum, string ItemRemark, string ItemCode, string Dscription, int Quantity, decimal ItemPrice,decimal ItemAmount, string Remark)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_ITEM(ShippingCode,SeqNo,Docentry,linenum,ItemRemark,ItemCode,Dscription,Quantity,ItemPrice,ItemAmount,Remark) values(@ShippingCode,@SeqNo,@Docentry,@linenum,@ItemRemark,@ItemCode,@Dscription,@Quantity,@ItemPrice,@ItemAmount,@Remark)", Connection);
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

        private void btnTVExcel_Click(object sender, EventArgs e)
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
                TV(opdf.FileName);
                textBox1.Text = "TV";
                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());
                        int ROW = Convert.ToInt32(GetOPOR2TV(fmLogin.LoginID.ToString()).Rows[0][0].ToString());//全部同PO
                        if (G3.Rows.Count > 0)
                        {
                            int h = -1;
                            for (int s = 0; s <= G3.Rows.Count - 1; s++)
                            {
                                h++;

                                if (h >= ROW)
                                {

                                    l++;
                                    h = 0;
                                }

                                UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }
                }


                System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                dataGridView1.DataSource = G1;
                System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

            }
        }
        private void TV(string ExcelFile)
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
                string ITEMCODE;
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

                    decimal n;
                    
                    string partno = "";//產品大項分類
                    string model = "";//原廠型號共計8碼
                    string grade = "";
                    string partno2 = "";//原廠Ver. & PART No.共計3碼
                    string fab = "";//交易別 & 產地別 & 產品特殊狀況描述
                    //model
                    switch (AuoPN.Substring(0, 2))
                    {
                        case "91":
                            //若為91則改為O
                            model = "O" + ModelName.Substring(1, 8);
                            ModelName = "O" + ModelName.Substring(1, ModelName.Length - 1)+"."+ AuoPN.Split('.')[2].Substring(0,1);
                            break;
                        default:
                            model = ModelName.Split('.')[0].Substring(0, 9);
                            ModelName = ModelName + "." + AuoPN.Split('.')[2].Substring(0, 1);
                            break;
                    }
                    //grade
                    switch (Grade)
                    {
                        case "A":
                            grade = "1";
                            break;
                        case "A-":
                            grade = "5";
                            break;
                        case "V":
                            grade = "3";
                            break;

                    }
                    //AMT
                    string s = "";
                    if (ModelName.Substring(1, 3) == "320" || ModelName.Substring(1, 3) == "430" || ModelName.Substring(1, 3) == "750" || ModelName.Substring(1, 3) == "850")
                    {
                        s = ModelName.Substring(1, 6);//32,43才有差別要取到六碼
                    }
                    else 
                    {
                        s = ModelName.Substring(1, 3);//其餘三碼判斷以免增加後三碼(QVR、QVN就要修改程式)                        )
                    }
                    
                    Price = GETPrice(s, Grade);

                    int x, y;
                    x = int.Parse(Price.Split(',')[1]);
                    y = int.Parse(Price.Split(',')[0]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                    range.Select();
                    Price = range.Text.ToString().Trim();


                    if (decimal.TryParse(Price, out n))
                    {
                        PRICE = Convert.ToDecimal(Price);
                    }
                    else
                    {
                        PRICE = 0;
                    }
                    if (decimal.TryParse(Apr, out n))
                    {
                        QTY = Convert.ToDecimal(Apr);
                    }
                    else
                    {
                        QTY = 0;
                    }

                    AMT = PRICE * QTY;
                    //partno2
                    partno2 = AuoPN.Split('.')[2];
                    //fab
                    switch (FAB)
                    {
                        case "M02":
                        case "M11":
                        case "L8B":
                            fab = "1";
                            break;
                        default:
                            fab = "2";
                            break;

                    }
                    ITEMCODE = model + "." + grade + partno2 + fab;
                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {
                            ADDOPOR(Convert.ToInt16(LINE), ITEMCODE, QTY, PRICE, AMT, Remarks, "S0001-TV", fmLogin.LoginID.ToString(), ModelName, "", "");
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    //處理tcon
                    if (Tcon != "" && Tcon != "TCON PN" && Tqty != "") 
                    {
                        System.Data.DataTable dt = GetTCONItemCode(Tcon);
                        string Location = "";
                        if (dt.Rows.Count > 0)
                        {
                            ITEMCODE = dt.Rows[0]["ItemCode"].ToString().Substring(0, dt.Rows[0]["ItemCode"].ToString().Length - 1) + fab;

                            switch (Tcon.Substring(0, 5))
                            {
                                case "55.32":
                                    Location = "14,6";
                                    break;
                                case "55.43":
                                    Location = "14,9";
                                    break;
                                case "55.50":
                                    Location = "14,10";
                                    break;
                                case "55.55":
                                    Location = "14,11";
                                    break;
                                case "55.65":
                                    Location = "14,12";
                                    break;
                                case "55.75":
                                    Location = "14,13";//先抓4K
                                    break;
                                case "55.85":
                                    Location = "14,15";//先抓4K
                                    break;

                            }
                            x = int.Parse(Location.Split(',')[1]);
                            y = int.Parse(Location.Split(',')[0]);
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                            range.Select();
                            Price = range.Text.ToString().Trim();

                            PRICE = decimal.Parse(Price);
                            QTY = Tqty == "" ? QTY : int.Parse(Tqty);
                            AMT = PRICE * QTY;
                            ModelName = ITEMCODE.Split('.')[0] + "." + ITEMCODE.Split('.')[1].Substring(1, 1);

                           
                        }
                        else 
                        {
                            //一次性料號
                            ITEMCODE = "ACME00001.00001";
                            QTY = Tqty == "" ? QTY : int.Parse(Tqty);
                            switch (Tcon.Substring(0, 5))
                            {
                                case "55.32":
                                    Location = "14,6";
                                    break;
                                case "55.43":
                                    Location = "14,9";
                                    break;
                                case "55.50":
                                    Location = "14,10";
                                    break;
                                case "55.55":
                                    Location = "14,11";
                                    break;
                                case "55.65":
                                    Location = "14,12";
                                    break;
                                case "55.75":
                                    Location = "14,13";//先抓4K
                                    break;
                                case "55.85":
                                    Location = "14,15";//先抓4K
                                    break;

                            }
                            x = int.Parse(Location.Split(',')[1]);
                            y = int.Parse(Location.Split(',')[0]);
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                            range.Select();
                            Price = range.Text.ToString().Trim();

                            PRICE = decimal.Parse(Price);
                            AMT = PRICE * QTY;
                            ModelName = ITEMCODE.Split('.')[0] + "." + ITEMCODE.Split('.')[1].Substring(1, 1);//可能會沒有這筆tcon料號 另外再手動上傳
                        }
                        try
                        {
                            if (!String.IsNullOrEmpty(ITEMCODE))
                            {
                                ADDOPOR(Convert.ToInt16(LINE), ITEMCODE, QTY, PRICE, AMT, Remarks, "S0001-TV", fmLogin.LoginID.ToString(), ModelName, "", "");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
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
        private void PriceImport(string ExcelFile)
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
            }
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
        private string GETPrice(string Model, string Grade)
        {
            string Location = "";
            int PRICE = 0;
            if (Grade == "A")
            {

                switch (Model)
                {
                    case "215":
                        Location = "15,3";
                        break;
                    case "238":
                        Location = "15,4";
                        break;
                    case "320HVN":
                        Location = "15,6";
                        break;
                    case "320XVN":
                        Location = "15,5";
                        break;
                    case "390":
                        Location = "15,7";
                        break;
                    case "430HVN":
                        Location = "15,8";
                        break;
                    case "430QVN":
                        Location = "15,9";
                        break;
                    case "500":
                        Location = "15,10";
                        break;
                    case "550":
                        Location = "15,11";
                        break;
                    case "650":
                        Location = "15,12";
                        break;
                    case "750QVN":
                        Location = "15,13";
                        break;
                    case "750MVR":
                        Location = "15,14";
                        break;
                    case "850QVN":
                        Location = "15,15";
                        break;
                    case "850MVR":
                        Location = "15,16";
                        break;
                }
            }
            else if (Grade == "A-")
            {
                switch (Model)
                {
                    case "215":
                        Location = "16,3";
                        break;
                    case "238":
                        Location = "16,4";
                        break;
                    case "320HVN":
                        Location = "16,6";
                        break;
                    case "320XVN":
                        Location = "16,5";
                        break;
                    case "390":
                        Location = "16,7";
                        break;
                    case "430HVN":
                        Location = "16,8";
                        break;
                    case "430QVN":
                        Location = "16,9";
                        break;
                    case "500":
                        Location = "16,10";
                        break;
                    case "550":
                        Location = "16,11";
                        break;
                    case "650":
                        Location = "16,12";
                        break;
                    case "750QVN":
                        Location = "16,13";
                        break;
                    case "750MVR":
                        Location = "16,14";
                        break;
                    case "850QVN":
                        Location = "16,15";
                        break;
                    case "850MVR":
                        Location = "16,16";
                        break;
                }
            }
            else if (Grade == "V")
            {
                switch (Model)
                {
                    case "215":
                        Location = "17,3";
                        break;
                    case "238":
                        Location = "17,4";
                        break;
                    case "320HVN":
                        Location = "17,6";
                        break;
                    case "320XVN":
                        Location = "17,5";
                        break;
                    case "390":
                        Location = "17,7";
                        break;
                    case "430HVN":
                        Location = "17,8";
                        break;
                    case "430QVN":
                    case "430QVR":
                        Location = "17,9";
                        break;
                    case "500":
                        Location = "17,10";
                        break;
                    case "550":
                        Location = "17,11";
                        break;
                    case "650":
                        Location = "17,12";
                        break;
                    case "750QVN":
                    case "750QVR":
                        Location = "17,13";
                        break;
                    case "750MVR":
                        Location = "17,14";
                        break;
                    case "850QVN":
                    case "850QVR":
                        Location = "17,15";
                        break;
                    case "850MVR":
                        Location = "17,16";
                        break;
                }
            }
            return Location;
        }
        public enum APrice
        {
            //A等級
            O215HVN = 55,
            O238HVN = 59,
            O320XVN = 68,
            O320HVN = 67,
            O390XVN = 80,
            O430HVN = 103,
            O430QVN = 108,
            O500QVN = 138,
            O550QVN = 163,
            O650QVN = 205,
            O750QVN = 278,
            O750MVR = 278,
            O850QVN = 438,
            O850MVR = 438
        }
        public enum A_Price
        {
            //A-等級
            O215HVN = 44,
            O238HVN = 47,
            O320XVN = 60,
            O320HVN = 58,
            O390XVN = 70,
            O430HVN = 90,
            O430QVN = 94,
            O500QVN = 122,
            O550QVN = 143,
            O650QVN = 178,
            O750QVN = 242,
            O750MVR = 242,
            O850QVN = 382,
            O850MVR = 382
        }
        public enum  TCON
        {
            T32F = 5,
            T43F = 7,
            T43U = 10,
            T50U = 10,
            T55U = 10,
            T65U = 10,
            T75U4K = 10,
            T75U8K = 40,
            T85U4K = 10,
            T85U8K = 46,
        }
        private void btnPIDExcel_Click(object sender, EventArgs e)
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
                PID(opdf.FileName);

                textBox1.Text = "PID";

                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());

                        if (G3.Rows.Count > 0)
                        {
                            int h = -1;
                            for (int s = 0; s <= G3.Rows.Count - 1; s++)
                            {

                                UPOPOR(1, 0, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }
                }

                System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                dataGridView1.DataSource = G1;
                System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();
            }
        }
        private void PID(string ExcelFile)
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
                string ITEMCODE = "ACME00001.00001";
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

                string ModelVersion;
                string PartNo;
                string Grade;
                string FAB;
                string Price;
                string Remarks;
                string Apr;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                ModelVersion = range.Text.ToString().Trim();

                if (ModelVersion != "" && ModelVersion != "MODEL_VERSION")
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    PartNo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    Grade = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    FAB = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    Price = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    Remarks = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    Apr = range.Text.ToString().Trim();

                    decimal n;
                    if (decimal.TryParse(Price, out n))
                    {
                        PRICE = Convert.ToDecimal(Price);
                    }
                    else
                    {
                        PRICE = 0;
                    }
                    if (decimal.TryParse(Apr, out n))
                    {
                        QTY = Convert.ToDecimal(Apr);
                    }
                    else
                    {
                        QTY = 0;
                    }

                    AMT = PRICE * QTY;
                    string partno = "";//產品大項分類
                    string model = "";//原廠型號共計8碼
                    string grade = "";
                    string partno2 = "";//原廠Ver. & PART No.共計3碼
                    string fab = "";//交易別 & 產地別 & 產品特殊狀況描述

                    bool TconTag = false;
                    //model
                    switch (PartNo.Substring(0, 2)) 
                    {
                        case "91":
                            //若為91則改為O
                            model = "O" + ModelVersion.Split('.')[0].Substring(1, 8);
                            ModelVersion = "O" + ModelVersion.Substring(1, ModelVersion.Length-1);
                            break;
                        case "55":
                            string sql = "SELECT ITEMCODE FROM OITM WHERE U_MODEL LIKE '{0}'";
                            sql = string.Format(sql,PartNo);
                            System.Data.DataTable table = GetData(sql);
                            if (table.Rows.Count > 0)
                            {
                                ITEMCODE = table.Rows[0]["ITEMCODE"].ToString();
                            }
                            else 
                            {
                                ITEMCODE = "ACME00001.00001";
                            }
                            TconTag = true;
                            break;
                        default:
                            model = ModelVersion.Split('.')[0].Substring(0, 9);
                            break;
                    }
                    //grade
                    if (TconTag == false)
                    {
                        switch (Grade)
                        {
                            case "Z":
                                grade = "0";
                                break;
                            case "P":
                                grade = "1";
                                break;
                            case "N":
                                grade = "5";
                                break;

                        }
                        //partno2
                        partno2 = PartNo.Split('.')[2];
                        //fab
                        switch (FAB.Split('_')[1])
                        {
                            case "M02":
                            case "M11":
                            case "L8B":
                                fab = "1";
                                break;
                            default:
                                fab = "2";
                                break;

                        }
                        ITEMCODE = model + "." + grade + partno2 + fab;
                    }
                    else 
                    {
                        if (ITEMCODE != "ACME00001.00001") 
                        {
                            switch (FAB.Split('_')[1])
                            {
                                case "M02":
                                case "M11":
                                case "L8B":
                                    fab = "1";
                                    break;
                                default:
                                    fab = "2";
                                    break;
                            }
                            ITEMCODE = ITEMCODE.Substring(0, ITEMCODE.Length - 1);
                            ITEMCODE = ITEMCODE + fab;
                        }
                    }
                   
                   


                    
                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {
                            ADDOPOR(Convert.ToInt16(LINE), ITEMCODE, QTY, PRICE, AMT, Remarks, "S0623-PID", fmLogin.LoginID.ToString(), ModelVersion, "", "");
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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

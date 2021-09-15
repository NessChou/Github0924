using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Web.UI;
using System.Net.Mime;
using System.Threading.Tasks;
namespace ACME
{
    public partial class HRNO : Form
    {
        string strCn = "Data Source=10.10.1.45;Initial Catalog=89206602;Persist Security Info=True;User ID=ehrview;Password=viewehr";
        private string FileName;
        public HRNO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                if (comboBox1.Text == "超時出勤異常")
                {
                    DELCARDDOOR();
                    WriteExcelAP(FileName);
                }
                if (comboBox1.Text == "特休時數提醒")
                {
                    DELCARDDOOR2();
                    WriteExcelAP2(FileName);
                }
       

            }
        }
        public void DELCARDDOOR()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE HR_CARDDOOR", connection);

            command.CommandType = CommandType.Text;

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

        public void DELCARDDOOR2()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE HR_CARDDOOR2", connection);

            command.CommandType = CommandType.Text;

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
        public void AddCARDDOOR(string EMPID, string EMPCNAME, string FNAME, string LNAME, string GROUPCODE, string GROUPNAME, string EMPTYPE, string EMPDATE, string EMPCARD1, string EMPCARD2, string CARDTYPE, string CARDNUM, string CARDSTATUS, string CARDMEMO, string EMAIL, string MANAGER, string EMAIL2)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into HR_CARDDOOR(EMPID,EMPCNAME,FNAME,LNAME,GROUPCODE,GROUPNAME,EMPTYPE,EMPDATE,EMPCARD1,EMPCARD2,CARDTYPE,CARDNUM,CARDSTATUS,CARDMEMO,EMAIL,MANAGER,EMAIL2) values(@EMPID,@EMPCNAME,@FNAME,@LNAME,@GROUPCODE,@GROUPNAME,@EMPTYPE,@EMPDATE,@EMPCARD1,@EMPCARD2,@CARDTYPE,@CARDNUM,@CARDSTATUS,@CARDMEMO,@EMAIL,@MANAGER,@EMAIL2)", connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
            command.Parameters.Add(new SqlParameter("@EMPCNAME", EMPCNAME));
            command.Parameters.Add(new SqlParameter("@FNAME", FNAME));
            command.Parameters.Add(new SqlParameter("@LNAME", LNAME));
            command.Parameters.Add(new SqlParameter("@GROUPCODE", GROUPCODE));
            command.Parameters.Add(new SqlParameter("@GROUPNAME", GROUPNAME));
            command.Parameters.Add(new SqlParameter("@EMPTYPE", EMPTYPE));
            command.Parameters.Add(new SqlParameter("@EMPDATE", EMPDATE));
            command.Parameters.Add(new SqlParameter("@EMPCARD1", EMPCARD1));
            command.Parameters.Add(new SqlParameter("@EMPCARD2", EMPCARD2));
            command.Parameters.Add(new SqlParameter("@CARDTYPE", CARDTYPE));
            command.Parameters.Add(new SqlParameter("@CARDNUM", CARDNUM));
            command.Parameters.Add(new SqlParameter("@CARDSTATUS", CARDSTATUS));
            command.Parameters.Add(new SqlParameter("@CARDMEMO", CARDMEMO));

            command.Parameters.Add(new SqlParameter("@EMAIL", EMAIL));
            command.Parameters.Add(new SqlParameter("@MANAGER", MANAGER));
            command.Parameters.Add(new SqlParameter("@EMAIL2", EMAIL2));
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

        public void AddCARDDOOR2(string EMPID,string HNAME, string HDAY, string HD1, string HD2, string HS1, string HS2, string EMAIL, string MANAGER, string EMAIL2)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into HR_CARDDOOR2(EMPID,HNAME,HDAY,HD1,HD2,HS1,HS2,EMAIL,MANAGER,EMAIL2) values(@EMPID,@HNAME,@HDAY,@HD1,@HD2,@HS1,@HS2,@EMAIL,@MANAGER,@EMAIL2)", connection);
           
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
            command.Parameters.Add(new SqlParameter("@HNAME", HNAME));
            command.Parameters.Add(new SqlParameter("@HDAY", HDAY));
            command.Parameters.Add(new SqlParameter("@HD1", HD1));
            command.Parameters.Add(new SqlParameter("@HD2", HD2));
            command.Parameters.Add(new SqlParameter("@HS1", HS1));
            command.Parameters.Add(new SqlParameter("@HS2", HS2));
            command.Parameters.Add(new SqlParameter("@EMAIL", EMAIL));
            command.Parameters.Add(new SqlParameter("@MANAGER", MANAGER));
            command.Parameters.Add(new SqlParameter("@EMAIL2", EMAIL2));

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

        public System.Data.DataTable GETETH(string EMPLOYEE_ID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T0.EMPLOYEE_CNAME 員工, ");
            sb.Append(" CASE WHEN  T0.EMPLOYEE_CNAME=  T2.EMPLOYEE_CNAME THEN  T4.EMPLOYEE_CNAME ELSE T2.EMPLOYEE_CNAME  END MANAGER  ");
            sb.Append(" ,T0.EMPLOYEE_EMAIL_1 EMAIL,CASE WHEN  T0.EMPLOYEE_CNAME=  T2.EMPLOYEE_CNAME THEN  T4.EMPLOYEE_EMAIL_1 ELSE T2.EMPLOYEE_EMAIL_1  END  EMAIL2 FROM vwZZ_EMPLOYEE T0  ");
            sb.Append(" LEFT JOIN vwZZ_DEPARTMENT T1 ON (T0.DEPARTMENT_ID =T1.DEPARTMENT_ID) ");
            sb.Append(" LEFT JOIN vwZZ_EMPLOYEE T2 ON (T1.DEPARTMENT_LEADER_ID  =T2.EMPLOYEE_ID) ");
            sb.Append(" LEFT JOIN vwZZ_DEPARTMENT T3 ON (T1.PART_DEPARTMENT_ID =T3.DEPARTMENT_ID) ");
            sb.Append(" LEFT JOIN vwZZ_EMPLOYEE T4 ON (T3.DEPARTMENT_LEADER_ID  =T4.EMPLOYEE_ID) ");
            sb.Append(" WHERE T0.EMPLOYEE_CNAME  NOT IN ('許心如','黃舉昇')");
            sb.Append("  AND T0.EMPLOYEE_NO=@EMPLOYEE_ID");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPLOYEE_ID", EMPLOYEE_ID));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GETETHID(string EMPLOYEE_CNAME)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT EMPLOYEE_NO ID  FROM vwZZ_EMPLOYEE WHERE EMPLOYEE_CNAME=@EMPLOYEE_CNAME AND DEPARTMENT_CNAME <> '離職人員' ORDER BY EMPLOYEE_ID DESC ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPLOYEE_CNAME", EMPLOYEE_CNAME));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        private void WriteExcelAP(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
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


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string EMPID = "";
                string EMPCNAME = "";
                string FNAME = "";
                string LNAME = "";
                string GROUPCODE = "";
                string GROUPNAME = "";
                string EMPTYPE = "";
                string EMPDATE = "";
                string EMPCARD1 = "";
                string EMPCARD2 = "";
                string CARDTYPE = "";
                string CARDNUM = "";
                string CARDSTATUS = "";
                string CARDMEMO = "";

                string DEMPID = "";
                string DEMPCNAME = "";
                string DFNAME = "";
                string DLNAME = "";
                string DGROUPCODE = "";
                string DGROUPNAME = "";
                string DEMPTYPE = "";

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    EMPID = range.Text.ToString().Trim();

                    if (EMPID == "")
                    {
                        EMPID = DEMPID;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    EMPCNAME = range.Text.ToString().Trim();
                    if (EMPCNAME == "")
                    {
                        EMPCNAME = DEMPCNAME;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    FNAME = range.Text.ToString().Trim();
                    if (FNAME == "")
                    {
                        FNAME = DFNAME;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    LNAME = range.Text.ToString().Trim();
                    if (LNAME == "")
                    {
                        LNAME = DLNAME;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    GROUPCODE = range.Text.ToString().Trim();
                    if (GROUPCODE == "")
                    {
                        GROUPCODE = DGROUPCODE;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    GROUPNAME = range.Text.ToString().Trim();
                    if (GROUPNAME == "")
                    {
                        GROUPNAME = DGROUPNAME;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    EMPTYPE = range.Text.ToString().Trim();
                    if (EMPTYPE == "")
                    {
                        EMPTYPE = DEMPTYPE;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    EMPDATE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    EMPCARD1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    EMPCARD2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    CARDTYPE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    CARDNUM = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    CARDSTATUS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    CARDMEMO = range.Text.ToString().Trim();

                    System.Data.DataTable H1 = GETETH(EMPID);
                    if (H1.Rows.Count > 0)
                    {
                        if (CARDSTATUS != "")
                        {
                            string EMAIL = H1.Rows[0]["EMAIL"].ToString();
                            if (EMAIL != "")
                            {
                                string MANAGER = H1.Rows[0]["MANAGER"].ToString();

                                string EMAIL2 = H1.Rows[0]["EMAIL2"].ToString();
                                if (MANAGER == "許心如" || MANAGER == "黃舉昇")
                                {
                                    EMAIL2 = "";
                                }
                                if (MANAGER == "陳思怡")
                                {
                                    MANAGER = "陳彥琪";
                                    EMAIL2 = "applechen@acmepoint.com";
                                }
                                AddCARDDOOR(EMPID, EMPCNAME, FNAME, LNAME, GROUPCODE, GROUPNAME, EMPTYPE, EMPDATE, EMPCARD1, EMPCARD2, CARDTYPE, CARDNUM, CARDSTATUS, CARDMEMO, EMAIL, MANAGER, EMAIL2);
                            }
                        }
                    }


                    DEMPID = EMPID;
                    DEMPCNAME = EMPCNAME;
                    DFNAME = FNAME;
                    DLNAME = LNAME;
                    DGROUPCODE = GROUPCODE;
                    DGROUPNAME = GROUPNAME;
                    DEMPTYPE = EMPTYPE;

                }




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
                dataGridView1.DataSource = GetDOOR();
                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelAP2(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false ;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string EMPID = "";
                string HNAME = "";
                string HDAY = "";
                string HD1 = "";
                string HD2 = "";
                string HS1 = "";
                string HS2 = "";
           


                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    HNAME = range.Text.ToString().Trim();

                    System.Data.DataTable F1 = GETETHID(HNAME);
                    if (F1.Rows.Count > 0)
                    {
                        EMPID = F1.Rows[0][0].ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        range.Select();
                        HDAY = range.Text.ToString().Trim();
                     
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        range.Select();
                        HD1 = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        range.Select();
                        HD2 = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        range.Select();
                        HS1 = range.Text.ToString().Trim();
   
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        range.Select();
                        HS2 = range.Text.ToString().Trim();

                       
                        System.Data.DataTable H1 = GETETH(EMPID);
                        if (H1.Rows.Count > 0)
                        {

                                string EMAIL = H1.Rows[0]["EMAIL"].ToString();
                     
                                string MANAGER = H1.Rows[0]["MANAGER"].ToString();

                                string EMAIL2 = H1.Rows[0]["EMAIL2"].ToString();
                                if (MANAGER == "許心如" || MANAGER == "黃舉昇")
                                {
                                    EMAIL2 = "";
                                }
                                if (MANAGER == "陳思怡")
                                {
                                    MANAGER = "陳彥琪";
                                    EMAIL2 = "applechen@acmepoint.com";
                                }
                                if (EMPID == "130002" || EMPID == "130010" || EMPID == "130017" || EMPID == "130005" || EMPID == "130020" || EMPID == "130023")
                                {
                                    EMAIL = "sylviashih@aresopto.com";
                                    EMAIL2 = "";
                                }
                                AddCARDDOOR2(EMPID,HNAME, HDAY, HD1, HD2, HS1, HS2, EMAIL, MANAGER, EMAIL2);
                          
                        }



                    }

                }


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
dataGridView1.DataSource  = GetDOOR2();
                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        public static System.Data.DataTable GetDOOR()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM HR_CARDDOOR ");

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

        public static System.Data.DataTable GetDOOR2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT EMPID,HNAME 中文姓名,HDAY 公司給假,HD1 合計已休時數,HD2 合計未休時數,HS1 給假起始日,HS2 可休截止日,EMAIL,MANAGER,EMAIL2 FROM HR_CARDDOOR2  ");

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
        public static System.Data.DataTable GetDOOR1()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DISTINCT EMPID FROM HR_CARDDOOR  ");

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
        public static System.Data.DataTable GetDOOR1F()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DISTINCT EMPID FROM HR_CARDDOOR2  ");

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
        public static System.Data.DataTable GetDOOR2(string EMPID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT TOP 1 FNAME,EMAIL,MANAGER,EMAIL2   FROM HR_CARDDOOR WHERE EMPID =@EMPID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
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
        public static System.Data.DataTable GetDOOR2F(string EMPID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT TOP 1 EMAIL,MANAGER,EMAIL2   FROM HR_CARDDOOR2 WHERE EMPID =@EMPID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
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
        public static System.Data.DataTable GetDOOR3(string EMPID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT GROUPNAME 部門名稱,EMPDATE 出勤日期,EMPCARD1 應刷卡時段,EMPCARD2 當日卡鐘資料,CARDSTATUS 異常狀態 FROM HR_CARDDOOR  WHERE EMPID =@EMPID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
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

        public static System.Data.DataTable GetDOOR4(string EMPID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT HNAME 中文姓名,HDAY 公司給假,HD1 合計已休時數,HD2 合計未休時數,HS1 給假起始日,HS2 可休截止日 FROM HR_CARDDOOR2  WHERE EMPID =@EMPID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPID", EMPID));
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


        public void sendGmail(string EMPID)
        {
            System.Data.DataTable M1 = GetDOOR2(EMPID);
            string FNAME = M1.Rows[0]["FNAME"].ToString();
            string EMAIL = M1.Rows[0]["EMAIL"].ToString();

            string EMAIL2 = M1.Rows[0]["EMAIL2"].ToString();
            Encoding encode = Encoding.GetEncoding("UTF-8");
            MailMessage mail = new MailMessage();
            //前面是發信email後面是顯示的名稱
          //  mail.From = new MailAddress("gmail", "xxx");
            //收信者email
            mail.To.Add(EMAIL);
            if (!String.IsNullOrEmpty(EMAIL2))
            {
                mail.CC.Add(EMAIL2);
            }
            //設定優先權
            mail.Priority = MailPriority.Normal;
            //標題
            mail.Subject = "超時出勤異常";
            mail.SubjectEncoding = encode;

            //內容
            mailBody(mail, EMPID, FNAME);

            //內容使用html
            mail.IsBodyHtml = true;
            //設定gmail的smtp
                   SmtpClient smtp = new SmtpClient();
            smtp.Send(mail);

            //放掉宣告出來的MySmtp
            smtp = null;
            //放掉宣告出來的mail
            mail.Dispose();
        }

        public void sendGmail2(string EMPID)
        {
            System.Data.DataTable M1 = GetDOOR2F(EMPID);
  
            string EMAIL = M1.Rows[0]["EMAIL"].ToString();

            string EMAIL2 = M1.Rows[0]["EMAIL2"].ToString();
            Encoding encode = Encoding.GetEncoding("UTF-8");
            MailMessage mail = new MailMessage();

            mail.To.Add(EMAIL);
            if (!String.IsNullOrEmpty(EMAIL2))
            {
                mail.CC.Add(EMAIL2);
            }
            //設定優先權
            mail.Priority = MailPriority.Normal;
            //標題
            mail.Subject = "特休時數提醒";
            mail.SubjectEncoding = encode;

            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\HR2.htm";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
            dataGridView2.DataSource = GetDOOR4(EMPID);
            string MailContent = htmlMessageBody(dataGridView2).ToString();

            template = template.Replace("##Content1##", "親愛的同仁，您好！");
            template = template.Replace("##Content2##", "提醒您，截至目前，您的特別休假天數及可申請使用之期限，如下方欄位所顯示。");
            template = template.Replace("##Content3##", MailContent);
            template = template.Replace("##Content4##", "*貼心提醒：為了您的身心健康著想，請定期安排休假，以紓解工作之辛勞，保持在工作、健康與家庭取得平衡～");
          
      
            //內容使用html
            mail.IsBodyHtml = true;
            mail.Body = template;
            SmtpClient smtp = new SmtpClient();
            smtp.Send(mail);

            //放掉宣告出來的MySmtp
            smtp = null;
            //放掉宣告出來的mail
            mail.Dispose();
        }


        public void mailBody(MailMessage mail, string EMPID, string FNAME)
        {


            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);



            dataGridView2.DataSource = GetDOOR3(EMPID);
            string MailContent = htmlMessageBody(dataGridView2).ToString();
            string palinBody = "【XXXX】";
            AlternateView plainView = AlternateView.CreateAlternateViewFromString(
                     palinBody, null, "text/plain");
            string htmlBody = string.Format("<html><body><P>Hi {0},</P><P>以下是您的超時出勤明細，請至portal填寫加班單或依照以下步驟填寫超時出勤原因(個人因素)，謝謝!</P>{1}<img alt=\"\" hspace=0 src=\"cid:HR\" align=baseline border=0 ></body></html>", FNAME, MailContent);
            //string htmlBody = "<p> 此為系統主動發送信函，請勿直接回覆此封信件。</p> ";
            //htmlBody += "<img alt=\"\" hspace=0 src=\"cid:HR\" align=baseline border=0 >";

            AlternateView htmlView =
                    AlternateView.CreateAlternateViewFromString(htmlBody, null, "text/html");
            imgResource(htmlView, "HR.JPG", "image/jpg");


            // add the views
            mail.AlternateViews.Add(plainView);
            mail.AlternateViews.Add(htmlView);
        }


        public void imgResource(AlternateView htmlView, string imgName, string imgType)
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string Chart = lsAppDir + "\\MailTemplates\\HR.JPG";
            LinkedResource imageResource = new LinkedResource(Chart, imgType);
            string[] imgArr = imgName.Split('.');
            imageResource.ContentId = imgArr[0];
            imageResource.TransferEncoding = TransferEncoding.Base64;
            htmlView.LinkedResources.Add(imageResource);

        }

        private string getImgPath(string strImgName)
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string Chart = lsAppDir + "\\MailTemplates\\HR.JPG";
            string strImgPath = @"C:\xxxxx\img\" + strImgName;
            return strImgPath;
        }

        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  今日無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }

            //create html & table
            //strB.AppendLine("<html><body><center><table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border-collapse:collapse;'>");
            strB.AppendLine("<tr class='HeaderBorder'>");
            //cteate table header
            for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
            {
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                {

    

                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;

                    if (KeyValue.IndexOf("{") >= 0)
                    {
                        tmpKeyValue = "";
                    }
                }
                else
                {
                    tmpKeyValue = "";
                }


                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }



                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                }


                for (int d = 1; d <= dg.Rows[i].Cells.Count - 1; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    // HttpUtility.HtmlDecode("&nbsp;&nbsp;&nbsp;")

                    if (dgvc.ValueType == typeof(Int32))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Int32 x = Convert.ToInt32(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }

                    else if (dgvc.ValueType == typeof(Decimal) || dgvc.ValueType == typeof(Double))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Decimal x = Convert.ToDecimal(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }
                    else
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (comboBox1.Text == "超時出勤異常")
            {
                System.Data.DataTable M1 = GetDOOR1();
                if (M1.Rows.Count > 0)
                {
                    for (int i = 0; i <= M1.Rows.Count - 1; i++)
                    {
                        string EMPID = M1.Rows[i][0].ToString();
                        sendGmail(EMPID);
                    }

                    MessageBox.Show("寄信成功");
                }
            }

            if (comboBox1.Text == "特休時數提醒")
            {

                System.Data.DataTable M1 = GetDOOR1F();
                if (M1.Rows.Count > 0)
                {
                    for (int i = 0; i <= M1.Rows.Count - 1; i++)
                    {
                        string EMPID = M1.Rows[i][0].ToString();
                        sendGmail2(EMPID);
                    }

                    MessageBox.Show("寄信成功");
                }
            }
        }

        private void HRNO_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "超時出勤異常";
        }

    }
}

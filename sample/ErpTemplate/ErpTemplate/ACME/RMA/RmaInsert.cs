
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Net;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
using System.Diagnostics;
using System.Threading;
using SAPbobsCOM;
namespace ACME
{
    public partial class RmaInsert : Form
    {
        string FA = "acmesql02";
        int G1;
        int G2;
        int s1 = 0;
        int inint = 0;
        string NewFileName;
        private string FileName;
        private System.Data.DataTable TempDt;
        private System.Data.DataTable TempDtS;
        string RMA;
        int ROW;
        System.Net.Mail.Attachment data = null;
        string TRMA = "";
        string TMODEL = "";
        string TVER = "";
        string TQTY = "";
        System.Data.DataTable dtCost2 = null;
        public RmaInsert()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
          
                   System.Data.DataTable dt = Getinvoiced(textBox2.Text);
                    DataRow drw = dt.Rows[0];
                    string model = drw["model"].ToString();
                    string rmano = drw["rmano"].ToString();
                    string startdate = drw["startdate"].ToString();
                    string enddate = drw["enddate"].ToString();
                    if (globals.GroupID.ToString().Trim() != "EEP")
                    {
                        AddAUOGD2(Convert.ToInt32(textBox2.Text));
                    }
                    GetExcelContentGD4(textBox1.Text, rmano, model,startdate, enddate);
                    MessageBox.Show("匯入成功");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetExcelContentGD4(string ExcelFile, string u_RMA_1, string u_model_2, string startdate, string enddate)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString())+1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }

            }

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            int Line = 1;
            int Itemgroup = 104 ;

            string u_seq_no1;
            string u_ver;
            string u_month_seq = "";
            string u_c_complain = "";
            string u_iqc = "";
            string u_acme_confirm = "";
            string u_acme_judge = "";
            string u_place_1 = "";
            string u_remark1 = "";
            string NO;
            string u_model_1="";
            int H = 0;
            string VER="";
            string ENG = "";
            int k = 0;
            for (int i = 3; i <= iRowCnt; i++)
            {
             
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                //range.Select();
     

                NO = range.Text.ToString().Trim() ;
                int g = NO.ToUpper().LastIndexOf("MODEL");
                if (g.ToString() != "-1")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    u_model_1 = range.Text.ToString();
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
             
                u_seq_no1 = range.Text.ToString();
             
   
                if (comboBox4.Text =="KIT")
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    u_ver = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    u_month_seq = range.Text.ToString();

                    u_iqc = "";


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    u_c_complain = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    u_acme_confirm = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    u_acme_judge = range.Text.ToString();


                    u_place_1 = "";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    u_remark1 = range.Text.ToString();
                }
                else if (comboBox4.Text == "Open frame")
                {

                    u_ver = "";

                 
                    u_month_seq = "";

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    u_iqc = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    u_c_complain = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    u_acme_confirm = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    u_acme_judge = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    u_place_1 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    u_remark1 = range.Text.ToString();
                }
                else
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    u_ver = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    u_month_seq = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    u_iqc = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    u_c_complain = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    u_acme_confirm = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    u_acme_judge = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    u_place_1 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                    u_remark1 = range.Text.ToString();
                }

                try
                {

                    int P4 = NO.ToUpper().IndexOf("製表者");
                    if (P4 != -1)
                    {
                        k = i;

                        ENG = NO.Replace("製表者", "").Replace("：", "").Replace(":", "").Trim();
                    }

                    int s = Convert.ToInt32(textBox2.Text);
                    if (checkBox7.Checked)
                    {
                        if (NO != "No." && NO != "編號" && IsNumber(NO))
                        {
                            if (k == 0)
                            {
                                if (globals.GroupID.ToString().Trim() != "EEP")
                                {
                                    AddAUOGD4(s, Line, u_ver, u_month_seq, u_iqc, u_c_complain, u_seq_no1, u_acme_confirm, u_acme_judge, u_place_1, u_model_1, 1, Itemgroup, u_RMA_1, u_model_1, u_remark1, startdate, enddate);
                                }
                                Line = Line + 1;
                                Itemgroup = Itemgroup + 1;

                                if (u_acme_judge.Trim() == "NG")
                                {
                                    H += 1;
                                }
                                VER = u_ver;
                            }
                        }
                    }
                    else
                    {
                        if (u_seq_no1 != "" & NO != "No." && NO != "編號" && IsNumber(NO))
                        {
                            if (k == 0)
                            {
                                if (globals.GroupID.ToString().Trim() != "EEP")
                                {
                                    AddAUOGD4(s, Line, u_ver, u_month_seq, u_iqc, u_c_complain, u_seq_no1, u_acme_confirm, u_acme_judge, u_place_1, u_model_1, 1, Itemgroup, u_RMA_1, u_model_1, u_remark1, startdate, enddate);
                                }
                                Line = Line + 1;
                                Itemgroup = Itemgroup + 1;

                                if (u_acme_judge.Trim() == "NG")
                                {
                                    H += 1;
                                }
                                VER = u_ver;
                            }
                        }
                    }
                    if (i == iRowCnt)
                    {
                        if (!String.IsNullOrEmpty(ENG))
                        {
                            UPDATE(s, u_model_1, VER, H, ENG);

                        }
                        System.Data.DataTable J1 = GetF1(s.ToString());
                        if (J1.Rows.Count > 0)
                        {
                            string R1 = J1.Rows[0][0].ToString();
                            UPDATE2(s, R1);
                        }
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

        private void GetExcel4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;
          

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            string CONTRACT = "";
            string EXCEL = "";
            for (int h = 0; h <= dataGridView1.Rows.Count - 2; h++)
            {
                System.Data.DataTable dt = TempDt;
                DataRow drw = dt.Rows[h];

                CONTRACT = dataGridView1.Rows[h].Cells["契約號碼"].Value.ToString();
                EXCEL = dataGridView1.Rows[h].Cells["EXCEL分頁"].Value.ToString();
                if (!String.IsNullOrEmpty(CONTRACT))
                {
                    AddAUOGD2(Convert.ToInt32(CONTRACT));
                    System.Data.DataTable dtt = Getinvoiced(CONTRACT);
                    DataRow drw1 = dtt.Rows[0];
                    string model = drw1["model"].ToString();
                    string rmano = drw1["rmano"].ToString();
                    string startdate = drw1["startdate"].ToString();
                    string enddate = drw1["enddate"].ToString();



                    int sd1 = h + 1;
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

                    int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                    int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                    if (!checkBox6.Checked)
                    {
                        if (iRowCnt > 200)
                        {
                            iRowCnt = 200;
                        }
                    }



                    Microsoft.Office.Interop.Excel.Range range = null;



                    object SelectCell = "A1";
                    range = excelSheet.get_Range(SelectCell, SelectCell);

                    int Line = 1;
                    int Itemgroup = 104;

                    string u_seq_no1;
                    string u_ver;
                    string u_month_seq;
                    string u_c_complain;
                    string u_iqc;
                    string u_acme_confirm;
                    string u_acme_judge;
                    string u_place_1;
                    string u_remark1;
                    string NO;
                    string u_model_1 = "";
                    int H = 0;
                    int k=0;
                    string ENG = "";
                    string VER = "";
                    for (int i = 3; i <= iRowCnt; i++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);

                        NO = range.Text.ToString().Trim();
                        int g = NO.ToUpper().LastIndexOf("MODEL");
                        if (g.ToString() != "-1")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                            u_model_1 = range.Text.ToString();
                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);

                        u_seq_no1 = range.Text.ToString();

                    
                        if (comboBox5.Text == "KIT")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                            u_ver = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                            u_month_seq = range.Text.ToString();

                            u_iqc = "";
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                            u_c_complain = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                            u_acme_confirm = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                            u_acme_judge = range.Text.ToString();


                            u_place_1 = "";

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                            u_remark1 = range.Text.ToString();
                        }
                        else if (comboBox5.Text == "Open frame")
                        {


                            u_ver = "";


                            u_month_seq = "";

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                            u_iqc = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                            u_c_complain = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                            u_acme_confirm = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                            u_acme_judge = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                            u_place_1 = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                            u_remark1 = range.Text.ToString();
                        }
                        else
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                            u_ver = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                            u_month_seq = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                            u_iqc = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                            u_c_complain = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                            u_acme_confirm = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                            u_acme_judge = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                            u_place_1 = range.Text.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                            u_remark1 = range.Text.ToString();
                        }

                        try
                        {
                            int P4 = NO.ToUpper().IndexOf("製表者");
                            if (P4 != -1)
                            {
                                k = i;

                                ENG = NO.Replace("製表者", "").Replace("：", "").Replace(":", "").Trim();
                            }

                            int s = Convert.ToInt32(CONTRACT);
                            if (u_seq_no1 != "" & NO != "No." && NO != "編號" && IsNumber(NO))
                            {
                                if (k == 0)
                                {

                                    Line = Line + 1;
                                    Itemgroup = Itemgroup + 1;

                                    if (u_acme_judge.Trim() == "NG")
                                    {
                                        H += 1;
                                    }
                                    VER = u_ver;

                                    AddAUOGD4(s, Line, u_ver, u_month_seq, u_iqc, u_c_complain, u_seq_no1, u_acme_confirm, u_acme_judge, u_place_1, u_model_1, 1, Itemgroup, rmano, u_model_1, u_remark1, startdate, enddate);
                                }
                            }


                            if (i == iRowCnt)
                            {
                                if (!String.IsNullOrEmpty(ENG))
                                {
                                  
                                    UPDATE(s, u_model_1, VER, H,ENG);
                                }
                            }
            
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                    range = null;
                    excelSheet = null;
                }
            }
            //Quit
            excelApp.Quit();


            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


         
            excelApp = null;
            excelBook = null;


            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();



        }

        private System.Data.DataTable GetF1(string CONTRACTID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

          
            sb.Append(" SELECT  T2.slpname FROM OCTR T0");
            sb.Append(" LEFT JOIN ocrd T1 ON (T0.CstmrCode =T1.CardCode)");
            sb.Append(" LEFT JOIN OSLP T2 ON (T1.SlpCode =T2.SlpCode)");
            sb.Append(" WHERE CONTRACTID=@CONTRACTID AND ISNULL(T2.slpname,'')<>''");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));

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
        public void AddAUOGD2(int ContrX)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Delete from ctr1 where ContractID = @ContrX", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ContrX", ContrX));
          
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
        public void AddAUOGD4(int ContractID, int Line, string U_U_Ver, string U_U_month_seq, string U_U_iqc, string U_U_C_complain, string U_S_seq, string U_U_acme_confirm, string U_U_Acme_judge, string U_U_PLACE_1, string ItemCode, int insid, int Itemgroup, string U_RMA_no, string U_Model, string U_rRemark, string startdate, string enddate)
        {

            SqlConnection connection = globals.shipConnection; 
            SqlCommand command = new SqlCommand("Insert into CTR1(ContractID,Line,U_U_Ver,U_U_month_seq,U_U_iqc,U_U_C_complain,U_S_seq,U_U_acme_confirm,U_U_Acme_judge,U_U_PLACE_1,ItemCode,insid,Itemgroup,U_RMA_no,U_Model,U_rRemark,startdate,enddate) values(@ContractID,@Line,@U_U_Ver,@U_U_month_seq,@U_U_iqc,@U_U_C_complain,@U_S_seq,@U_U_acme_confirm,@U_U_Acme_judge,@U_U_PLACE_1,@ItemCode,@insid,@Itemgroup,@U_RMA_no,@U_Model,@U_rRemark,@startdate,@enddate)", connection);


            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ContractID", ContractID));
            command.Parameters.Add(new SqlParameter("@Line", Line));

            command.Parameters.Add(new SqlParameter("@insid", insid));
            command.Parameters.Add(new SqlParameter("@Itemgroup", Itemgroup));
            command.Parameters.Add(new SqlParameter("@startdate", startdate));
            command.Parameters.Add(new SqlParameter("@enddate", enddate));
         

            command.Parameters.Add(new SqlParameter("@U_U_Ver", U_U_Ver));
            if (String.IsNullOrEmpty(U_U_Ver))
            {
                command.Parameters["@U_U_Ver"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_month_seq", U_U_month_seq));
            if (String.IsNullOrEmpty(U_U_month_seq))
            {
                command.Parameters["@U_U_month_seq"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_iqc", U_U_iqc));
            if (String.IsNullOrEmpty(U_U_iqc))
            {
                command.Parameters["@U_U_iqc"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_C_complain", U_U_C_complain));
            if (String.IsNullOrEmpty(U_U_C_complain))
            {
                command.Parameters["@U_U_C_complain"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_S_seq", U_S_seq));
            if (String.IsNullOrEmpty(U_S_seq))
            {
                command.Parameters["@U_S_seq"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_acme_confirm", U_U_acme_confirm));
            if (String.IsNullOrEmpty(U_U_acme_confirm))
            {
                command.Parameters["@U_U_acme_confirm"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_Acme_judge", U_U_Acme_judge));
            if (String.IsNullOrEmpty(U_U_Acme_judge))
            {
                command.Parameters["@U_U_Acme_judge"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_U_PLACE_1", U_U_PLACE_1));
            if (String.IsNullOrEmpty(U_U_PLACE_1))
            {
                command.Parameters["@U_U_PLACE_1"].Value = "";
            }

            if (ItemCode.Length >= 20)
            {
               ItemCode= ItemCode.Substring(0, 20);
            }
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            if (String.IsNullOrEmpty(ItemCode))
            {
                command.Parameters["@ItemCode"].Value = "";
            }



            command.Parameters.Add(new SqlParameter("@U_RMA_no", U_RMA_no));
            if (String.IsNullOrEmpty(U_RMA_no))
            {
                command.Parameters["@U_RMA_no"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_Model", U_Model));

            if (String.IsNullOrEmpty(U_Model))
            {
                command.Parameters["@U_Model"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_rRemark", U_rRemark));

            if (String.IsNullOrEmpty(U_rRemark))
            {
                command.Parameters["@U_rRemark"].Value = "";
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


        public System.Data.DataTable Getinvoiced(string contractid)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "select Convert(varchar(10),startdate,111) startdate,Convert(varchar(10),enddate,111 ) enddate,u_rma_no rmano,u_rmodel model from octr where contractid=@contractid ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@contractid", contractid));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public void UPDATE(int ContractID, string U_RMODEL, string U_RVER, int U_RQUINITY, string U_RENGINEER)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE OCTR SET UPDATE OCTR SET U_RMODEL=@U_RMODEL,U_RVER=@U_RVER,U_RQUINITY=@U_RQUINITY,U_RENGINEER=@U_RENGINEER,U_RSales=@U_RSales WHERE ContractID=@ContractID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ContractID", ContractID));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_RQUINITY", U_RQUINITY));
            command.Parameters.Add(new SqlParameter("@U_RENGINEER", U_RENGINEER));

            //U_RSales
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

        public void UPDATE2(int ContractID,string U_RSales)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE OCTR SET U_RSales=@U_RSales WHERE ContractID=@ContractID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ContractID", ContractID));

            command.Parameters.Add(new SqlParameter("@U_RSales", U_RSales));
            //U_RSales
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
        private void button1_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                comboBox1.SelectedIndex = -1;
                comboBox1.Items.Clear();
                FileName = openFileDialog1.FileName;
                this.textBox1.Text = FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
              excelApp.Visible = false;
                object oMissing = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                string Count_Sheet = excelBook.Sheets.Count.ToString();
                int i = excelBook.Sheets.Count;

                for (int xi = 1; xi <= i; xi++)
                {

                    Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                    string X1 = xi.ToString();
                    string X2 = excelsheet.Name.ToString();
                    string name_sheet = X1 + ":" + X2;
                    comboBox1.Items.Add(name_sheet);

                    Microsoft.Office.Interop.Excel.Range range = null;
                    int iRowCnt = excelsheet.UsedRange.Cells.Rows.Count;
                    string NO;
                    string MODEL = "";
                    string VER = "";
                    string JUD = "";
                    string SN = "";
                    for (int i2 = 3; i2 <= iRowCnt; i2++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 1]);

                        NO = range.Text.ToString();
                        int g = NO.ToUpper().LastIndexOf("MODEL");
                        if (g.ToString() != "-1")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 3]);
                            MODEL = range.Text.ToString();
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 2]);
                        SN = range.Text.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 3]);
                        VER = range.Text.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 8]);
                        JUD = range.Text.ToString().Trim();



                        if (!String.IsNullOrEmpty(SN) && !String.IsNullOrEmpty(JUD) && !String.IsNullOrEmpty(VER))
                        {
                            System.Data.DataTable G1 = GetDRS3(SN, X2);

                            if (G1.Rows.Count > 0)
                            {
                                for (int f = 0; f <= G1.Rows.Count - 1; f++)
                                {
                                    string RMA = G1.Rows[f][0].ToString();
                                    string MESSAGE = X2 + "_" + SN + "與RMANO#" + RMA + "重複";

                                    MessageBox.Show(MESSAGE);
                                }

                            }
                        }
                    }

                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                excelApp = null;
                excelBook = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                GetExcelContentGD5(textBox1.Text);
                MessageBox.Show("已匯入");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetExcelContentGD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            int Line = 1;
            int Itemgroup = 104;

            string x1;
            string x5;
            string x7;
            string x8;
            string x9;
         


            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                x1 = range.Text.ToString().Trim();

                int G1 = x1.Length;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                x5 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                x7 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                x9 = range.Text.ToString().Trim();

                try
                {
                    if (G1 >= 10)
                    {
                        AddRma(x1, x9, x7, x5);
                        AddRma2(x1, x5);
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
        private void GetVender(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


    
            string x1;
            string x3;
            string x4;
            string x5;

            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                x1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                x5 = range.Text.ToString().Trim();
                string a1 = x5.Substring(0, 8);
                int d = x5.Length;
                string a2 = x5.Substring(9, d - 9);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                x4 = range.Text.ToString().Trim();

                DateTime AA = Convert.ToDateTime(x4);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                x5 = range.Text.ToString().Trim();

                try
                {
                    UpVender(x1, a1, a2, AA, x5);

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
        private void GetVenderNOT(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }


            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);



            string x1 = "";
            string x2 = "";
            string x3 = "";
            string x4 = "";
            string x5 = "";
            string x6 = "";
            string x7 = "";
            string x8 = "";
            string x9 = "";
            string x10 = "";
            string MODEL = "";
            string VER = "";
            string QTY = "";
            TempDtS = MakeTableS();
            DataRow drS = null;
            for (int i = 2; i <= iRowCnt; i++)
            {
             

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                x1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                x2 = range.Text.ToString().Trim().ToUpper();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                x4 = range.Text.ToString().Trim().ToUpper();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                x5 = range.Text.ToString().Trim().ToUpper();

     
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                x3 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                x6 = range.Text.ToString().Trim();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                x7 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                x8 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                x9 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                x10 = range.Text.ToString().Trim();
                System.Data.DataTable G1 = GetTEMP(x1, x8);
                if (G1.Rows.Count > 0)
                {
                    MODEL = G1.Rows[0]["MODEL"].ToString().ToUpper();
                    VER = G1.Rows[0]["VER"].ToString().ToUpper();
                    QTY = G1.Rows[0]["QTY"].ToString().ToUpper();
                }
                int K1 = x4.Length + x5.Length;
                if (K1 > 0)
                {
                    if (MODEL != x4 || VER != x5 || QTY != x6 || QTY != x7)
                    {
                        drS = TempDtS.NewRow();
                        drS["RMANO"] = x1;
                        drS["ASP"] = x2;
                        drS["Vender收貨日"] = x3;
                        if (MODEL != x4)
                        {
                            drS["MODEL"] = x4 + "--異常";
                        }
                        else
                        {
                            drS["MODEL"] =  x4;
                        }
                        if (VER != x5)
                        {
                            drS["VER"] = x5 + "--異常";
                        }
                        else
                        {
                            drS["VER"] = x5;
                        }
                        if (QTY != x6)
                        {
                            drS["RequestQty"] = x6 + "--異常";
                        }
                        else
                        {
                            drS["RequestQty"] = x6;
                        }
                        if (QTY != x7)
                        {
                            drS["Vender未還數量"] = x7 + "--異常";
                        }
                        else
                        {
                            drS["Vender未還數量"] = x7;
                        }
                        drS["REF"] = x8;
                        TempDtS.Rows.Add(drS);

                    }
                }
           
                try
                {
                    System.Data.DataTable G2 = GetTEMP2(x1, x4, x5, x6);
                    if (G2.Rows.Count > 0)
                    {
                        MessageBox.Show("RMA NO " + x1 + " 已有日期");
                    }

                    System.Data.DataTable G3 = GetTEMP3(x1, x4, x5, x6);
                    if (G3.Rows.Count > 0)
                    {
                        MessageBox.Show("RMA NO " + x1 + " 匯入完成");
                    }

                    if (x1 != "")
                    {
                        DateTime AA = Convert.ToDateTime(x3);
                        if (globals.GroupID.ToString().Trim() != "EEP")
                        {
                            UpVenderNOT(x1, x4, x5, AA, x6);
                        }

                        if (x9 != "")
                        {
                            DateTime AB = Convert.ToDateTime(x9);
                            if (globals.GroupID.ToString().Trim() != "EEP")
                            {
                                UpVenderNOTDATE1(x1, x4, x5, AB);
                            }
                        }

                        if (x10 != "")
                        {
                            DateTime AC = Convert.ToDateTime(x10);
                            if (globals.GroupID.ToString().Trim() != "EEP")
                            {
                                UpVenderNOTDATE2(x1, x4, x5, AC);
                            }
                        }
          
                    }



                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            dataGridView3.DataSource = TempDtS;


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


        private void GetVenderOUT(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;
            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }


            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);



            string x1;
            string x2;
         
            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                x1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                x2 = range.Text.ToString().Trim();



                try
                {
                    if (x1 != "" && x2 !="")
                    {
                        DateTime AA = Convert.ToDateTime(x2);
                        UpVenderOUT(x1, AA);
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
        public void AddRma(string U_AUO_RMA_NO, string U_RepairCenter, string U_yetqty, string U_RMA_NO)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR  set  U_AUO_RMA_NO = @U_AUO_RMA_NO,U_RepairCenter = @U_RepairCenter ,U_yetqty =@U_yetqty  where U_RMA_NO =@U_RMA_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            if (String.IsNullOrEmpty(U_AUO_RMA_NO))
            {
                command.Parameters["@U_AUO_RMA_NO"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_RepairCenter", U_RepairCenter));
            if (String.IsNullOrEmpty(U_RepairCenter))
            {
                command.Parameters["@U_RepairCenter"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@U_yetqty", U_yetqty));
            if (String.IsNullOrEmpty(U_yetqty))
            {
                command.Parameters["@U_yetqty"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
        
            
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

        public void AddRma2(string U_AUO_RMA, string U_RMA_No)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update CTR1 set  U_AUO_RMA = @U_AUO_RMA where U_RMA_No = @U_RMA_No ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA", U_AUO_RMA));
            if (String.IsNullOrEmpty(U_AUO_RMA))
            {
                command.Parameters["@U_AUO_RMA"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@U_RMA_No", U_RMA_No));
      

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
        public void UpVender(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, DateTime U_ACME_BACKDATE, string U_ACME_QBACK)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR set  U_ACME_BACKDATE =@U_ACME_BACKDATE,U_ACME_QBACK=@U_ACME_QBACK where U_AUO_RMA_NO = @U_AUO_RMA_NO AND U_RMODEL = @U_RMODEL AND U_RVER = @U_RVER ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_ACME_BACKDATE", U_ACME_BACKDATE));
            command.Parameters.Add(new SqlParameter("@U_ACME_QBACK", U_ACME_QBACK));
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
        public void UpVenderNOT(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, DateTime U_ACME_RECEDATE, string U_YETQTY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR set  U_ACME_RECEDATE =@U_ACME_RECEDATE where U_AUO_RMA_NO = @U_AUO_RMA_NO AND SUBSTRING(U_RMODEL,1,9) = @U_RMODEL AND U_RVER = @U_RVER AND U_YETQTY=@U_YETQTY AND ISNULL(U_ACME_RECEDATE,'') = '' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_ACME_RECEDATE", U_ACME_RECEDATE));
            command.Parameters.Add(new SqlParameter("@U_YETQTY", U_YETQTY));
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
        public void UpVenderNOTDATE1(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, DateTime U_RTORECEIVING)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR set  U_RTORECEIVING =@U_RTORECEIVING where U_AUO_RMA_NO = @U_AUO_RMA_NO AND SUBSTRING(U_RMODEL,1,9) = @U_RMODEL AND U_RVER = @U_RVER AND U_YETQTY=@U_YETQTY AND ISNULL(U_ACME_RECEDATE,'') = '' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));

            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));

            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_RTORECEIVING", U_RTORECEIVING));

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
        public void UpVenderNOTDATE2(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, DateTime U_ACME_OUT)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR set  U_ACME_OUT =@U_ACME_OUT where U_AUO_RMA_NO = @U_AUO_RMA_NO AND SUBSTRING(U_RMODEL,1,9) = @U_RMODEL AND U_RVER = @U_RVER AND U_YETQTY=@U_YETQTY AND ISNULL(U_ACME_RECEDATE,'') = '' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_ACME_OUT", U_ACME_OUT));

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
        public void UpVenderOUT(string U_AUO_RMA_NO, DateTime U_ACME_OUT)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" Update OCTR set  U_ACME_OUT =@U_ACME_OUT  where U_AUO_RMA_NO = @U_AUO_RMA_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));

            command.Parameters.Add(new SqlParameter("@U_ACME_OUT", U_ACME_OUT));

  
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
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("請匯入EXCEL");
                    
                    return;
                }

                if (comboBox1.Text == "")
                {
                    MessageBox.Show("請選擇分頁");

                    return;
                }
                GetVender(textBox1.Text);
                MessageBox.Show("匯入完成");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
        
                if (textBox1.Text == "")
                {
                    MessageBox.Show("請匯入EXCEL");

                    return;
                }
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("請選擇分頁");

                    return;
                }

                GetVenderNOT(textBox1.Text);
           
        }
        private System.Data.DataTable MakeTableS()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("RMANO", typeof(string));
            dt.Columns.Add("ASP", typeof(string));
            dt.Columns.Add("Vender收貨日", typeof(string));
            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("RequestQty", typeof(string));
            dt.Columns.Add("Vender未還數量", typeof(string));
            dt.Columns.Add("REF", typeof(string));


            return dt;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            S1();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //comboBox1.SelectedIndex = -1;
                //comboBox1.Items.Clear();
                FileName = openFileDialog1.FileName;
                this.textBox3.Text = FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;

                object oMissing = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                string Count_Sheet = excelBook.Sheets.Count.ToString();
                int i = excelBook.Sheets.Count;
                
                TempDt = MakeTable();
                DataRow dr = null;
                for (int xi = 1; xi <= i; xi++)
                {
                    dr = TempDt.NewRow();
                    Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                    string X1 = xi.ToString();
                    string X2 = excelsheet.Name.ToString();
                    dr["NO"] = X1;
                    dr["EXCEL分頁"] = X2;
                    TempDt.Rows.Add(dr);
                    Microsoft.Office.Interop.Excel.Range range = null;
                    int iRowCnt = excelsheet.UsedRange.Cells.Rows.Count;
                    string NO;
                    string MODEL = "";
                    string VER = "";
                    string JUD = "";
                    string SN = "";
                    for (int i2 = 3; i2 <= iRowCnt; i2++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 1]);

                        NO = range.Text.ToString();
                        int g = NO.ToUpper().LastIndexOf("MODEL");
                        if (g.ToString() != "-1")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 3]);
                            MODEL = range.Text.ToString();
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 2]);
                        SN = range.Text.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 3]);
                        VER = range.Text.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 8]);
                        JUD = range.Text.ToString().Trim();

                

                        if (!String.IsNullOrEmpty(SN) && !String.IsNullOrEmpty(JUD) && !String.IsNullOrEmpty(VER))
                        {
                            System.Data.DataTable G1 = GetDRS3(SN, X2);

                            if (G1.Rows.Count > 0)
                            {
                                for (int f = 0; f <= G1.Rows.Count - 1; f++)
                                {
                                    string RMA = G1.Rows[f][0].ToString();
                                    string MESSAGE = X2 + "_" + SN + "與RMANO#" + RMA + "重複";

                                    MessageBox.Show(MESSAGE);
                                }

                            }
                        }
                    }
                }

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                excelApp = null;
                excelBook = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


                dataGridView1.DataSource = TempDt;
            }
        }
        private System.Data.DataTable GetDRS3(string SN, string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT U_RMA_NO  FROM  CTR1   WHERE U_S_seq =@SN AND U_RMA_NO <> @U_RMA_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SN", SN));
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
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
        private void SG(string CONTRACTID)
        {
            TRMA = "";
            TMODEL = "";
            TVER = "";
            TQTY = "";
            System.Data.DataTable S1 = GetTEMPT(CONTRACTID);
            if (S1.Rows.Count > 0)
            {
                TRMA = S1.Rows[0]["RMA"].ToString().ToUpper();
                TMODEL = S1.Rows[0]["MODEL"].ToString().ToUpper();
                TVER = S1.Rows[0]["VER"].ToString().ToUpper();
                TQTY = S1.Rows[0]["QTY"].ToString().ToUpper();
            }
        }
        private void  S1()
        {
            try
            {

                RMA = "";
                if (textBox3.Text == "")
                {
                    MessageBox.Show("請選擇檔案");
                    return;
                }

                string EXCEL = "";

                System.Data.DataTable K12 = GetMenu.GETDEFECTCODE2();
                if (K12.Rows.Count == 0)
                {
                    MessageBox.Show("DEFECTCODE沒有資料，請重新匯入");
                    return;
                }

                G1 = 0;
                try
                {
                    DELETEFILE();
                    DELETEFILE2();
                }
                catch
                { }


                if (comboBox5.Text == "KIT")
                {

                    for (int h = 0; h <= dataGridView1.Rows.Count - 2; h++)
                    {
                        G2 = 0;
                        System.Data.DataTable dt = TempDt;
                        DataRow drw = dt.Rows[h];
                        string CONTRACT = dataGridView1.Rows[h].Cells["契約號碼"].Value.ToString();
                         EXCEL = dataGridView1.Rows[h].Cells["EXCEL分頁"].Value.ToString();
                        if (!String.IsNullOrEmpty(CONTRACT))
                        {
                            GetMULTI2(textBox3.Text, h + 1);

                            if (G2 == 1)
                            {
                                RMA += textBox5.Text + " " + EXCEL + ".";
                                GetEMULTI3(h + 1, EXCEL);
                            }

                            DELETEFILE();
                        }
                    }

                }
                else if (comboBox5.Text == "Open frame")
                {
                    for (int h = 0; h <= dataGridView1.Rows.Count - 2; h++)
                    {
                        G2 = 0;
                        System.Data.DataTable dt = TempDt;
                        DataRow drw = dt.Rows[h];

                        string CONTRACT = dataGridView1.Rows[h].Cells["契約號碼"].Value.ToString();
                         EXCEL = dataGridView1.Rows[h].Cells["EXCEL分頁"].Value.ToString();
                        if (!String.IsNullOrEmpty(CONTRACT))
                        {

                            GetMULTI2(textBox3.Text, h + 1);

                            if (G2 == 1)
                            {

                                RMA += textBox5.Text + " " + EXCEL + ".";
                                GetEMULTI3(h + 1, EXCEL);
                            }
                            DELETEFILE();
                        }
                    }

                }
                else
                {
                    //for (int h = 0; h <= dataGridView1.Rows.Count - 2; h++)
                    //{
                    //    G2 = 0;
                    //    System.Data.DataTable dt = TempDt;
                    //    DataRow drw = dt.Rows[h];

                    //    string CONTRACT = dataGridView1.Rows[h].Cells["契約號碼"].Value.ToString();
                    //    if (CONTRACT == "")
                    //    {
                    //        MessageBox.Show("請輸入契約號碼");

                    //        return;

                    //    }

                    //     EXCEL = dataGridView1.Rows[h].Cells["EXCEL分頁"].Value.ToString();
                    //    if (!String.IsNullOrEmpty(CONTRACT))
                    //    {

                    //        GetMULTI(textBox3.Text, h + 1,CONTRACT);

                    //        if (G2 == 1)
                    //        {

                    //            RMA += textBox5.Text + " " + EXCEL + ".";
                    //            GetEMULTI3(h + 1, EXCEL);
                    //        }
                    //        DELETEFILE();
                    //    }
                    //}

                }

                if (G1 == 1)
                {

                    string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";



                    string J1 = RMA.Remove(RMA.Length - 1, 1);

                    MAILFILE2(J1);


                    DialogResult result;
                    result = MessageBox.Show(EXCEL + " 匯入格式有錯，是否要繼續匯入？", "YES/NO", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        GetExcel4(textBox3.Text);
                        MessageBox.Show(EXCEL + " 匯入成功");
                    }
                }
                else
                {
                    GetExcel4(textBox3.Text);
                    MessageBox.Show(EXCEL + " 匯入成功");
                }



                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NO", typeof(string));
            dt.Columns.Add("EXCEL分頁", typeof(string));
            dt.Columns.Add("契約號碼", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTableF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("RMANO", typeof(string));
            dt.Columns.Add("客戶", typeof(string));

            return dt;
        }
        private void RmaInsert_Load(object sender, EventArgs e)
        {



            System.Data.DataTable G1 = GetMenu.Getdata("RMAIND");
            if (G1.Rows.Count > 0)
            {
                textBox6.Text = G1.Rows[0][0].ToString();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("請選擇檔案");
                    return;
                }

                 if (textBox2.Text  == "")
                {
                    MessageBox.Show("請填寫契約號碼");
                    return;
                }
                 if (comboBox1.SelectedIndex  == -1)
                {
                    MessageBox.Show("請選擇SHEET");
                    return;
                }

                System.Data.DataTable K12 = GetMenu.GETDEFECTCODE2();
                if (K12.Rows.Count == 0)
                {
                    MessageBox.Show("DEFECTCODE沒有資料，請重新匯入");
                    return;
                }
            
            G1 = 0;
            try
            {
                DELETEFILE();
                DELETEFILE2();
            }
            catch
            { }

                try
                {
                    if (globals.GroupID.ToString().Trim() == "EEP")
                    {
                        button2_Click(sender, e);
                    }
                    else if (checkBox1.Checked)
                    {
                        button2_Click(sender, e);
                    }
                    else
                    {
                        if (comboBox4.Text == "KIT")
                        {
                            GetExcelProduct3(textBox1.Text);
                        }

                        else if (comboBox4.Text == "Open frame")
                        {
                            GetExcelProduct4(textBox1.Text);
                        }
                        else
                        {
                            GetExcelProduct(textBox1.Text);
                        }
                        string t1 = comboBox1.Text.ToString();
                        int j1 = t1.IndexOf(":");
                        int j2 = t1.Length;
                        string t2 = t1.Substring(j1 + 1, j2 - 1 - j1);
                        GetExcelProduct2(t2);




                        if (G1 == 1)
                        {



                            MAILFILE(textBox4.Text);


                            DialogResult result;
                            result = MessageBox.Show("匯入格式有錯，是否要繼續匯入？", "YES/NO", MessageBoxButtons.YesNo);
                            if (result == DialogResult.Yes)
                            {
                                button2_Click(sender, e);
                            }
                        }
                        else
                        {
                            button2_Click(sender, e);
                        }
                    }
       
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }

    
        }

        private void MAILFILE(string RMAA)
        {
            string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL3.xls";
           

         
                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\RMA.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);


                template = template.Replace("##Content##", "Dear,");
                template = template.Replace("##MARK##", textBox5.Text + " " + RMAA);
                MailMessage message = new MailMessage();


                message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
                //if (fmLogin.LoginID.ToString().ToUpper() == "LLEYTONCHEN")
                //{
                //    NAME = "LLEYTONCHEN";
                //}
              //  message.To.Add(new MailAddress(NAME + "@acmepoint.com"));
                System.Data.DataTable F1 = GetMenu.GETENG();
                for (int i = 0; i <= F1.Rows.Count - 1; i++)
                {
                    string ff = F1.Rows[i]["DataValue"].ToString();
                    message.To.Add(ff + "@acmepoint.com");

                }
    

                message.Subject = textBox5.Text + " " + RMAA + " 工程師覆判報告不符";
                message.Body = template;


                string OutPutFile = lsAppDir + "\\Excel\\temp\\rma";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {

                    string m_File = "";

                    m_File = file;
                    data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                    //附件资料
                    ContentDisposition disposition = data.ContentDisposition;


                    // 加入邮件附件
                    message.Attachments.Add(data);

                }



                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                try
                {
                    client.Send(message);
                    data.Dispose();


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }

            //}
        }


        private void MAILFILE2(string RMAA)
        {
            string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL3.xls";



            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\RMA.htm";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

            template = template.Replace("##Content##", "Dear,");
            template = template.Replace("##MARK##", RMAA);
            MailMessage message = new MailMessage();


            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            System.Data.DataTable F1 = GetMenu.GETENG();
            for (int i = 0; i <= F1.Rows.Count - 1; i++)
            {
                string ff = F1.Rows[i]["DataValue"].ToString();
                message.To.Add(ff + "@acmepoint.com");

            }

            message.Subject = RMAA + " 工程師覆判報告不符";
            message.Body = template;


            string OutPutFile = lsAppDir + "\\Excel\\temp\\rma";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                string m_File = "";

                m_File = file;
                data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                //附件资料
                ContentDisposition disposition = data.ContentDisposition;


                // 加入邮件附件
                message.Attachments.Add(data);

            }



            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            try
            {
                client.Send(message);
                data.Dispose();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            //}
        }
        private void GetExcelProduct(string ExcelFile)
        
{

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
   
            excelApp.DisplayAlerts =false;

            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

    
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }
            int J1 = 0;
            string NO;
            string SN;
            string VER;
            string DATE;
            string GRADE;
            string PLACE;
            string DEFECT;
            string IQC;
            int k = 0;

            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {
           
                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
           
                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
            
           
                    int RMA1 = RMA.Length;
                    if (RMA1 != 8)
                    {
                        for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                            sTemp = (string)range.Text;
                            string RMAS = sTemp.Trim();
                            int F4 = RMAS.ToUpper().IndexOf("RMA#");
                            if (F4 != -1)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                                sTemp = (string)range.Text;
                                string RMAS2 = sTemp.Trim();
                                int RMA2 = RMAS2.Length;
                                if (RMA2 != 8)
                                {
                                    G1 = 1;
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }


                    }
                    //SAP檢查
                    SG(textBox2.Text);
                    if (RMA != TRMA)
                    {
                        G1 = 1;
                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 3]);
                sTemp = (string)range.Text;
                string MODEL = sTemp.Trim();
                if (MODEL.ToUpper() != TMODEL.ToUpper())
                {
                    G1 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                textBox4.Text = RMA;
           
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();

                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                        ROW = iRecord;
                    }

                    if (k == 0)
                    {
                        int F = NO.ToUpper().IndexOf("NO");
                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);

                        sTemp = (string)range.Text;
                        VER = sTemp.Trim();


                        int V1 = VER.ToUpper().IndexOf("V");
                        int V2 = VER.ToUpper().IndexOf(".");
         
                            if (V1 != -1 || V2 != -1 || String.IsNullOrEmpty(VER))
                            {
                                if (F == -1 && IsNumber(NO))
                                {
                                    if (VER != "V")
                                    {
                                        G1 = 1;
                                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                    }
                                }

                            }
                            //SAP檢查
                            if (IsNumber(NO))
                            {
                                if (VER != TVER)
                                {
                                    G1 = 1;
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                }
                            }

                        int P3 = NO.ToUpper().IndexOf("MODEL");
                        if (P3 != -1)
                        {
                            int VERP = VER.ToUpper().IndexOf("V.");
                            if (VERP != -1)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        sTemp = (string)range.Text;
                        DATE = sTemp.Trim();
                        int LEN = DATE.Length;
                        int LEN1 = DATE.ToUpper().IndexOf("/");
                        if (!String.IsNullOrEmpty(DATE))
                        {
                            if (LEN != 5 || LEN1 == -1)
                            {
                                if (F == -1 && IsNumber(NO))
                                {
                                    G1 = 1;
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                }

                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        IQC = sTemp.Trim();


                        int IQC1 = IQC.IndexOf("IQC");
                        int LR = IQC.IndexOf("LR");
                        int FR = IQC.IndexOf("FR");

                        if (IQC1 == -1 && LR == -1 && FR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

    

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();

                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "PANEL");
                        if (K1.Rows.Count == 0 )
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");
                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1 && OOW == -1 && ACME == -1)
                        {

                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);

                        sTemp = (string)range.Text;
                        PLACE = sTemp.Trim();

                        if (F == -1 && IsNumber(NO))
                        {
                            //China
                            if (PLACE != "China" && PLACE != "Taiwan" && PLACE != "Turkey")
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
           

                        int P2 = NO.ToUpper().IndexOf("CUST");
                        if (P2 != -1)
                        {
                            textBox5.Text = VER;
                        }



                      
                    }

             
                }

                //SAP檢查
                string QTY = J1.ToString();
                if (QTY != TQTY)
                {
                    G1 = 1;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                    range.Value2 = RMA + "數量錯誤";
                }
         
            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                try
                {




                    excelSheet.SaveAs(FileName1, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

         
                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }

        private void GetExcelProduct4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);


            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }
            int J1 = 0;
            string NO;
            string SN;
            string VER;
            string DATE;
            string GRADE;
            string PLACE;
            string DEFECT;
            string IQC;
            int k = 0;
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);

                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
                int RMA1 = RMA.Length;
                if (RMA1 != 8)
                {
                    for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                        sTemp = (string)range.Text;
                        string RMAS = sTemp.Trim();
                        int F4 = RMAS.ToUpper().IndexOf("RMA#");
                        if (F4 != -1)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            sTemp = (string)range.Text;
                            string RMAS2 = sTemp.Trim();
                            int RMA2 = RMAS2.Length;
                            if (RMA2 != 8)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                }
                textBox4.Text = RMA;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();

                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                        ROW = iRecord;
                    }

                    if (k == 0)
                    {
                        int F = NO.ToUpper().IndexOf("NO");
                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

           

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        sTemp = (string)range.Text;
                        IQC = sTemp.Trim();


                        int IQC1 = IQC.IndexOf("IQC");
                        int LR = IQC.IndexOf("LR");
                        int FR = IQC.IndexOf("FR");

                        if (IQC1 == -1 && LR == -1 && FR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }



                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();

                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "PANEL");
                        if (K1.Rows.Count == 0)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");
                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1 && OOW == -1 && ACME == -1)
                        {

                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);

                        sTemp = (string)range.Text;
                        PLACE = sTemp.Trim();

                        if (F == -1 && IsNumber(NO))
                        {
                            //China
                            if (PLACE != "China" && PLACE != "Taiwan" && PLACE != "Turkey")
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }


                      

                    }
                }
          
            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                try
                {




                    excelSheet.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }
        private void GetExcelProduct3(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);


            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }

            string NO;
            string SN;
            string S3;
            string S4;
            string S5;
            string GRADE;
            string DEFECT;
            string ENG;
            int J1 = 0;
            int k = 0;
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
                int RMA1 = RMA.Length;
                if (RMA1 != 8)
                {
                    G1 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                textBox4.Text = RMA;

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();
                    int F = NO.ToUpper().IndexOf("NO");
                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                    }

                    if (k == 0)
                    {

         


                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        sTemp = (string)range.Text;
                        S3 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S3))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        sTemp = (string)range.Text;
                        S4 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S4))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;

                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        S5 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S5))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();



                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "KIT");
                        if (K1.Rows.Count == 0)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");

                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        int P2 = NO.ToUpper().IndexOf("CUST");
                        if (P2 != -1)
                        {
                            textBox5.Text = S3;
                        }
                    }
                }

            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                try
                {


                    excelSheet.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }
        private void GetMULTI(string ExcelFile, int sd1, string CONTRACT)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible =  true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;


            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);


            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox6.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }
            int J1 = 0;
            string NO;
            string SN;
            string VER;
            string DATE;
            string GRADE;
            string PLACE;
            string DEFECT;
            string IQC;
            string REA;
            int k = 0;
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
                int RMA1 = RMA.Length;
                if (RMA1 != 8)
                {
                    G1 = 1;
                    G2 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                //SAP檢查
                SG(CONTRACT);
                if (RMA != TRMA)
                {
                    G1 = 1;
                    G2 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 3]);
                sTemp = (string)range.Text;
                string MODEL = sTemp.Trim();
                if (MODEL != TMODEL)
                {
                    G1 = 1;
                    G2 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

              
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();
                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                        ROW = iRecord;
                    }

                    if (k == 0)
                    {
                        int F = NO.ToUpper().IndexOf("NO");
                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);

                        sTemp = (string)range.Text;
                        VER = sTemp.Trim();


                        int V1 = VER.ToUpper().IndexOf("V");
                        int V2 = VER.ToUpper().IndexOf(".");

                        if (V1 != -1 || V2 != -1 || String.IsNullOrEmpty(VER))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                if (VER != "V")
                                {
                                    G1 = 1;
                                    G2 = 1;
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                }
                            }

                        }
                        //SAP檢查
                        if (IsNumber(NO))
                        {
                            if (VER != TVER)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P3 = NO.ToUpper().IndexOf("MODEL");
                        if (P3 != -1)
                        {
                            int VERP = VER.ToUpper().IndexOf("V.");
                            if (VERP != -1)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        sTemp = (string)range.Text;
                        DATE = sTemp.Trim();
                        int LEN = DATE.Length;
                        int LEN1 = DATE.ToUpper().IndexOf("/");

                        if (LEN != 5 || LEN1 == -1 || String.IsNullOrEmpty(DATE))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        IQC = sTemp.Trim();


                        int IQC1 = IQC.IndexOf("IQC");
                        int LR = IQC.IndexOf("LR");
                        int FR = IQC.IndexOf("FR");

                        if (IQC1 == -1 && LR == -1 && FR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }



                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();

                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "PANEL");
                        if (K1.Rows.Count == 0)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");
                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1 && OOW == -1 && ACME == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);

                        sTemp = (string)range.Text;
                        PLACE = sTemp.Trim();

                        if (IsNumber(NO))
                        {
                            if (PLACE != "China" && PLACE != "Taiwan" && PLACE != "Turkey")
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
     

                        int P2 = NO.ToUpper().IndexOf("CUST");
                        if (P2 != -1)
                        {
                            
                            textBox5.Text = VER;
                        }

                    }
                }

                //SAP檢查
                string QTY = J1.ToString();
                if (QTY != TQTY)
                {
                    G1 = 1;
                    G2 = 1;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                    range.Value2 = RMA + "數量錯誤";
                }
            }
            finally
            {
                if (G1 == 1)
                {
                    string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                    try
                    {




                        excelSheet.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }



                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }

        private void GetMULTI2(string ExcelFile, int sd1)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;



            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);


            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox6.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }

            string NO;
            string SN;
            string S3;
            string S4;
            string S5;
            string GRADE;
            string DEFECT;
            int J1 = 0;
            int k=0;
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
                int RMA1 = RMA.Length;
                if (RMA1 != 8)
                {
                    G1 = 1;
                    G2 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();
                    int F = NO.ToUpper().IndexOf("NO");
                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                    }

                    if (k == 0)
                    {
   


                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        sTemp = (string)range.Text;
                        S3 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S3))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        sTemp = (string)range.Text;
                        S4 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S4))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        S5 = sTemp.Trim();

                        if (String.IsNullOrEmpty(S5))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();

                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "KIT");
                        if (K1.Rows.Count == 0)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");

                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

             
                        int P2 = NO.ToUpper().IndexOf("CUST");
                        if (P2 != -1)
                        {
                            textBox5.Text = S3;
                        }
                    }
                }


            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                try
                {


                    excelSheet.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }

        private void GetMULT2(string ExcelFile, int sd1)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;


            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);


            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox6.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }
            int J1 = 0;
            string NO;
            string SN;
            string VER;
            string DATE;
            string GRADE;
            string PLACE;
            string DEFECT;
            string IQC;
            string REA;
            int k = 0;
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                sTemp = (string)range.Text;
                string RMA = sTemp.Trim();
                int RMA1 = RMA.Length;
                if (RMA1 != 8)
                {
                    G1 = 1;
                    G2 = 1;
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    sTemp = (string)range.Text;
                    NO = sTemp.Trim();
                    int F4 = NO.ToUpper().IndexOf("製表者");
                    if (F4 != -1)
                    {
                        k = iRecord;
                        ROW = iRecord;
                    }

                    if (k == 0)
                    {
                        int F = NO.ToUpper().IndexOf("NO");
                        if (F == -1 && IsNumber(NO))
                        {
                            J1++;

                            if (J1.ToString() != NO)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        int P4 = NO.ToUpper().IndexOf("製表者");
                        if (P4 != -1)
                        {
                            int N1 = NO.Length;

                            if (N1 == 4)
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        sTemp = (string)range.Text;
                        SN = sTemp.Trim();

                        int iPos = SN.IndexOf("-");

                        if (iPos > 0 || String.IsNullOrEmpty(SN))
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        sTemp = (string)range.Text;
                        IQC = sTemp.Trim();


                        int IQC1 = IQC.IndexOf("IQC");
                        int LR = IQC.IndexOf("LR");
                        int FR = IQC.IndexOf("FR");

                        if (IQC1 == -1 && LR == -1 && FR == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }



                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                        sTemp = (string)range.Text;
                        DEFECT = sTemp.Trim();

                        System.Data.DataTable K1 = GetMenu.GETDEFECTCODE(DEFECT, "PANEL");
                        if (K1.Rows.Count == 0)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        sTemp = (string)range.Text;
                        GRADE = sTemp.Trim();


                        int OK = GRADE.IndexOf("OK");
                        int NDF = GRADE.IndexOf("NDF");
                        int NG = GRADE.IndexOf("NG");
                        int CR = GRADE.IndexOf("CR");
                        int OOW = GRADE.IndexOf("OOW");
                        int ACME = GRADE.IndexOf("ACME");
                        if (OK == -1 && NDF == -1 && NG == -1 && CR == -1 && OOW == -1 && ACME == -1)
                        {
                            if (F == -1 && IsNumber(NO))
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);

                        sTemp = (string)range.Text;
                        PLACE = sTemp.Trim();

                        if (IsNumber(NO))
                        {
                            if (PLACE != "China" && PLACE != "Taiwan" && PLACE != "Turkey")
                            {
                                G1 = 1;
                                G2 = 1;
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }



                    }
                }


            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\EXCEL.xls";



                try
                {




                    excelSheet.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



                //增加一個 Close
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }
        private void GetExcelProduct2(string RMA)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

   

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Excel\\EXCEL2.xls";
            string lsAppDir3 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Excel\\temp\\EXCEL.xls";
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(lsAppDir3, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            int sd1 = Convert.ToInt16(this.comboBox1.SelectedIndex.ToString()) + 1;
            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            Microsoft.Office.Interop.Excel.Workbook excelBook2 = excelApp.Workbooks.Open(lsAppDir, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox5.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }

            object Cell_From = null;
            object Cell_To = null;
   
            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {

  


            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
       
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\rma\\" + RMA;
                if (comboBox4.Text == "KIT" || comboBox4.Text == "Open frame")
                {
                    // 指定 複製 的範圍
                     Cell_From = "A1";
                     Cell_To = "G" + Convert.ToString(iRowCnt + 1);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);

                }
                else
                {
                    // 指定 複製 的範圍
                     Cell_From = "A1";
                     Cell_To = "I" + Convert.ToString(iRowCnt + 1);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                }

                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook2.Sheets.get_Item(1);
                excelSheet2.Paste(oMissing, oMissing);
                range = excelSheet2.get_Range(Cell_From, Cell_To);
                range.Select();
                range.Columns.AutoFit();

                try
                {

                //    excelSheet2.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    excelSheet2.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }



                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
                excelBook2.Close(oMissing, oMissing, oMissing);
                //Quit
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet2);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook2);
       
                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                excelBook2 = null;
                excelSheet2 = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }
        private void GetEMULTI3(int sd1,string RMA)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths



            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Excel\\EXCEL2.xls";
            string lsAppDir3 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Excel\\temp\\EXCEL.xls";
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(lsAppDir3, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

          
            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            Microsoft.Office.Interop.Excel.Workbook excelBook2 = excelApp.Workbooks.Open(lsAppDir, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            if (!checkBox6.Checked)
            {
                if (iRowCnt > 200)
                {
                    iRowCnt = 200;
                }
            }
            object Cell_From = null;
            object Cell_To = null;

            Microsoft.Office.Interop.Excel.Range range = null;

            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);
            try
            {




            }
            finally
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\rma\\" + RMA;


                if (comboBox5.Text == "KIT" || comboBox5.Text == "Open frame")
                {
                    // 指定 複製 的範圍
                    Cell_From = "A1";
                    Cell_To = "G" + Convert.ToString(iRowCnt + 1);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);

                }
                else
                {
                    // 指定 複製 的範圍
                    Cell_From = "A1";
                    Cell_To = "I" + Convert.ToString(iRowCnt + 1);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                }

                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook2.Sheets.get_Item(1);
                excelSheet2.Paste(oMissing, oMissing);
                range = excelSheet2.get_Range(Cell_From, Cell_To);
                range.Select();
                range.Columns.AutoFit();

                try
                {


                    excelSheet2.SaveAs(FileName1, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }



                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
                excelBook2.Close(oMissing, oMissing, oMissing);
                //Quit
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet2);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook2);

                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                excelBook2 = null;
                excelSheet2 = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);



            }

        }
        public static bool IsNumber(string strNumber)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"^\d+(\.)?\d*$");
            return r.IsMatch(strNumber);
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(FileName1);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void DELETEFILE2()
        {
            try
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\rma";
                string[] filenames = Directory.GetFiles(FileName1);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }


        private void button10_Click(object sender, EventArgs e)
        {
            //Process[] _proceses = null;
            //_proceses = Process.GetProcessesByName("ACME");
            //foreach (Process proces in _proceses)
            //{
            //    proces.Kill();
            //}
            //try
            ////{
            //TruncateTable();

                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {
                    TruncateTable();
                    GetExcelContentGD44(opdf.FileName);

                    dataGridView2.DataSource = GetOrderData2();
                }
                
              
        }
        private void GetExcelContentGD44(string ExcelFile)
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




            string id1;
            string id2;

            for (int i = 2; i <= iRowCnt; i++)
            {



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id1 = range.Text.ToString().Trim();
                
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString().Trim();

                try
                {
                    if (!String.IsNullOrEmpty(id1))
                    {


                        AddProduct(id1, id2);

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
        public void AddProduct(string DTYPE, string DefectCode)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO [AcmeSqlSP].[dbo].[RMA_DEFECTCODE]");
            sb.Append("            (DTYPE,DefectCode)");
            sb.Append("      VALUES(@DTYPE,@DefectCode)");

            sb.Append(" INSERT INTO [AcmeSqlSPDRS].[dbo].[RMA_DEFECTCODE]");
            sb.Append("            (DTYPE,DefectCode)");
            sb.Append("      VALUES(@DTYPE,@DefectCode)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
            command.Parameters.Add(new SqlParameter("@DefectCode", DefectCode));

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
        private void TruncateTable()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("truncate table ACMESQLSP.DBO.RMA_DEFECTCODE ");
            sb.Append("truncate table ACMESQLSPDRS.DBO.RMA_DEFECTCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


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
        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DTYPE [TYPE],DEFECTCODE FROM RMA_DEFECTCODE ORDER BY DTYPE DESC,DEFECTCODE ");

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
        private System.Data.DataTable GetTEMP(string U_AUO_RMA_NO, string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select U_RMODEL MODEL,U_RVER VER,U_YETQTY QTY from   OCTR where U_AUO_RMA_NO=@U_AUO_RMA_NO AND U_RMA_NO=@U_RMA_NO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
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

        private System.Data.DataTable GETCUST(string U_Cusname_S)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT  TOP 1 CstmrCode  FROM OCTR WHERE U_Cusname_S =@U_Cusname_S ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_Cusname_S", U_Cusname_S));

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
        private System.Data.DataTable GetTEMPT(string CONTRACTID)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select U_RMA_NO RMA,U_RMODEL MODEL,U_RVER VER,U_rquinity QTY from   OCTR where CONTRACTID=@CONTRACTID  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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
        private System.Data.DataTable MakeMain2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("片序", typeof(string));
            dt.Columns.Add("箱序", typeof(string));
            dt.Columns.Add("棧板號", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("收採單號", typeof(string));
            dt.Columns.Add("供應商名稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("INVOICE", typeof(string));

            dt.Columns.Add("INVOICE日期", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("訂單單號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));

            return dt;
        }

        private void GetExcelContent3(string ExcelFile)
        {

            //Create an Excel App
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

            //Find the predefined barcode cell into the worksheet
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);


            Microsoft.Office.Interop.Excel.Range range = null;




            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            dtCost2 = MakeMain2();
            for (int i = 2; i <= iRowCnt; i++)
            {

                string SO;
                string CO;
                string PO;
                string CARTON = "";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                SO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                CO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                PO = range.Text.ToString().Trim();
                string TT = SO + CO + PO;

                if (TT != "")
                {
                    System.Data.DataTable dt = GetScheSap2(SO, CO, PO);
                    DataRow dr = null;
                    if (dt.Rows.Count > 0)
                    {

                        dr = dtCost2.NewRow();
                        string INVOICE = dt.Rows[0]["INVOICE"].ToString();
                        string ITEMCODE = dt.Rows[0]["產品編號"].ToString();
                        dt.Columns.Add("片序", typeof(string));
                        dt.Columns.Add("箱序", typeof(string));
                        dt.Columns.Add("棧板號", typeof(string));

                        dr["片序"] = SO;
                        dr["箱序"] = CO;
                        dr["棧板號"] = PO;

                        if (String.IsNullOrEmpty(CO))
                        {
                            if (!String.IsNullOrEmpty(PO))
                            {
                                System.Data.DataTable GETCART = GetCARTON3(PO);
                                if (GETCART.Rows.Count > 0)
                                {

                                    CARTON = GETCART.Rows[0][0].ToString();
                                }
                            }
                            if (!String.IsNullOrEmpty(SO))
                            {
                                System.Data.DataTable GETCART = GetCARTON(SO);
                                System.Data.DataTable GETCART2 = GetCARTON2(SO);
                                if (GETCART.Rows.Count > 0)
                                {

                                    CARTON = GETCART.Rows[0][0].ToString();
                                }
                                else if (GETCART2.Rows.Count > 0)
                                {

                                    CARTON = GETCART2.Rows[0][0].ToString();
                                }

                            }


                        }
                        else
                        {
                            CARTON = CO;
                        }

                        dr["過帳日期"] = dt.Rows[0]["過帳日期"].ToString();
                        dr["採購單號"] = dt.Rows[0]["採購單號"].ToString();
                        dr["收採單號"] = dt.Rows[0]["收採單號"].ToString();

                        dr["供應商名稱"] = dt.Rows[0]["供應商名稱"].ToString();
                        dr["產品編號"] = ITEMCODE;
                        dr["品名"] = dt.Rows[0]["品名"].ToString();
                        dr["數量"] = dt.Rows[0]["數量"].ToString();
                        dr["INVOICE"] = dt.Rows[0]["INVOICE"].ToString();
                        dr["INVOICE日期"] = dt.Rows[0]["INVOICE日期"].ToString();
                        dr["備註"] = dt.Rows[0]["備註"].ToString();

                        System.Data.DataTable TODLN = GetODLN(INVOICE, ITEMCODE, CARTON);
                        if (TODLN.Rows.Count > 0)
                        {
                            UPDATITEM(INVOICE, ITEMCODE);
                            dr["出貨日期"] = TODLN.Rows[0]["出貨日期"].ToString();
                            dr["訂單單號"] = TODLN.Rows[0]["訂單單號"].ToString();
                            dr["客戶名稱"] = TODLN.Rows[0]["客戶名稱"].ToString();

                        }

                        dtCost2.Rows.Add(dr);


                    }
                    else
                    {

                        dr = dtCost2.NewRow();

                        dt.Columns.Add("片序", typeof(string));
                        dt.Columns.Add("箱序", typeof(string));
                        dt.Columns.Add("棧板號", typeof(string));

                        dr["片序"] = SO;
                        dr["箱序"] = CO;
                        dr["棧板號"] = PO;


                        dr["過帳日期"] = "";
                        dr["採購單號"] = "";
                        dr["收採單號"] = "";

                        dr["供應商名稱"] = "";
                        dr["產品編號"] = "";
                        dr["品名"] = "";
                        dr["數量"] = "";
                        dr["INVOICE"] = "";
                        dr["INVOICE日期"] = "";
                        dr["備註"] = "";

                        dr["出貨日期"] = "";
                        dr["訂單單號"] = "";
                        dr["客戶名稱"] = "";

                        dtCost2.Rows.Add(dr);
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
            //可以將 Excel.exe 清除
            System.GC.WaitForPendingFinalizers();


        }
        public void UPDATITEM(string INVOICE_NO, string ITEMCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE ACMESQLSP.DBO.WH_AUO   SET ITEMCODE=@ITEMCODE ");
            sb.Append(" WHERE ID IN (SELECT T0.ID FROM ACMESQLSP.DBO.WH_AUO  T0  ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.AP_INVOICEIN T1 ON (T0.INVOICE_NO =T1.INV AND T0.CARTON_NO =T1.CARTON)  ");
            sb.Append(" WHERE INVOICE_NO=@INVOICE_NO   AND T1.ITEMCODE=@ITEMCODE  ");
            sb.Append(" AND ISNULL(T0.ITEMCODE,'') ='')");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICE_NO", INVOICE_NO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //Response.Write(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }
        }
        private System.Data.DataTable GetODLN(string INVOICE_NO, string ITEMCODE, string CARTON)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                SELECT DISTINCT T9.CARDNAME 客戶名稱,Convert(varchar(10),(t9.docdate),112) 出貨日期,T5.DOCENTRY 訂單單號,CAST(T4.Quantity AS INT) 數量   FROM  ACMESQLSP.DBO.AP_INVOICEIN T1 ");
            sb.Append("                    LEFT JOIN ACMESQLSP.DBO.WH_Item4 T2 ON (T1.WHNO =T2.ShippingCode AND T1.ITEMCODE =T2.ItemCode)  ");
            sb.Append("                    left join ACMESQL02.DBO.dln1 t4 on (t4.baseentry=T2.DOCENTRY and  t4.baseline=t2.linenum  and t4.basetype='17')  ");
            sb.Append("                    left join ACMESQL02.DBO.odln t9 on (t4.docentry=T9.docentry )  ");
            sb.Append("					left join ACMESQL02.DBO.rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");

            sb.Append("           WHERE T1.INV=@INVOICE_NO   AND T1.ITEMCODE=@ITEMCODE ");
            if (!String.IsNullOrEmpty(CARTON))
            {
                sb.Append("         AND T1.CARTON =@CARTON ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE_NO", INVOICE_NO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON_NO CARTON FROM WH_AUO WHERE SHIPPING_NO=@SHIPPING ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON3(string PALLET_NO)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON_NO CARTON FROM WH_AUO WHERE PALLET_NO=@PALLET_NO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@PALLET_NO", PALLET_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetCARTON2(string SHIPPING)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARTON FROM RMA_INVOICEOUT WHERE SHIPPING=@SHIPPING ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetTEMP2(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, string U_YETQTY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select U_RMODEL  from   OCTR where U_AUO_RMA_NO=@U_AUO_RMA_NO AND SUBSTRING(U_RMODEL,1,9) = @U_RMODEL AND U_RVER = @U_RVER AND U_YETQTY=@U_YETQTY  AND ISNULL(U_ACME_RECEDATE,'') <> ''  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_YETQTY", U_YETQTY));
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

        private System.Data.DataTable GetScheSap2(string SO, string CO, string PO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select  Convert(varchar(10),(t0.docdate),112) 過帳日期,T5.DOCENTRY 採購單號,t0.docentry 收採單號 ,t0.cardname 供應商名稱");
            sb.Append(" ,t1.ItemCode 產品編號,t1.dscription 品名,cast(t1.quantity as int) 數量, ");
            sb.Append(" LTRIM(RTRIM(T0.U_ACME_INV)) INVOICE,Convert(varchar(10),T0.U_ACME_Invoice ,111) INVOICE日期,T0.Comments 備註  from opdn t0 ");
            sb.Append(" LEFT JOIN PDN1 T1 ON(T0.docentry=t1.docentry) ");
            sb.Append(" LEFT JOIN POR1 T5 ON (T5.docentry=T1.baseentry AND T5.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN OITM T10 ON T1.ITEMCODE = T10.ITEMCODE    ");
            sb.Append(" WHERE  ISNULL(T10.U_GROUP,'') <> 'Z&R-費用類群組' ");



            string STYPE = "";
            string SER = "";
            if (SO != "")
            {
                SER = SO;
                STYPE = "A";
            }
            else if (CO != "")
            {
                SER = CO;
                STYPE = "C";
            }
            else if (PO != "")
            {
                SER = PO;
                STYPE = "B";
            }
            System.Data.DataTable DTSFRMA = GetSERFRMA(SER.Trim(), STYPE);
            System.Data.DataTable DTSF = GetSERF(SER.Trim(), STYPE);
            if (DTSF.Rows.Count > 0)
            {
                string INVOICE = DTSF.Rows[0]["INVOICE"].ToString();
                string MODEL = DTSF.Rows[0]["MODEL"].ToString();
                string GRADE = DTSF.Rows[0]["GRADE"].ToString();
                string PARTNO = DTSF.Rows[0]["PARTNO"].ToString();
                string ITEM = "";
                if (String.IsNullOrEmpty(MODEL))
                {
                    System.Data.DataTable DTSF2 = GetSERF2(INVOICE, PARTNO, GRADE);
                    if (DTSF2.Rows.Count > 0)
                    {
                        ITEM = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND  ( t0.u_acme_inv like '%" + INVOICE + "%' OR t0.U_AUOINV like '%" + INVOICE + "%' )  AND T1.ITEMCODE = '" + ITEM + "' ");

                        UPDATESER(ITEM, INVOICE, "", GRADE, PARTNO);
                    }
                    else
                    {
                        sb.Append("and 1=0  ");
                    }


                }
                else
                {

                    System.Data.DataTable DTS = GetSER(SER.Trim(), STYPE);
                    ITEM = DTS.Rows[0][1].ToString();
                    System.Data.DataTable DTSF2 = GetSERF3(INVOICE, ITEM);
                    if (DTSF2.Rows.Count > 0)
                    {
                        string ITEMCODE = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND ( t0.u_acme_inv = '" + INVOICE + "' OR  t0.U_AUOINV = '" + INVOICE + "') AND T1.ITEMCODE = '" + ITEMCODE + "' ");
                        UPDATESER(ITEMCODE, INVOICE, MODEL, GRADE, PARTNO);
                    }
                    else
                    {
                        sb.Append("and 1=0  ");
                    }

                }

            }
            else if (DTSFRMA.Rows.Count > 0)
            {

                string INVOICE = DTSFRMA.Rows[0]["INVOICE"].ToString();
                string MODEL = DTSFRMA.Rows[0]["MODEL"].ToString();
                string PARTNO = DTSFRMA.Rows[0]["PARTNO"].ToString();
                string MODEL1 = DTSFRMA.Rows[0]["MODEL1"].ToString();
                string VER = DTSFRMA.Rows[0]["VER"].ToString();
                string CARTON = DTSFRMA.Rows[0]["CARTON"].ToString();
                string ITEM = "";
                if (String.IsNullOrEmpty(MODEL))
                {
                    System.Data.DataTable DTSF2 = GetSERF2RMA(INVOICE, PARTNO);
                    if (DTSF2.Rows.Count > 0)
                    {
                        ITEM = DTSF2.Rows[0][0].ToString();
                        sb.Append("  AND  (t0.u_acme_inv = '" + INVOICE + "' OR t0.U_AUOINV = '" + INVOICE + "') AND T1.ITEMCODE = '" + ITEM + "' ");

                        UPDATESERRMA(ITEM, INVOICE, "", PARTNO);
                    }
                    else
                    {
                        sb.Append("and 1=0  ");
                    }

                }
                else
                {

                    System.Data.DataTable DTSF2 = GetSERF2RMA2(INVOICE, MODEL1, VER);
                    if (DTSF2.Rows.Count > 0)
                    {
                        string ITEMCODE = "";
                        if (DTSF2.Rows.Count > 1)
                        {
                            System.Data.DataTable DTSF3 = GetSERF2RMA3(INVOICE, CARTON);
                            ITEMCODE = DTSF3.Rows[0][0].ToString();
                        }
                        else
                        {
                            ITEMCODE = DTSF2.Rows[0][0].ToString();
                        }
                        sb.Append("  AND  (t0.u_acme_inv = '" + INVOICE + "' OR  t0.U_AUOINV = '" + INVOICE + "' ) AND T1.ITEMCODE = '" + ITEMCODE + "' ");
                        UPDATESERRMA(ITEMCODE, INVOICE, MODEL, PARTNO);
                    }
                    else
                    {
                        sb.Append("and 1=0  ");
                    }

                }

            }
            else
            {
                sb.Append("and 1=0  ");
            }




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ODLN");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ODLN"];
        }

        private System.Data.DataTable GetSERF3(string INVOICE, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPDN T0 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" WHERE ( T0.U_ACME_INV  LIKE '%" + INVOICE + "%'  OR T0.U_AUOINV  LIKE '%" + INVOICE + "%'  ) ");
            sb.Append(" AND ACMESQLSP.dbo.fn_RemoveLastChar(T1.ITEMCODE)  COLLATE Chinese_Taiwan_Stroke_CI_AS =@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF2(string INVOICE, string PARTNO, string GRADE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV  LIKE '%" + INVOICE + "%' ");
            sb.Append(" AND T2.U_PARTNO =@PARTNO AND T2.U_GRADE =@GRADE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSER(string SHIPPING_NO, string TYPE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT T0.INVOICE_NO INVOICE,ACMESQLSP.dbo.fn_RemoveLastChar(T1.ITEMCODE) ITEMCODE FROM ACMESQLSP.DBO.WH_AUO  T0   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (CASE SUBSTRING(PART_NO,1,2) WHEN '91' THEN 'O' ELSE SUBSTRING(T0.MODEL_NO,1,1) END+  SUBSTRING(T0.MODEL_NO,2,CHARINDEX('.', T0.MODEL_NO)-2)=T1.U_TMODEL  COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.FINAL_GRADE=T1.U_GRADE  COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.PART_NO =T1.U_PARTNO COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            if (TYPE == "A")
            {
                sb.Append(" WHERE T0.SHIPPING_NO = @SHIPPING_NO ");
            }

            if (TYPE == "B")
            {
                sb.Append(" WHERE T0.PALLET_NO = @SHIPPING_NO ");
            }

            if (TYPE == "C")
            {
                sb.Append(" WHERE T0.CARTON_NO = @SHIPPING_NO ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING_NO", SHIPPING_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        public void UPDATESERRMA(string ITEMCODE, string INVOICE, string MODEL, string PART)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE RMA_INVOICEOUT SET ITEMCODE=@ITEMCODE WHERE  INVOICE=@INVOICE   AND  (PART =@PART　OR PART=@MODEL')    ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@PART", PART));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //Response.Write(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }
        }
        public void UPDATESER(string ITEMCODE, string INVOICE_NO, string MODEL_NO, string FINAL_GRADE, string PART_NO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_AUO SET ITEMCODE=@ITEMCODE WHERE  INVOICE_NO=@INVOICE_NO AND MODEL_NO=@MODEL_NO AND FINAL_GRADE=@FINAL_GRADE AND PART_NO=@PART_NO   ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE_NO", INVOICE_NO));
            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@FINAL_GRADE", FINAL_GRADE));
            command.Parameters.Add(new SqlParameter("@PART_NO", PART_NO));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //Response.Write(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }
        }
        private System.Data.DataTable GetSERF2RMA(string INVOICE, string PARTNO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE ");
            sb.Append(" AND T2.U_PARTNO =@PARTNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF(string SHIPPING_NO, string TYPE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT INVOICE_NO INVOICE,MODEL_NO MODEL,FINAL_GRADE GRADE,PART_NO PARTNO  FROM WH_AUO T0");

            if (TYPE == "A")
            {
                sb.Append(" WHERE T0.SHIPPING_NO = @SHIPPING_NO ");
            }

            if (TYPE == "B")
            {
                sb.Append(" WHERE T0.PALLET_NO = @SHIPPING_NO ");
            }

            if (TYPE == "C")
            {
                sb.Append(" WHERE T0.CARTON_NO = @SHIPPING_NO ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING_NO", SHIPPING_NO));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }

        private System.Data.DataTable GetSERFRMA(string SHIPPING, string TYPE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT INVOICE,CASE SUBSTRING(PART,1,2) WHEN '97' THEN PART END PARTNO,");
            sb.Append("  CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN PART END MODEL,");
            sb.Append("  CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN SUBSTRING(PART,0,CHARINDEX('.', PART)) END MODEL1,");
            sb.Append("   CASE WHEN SUBSTRING(PART,1,2) <>  '97' THEN SUBSTRING(PART,CHARINDEX('.', PART)+1,3) END VER,CARTON");
            sb.Append("   FROM RMA_INVOICEOUT");
            sb.Append("   WHERE ISNULL(PART,'') <>'' ");

            if (TYPE == "A")
            {
                sb.Append(" AND SHIPPING = @SHIPPING ");
            }


            if (TYPE == "C")
            {
                sb.Append(" AND CARTON = @SHIPPING ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetSERF2RMA2(string INVOICE, string MODEL1, string VER)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE FROM ACMESQL02.DBO.OPCH T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" WHERE T0.U_ACME_INV =@INVOICE ");
            sb.Append(" AND T2.U_TMODEL  LIKE '%" + MODEL1 + "%' AND T2.U_PARTNO LIKE '%" + VER + "%'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }

        private System.Data.DataTable GetSERF2RMA3(string INVOICE, string CARTON)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE FROM AP_INVOICEIN WHERE INV=@INVOICE AND CARTON=@CARTON");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "stkBillMain");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["stkBillMain"];
        }
        private System.Data.DataTable GetTEMP3(string U_AUO_RMA_NO, string U_RMODEL, string U_RVER, string U_YETQTY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select U_RMODEL from   OCTR where U_AUO_RMA_NO=@U_AUO_RMA_NO AND SUBSTRING(U_RMODEL,1,9) = @U_RMODEL AND U_RVER = @U_RVER AND U_YETQTY=@U_YETQTY AND ISNULL(U_ACME_RECEDATE,'') = ''  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_RMODEL", U_RMODEL));
            command.Parameters.Add(new SqlParameter("@U_RVER", U_RVER));
            command.Parameters.Add(new SqlParameter("@U_YETQTY", U_YETQTY));
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

        private System.Data.DataTable GetCONTRACTID(string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CONTRACTID FROM OCTR WHERE  U_RMA_NO=@U_RMA_NO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));

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
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("請匯入EXCEL");

                    return;
                }
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("請選擇分頁");

                    return;
                }

                GetVenderOUT(textBox1.Text);
                MessageBox.Show("匯入完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




  
        private void button13_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("RMAIND").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "RMAIND");
                }
                else
                {
                    GetMenu.UP(t1, "RMAIND");
                }

                textBox6.Text = t1;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            inint = 0;
            s1 = 0;
            string d = textBox6.Text;

            string[] filenames = Directory.GetFiles(d);
            foreach (string file in filenames)
            {
                string FileName = string.Empty;

                FileName = file;
                    this.textBox3.Text = FileName;
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    object oMissing = System.Reflection.Missing.Value;

                    Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    string Count_Sheet = excelBook.Sheets.Count.ToString();
                    int i = excelBook.Sheets.Count;

                    TempDt = MakeTable();
                    DataRow dr = null;
                    for (int xi = 1; xi <= i; xi++)
                    {
                        dr = TempDt.NewRow();
                        Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                        string X1 = xi.ToString();
                        string X2 = excelsheet.Name.ToString().Trim();
                        dr["NO"] = X1;
                        dr["EXCEL分頁"] = X2;
                        System.Data.DataTable CT1 = GetCONTRACTID(X2);
                        if (CT1.Rows.Count > 0)
                        {
                            string ID = CT1.Rows[0][0].ToString();
                            dr["契約號碼"] = ID;

                            TempDt.Rows.Add(dr);

                            Microsoft.Office.Interop.Excel.Range range = null;
                            int iRowCnt = excelsheet.UsedRange.Cells.Rows.Count;
          
                            string VER = "";
                            string JUD = "";
                            string SN = "";
                            for (int i2 = 3; i2 <= iRowCnt; i2++)
                            {

                              
                            
                                range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 2]);
                                SN = range.Text.ToString();

                                range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 3]);
                                VER = range.Text.ToString();

                                range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 8]);
                                JUD = range.Text.ToString().Trim();

                                if (!String.IsNullOrEmpty(SN) && !String.IsNullOrEmpty(JUD) && !String.IsNullOrEmpty(VER))
                                {
                                    System.Data.DataTable G1 = GetDRS3(SN, X2);

                                    if (G1.Rows.Count > 0)
                                    {
                                        for (int f = 0; f <= G1.Rows.Count - 1; f++)
                                        {
                                            string RMA = G1.Rows[f][0].ToString();
                                            string MESSAGE = X2 + "_" + SN + "與RMANO#" + RMA + "重複";

                                            MessageBox.Show(MESSAGE);
                                        }

                                    }
                                }
                            }
                        }
                   
                    }

                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                    excelApp = null;
                    excelBook = null;

                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();


                    dataGridView1.DataSource = TempDt;



                    S1();

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                FileName = openFileDialog1.FileName;
                this.textBox7.Text = FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;

                object oMissing = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                string Count_Sheet = excelBook.Sheets.Count.ToString();
                int i = excelBook.Sheets.Count;

                TempDt = MakeTableF();
                DataRow dr = null;
           
                   
                    Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
  
                    Microsoft.Office.Interop.Excel.Range range = null;
                    int iRowCnt = excelsheet.UsedRange.Cells.Rows.Count;
                    string RMANO = "";
                    string CUST = "";
             
                    for (int i2 = 1; i2 <= iRowCnt; i2++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 1]);

                        RMANO = range.Text.ToString();
             

                        range = ((Microsoft.Office.Interop.Excel.Range)excelsheet.UsedRange.Cells[i2, 2]);
                        CUST = range.Text.ToString();





                        if (!String.IsNullOrEmpty(RMANO) && !String.IsNullOrEmpty(CUST) )
                        {
                            dr = TempDt.NewRow();

                            dr["RMANO"] = RMANO;
                            dr["客戶"] = CUST;
                            TempDt.Rows.Add(dr);
                        }
                    }
    

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                excelApp = null;
                excelBook = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


                dataGridView4.DataSource = TempDt;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {

            if (dataGridView4.Rows.Count > 0)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int SS = 0; //  to be used as an index

                oCompany.CompanyDB = FA;
                oCompany.UserName = "R01";
                oCompany.Password = "1234";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    SAPbobsCOM.ServiceContracts oCUSTSERVICE = null;
                    oCUSTSERVICE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts);

                  

                    for (int i2 = 0; i2 <= dataGridView4.Rows.Count - 2; i2++)
                    {

                        DataGridViewRow row;

                        row = dataGridView4.Rows[i2];
                       
                        string RMANO = row.Cells["RMANO"].Value.ToString();
                        string CUST = row.Cells["客戶"].Value.ToString();
                        System.Data.DataTable K1 = GETCUST(CUST);
                        if (K1.Rows.Count > 0)
                        {
                            DateTime StartDate = DateTime.Now;
                            DateTime EndDate = StartDate.AddDays(30);
                            oCUSTSERVICE.CustomerCode = K1.Rows[0][0].ToString();
                            oCUSTSERVICE.ContractType = BoContractTypes.ct_ItemGroup;
                            oCUSTSERVICE.StartDate = StartDate;
                            oCUSTSERVICE.EndDate = EndDate;
                            oCUSTSERVICE.UserFields.Fields.Item("U_RMA_NO").Value = RMANO;
                            oCUSTSERVICE.UserFields.Fields.Item("U_Cusname_S").Value = CUST;
                            
                            string R1 = RMANO.Substring(0, 1);
                            if (R1 == "A")
                            {
                                oCUSTSERVICE.UserFields.Fields.Item("U_Routwharehouse").Value = "客戶端"; 
                            }
                            if (R1 == "R")
                            {
                                oCUSTSERVICE.UserFields.Fields.Item("U_Routwharehouse").Value = "內湖"; 
                            }
                            oCUSTSERVICE.Lines.ItemGroup = 1035;
                            int res = oCUSTSERVICE.Add();
                            if (res != 0)
                            {
                                MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                            }
                            else
                            {
                                SS = 1;


                            }
                        }
                    }

                    if (SS == 1)
                    {
                        MessageBox.Show("新增成功");
                    
                    }



                  




                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
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
                    GetExcelContent3(opdf.FileName);
                    if (dtCost2.Rows.Count > 0)
                    {
                        dataGridView5.DataSource = dtCost2;
                        ExcelReport.GridViewToExcel(dataGridView5);
                    }

                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
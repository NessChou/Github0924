using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ACME
{
    public partial class HR104 : Form
    {
        string NewFileName;
        string strCn = "Data Source=10.10.1.45;Initial Catalog=89206602;Persist Security Info=True;User ID=ehrview;Password=viewehr";
        public HR104()
        {
            InitializeComponent();
        }


        public System.Data.DataTable GetCHO5()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("          SELECT  convert(varchar(8), [WORK_CARD_DATE],112) EMPDATE,T0.EMPLOYEE_NO EmpID,T0.EMPLOYEE_CNAME CHNAME,COMPANY_CNAME Company,T2.DEPARTMENT_CNAME BU,CASE WHEN convert(varchar(8), CARD_DATA_DATE,112)='19000101' THEN '' ELSE convert(VARCHAR(5),CARD_DATA_DATE,114) END CLOCKTIME, ");
            sb.Append("                   case  ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='0' then  CASE ISNULL(T1.LEAVE_SETUP_CLASS,'') WHEN '' THEN '應刷未刷' ELSE T1.LEAVE_SETUP_CLASS END ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='1' then '遲到'  ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='2' then '早退' ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='4' then CASE ISNULL(T1.LEAVE_SETUP_CLASS,'') WHEN '' THEN '曠職' ELSE T1.LEAVE_SETUP_CLASS END ");
            sb.Append("                      END DOCTYPE, ");
            sb.Append("               	   case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(11),T0.[WORK_CARD_DATE],111)  ELSE CONVERT(VARCHAR(11),T3.ASK_LEAVE_START,111)    END 開始日期, ");
            sb.Append("               	     case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(5),T0.[WORK_CARD_DATE],114)  ELSE CONVERT(VARCHAR(5),T3.ASK_LEAVE_START,114)    END 開始時間, ");
            sb.Append("               		 	   case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(11),T0.CARD_DATA_DATE,111)  ELSE CONVERT(VARCHAR(11),T3.ASK_LEAVE_END,111)    END 結束日期, ");
            sb.Append("               	     case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(5),T0.CARD_DATA_DATE,114)  ELSE CONVERT(VARCHAR(5),T3.ASK_LEAVE_END,114)    END 結束時間, ");
            sb.Append("               		 EMPLOYEE_ACCOUNT ENGNAME,CASE WHEN T2.DEPARTMENT_CNAME='節能事業部' THEN 'A00400' WHEN SUBSTRING(T5.DEPARTMENT_CODE ,1,3)= 'A01' THEN 'A01000'  WHEN T5.DEPARTMENT_CNAME IN ('董事長','總經理') THEN T4.DEPARTMENT_CODE  ELSE T5.DEPARTMENT_CODE END DEP,T3.ASK_LEAVE_MEMO MEMO ");
            sb.Append("                 FROM [vwZZ_CARD_DATA_MATCH] T0 ");
            sb.Append("                 Left Join [vwZZ_ASK_LEAVE_DETAIL] T1 ON T1.[EMPLOYEE_NO] =T0.EMPLOYEE_NO and  convert(varchar(8),T1.ASK_LEAVE_DETAIL_START,112)=@WORK_CARD_DATE");
            sb.Append("                 LEFT JOIN vwZZ_EMPLOYEE T2 ON (T0.EMPLOYEE_NO =T2.EMPLOYEE_NO) ");
            sb.Append("                   LEFT JOIN [vwZZ_ASK_LEAVE] T3 ON (T1.ASK_LEAVE_ID =T3.ASK_LEAVE_ID) ");
            sb.Append("                LEFT JOIN vwZZ_DEPARTMENT T4 ON (T2.DEPARTMENT_ID=T4.DEPARTMENT_ID) ");
            sb.Append("                 LEFT JOIN vwZZ_DEPARTMENT T5 ON (T4.PART_DEPARTMENT_ID=T5.DEPARTMENT_ID) ");
            sb.Append("                 where 1=1 ");
            sb.Append("                 and T0.WORK_CARD_TYPE='0' AND T2.DEPARTMENT_CNAME <> '董事長' ");
            sb.Append("                and  convert(varchar(8), [WORK_CARD_DATE],112) =@WORK_CARD_DATE");
            sb.Append("                 and CARD_DATA_CODE <> '' AND CARD_DATA_CODE <>'3' AND EMPLOYEE_WORK_STATUS=1 ");
            if (globals.DBNAME == "達睿生")
            {
                sb.Append("  and COMPANY_CNAME = '達睿生科技發展深圳有限公司' ");
            }
            else
            {
                sb.Append(" and  COMPANY_CNAME <> '達睿生科技發展深圳有限公司' ");
            }
            sb.Append("  UNION ALL ");
            sb.Append("   SELECT   convert(varchar(8),T1.ASK_LEAVE_DETAIL_START,112) EMPDATE,T1.EMPLOYEE_NO EmpID,T1.EMPLOYEE_CNAME CHNAME,COMPANY_CNAME Company,T4.DEPARTMENT_CNAME BU,");
            sb.Append("   '' CLOCKTIME,T1.LEAVE_SETUP_CLASS DOCTYPE,");
            sb.Append("    CONVERT(VARCHAR(11),T3.ASK_LEAVE_START,111)     開始日期,  ");
            sb.Append("            CONVERT(VARCHAR(5),T3.ASK_LEAVE_START,114)     開始時間,  ");
            sb.Append("             CONVERT(VARCHAR(11),T3.ASK_LEAVE_END,111)     結束日期,  ");
            sb.Append("                         CONVERT(VARCHAR(5),T3.ASK_LEAVE_END,114)     結束時間,");
            sb.Append(" 							EMPLOYEE_ACCOUNT ENGNAME,CASE WHEN T4.DEPARTMENT_CNAME='節能事業部' THEN 'A00400' WHEN SUBSTRING(T5.DEPARTMENT_CODE ,1,3)= 'A01' THEN 'A01000'     WHEN T5.DEPARTMENT_CNAME IN ('董事長','總經理') THEN T4.DEPARTMENT_CODE  ELSE T5.DEPARTMENT_CODE END DEP,T3.ASK_LEAVE_MEMO MEMO  ");
            sb.Append(" 						  FROM [dbo].[vwZZ_ASK_LEAVE_DETAIL] T1");
            sb.Append("     LEFT JOIN [vwZZ_ASK_LEAVE] T3 ON (T1.ASK_LEAVE_ID =T3.ASK_LEAVE_ID)  ");
            sb.Append(" 	             LEFT JOIN vwZZ_EMPLOYEE T2 ON (T1.EMPLOYEE_NO =T2.EMPLOYEE_NO)  ");
            sb.Append(" 	                            LEFT JOIN vwZZ_DEPARTMENT T4 ON (T2.DEPARTMENT_ID=T4.DEPARTMENT_ID)  ");
            sb.Append("                             LEFT JOIN vwZZ_DEPARTMENT T5 ON (T4.PART_DEPARTMENT_ID=T5.DEPARTMENT_ID)  ");
            sb.Append("    WHERE convert(varchar(8),T1.ASK_LEAVE_DETAIL_START,112)=@WORK_CARD_DATE AND  JOB_CNAME in ('協理','總經理') ");
            sb.Append("    AND T1.EMPLOYEE_CNAME NOT IN ( ");
            sb.Append("          SELECT  T0.EMPLOYEE_CNAME  FROM [vwZZ_CARD_DATA_MATCH] T0 ");
            sb.Append("                 Left Join [vwZZ_ASK_LEAVE_DETAIL] T1 ON T1.[EMPLOYEE_NO] =T0.EMPLOYEE_NO and  convert(varchar(8),T1.ASK_LEAVE_DETAIL_START,112)=@WORK_CARD_DATE");
            sb.Append("                 LEFT JOIN vwZZ_EMPLOYEE T2 ON (T0.EMPLOYEE_NO =T2.EMPLOYEE_NO) ");
            sb.Append("                   LEFT JOIN [vwZZ_ASK_LEAVE] T3 ON (T1.ASK_LEAVE_ID =T3.ASK_LEAVE_ID) ");
            sb.Append("                LEFT JOIN vwZZ_DEPARTMENT T4 ON (T2.DEPARTMENT_ID=T4.DEPARTMENT_ID) ");
            sb.Append("                 LEFT JOIN vwZZ_DEPARTMENT T5 ON (T4.PART_DEPARTMENT_ID=T5.DEPARTMENT_ID) ");
            sb.Append("                 where 1=1 ");
            sb.Append("                 and T0.WORK_CARD_TYPE='0' AND T2.DEPARTMENT_CNAME <> '董事長' ");
            sb.Append("                and  convert(varchar(8), [WORK_CARD_DATE],112) =@WORK_CARD_DATE");
            sb.Append("                 and CARD_DATA_CODE <> '' AND CARD_DATA_CODE <>'3' AND EMPLOYEE_WORK_STATUS=1 ");
            if (globals.DBNAME == "達睿生")
            {
                sb.Append("  and COMPANY_CNAME = '達睿生科技發展深圳有限公司') ");
            }
            else
            {
                sb.Append(" and  COMPANY_CNAME <> '達睿生科技發展深圳有限公司') ");
            }
       
            sb.Append("   order by  COMPANY_CNAME,T0.EMPLOYEE_NO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WORK_CARD_DATE", toolStripTextBox1.Text));
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

        public System.Data.DataTable GetCHO6(string EmpID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("           SELECT        case  ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='0' then CASE ISNULL(T1.LEAVE_SETUP_CLASS,'') WHEN '' THEN '應刷未刷' ELSE T1.LEAVE_SETUP_CLASS END  ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='1' then '遲到'  ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='2' then '早退' ");
            sb.Append("                   when T0.[CARD_DATA_CODE]='4' then CASE ISNULL(T1.LEAVE_SETUP_CLASS,'') WHEN '' THEN '曠職' ELSE T1.LEAVE_SETUP_CLASS END ");
            sb.Append("                      END DOCTYPE, ");
            sb.Append("               	   case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(11),T0.[WORK_CARD_DATE],111)  ELSE CONVERT(VARCHAR(11),T3.ASK_LEAVE_START,111)    END 開始日期, ");
            sb.Append("               	     case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(5),T0.[WORK_CARD_DATE],114)  ELSE CONVERT(VARCHAR(5),T3.ASK_LEAVE_START,114)    END 開始時間, ");
            sb.Append("               		 	   case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(11),T0.CARD_DATA_DATE,111)  ELSE CONVERT(VARCHAR(11),T3.ASK_LEAVE_END,111)    END 結束日期, ");
            sb.Append("               	     case  when T0.[CARD_DATA_CODE]='1' then CONVERT(VARCHAR(5),T0.CARD_DATA_DATE,114)  ELSE CONVERT(VARCHAR(5),T3.ASK_LEAVE_END,114)    END 結束時間");
            sb.Append("                 FROM [vwZZ_CARD_DATA_MATCH] T0 ");
            sb.Append("                 Left Join [vwZZ_ASK_LEAVE_DETAIL] T1 ON T1.[EMPLOYEE_NO] =T0.EMPLOYEE_NO and  convert(varchar(8),T1.ASK_LEAVE_DETAIL_START,112)=@WORK_CARD_DATE ");
            sb.Append("                 LEFT JOIN vwZZ_EMPLOYEE T2 ON (T0.EMPLOYEE_NO =T2.EMPLOYEE_NO) ");
            sb.Append("                   LEFT JOIN [vwZZ_ASK_LEAVE] T3 ON (T1.ASK_LEAVE_ID =T3.ASK_LEAVE_ID) ");
            sb.Append("                 where 1=1 ");
            sb.Append("                 and T0.WORK_CARD_TYPE='0' AND T2.DEPARTMENT_CNAME <> '董事長' ");
            sb.Append("                and  convert(varchar(8), [WORK_CARD_DATE],112) =@WORK_CARD_DATE AND T0.EMPLOYEE_NO=@EmpID ");
            sb.Append("                 and CARD_DATA_CODE <> '' AND CARD_DATA_CODE <>'3' AND EMPLOYEE_WORK_STATUS=1 ");
            if (globals.DBNAME == "達睿生")
            {
                sb.Append("  and COMPANY_CNAME = '達睿生科技發展深圳有限公司' ");
            }
            else
            {
                sb.Append(" and  COMPANY_CNAME <> '達睿生科技發展深圳有限公司' ");
            }
            sb.Append("   order by  COMPANY_CNAME,T0.EMPLOYEE_NO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WORK_CARD_DATE", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@EmpID", EmpID));
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
        private void hR_Main104BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.hR_Main104BindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.hR);

            MessageBox.Show("更新成功");

        }



        private void button3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = GetGROUP4();
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("已有資料");
                return;
            }


            System.Data.DataTable dt = GetCHO5();
            if (dt.Rows.Count > 0)
            {

                for (int h = 0; h <= dt.Rows.Count - 1; h++)
                {

                    AddAUOGD4(dt.Rows[h]["EmpID"].ToString(), dt.Rows[h]["Company"].ToString(), toolStripTextBox1.Text, dt.Rows[h]["CHNAME"].ToString(), dt.Rows[h]["ENGNAME"].ToString(), dt.Rows[h]["BU"].ToString(), dt.Rows[h]["CLOCKTIME"].ToString(), dt.Rows[h]["DOCTYPE"].ToString(), dt.Rows[h]["開始日期"].ToString(), dt.Rows[h]["開始時間"].ToString(), dt.Rows[h]["結束日期"].ToString(), dt.Rows[h]["結束時間"].ToString(), dt.Rows[h]["DEP"].ToString(), dt.Rows[h]["MEMO"].ToString());
                }



                try
                {
                    this.hR_Main104TableAdapter.Fill(this.hR.HR_Main104, toolStripTextBox1.Text);


                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("今日無資料");
            }
        }

        public void AddAUOGD4(string EmpID, string company, string EMPDATE, string CHNAME, string ENGNAME, string BU, string CLOCKTIME, string DOCTYPE, string STARTDATE, string STARTTIME, string ENDDATE, string ENDTIME, string DEP,string MEMO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into hr_main104(EmpID,company,EMPDATE,CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE,STARTTIME,ENDDATE,ENDTIME,DEP,MEMO) values(@EmpID,@company,@EMPDATE,@CHNAME,@ENGNAME,@BU,@CLOCKTIME,@DOCTYPE,@STARTDATE,@STARTTIME,@ENDDATE,@ENDTIME,@DEP,@MEMO)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EmpID", EmpID));
            command.Parameters.Add(new SqlParameter("@company", company));
            command.Parameters.Add(new SqlParameter("@EMPDATE", EMPDATE));
            command.Parameters.Add(new SqlParameter("@CHNAME", CHNAME));
            command.Parameters.Add(new SqlParameter("@ENGNAME", ENGNAME));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@CLOCKTIME", CLOCKTIME));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@STARTDATE", STARTDATE));
            command.Parameters.Add(new SqlParameter("@STARTTIME", STARTTIME));
            command.Parameters.Add(new SqlParameter("@ENDDATE", ENDDATE));
            command.Parameters.Add(new SqlParameter("@ENDTIME", ENDTIME));
            command.Parameters.Add(new SqlParameter("@DEP", DEP));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));

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
        public void AddAUOGD41(string EmpID, string STARTDATE, string STARTTIME, string ENDDATE, string ENDTIME, string DOCTYPE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand(" UPDATE hr_main104 SET DOCTYPE=@DOCTYPE,STARTDATE=@STARTDATE,STARTTIME=@STARTTIME,ENDDATE=@ENDDATE,ENDTIME=@ENDTIME WHERE EmpID=@EmpID AND EMPDATE=@EMPDATE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EmpID", EmpID));
            command.Parameters.Add(new SqlParameter("@EMPDATE", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@STARTDATE", STARTDATE));
            command.Parameters.Add(new SqlParameter("@STARTTIME", STARTTIME));
            command.Parameters.Add(new SqlParameter("@ENDDATE", ENDDATE));
            command.Parameters.Add(new SqlParameter("@ENDTIME", ENDTIME));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));

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
        private void HR104_Load(object sender, EventArgs e)
        {

            this.hR_BUTableAdapter.Fill(this.hR.HR_BU);
            toolStripTextBox1.Text = GetMenu.Day();

            try
            {
                hR_Main104TableAdapter.Connection = globals.Connection;
                this.hR_Main104TableAdapter.Fill(this.hR.HR_Main104, toolStripTextBox1.Text);


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }


        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.hR_Main104TableAdapter.Fill(this.hR.HR_Main104, toolStripTextBox1.Text);


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelReport.DELETEFILE();

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                FileName = lsAppDir + "\\Excel\\HR\\出缺勤104.xls";
                System.Data.DataTable DT1 = GetSAPRevenue13();
                System.Data.DataTable DT2 = GetSAPRevenue14();
                string G1 = DT1.Rows[0][0].ToString();
                string G2 = "";
                string G3 = "";
                string G4 = "";
                try
                {
                    G2 = DT1.Rows[1][0].ToString();
                }
                catch
                {
                    G2 = "";
                }
                try
                {
                    G3 = DT1.Rows[2][0].ToString();
                }
                catch
                {
                    G3 = "";
                }
                try
                {
                    G4 = DT1.Rows[3][0].ToString();
                }
                catch
                {
                    G4 = "";
                }
                string F2 = "";
                string F3 = "";
                string F4 = "";
                string F1 = DT2.Rows[0][0].ToString();
                try
                {
                    F2 = DT2.Rows[1][0].ToString();
                }
                catch
                {
                    F2 = "";
                }
                try
                {
                    F3 = DT2.Rows[2][0].ToString();
                }
                catch
                {
                    F3 = "";
                }
                try
                {
                    F4 = DT2.Rows[3][0].ToString();
                }
                catch
                {
                    F4 = "";
                }

                System.Data.DataTable OrderData = ExecuteQuery11(G1, F1, G2, F2, G3, F3, G4, F4);


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                     DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report


                ExcelReport.ExcelReportOutputHR104(OrderData, ExcelTemplate, OutPutFile, toolStripTextBox1.Text);


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            //GetExcelProduct(OutPutFile1);
       
        }
        private void GetExcelProduct(string ExcelFile)
        {

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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            int 資產 = 0;
            int 負債 = 0;
            int 業主權益 = 0;




            Microsoft.Office.Interop.Excel.Range range = null;

            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, 4]);
            range.Select();
            //資產 = Convert.ToInt32(range.Value2);

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                //20080509
                bool IsDetail = false;
                int DetailRow = 0;

                int Line_Liab = 0;

                for (int iRecord = iRowCnt; iRecord >= 1; iRecord--)
                {


                    //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    range.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    if (sTemp == "進金生" || sTemp == "博豐" || sTemp == "聿豐")
                    {

                        range.Select();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);


                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);

                    }
             

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    int g = sTemp.IndexOf("出勤率");

                    if (g != -1)
                    {

                        range.Select();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);


                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);

                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    if (sTemp != "")
                    {
                        range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    int g2 = sTemp.IndexOf("%");

                    if (g2 != -1)
                    {

                        range.Select();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);


                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);

                    }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    if (sTemp != "")
                    {
                        range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    int iPos2 = sTemp.IndexOf("班");
                    if (iPos2 < 0)
                    {
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    if (sTemp == "acme")
                    {

                        range.Select();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
                        range.Value2 = "";
                    }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    if (sTemp != "")
                    {
                        range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    int iPos1 = sTemp.IndexOf("人");
                    int iPos3 = sTemp.IndexOf("註");
                    if (iPos3 < 0)
                    {
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }

                    if (iPos1 > 0)
                    {

                        range.Select();

                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);

                    }

                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                range.Select();

                ////插入三行
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1]);
                range.Select();
                range.Value2 = toolStripTextBox1.Text + "出勤異常人員";
                object SelectCell_From = "A1";
                object SelectCell_To = "E1";
                range = excelSheet.get_Range(SelectCell_From, SelectCell_To);
                range.Select();
                range.Merge(true);
                range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                range.Font.Size = 14;
                range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                range.Font.Name = "Times New Roman";
            }
            finally
            {
                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
            "每日出勤" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";


                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                System.Diagnostics.Process.Start(NewFileName);


            }

        }
        private System.Data.DataTable GetSAPRevenue13()
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區96")
            {
                sb.Append("    SELECT count(*) FROM vwZZ_EMPLOYEE where EMPLOYEE_WORK_STATUS=1 and COMPANY_CNAME ='進金生實業股份有限公司'");
                sb.Append("       UNION ALL");
                sb.Append("    SELECT count(*) FROM vwZZ_EMPLOYEE where EMPLOYEE_WORK_STATUS=1 and COMPANY_CNAME ='進金生能源服務股份有限公司'");
                sb.Append("       UNION ALL");
                sb.Append("    SELECT count(*) FROM vwZZ_EMPLOYEE where EMPLOYEE_WORK_STATUS=1 and COMPANY_CNAME ='聿豐實業股份有限公司'");
                sb.Append("             UNION ALL");
                sb.Append("    SELECT count(*) FROM vwZZ_EMPLOYEE where EMPLOYEE_WORK_STATUS=1 and COMPANY_CNAME ='博豐光電股份有限公司'");
            }
            if (globals.DBNAME == "達睿生")
            {
                sb.Append("    SELECT count(*) FROM vwZZ_EMPLOYEE where EMPLOYEE_WORK_STATUS=1 and COMPANY_CNAME ='達睿生科技發展深圳有限公司'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable ExecuteQuery11(string ACMEEMP, string ACMENO, string SOLAREMP, string SOLARNO, string ARMASEMP, string ARMASNO, string ARESEMP, string ARESNO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區96")
            {

                if (GetSGROUP("進金生實業股份有限公司").Rows.Count > 0)
                {
                    sb.Append(" SELECT TOP 1 '進金生（公司別）' CHNAME,'' ENGNAME,'總員工數: '+cast(@A1 as varchar) BU ,'異常出勤人數:' CLOCKTIME,DOCTYPE=@A2,'正常出勤人數：' 開始日期,開始時間=CAST(CAST(@A1 AS INT) - CAST(@A2 AS INT) AS VARCHAR),'出勤率：' 結束日期,結束時間=case when @A1='0' then '100%' else cast(cast(round((cast(@A1 as decimal(3,0))-cast(@A2 as decimal(3,0)))/cast(@A1 as decimal(3,0)),2)*100 as int) as nvarchar)+'%' end FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT TOP 1 '姓名' CHNAME,'Portal帳號' ENGNAME,'部門名稱' BU,'當日卡鐘資料' CLOCKTIME,'假勤項目' DOCTYPE,'開始日期' 開始日期,'開始時間' 開始時間,'結束日期' 結束日期,'結束時間' 結束時間  FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    System.Data.DataTable H2 = GetGROUP5("進金生實業股份有限公司");
                    for (int i = 0; i <= H2.Rows.Count -1; i++)
                    {
                        string BU = H2.Rows[i]["BU"].ToString();
                        string BUCODE = H2.Rows[i]["BUCODE"].ToString();

                        if (GetSGROUP2("進金生實業股份有限公司", BUCODE).Rows.Count > 0)
                        {
                            sb.Append(" select TOP 1  CHNAME='" + BU + "','','','','','','','',''  from hr_main104  ");
                            sb.Append(" UNION ALL");
                            sb.Append("  SELECT CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE 開始日期,STARTTIME 開始時間,ENDDATE 結束日期,ENDTIME  結束時間 FROM HR_MAIN104 where company='進金生實業股份有限公司' AND EMPDATE=@EMPDATE AND DEP='" + BUCODE + "' ");
                            sb.Append(" UNION ALL");
                        }

                    }
                }
                if (GetSGROUP("進金生能源服務股份有限公司").Rows.Count > 0)
                {

                    sb.Append(" SELECT TOP 1 '太陽能（公司別）' CHNAME,'' ENGNAME,'總員工數: '+cast(@B1 as varchar) BU ,'異常出勤人數:' CLOCKTIME,DOCTYPE=@B2,'正常出勤人數：' 開始日期,開始時間=CAST(CAST(@B1 AS INT) - CAST(@B2 AS INT) AS VARCHAR),'出勤率：' 結束日期,結束時間=case when @B1='0' then '100%' else cast(cast(round((cast(@B1 as decimal(3,0))-cast(@B2 as decimal(3,0)))/cast(@B1 as decimal(3,0)),2)*100 as int) as nvarchar)+'%' end FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT TOP 1 '姓名' CHNAME,'Portal帳號' ENGNAME,'部門名稱' BU,'當日卡鐘資料' CLOCKTIME,'假勤項目' DOCTYPE,'開始日期' 開始日期,'開始時間' 開始時間,'結束日期' 結束日期,'結束時間' 結束時間  FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    System.Data.DataTable H2 = GetGROUP5("進金生能源服務股份有限公司");
                    for (int i = 0; i <= H2.Rows.Count - 1; i++)
                    {
                        string BU = H2.Rows[i]["BU"].ToString();
                        string BUCODE = H2.Rows[i]["BUCODE"].ToString();
                        System.Data.DataTable G1 = GetSGROUP2("進金生能源服務股份有限公司", BUCODE);
                        if (G1.Rows.Count > 0)
                        {
                            sb.Append(" select TOP 1  CHNAME='" + BU + "','','','','','','','',''  from hr_main104  ");
                            sb.Append(" UNION ALL");
                            sb.Append("  SELECT CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE 開始日期,STARTTIME 開始時間,ENDDATE 結束日期,ENDTIME  結束時間 FROM HR_MAIN104 where company='進金生能源服務股份有限公司' AND EMPDATE=@EMPDATE AND DEP='" + BUCODE + "' ");
                                    sb.Append(" UNION ALL");                     
                        }

                    }
                }
                if (GetSGROUP("聿豐實業股份有限公司").Rows.Count > 0)
                {
                    sb.Append(" SELECT TOP 1 '聿豐（公司別）' CHNAME,'' ENGNAME,'總員工數: '+cast(@C1 as varchar) BU  ,'異常出勤人數:' CLOCKTIME,DOCTYPE=@C2,'正常出勤人數：' 開始日期,開始時間=CAST(CAST(@C1 AS INT) - CAST(@C2 AS INT) AS VARCHAR),'出勤率：' 結束日期,結束時間=case when @C1='0' then '100%' else cast(cast(round((cast(@C1 as decimal(3,0))-cast(@C2 as decimal(3,0)))/cast(@C1 as decimal(3,0)),2)*100 as int) as nvarchar)+'%' end FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT TOP 1 '姓名' CHNAME,'Portal帳號' ENGNAME,'部門名稱' BU,'當日卡鐘資料' CLOCKTIME,'假勤項目' DOCTYPE,'開始日期' 開始日期,'開始時間' 開始時間,'結束日期' 結束日期,'結束時間' 結束時間  FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    System.Data.DataTable H2 = GetGROUP5("聿豐實業股份有限公司");
                    for (int i = 0; i <= H2.Rows.Count - 1; i++)
                    {
                        string BU = H2.Rows[i]["BU"].ToString();
                        string BUCODE = H2.Rows[i]["BUCODE"].ToString();
                        System.Data.DataTable G1 = GetSGROUP2("聿豐實業股份有限公司", BUCODE);
                        if (G1.Rows.Count > 0)
                        {
                            sb.Append(" select TOP 1  CHNAME='" + BU + "','','','','','','','',''  from hr_main104  ");
                            sb.Append(" UNION ALL");
                            sb.Append("  SELECT CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE 開始日期,STARTTIME 開始時間,ENDDATE 結束日期,ENDTIME  結束時間 FROM HR_MAIN104 where company='聿豐實業股份有限公司' AND EMPDATE=@EMPDATE AND DEP='" + BUCODE + "' ");
              
                       
                                    sb.Append(" UNION ALL");
                                
                   
                        }

                    }
                }
                if (GetSGROUP("博豐光電股份有限公司").Rows.Count > 0)
                {
        
                    sb.Append(" SELECT TOP 1 '博豐（公司別）' CHNAME,'' ENGNAME,'總員工數: '+cast(@D1 as varchar) BU  ,'異常出勤人數:' CLOCKTIME,DOCTYPE=@D2,'正常出勤人數：' 開始日期,開始時間=CAST(CAST(@D1 AS INT) - CAST(@D2 AS INT) AS VARCHAR),'出勤率：' 結束日期,結束時間=case when @D1='0' then '100%' else cast(cast(round((cast(@D1 as decimal(3,0))-cast(@D2 as decimal(3,0)))/cast(@D1 as decimal(3,0)),2)*100 as int) as nvarchar)+'%' end FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT TOP 1 '姓名' CHNAME,'Portal帳號' ENGNAME,'部門名稱' BU,'當日卡鐘資料' CLOCKTIME,'假勤項目' DOCTYPE,'開始日期' 開始日期,'開始時間' 開始時間,'結束日期' 結束日期,'結束時間' 結束時間  FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE 開始日期,STARTTIME 開始時間,ENDDATE 結束日期,ENDTIME  結束時間 FROM HR_MAIN104 where company='博豐光電股份有限公司' AND EMPDATE=@EMPDATE ");
                }
            }
            
            if (globals.DBNAME == "達睿生")
            {
                if (GetSGROUP("達睿生科技發展深圳有限公司").Rows.Count > 0)
                {

                    sb.Append(" SELECT TOP 1 '達睿生（公司別）' CHNAME,'' ENGNAME,'總員工數: '+cast(@A1 as varchar) BU  ,'異常出勤人數:' CLOCKTIME,DOCTYPE=@A2,'正常出勤人數：' 開始日期,開始時間=CAST(CAST(@A1 AS INT) - CAST(@A2 AS INT) AS VARCHAR),'出勤率：' 結束日期,結束時間=case when @A1='0' then '100%' else cast(cast(round((cast(@A1 as decimal(3,0))-cast(@A2 as decimal(3,0)))/cast(@A1 as decimal(3,0)),2)*100 as int) as nvarchar)+'%' end FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append(" SELECT TOP 1 '姓名' CHNAME,'Portal帳號' ENGNAME,'部門名稱' BU,'當日卡鐘資料' CLOCKTIME,'假勤項目' DOCTYPE,'開始日期' 開始日期,'開始時間' 開始時間,'結束日期' 結束日期,'結束時間' 結束時間  FROM HR_MAIN104 ");
                    sb.Append(" UNION ALL");
                    sb.Append("  SELECT CHNAME,ENGNAME,BU,CLOCKTIME,DOCTYPE,STARTDATE 開始日期,STARTTIME 開始時間,ENDDATE 結束日期,ENDTIME  結束時間 FROM HR_MAIN104 where company='達睿生科技發展深圳有限公司' AND EMPDATE=@EMPDATE ");
                }
            }
            string H1 = sb.ToString();
            int R = H1.Length;
            string ALL = H1.Substring(R - 3, 3);
            if (ALL == "ALL")
            {
                H1 = H1.Substring(0, R - 9);
            }
            SqlCommand command = new SqlCommand(H1, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", ACMEEMP));
            command.Parameters.Add(new SqlParameter("@A2", ACMENO));
            command.Parameters.Add(new SqlParameter("@B1", SOLAREMP));
            command.Parameters.Add(new SqlParameter("@B2", SOLARNO));
            command.Parameters.Add(new SqlParameter("@C1", ARMASEMP));
            command.Parameters.Add(new SqlParameter("@C2", ARMASNO));
            command.Parameters.Add(new SqlParameter("@D1", ARESEMP));
            command.Parameters.Add(new SqlParameter("@D2", ARESNO));
            command.Parameters.Add(new SqlParameter("@EMPDATE", toolStripTextBox1.Text));

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

        private System.Data.DataTable GetSAPRevenue14()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區96")
            {
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='進金生實業股份有限公司'");
                sb.Append(" union all");
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='進金生能源服務股份有限公司'");
                sb.Append(" union all");
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='聿豐實業股份有限公司'");
                sb.Append(" union all");
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='博豐光電股份有限公司'");
            }
            if (globals.DBNAME == "達睿生")
            {
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='達睿生科技發展深圳有限公司'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@empdate", toolStripTextBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }

        private System.Data.DataTable GetSGROUP(string COMPANY)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區96")
            {
                sb.Append(" SELECT * FROM HR_MAIN104 WHERE EMPDATE=@EMPDATE AND COMPANY=@COMPANY ");
  
            }
            if (globals.DBNAME == "達睿生")
            {
                sb.Append(" SELECT count(*) FROM HR_MAIN104 where empdate=@empdate and company='達睿生科技發展深圳有限公司'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPDATE", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSGROUP2(string COMPANY, string DEP)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM HR_MAIN104 WHERE EMPDATE=@EMPDATE AND COMPANY=@COMPANY  AND DEP=@DEP");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPDATE", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@DEP", DEP));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }


        private System.Data.DataTable GetGROUP4()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區96")
            {
                sb.Append(" SELECT * FROM HR_MAIN104 WHERE EMPDATE=@EMPDATE and company <> '達睿生科技發展深圳有限公司' ");

            }
            if (globals.DBNAME == "達睿生")
            {
                sb.Append(" SELECT * FROM HR_MAIN104 WHERE EMPDATE=@EMPDATE and company='達睿生科技發展深圳有限公司'  ");
        
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMPDATE", toolStripTextBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }

        private System.Data.DataTable GetGROUP5(string COMPANY)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
           
                sb.Append(" SELECT '部門'+BU BU,BUCODE FROM HR_BU WHERE COMPANY=@COMPANY ORDER BY LINE ");

            

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }


        private void button1_Click(object sender, EventArgs e)
        {
            int CC = 0;
            for (int i = 0; i <= hR_Main104DataGridView.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = hR_Main104DataGridView.Rows[i];
                string EmpID = row.Cells["EmpID"].Value.ToString();
                string DOCTYPE = row.Cells["DOCTYPE"].Value.ToString();
                System.Data.DataTable K1 = GetCHO6(EmpID);
                         string DOCTYPE2  = K1.Rows[0]["DOCTYPE"].ToString();
                         string STARTDATE = K1.Rows[0]["開始日期"].ToString();
                         string STARTTIME = K1.Rows[0]["開始時間"].ToString();
                         string ENDDATE = K1.Rows[0]["結束日期"].ToString();
                         string ENDTIME = K1.Rows[0]["結束時間"].ToString();
              
                if (DOCTYPE2 != DOCTYPE)
                {
                    CC = 1;
                    AddAUOGD41(EmpID, STARTDATE, STARTTIME, ENDDATE, ENDTIME, DOCTYPE2);
                   
                }

            }

            if (CC == 0)
            {
                MessageBox.Show("沒有假別被更新");
            }
            

            try
            {
                this.hR_Main104TableAdapter.Fill(this.hR.HR_Main104, toolStripTextBox1.Text);


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void 儲存SToolStripButton_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.hR_BUBindingSource.EndEdit();
            this.hR_BUTableAdapter.Update(this.hR);

            MessageBox.Show("更新成功");
        }

  
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//Excel
using Microsoft.Office.Interop.Excel;
//HashTable
using System.Collections;
using System.IO;

//
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ACME
{
    public partial class fmAcmeOJDT : Form
    {

        private string ConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database={0}";
        private string SAP;

        int ErrCount = 0;

        public static SAPbobsCOM.Company oCompany;

        public static string sErrMsg;
        public static int lErrCode;
        public static int lRetCode;

        System.Data.DataTable dt;

        private SAPbobsCOM.Recordset oRecordSet;

        public fmAcmeOJDT()
        {
            InitializeComponent();
        }

        private void GetCompany()
        {
            
            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;

            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";


            int i = 0; 
            oRecordSet = oCompany.GetCompanyList();


            oCompany.GetLastError(out lErrCode, out sErrMsg);

            //if (lErrCode != 0)
            //{
            //    MessageBox.Show(sErrMsg);
            //}
            //else
            //{
                
            //    while (!(oRecordSet.EoF == true))
            //    {
            //        // add the value of the first field of the Recordset
            //        Combo1.Items.Add(oRecordSet.Fields.Item(0).Value);
            //        // move the record pointer to the next row
            //        oRecordSet.MoveNext();
            //    }
            //}


            if (oCompany.Connected == true)
            {
                
                Command1.Enabled = false;
                Combo1.SelectedText = oCompany.CompanyDB;
                Text1.Text = oCompany.UserName;
                Text2.Text = oCompany.Password;
                this.Text = this.Text + ": Connected";
            }


            //MessageBox.Show("Connected....");
        }


        private void ConnetToCompany()
        {

            // setting the rest of the mandatory properties

            this.Cursor = Cursors.WaitCursor;

            try
            {


                oCompany.CompanyDB = Combo1.Text;
                oCompany.UserName = Text1.Text;
                oCompany.Password = Text2.Text;

                // Connecting to a company DB
                lRetCode = oCompany.Connect();

                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lErrCode, out sErrMsg);
                    MessageBox.Show(sErrMsg);
                }
                else
                {
                    //MessageBox.Show("Connected to " + oCompany.CompanyName);
                    this.Text = oCompany.CompanyName + ": Connected";
                    Command1.Enabled = false;
                    Command2.Enabled = true;

                    Combo1.Enabled = false;
                }
            }
            finally
            {
                this.Cursor = Cursors.Default;
            
            }

        }

        private void fmAcmeOJDT_Load(object sender, EventArgs e)
        {
            GetCompany();
            Combo1.Text = "AcmeSql05";

            
            Command2.Enabled = false;


        }

        private void Command1_Click(object sender, EventArgs e)
        {
            ConnetToCompany();

            SAP = string.Format(ConnStr, Combo1.Text);
         
            //if (MessageBox.Show("確定執行嗎？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            //{
            //    return;
            //}
        }

        private void Command2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("確定執行嗎？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }


            if (Combo1.Text == "AcmeSql02")
            {

                if (MessageBox.Show("這個是正式區----確定執行嗎？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }
            
            }

            


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                this.UseWaitCursor = true;
                
                string FileName = openFileDialog1.FileName;
                //MessageBox.Show(FileName);
                GetExcelContent(FileName);



                CheckDt(dt);

                if (ErrCount > 0)
                {
                    MessageBox.Show("資料有錯誤,請修正 !");
                    this.UseWaitCursor = false;
                    return;
                }




                DtToSap(dt);


                this.UseWaitCursor = false;

                MessageBox.Show("匯入資料完成");

            }

        }


        private void AppendToError(string Msg)
        {
            txtError.Text += Msg +"\r\n";
        }

        private void AppendToSuccess(string Msg)
        {
            txtSuccess.Text += Msg + "\r\n";
        }


        private void CheckDt(System.Data.DataTable dt)
        {

            txtError.Text = "";
            txtSuccess.Text = "";
            
            string Group = "";

            string DocNo="";
            DateTime DocDate;
            string AccountCode;
            Int32 Credit;
            Int32 Debit;
            string Memo;

            string ProjectCode;
            string CostingCode;

            ErrCount = 0;
         

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {


                DocNo = Convert.ToString(dt.Rows[i]["DocNo"]);
                DocDate = Convert.ToDateTime(dt.Rows[i]["DocDate"]);

                AccountCode = Convert.ToString(dt.Rows[i]["AccountCode"]);

                Credit = Convert.ToInt32(dt.Rows[i]["Credit"]);
                Debit = Convert.ToInt32(dt.Rows[i]["Debit"]);
                Memo = Convert.ToString(dt.Rows[i]["Memo"]);


                ProjectCode = Convert.ToString(dt.Rows[i]["ProjectCode"]);

                CostingCode = Convert.ToString(dt.Rows[i]["CostingCode"]);


                if (string.IsNullOrEmpty(AccountCode))
                {
                    AppendToError(string.Format("單號:{0} 科目未輸入", DocNo));
                    ErrCount++;
                }

                if (Credit > 0 && Debit > 0)
                {
                    AppendToError(string.Format("單號:{0} 借貸方金額錯誤", DocNo));
                    ErrCount++;
                }


                if (Credit == 0 && Debit == 0)
                {
                    AppendToError(string.Format("單號:{0} 借貸方金額錯誤", DocNo));
                    ErrCount++;
                }


                if (AccountCode.Length < 8)
                {
                    if (CheckBP(AccountCode) != 1)
                    {
                        AppendToError(string.Format("單號:{0} 業務夥伴:{1} 不存在", DocNo, AccountCode));
                        ErrCount++;
                    }

                }
                else
                {
                    if (CheckAcctCode(AccountCode) != 1)
                    {
                        AppendToError(string.Format("單號:{0} 科目:{1} 不存在", DocNo, AccountCode));
                        ErrCount++;
                    }
                }

                if (!string.IsNullOrEmpty(ProjectCode))
                {

                    if (CheckProject(ProjectCode) != 1)
                    {
                        AppendToError(string.Format("單號:{0} 專案:{1} 不存在", DocNo, ProjectCode));
                        ErrCount++;
                    }
                    
                    
                   
                }

                if (!string.IsNullOrEmpty(CostingCode))
                {

                    if (CheckDept(CostingCode) != 1)
                    {
                        AppendToError(string.Format("單號:{0} 部門:{1} 不存在", DocNo, CostingCode));
                        ErrCount++;
                    }



                }

                

            }

        }

        private void DtToSap(System.Data.DataTable dt)
        {

            string Group = "";

            string DocNo="";
            DateTime DocDate;
            string AccountCode;
            Int32 Credit;
            Int32 Debit;
            string Memo;

            string ProjectCode;
            string CostingCode;



            //傳票分錄
            SAPbobsCOM.JournalEntries oJE=null;

            for (int i =0;i <=dt.Rows.Count -1;i++)
            {


                DocNo = Convert.ToString(dt.Rows[i]["DocNo"]);
                DocDate = Convert.ToDateTime(dt.Rows[i]["DocDate"]);

                AccountCode = Convert.ToString(dt.Rows[i]["AccountCode"]);

                Credit = Convert.ToInt32(dt.Rows[i]["Credit"]);

                Debit = Convert.ToInt32(dt.Rows[i]["Debit"]);
                Memo = Convert.ToString(dt.Rows[i]["Memo"]);


                ProjectCode = Convert.ToString(dt.Rows[i]["ProjectCode"]);

                CostingCode = Convert.ToString(dt.Rows[i]["CostingCode"]);

                if (Group != DocNo)
                {
                    


                    if (i != 0)
                    {

                        try
                        {
                            int lRetCode = oJE.Add();
                            if (lRetCode != 0)
                            {
                                oCompany.GetLastError(out lErrCode, out sErrMsg);
                                //MessageBox.Show(lErrCode + " " + sErrMsg);

                                string Msg = string.Format("單號{0} 錯誤訊息:{1}", Group, lErrCode + " " + sErrMsg);
                                AppendToError(Msg);
                            }
                            else
                            {
                                AppendToSuccess(string.Format("單號{0} ->OK", Group));
                            }
                        }
                        catch
                        {

                        }


                    }

                    //Block by row 要解除
                    //科目的稅要定義


                    oJE = (SAPbobsCOM.JournalEntries)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));

                    //到期日期
                    oJE.DueDate = DocDate;
                    //過帳日期
                    oJE.ReferenceDate = DocDate;
                    //文件日期
                   // oJE.TaxDate = DocDate;

                    oJE.Series = 14;


                    //如果是 BP 則自動代入所對應的應收帳款

                    if (AccountCode.Length < 8)
                    {
                        oJE.Lines.ShortName = AccountCode;
                    }
                    else
                    {
                        oJE.Lines.AccountCode = AccountCode;
                    }
                    //沖銷科目
                    //oJE.Lines.ContraAccount = "10090101";
                    oJE.Lines.Credit = Credit;
                    oJE.Lines.Debit = Debit;
                    oJE.Lines.LineMemo = Memo;

                    oJE.Lines.ProjectCode = ProjectCode;
                    oJE.Lines.CostingCode = CostingCode;


                 
                    oJE.Lines.Add();
                    Group = DocNo;

            //        'Second line
            //entries.Lines.SetCurrentLine(1)

                }
                else
                {

                    if (AccountCode.Length < 8)
                    {
                        oJE.Lines.ShortName = AccountCode;
                    }
                    else
                    {
                        oJE.Lines.AccountCode = AccountCode;
                    }
               
                    oJE.Lines.Credit = Credit;
                    oJE.Lines.Debit = Debit;

                    oJE.Lines.LineMemo = Memo;

                    oJE.Lines.ProjectCode = ProjectCode;
                    oJE.Lines.CostingCode = CostingCode;

                    //oJE.Lines.TaxGroup = "AR0%";

                    //oJE.Lines.DueDate = DateTime.Now;
                    //oJE.Lines.ReferenceDate1 = DateTime.Now;
                   // oJE.Lines.ShortName = "10090101";
                    //oJE.Lines.TaxDate = DateTime.Now;
                    //oJE.Lines.TaxGroup = "";
                    //oJE.Lines.vat

                    oJE.Lines.Add();
                }





            }

            try
            {
                int lRetCode = oJE.Add();
                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lErrCode, out sErrMsg);
                   // MessageBox.Show(lErrCode + " " + sErrMsg);
                    string Msg = string.Format("單號{0} 錯誤訊息:{1}", Group, lErrCode + " " + sErrMsg);
                    AppendToError(Msg);
                }
                else
                {
                    AppendToSuccess(string.Format("單號{0} ->OK", Group));
                }
            }
            catch
            {

            }

            
        }

        private void GetExcelContent(string ExcelFile)
        {

            dt = MakeTable();

      

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
 

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;




            Hashtable ht = new Hashtable(iRowCnt);

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;


            object SelectCell = "C7";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string DocNo;
            string DocDate;
            string AccountCode;
            string Credit;
            string Debit;
            string Memo;

            string ProjectCode;
            string CostingCode;


        

            //for (int i = 10; i <= iRowCnt; i++)
            for (int i = 2; i <= iRowCnt; i++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                DocNo = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();

                DocDate = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                AccountCode = range.Text.ToString();

               

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                Debit = range.Text.ToString();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                Credit = range.Text.ToString();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                Memo = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                CostingCode = range.Text.ToString();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                range.Select();
                ProjectCode = range.Text.ToString();



                ht.Add(i.ToString(), strText);

                if (string.IsNullOrEmpty(DocNo)) continue;


                try
                {
                    DataRow dr = dt.NewRow();
                    dr["DocNo"] = DocNo;
                    dr["DocDate"] = DocDate;


                    dr["AccountCode"] = AccountCode;

                    if (string.IsNullOrEmpty(Credit)) Credit="0";
                    if (string.IsNullOrEmpty(Debit)) Debit = "0";


                    dr["Credit"] = Credit;
                    dr["Debit"] = Debit;
                    dr["Memo"] = Memo;

                    dr["CostingCode"] = CostingCode;

                    dr["ProjectCode"] = ProjectCode;

                    dt.Rows.Add(dr);
                }
                catch
                { 
                
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

           // dataGridView1.DataSource = dt;
        }


        //動態產生資料結構
        private System.Data.DataTable MakeTable()
        {

            //string DocNo;
            //string DocDate;
            //string AccountCode;
            //string Credit;
            //string Debit;
            //string Memo;

            //string ProjectCode;
            //string CostingCode;


            System.Data.DataTable dt = new System.Data.DataTable();
            ///旗標,產品編號,日期,數量
            //第一個固定欄位
            dt.Columns.Add("DocNo", typeof(string));
            dt.Columns.Add("DocDate", typeof(string));
            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("Credit", typeof(Int32));
            dt.Columns.Add("Debit", typeof(Int32));
            dt.Columns.Add("Memo", typeof(string));

            dt.Columns.Add("ProjectCode", typeof(string));
            dt.Columns.Add("CostingCode", typeof(string));

            /*
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["SERIAL_NO"];
            dt.PrimaryKey = colPk;
            */

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }



        private DateTime StrToDate1(string sDate)
        {

            string[] s = sDate.Split('/');


            UInt16 Year = Convert.ToUInt16(s[0]);
            UInt16 Month = Convert.ToUInt16(s[1]);
            UInt16 Day = Convert.ToUInt16(s[2]);

            return new DateTime(Year, Month, Day);
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
        }



        public  int CheckAcctCode(string AcctCode)
        {
            SqlConnection connection = new SqlConnection(SAP);
            string sql = "SELECT COUNT(*) FROM OACT WHERE AcctCode=@AcctCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AcctCode", AcctCode));
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }

        public int CheckBP(string CardCode)
        {
            SqlConnection connection = new SqlConnection(SAP);
            string sql = "SELECT COUNT(*) FROM OCRD WHERE CardCode=@CardCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }

        public int CheckProject(string PrjCode)
        {
            SqlConnection connection = new SqlConnection(SAP);
            string sql = "SELECT COUNT(*) FROM OPRJ WHERE PrjCode=@PrjCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }


        public int CheckDept(string OcrCode)
        {
            SqlConnection connection = new SqlConnection(SAP);
            string sql = "SELECT COUNT(*) FROM OOCR WHERE OcrCode=@OcrCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OcrCode", OcrCode));
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }

    }
}
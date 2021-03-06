using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;

//HashTable
using System.Collections;
using System.IO;
//
using System.Data.OleDb;

namespace ACME
{
    
    
    public partial class APSAP : Form
    {
        private string FileName;

        private static string EEPConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=Acmesqlsp";

        private string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
        private SAPbobsCOM.Recordset oRecordSet;


        private System.Data.DataTable dt;


        public APSAP()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        public void  LOAD()
        {

            globals.oCompany = new SAPbobsCOM.Company();

            globals.oCompany.Server = "acmesap";
            globals.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            globals.oCompany.UseTrusted = false;
            globals.oCompany.DbUserName = "sapdbo";
            globals.oCompany.DbPassword = "@rmas";


            int i = 0; //  to be used as an index

            oRecordSet = globals.oCompany.GetCompanyList();

            globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);


            globals.oCompany.CompanyDB = "acmesql02";
            globals.oCompany.UserName = "manager";
            globals.oCompany.Password = "19571215";

            // Connecting to a company DB
            globals.lRetCode = globals.oCompany.Connect();

            if (globals.lRetCode != 0)
            {
                globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                Interaction.MsgBox(globals.sErrMsg, (Microsoft.VisualBasic.MsgBoxStyle)(0), null);
            }
            else
            {
                Interaction.MsgBox("連結公司 " + globals.oCompany.CompanyName, (Microsoft.VisualBasic.MsgBoxStyle)(0), null);
                this.Text = this.Text + ": Connected";
        
            }  

        }
        public void DIADDSAP(string DOCENTRY, string ITEMCODE, string ITEMNAME, double QTY, double PRICE, int SLPCODE)
        {


            SAPbobsCOM.Documents vPO = (SAPbobsCOM.Documents)(globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders));


            bool RetVal;

            //找出那一張採購單
            RetVal = vPO.GetByKey(Convert.ToInt16(DOCENTRY));

            if (RetVal == false)
            {

                globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                MessageBox.Show(globals.lErrCode + " " + globals.sErrMsg);
                return;
            }
            SAPbobsCOM.Items oItem = null;

            oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;

           // int OldLineNum = Convert.ToInt16(textBox2.Text);



            //  vPO.Lines.SetCurrentLine(OldLineNum);

            //   MessageBox.Show(vPO.Lines.ItemCode);
            vPO.Lines.Add();
            vPO.Lines.ItemCode = ITEMCODE;
            vPO.Lines.ItemDescription = ITEMNAME;
            vPO.Lines.Quantity = QTY;
            vPO.Lines.UnitPrice  = PRICE;
            vPO.Lines.SalesPersonCode = SLPCODE;

            // vPO.Lines.Add();


            int RetVal1 = vPO.Update();

            if (RetVal1 != 0)
            {
                globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                MessageBox.Show(globals.lErrCode + " " + globals.sErrMsg);
            }

        }

        public void DIUPSAP(string DOCENTRY,int OldLineNum,double QTY)
        {


            SAPbobsCOM.Documents vPO = (SAPbobsCOM.Documents)(globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders));


            bool RetVal;

            //找出那一張採購單
            RetVal = vPO.GetByKey(Convert.ToInt16(DOCENTRY));

            if (RetVal == false)
            {

                globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                MessageBox.Show(globals.lErrCode + " " + globals.sErrMsg);
                return;
            }


        



              vPO.Lines.SetCurrentLine(OldLineNum);


            vPO.Lines.Quantity = QTY;

            // vPO.Lines.Add();


            int RetVal1 = vPO.Update();

            if (RetVal1 != 0)
            {
                globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                MessageBox.Show(globals.lErrCode + " " + globals.sErrMsg);
            }

        }
        //日期處理--------------------------------------------------------------------------------------------
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }

        private void AddToData()
        {

            //M170EP01 ->產品編號增加 Model no 以避免產品編號找不到
            SAPbobsCOM.Documents vPO = (SAPbobsCOM.Documents)(globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders));

            int lRetCode;


            //自行下 SQL 取得 Key 值
            //if (oCard.GetByKey(SERIAL_NO))
            //{

            //   // RetVal = oCard.Update();

            //}
            //else
            //{

            //if (!CheckSAPSerial(SERIAL_NO))
            //{

            vPO.Lines.ItemCode = "TG150XG03.03002";
            vPO.Lines.ItemDescription = "系統編輯";
            //   vPO.Lines.TaxBeforeDPMFC
          lRetCode = vPO.Add();

                if (lRetCode != 0)
                {
                    globals.oCompany.GetLastError(out globals.lErrCode, out globals.sErrMsg);
                    MessageBox.Show(globals.lErrCode + " " + globals.sErrMsg);

                    //  listBox1.Items.Add(row.Cells["單號"].Value.ToString() + " " + globals.lErrCode + " " + globals.sErrMsg);

                }

            //}


        }



        private void APSAP_Load(object sender, EventArgs e)
        {
            LOAD();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string USER = fmLogin.LoginID.ToString().ToUpper();
                DELSAPOPOR(USER);
                FileName = openFileDialog1.FileName;
                GetExcelContentGD44(FileName);
                System.Data.DataTable T1 = GETSAPOPOR(USER);
                if (T1.Rows.Count > 0)
                {
                    for (int h = 0; h <= T1.Rows.Count - 1; h++)
                    {
                        string DOCENTRY = T1.Rows[h]["DOCENTRY"].ToString();
                        string MODEL = T1.Rows[h]["MODEL"].ToString();
                        double QTY = Convert.ToDouble(T1.Rows[h]["QTY"].ToString());
                        System.Data.DataTable G1 = GETOPOR(DOCENTRY, MODEL);
                        if (G1.Rows.Count == 1)
                        {
                            string ITEMCODEO = G1.Rows[0]["ITEMCODE"].ToString();
                            int LINENUM = Convert.ToInt16(G1.Rows[0]["LINENUM"]);
                            double QTY1 = Convert.ToDouble(G1.Rows[0]["QUANTITY"]);
                            double QTY2 = QTY1 - QTY;
                            int VISORDER = Convert.ToInt16(G1.Rows[0]["VISORDER"]) + 1;

                            AddSAPOPOR2(Convert.ToInt16(DOCENTRY), LINENUM, Convert.ToInt32(QTY1), Convert.ToInt32(QTY2), USER, ITEMCODEO, VISORDER);
                        }
                    }

                    dataGridView1.DataSource = GETSAPOPOR2D(USER);
                    dataGridView2.DataSource = GETSAPOPOR3D(USER);
                }


            }
        }

        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
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


            string DOCENTRY;
            string MODEL = "";
            string PARTNO = "";
            double  PRICE = 0;
            double QTY = 0;
            for (int i = 1; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                DOCENTRY = range.Text.ToString().Trim();
                int num1;

                if (int.TryParse(DOCENTRY, out num1) != false)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    PARTNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    PRICE = Convert.ToDouble(range.Text.ToString().Trim());

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 13]);
                    range.Select();
                    QTY = Convert.ToDouble(range.Text.ToString().Trim());
                    System.Data.DataTable G2 = GETPARTNO(PARTNO);
                    System.Data.DataTable G1 = GETOPOR(DOCENTRY, MODEL);
                    if (G1.Rows.Count == 1)
                    {
                        if (G2.Rows.Count > 0)
                        {
                            string ITEMCODEO = G1.Rows[0]["ITEMCODE"].ToString();
                            string ITEMCODE = G2.Rows[0]["ITEMCODE"].ToString();
                            string ITEMNAME = G2.Rows[0]["ITEMNAME"].ToString();
                            int SLPCODE = Convert.ToInt16(G1.Rows[0]["SLPCODE"]);
                            //SLPCODE
                            //    ADDSAP(DOCENTRY, ITEMCODE, ITEMNAME, QTY, PRICE);
                            AddSAPOPOR(Convert.ToInt16(DOCENTRY), MODEL, PARTNO, fmLogin.LoginID.ToString().ToUpper(), Convert.ToInt32(QTY), ITEMCODE, ITEMNAME, ITEMCODEO, PRICE, SLPCODE);

                        }
                    }

                }


            }



            //Quit
            excelBook.Close(0);
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

        public System.Data.DataTable GETOPOR(string DOCENTRY, string U_MODEL)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,LINENUM,QUANTITY,SLPCODE,VISORDER FROM ACMESQL02.DBO.POR1 T0 LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE) WHERE T0.DOCENTRY=@DOCENTRY AND REPLACE(U_MODEL,' V','') LIKE '%" + U_MODEL + "%'  AND T0.LINESTATUS='O' AND T0.GROSSBUYPR IS NOT NULL ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@U_MODEL", U_MODEL));
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

        public System.Data.DataTable GETSAPOPOR(string USERNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY) QTY,MODEL,DOCENTRY FROM AP_SAPOPOR WHERE USERNAME=@USERNAME GROUP BY MODEL,DOCENTRY");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
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
        public System.Data.DataTable GETSAPOPOR2(string USERNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT *  FROM AP_SAPOPOR2 where USERNAME=@USERNAME ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
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
        public System.Data.DataTable GETSAPOPOR2D(string USERNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY,VISORDER,ITEMCODEO,QTY1,QTY2  FROM AP_SAPOPOR2 where USERNAME=@USERNAME ORDER BY DOCENTRY,VISORDER ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
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
        public System.Data.DataTable GETSAPOPOR3(string USERNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT *  FROM AP_SAPOPOR where USERNAME=@USERNAME ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
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
        public System.Data.DataTable GETSAPOPOR3D(string USERNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY,ITEMCODEO,ITEMCODE,ITEMNAME,MODEL,PARTNO,QTY,PRICE  FROM AP_SAPOPOR where USERNAME=@USERNAME ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
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
        public System.Data.DataTable GETPARTNO(string U_PARTNO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,ITEMNAME FROM ACMESQL02.DBO.OITM  WHERE U_PARTNO=@U_PARTNO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
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
        public void AddSAPOPOR(int DOCENTRY, string MODEL, string PARTNO, string USERNAME, int QTY, string ITEMCODE, string ITEMNAME, string ITEMCODEO, double PRICE, int SLPCODE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_SAPOPOR(DOCENTRY,MODEL,PARTNO,USERNAME,QTY,ITEMCODE,ITEMNAME,ITEMCODEO,PRICE,SLPCODE) values(@DOCENTRY,@MODEL,@PARTNO,@USERNAME,@QTY,@ITEMCODE,@ITEMNAME,@ITEMCODEO,@PRICE,@SLPCODE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODEO", ITEMCODEO));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@SLPCODE", SLPCODE));
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
        public void AddSAPOPOR2(int DOCENTRY, int LINENUM, int QTY1, int QTY2, string USERNAME, string ITEMCODEO, int VISORDER)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_SAPOPOR2(DOCENTRY,LINENUM,QTY1,QTY2,USERNAME,ITEMCODEO,VISORDER) values(@DOCENTRY,@LINENUM,@QTY1,@QTY2,@USERNAME,@ITEMCODEO,@VISORDER)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@QTY1", QTY1));
            command.Parameters.Add(new SqlParameter("@QTY2", QTY2));
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODEO", ITEMCODEO));
            command.Parameters.Add(new SqlParameter("@VISORDER", VISORDER));
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
        public void DELSAPOPOR(string USERNAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_SAPOPOR WHERE USERNAME=@USERNAME DELETE AP_SAPOPOR2 WHERE USERNAME=@USERNAME ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));

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
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("項目號碼", typeof(string));
            dt.Columns.Add("LINE", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            return dt;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("是否確定要匯入SAP ?", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                string USER = fmLogin.LoginID.ToString().ToUpper();
                System.Data.DataTable H1 = GETSAPOPOR3(USER);
                if (H1.Rows.Count > 0)
                {
                    for (int h = 0; h <= H1.Rows.Count - 1; h++)
                    {

                        string DOCENTRY = H1.Rows[h]["DOCENTRY"].ToString();
                        string ITEMCODE = H1.Rows[h]["ITEMCODE"].ToString();
                        string ITEMNAME = H1.Rows[h]["ITEMNAME"].ToString();
                        double QTY = Convert.ToDouble(H1.Rows[h]["QTY"]);
                        double PRICE = Convert.ToDouble(H1.Rows[h]["PRICE"]);
                        int SLPCODE = Convert.ToInt16(H1.Rows[h]["SLPCODE"]);
                        DIADDSAP(DOCENTRY, ITEMCODE, ITEMNAME, QTY, PRICE, SLPCODE);
                    }

                }
                else
                {
                    MessageBox.Show("請重新匯入EXCEL");
                    return;
                }
                System.Data.DataTable H2 = GETSAPOPOR2(USER);
                for (int h = 0; h <= H2.Rows.Count - 1; h++)
                {
                    string DOCENTRY = H2.Rows[h]["DOCENTRY"].ToString();
                    int LINENUM = Convert.ToInt16(H2.Rows[h]["LINENUM"]);
                    double QTY2 = Convert.ToDouble(H2.Rows[h]["QTY2"]);
                    DIUPSAP(DOCENTRY, LINENUM, QTY2);
                }

                MessageBox.Show("匯入成功");
                DELSAPOPOR(USER);
            }
        }
        
    }

}


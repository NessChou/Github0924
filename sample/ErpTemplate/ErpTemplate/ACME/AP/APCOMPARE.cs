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
    public partial class APCOMPARE : Form
    {
        private string FileName;
        public string c;
        public string d;
        StringBuilder sb = new StringBuilder();
        StringBuilder sb2 = new StringBuilder();
        StringBuilder sb3 = new StringBuilder();
        public APCOMPARE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
               

                WriteExcelAP(FileName);
  
                MessageBox.Show("匯入完成");
            }
        }
        public void AddAP(string BU, string MODEL, string VER, string TYPE, string GRADE, string CDATE, string PRICE, string PPRICE, string REMARK, string CUST, string FILENAME, DateTime DATETIME)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = null;
            System.Data.DataTable G1 = GETC2M(MODEL, VER, TYPE, GRADE, CDATE, BU);
            if (G1.Rows.Count > 0)
            {
                command = new SqlCommand(" UPDATE  AP_COMPARE SET PRICE=@PRICE,REMARK=@REMARK,PPRICE=@PPRICE,FILENAME=@FILENAME,DATETIME=@DATETIME  WHERE BU=@BU AND MODEL=@MODEL AND VER=@VER AND  TYPE=@TYPE AND GRADE=@GRADE AND CDATE=@CDATE", connection);
            }
            else
            {
                command = new SqlCommand("Insert into AP_COMPARE(BU,MODEL,VER,TYPE,GRADE,CDATE,PRICE,PPRICE,REMARK,CUST,FILENAME,DATETIME) values(@BU,@MODEL,@VER,@TYPE,@GRADE,@CDATE,@PRICE,@PPRICE,@REMARK,@CUST,@FILENAME,@DATETIME)", connection);
            }
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@CDATE", CDATE));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@PPRICE", PPRICE));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@CUST", CUST));
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
            command.Parameters.Add(new SqlParameter("@DATETIME", DATETIME));
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
        public void UPAP(string BU, string MODEL, string VER, string TYPE, string GRADE, string CDATE, string PRICE, string PPRICE, string REMARK)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE  AP_COMPARE SET PRICE=@PRICE,REMARK=@REMARK,PPRICE=@PPRICE WHERE BU=@BU AND MODEL=@MODEL AND VER=@VER AND  TYPE=@TYPE AND GRADE=@GRADE AND CDATE=@CDATE", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@CDATE", CDATE));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@PPRICE", PPRICE));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
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

            int TTT1 = excelBook.Sheets.Count;


            Microsoft.Office.Interop.Excel.Range range = null;
         

            try
            {
             
                string BU = "";
                string MODEL = "";
                string VER = "";
                string TYPE = "";
                string GRADE = "";
                string CDATE = "";
                string PRICE = "";
                string PPRICE = "";
                string REMARK = "";
                string CUST = "";
                for (int s = 1; s <= TTT1; s++)
                {


                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(s);
                    excelSheet.Activate();
                    //取得 Excel 的使用區域
                    int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                    int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
     
                    string L1 = excelSheet.Name.Trim();
                    string NAME = excelBook.Name.Trim();
                    label7.Text = NAME;
                    if (L1.IndexOf("欄位明細") != -1)
                    {
                        progressBar1.Maximum = iRowCnt;
                        label6.Text = L1 + " 上傳進度";

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                        range.Select();
                        BU = range.Text.ToString().Trim().Replace("BU", "");
                        for (int iRecord = 3; iRecord <= iRowCnt; iRecord++)
                        {

                            label8.Text = iRecord.ToString() + " / " + iRowCnt.ToString();
                            progressBar1.Value = iRecord;
                            progressBar1.PerformStep();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                            range.Select();
                            MODEL = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            VER = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            range.Select();
                            TYPE = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            GRADE = range.Text.ToString().Trim().ToUpper();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                            range.Select();
                            CDATE = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            range.Select();
                            PRICE = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                            range.Select();
                            PPRICE = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                            range.Select();
                            REMARK = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                            range.Select();
                            CUST = range.Text.ToString().Trim();
                            string FR = range.MergeCells.ToString().ToUpper();

                            if (FR == "TRUE")
                            {
                                range = (Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[range.MergeArea.Row, range.MergeArea.Column];
                                range.Select();
                                string RR = range.Text.ToString().Trim();
                                int A1 = RR.IndexOf("與");
                                if (A1 != -1)
                                {
                                    int A2 = RR.IndexOf("同");
                                    if (A2 != -1)
                                    {
                                        string RR2 = RR.Replace("與", "").Replace("同", "");
                                        int A4 = RR2.IndexOf(".");
                                        string MODEL1 = RR2.Substring(0, A4);
                                        string VER1 = "V" + RR2.Substring(A4 + 1, 1);
                                        if (GRADE == "PN")
                                        {
                                            AddAP(BU, MODEL1, VER1, TYPE, "P", CDATE, PRICE, PPRICE, REMARK, CUST, NAME,DateTime.Now);
                                            AddAP(BU, MODEL1, VER1, TYPE, "N", CDATE, PRICE, PPRICE, REMARK, CUST, NAME,DateTime.Now);
                                        }
                                        else if (GRADE == "ZP")
                                        {
                                            AddAP(BU, MODEL1, VER1, TYPE, "Z", CDATE, PRICE, PPRICE, REMARK, CUST, NAME,DateTime.Now);
                                            AddAP(BU, MODEL1, VER1, TYPE, "P", CDATE, PRICE, PPRICE, REMARK, CUST, NAME,DateTime.Now);
                                        }
                                        else
                                        {
                                            AddAP(BU, MODEL1, VER1, TYPE, GRADE, CDATE, PRICE, PPRICE, REMARK, CUST, NAME,DateTime.Now);
                                        }
                                    }

                                    int A3 = RR.IndexOf("+");
                                    if (A3 != -1)
                                    {
                                        int T1 = Convert.ToInt32(PRICE);

                                        string RR2 = RR.Replace("與", "");
                                        int A4 = RR2.IndexOf(".");
                                        string MODEL1 = RR2.Substring(0, A4);
                                        string VER1 = "V" + RR2.Substring(A4 + 1, 1);
                                        int A5 = RR2.IndexOf("+");
                                        string F1 = RR2.Substring(A5 + 1, RR2.Length - 1 - A5);
                                        int I4 = Convert.ToInt32(F1);
                                        string PRICE1 = (T1 + I4).ToString();
                                        AddAP(BU, MODEL1, VER1, TYPE, GRADE, CDATE, PRICE1, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                                    }
                                }
                            }

                            if (GRADE == "PN")
                            {
                                AddAP(BU, MODEL, VER, TYPE, "P", CDATE, PRICE, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                                AddAP(BU, MODEL, VER, TYPE, "N", CDATE, PRICE, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                            }
                           else  if (GRADE == "ZP")
                            {
                                AddAP(BU, MODEL, VER, TYPE, "Z", CDATE, PRICE, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                                AddAP(BU, MODEL, VER, TYPE, "P", CDATE, PRICE, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                            }
                            else
                            {

                                AddAP(BU, MODEL, VER, TYPE, GRADE, CDATE, PRICE, PPRICE, REMARK, CUST, NAME, DateTime.Now);
                            }
                        }


                    }

                }





            }
            finally
            {

       
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



        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            sb.Clear();
            sb2.Clear();
            sb3.Clear();
            if (cGD.Checked)
            {
                sb.Append("'GD',");
            }

            if (cDT.Checked)
            {
                sb.Append("'DT',");
            }

            if (cPID.Checked)
            {
                sb.Append("'PID',");
            }

            if (cTV.Checked)
            {
                sb.Append("'TV',");
            }

            sb.Remove(sb.Length - 1, 1);


            if (cZ.Checked)
            {
                sb2.Append("'Z',");
            }

            if (cP.Checked)
            {
                sb2.Append("'P',");
            }

            if (cN.Checked)
            {
                sb2.Append("'N',");
            }


            sb2.Remove(sb2.Length - 1, 1);

            if (cMDL.Checked)
            {
                sb3.Append("'MDL',");
            }

            if (cOC.Checked)
            {
                sb3.Append("'OC',");
            }

            sb3.Remove(sb3.Length - 1, 1);

            System.Data.DataTable dtCost = MakeTable();


            System.Data.DataTable dt = GETC1(sb2.ToString());

            DataRow dr = null;
            string D = sb3.ToString().Replace("'", "");
            //ZPN
            string D2 = sb2.ToString().Replace("'", "");
            string D3 = sb.ToString().Replace("'", "");
            string[] GD = D3.Split(new Char[] { ',' });
            string[] ZPN = D2.Split(new Char[] { ',' });
            string[] MDL = D.Split(new Char[] { ',' });
            for (int i2 = 0; i2 <= dt.Rows.Count - 1; i2++)
            {

                foreach (string S in GD)
                {
                    //string DS = "MDL,OC";
                 //   string[] GDS = DS.Split(new Char[] { ',' });
                    foreach (string S2 in MDL)
                    {
                        for (int i3 = 1; i3 <= 2; i3++)
                        {

                            if (i3 == 1)
                            {
                                System.Data.DataTable dt2 = GETC13(sb3.ToString(), dt.Rows[i2]["MODEL"].ToString(), dt.Rows[i2]["VER"].ToString(), sb2.ToString());
                                if (dt2.Rows.Count > 0)
                                {
                                    dr = dtCost.NewRow();
                                    string MODEL = dt.Rows[i2]["MODEL"].ToString();
                                    string VER2 = dt.Rows[i2]["VER"].ToString();
                                    string VER = dt.Rows[i2]["VER"].ToString().Replace("V", "");
                                    dr["MODEL"] = "";
                                    dr["VER"] = "";
                                    dr["MDL/OC"] = "";
                                    int StartMon = Convert.ToInt32(comboBox2.Text);
                                    string YEAR = comboBox1.Text;
                                    int EndMon = Convert.ToInt32(comboBox3.Text);
       
                                    foreach (string SZ in ZPN)
                                    {
                                        for (int i = StartMon; i <= EndMon; i++)
                                        {
                                            string FD = "(已成交)";

                                            System.Data.DataTable G1 = GETSAP1(MODEL, VER, S2, SZ, YEAR + i.ToString("00"), S);
                                            string RR = G1.Rows[0][0].ToString();
                                            if (String.IsNullOrEmpty(RR))
                                            {

                                                G1 = GETC2(MODEL, VER2, S2, SZ, YEAR, i.ToString("00"), S);
                                                if (G1.Rows.Count > 0)
                                                {
                                                    FD = "(未成交)";
                                                }

                                            }

                                        
                                            string K1 = (YEAR + i.ToString("00") + "_" + SZ + "_" + S + FD);

                                            dr[SZ + i.ToString()] = K1;
                                        }

      

                                        //string K2 = (YEAR + EndMon2.ToString("00") + "_" + SZ + "_" + S + "(報價未成交)");
                                        //dr[SZ + EndMon2.ToString()] = K2;
                                    }

                                    dtCost.Rows.Add(dr);

                                }
                            }

                            if (i3 == 2)
                            {

                    
                                    dr = dtCost.NewRow();
                                    string MODEL = dt.Rows[i2]["MODEL"].ToString();
                                    string VER2 = dt.Rows[i2]["VER"].ToString();
                                    string VER = dt.Rows[i2]["VER"].ToString().Replace("V", "");
                                    dr["MODEL"] = MODEL;
                                    dr["VER"] = VER;
                                    dr["MDL/OC"] = S2;
                                    int StartMon = Convert.ToInt32(comboBox2.Text);
                                    string YEAR = comboBox1.Text;
                                    int EndMon = Convert.ToInt32(comboBox3.Text);
     
                                    foreach (string SZ in ZPN)
                                    {
                                        for (int i = StartMon; i <= EndMon; i++)
                                        {

                                            System.Data.DataTable G1 = GETSAP1(MODEL, VER, S2, SZ, YEAR + i.ToString("00"), S);
                                            string RR = G1.Rows[0][0].ToString();
                                            if (!String.IsNullOrEmpty(RR))
                                            {
                                                dr[SZ + i.ToString()] = G1.Rows[0][0].ToString();

                                            }
                                            else
                                            {
                                                G1 = GETC2(MODEL, VER2, S2, SZ, YEAR, i.ToString("00"), S);
                                                if (G1.Rows.Count > 0)
                                                {
                                                    dr[SZ + i.ToString()] = G1.Rows[0][0].ToString();
                                                }
                                            }

                                        }

                                        //System.Data.DataTable G2 = GETC2(MODEL, VER2, S2, SZ, YEAR, EndMon2.ToString("00"), S);
                                        //if (G2.Rows.Count > 0)
                                        //{

                                        //    dr[SZ + EndMon2.ToString()] = G2.Rows[0][0].ToString();

                                        //}
                                    }

                                    dtCost.Rows.Add(dr);
                  
                            }
                        }
                    }
                }
                
 
              
                    

                //}
            }


            dataGridView1.DataSource = dtCost;
        }

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("MDL/OC", typeof(string));
            int StartMon = Convert.ToInt32(comboBox2.Text);
            int EndMon2 = Convert.ToInt32(comboBox3.Text) + 1;
            string D2 = sb2.ToString().Replace("'", "");
            string[] A2 = D2.Split(new Char[] { ',' });

            //1Z
            foreach (string S2 in A2)
            {
                for (int i = StartMon; i <= EndMon2; i++)
                {

                    //  string K1 = (YEAR + i.ToString("00") + "_" + S2 + " " + S + "_" + S3 + "(已成交)").Replace("'", "");
                    string D1 = S2+i.ToString();
                    dt.Columns.Add(D1, typeof(string));


                }
            }

 


            return dt;
        }


        private void APCOMPARE_Load(object sender, EventArgs e)
        {

            System.Data.DataTable K1 = GETF1();
            if (K1.Rows.Count > 0)
            {

                label7.Text = K1.Rows[0][0].ToString();


            
            }

            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Month(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetMenu.Month(), "DataValue", "DataValue");

            System.Data.DataTable dtSIZE = GETSIZE();

            comboBox4.Items.Clear();


            for (int i = 0; i <= dtSIZE.Rows.Count - 1; i++)
            {
                comboBox4.Items.Add(Convert.ToString(dtSIZE.Rows[i][0]));
            }

            System.Data.DataTable dtMODEL = GETMODEL();

            comboBox5.Items.Clear();


            for (int i = 0; i <= dtMODEL.Rows.Count - 1; i++)
            {
                comboBox5.Items.Add(Convert.ToString(dtMODEL.Rows[i][0]));
            }

            System.Data.DataTable dtVER = GETVER();

            comboBox6.Items.Clear();


            for (int i = 0; i <= dtVER.Rows.Count - 1; i++)
            {
                comboBox6.Items.Add(Convert.ToString(dtVER.Rows[i][0]));
            }



        }
        public System.Data.DataTable GETF1()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 1[FILENAME]DATETIME FROM AP_COMPARE ORDER BY id DESC ");

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

        public System.Data.DataTable GETSIZE()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT SUBSTRING(MODEL,2,2)+'.'+SUBSTRING(MODEL,4,1) FROM AP_COMPARE WHERE LEN(MODEL) > 4  ORDER BY SUBSTRING(MODEL,2,2)+'.'+SUBSTRING(MODEL,4,1)  ");

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


        public System.Data.DataTable GETMODEL()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT MODEL FROM AP_COMPARE  ORDER BY MODEL ");

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

        public System.Data.DataTable GETVER()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT VER FROM AP_COMPARE  ORDER BY VER ");

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
        private System.Data.DataTable GETC1( string SB2)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT MODEL,VER   FROM AP_COMPARE WHERE 1=1   ");
       //     sb.Append(" AND   BU IN  ( " + SB + ") ");
            sb.Append(" AND   GRADE IN  ( " + SB2 + ") ");

            if (checkBox5.Checked)
            {
                sb.Append(" AND   SUBSTRING(MODEL,2,2)+'.'+SUBSTRING(MODEL,4,1) IN  ( " + c + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {
                    sb.Append(" AND SUBSTRING(MODEL,2,2)+'.'+SUBSTRING(MODEL,4,1)=@SIZE ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" AND   MODEL IN  ( " + d + ") ");
            }
            else
            {
                if (comboBox5.Text != "")
                {
                    sb.Append(" AND MODEL=@MODEL ");
                }

            }
            if (comboBox6.Text != "")
            {
                sb.Append(" AND VER=@VER ");
            }

            sb.Append(" ORDER BY MODEL,VER   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SIZE", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@MODEL", comboBox5.Text));
            command.Parameters.Add(new SqlParameter("@VER", comboBox6.Text));
           
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }

        private System.Data.DataTable GETC12(string SB3, string TYPE, string MODEL, string VER, string SBG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT MODEL,VER,[TYPE]  FROM AP_COMPARE WHERE 1=1   ");
            sb.Append(" AND   [TYPE] IN  ( " + SB3 + ") ");
            sb.Append(" AND   GRADE IN  ( " + SBG + ") ");
            sb.Append(" AND MODEL=@MODEL ");
            sb.Append(" AND VER=@VER ");
            sb.Append(" AND [TYPE]=@TYPE    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }
        private System.Data.DataTable GETC13(string SB3, string MODEL, string VER, string SBG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT MODEL,VER,[TYPE]  FROM AP_COMPARE WHERE 1=1   ");
            sb.Append(" AND   [TYPE] IN  ( " + SB3 + ") ");
            sb.Append(" AND  GRADE IN  ( " + SBG + ") ");
            sb.Append(" AND MODEL=@MODEL ");
            sb.Append(" AND VER=@VER ");
            //sb.Append(" AND  BU=@BU ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            //command.Parameters.Add(new SqlParameter("@BU", BU));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }
        private System.Data.DataTable GETC2(string MODEL, string VER, string TYPE, string GRADE, string YEAR, string MONTH, string BU)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PRICE  FROM AP_COMPARE WHERE MODEL=@MODEL AND VER=@VER  AND [TYPE]=@TYPE AND GRADE=@GRADE AND BU=@BU  ");
            sb.Append(" AND SUBSTRING(CDATE,1,4)=@YEAR AND SUBSTRING(CDATE,5,2)=@MONTH ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }

        private System.Data.DataTable GETC2M(string MODEL, string VER, string TYPE, string GRADE, string CDATE, string BU)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PRICE  FROM AP_COMPARE WHERE MODEL=@MODEL AND VER=@VER  AND [TYPE]=@TYPE AND GRADE=@GRADE AND BU=@BU  ");
            sb.Append(" AND CDATE=@CDATE ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@CDATE", CDATE));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }
        private System.Data.DataTable GETSAP1(string MODEL, string VER, string TYPE, string GRADE, string YM, string GD)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            string GD2 = GD;
            if (GD2 == "DT")
            {
                GD2 = "DD";
            }

            sb.Append("   Declare @N1 varchar(200) ");
            sb.Append("  select @N1 =SUBSTRING(COALESCE(@N1 + '/',''),0,199) + 單價 ");
            sb.Append("   from   (");
            sb.Append("    SELECT DISTINCT  CAST(CAST(T1.Price AS decimal(16,2)) AS VARCHAR) 單價");
            sb.Append("              FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append("              INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)         ");
            sb.Append("              left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
            sb.Append("              left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)     ");
            sb.Append("  INNER join opdn t44 on (t4.docentry=t44.docentry)   ");
            sb.Append("              left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )    ");
            sb.Append("              left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )    ");
            sb.Append("              WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND (T1.[Quantity]-ISNULL(T8.QUANTITY,0) <> 0)  AND (T1.[Quantity]-ISNULL(T9.QUANTITY,0) <> 0)      ");
            sb.Append("			  AND   T11.U_VERSION=@VER AND T11.U_GRADE=@GRADE AND  Convert(varchar(6),T44.[DOCDate],112) =@YM  AND REPLACE(REPLACE(T0.CARDCODE,'S0001-',''),'S0623-','')=@GD ");
            if (TYPE == "MDL")
            {
                sb.Append(" AND T11.U_TMODEL=@MODEL ");
            }
            else
            {
                sb.Append(" AND T11.U_TMODEL='O'+SUBSTRING(@MODEL,2,8) ");
            }
            sb.Append("			  AND T1.PRICE <> 0 ) AS A");
            sb.Append("			  SELECT ISNULL(@N1,'') A");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@YM", YM));
            command.Parameters.Add(new SqlParameter("@GD", GD2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }

        private void button10_Click(object sender, EventArgs e)
        {

            APSIZE frm1 = new APSIZE();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox5.Checked = true;
                c = frm1.q;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            APSIZE2 frm1 = new APSIZE2();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox3.Checked = true;
                d = frm1.q;

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//Excel
using Microsoft.Office.Interop.Excel;
//
using System.IO;
//
using System.Data.SqlClient;

//20081204 �ץ� �W�[
//���o�Y�@���I���v�w�s�q
//private int GetHisoryStock(string ItemCode,  string BaseDate)

namespace ACME
{
    public partial class AcmeItemHelper : Form
    {
        private string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
        
        private System.Data.DataTable TempDt;

        private System.Data.DataTable dtItemNo;

        private System.Data.DataTable dtPrint;

        private string FileName;

        private int[] iCount;

        public AcmeItemHelper()
        {
            InitializeComponent();
        }

      
        private void GetExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
           
            
            excelApp.Visible = checkBox1.Checked;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
         
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //���o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
         
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //���o Excel ���ϥΰϰ�
              int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
              int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

              progressBar1.Maximum = iRowCnt;

            
            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            //�ϥΩT�w�榡
            //C7 �O Model no
            //�q�ĤQ��_��
            //No.	Serial Number	Version	W/C	IQC/LR/FR	Defect Reason	Defect Confirm	Judge	Made in Taiwan/China
            //Serial Number �p�G�ťիh���@�F


            //object SelectCell = "C7";
            //range = excelSheet.get_Range(SelectCell, SelectCell);

            //���wAV ���d��
            //object AV_SelectCell_From = "B4";
            //object AV_SelectCell_To = "AL23";
            //FixedRange = AV_Sheet.get_Range(AV_SelectCell_From, AV_SelectCell_To);


            //object GD_SelectCell_From = "B3";
            //object GD_SelectCell_To = "O21";
            //range = GD_Sheet.get_Range(GD_SelectCell_From, GD_SelectCell_To);

            //int iColCnt = FixedRange.Rows.Count;
            //int iRowCnt = FixedRange.Columns.Count;

            //  MessageBox.Show(range.Text.ToString());

            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //�Ĥ@��n
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {
                    
                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord,1 ]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().Replace("-", "");
                    //SERIAL_NO = SERIAL_NO.Substring(0, 7);
                    range.Select();


                    //�p�G�䤣��ɤ~�s�W
                     drFind =TempDt.Rows.Find(SERIAL_NO);

                     if (drFind == null)
                     {

                         dr = TempDt.NewRow();

                         dr["Mo"] = SERIAL_NO;

                         TempDt.Rows.Add(dr);
                     }

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
                //�i�H�N Excel.exe �M��
                System.GC.WaitForPendingFinalizers();
                
            }
        }

        //�ʺA���͸�Ƶ��c
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
        
            //�Ĥ@�өT�w���(�u�渹�X)
            dt.Columns.Add("Mo", typeof(string));

            //�̫�@���`�p
          //  dt.Columns.Add("Qty", typeof(int));

            
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Mo"];
            dt.PrimaryKey = colPk;
            
            return dt;
        }

        private System.Data.DataTable MakeTableItemNo()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //�Ĥ@�өT�w���(�u�渹�X)
            dt.Columns.Add("ItemNo", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["ItemNo"];
            dt.PrimaryKey = colPk;

            return dt;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
              //TempDt= MakeTable();

              //CreateboundColumn("Mo", "�u�渹�X");
              //CreateboundColumn("Qty", "�ƶq");
             // dataGridView1.DataSource = TempDt;

            dtItemNo = MakeTableItemNo();


            DateTime  aDate = DateTime.Now.AddMonths(-1);

            textBox2.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM")+"01";
            textBox3.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + DateTime.DaysInMonth(aDate.Year, aDate.Month).ToString();
            
        }

        private void CreateboundColumn(string FieldName,string Caption)
        {
            // Initialize the button column.
            DataGridViewTextBoxColumn TextBoxColumn = new DataGridViewTextBoxColumn();
            TextBoxColumn.Name = FieldName;
            TextBoxColumn.HeaderText = Caption;
            TextBoxColumn.DataPropertyName = FieldName;

            // Add the button column to the control.
            //dataGridView1.Columns.Add(TextBoxColumn);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                //iCount = new int[TempDt.Rows.Count];
                //GetExcelProduct(FileName);
            }
        }


        private bool CheckSerial(string Serial_no)
        {
            string Mo = string.Empty;

            int iChecked=-1;

            for (int i = 0; i <= TempDt.Rows.Count - 1; i++)
            {
                Mo =Convert.ToString(TempDt.Rows[i]["Mo"]);
                Mo = Mo.Substring(3, 7);

                if (Mo == Serial_no)
                {
                    iChecked = i;

                    iCount[i] = iCount[i] + 1;

                    if (iCount[i] > Convert.ToInt32(TempDt.Rows[i]["Qty"]))
                    {
                        return false;
                    }

                    return true;
                }
            
            }
            return false;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
           // {
                //FileName = openFileDialog1.FileName;
                //iCount = new int[TempDt.Rows.Count];
                WriteExcelProduct(FileName);
                MessageBox.Show("OK");

               // dataGridView1.DataSource = dtItemNo;
            //}
        }


        private void WriteExcelProduct(string ExcelFile)
        {

            string DocDate1 = textBox2.Text;
            string DocDate2 = textBox3.Text;


            try
            {
                StrToDate(DocDate1);
            }
            catch
            {
                MessageBox.Show("�����J���~");
                textBox2.Focus();
                return;
            }


            try
            {
                StrToDate(DocDate2);
            }
            catch
            {
                MessageBox.Show("�����J���~");
                textBox3.Focus();
                return;
            }


            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = checkBox1.Checked;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //���o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //���o Excel ���ϥΰϰ�
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

          //  MessageBox.Show(iRowCnt.ToString());

            progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;

                DataRow dr;
                DataRow drFind;
                string ItemCode="";


                //�Ĥ@��n
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    progressBar1.Value = iRecord;
                    progressBar1.PerformStep();
                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();

                    ItemCode = range.Text.ToString().Trim();

                    if (ItemCode.Length != 15)
                    {
                       // SERIAL_NO = SERIAL_NO.Substring(0, 16);
                        continue;
                    }

                    //SERIAL_NO = SERIAL_NO.Substring(0, 7);
                    range.Select();

                  
                    //drFind =TempDt.Rows.Find(SERIAL_NO);

                    //if (drFind != null)
            
                    //{
                    //   //�g�J��m
                    //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord,Convert.ToInt32(textBox1.Text)+1]);
                    //    range.Value2 = "Y";
                   

                    //}
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, Convert.ToInt32(textBox1.Text)]);
                    range.Value2 = GetDataByProd_noSAP(ItemCode,DocDate1,DocDate2);


                    //20081204
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, Convert.ToInt32(textBox4.Text)]);
                    range.Value2 = GetHisoryStock(ItemCode,  DocDate2);



                    //�g�J
                    //�p�G�䤣��ɤ~�s�W
                    drFind = dtItemNo.Rows.Find(ItemCode);

                     if (drFind == null)
                     {

                         dr = dtItemNo.NewRow();

                         dr["ItemNo"] = ItemCode;

                         dtItemNo.Rows.Add(dr);
                     }

                }


            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
               "Acme_" + Path.GetFileName(FileName);

                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                //�i�H�N Excel.exe �M��
                System.GC.WaitForPendingFinalizers();
                
            }
        }


//SELECT SUM(�ƶq) �ƶq
//FROM
//(
//SELECT SUM(T1.Quantity) �ƶq 
//from ODLN T0 
//inner join DLN1 T1 on T0.DocEntry = T1.DocEntry 
//WHERE T1.ItemCode='TAM17EG01.3DZZ1'
//Union All
//SELECT SUM(T1.Quantity) �ƶq 
//from ORDN T0 
//inner join RDN1 T1 on T0.DocEntry = T1.DocEntry 
//WHERE T1.ItemCode='TAM17EG01.3DZZ1'
//)  T

        private int  GetDataByProd_noSAP(string ItemCode,string DocDate1,string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT Convert(int,SUM(�ƶq)) �ƶq");
            sb.Append(" FROM");
            sb.Append(" (");
            sb.Append(" SELECT SUM(T1.Quantity) �ƶq ");
            sb.Append(" from ODLN T0 ");
            sb.Append(" inner join DLN1 T1 on T0.DocEntry = T1.DocEntry ");
            sb.Append(" WHERE T1.ItemCode=@ItemCode");
            sb.Append(" AND   T0.DocDate >= @DocDate1");
            sb.Append(" AND   T0.DocDate <= @DocDate2");

            sb.Append(" Union All");
            sb.Append(" SELECT SUM(T1.Quantity) �ƶq ");
            sb.Append(" from ORDN T0 ");
            sb.Append(" inner join RDN1 T1 on T0.DocEntry = T1.DocEntry ");
            sb.Append(" WHERE T1.ItemCode=@ItemCode");
            sb.Append(" AND   T0.DocDate >= @DocDate1");
            sb.Append(" AND   T0.DocDate <= @DocDate2");
            sb.Append(" )  T");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

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
            //return ds.Tables["PRODUCT"];

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            try
            {
                return Convert.ToInt32(dt.Rows[0]["�ƶq"]);
            }
            catch
            {
                return 0;
            }

        }

        //����B�z--------------------------------------------------------------------------------------------
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

        private void button3_Click(object sender, EventArgs e)
        {
            //�v���d�߼g�ܼȦs��

            DataRow dr;
            System.Data.DataTable dtGetOinm;
            string ItemCode;


            dtPrint = GetOinmEmpty();

            //dr = dtPrint.NewRow();

            //dr["���ؽs��"] = "12234";

            //dtPrint.Rows.Add(dr);

            progressBar1.Maximum = dtItemNo.Rows.Count;

            for (int i = 0; i <= dtItemNo.Rows.Count - 1; i++)
            {

                progressBar1.Value = i;
                progressBar1.PerformStep();

                ItemCode =Convert.ToString(dtItemNo.Rows[i][0]);

                //���o�s�f�L�b�M��
                dtGetOinm = GetOinm(ItemCode, textBox2.Text, textBox3.Text);


                for (int RecCount = 0; RecCount <= dtGetOinm.Rows.Count - 1; RecCount++)
                {

                    dr = dtPrint.NewRow();


                    for (int j = 0; j <= dtGetOinm.Columns.Count - 1; j++)
                    {

                        dr[j] = dtGetOinm.Rows[RecCount][j];
                    }

                    dtPrint.Rows.Add(dr);
                }
            }

            dataGridView1.DataSource =dtPrint;


        }


        private System.Data.DataTable GetOinm(string ItemCode, string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ItemCode] as ���ؽs��, T0.[Dscription] as ���ػ���, ");
            sb.Append(" Convert(Varchar(8),T0.[DocDate],112) as �L�b��� , ");
            sb.Append(" ��� =( ");
            sb.Append(" CASE ");
            sb.Append(" WHEN T0.[TransType]=15  THEN 'Delivery' ");
            sb.Append(" WHEN T0.[TransType]=16  THEN 'Returns'");
            sb.Append(" WHEN T0.[TransType]=13  THEN 'A/R Invoice'");
            sb.Append(" WHEN T0.[TransType]=14  THEN 'A/R Credit Memo'");
            sb.Append(" WHEN T0.[TransType]=132 THEN 'Correction Invoice'");
            sb.Append(" WHEN T0.[TransType]=20  THEN 'Goods Receipt'");
            sb.Append(" WHEN T0.[TransType]=21  THEN 'Goods Returns'");
            sb.Append(" WHEN T0.[TransType]=18  THEN 'A/P Invoice'");
            sb.Append(" WHEN T0.[TransType]=19  THEN 'A/P Credit Memo'");
            sb.Append(" WHEN T0.[TransType]=-2  THEN 'Opening Balance'");
            sb.Append(" WHEN T0.[TransType]=58  THEN 'Stock Update'");
            sb.Append(" WHEN T0.[TransType]=59  THEN 'Goods Receipt'");
            sb.Append(" WHEN T0.[TransType]=60  THEN 'Goods Issue'");
            sb.Append(" WHEN T0.[TransType]=67  THEN 'Inventory Transfers'");
            sb.Append(" WHEN T0.[TransType]=68  THEN 'Work Instructions'");
            sb.Append(" WHEN T0.[TransType]=-1  THEN 'All Transactions'");
            sb.Append(" ELSE 'Other'");
            sb.Append(" END) , ");
            sb.Append(" T0.[CreatedBy] as ���, ");
            sb.Append(" Convert(int,T0.[InQty]) as ���f�q, Convert(int,T0.[OutQty]) as �o�f�q, T0.[Price ] as ����,");
            sb.Append(" T0.[CardCode] as �~�ȹ٦�, T0.[CardName]  as �~�ȹ٦�W��");
            sb.Append(" FROM OINM T0 ");
            sb.Append(" WHERE T0.[ItemCode] =@ItemCode");
            sb.Append(" AND T0.[DocDate] >=@DocDate1 AND  T0.[DocDate] <=@DocDate2");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

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
            return ds.Tables["PRODUCT"];
        }

        private System.Data.DataTable GetOinm2(string ItemCode, string DocDate1, string DocDate2)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ItemCode] as ���ؽs��, T0.[Dscription] as ���ػ���, ");
            sb.Append(" Convert(int,SUM(T0.[InQty])) as ���f�q, Convert(int,SUM(T0.[OutQty])) as �o�f�q ");
            sb.Append(" FROM OINM T0 ");
            sb.Append(" WHERE T0.[ItemCode] =@ItemCode");
            sb.Append(" AND T0.[DocDate] >=@DocDate1 AND  T0.[DocDate] <=@DocDate2");
            sb.Append(" Group by  T0.[ItemCode] ,T0.[Dscription] ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

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
            return ds.Tables["PRODUCT"];
        }

        private System.Data.DataTable GetOinmEmpty()
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ItemCode] as ���ؽs��, T0.[Dscription] as ���ػ���, ");
            sb.Append(" Convert(Varchar(8),T0.[DocDate],112) as �L�b��� , ");
            sb.Append(" ��� =( ");
            sb.Append(" CASE ");
            sb.Append(" WHEN T0.[TransType]=15  THEN 'Delivery' ");
            sb.Append(" WHEN T0.[TransType]=16  THEN 'Returns'");
            sb.Append(" WHEN T0.[TransType]=13  THEN 'A/R Invoice'");
            sb.Append(" WHEN T0.[TransType]=14  THEN 'A/R Credit Memo'");
            sb.Append(" WHEN T0.[TransType]=132 THEN 'Correction Invoice'");
            sb.Append(" WHEN T0.[TransType]=20  THEN 'Goods Receipt'");
            sb.Append(" WHEN T0.[TransType]=21  THEN 'Goods Returns'");
            sb.Append(" WHEN T0.[TransType]=18  THEN 'A/P Invoice'");
            sb.Append(" WHEN T0.[TransType]=19  THEN 'A/P Credit Memo'");
            sb.Append(" WHEN T0.[TransType]=-2  THEN 'Opening Balance'");
            sb.Append(" WHEN T0.[TransType]=58  THEN 'Stock Update'");
            sb.Append(" WHEN T0.[TransType]=59  THEN 'Goods Receipt'");
            sb.Append(" WHEN T0.[TransType]=60  THEN 'Goods Issue'");
            sb.Append(" WHEN T0.[TransType]=67  THEN 'Inventory Transfers'");
            sb.Append(" WHEN T0.[TransType]=68  THEN 'Work Instructions'");
            sb.Append(" WHEN T0.[TransType]=-1  THEN 'All Transactions'");
            sb.Append(" ELSE 'Other'");
            sb.Append(" END) , ");
            sb.Append(" T0.[CreatedBy] as ���, ");
            sb.Append(" Convert(int,T0.[InQty]) as ���f�q, Convert(int,T0.[OutQty]) as �o�f�q, T0.[Price ] as ����,");
            sb.Append(" T0.[CardCode] as �~�ȹ٦�, T0.[CardName]  as �~�ȹ٦�W��");
            sb.Append(" FROM OINM T0 ");
            
            sb.Append(" WHERE 1=0");
        


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            //command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            //command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

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
            return ds.Tables["PRODUCT"];
        }

        /// <summary>
        /// 20090805
        /// </summary>
        /// <returns></returns>
        private System.Data.DataTable GetOinmEmpty2()
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ItemCode] as ���ؽs��, T0.[Dscription] as ���ػ���, ");
            sb.Append(" Convert(int,T0.[InQty]) as ���f�q, Convert(int,T0.[OutQty]) as �o�f�q, ");
            sb.Append(" 0 as �w�s�q ");
            sb.Append(" FROM OINM T0 ");
            sb.Append(" WHERE 1=0");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            //command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            //command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));

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
            return ds.Tables["PRODUCT"];
        }

        private void button4_Click(object sender, EventArgs e)
        {
           ExcelReport.GridViewToExcel(dataGridView1);
        }


       


        //���o�Y�@���I���v�w�s�q
        private int GetHisoryStock(string ItemCode,  string BaseDate)
        {
            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  SUM(T0.[InQty])-SUM(T0.[OutQty]) �w�s�q");
            sb.Append(" FROM  [OINM] T0  ");
            sb.Append(" INNER  JOIN [OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' ");
            sb.Append(" and  (T0.[docdate] >='2007.12.31' AND  T0.[docdate] <= @BaseDate) ");
            sb.Append(" and T0.ItemCode = @ItemCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@BaseDate", BaseDate));

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
            //return ds.Tables["PRODUCT"];

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            try
            {
                return Convert.ToInt32(dt.Rows[0]["�w�s�q"]);
            }
            catch
            {
                return 0;
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //�v���d�߼g�ܼȦs��
            string BaseDate =textBox3.Text;

            DataRow dr;
            System.Data.DataTable dtGetOinm;
            string ItemCode;


            dtPrint = GetOinmEmpty2();


            progressBar1.Maximum = dtItemNo.Rows.Count;

            for (int i = 0; i <= dtItemNo.Rows.Count - 1; i++)
            {

                progressBar1.Value = i;
                progressBar1.PerformStep();

                ItemCode = Convert.ToString(dtItemNo.Rows[i][0]);

                //���o�s�f�L�b�M��
                dtGetOinm = GetOinm2(ItemCode, textBox2.Text, textBox3.Text);


                for (int RecCount = 0; RecCount <= dtGetOinm.Rows.Count - 1; RecCount++)
                {

                    dr = dtPrint.NewRow();

                    int StockQty = GetHisoryStock(ItemCode, BaseDate);


                    dr[0] = dtGetOinm.Rows[RecCount][0];
                    dr[1] = dtGetOinm.Rows[RecCount][1];
                    dr[2] = dtGetOinm.Rows[RecCount][2];
                    dr[3] = dtGetOinm.Rows[RecCount][3];
                    dr[4] = StockQty;

                    dtPrint.Rows.Add(dr);
                }
            }

            dataGridView1.DataSource = dtPrint;


        }


    }




 





}
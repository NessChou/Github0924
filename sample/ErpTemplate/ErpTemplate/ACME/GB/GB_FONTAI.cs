using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ACME
{
    public partial class GB_FONTAI : Form
    {
        System.Data.DataTable dtCost = null;
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_FONTAI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                string QTY2 = "";
                string UNIT = "";
                string ITEMNAME = "";
                string COST = "";
                string AMT = "";
                TRUNFONTAI();
                dtCost = MakeTableCombine();
                string F = opdf.FileName;
                GetExcelContentGD44(F);
                System.Data.DataTable FON1 = GetFONTAI();
                if (FON1.Rows.Count > 0)
                {
                    for (int i = 0; i <= FON1.Rows.Count - 1; i++)
                    {
                        string ITEMCODE = FON1.Rows[i]["ITEMCODE"].ToString();
                        string QTY = FON1.Rows[i]["QTY"].ToString();
                        System.Data.DataTable G1 = GetCHO(ITEMCODE);
                        if (G1.Rows.Count > 0)
                        {
                            DataRow dd = G1.Rows[0];

                            UNIT = dd["單位"].ToString();
                            QTY2 = dd["數量"].ToString();
                            ITEMNAME = dd["品名"].ToString();
                            COST = dd["單價"].ToString();
                            AMT = dd["金額"].ToString();
                            if (!String.IsNullOrEmpty(QTY) && !String.IsNullOrEmpty(QTY2))
                            {
                                decimal Q1 = Convert.ToDecimal(QTY);
                                decimal Q2 = Convert.ToDecimal(QTY2);
                                decimal CC = Convert.ToDecimal(COST);
                                decimal QTY3 = 0;

                                DataRow dr = null;
                                //if (Q1 != Q2)
                                //{
                                    decimal n;
                                    if (decimal.TryParse(QTY, out n) && decimal.TryParse(QTY2, out n))
                                    {
                                        QTY3 = Q1 - Q2;

                                    }

                                    dr = dtCost.NewRow();
                                    dr["編號"] = (i + 1).ToString();
                                    dr["品項"] = dd["品項"].ToString();
                                    dr["零售批發"] = dd["零售批發"].ToString();
                                    dr["料號"] = ITEMCODE;
                                    dr["品名規格"] = ITEMNAME;
                                    dr["單位"] = UNIT;
                                    dr["逢泰庫存"] = Convert.ToDecimal(QTY).ToString("#,##0.00");
                                    dr["正航逢泰倉"] = Convert.ToDecimal(QTY2).ToString("#,##0.00");
                                    dr["差異"] = Convert.ToDecimal(QTY3).ToString("#,##0.00");
                                    dr["成本單價"] = Convert.ToDecimal(COST).ToString("#,##0.00");
                                    dr["成本金額"] = Convert.ToDecimal(AMT).ToString("#,##0.00");
                                    dr["差異金額"] = (CC * QTY3).ToString("#,##0.00");
                                    System.Data.DataTable G2 = GetCHO1("100",ITEMCODE);
            
                                    if (G2.Rows.Count > 0)
                                    {
                                        dr["進貨量"] = G2.Rows[0][0].ToString();
                                        
                                    }
                                    System.Data.DataTable G3 = GetCHO1("500", ITEMCODE);
                                    if (G3.Rows.Count > 0)
                                    {
                                        dr["銷貨量"] = G3.Rows[0][0].ToString();

                                    }
                                    System.Data.DataTable G4 = GetCHO1("300", ITEMCODE);
                                    if (G4.Rows.Count > 0)
                                    {
                                        dr["調整量"] = G4.Rows[0][0].ToString();

                                    }
                                    dtCost.Rows.Add(dr);


                              //  }
                            }
                        }
                 
                    }
                }
           
                dataGridView1.DataSource = dtCost;

                for (int i = 5; i <= 10; i++)
                {
                    DataGridViewColumn col = dataGridView1.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.00";


                }
            }
        }

        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false  ;

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


            string ITEMCODE;
            string QTY ;

  
            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                ITEMCODE = range.Text.ToString().ToUpper().Replace("RGN-", "");



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 17]);
                QTY = range.Text.ToString();

                if (!String.IsNullOrEmpty(ITEMCODE))
                {
                    AddFONTAI(ITEMCODE, Convert.ToDecimal(QTY));
              

       
                  

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
        public void AddFONTAI(string  ITEMCODE,decimal  QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_FONTAI(ITEMCODE,QTY) values(@ITEMCODE,@QTY)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
   
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
        private System.Data.DataTable GetFONTAI()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,SUM(QTY) QTY FROM GB_FONTAI GROUP BY ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public void TRUNFONTAI()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE GB_FONTAI", connection);
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
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("編號", typeof(string));
            dt.Columns.Add("品項", typeof(string));
            dt.Columns.Add("零售批發", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("單位", typeof(string));
            dt.Columns.Add("逢泰庫存", typeof(string));
            dt.Columns.Add("正航逢泰倉", typeof(string));
            dt.Columns.Add("差異", typeof(string));
            dt.Columns.Add("成本單價", typeof(string));
            dt.Columns.Add("成本金額", typeof(string));
            dt.Columns.Add("差異金額", typeof(string));
            dt.Columns.Add("進貨量", typeof(string));
            dt.Columns.Add("銷貨量", typeof(string));
            dt.Columns.Add("調整量", typeof(string));
            return dt;
        }
        public System.Data.DataTable GetCHO(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T0.ProdID 料號,T0.ProdName 品名,T0.Unit 單位   ,T0.CAvgCost 單價 ");
            sb.Append(" ,ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID ),0) 數量 ");
            sb.Append("  ,ISNULL((select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID )*T0.CAvgCost,0) 金額 ");
            sb.Append(" ,CASE WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T0.ProdID,1,1)='P' THEN '加工品'         ");
            sb.Append(" END 品項,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'       ");
            sb.Append(" END '零售批發'     FROM comProduct T0  Left Join comProductClass K On T0.ClassID =K.ClassID WHERE T0.ProdID =@ProdID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        public System.Data.DataTable GetCHO1(string Flag, string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  Select   ISNULL(SUM(A.Quantity),0) 數量 From DBO.ComProdRec A  Where A.Flag=@Flag AND A.ProdID =@ProdID AND BILLDATE=@BILLDATE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@BILLDATE",DateTime.Now.ToString("yyyyMMdd")));
          //  command.Parameters.Add(new SqlParameter("@BILLDATE", "20190704"));
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

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

    }
}

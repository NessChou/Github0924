using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class WH_OWTR : Form
    {
        private string FileName;
        public WH_OWTR()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                WriteExcelProduct6(FileName);
            }
        }
        private System.Data.DataTable MakeTableF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("項次", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("出倉倉別", typeof(string));
            dt.Columns.Add("出倉倉別編號", typeof(string));
            dt.Columns.Add("調撥倉別", typeof(string));
            dt.Columns.Add("調撥倉別編號", typeof(string));
            return dt;
        }
        private void WriteExcelProduct6(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;






            try
            {
                System.Data.DataTable dtCost  = MakeTableF();
                DataRow dr = null;
                string 項次;
                string 料號;
                string 品名;
                string 數量;
                string 出倉倉別;
                string 調撥倉別;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    項次 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    料號 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    品名 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    數量 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    出倉倉別 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    調撥倉別 = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(品名))
                    {
                        System.Data.DataTable G2 = GetDI2(出倉倉別);
                        System.Data.DataTable G3 = GetDI2(調撥倉別);
                        if (G2.Rows.Count > 0 && G3.Rows.Count > 0)
                        {
                            dr = dtCost.NewRow();
                            dr["項次"] = 項次;
                            dr["料號"] = 料號;
                            dr["品名"] = 品名;
                            dr["數量"] = 數量;
                            dr["出倉倉別"] = 出倉倉別;
                            dr["出倉倉別編號"] = G2.Rows[0][0].ToString();
                            dr["調撥倉別"] = 調撥倉別;
                            dr["調撥倉別編號"] = G3.Rows[0][0].ToString();
                            dtCost.Rows.Add(dr);
                        }
                    }

                }

                dataGridView1.DataSource = dtCost;


            }
            finally
            {

         


                try
                {
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
                System.GC.WaitForPendingFinalizers();


               
            }



        }
        public System.Data.DataTable GetDI2(string WhsName)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select WHSCODE   from ACMESQL02.DBO.OWHS  WHERE (STREET =@WhsName OR CITY =@WhsName) ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WhsName", WhsName));

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

        public System.Data.DataTable GetDI4()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OWTR");
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
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult resultS;
            resultS = MessageBox.Show("請確認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (resultS == DialogResult.Yes)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = "acmesql98";
                oCompany.UserName = "manager";
                oCompany.Password = "19571215";
                int result = oCompany.Connect();
                if (result == 0)
                {

                    SAPbobsCOM.StockTransfer oStock = null;
                    oStock = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);


                        string WHNAME = dataGridView1.Rows[0].Cells["調撥倉別編號"].Value.ToString();
                        string WHNAME1 = dataGridView1.Rows[0].Cells["出倉倉別編號"].Value.ToString();
  
                            oStock.CardCode = "";
                            oStock.FromWarehouse = WHNAME1;

                            for (int f = 0; f <= dataGridView1.Rows.Count - 1; f++)
                            {
                        DataGridViewRow row;

                        row = dataGridView1.Rows[f];
             
                                oStock.Lines.ItemCode = row.Cells["料號"].Value.ToString();
                                oStock.Lines.ItemDescription = row.Cells["品名"].Value.ToString();
                                oStock.Lines.Quantity = Convert.ToDouble(row.Cells["數量"].Value.ToString());
                                oStock.Lines.WarehouseCode = WHNAME;
                                oStock.Lines.Add();
                            }


                        

          
                    int res = oStock.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
          
                        MessageBox.Show("上傳成功 調撥單號 : " + OWTR);

                    }


                }
                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }
            }
        }

    }
}

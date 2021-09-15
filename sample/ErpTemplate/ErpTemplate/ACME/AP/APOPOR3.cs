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
    public partial class APOPOR3 : Form
    {
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public APOPOR3()
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
                               DELOPOR();
                               GD5(opdf.FileName);
                             
                 
                        
                                   System.Data.DataTable G2 = GetOPOR2();

                                   if (G2.Rows.Count > 0)
                                   {
                                       int l = 0;
                                       for (int i = 0; i <= G2.Rows.Count - 1; i++)
                                       {
                                           l++;

                                           string T1 = G2.Rows[i][0].ToString();
                                           System.Data.DataTable G3 = GetOPOR3(T1);
                                        
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

                                                   UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]));
                                               }
                                           }
                                       }

                                   

                                   System.Data.DataTable G1 = GetOPOR();
                                   dataGridView1.DataSource = G1;
                                                     System.Data.DataTable GG2 = GetOPOR6();
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
        private void GD5(string ExcelFile)
        {

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
                            System.Data.DataTable F1 = GetITEMCODE(ITEMCODE, GRADE, VER);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();
                            }
                            else
                            {
                                ITEM = "ACME00004.00004";
                            }
                          

                            ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK1 +" "+ REMARK2);
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

        public static System.Data.DataTable GetITEMCODE(string MODEL, string U_GRADE, string VERSION)
        {
            SqlConnection MyConnection = globals.shipConnection;
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

        public static System.Data.DataTable GetOPOR()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR ORDER BY DOC,ID");

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
        public static System.Data.DataTable GetOPOR6()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR ");

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
        public static System.Data.DataTable GetOPOR6IN(string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR WHERE ID IN ( " + cs + ") ");

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
        public static System.Data.DataTable GetOPOR2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPOR ORDER BY DOCENTRY ");

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
        public static System.Data.DataTable GetOPOR4()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR ORDER BY DOC ");

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
        public static System.Data.DataTable GetOPOR3(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPOR WHERE DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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
        public static System.Data.DataTable GetOPOR5(string DOC)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE DOC=@DOC ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
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
        public static System.Data.DataTable GetOPOR5ID(string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE ID IN ( " + cs + ")  ");

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

        public void DELOPOR()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE AP_OPOR ", connection);
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

        public void ADDOPOR(int DOCENTRY, string ITEMCODE, decimal QTY, decimal PRICE, decimal AMT, string REMARK)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPOR(DOCENTRY,ITEMCODE,QTY,PRICE,AMT,REMARK) values(@DOCENTRY,@ITEMCODE,@QTY,@PRICE,@AMT,@REMARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
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


        public void UPOPOR(int DOC,int LINENUM, int ID)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPOR SET DOC=@DOC,LINENUM=@LINENUM WHERE ID=@ID ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ID", ID));


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
            oCompany.UserName = "A02";
            oCompany.Password = "6500";
            int result = oCompany.Connect();
            if (result == 0)
            {


                if (listBox1.Items.Count == 0)
                {

                    System.Data.DataTable G2 = GetOPOR4();

                    if (G2.Rows.Count > 0)
                    {

                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {
                            SAPbobsCOM.Documents oPURCH = null;
                            oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                            string T1 = G2.Rows[n][0].ToString();

                            System.Data.DataTable G3 = GetOPOR5(T1);

                            if (G3.Rows.Count > 0)
                            {
                                oPURCH.CardCode = "S0001-GD";
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

                                    System.Data.DataTable GG3 = GetOPOR5(T1);

                                    if (GG3.Rows.Count > 0)
                                    {
                                        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                        {

                                            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                            string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                            UPDATEOPOR(OWTR, LINENUM, REMARK);
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

                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                    System.Data.DataTable G3 = GetOPOR5ID(sb.ToString());

                    if (G3.Rows.Count > 0)
                    {
                        oPURCH.CardCode = "S0001-GD";
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
                            if (ITEMCODE == "ACME00004.00004")
                            {
                                oPURCH.Lines.ItemDescription = ITEMCODE;
                            }
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

                            System.Data.DataTable GG3 = GetOPOR5ID(sb.ToString());

                            if (GG3.Rows.Count > 0)
                            {
                                for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                                {

                                    string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                                    string REMARK = GG3.Rows[j]["REMARK"].ToString();
                                    UPDATEOPOR(OWTR, LINENUM, REMARK);
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
        }


        public System.Data.DataTable GetDI4()
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
         //   SqlConnection MyConnection = globals.shipConnection;
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

                    System.Data.DataTable GG2 = GetOPOR6IN(sb.ToString());
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
    }
}

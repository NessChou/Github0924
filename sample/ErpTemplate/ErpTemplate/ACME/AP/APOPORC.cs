using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualBasic.Devices;
namespace ACME
{
    public partial class APOPORC : Form
    {
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string FA = "acmesql98";
        public APOPORC()
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
                DELOPOR(fmLogin.LoginID.ToString());
                GD5(opdf.FileName);

                textBox1.Text = "DD";

                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    //int l = 0;
                    //for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    //{
                    //    l++;

                    //    string T1 = G2.Rows[i][0].ToString();
                    //    System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());

                    //    if (G3.Rows.Count > 0)
                    //    {
                    //        int h = -1;
                    //        for (int s = 0; s <= G3.Rows.Count - 1; s++)
                    //        {
                    //            h++;

                    //            if (h >= 20)
                    //            {

                    //                l++;
                    //                h = 0;
                    //            }

                    //            UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                    //        }
                    //    }
                    //}



                    System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                    dataGridView1.DataSource = G1;
                    System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                    label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                    label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

                }
            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        private void D1(string CS)
        {
            UPOPORD(0, 0, fmLogin.LoginID.ToString());
            System.Data.DataTable G2 = null;
            int ROW = 0;
            if (textBox1.Text == "GD" || (listBox1.Items.Count == 0))
            {
                G2 = GetOPOR2(fmLogin.LoginID.ToString());
            }
            else
            {
                G2 = GetOPOR2D(CS, fmLogin.LoginID.ToString());
            }
            if (textBox1.Text == "DD")
            {
                ROW = 20;
            }

            if (textBox1.Text == "GD")
            {
                ROW = 20;
            }
            if (textBox1.Text == "TV")
            {
                ROW = Convert.ToInt32(GetOPOR2TV(fmLogin.LoginID.ToString()).Rows[0][0].ToString());//全部同PO
            }

            if (textBox1.Text == "PID")
            {
                ROW = 20;
            }
            if (G2.Rows.Count > 0)
            {
                int l = 0;
                for (int i = 0; i <= G2.Rows.Count - 1; i++)
                {
                    l++;

                    string T1 = G2.Rows[i][0].ToString();
                    System.Data.DataTable G3 = null;
                    if (textBox1.Text == "GD" || (listBox1.Items.Count == 0))
                    {
                        G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());
                    }
                    else
                    {
                        G3 = GetOPOR3D(CS, T1, fmLogin.LoginID.ToString());
                    }

                    if (G3.Rows.Count > 0)
                    {
                        int h = -1;
                        for (int s = 0; s <= G3.Rows.Count - 1; s++)
                        {
                            h++;

                            if (h >= ROW)
                            {

                                l++;
                                h = 0;
                            }

                            UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                        }
                    }
                }


            }
        }
        private void D2(string CS)
        {
            UPOPORDQ(0, 0, fmLogin.LoginID.ToString());
            System.Data.DataTable G2 = null;
            int ROW = 0;
            if (textBox1.Text == "GD" || (listBox2.Items.Count == 0))
            {
                G2 = GetOPOR2Q(fmLogin.LoginID.ToString());
            }
            else
            {
                G2 = GetOPOR2DQ(CS, fmLogin.LoginID.ToString());
            }


            if (G2.Rows.Count > 0)
            {
                int l = 0;
                for (int i = 0; i <= G2.Rows.Count - 1; i++)
                {
                    l++;

                    string T1 = G2.Rows[i][0].ToString();
                    System.Data.DataTable G3 = null;
                    if (textBox1.Text == "GD" || (listBox2.Items.Count == 0))
                    {
                        G3 = GetOPOR3Q(T1, fmLogin.LoginID.ToString());
                    }
                    else
                    {
                        G3 = GetOPOR3D(CS, T1, fmLogin.LoginID.ToString());
                    }

                    if (G3.Rows.Count > 0)
                    {
                        int h = -1;
                        for (int s = 0; s <= G3.Rows.Count - 1; s++)
                        {
                            h++;

                            if (h >= ROW)
                            {

                                l++;
                                h = 0;
                            }

                            UPOPORQ(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                        }
                    }
                }


            }
        }
        private void GD5Q(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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



            for (int i = 2; i <= iRowCnt; i++)
            {
                //Customer,Fab,Brand,SHIPDATE,InvoiceNo,SINo,PONo,Model,ver,PartNo,Grade,FREIGHTTERM,PAYMENT,LCNo,
                //Currency,Price,QTY,Amount,ETD,ETA,ShipCountry,ShipCity,TradeTerm

                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }

                //U_ACME_Dscription
                string Customer;
                string Fab;
                string Brand;
                string SHIPDATE;
                string InvoiceNo;
                string SINo;
                string PONo;
                string Model = "";
                string ver = "";
                string PartNo;
                string Grade;
                string FREIGHTTERM;
                string PAYMENT;
                string LCNo;
                string Currency;
                string Price;
                string QTY;
                string Amount;
                string ETD;
                string ETA;
                string ShipCountry;
                string ShipCity;
                string TradeTerm;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                Customer = range.Text.ToString().Trim();

                if (Customer != "")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    Fab = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    Brand = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    SHIPDATE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    InvoiceNo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    SINo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    PONo = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    string MV = range.Text.ToString().Trim();
                    int Y = MV.IndexOf(".");

                    if (Y != -1)
                    {
                        Model = MV.Substring(0, Y);
                        ver = MV.Substring(Y + 1, 1);

                    }


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    PartNo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                    range.Select();
                    Grade = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                    range.Select();
                    FREIGHTTERM = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    PAYMENT = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 13]);
                    range.Select();
                    LCNo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 14]);
                    range.Select();
                    Currency = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 15]);
                    range.Select();
                    Price = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 16]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 17]);
                    range.Select();
                    Amount = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 18]);
                    range.Select();
                    ETD = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 19]);
                    range.Select();
                    ETA = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 20]);
                    range.Select();
                    ShipCountry = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 21]);
                    range.Select();
                    ShipCity = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 22]);
                    range.Select();
                    TradeTerm = range.Text.ToString().Trim().ToUpper();

                    string CARDCODE = "";
                    if (comboBox1.Text == "AUOTV")
                    {
                        CARDCODE = "S0001-TV";
                    }
                    if (comboBox1.Text == "AUODD")
                    {
                        CARDCODE = "S0001-DD";
                    }
                    if (comboBox1.Text == "AUOGD")
                    {
                        CARDCODE = "S0001-GD";
                    }
                    if (comboBox1.Text == "達擎GD")
                    {
                        CARDCODE = "S0623-GD";
                    }
                    if (comboBox1.Text == "達擎PID")
                    {
                        CARDCODE = "S0623-PID";
                    }

                    try
                    {


                        ADDOPORQ(Customer, Fab, Brand, SHIPDATE, InvoiceNo, SINo, PONo, Model, ver, PartNo, Grade, FREIGHTTERM, PAYMENT, LCNo, Currency, Price, QTY, Amount, ETD, ETA, ShipCountry, ShipCity, TradeTerm, PONo, CARDCODE);



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
        private System.Data.DataTable MakeTableCombine()
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("INVOICE", typeof(string));
            dt.Columns.Add("INVOICE日期", typeof(string));
            dt.Columns.Add("SINO", typeof(string));
            dt.Columns.Add("SINO異常", typeof(string));
            dt.Columns.Add("輔助SINO", typeof(string));
            dt.Columns.Add("PARTNO", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("付款方式", typeof(string));
            dt.Columns.Add("採購報價單號", typeof(string));
            dt.Columns.Add("採購報價單號異常", typeof(string));
            dt.Columns.Add("採購報價料號", typeof(string));
            dt.Columns.Add("採購料號", typeof(string));
            dt.Columns.Add("採購報價單價", typeof(string));
            dt.Columns.Add("採購單價", typeof(string));
            dt.Columns.Add("TRADETERM", typeof(string));
            dt.Columns.Add("TRADETERM異常", typeof(string));

            dt.Columns.Add("SINO2", typeof(string));
            dt.Columns.Add("LCNO", typeof(string));
            dt.Columns.Add("產地", typeof(string));
            dt.Columns.Add("LINENUM", typeof(string));
            dt.Columns.Add("採購報價LINE", typeof(string));
            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("SINO3", typeof(string));
            dt.Columns.Add("ShipCity", typeof(string));
            dt.Columns.Add("採購報價備註", typeof(string));
            dt.Columns.Add("不覆蓋SI", typeof(string));
            //不覆蓋SI

            return dt;
        }

        private System.Data.DataTable MakeTableCombine3()
        {



            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("LINENUM", typeof(string));
            dt.Columns.Add("採購單過帳日期", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("項目料號", typeof(string));
            dt.Columns.Add("項目說明", typeof(string));
            dt.Columns.Add("PART NO", typeof(string));
            dt.Columns.Add("原廠發票項次", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("美金單價", typeof(decimal));
            dt.Columns.Add("原廠進貨匯率", typeof(string));
            dt.Columns.Add("台幣單價", typeof(decimal));
            dt.Columns.Add("台幣未稅金額", typeof(decimal));
            dt.Columns.Add("台幣稅額", typeof(decimal));
            dt.Columns.Add("台幣含稅金額", typeof(decimal));
            dt.Columns.Add("稅碼", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("SHIPPING工單號碼", typeof(string));
            dt.Columns.Add("原廠INVOICE", typeof(string));
            dt.Columns.Add("INVOICE日期", typeof(string));
            dt.Columns.Add("產地", typeof(string));
            dt.Columns.Add("付款方式", typeof(string));
            dt.Columns.Add("LCNO", typeof(string));


            return dt;
        }
        private void GD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK;
                string P1;
                string SITE;
                string ARRIVE;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                ITEMCODE = range.Text.ToString().Trim();

                if (ITEMCODE != "")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    LINE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    SITE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    ARRIVE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    AU97 = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
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

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 12]);
                    range.Select();
                    REMARK = range.Text.ToString().Trim();

                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {

                            string ITEM = "";
                            System.Data.DataTable F1 = GetITEMCODE(ITEMCODE, GRADE, AU97);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();
                            }
                            else
                            {
                                System.Data.DataTable F2 = GetITEMCODE2(ITEMCODE, GRADE, AU97);
                                if (F2.Rows.Count > 0)
                                {
                                    ITEM = F2.Rows[0][0].ToString();
                                }
                                else
                                {

                                    System.Data.DataTable F3 = GetITEMCODE3(ITEMCODE, GRADE, AU97);
                                    if (F3.Rows.Count > 0)
                                    {
                                        ITEM = F3.Rows[0][0].ToString();
                                    }
                                }
                            }

                            int GA = ARRIVE.IndexOf("回台");
                            if (GA != -1)
                            {
                                ARRIVE = "回台";
                            }

                            if (String.IsNullOrEmpty(ITEM))
                            {
                                ITEM = "ACME00004.00004";
                            }
                            if (LINE == "")
                            {
                                LINE = "1";
                            }
                            ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK, "S0001-DD", fmLogin.LoginID.ToString(), "", SITE, ARRIVE);
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

        public System.Data.DataTable GetITEMCODE(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 T0.ITEMCODE  FROM PDN1 T0 LEFT JOIN  OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;


                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                        s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (ITEMNAME LIKE '%" + ITEM1 + "%' OR ITEMNAME LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode + "%'   ");
            }


            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }


            sb.Append("  ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
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
        public System.Data.DataTable GETF1VILEN(string U_PARTNO, string U_GRADE, string U_VERSION, string DOCENTRY, string LEN)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE U_PARTNO=@U_PARTNO");
            sb.Append(" AND U_GRADE=@U_GRADE AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY");
            sb.Append("  AND SUBSTRING(T1.ITEMCODE,LEN(T1.ITEMCODE),1)=@LEN AND ISNULL(T1.LineStatus,'') <> 'C' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LEN", LEN));
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
        public System.Data.DataTable GETF1VI(string U_PARTNO, string U_GRADE, string U_VERSION, string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE U_PARTNO=@U_PARTNO");
            sb.Append(" AND U_GRADE=@U_GRADE AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY AND ISNULL(T1.LineStatus,'') <> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
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
        public System.Data.DataTable GETF1(string U_PARTNO, string U_GRADE, string U_VERSION, string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE SUBSTRING(U_PARTNO,0,11)=@U_PARTNO");
            sb.Append(" AND U_GRADE=@U_GRADE AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY AND ISNULL(T1.LineStatus,'') <> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
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
        public System.Data.DataTable GETF1F(string U_PARTNO, string U_GRADE, string U_VERSION, string DOCENTRY, decimal PP)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE SUBSTRING(U_PARTNO,0,11)=@U_PARTNO");
            sb.Append(" AND U_GRADE=@U_GRADE AND T1.PRICE=@PP AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY AND ISNULL(T1.LineStatus,'') <> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@PP", PP));
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
        public System.Data.DataTable GETF1FS(string U_PARTNO, string U_VERSION, string DOCENTRY, decimal PP)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE SUBSTRING(U_PARTNO,0,11)=@U_PARTNO");
            sb.Append("  AND T1.PRICE=@PP AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY AND ISNULL(T1.LineStatus,'') <> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));

            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@PP", PP));
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
        public System.Data.DataTable GETF1FSIT(string U_PARTNO, string DOCENTRY, decimal PP)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE U_PARTNO=@U_PARTNO");
            sb.Append("  AND T1.PRICE=@PP AND T1.DOCENTRY=@DOCENTRY AND ISNULL(T1.LineStatus,'') <> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));


            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@PP", PP));
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
        public System.Data.DataTable GETKIT(string U_PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE U_PARTNO=@U_PARTNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
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
        public System.Data.DataTable GETF1FS2(string U_PARTNO, string U_VERSION, string DOCENTRY, decimal PP)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM,T1.U_MEMO MEMO     FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE U_PARTNO=@U_PARTNO");
            sb.Append("  AND T1.PRICE=@PP AND U_VERSION=@U_VERSION AND T1.DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));

            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@PP", PP));
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


        public System.Data.DataTable GETF2(string U_PARTNO, string U_GRADE, string U_VERSION)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,0,11)=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
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
        public System.Data.DataTable GETF2P(string U_PARTNO, string U_GRADE, string U_VERSION)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE U_PARTNO=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION ");
            //sb.Append("  WHERE U_PARTNO=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION AND ITEMCODE=@ITEMCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            // command.Parameters.Add(new SqlParameter("@ITE//MCODE", ITEMCODE));

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

        public System.Data.DataTable GETF2P2(string U_PARTNO, string U_GRADE, string U_VERSION, string LEN)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE U_PARTNO=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION ");
            sb.Append("  AND SUBSTRING(ITEMCODE,LEN(ITEMCODE),1)=@LEN");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@LEN", LEN));

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

        public System.Data.DataTable GETF2P3(string U_PARTNO, string U_GRADE, string U_VERSION, string LEN, string VER)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE U_PARTNO=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION ");
            sb.Append("  AND SUBSTRING(ITEMCODE,LEN(ITEMCODE),1)=@LEN AND SUBSTRING(ITEMCODE,CHARINDEX('.', ITEMCODE)+2,3)=@VER");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@LEN", LEN));
            command.Parameters.Add(new SqlParameter("@VER", VER));
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
        public System.Data.DataTable GETF2P4(string U_PARTNO, string U_GRADE, string U_VERSION, string VER)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.ITEMCODE  FROM OITM T0");
            sb.Append("  WHERE U_PARTNO=@U_PARTNO AND CASE U_GRADE WHEN 'NN' THEN 'N' ELSE U_GRADE END  =@U_GRADE AND U_VERSION=@U_VERSION ");
            sb.Append("  AND SUBSTRING(ITEMCODE,LEN(ITEMCODE),1)='M' AND SUBSTRING(ITEMCODE,CHARINDEX('.', ITEMCODE)+2,3)=@VER");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@VER", VER));
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

        public System.Data.DataTable GETF1H(string U_PARTNO, string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.PRICE,T1.LINENUM    FROM OITM T0");
            sb.Append(" LEFT JOIN PQT1 T1 ON (T0.ITEMCODE=T1.ITEMCODE)   WHERE SUBSTRING(U_PARTNO,0,11)=@U_PARTNO");
            sb.Append(" AND T1.DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));

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
        public System.Data.DataTable GETF1HH(string U_PARTNO, string U_VERSION, string U_GRADE)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" 			  select ITEMCODE from oitm where U_PARTNO=@U_PARTNO and U_VERSION=@U_VERSION AND U_GRADE=@U_GRADE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            command.Parameters.Add(new SqlParameter("@U_VERSION", U_VERSION));
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
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
        public static System.Data.DataTable GETF3(string DOCENTRY, string ITEMCODE, string LINENUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM Shipping_Item WHERE DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE AND LINENUM=@LINENUM");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        public static System.Data.DataTable GETF4(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM Shipping_Item WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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


        public static System.Data.DataTable GETF6(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DocNum FROM LcInstro WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable GetITEMCODEGD(string MODEL, string U_GRADE, string VERSION)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
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
        public static System.Data.DataTable GetOPOR(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_OPOR WHERE USERS=@USERS ORDER BY DOCENTRY,DOC,ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR6(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR WHERE USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR6IN(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(QTY) QTY,SUM(AMT) AMT FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GERPQT(string DOCENTRY, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LINENUM FROM PQT1 WHERE DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
            SqlConnection MyConnection = globals.shipConnection;
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
        public static System.Data.DataTable GetOPOR2(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPOR WHERE USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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

        public static System.Data.DataTable GetOPOR2Q(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPORQ WHERE USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR2D(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPOR WHERE  ID IN ( " + cs + ") AND USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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

        public static System.Data.DataTable GetOPOR2DQ(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  DISTINCT  DOCENTRY FROM AP_OPORQ WHERE  ID IN ( " + cs + ") AND USERS=@USERS AND DOCENTRY >0 ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR WHERE USERS=@USERS AND DOC >0 ORDER BY DOC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4Q(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPORQ WHERE USERS=@USERS AND DOC >0 ORDER BY DOC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4D(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS AND DOC >0 ORDER BY DOC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR4DQ(string cs, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT  DOC FROM AP_OPOR WHERE ID IN ( " + cs + ") AND USERS=@USERS AND DOC >0 ORDER BY DOC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public static System.Data.DataTable GetOPOR3(string DOCENTRY, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPOR WHERE DOCENTRY=@DOCENTRY AND DOCENTRY >0 AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR3Q(string DOCENTRY, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPORQ WHERE DOCENTRY=@DOCENTRY AND DOCENTRY >0 AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR3D(string cs, string DOCENTRY, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_OPOR WHERE DOCENTRY=@DOCENTRY AND ID IN ( " + cs + ")  AND DOCENTRY >0 AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5(string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE DOC=@DOC AND DOC >0  AND USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5Q(string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPORQ WHERE DOC=@DOC AND DOC >0  AND USERS=@USERS  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5ID(string cs, string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR WHERE  DOC=@DOC AND ID IN ( " + cs + ")  AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetOPOR5IDQ(string cs, string DOC, string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPORQ WHERE  DOC=@DOC AND ID IN ( " + cs + ")  AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public static System.Data.DataTable GetSI1(string cs)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [SITE],ARRIVE 到貨地 FROM AP_OPOR 　WHERE USERS=@USERS");
            if (!String.IsNullOrEmpty(cs))
            {
                sb.Append(" AND ID IN ( " + cs + ") ");
            }
            sb.Append("  GROUP BY [SITE],ARRIVE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public static System.Data.DataTable GetSI2(string SITE, string ARRIVE, string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM AP_OPOR 　WHERE USERS=@USERS AND [SITE]=@SITE AND ARRIVE =@ARRIVE ");
            if (!String.IsNullOrEmpty(cs))
            {
                sb.Append(" AND ID IN ( " + cs + ") ");
            }
            sb.Append(" ORDER BY SAPDOC,LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@SITE", SITE));
            command.Parameters.Add(new SqlParameter("@ARRIVE", ARRIVE));
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
        public System.Data.DataTable GetITEMCODE2(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 T0.ITEMCODE  FROM PDN1 T0 LEFT JOIN  OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            sb.Append(" AND  U_TMODEL +'.'+U_VERSION  like '%" + ItemCode + "%'  ");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;


                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                        s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (   U_TMODEL +'.'+U_VERSION LIKE '%" + ITEM1 + "%' OR U_TMODEL +'.'+U_VERSION LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND U_TMODEL +'.'+U_VERSION  LIKE '%" + ItemCode + "%'   ");
            }
            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }

            sb.Append("  ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));

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

        public System.Data.DataTable GetITEMCODE3(string ItemCode, string U_GRADE, string PARTNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 ITEMCODE  FROM   OITM ");
            sb.Append("  WHERE SUBSTRING(U_PARTNO,1,2)=@PARTNO");
            int F1 = ItemCode.IndexOf("/");
            int F2 = ItemCode.IndexOf("_");
            int F3 = ItemCode.IndexOf("(");
            if (F1 != -1)
            {
                string[] arrurl = ItemCode.Split(new Char[] { '/' });
                StringBuilder sbs = new StringBuilder();
                string ITEM1 = "";
                string ITEM2 = "";
                string s1 = "";
                int G = 0;
                foreach (string ESi in arrurl)
                {
                    G++;


                    int g = ESi.IndexOf(".");
                    if (g != -1)
                    {
                        s1 = ESi.Substring(0, g);
                    }
                    if (G == 1)
                    {
                        ITEM1 = ESi;
                    }
                    if (G == 2)
                    {
                        ITEM2 = s1 + "." + ESi;
                    }
                }
                sb.Append(" AND (ITEMNAME LIKE '%" + ITEM1 + "%' OR ITEMNAME LIKE '%" + ITEM2 + "%')   ");
            }
            else if (F2 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F2) + "%'   ");
            }
            else if (F3 != -1)
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode.Substring(0, F3) + "%'   ");
            }
            else
            {
                sb.Append(" AND ITEMNAME LIKE '%" + ItemCode + "%'   ");
            }


            if (U_GRADE == "PN")
            {
                sb.Append(" AND ( U_GRADE ='N' OR  U_GRADE ='P')");
            }
            else if (U_GRADE == "N")
            {
                sb.Append(" AND ( U_GRADE ='NN' OR  U_GRADE ='N')");
            }
            else
            {
                sb.Append(" AND U_GRADE=@U_GRADE ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_GRADE", U_GRADE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
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
        public void DELOPOR(string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_OPOR WHERE USERS=@USERS ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public void DELOPORQ(string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_OPORQ WHERE USERS=@USERS ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        public void ADDOPOR(int DOCENTRY, string ITEMCODE, decimal QTY, decimal PRICE, decimal AMT, string REMARK, string CARDCODE, string USERS, string OCARDCODE, string SITE, string ARRIVE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPOR(DOCENTRY,ITEMCODE,QTY,PRICE,AMT,REMARK,CARDCODE,USERS,OCARDCODE,SITE,ARRIVE) values(@DOCENTRY,@ITEMCODE,@QTY,@PRICE,@AMT,@REMARK,@CARDCODE,@USERS,@OCARDCODE,@SITE,@ARRIVE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@OCARDCODE", OCARDCODE));
            command.Parameters.Add(new SqlParameter("@SITE", SITE));
            command.Parameters.Add(new SqlParameter("@ARRIVE", ARRIVE));
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
        public void ADDOPORQ(string Customer, string Fab, string Brand, string SHIPDATE, string InvoiceNo, string SINo, string PONo, string Model, string ver, string PartNo, string Grade, string FREIGHTTERM, string PAYMENT, string LCNo, string Currency, string Price, string QTY, string Amount, string ETD, string ETA, string ShipCountry, string ShipCity, string TradeTerm, string DOCENTRY, string CARDCODE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPORQ(Customer,Fab,Brand,SHIPDATE,InvoiceNo,SINo,PONo,Model,ver,PartNo,Grade,FREIGHTTERM,PAYMENT,LCNo,Currency,Price,QTY,Amount,ETD,ETA,ShipCountry,ShipCity,TradeTerm,USERS,DOCENTRY,CARDCODE) values(@Customer,@Fab,@Brand,@SHIPDATE,@InvoiceNo,@SINo,@PONo,@Model,@ver,@PartNo,@Grade,@FREIGHTTERM,@PAYMENT,@LCNo,@Currency,@Price,@QTY,@Amount,@ETD,@ETA,@ShipCountry,@ShipCity,@TradeTerm,@USERS,@DOCENTRY,@CARDCODE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Customer", Customer));
            command.Parameters.Add(new SqlParameter("@Fab", Fab));
            command.Parameters.Add(new SqlParameter("@Brand", Brand));
            command.Parameters.Add(new SqlParameter("@SHIPDATE", SHIPDATE));
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            command.Parameters.Add(new SqlParameter("@SINo", SINo));
            command.Parameters.Add(new SqlParameter("@PONo", PONo));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@ver", ver));
            command.Parameters.Add(new SqlParameter("@PartNo", PartNo));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@FREIGHTTERM", FREIGHTTERM));
            command.Parameters.Add(new SqlParameter("@PAYMENT", PAYMENT));
            command.Parameters.Add(new SqlParameter("@LCNo", LCNo));
            command.Parameters.Add(new SqlParameter("@Currency", Currency));
            command.Parameters.Add(new SqlParameter("@Price", Price));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@ETD", ETD));
            command.Parameters.Add(new SqlParameter("@ETA", ETA));
            command.Parameters.Add(new SqlParameter("@ShipCountry", ShipCountry));
            command.Parameters.Add(new SqlParameter("@ShipCity", ShipCity));
            command.Parameters.Add(new SqlParameter("@TradeTerm", TradeTerm));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));

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

        public void UPOPQT(int TrgetEntry, int DOCENTRY, int LINENUM)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE PQT1 SET TargetType=22,TrgetEntry=@TrgetEntry WHERE DOCENTRY=@DOCENTRY  AND LINENUM=@LINENUM", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TrgetEntry", TrgetEntry));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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

        public void UPOPQT2(int TrgetEntry, int DOCENTRY, int LINENUM)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE PQT1 SET TargetType=22,TrgetEntry=@TrgetEntry WHERE DOCENTRY=@DOCENTRY  AND LINENUM=@LINENUM", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TrgetEntry", TrgetEntry));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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

        public void UPOPORF(int BASEENTRY, int BASELINE, int DOCENTRY, int LINENUM)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);

            SqlCommand command = new SqlCommand("UPDATE POR1 SET BASETYPE=540000006,BASEENTRY=@BASEENTRY,BASEREF=@BASEENTRY,BASELINE=@BASELINE WHERE DOCENTRY=@DOCENTRY  AND LINENUM=@LINENUM", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BASEENTRY", BASEENTRY));
            command.Parameters.Add(new SqlParameter("@BASELINE", BASELINE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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

        public void UPOPORF2(int DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append("               UPDATE PQT1 SET OpenCreQty =T0.Quantity -T1.QTY,OpenQty =T0.Quantity -T1.QTY, ");
            sb.Append("              LineStatus =CASE  WHEN (T0.Quantity -T1.QTY) <= 0 THEN 'C'  END   FROM PQT1 T0 ");
            sb.Append("              ,(SELECT  BASEENTRY,BASELINE,SUM(Quantity) QTY FROM POR1 T0");
            sb.Append("			   LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE BASEENTRY=@DOCENTRY AND T1.CANCELED<>'Y'");
            sb.Append("              GROUP BY  BASEENTRY,BASELINE) T1 WHERE T0.DOCENTRY=T1.BASEENTRY AND T0.LINENUM=T1.BASELINE ");
            sb.Append("              AND T0.DOCENTRY=@DOCENTRY  AND  ISNULL(T0.LineStatus,'')<> 'C' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        public void UPOPOR(int DOC, int LINENUM, int ID, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPOR SET DOC=@DOC,LINENUM=@LINENUM WHERE ID=@ID AND USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public void UPOPORQ(int DOC, int LINENUM, int ID, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPORQ SET DOC=@DOC,LINENUM=@LINENUM WHERE ID=@ID AND USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public void UPOPORD(int DOC, int LINENUM, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPOR SET DOC=@DOC,LINENUM=@LINENUM WHERE USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));

            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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

        public void UPOPORDQ(int DOC, int LINENUM, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_OPORQ SET DOC=@DOC,LINENUM=@LINENUM WHERE USERS=@USERS  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));

            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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

            oCompany.CompanyDB = FA;
            oCompany.UserName = "A02";
            oCompany.Password = "6500";
            int result = oCompany.Connect();
            if (result == 0)
            {


                if (listBox1.Items.Count == 0)
                {
                    D1("");

                    System.Data.DataTable G2 = GetOPOR4(fmLogin.LoginID.ToString());

                    if (G2.Rows.Count > 0)
                    {

                        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                        {
                            SAPbobsCOM.Documents oPURCH = null;
                            oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations);

                            string T1 = G2.Rows[n][0].ToString();

                            System.Data.DataTable G3 = GetOPOR5(T1, fmLogin.LoginID.ToString());

                            if (G3.Rows.Count > 0)
                            {

                                oPURCH.CardCode = G3.Rows[0]["CARDCODE"].ToString();
                                oPURCH.DocCurrency = "USD";
                                oPURCH.VatPercent = 0;
                                oPURCH.RequriedDate = DateTime.Now;
                                System.Data.DataTable G1 = GetOPOR7();
                                if (G1.Rows.Count > 0)
                                {
                                    oPURCH.DocRate = Convert.ToDouble(G1.Rows[0][0]);
                                }
                                System.Data.DataTable G7 = GetOPORH();
                                if (G7.Rows.Count > 0)
                                {
                                    oPURCH.DocumentsOwner = Convert.ToInt32(G7.Rows[0][0]);
                                    oPURCH.SalesPersonCode = Convert.ToInt32(G7.Rows[0][1]);
                                }


                                oPURCH.Comments = "OA 45天//LC 45天,Z/P/N保固18個月(保固以INVOICE DATE起算)";
                                for (int s = 0; s <= G3.Rows.Count - 1; s++)
                                {
                                    string ITEMCODE = G3.Rows[s]["ITEMCODE"].ToString();
                                    if (ITEMCODE == "KTCAU43TX.00102") 
                                    {
                                        ITEMCODE = "KTCAU43TX.00162";
                                    }
                                    System.Data.DataTable tt = GetDESC(ITEMCODE);
                                    if (tt.Rows.Count == 0)
                                    {
                                        ITEMCODE = "ACME00004.00004";
                                    }
                                    string OCARDCODE = G3.Rows[s]["OCARDCODE"].ToString();
                                    double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                                    double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                                    oPURCH.Lines.WarehouseCode = "TW017";
                                    oPURCH.Lines.ItemCode = ITEMCODE;
                                    oPURCH.Lines.Quantity = QTY;
                                    oPURCH.Lines.UnitPrice = PRICE;
                                    oPURCH.Lines.Price = PRICE;
                                    oPURCH.Lines.VatGroup = "AP0%";
                                    oPURCH.Lines.Currency = "USD";
                                    oPURCH.Lines.UserFields.Fields.Item("U_ACME_Dscription").Value = "OA 45 DAYS";
                                    oPURCH.Lines.UserFields.Fields.Item("U_MEMO").Value = G3.Rows[s]["REMARK"].ToString();
                                    if (ITEMCODE == "ACME00004.00004")
                                    {
                                        oPURCH.Lines.ItemDescription = OCARDCODE;
                                    }
                                    if (textBox1.Text == "PID" || textBox1.Text == "TV")
                                    {
                                        DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
                                        oPURCH.DocDate = LastDay;
                                        oPURCH.Lines.ShipDate = LastDay;
                                    }
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
                                    MessageBox.Show("上傳成功 採購報價單號 : " + OWTR);


                                }
                            }
                        }

                    }
                }
                //else
                //{

                //    ArrayList al = new ArrayList();

                //    for (int i2 = 0; i2 <= listBox1.Items.Count - 1; i2++)
                //    {
                //        al.Add(listBox1.Items[i2].ToString());
                //    }
                //    StringBuilder sb = new StringBuilder();



                //    foreach (string v in al)
                //    {
                //        sb.Append("'" + v + "',");
                //    }

                //    sb.Remove(sb.Length - 1, 1);

                //    D1(sb.ToString());

                //    System.Data.DataTable G2 = GetOPOR4D(sb.ToString(), fmLogin.LoginID.ToString());

                //    if (G2.Rows.Count > 0)
                //    {

                //        for (int n = 0; n <= G2.Rows.Count - 1; n++)
                //        {
                //            SAPbobsCOM.Documents oPURCH = null;
                //            oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                //            string T1 = G2.Rows[n][0].ToString();
                //            System.Data.DataTable G3 = GetOPOR5ID(sb.ToString(), T1, fmLogin.LoginID.ToString());

                //            if (G3.Rows.Count > 0)
                //            {
                //                oPURCH.CardCode = G3.Rows[0]["CARDCODE"].ToString();
                //                oPURCH.DocCurrency = "USD";
                //                oPURCH.VatPercent = 0;
                //                oPURCH.RequriedDate = DateTime.Now;
                //                System.Data.DataTable G1 = GetOPOR7();
                //                if (G1.Rows.Count > 0)
                //                {
                //                    oPURCH.DocRate = Convert.ToDouble(G1.Rows[0][0]);
                //                }
                //                System.Data.DataTable G7 = GetOPORH();
                //                if (G7.Rows.Count > 0)
                //                {
                //                    oPURCH.DocumentsOwner = Convert.ToInt32(G7.Rows[0][0]);
                //                    oPURCH.SalesPersonCode = Convert.ToInt32(G7.Rows[0][1]);
                //                }
                //                oPURCH.Comments = "OA 45天//LC 45天,Z/P/N保固18個月(保固以INVOICE DATE起算)";
                //                for (int s = 0; s <= G3.Rows.Count - 1; s++)
                //                {
                //                    string ITEMCODE = G3.Rows[s]["ITEMCODE"].ToString();
                //                    string OCARDCODE = G3.Rows[s]["OCARDCODE"].ToString();
                //                    double QTY = Convert.ToDouble(G3.Rows[s]["QTY"]);
                //                    double PRICE = Convert.ToDouble(G3.Rows[s]["PRICE"]);
                //                    oPURCH.Lines.WarehouseCode = "TW017";
                //                    oPURCH.Lines.ItemCode = ITEMCODE;
                //                    oPURCH.Lines.Quantity = QTY;
                //                    oPURCH.Lines.Price = PRICE;
                //                    oPURCH.Lines.VatGroup = "AP0%";
                //                    oPURCH.Lines.Currency = "USD";
                //                    oPURCH.Lines.UserFields.Fields.Item("U_ACME_Dscription").Value = "OA 45 DAYS";
                //                    if (ITEMCODE == "ACME00004.00004")
                //                    {
                //                        oPURCH.Lines.ItemDescription = OCARDCODE;
                //                    }
                //                    if (textBox1.Text == "PID" || textBox1.Text == "TV")
                //                    {
                //                        DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
                //                        oPURCH.DocDate = LastDay;
                //                        oPURCH.Lines.ShipDate = LastDay;
                //                    }
                //                    oPURCH.Lines.Add();

                //                }


                //                int res = oPURCH.Add();
                //                if (res != 0)
                //                {
                //                    MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                //                }
                //                else
                //                {
                //                    System.Data.DataTable G4 = GetDI4();
                //                    string OWTR = G4.Rows[0][0].ToString();
                //                    MessageBox.Show("上傳成功 採購單號 : " + OWTR);

                //                    System.Data.DataTable GG3 = GetOPOR5ID(sb.ToString(), T1, fmLogin.LoginID.ToString());

                //                    if (GG3.Rows.Count > 0)
                //                    {
                //                        for (int j = 0; j <= GG3.Rows.Count - 1; j++)
                //                        {

                //                            string LINENUM = GG3.Rows[j]["LINENUM"].ToString();
                //                            string REMARK = GG3.Rows[j]["REMARK"].ToString();
                //                            UPDATEOPOR(OWTR, LINENUM, REMARK);

                //                            UPDATEAPOPORD(sb.ToString(), T1, LINENUM, OWTR);
                //                        }

                //                    }
                //                }
                //            }
                //        }
                //    }
                //}




            }
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }


            System.Data.DataTable GG1 = GetOPOR(fmLogin.LoginID.ToString());
            dataGridView1.DataSource = GG1;
        }
        public static System.Data.DataTable GetOPORH()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT T0.EmpID ,T1.SlpCode,LASTNAME+firstName  FNAME   FROM OHEM T0");
            sb.Append("  LEFT JOIN OSLP T1 ON (T0.lastName +T0.firstName =T1.SlpName)");
            sb.Append("  WHERE T0.homeTel=@homeTel");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@homeTel", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetDI4Q(string u_shipping_no)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY)   FROM POR1 WHERE  u_shipping_no=@u_shipping_no");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@u_shipping_no", u_shipping_no));
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
        public System.Data.DataTable GetODRF()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCNUM)  DOC FROM ODRF WHERE ObjType =20");
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

        public System.Data.DataTable GetDI4Q2(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY,LINENUM,U_PAY,U_CUSTITEMCODE  FROM POR1 WHERE DOCENTRY=@DOCENTRY");
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
        public System.Data.DataTable GetDI4()
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OPQT");
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

        public System.Data.DataTable GetDESC(string ITEMCODE)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMNAME FROM OITM WHERE ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
            sb.Append(" UPDATE PQT1 SET U_MEMO=@U_MEMO,U_ACME_Dscription='OA 45 days'  WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

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

        private void UPDATEAPOPOR(string DOC, string LINENUM, string SAPDOC)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE AP_OPOR SET SAPDOC=@SAPDOC WHERE DOC=@DOC AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
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

        private void UPDATEAPOPORD(string cs, string DOC, string LINENUM, string SAPDOC)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE AP_OPOR SET SAPDOC=@SAPDOC WHERE DOC=@DOC AND ID IN ( " + cs + ")  AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
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
        //ID IN ( " + cs + ") 
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

                    System.Data.DataTable GG2 = GetOPOR6IN(sb.ToString(), fmLogin.LoginID.ToString());
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
        private void GD6(string ExcelFile)
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
                            System.Data.DataTable F1 = GetITEMCODEGD(ITEMCODE, GRADE, VER);
                            if (F1.Rows.Count > 0)
                            {
                                ITEM = F1.Rows[0][0].ToString();
                            }
                            else
                            {
                                ITEM = "ACME00004.00004";
                            }

                            if (String.IsNullOrEmpty(LINE))
                            {
                                LINE = "1";
                            }

                            ADDOPOR(3, ITEM, QTY, PRICE, AMT, REMARK1 + " " + REMARK2, "S0623-GD", fmLogin.LoginID.ToString(), ITEMCODE + "." + VER.Substring(0, 1), "", "");
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
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR(fmLogin.LoginID.ToString());
                GD6(opdf.FileName);

                textBox1.Text = "GD";

                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());

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

                                UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }



                    System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                    dataGridView1.DataSource = G1;
                    System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                    label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                    label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

                }
            }
        }



        public void AddSHIPMAIN2(string ShippingCode, string ShippingCode2)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("INSERT INTO SHIPPING_MAIN (ShippingCode,CardName,CardCode,TradeCondition,receivePlace,shipment,goalPlace,unloadCargo,receiveDay,boardCountNo,ADD10,CreateName)  select @shippingcode2,CardName,CardCode,TradeCondition,receivePlace,shipment,goalPlace,unloadCargo,receiveDay,boardCountNo,ADD10,CreateName from SHIPPING_MAIN where shippingcode=@shippingcode", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ShippingCode2", ShippingCode2));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void AddSHIPMAIN3(string ShippingCode, string ShippingCode2, string DocNum)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("INSERT INTO LcInstro (ShippingCode,DocNum,Memo,WHSCODE,RUSH,ITEMS,DLC)  select @shippingcode2,@DocNum,Memo,WHSCODE,RUSH,ITEMS,DLC from LcInstro where shippingcode=@shippingcode", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ShippingCode2", ShippingCode2));
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void AddSHIPITEM(string ShippingCode, int SeqNo, string Docentry, int linenum, string ItemRemark, string ItemCode, string Dscription, int Quantity, decimal ItemPrice, decimal ItemAmount, string Remark)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_ITEM(ShippingCode,SeqNo,Docentry,linenum,ItemRemark,ItemCode,Dscription,Quantity,ItemPrice,ItemAmount,Remark) values(@ShippingCode,@SeqNo,@Docentry,@linenum,@ItemRemark,@ItemCode,@Dscription,@Quantity,@ItemPrice,@ItemAmount,@Remark)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Dscription", Dscription));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@ItemPrice", ItemPrice));
            command.Parameters.Add(new SqlParameter("@ItemAmount", ItemAmount));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void AddSHIPITEM2(string shippingcode, string docnum, string seqno, string Docentry, string ItemCode, string Dscription, string Quantity, decimal ItemPrice, decimal ItemAmount, string LC, string linenum)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into LcInstro1(shippingcode,docnum,seqno,Docentry,ItemCode,Dscription,Quantity,ItemPrice,ItemAmount,LC,linenum) values(@shippingcode,@docnum,@seqno,@Docentry,@ItemCode,@Dscription,@Quantity,@ItemPrice,@ItemAmount,@LC,@linenum)", Connection);
            command.CommandType = CommandType.Text;
            //shippingcode,docnum,seqno,Docentry,ItemCode,Dscription,Quantity,ItemPrice,ItemAmount
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingcode));
            command.Parameters.Add(new SqlParameter("@docnum", docnum));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Dscription", Dscription));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@ItemPrice", ItemPrice));
            command.Parameters.Add(new SqlParameter("@ItemAmount", ItemAmount));
            command.Parameters.Add(new SqlParameter("@LC", LC));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));


            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void DELINV()
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand(" DELETE AP_OPORQINV WHERE USERS=@USERS", Connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }


        public void ADDINV(string INV, string INVPRICE, string INVITEM)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into AP_OPORQINV(INV,INVPRICE,USERS,INVITEM) values(@INV,@INVPRICE,@USERS,@INVITEM)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@INVPRICE", INVPRICE));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@INVITEM", INVITEM));
            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public void DELITEM(string ShippingCode)
        {
            SqlConnection Connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("DELETE SHIPPING_ITEM WHERE ShippingCode=@ShippingCode DELETE LcInstro1 WHERE ShippingCode=@ShippingCode", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));


            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }
        public System.Data.DataTable GSHIP(string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT *  FROM OPQT  WHERE DOCENTRY=@DOCENTRY");

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
        public System.Data.DataTable GOPQT(string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT *  FROM OPQT  WHERE DOCENTRY=@DOCENTRY");

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
        public System.Data.DataTable GetOPQT(string DOCENTRY)
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.CARDCODE,T0.CardName,T1.ITEMCODE,T1.PRICE,CAST(T1.Quantity AS INT) QTY  FROM OPQT T0");
            sb.Append("  LEFT JOIN PQT1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("  WHERE T0.DOCENTRY=@DOCENTRY");


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
        public System.Data.DataTable GETITEMNAME(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMNAME FROM OITM WHERE ITEMCODE=@ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetOPQTF(string USERS)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT *, SHIPDATE 出貨日期,SUBSTRING(PARTNO,0,11)  PARTNO2,SUBSTRING(PARTNO,1,2) PARTNO3,SUBSTRING(PARTNO,10,3) PARTNO4  FROM AP_OPORQ WHERE USERS=@USERS order by sino");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public System.Data.DataTable GetAPPLE()
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT T0.DOCENTRY 採購單號, T1.LINENUM,Convert(varchar(10),T0.DOCDATE,111)  採購單過帳日期,T1.ITEMCODE 項目料號,T1.Dscription 項目說明 ");
            sb.Append("               ,T2.U_PARTNO PARTNO,CAST(T1.QUANTITY AS INT) 數量,T1.PRICE 單價,T1.U_ACME_INV 原廠INVOICE,Convert(varchar(10),T1.U_ACME_SHIPDAY,111)  INVOICE日期,T1.U_BASE_DOC LCNO ");
            sb.Append("			   ,T1.U_ACME_KIND 產地,T1.U_SHIPPING_NO 工單號碼,T1.VatGroup 稅碼,T1.U_ACME_DSCRIPTION 付款方式,T0.CARDCODE 廠商編號");
            sb.Append("              FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("               LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append("               WHERE       ISNULL(T1.LineStatus,'') <> 'C' ");
       //     sb.Append("        WHERE T0.DOCENTRY=44451 ");
            
            if (textBox2.Text != "")
            {
                sb.Append("   AND T0.DOCENTRY =@DOCENTRY");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" AND Convert(varchar(8), T0.DOCDATE, 112) =@DOCDATE");
            }
            if (textBox3.Text != "")
            {
                sb.Append(" AND T0.CARDNAME =@CARDNAME");
            }
            if (comboBox2.Text != "")
            {
                sb.Append(" AND T1.U_ACME_INV  =@U_ACME_INV");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CARDNAME", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", comboBox2.Text));
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
        public System.Data.DataTable GetAPPLECOMB()
        {
            SqlConnection MyConnection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT DISTINCT T1.U_ACME_INV 原廠INVOICE ");
            sb.Append("              FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("               LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append("               WHERE      ISNULL(T1.LineStatus,'') <> 'C'  ");
          //  sb.Append("        WHERE T0.DOCENTRY=44451 ");
            if (textBox2.Text != "")
            {
                sb.Append("   AND T0.DOCENTRY =@DOCENTRY");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" AND Convert(varchar(8), T0.DOCDATE, 112) =@DOCDATE");
            }
            if (textBox3.Text != "")
            {
                sb.Append(" AND T0.CARDNAME =@CARDNAME");
            }
            if (comboBox2.Text != "")
            {
                sb.Append(" AND T1.U_ACME_INV  =@U_ACME_INV");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CARDNAME", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", comboBox2.Text));
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
        public System.Data.DataTable GETA2()
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();


            sb.Append("          SELECT DISTINCT INV FROM AP_OPORQINV WHERE USERS=@USERS ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GETA2INV(string INV)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();


            sb.Append("          SELECT  INVITEM FROM AP_OPORQINV WHERE USERS=@USERS AND ISNULL(INVITEM,'') <> '' AND  INV=@INV ORDER BY CAST(INVITEM AS INT)  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@INV", INV));
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
        public System.Data.DataTable GETA3()
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();


            sb.Append(" Declare @N1 varchar(200) ");
            sb.Append(" select @N1 =SUBSTRING(COALESCE(@N1 + '/',''),0,190) + INVPRICE ");
            sb.Append(" from   (SELECT    INVPRICE    FROM AP_OPORQINV WHERE USERS=@USERS ) pc");
            sb.Append(" SELECT @N1 S");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

        public System.Data.DataTable GetAPPLE2(DateTime DDATE)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 ISNULL(HFAR,HBUY) FROM ACMESQLSP.DBO.WH_HAIGUAN");
            sb.Append(" WHERE HYEAR=YEAR(@DDATE) AND HMON=MONTH(@DDATE)");
            sb.Append(" AND DAY(@DDATE) BETWEEN HDAY2 AND HDAY3");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DDATE", DDATE));
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
        public System.Data.DataTable GETLC(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DocNum FROM LcInstro WHERE SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable GetOPQTF2(string USERS, string SINO)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT * FROM AP_OPORQ WHERE USERS=@USERS AND SINO=@SINO");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@SINO", SINO));
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

        public System.Data.DataTable GetOPCH1(DateTime date)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT TOP 1 ISNULL(HFAR,HBUY) FROM ACMESQLSP.DBO.WH_HAIGUAN");
            sb.Append("  WHERE HYEAR=YEAR(@date) AND HMON=MONTH(@date)");
            sb.Append("  AND DAY(@date) BETWEEN HDAY2 AND HDAY3");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@date", date));

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
        private void APOPOR_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "AUOTV";
            textBox6.Text = GetMenu.Day();
            //if (globals.GroupID.ToString().Trim() != "EEP")
            //{
            if (globals.DBNAME == "進金生")
            {
                strCn98 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                FA = "acmesql02";
            }
            //   }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPORQ(fmLogin.LoginID.ToString());
                GD5Q(opdf.FileName);
                DataRow dr = null;

                System.Data.DataTable dtCost = MakeTableCombine();
                System.Data.DataTable G1 = GetOPQTF(fmLogin.LoginID.ToString());
                int LI = 0;
                string DUP = "";
                for (int i = 0; i <= G1.Rows.Count - 1; i++)
                {
                    DataRow dd = G1.Rows[i];
                    dr = dtCost.NewRow();
                    string SINO = dd["SINO"].ToString();
                    if (DUP != "")
                    {
                        if (DUP == SINO)
                        {
                            LI = LI + 1;


                        }
                        else
                        {
                            LI = 1;
                        }
                    }
                    else
                    {
                        LI = 1;
                    }



                    dr["ID"] = dd["ID"].ToString();
                    dr["廠商編號"] = dd["CARDCODE"].ToString();
                    dr["出貨日期"] = dd["出貨日期"].ToString();
                    dr["LINENUM"] = LI.ToString();




                    DUP = SINO;
                    int N1 = 0;
                    int N2 = 0;
                    System.Data.DataTable G2 = GetOPQTF2(fmLogin.LoginID.ToString(), SINO);
                    System.Data.DataTable G3 = GETF4(SINO);
                    if (G2.Rows.Count > 0)
                    {
                        N1 = G2.Rows.Count;
                    }
                    if (G3.Rows.Count > 0)
                    {
                        N2 = G3.Rows.Count;
                    }
                    string PARTNOF = dd["PARTNO"].ToString();
                    //if (PARTNOF == "91.27M06.3UA")
                    //{
                    //    MessageBox.Show("s");
                    //}
                    string docnum = "";
                    System.Data.DataTable SI1 = GETF6(SINO);

                    if (SI1.Rows.Count > 0)
                    {
                        docnum = SI1.Rows[0][0].ToString();
                    }
                    dr["SINO"] = SINO;
                    dr["SINO2"] = docnum;
                    dr["SINO3"] = SINO;
                    dr["LCNO"] = dd["LCNO"].ToString();
                    dr["產地"] = dd["Fab"].ToString();
                    if (N1 != N2)
                    {
                        dr["SINO異常"] = "採購SI列數 : " + N2.ToString();

                    }
                    //Price
                    decimal PP = Convert.ToDecimal(dd["Price"]);

                    string PARTNO = dd["PARTNO2"].ToString();
                    string PARTNO3 = dd["PARTNO3"].ToString();
                    string PARTNO4 = dd["PARTNO4"].ToString();

                    string MODEL = dd["MODEL"].ToString();
                    string PONO = dd["PONO"].ToString();
                    string Fab = dd["Fab"].ToString();
                    string LW = "";

                    if (Fab == "M02" || Fab == "M11" || Fab == "L8B" || Fab == "X12")
                    {
                        LW = "1";
                    }
                    else
                    {
                        LW = "2";
                    }

                    dr["採購報價單號"] = PONO;
                    dr["PARTNO"] = PARTNO;
                    dr["QTY"] = dd["QTY"].ToString();
                    dr["TRADETERM"] = dd["TRADETERM"].ToString();
                    dr["付款方式"] = dd["PAYMENT"].ToString();
                    dr["ShipCity"] = dd["ShipCity"].ToString();
                    string GRADE = dd["GRADE"].ToString();
                    string VER = dd["VER"].ToString();


                    string 採購報價單號異常 = "";
                    string 採購報價料號 = "";
                    string 採購報價單價 = "";
                    string 採購報價備註 = "";
                    string 採購料號 = "";
                    string LINENUM = "";
                    //if (PARTNOF == "91.27M06.3UA")
                    //{
                    //    MessageBox.Show("s");
                    //}
                    System.Data.DataTable IOPQT = GOPQT(PONO);
                    System.Data.DataTable IOPQT2 = GSHIP(PONO);
                    if (IOPQT.Rows.Count == 0 || IOPQT2.Rows.Count == 0)
                    {
                        採購報價單號異常 = "異常";
                    }
                    dr["採購報價單號異常"] = 採購報價單號異常;
                    System.Data.DataTable ITEM1 = null;
                    int FE = 0;
                    ITEM1 = GETF1VILEN(PARTNOF, GRADE, VER, PONO, LW);
                    if (ITEM1.Rows.Count == 0)
                    {
                        ITEM1 = GETF1VI(PARTNOF, GRADE, VER, PONO);
                    }
                    if (ITEM1.Rows.Count == 0)
                    {
                        ITEM1 = GETF1(PARTNO, GRADE, VER, PONO);
                    }
                    if (ITEM1.Rows.Count == 0)
                    {
                        ITEM1 = GETF1F(PARTNO, GRADE, VER, PONO, PP);
                        if (ITEM1.Rows.Count > 0)
                        {
                            FE = 1;

                        }
                        if (ITEM1.Rows.Count == 0)
                        {
                            System.Data.DataTable ITEM2H = GETF1FS2(PARTNOF, VER, PONO, PP);
                            if (ITEM2H.Rows.Count == 1)
                            {
                                ITEM1 = ITEM2H;
                            }
                        }
                        if (ITEM1.Rows.Count == 0)
                        {
                            System.Data.DataTable ITEM2H = GETF1FS(PARTNO, VER, PONO, PP);
                            if (ITEM2H.Rows.Count == 1)
                            {
                                ITEM1 = ITEM2H;
                            }
                        }

                        if (ITEM1.Rows.Count == 0)
                        {
                            System.Data.DataTable ITEM2H = GETF1FS2(PARTNOF, VER, PONO, PP);
                            if (ITEM2H.Rows.Count > 0)
                            {
                                ITEM1 = ITEM2H;
                            }
                        }
                        if (ITEM1.Rows.Count == 0)
                        {
                            if (PARTNO3 == "55")
                            {
                                System.Data.DataTable ITEM2H = GETF1FSIT(PARTNOF, PONO, PP);
                                if (ITEM2H.Rows.Count > 0)
                                {
                                    ITEM1 = ITEM2H;
                                }
                            }

                        }
                        //
                    }
                    System.Data.DataTable ITEM2 = GETF2(PARTNO, GRADE, VER);

                    //string QTY = dd["QTY"].ToString();
                    if (ITEM1.Rows.Count > 0)
                    {
                        採購報價料號 = ITEM1.Rows[0][0].ToString();
                        採購報價單價 = ITEM1.Rows[0][1].ToString();
                        LINENUM = ITEM1.Rows[0][2].ToString();
                        採購報價備註 = ITEM1.Rows[0][3].ToString();
                        System.Data.DataTable ITEM3 = GETF3(PONO, 採購報價料號, LINENUM);

                        if (ITEM3.Rows.Count > 0)
                        {
                            dr["輔助SINO"] = ITEM3.Rows[0][0].ToString();

                        }
                        dr["採購報價LINE"] = LINENUM;
                    }

                    //if (MODEL.IndexOf("430QVN02") != -1)
                    //{
                    //    LW = "M";
                    //}
                    System.Data.DataTable ITEM2P = GETF2P(PARTNOF, GRADE, VER);
                    System.Data.DataTable ITEM2P2 = GETF2P2(PARTNOF, GRADE, VER, LW);
                    System.Data.DataTable ITEM2P3 = GETF2P3(PARTNOF, GRADE, VER, LW, PARTNO4);
                    System.Data.DataTable ITEM2P4 = GETF2P4(PARTNOF, GRADE, VER, PARTNO4);

                    if (ITEM2P.Rows.Count == 0 && FE == 0)
                    {
                        採購料號 = PARTNOF;
                    }
                    else
                    {
                        if (ITEM2.Rows.Count > 0)
                        {
                            if (ITEM2P2.Rows.Count == 1)
                            {
                                採購料號 = ITEM2P2.Rows[0][0].ToString();
                            }
                            else if (ITEM2P.Rows.Count == 1)
                            {
                                採購料號 = ITEM2P.Rows[0][0].ToString();
                            }
                            else if (ITEM2.Rows.Count == 1)
                            {
                                採購料號 = ITEM2.Rows[0][0].ToString();
                            }
                            else if (ITEM2P3.Rows.Count == 1)
                            {
                                採購料號 = ITEM2P3.Rows[0][0].ToString();
                            }
                            else
                            {

                                if (MODEL.IndexOf("430QVN02") != -1)
                                {
                                    if (ITEM2P4.Rows.Count == 1)
                                    {
                                        採購料號 = ITEM2P4.Rows[0][0].ToString();
                                    }
                                    else
                                    {
                                        採購料號 = 採購報價料號;

                                    }

                                }
                                else
                                {
                                    採購料號 = 採購報價料號;
                                }

                            }
                        }
                        else
                        {
                            採購料號 = 採購報價料號;
                        }



                    }

                    if (PARTNO3 == "55")
                    {
                        System.Data.DataTable ITEM2H = GETF1FSIT(PARTNOF, PONO, PP);
                        if (ITEM2H.Rows.Count > 0)
                        {
                            採購料號 = ITEM2H.Rows[0][0].ToString();
                        }
                        else
                        {
                            System.Data.DataTable ITEM2PKIT = GETKIT(PARTNOF);
                            if (ITEM2PKIT.Rows.Count > 0)
                            {
                                採購料號 = ITEM2PKIT.Rows[0][0].ToString();
                            }
                        }
                    }

                    //不覆蓋SI
                    dr["不覆蓋SI"] = "0";
                    dr["採購報價料號"] = 採購報價料號;
                    dr["採購料號"] = 採購料號;
                    dr["採購報價單價"] = 採購報價單價;
                    dr["採購報價備註"] = 採購報價備註;
                    dr["採購單價"] = dd["PRICE"].ToString();
                    dr["INVOICE"] = dd["INVOICENO"].ToString();
                    dr["INVOICE日期"] = dd["SHIPDATE"].ToString();

                    if (採購料號 != "")
                    {
                        採購料號 = "ACME00004.00004";

                    }

                    dtCost.Rows.Add(dr);

                }

                dataGridView2.DataSource = dtCost;
            }
        }
        private void D1()
        {
            if (dataGridView2.Rows.Count == 0)
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

            int i = 0; 

            oCompany.CompanyDB = FA;
            oCompany.UserName = "A02";
            oCompany.Password = "6500";
            int result = oCompany.Connect();
            if (result == 0)
            {

                string SI = "";
                string DDATE = "";
  

                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                    oPURCH.CardCode = dataGridView2.Rows[0].Cells["廠商編號"].Value.ToString();
                    oPURCH.DocCurrency = "USD";
                    oPURCH.VatPercent = 0;
                    DDATE = dataGridView2.Rows[0].Cells["出貨日期"].Value.ToString();
                    oPURCH.RequriedDate = Convert.ToDateTime(DDATE);
                    oPURCH.DocDate = Convert.ToDateTime(DDATE);
                    oPURCH.DocDueDate = Convert.ToDateTime(DDATE);
                    System.Data.DataTable OPQT = GSHIP(dataGridView2.Rows[0].Cells["採購報價單號"].Value.ToString());
                    if (OPQT.Rows.Count > 0)
                    {
                        int OwnerCode = Convert.ToInt32(OPQT.Rows[0]["OwnerCode"]);
                        string Comments = OPQT.Rows[0]["Comments"].ToString();

                        oPURCH.DocumentsOwner = OwnerCode;
                        oPURCH.Comments = Comments;
                    }
                    System.Data.DataTable G1 = GetOPOR7();

                    string ShipCity = dataGridView2.Rows[0].Cells["ShipCity"].Value.ToString();
                    string TAX = "";

                    int H1 = ShipCity.IndexOf("新得利");
                    int H2 = ShipCity.IndexOf("聯倉");
                    int H3 = ShipCity.IndexOf("內湖");
                    int H4 = ShipCity.IndexOf("博豐");
                    int H5 = ShipCity.IndexOf("大發");

                    if (H1 != -1 || H2 != -1 || H3 != -1 || H4 != -1 || H5 != -1)
                    {
                        TAX = "AP5%";
                    }
                    else
                    {
                        TAX = "AP0%";
                    }

                    if (G1.Rows.Count > 0)
                    {
                        oPURCH.DocRate = Convert.ToDouble(G1.Rows[0][0]);
                    }


                    for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
                    {
                        DataGridViewRow row;

                        row = dataGridView2.Rows[i2];
                        //採購報價單號
                        string ITEMCODE = row.Cells["採購料號"].Value.ToString();
                        double QTY = Convert.ToDouble(row.Cells["QTY"].Value);
                        double PRICE = Convert.ToDouble(row.Cells["採購單價"].Value);
                        //oPURCH.Lines.BaseEntry = Convert.ToInt16(row.Cells["採購報價單號"].Value);

                        string DOC = row.Cells["採購報價單號"].Value.ToString();
                        // oPURCH.Lines.BaseType = 540000006;
                        oPURCH.Lines.WarehouseCode = "OT001";
                        oPURCH.Lines.ItemCode = ITEMCODE;
                        oPURCH.Lines.Quantity = QTY;
                        oPURCH.Lines.UnitPrice = PRICE;
                        oPURCH.Lines.Price = PRICE;
                        oPURCH.Lines.VatGroup = TAX;
                        oPURCH.Lines.Currency = "USD";
                        oPURCH.Lines.UserFields.Fields.Item("U_ACME_Dscription").Value = row.Cells["付款方式"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_Shipping_no").Value = row.Cells["SINO"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_PAY").Value = row.Cells["採購報價單號"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_ACME_INV").Value = row.Cells["INVOICE"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = Convert.ToDateTime(row.Cells["出貨日期"].Value);
                        oPURCH.Lines.UserFields.Fields.Item("U_BASE_DOC").Value = row.Cells["LCNO"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_ACME_Kind").Value = row.Cells["產地"].Value;
                        oPURCH.Lines.UserFields.Fields.Item("U_MEMO").Value = row.Cells["採購報價備註"].Value;
                        string LINE = "";
                        string 採購報價LINE = row.Cells["採購報價LINE"].Value.ToString();
                        if (String.IsNullOrEmpty(採購報價LINE))
                        {
                            System.Data.DataTable h1 = GERPQT(DOC, ITEMCODE);
                            if (h1.Rows.Count > 0)
                            {
                                LINE = h1.Rows[0][0].ToString();
                            }
                        }
                        else
                        {
                            LINE = 採購報價LINE;
                        }
                        if (LINE == "")
                        {
                            MessageBox.Show("採購報價料號無法對應");
                            return;
                        }
                        oPURCH.Lines.UserFields.Fields.Item("U_CUSTITEMCODE").Value = LINE;
                        if (!String.IsNullOrEmpty(row.Cells["SINO"].Value.ToString()))
                        {
                            SI = row.Cells["SINO"].Value.ToString();
                        }
                        oPURCH.Lines.Add();

                    }


                    int res = oPURCH.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4Q(SI);
                        string OWTR = G4.Rows[0][0].ToString();
                        MessageBox.Show("上傳成功 採購單號 : " + OWTR);
                        System.Data.DataTable G5 = GetDI4Q2(OWTR);
                        if (G5.Rows.Count > 0)
                        {
                            for (int s = 0; s <= G5.Rows.Count - 1; s++)
                            {
                                int DOCENTRY = Convert.ToInt32(G5.Rows[s]["DOCENTRY"]);
                                int LINENUM = Convert.ToInt32(G5.Rows[s]["LINENUM"]);
                                int U_PAY = Convert.ToInt32(G5.Rows[s]["U_PAY"]);
                                int U_CUSTITEMCODE = Convert.ToInt32(G5.Rows[s]["U_CUSTITEMCODE"]);

                                UPOPQT(DOCENTRY, U_PAY, U_CUSTITEMCODE);
                                UPOPORF(U_PAY, U_CUSTITEMCODE, DOCENTRY, LINENUM);
                            }


                            for (int s = 0; s <= G5.Rows.Count - 1; s++)
                            {

                                int U_PAY = Convert.ToInt32(G5.Rows[s]["U_PAY"]);

                                UPOPORF2(U_PAY);
                            }
                        }


                    }


            }
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }



        }
        private void D2F()
        {
            if (dataGridView3.Rows.Count == 0)
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
            oCompany.CompanyDB = FA;
            oCompany.UserName = "A01";
            oCompany.Password = "89206602";
            int result = oCompany.Connect();
            if (result == 0)
            {


                DELINV();
                string 原廠發票項次 = "";
                StringBuilder sb3 = new StringBuilder();

                for (int i2 = 0; i2 <= dataGridView3.Rows.Count - 1; i2++)
                {
                    DataGridViewRow row;

                    row = dataGridView3.Rows[i2];


                    string INVOICE = row.Cells["原廠INVOICE"].Value.ToString();
                    string 單價 = row.Cells["美金單價"].Value.ToString();
                    string UAN = row.Cells["原廠發票項次"].Value.ToString();

       

                    ADDINV(INVOICE, 單價, UAN);

                }

     
                System.Data.DataTable F1 = GETA2();
                if (F1.Rows.Count > 0)
                {
                    for (int s = 0; s <= F1.Rows.Count - 1; s++)
                    {
                        string INV = F1.Rows[s][0].ToString();
                        //SAPbobsCOM.Documents oPURCH = null;
                        //oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes );
                        System.Data.DataTable FINV = GETA2INV(INV);
                        if (FINV.Rows.Count > 0)
                        {
                            for (int s2 = 0; s2 <= FINV.Rows.Count - 1; s2++)
                            {
                                string UAN = FINV.Rows[s2][0].ToString();

                                sb3.Append(UAN + "/");
                            }

                        }
                        if (sb3.Length > 1)
                        {
                            sb3.Remove(sb3.Length - 1, 1);
                            原廠發票項次 = sb3.ToString();
                        }
                        SAPbobsCOM.Documents oPURCH = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oPURCH.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;


                        for (int i2 = 0; i2 <= dataGridView3.Rows.Count - 1; i2++)
                        {
                            DataGridViewRow row;

                            row = dataGridView3.Rows[i2];
                            string INV2 = row.Cells["原廠INVOICE"].Value.ToString();
                            if (INV == INV2)
                            {
                                oPURCH.CardCode = dataGridView3.Rows[i2].Cells["廠商編號"].Value.ToString();
                                oPURCH.DocCurrency = "NTD";
                                oPURCH.VatPercent = 0;

                                oPURCH.DocDate = Convert.ToDateTime(textBox6.Text);
                                System.Data.DataTable G7 = GetOPORH();
                                if (G7.Rows.Count > 0)
                                {
                                    oPURCH.UserFields.Fields.Item("U_ACME_USER").Value = G7.Rows[0][2].ToString();

                                }
                                oPURCH.UserFields.Fields.Item("U_ACME_INV").Value = row.Cells["原廠INVOICE"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_Invoice").Value = row.Cells["INVOICE日期"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_rate1").Value = row.Cells["原廠進貨匯率"].Value;
                                oPURCH.UserFields.Fields.Item("U_LOCATION").Value = row.Cells["產地"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_LC").Value = row.Cells["LCNO"].Value;
                                oPURCH.UserFields.Fields.Item("U_Shipping_no").Value = row.Cells["SHIPPING工單號碼"].Value;
                                oPURCH.UserFields.Fields.Item("U_INVITEM").Value = 原廠發票項次;
                                System.Data.DataTable gg1 = GETA3();
                                if (gg1.Rows.Count > 0)
                                {

                                    oPURCH.UserFields.Fields.Item("U_ACME_Price1").Value = gg1.Rows[0][0].ToString();

                                }

                                oPURCH.DocRate = 1;
                                string ITEMCODE = row.Cells["項目料號"].Value.ToString();
                                double 數量 = Convert.ToDouble(row.Cells["數量"].Value);
                                double 單價 = Convert.ToDouble(row.Cells["台幣單價"].Value);
                                oPURCH.Lines.BaseEntry = Convert.ToInt32(row.Cells["採購單號"].Value);
                                oPURCH.Lines.BaseLine = Convert.ToInt16(row.Cells["LINENUM"].Value);
                                oPURCH.Lines.BaseType = 22;
                                oPURCH.Lines.WarehouseCode = "OT001";
                                oPURCH.Lines.ItemCode = ITEMCODE;
                                oPURCH.Lines.Quantity = 數量;
                                oPURCH.Lines.UnitPrice = 單價;
                                oPURCH.Lines.Price = 單價;
                                oPURCH.Lines.VatGroup = row.Cells["稅碼"].Value.ToString();
                                oPURCH.Lines.Currency = "NTD";
                                oPURCH.Lines.Rate = 1;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_Dscription").Value = row.Cells["付款方式"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_Shipping_no").Value = row.Cells["SHIPPING工單號碼"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_INV").Value = row.Cells["原廠INVOICE"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = row.Cells["INVOICE日期"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_BASE_DOC").Value = row.Cells["LCNO"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_Kind").Value = row.Cells["產地"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_WhsName").Value = "在途倉";

                                oPURCH.Lines.Add();

                            }

                        }
                        int res = oPURCH.Add();
                        if (res != 0)
                        {
                            MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                        }
                        else
                        {
                            System.Data.DataTable G4 = GetODRF();
                            string OWTR = G4.Rows[0][0].ToString();
                            MessageBox.Show("上傳成功 收貨採購草稿單號 : " + OWTR);



                        }

                    }

                }

            }
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }



        }
        private void D2FSELECT()
        {
            if (dataGridView3.Rows.Count == 0)
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
            oCompany.CompanyDB = FA;
            oCompany.UserName = "A01";
            oCompany.Password = "89206602";
            int result = oCompany.Connect();
            if (result == 0)
            {


                DELINV();
                string 原廠發票項次 = "";
                StringBuilder sb3 = new StringBuilder();
                for (int i2 = 0; i2 <= dataGridView3.SelectedRows.Count - 1; i2++)
                {
                    DataGridViewRow row;

                    row = dataGridView3.SelectedRows[i2];


                    string INVOICE = row.Cells["原廠INVOICE"].Value.ToString();
                    string 單價 = row.Cells["美金單價"].Value.ToString();
                    string UAN = row.Cells["原廠發票項次"].Value.ToString();

                

                    ADDINV(INVOICE, 單價, UAN);

                }
      
                System.Data.DataTable F1 = GETA2();
                if (F1.Rows.Count > 0)
                {
                    for (int s = 0; s <= F1.Rows.Count - 1; s++)
                    {
                        string INV = F1.Rows[s][0].ToString();
                        //SAPbobsCOM.Documents oPURCH = null;
                        //oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes );
                        System.Data.DataTable FINV = GETA2INV(INV);
                        if (FINV.Rows.Count > 0)
                        {
                            for (int s2 = 0; s2 <= FINV.Rows.Count - 1; s2++)
                            {
                                string UAN = FINV.Rows[s2][0].ToString();

                                sb3.Append(UAN + "/");
                            }

                        }
                        if (sb3.Length > 1)
                        {
                            sb3.Remove(sb3.Length - 1, 1);
                            原廠發票項次 = sb3.ToString();
                        }
                        SAPbobsCOM.Documents oPURCH = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oPURCH.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;


                        for (int i2 = 0; i2 <= dataGridView3.SelectedRows.Count - 1; i2++)
                        {
                            DataGridViewRow row;

                            row = dataGridView3.SelectedRows[i2];
                            string INV2 = row.Cells["原廠INVOICE"].Value.ToString();
                            if (INV == INV2)
                            {
                                oPURCH.CardCode = dataGridView3.SelectedRows[i2].Cells["廠商編號"].Value.ToString();
                                oPURCH.DocCurrency = "NTD";
                                oPURCH.VatPercent = 0;
                                System.Data.DataTable G7 = GetOPORH();
                                if (G7.Rows.Count > 0)
                                {
                                    oPURCH.UserFields.Fields.Item("U_ACME_USER").Value = G7.Rows[0][2].ToString();

                                }
                                oPURCH.UserFields.Fields.Item("U_ACME_INV").Value = row.Cells["原廠INVOICE"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_Invoice").Value = row.Cells["INVOICE日期"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_rate1").Value = row.Cells["原廠進貨匯率"].Value;
                                oPURCH.UserFields.Fields.Item("U_LOCATION").Value = row.Cells["產地"].Value;
                                oPURCH.UserFields.Fields.Item("U_ACME_LC").Value = row.Cells["LCNO"].Value;
                                oPURCH.UserFields.Fields.Item("U_Shipping_no").Value = row.Cells["SHIPPING工單號碼"].Value;
                                System.Data.DataTable gg1 = GETA3();
                                if (gg1.Rows.Count > 0)
                                {
                                    
                                        oPURCH.UserFields.Fields.Item("U_ACME_Price1").Value = gg1.Rows[0][0].ToString();
                                    
                                }
                                oPURCH.UserFields.Fields.Item("U_INVITEM").Value = 原廠發票項次;
                                oPURCH.DocRate = 1;
                 
                                string ITEMCODE = row.Cells["項目料號"].Value.ToString();
                                double 數量 = Convert.ToDouble(row.Cells["數量"].Value);
                                double 單價 = Convert.ToDouble(row.Cells["台幣單價"].Value);
                                oPURCH.Lines.BaseEntry = Convert.ToInt32(row.Cells["採購單號"].Value);
                                oPURCH.Lines.BaseLine = Convert.ToInt16(row.Cells["LINENUM"].Value);
                                oPURCH.Lines.BaseType = 22;
                                oPURCH.Lines.WarehouseCode = "OT001";
                                oPURCH.Lines.ItemCode = ITEMCODE;
                                oPURCH.Lines.Quantity = 數量;
                                oPURCH.Lines.UnitPrice = 單價;
                                oPURCH.Lines.Price = 單價;
                                oPURCH.Lines.VatGroup = row.Cells["稅碼"].Value.ToString();
                                oPURCH.Lines.Currency = "NTD";
                                oPURCH.Lines.Rate = 1;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_Dscription").Value = row.Cells["付款方式"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_Shipping_no").Value = row.Cells["SHIPPING工單號碼"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_INV").Value = row.Cells["原廠INVOICE"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = row.Cells["INVOICE日期"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_BASE_DOC").Value = row.Cells["LCNO"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_Kind").Value = row.Cells["產地"].Value;
                                oPURCH.Lines.UserFields.Fields.Item("U_ACME_WhsName").Value = "在途倉";
                             
                                //U_INVITEM
                                oPURCH.Lines.Add();

                            }

                        }
                        int res = oPURCH.Add();
                        if (res != 0)
                        {
                            MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                        }
                        else
                        {
                            System.Data.DataTable G4 = GetODRF();
                            string OWTR = G4.Rows[0][0].ToString();
                            MessageBox.Show("上傳成功 收貨採購草稿單號 : " + OWTR);



                        }

                    }

                }

            }
            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }



        }
        public System.Data.DataTable GETLC1(string SHIPPINGCODE, string DOCNUM)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT LC,DOCENTRY,ITEMCODE,Dscription,Quantity,ITEMPRICE,LINENUM FROM lcInstro1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCNUM=@DOCNUM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
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
        public System.Data.DataTable GETLC1F(string SHIPPINGCODE, string DOCNUM, string ITEMCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Quantity FROM lcInstro1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCNUM=@DOCNUM AND ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
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

        public System.Data.DataTable GETLC2(string LCNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DOCNUM,LcAmt FROM APLC WHERE LCNO=@LCNO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCNO", LCNO));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }


        public System.Data.DataTable GETLC3(string DOCNUM)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ISNULL(SUM(AMT),0) AMT FROM PLC1  WHERE DOCNUM=@DOCNUM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
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
        public System.Data.DataTable GETLC4(string DonNo, string LINENUM)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Qty FROM PLC1  WHERE DonNo=@DonNo AND LINENUM=@LINENUM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DonNo", DonNo));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        public System.Data.DataTable GETLC5(string DOCENTRY, string LINENUM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT QUANTITY FROM POR1  WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        public System.Data.DataTable GETCARDNAME(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT boardCountNo DTYPE,CARDNAME,CLOSEDAY FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        private void ss(string SINO, string DOCNO, string DTYPE, string SCARDNAME, string CLOSEDAY, string CARDCODE, string INVOICE)
        {




            System.Data.DataTable D1 = GETLC1(SINO, DOCNO);
            if (D1.Rows.Count > 0)
            {
                for (int i = 0; i <= D1.Rows.Count - 1; i++)
                {

                    string LC = D1.Rows[i]["LC"].ToString();
                    string DOCENTRY = D1.Rows[i]["DOCENTRY"].ToString();
                    string ITEMCODE = D1.Rows[i]["ITEMCODE"].ToString();
                    string Dscription = D1.Rows[i]["Dscription"].ToString();
                    int QTY = Convert.ToInt16(D1.Rows[i]["Quantity"]);
                    decimal ITEMPRICE = Convert.ToDecimal(D1.Rows[i]["ITEMPRICE"]);
                    string LINENUM = D1.Rows[i]["LINENUM"].ToString();
                    System.Data.DataTable D2 = GETLC2(LC);
                    if (D2.Rows.Count > 0)
                    {
                        string DOCNUM = D2.Rows[0][0].ToString();
                        decimal LCAMT = Convert.ToDecimal(D2.Rows[0][1]);
                        System.Data.DataTable D3 = GETLC3(DOCNUM);
                        int Q1 = 0;
                        int Q2 = 0;
                        System.Data.DataTable D4 = GETLC4(DOCENTRY, LINENUM);
                        if (D4.Rows.Count > 0)
                        {
                            Q1 = Convert.ToInt32(D4.Rows[0][0]);
                        }
                        System.Data.DataTable D5 = GETLC5(DOCENTRY, LINENUM);
                        if (D5.Rows.Count > 0)
                        {
                            Q2 = Convert.ToInt32(D5.Rows[0][0]);
                        }
                        int Q3 = Q2 - Q1 - QTY;
                        //if (D4.Rows.Count == 0 || Q3 > 0)
                        //{

                        decimal LCAMT2 = Convert.ToDecimal(D3.Rows[0][0]);
                        //  string DTYPE = boardCountNoTextBox.Text.Trim();
                        decimal taxx = 0;
                        if (DTYPE == "進口" || DTYPE == "內銷")
                        {
                            taxx = 5;
                        }


                        decimal taxx2 = taxx / 100;
                        decimal tax = (QTY * ITEMPRICE * taxx2);
                        decimal AMT = (QTY * ITEMPRICE) + Convert.ToDecimal(tax);
                        string CARDNAME = SCARDNAME.Replace("友達光電股份有限公司", "") + "出貨";
                        decimal d1 = LCAMT - LCAMT2 - AMT;
                        if (d1 < 0)
                        {
                            MessageBox.Show("LC金額不足");
                            return;
                        }
                        ADDPLC1(DOCNUM, "採購單", DOCENTRY, LINENUM, ITEMCODE, Dscription, QTY, ITEMPRICE, ITEMPRICE, taxx.ToString(), tax, AMT, CARDNAME, CLOSEDAY, CARDCODE, SCARDNAME, INVOICE);
                        MessageBox.Show("已新增至LC");
                    }
                    // }

                }
            }
            //  }

        }


        public void ADDPLC1(string DOCNUM, string PKIND, string DONNO, string LINENUM, string ITEMCODE, string ITEMNAME, int QTY, decimal PRICE, decimal COMMENTS, string TAXCODE, decimal TAX, decimal AMT, string CARDNAME, string CargoDate, string CARDCODE, string CARDNAME2, string InvoceNo)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into PLC1(DOCNUM,PKIND,DONNO,LINENUM,ITEMCODE,ITEMNAME,QTY,PRICE,COMMENTS,TAXCODE,TAX,AMT,CARDNAME,CargoDate,CARDCODE,CARDNAME2,InvoceNo) values(@DOCNUM,@PKIND,@DONNO,@LINENUM,@ITEMCODE,@ITEMNAME,@QTY,@PRICE,@COMMENTS,@TAXCODE,@TAX,@AMT,@CARDNAME,@CargoDate,@CARDCODE,@CARDNAME2,@InvoceNo)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            command.Parameters.Add(new SqlParameter("@PKIND", PKIND));
            command.Parameters.Add(new SqlParameter("@DONNO", DONNO));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@COMMENTS", COMMENTS));
            command.Parameters.Add(new SqlParameter("@TAXCODE", TAXCODE));
            command.Parameters.Add(new SqlParameter("@TAX", TAX));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@CargoDate", CargoDate));

            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME2", CARDNAME2));
            command.Parameters.Add(new SqlParameter("@InvoceNo", InvoceNo));
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
        public void ADDAPLC(string DOCNUM, string LCNO, string LCTYPE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into APLC(DOCNUM,LCNO,LCTYPE) values(@DOCNUM,@LCNO,@LCTYPE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            command.Parameters.Add(new SqlParameter("@LCNO", LCNO));
            command.Parameters.Add(new SqlParameter("@LCTYPE", LCTYPE));
            //LCTYPE


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

        private void button6_Click(object sender, EventArgs e)
        {
            D1();

            for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
            {
                DataGridViewRow row;

                row = dataGridView2.Rows[i2];

                string SINO異常 = row.Cells["SINO異常"].Value.ToString();
                string SINO = row.Cells["SINO"].Value.ToString();
                string SS = row.Cells["不覆蓋SI"].Value.ToString();

                if (SS == "0")
                {
                    DELITEM(SINO);
                }

            }

            for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
            {
                DataGridViewRow row;

                row = dataGridView2.Rows[i2];

                string SINO異常 = row.Cells["SINO異常"].Value.ToString();
                string SINO = row.Cells["SINO"].Value.ToString();
                string SINO2 = row.Cells["SINO2"].Value.ToString();
                string SINO3 = row.Cells["SINO3"].Value.ToString();
                string 採購報價單號 = row.Cells["採購報價單號"].Value.ToString();
                string 採購料號 = row.Cells["採購料號"].Value.ToString();
                string LCNO = row.Cells["LCNO"].Value.ToString();
                string CARDCODE = row.Cells["廠商編號"].Value.ToString();
                string LINENUM = row.Cells["LINENUM"].Value.ToString();
                string 採購報價LINE = row.Cells["採購報價LINE"].Value.ToString();
                string CARDNAME = "";
                string DTYPE = "";
                string CLOSEDAY = "";

                string NOSI = row.Cells["不覆蓋SI"].Value.ToString();

                System.Data.DataTable CARD1 = GETCARDNAME(SINO);
                if (CARD1.Rows.Count > 0)
                {
                    CARDNAME = CARD1.Rows[0]["CARDNAME"].ToString();
                    DTYPE = CARD1.Rows[0]["DTYPE"].ToString();
                    CLOSEDAY = CARD1.Rows[0]["CLOSEDAY"].ToString();

                }
                //CARDCODE
                int QTY = Convert.ToInt32(row.Cells["QTY"].Value);
                decimal PRICE = Convert.ToDecimal(row.Cells["採購單價"].Value);
                decimal AMT = Convert.ToDecimal(QTY) * PRICE;

                System.Data.DataTable I1 = GETITEMNAME(採購料號);

                string DESC = "";
                if (I1.Rows.Count > 0)
                {

                    DESC = I1.Rows[0][0].ToString();
                }


                if (NOSI == "0")
                {

                    AddSHIPITEM(SINO, Convert.ToInt32(LINENUM), 採購報價單號, Convert.ToInt16(採購報價LINE), "採購報價", 採購料號, DESC, QTY, PRICE, AMT, "");

                    //True
                    AddSHIPITEM2(SINO, SINO2, LINENUM, 採購報價單號, 採購料號, DESC, QTY.ToString(), PRICE, AMT, LCNO, 採購報價LINE);
                }

                //GETCARDNAME

                // }

            }


            for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
            {
                DataGridViewRow row;

                row = dataGridView2.Rows[i2];

                string SINO異常 = row.Cells["SINO異常"].Value.ToString();
                string SINO = row.Cells["SINO"].Value.ToString();
                string SINO2 = row.Cells["SINO2"].Value.ToString();
                string 採購報價單號 = row.Cells["採購報價單號"].Value.ToString();
                string 採購料號 = row.Cells["採購料號"].Value.ToString();
                string LCNO = row.Cells["LCNO"].Value.ToString();
                string CARDCODE = row.Cells["廠商編號"].Value.ToString();
                string LINENUM = row.Cells["LINENUM"].Value.ToString();
                string INVOICE = row.Cells["INVOICE"].Value.ToString();
                //INVOICE
                string CARDNAME = "";
                string DTYPE = "";
                string CLOSEDAY = "";
                //U_ACME_INV
                System.Data.DataTable CARD1 = GETCARDNAME(SINO);
                if (CARD1.Rows.Count > 0)
                {
                    CARDNAME = CARD1.Rows[0]["CARDNAME"].ToString();
                    DTYPE = CARD1.Rows[0]["DTYPE"].ToString();
                    CLOSEDAY = CARD1.Rows[0]["CLOSEDAY"].ToString();

                }
                //CARDCODE
                int QTY = Convert.ToInt32(row.Cells["QTY"].Value);
                decimal PRICE = Convert.ToDecimal(row.Cells["採購單價"].Value);
                decimal AMT = Convert.ToDecimal(QTY) * PRICE;

                System.Data.DataTable I1 = GETITEMNAME(採購料號);

                string DESC = "";
                if (I1.Rows.Count > 0)
                {

                    DESC = I1.Rows[0][0].ToString();
                }


                //if (SINO異常 != "")
                //{


                if (LCNO != "")
                {

                    ss(SINO, SINO2, DTYPE, CARDNAME, CLOSEDAY, CARDCODE, INVOICE);

                }
                //GETCARDNAME

                // }

            }

        }

        private void dataGridView2_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    listBox1.Items.Clear();
                    for (int i = dataGridView2.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView2.SelectedRows[i];

                        listBox2.Items.Add(row.Cells["ID"].Value.ToString());

                    }


                    ArrayList al = new ArrayList();

                    for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                    {
                        al.Add(listBox2.Items[i].ToString());
                    }
                    StringBuilder sb = new StringBuilder();



                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);


                }
                else
                {
                    listBox2.Items.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTableCombine3();
            System.Data.DataTable G1 = GetAPPLE();
            int LI = 0;
            string DUP = "";
            for (int i = 0; i <= G1.Rows.Count - 1; i++)
            {
                DataRow dd = G1.Rows[i];
                dr = dtCost.NewRow();




                dr["採購單號"] = dd["採購單號"].ToString();
                dr["LINENUM"] = dd["LINENUM"].ToString();
                dr["採購單過帳日期"] = dd["採購單過帳日期"].ToString();
                dr["廠商編號"] = dd["廠商編號"].ToString();
                dr["項目料號"] = dd["項目料號"].ToString();
                dr["項目說明"] = dd["項目說明"].ToString();
                dr["PART NO"] = dd["PARTNO"].ToString();
                dr["數量"] = Convert.ToDecimal(dd["數量"]);
                double P1 = Convert.ToDouble(dd["單價"]);
                double QTY = Convert.ToDouble(dd["數量"]);
                dr["美金單價"] = P1;
                string PRATE = dd["稅碼"].ToString();
                double A2 = 0;
                double A1 = 0;
                if (!String.IsNullOrEmpty(dd["INVOICE日期"].ToString()))
                {
                    DateTime D1 = Convert.ToDateTime(dd["INVOICE日期"]);
                    System.Data.DataTable GA2 = GetAPPLE2(D1);
                    if (GA2.Rows.Count > 0)
                    {
                        double R1 = Convert.ToDouble(GA2.Rows[0][0]);
                        double R3 = R1 * P1;
                        A1 = Math.Round(R3 * QTY, 0, MidpointRounding.AwayFromZero);
                        if (PRATE == "AP5%")
                        {
                            A2 = Math.Round(R3 * QTY * 0.050, 0, MidpointRounding.AwayFromZero);
                        }

                        dr["台幣單價"] = R3;
                        dr["原廠進貨匯率"] = R1.ToString();
                    }
                }
                dr["付款方式"] = dd["付款方式"].ToString();
                dr["稅碼"] = dd["稅碼"].ToString();
                dr["倉庫"] = "OT001";
                dr["SHIPPING工單號碼"] = dd["工單號碼"].ToString();
                dr["原廠INVOICE"] = dd["原廠INVOICE"].ToString();
                dr["INVOICE日期"] = dd["INVOICE日期"].ToString();
                dr["LCNO"] = dd["LCNO"].ToString();
                string LOC = dd["產地"].ToString();
                System.Data.DataTable GLOC = GetLOC(LOC);
                if (GLOC.Rows.Count > 0)
                {
                    dr["產地"] = GLOC.Rows[0][0].ToString();
                }
                dr["採購單號"] = dd["採購單號"].ToString();
                dr["台幣未稅金額"] = A1.ToString();
                dr["台幣稅額"] = A2.ToString();
                dr["台幣含稅金額"] = (A1 + A2).ToString();
                dr["原廠發票項次"] = "";
                
                dtCost.Rows.Add(dr);
            }
            dataGridView3.DataSource = dtCost;

            if (dtCost.Rows.Count > 0)
            {

                string gk1 = dtCost.Compute("Sum(台幣未稅金額)", null).ToString();
                string gk2 = dtCost.Compute("Sum(台幣稅額)", null).ToString();
                string gk3 = dtCost.Compute("Sum(台幣含稅金額)", null).ToString();

                decimal shk1 = Convert.ToDecimal(gk1);
                decimal shk2 = Convert.ToDecimal(gk2);
                decimal shk3 = Convert.ToDecimal(gk3);

                label9.Text = "台幣未稅金額:" + shk1.ToString("#,##0");
                label10.Text = "台幣稅額:" + shk2.ToString("#,##0");
                label11.Text = "台幣含稅金額:" + shk3.ToString("#,##0");
            }
            System.Data.DataTable dt4 = GetAPPLECOMB();


            comboBox2.Items.Clear();

            comboBox2.Items.Add("");
            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
             
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            //採購報價LINE
            if (dataGridView2.Columns[e.ColumnIndex].Name == "新增SINO")
            {
                string FF = this.dataGridView2.Rows[e.RowIndex].Cells["新增SINO"].Value.ToString();
                string SINO = this.dataGridView2.Rows[e.RowIndex].Cells["SINO"].Value.ToString();
                System.Data.DataTable GS = GETLC(SINO);
                string DOCNUM = "";
                if (GS.Rows.Count > 0)
                {
                    DOCNUM = GS.Rows[0][0].ToString();
                }
                if (FF == "1")
                {
                    DialogResult result;
                    result = MessageBox.Show("請確認是否要新增SINO", "YES/NO", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        string NumberName = "SI" + DateTime.Now.ToString("yyyyMMdd");
                        SqlConnection Connection = globals.Connection;
                        string AutoNum = util.GetAutoNumber(Connection, NumberName);

                        string NumberName2 = "LI" + DateTime.Now.ToString("yyyyMMdd");
                        string AutoNum2 = util.GetAutoNumber(Connection, NumberName2);
                        string kk2 = NumberName2 + AutoNum2;

                        string KK = NumberName + AutoNum + "X";

                        MessageBox.Show("新增單號" + KK);
                        try
                        {
                            AddSHIPMAIN2(SINO, KK);
                            AddSHIPMAIN3(SINO, KK, kk2);
                        }
                        catch { }

                        for (int i2 = 0; i2 <= dataGridView2.Rows.Count - 1; i2++)
                        {
                            DataGridViewRow row;

                            row = dataGridView2.Rows[i2];


                            string SINO2 = row.Cells["SINO"].Value.ToString();


                            if (SINO2 == SINO)
                            {
                                dataGridView2.Rows[i2].Cells["SINO"].Value = KK;
                                dataGridView2.Rows[i2].Cells["SINO2"].Value = kk2;
                            }

                        }
                    }
                }
            }

            if (dataGridView2.Columns[e.ColumnIndex].Name == "採購報價料號")
            {
                //採購報價LINE  採購報價單號 採購報價單價
                string 採購報價料號 = this.dataGridView2.Rows[e.RowIndex].Cells["採購報價料號"].Value.ToString();
                string 採購報價單號 = this.dataGridView2.Rows[e.RowIndex].Cells["採購報價單號"].Value.ToString();
                System.Data.DataTable J1 = GetCOLOU1(採購報價單號, 採購報價料號);
                if (J1.Rows.Count > 0)
                {

                    dataGridView2.Rows[e.RowIndex].Cells["採購報價LINE"].Value = J1.Rows[0][0].ToString();
                    dataGridView2.Rows[e.RowIndex].Cells["採購報價單價"].Value = J1.Rows[0][1].ToString();
                }

            }
        }
        public static System.Data.DataTable GetCOLOU1(string docentry, string itemcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select LINENUM,PRICE from pqt1 where docentry=@docentry and itemcode=@itemcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@itemcode", itemcode));

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
        public static System.Data.DataTable GetLOC(string PARAM_NO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PARAM_DESC  FROM RMA_PARAMS WHERE PARAM_KIND='APLOC' AND PARAM_NO =@PARAM_NO");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));


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
        public static System.Data.DataTable GetOPOR2TV(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  COUNT(DOCENTRY) FROM AP_OPOR WHERE USERS=@USERS AND DOCENTRY >0  GROUP BY DOCENTRY ORDER BY DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        private void btnTVExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR(fmLogin.LoginID.ToString());
                TV(opdf.FileName);
                textBox1.Text = "TV";
                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());
                        int ROW = Convert.ToInt32(GetOPOR2TV(fmLogin.LoginID.ToString()).Rows[0][0].ToString());//全部同PO
                        if (G3.Rows.Count > 0)
                        {
                            int h = -1;
                            for (int s = 0; s <= G3.Rows.Count - 1; s++)
                            {
                                h++;

                                if (h >= ROW)
                                {

                                    l++;
                                    h = 0;
                                }

                                UPOPOR(l, h, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }
                }


                System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                dataGridView1.DataSource = G1;
                System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();

            }
        }

        private void btnPIDExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR(fmLogin.LoginID.ToString());
                PID(opdf.FileName);

                textBox1.Text = "PID";

                System.Data.DataTable G2 = GetOPOR2(fmLogin.LoginID.ToString());

                if (G2.Rows.Count > 0)
                {
                    int l = 0;
                    for (int i = 0; i <= G2.Rows.Count - 1; i++)
                    {
                        l++;

                        string T1 = G2.Rows[i][0].ToString();
                        System.Data.DataTable G3 = GetOPOR3(T1, fmLogin.LoginID.ToString());

                        if (G3.Rows.Count > 0)
                        {
                            int h = -1;
                            for (int s = 0; s <= G3.Rows.Count - 1; s++)
                            {

                                UPOPOR(1, 0, Convert.ToInt16(G3.Rows[s]["ID"]), fmLogin.LoginID.ToString());
                            }
                        }
                    }
                }

                System.Data.DataTable G1 = GetOPOR(fmLogin.LoginID.ToString());
                dataGridView1.DataSource = G1;
                System.Data.DataTable GG2 = GetOPOR6(fmLogin.LoginID.ToString());
                label1.Text = "數量 : " + GG2.Rows[0]["QTY"].ToString();
                label2.Text = "金額 : " + GG2.Rows[0]["AMT"].ToString();
            }
        }
        private void TV(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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
                string LINE = "1";
                string AU97;
                string GRADE;
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK1;
                string REMARK2 = "";
                string P1;

                string ModelName;
                string AuoPN;
                string Grade;
                string FAB;
                string Price;
                string Remarks;
                string Apr;
                string Tcon;
                string Tqty;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                ModelName = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                Tcon = range.Text.ToString().Trim();

                if (ModelName != "" && ModelName != "MODEL_NAME")
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    AuoPN = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    Grade = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    FAB = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    Price = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                    range.Select();
                    Remarks = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    Apr = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    range.Select();
                    Tqty = range.Text.ToString().Trim();

                    decimal n;

                    string partno = "";//產品大項分類
                    string model = "";//原廠型號共計8碼
                    string grade = "";
                    string partno2 = "";//原廠Ver. & PART No.共計3碼
                    string fab = "";//交易別 & 產地別 & 產品特殊狀況描述
                    //model
                    switch (AuoPN.Substring(0, 2))
                    {
                        case "91":
                            //若為91則改為O
                            model = "O" + ModelName.Substring(1, 8);
                            ModelName = "O" + ModelName.Substring(1, ModelName.Length - 1) + "." + AuoPN.Split('.')[2].Substring(0, 1);
                            break;
                        default:
                            model = ModelName.Split('.')[0].Substring(0, 9);
                            ModelName = ModelName + "." + AuoPN.Split('.')[2].Substring(0, 1);
                            break;
                    }
                    //grade
                    switch (Grade)
                    {
                        case "A":
                            grade = "1";
                            break;
                        case "A-":
                            grade = "5";
                            break;
                        case "V":
                            grade = "3";
                            break;

                    }
                    //AMT
                    string s = "";
                    if (ModelName.Substring(1, 3) == "320" || ModelName.Substring(1, 3) == "430" || ModelName.Substring(1, 3) == "750" || ModelName.Substring(1, 3) == "850")
                    {
                        s = ModelName.Substring(1, 6);//32,43才有差別要取到六碼
                    }
                    else
                    {
                        s = ModelName.Substring(1, 3);//其餘三碼判斷以免增加後三碼(QVR、QVN就要修改程式)                        )
                    }
                    Price = GETPrice(s, Grade);

                    int x, y;
                    x = int.Parse(Price.Split(',')[1]);
                    y = int.Parse(Price.Split(',')[0]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                    range.Select();
                    Price = range.Text.ToString().Trim();


                    if (decimal.TryParse(Price, out n))
                    {
                        PRICE = Convert.ToDecimal(Price);
                    }
                    else
                    {
                        PRICE = 0;
                    }
                    if (decimal.TryParse(Apr, out n))
                    {
                        QTY = Convert.ToDecimal(Apr);
                    }
                    else
                    {
                        QTY = 0;
                    }

                    AMT = PRICE * QTY;
                    //partno2
                    partno2 = AuoPN.Split('.')[2];
                    //fab
                    switch (FAB)
                    {
                        case "M02":
                        case "M11":
                        case "L8B":
                            fab = "1";
                            break;
                        default:
                            fab = "2";
                            break;

                    }
                    ITEMCODE = model + "." + grade + partno2 + fab;
                    try
                    {
                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {
                            ADDOPOR(Convert.ToInt16(LINE), ITEMCODE, QTY, PRICE, AMT, Remarks, "S0001-TV", fmLogin.LoginID.ToString(), ModelName, "", "");
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    //處理tcon
                    if (Tcon != "" && Tcon != "TCON PN" && Tqty != "")
                    {
                        System.Data.DataTable dt = GetTCONItemCode(Tcon);
                        string Location = "";
                        if (dt.Rows.Count > 0)
                        {
                            ITEMCODE = dt.Rows[0]["ItemCode"].ToString().Substring(0, dt.Rows[0]["ItemCode"].ToString().Length - 1) + fab;

                            switch (Tcon.Substring(0, 5))
                            {
                                case "55.32":
                                    Location = "14,6";
                                    break;
                                case "55.43":
                                    Location = "14,9";
                                    break;
                                case "55.50":
                                    Location = "14,10";
                                    break;
                                case "55.55":
                                    Location = "14,11";
                                    break;
                                case "55.65":
                                    Location = "14,12";
                                    break;
                                case "55.75":
                                    Location = "14,13";//先抓4K
                                    break;
                                case "55.85":
                                    Location = "14,15";//先抓4K
                                    break;

                            }
                            x = int.Parse(Location.Split(',')[1]);
                            y = int.Parse(Location.Split(',')[0]);
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                            range.Select();
                            Price = range.Text.ToString().Trim();

                            PRICE = decimal.Parse(Price);
                            QTY = Tqty == "" ? QTY : int.Parse(Tqty);
                            AMT = PRICE * QTY;
                            ModelName = ITEMCODE.Split('.')[0] + "." + ITEMCODE.Split('.')[1].Substring(1, 1);


                        }
                        else
                        {
                            //一次性料號
                            ITEMCODE = "ACME00001.00001";
                            QTY = Tqty == "" ? QTY : int.Parse(Tqty);
                            switch (Tcon.Substring(0, 5))
                            {
                                case "55.32":
                                    Location = "14,6";
                                    break;
                                case "55.43":
                                    Location = "14,9";
                                    break;
                                case "55.50":
                                    Location = "14,10";
                                    break;
                                case "55.55":
                                    Location = "14,11";
                                    break;
                                case "55.65":
                                    Location = "14,12";
                                    break;
                                case "55.75":
                                    Location = "14,13";//先抓4K
                                    break;
                                case "55.85":
                                    Location = "14,15";//先抓4K
                                    break;

                            }
                            x = int.Parse(Location.Split(',')[1]);
                            y = int.Parse(Location.Split(',')[0]);
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[x, y]);
                            range.Select();
                            Price = range.Text.ToString().Trim();

                            PRICE = decimal.Parse(Price);
                            AMT = PRICE * QTY;
                            ModelName = ITEMCODE.Split('.')[0] + "." + ITEMCODE.Split('.')[1].Substring(1, 1);//可能會沒有這筆tcon料號 另外再手動上傳
                        }
                        try
                        {
                            if (!String.IsNullOrEmpty(ITEMCODE))
                            {
                                ADDOPOR(Convert.ToInt16(LINE), ITEMCODE, QTY, PRICE, AMT, Remarks, "S0001-TV", fmLogin.LoginID.ToString(), ModelName, "", "");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
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

        public static System.Data.DataTable GetTCONItemCode(string TCON)
        {
            SqlConnection MyConnection = globals.shipConnection; ;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ItemCode FROM OITM　WHERE ItemName LIKE @TCON AND FROZENFOR = 'N' AND CANCELED = 'N'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TCON", "%" + TCON + "%"));
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
        private string GETPrice(string Model, string Grade)
        {
            string Location = "";
            int PRICE = 0;
            if (Grade == "A")
            {

                switch (Model)
                {
                    case "215":
                        Location = "15,3";
                        break;
                    case "238":
                        Location = "15,4";
                        break;
                    case "320HVN":
                        Location = "15,6";
                        break;
                    case "320XVN":
                        Location = "15,5";
                        break;
                    case "390":
                        Location = "15,7";
                        break;
                    case "430HVN":
                        Location = "15,8";
                        break;
                    case "430QVN":
                        Location = "15,9";
                        break;
                    case "500":
                        Location = "15,10";
                        break;
                    case "550":
                        Location = "15,11";
                        break;
                    case "650":
                        Location = "15,12";
                        break;
                    case "750QVN":
                        Location = "15,13";
                        break;
                    case "750MVR":
                        Location = "15,14";
                        break;
                    case "850QVN":
                        Location = "15,15";
                        break;
                    case "850MVR":
                        Location = "15,16";
                        break;
                }
            }
            else if (Grade == "A-")
            {
                switch (Model)
                {
                    case "215":
                        Location = "16,3";
                        break;
                    case "238":
                        Location = "16,4";
                        break;
                    case "320HVN":
                        Location = "16,6";
                        break;
                    case "320XVN":
                        Location = "16,5";
                        break;
                    case "390":
                        Location = "16,7";
                        break;
                    case "430HVN":
                        Location = "16,8";
                        break;
                    case "430QVN":
                        Location = "16,9";
                        break;
                    case "500":
                        Location = "16,10";
                        break;
                    case "550":
                        Location = "16,11";
                        break;
                    case "650":
                        Location = "16,12";
                        break;
                    case "750QVN":
                        Location = "16,13";
                        break;
                    case "750MVR":
                        Location = "16,14";
                        break;
                    case "850QVN":
                        Location = "16,15";
                        break;
                    case "850MVR":
                        Location = "16,16";
                        break;
                }
            }
            else if (Grade == "V")
            {
                switch (Model)
                {
                    case "215":
                        Location = "17,3";
                        break;
                    case "238":
                        Location = "17,4";
                        break;
                    case "320HVN":
                        Location = "17,6";
                        break;
                    case "320XVN":
                        Location = "17,5";
                        break;
                    case "390":
                        Location = "17,7";
                        break;
                    case "430HVN":
                        Location = "17,8";
                        break;
                    case "430QVN":
                    case "430QVR":
                        Location = "17,9";
                        break;
                    case "500":
                        Location = "17,10";
                        break;
                    case "550":
                        Location = "17,11";
                        break;
                    case "650":
                        Location = "17,12";
                        break;
                    case "750QVN":
                    case "750QVR":
                        Location = "17,13";
                        break;
                    case "750MVR":
                        Location = "17,14";
                        break;
                    case "850QVN":
                    case "850QVR":
                        Location = "17,15";
                        break;
                    case "850MVR":
                        Location = "17,16";
                        break;
                }
            }
            return Location;
        }
        private void PID(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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
                string LINE = "1";
                string AU97;
                string GRADE;
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK1;
                string REMARK2 = "";
                string P1;

                string ModelVersion;
                string PartNo;
                string Grade;
                string FAB;
                string Price;
                string Remarks;
                string Apr;


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                ModelVersion = range.Text.ToString().Trim();

                if (ModelVersion != "" && ModelVersion != "MODEL_VERSION")
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    PartNo = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    Grade = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    FAB = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                    range.Select();
                    Price = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                    range.Select();
                    Remarks = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    range.Select();
                    Apr = range.Text.ToString().Trim();

                    decimal n;
                    if (decimal.TryParse(Price, out n))
                    {
                        PRICE = Convert.ToDecimal(Price);
                    }
                    else
                    {
                        PRICE = 0;
                    }
                    if (decimal.TryParse(Apr, out n))
                    {
                        QTY = Convert.ToDecimal(Apr);
                    }
                    else
                    {
                        QTY = 0;
                    }

                    AMT = PRICE * QTY;
                    string partno = "";//產品大項分類
                    string model = "";//原廠型號共計8碼
                    string grade = "";
                    string partno2 = "";//原廠Ver. & PART No.共計3碼
                    string fab = "";//交易別 & 產地別 & 產品特殊狀況描述
                    //model
                    switch (PartNo.Substring(0, 2))
                    {
                        case "91":
                            //若為91則改為O
                            model = "O" + ModelVersion.Split('.')[0].Substring(1, 8);
                            ModelVersion = "O" + ModelVersion.Substring(1, ModelVersion.Length - 1);
                            break;
                        default:
                            model = ModelVersion.Split('.')[0].Substring(0, 9);
                            break;
                    }
                    //grade
                    switch (Grade)
                    {
                        case "Z":
                            grade = "0";
                            break;
                        case "P":
                            grade = "1";
                            break;
                        case "N":
                            grade = "5";
                            break;

                    }
                    //partno2
                    partno2 = PartNo.Split('.')[2];
                    //fab
                    switch (FAB.Split('_')[1])
                    {
                        case "M02":
                        case "M11":
                        case "L8B":
                            fab = "1";
                            break;
                        default:
                            fab = "2";
                            break;

                    }


                    ITEMCODE = model + "." + grade + partno2 + fab;
                    try
                    {
                        if (String.IsNullOrEmpty(LINE))
                        {
                            LINE = "1";
                        }


                        if (String.IsNullOrEmpty(ITEMCODE))
                        {
                            ITEMCODE = "ACME00004.00004";
                        }

                        ADDOPOR(1, ITEMCODE, QTY, PRICE, AMT, Remarks, "S0623-PID", fmLogin.LoginID.ToString(), ModelVersion, "", "");



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

        private void button8_Click(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                D2FSELECT();
            }
            else
            {
                D2F();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();
            if (LookupValues != null)
            {

                textBox3.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G5 = GetDI4Q2(textBox5.Text);

            for (int s = 0; s <= G5.Rows.Count - 1; s++)
            {

                int U_PAY = Convert.ToInt32(G5.Rows[s]["U_PAY"]);

                UPOPORF2(U_PAY);
            }
           
            //if (G5.Rows.Count > 0)
            //{
            //    for (int s = 0; s <= G5.Rows.Count - 1; s++)
            //    {
            //        int DOCENTRY = Convert.ToInt32(G5.Rows[s]["DOCENTRY"]);
            //        int LINENUM = Convert.ToInt32(G5.Rows[s]["LINENUM"]);
            //        int U_PAY = Convert.ToInt32(G5.Rows[s]["U_PAY"]);
            //        int U_CUSTITEMCODE = Convert.ToInt32(G5.Rows[s]["U_CUSTITEMCODE"]);

            //        UPOPQT(DOCENTRY, U_PAY, U_CUSTITEMCODE);
            //        UPOPORF(U_PAY, U_CUSTITEMCODE, DOCENTRY, LINENUM);
            //    }


            //    for (int s = 0; s <= G5.Rows.Count - 1; s++)
            //    {

            //        int U_PAY = Convert.ToInt32(G5.Rows[s]["U_PAY"]);

            //        UPOPORF2(U_PAY);
            //    }
        

        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                CalcTotals2();
            }
        }
        private void CalcTotals2()
        {


            decimal shk1 = 0;
            decimal shk2 = 0;
            decimal shk3 = 0;

            int i = this.dataGridView3.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                shk1 += Convert.ToInt32(dataGridView3.SelectedRows[iRecs].Cells["台幣未稅金額"].Value);
                shk2 += Convert.ToDecimal(dataGridView3.SelectedRows[iRecs].Cells["台幣稅額"].Value);
                shk3 += Convert.ToDecimal(dataGridView3.SelectedRows[iRecs].Cells["台幣含稅金額"].Value);
            }


            label9.Text = "台幣未稅金額:" + shk1.ToString("#,##0");
            label10.Text = "台幣稅額:" + shk2.ToString("#,##0");
            label11.Text = "台幣含稅金額:" + shk3.ToString("#,##0");

 

        }



        private void button10_Click_1(object sender, EventArgs e)
        {
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\1.SummaryResult";

            //System.Data.DataTable GG = XmlStringToDataTable(OutPutFile);

            //int g = GG.Rows.Count;
            //dataGridView4.DataSource = XmlStringToDataTable(OutPutFile);
            //H1(OutPutFile);



        }


 
    }

}

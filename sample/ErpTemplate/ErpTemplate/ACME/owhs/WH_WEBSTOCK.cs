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
    public partial class WH_WEBSTOCK : Form
    {
       
        private System.Data.DataTable TempDt;
        DataRow dr = null;
        private System.Data.DataTable TempDt2;
        public WH_WEBSTOCK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string NAME = "";
            string COMPANY = "";
            try
            {
                TRUNWH_WEBSTOCK2();
                string d = @"\\acmesrv01\Public\進出貨序號\倉庫庫存表";
                string d2 = @"\\acmesrv01\Public\進出貨序號\倉庫庫存表\包裝分佈明細";
                DD2(d, 1);
                string[] filenames = Directory.GetFiles(d);
                TRUNCATE();
                foreach (string file in filenames)
                {

                    string WH = "";
                    FileInfo info = new FileInfo(file);
                    NAME = info.Name.ToString().Trim().Replace(" ", "");
                    if (NAME != "Thumbs.db")
                    {
                        int T1 = NAME.LastIndexOf("-");
                        int T2 = NAME.LastIndexOf(".");
                        int T3 = NAME.IndexOf("-");
                        COMPANY = NAME.Substring(T1 + 1, T2 - T1 - 1);
                        COMPANY = COMPANY.Replace(".", "");
                        WH = NAME.Substring(0, T3);
                        int T21 = NAME.IndexOf("聯揚");
                        int T41 = NAME.IndexOf("新得利");
                        int T32 = NAME.IndexOf("深圳巨航機保");
                        // 包裝分佈明細-深圳巨航機保
                        //          TRUNCATE(WH, COMPANY);
                        if (T21 != -1)
                        {
                            //聯揚
                            GetExcelContentGD44(file, WH, 1, 2, 6, 3, COMPANY, NAME);
                        }
                        else if (T41 != -1)
                        {
                            //新得利
                            GetExcelContentGD44(file, WH, 8, 5, 12, 9, COMPANY, NAME);
                        }

                        else
                        {
                            GetExcelContentGD44(file, WH, 1, 8, 11, 4, COMPANY, NAME);

                        }
                    }

                }
                string[] filenames2 = Directory.GetFiles(d2);
                foreach (string file in filenames2)
                {
                    FileInfo info = new FileInfo(file);
                    NAME = info.Name.ToString().Trim().Replace(" ", "");
                    int T32 = NAME.IndexOf("包裝分佈明細");
                    int T33 = NAME.IndexOf("深圳巨航機保");
                    if (T32 != -1)
                    {
                        if (T33 != -1)
                        {
                            GetExcelContentGD45(file);
                        }
                    }
                }
                //GetExcelContentGD45

                System.Data.DataTable GT5 = Get5();
                TempDt = MakeTable();

                for (int i = 0; i <= GT5.Rows.Count - 1; i++)
                {
                    string QTY5 = GT5.Rows[i]["QTY"].ToString();
                    string ITEMCODE5 = GT5.Rows[i]["ITEMCODE"].ToString();
                    string DOCTYPE5 = GT5.Rows[i]["DOCTYPE"].ToString();
                    string COMPANY2 = GT5.Rows[i]["COMPANY"].ToString().Trim().ToUpper();
                    System.Data.DataTable G1 = null;
                    int F = 0;
                    if (COMPANY2 == "ACME")
                    {
                        G1 = Get4(ITEMCODE5, DOCTYPE5);
                        System.Data.DataTable D1 = Get4D(ITEMCODE5, DOCTYPE5);
                        if (D1.Rows.Count > 0)
                        {
                            F = Convert.ToInt16(D1.Rows[0][0]);
                        }
                    }
                    else
                    {

                        string strCn = "";

                        if (COMPANY2 == "宇豐")
                        {
                            strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                        }

                        if (COMPANY2 == "IPGI")
                        {
                            strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                        }
                        if (COMPANY2 == "CHOICE")
                        {
                            strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                        }
                        if (COMPANY2 == "TOP")
                        {
                            strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                        }

                        G1 = Get41(ITEMCODE5, DOCTYPE5, strCn);
                    }
                    
                    if (G1.Rows.Count > 0)
                    {
                        string QQ = G1.Rows[0][0].ToString();

                        if (QQ != QTY5)
                        {
           
                            if (Convert.ToInt32(QQ) != Convert.ToInt32(QTY5)+F)
                            {
                                dr = TempDt.NewRow();
                                dr["公司"] = COMPANY2;
                                dr["倉庫"] = DOCTYPE5;
                                dr["料號"] = ITEMCODE5;
                                dr["儲位數量"] = (Convert.ToInt32(QTY5) + F).ToString();
                                dr["ERP數量"] = QQ;
                                TempDt.Rows.Add(dr);
                                DELETE(DOCTYPE5, ITEMCODE5, COMPANY2);
                            }
                        }
                    }
                }
                MessageBox.Show("匯入成功");
                dataGridView2.DataSource = TempDt;
                dataGridView1.DataSource = Get3();
              
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.Message + " " + NAME + "沒有放在第一頁" + COMPANY);
            }
        }

        public void DD2(string PATH,int DOCTYPE)
        {
            TempDt2 = MakeTable2();
            string NAME = "";
            string COMPANY = "";
            try
            {

                string d = @"\\acmesrv01\Public\進出貨序號\倉庫庫存表";

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {
                    DataRow dr;
                    string WH = "";
                    FileInfo info = new FileInfo(file);
                    NAME = info.Name.ToString().Trim().Replace(" ", "");
                    if (NAME != "Thumbs.db")
                    {
                        int T1 = NAME.LastIndexOf("-");
                        int T2 = NAME.LastIndexOf(".");
                        int T3 = NAME.IndexOf("-");
                        COMPANY = NAME.Substring(T1 + 1, T2 - T1 - 1);
                        WH = NAME.Substring(0, T3);


                        dr = TempDt2.NewRow();
                        dr["公司"] = COMPANY;
                        dr["EXCEL"] = NAME;
                        TempDt2.Rows.Add(dr);
                        if (DOCTYPE == 1)
                        {
                            AddWH_WEBSTOCK2(COMPANY, NAME);
                        }
                    }

                }


            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.Message);
            }

      
                       
        }
        private void DG()
        { 
        
}
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("EXCEL", typeof(string));

            return dt;
        }
        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT COMPANY 公司,DOCTYPE 倉庫, ITEMCODE 料號,DOCDATE 日期,QTY 數量,INVOICE  FROM WH_WEBSTOCK WHERE 1=1 ");
            if (textBox1.Text != "")
            {
                sb.Append("   AND ITEMCODE like '%" + textBox1.Text + "%'  ");
            }
          //  DOCTYPE=@DOCTYPE
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
        private System.Data.DataTable Get4(string ITEMCODE, string CITY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(T0.ONHAND AS INT) ONHAND FROM OITW T0");
            sb.Append(" LEFT JOIN OWHS T1 ON (T0.WHSCODE=T1.WHSCODE)");
            sb.Append(" WHERE ITEMCODE=@ITEMCODE AND T1.CITY =@CITY");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CITY", CITY));
        
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
        private System.Data.DataTable Get4D(string ITEMCODE, string CITY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(cast(T0.quantity as int)) 數量 FROM acmesql02.dbo.rdr1 T0     ");
            sb.Append(" left join  acmesql02.dbo.ORDR T3 on (T0.DOCENTRY=T3.DOCENTRY )     ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.owhs T7 ON T0.whscode=T7.whscode    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T8 ON T0.ITEMCODE=T8.ITEMCODE    ");
            sb.Append(" where 1=1 and t3.canceled <> 'Y' AND T3.doctype='I'     ");
            sb.Append(" and isnull(T0.u_acme_workday,'') <> ('缺貨')        ");
            sb.Append(" AND   (ItmsGrpCod =1032  OR T0. ITEMCODE ='ZBFREIGHT.00002')  and t0.linestatus ='O'        ");
            sb.Append(" AND isnull(Convert(varchar(8),T0.U_ACME_WORK,112),'')  <= '20190926'    ");
            sb.Append(" AND T0.ITEMCODE=@ITEMCODE AND T7.CITY =@CITY");
            sb.Append(" GROUP BY T0.itemcode,t7.whsname ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CITY", CITY));

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
        private System.Data.DataTable Get41(string ProdID, string WareHouseName, string strCn)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(T0.Quantity-T0.lendquan+T0.BorrowQuan AS INT)  ONHAND   FROM comWareAmount  T0 LEFT JOIN comWareHouse T1 ON (T0.WareID =T1.WareHouseID) WHERE T1.WareHouseName =@WareHouseName AND T0.ProdID=@ProdID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@WareHouseName", WareHouseName));

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
        private System.Data.DataTable Get5()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY) QTY,COMPANY,ITEMCODE,DOCTYPE,COMPANY FROM WH_WEBSTOCK  GROUP BY ITEMCODE,DOCTYPE,COMPANY");

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

        private System.Data.DataTable Get6(string ITEMCODE, string INVOICE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ID FROM WH_WEBSTOCK WHERE ITEMCODE=@ITEMCODE AND INVOICE =@INVOICE AND DOCTYPE='深圳巨航機保'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));


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
        private System.Data.DataTable GetWH_WEBSTOCK2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT COMPANY 公司,[FILENAME] EXCEL,[STATUS] ' ' FROM WH_WEBSTOCK2 ");

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
        private void GetExcelContentGD44(string ExcelFile, string DOCTYPE, int MDOCDATE, int MITEMCODE, int MQTY, int MINV, string COMPANY, string NAME)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false ;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (iRowCnt > 3000)
            {
                iRowCnt = 3000;
            }

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string DOCDATE;
            string ITEMCODE = "";
            string QTY = "";
            string INV = "";

            string P1 = "";
            string P2 = "";


            for (int i = 1; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, MDOCDATE]);
                range.Select();
                string DD = range.Text.ToString().Trim();
                if (DOCTYPE != "友福倉")
                {
                    range.NumberFormat = "yyyyMMdd";
                }

                DOCDATE = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, MITEMCODE]);
                ITEMCODE = range.Text.ToString().Trim();

                //if (ITEMCODE == "O300DVR01.01001")
                //{
                //    MessageBox.Show("A");
                //}

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, MQTY]);
                QTY = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, MINV]);
                INV = range.Text.ToString().Trim();
                //新得利倉
                if (DOCTYPE == "新得利倉")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 13]);
                    P1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 14]);
                    P2 = range.Text.ToString().Trim();
                }
                if (DOCTYPE == "深圳宏高")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 15]);
                    P1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 16]);
                    P2 = range.Text.ToString().Trim();
                }
                if (DOCTYPE == "聯揚倉")
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                    P1 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    P2 = range.Text.ToString().Trim();
                }
                //聯揚倉
                //包裝分佈明細
                int N = NAME.IndexOf("包裝分佈明細");
                if (DOCTYPE == "深圳巨航機保")
                {
                    if (N != -1)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 18]);
                        P1 = range.Text.ToString().Trim();
                    }
                }
                //深圳巨航機保
                if (!String.IsNullOrEmpty(ITEMCODE))
                {
                    if (!String.IsNullOrEmpty(QTY))
                    {
                        int num1;
                        if (int.TryParse(QTY, out num1) == true)
                        {
                            DOCDATE = DOCDATE.Replace(".", "");
                            AddAUOGD4(DOCTYPE, ITEMCODE, DOCDATE, Convert.ToInt32(QTY), INV,COMPANY,P1,P2);


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


            UPWH_WEBSTOCK2(NAME);

            dataGridView3.DataSource = GetWH_WEBSTOCK2();


        }
        private void GetExcelContentGD45(string ExcelFile)
        {

            UPWH_WEBSTOCK4();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (iRowCnt > 3000)
            {
                iRowCnt = 3000;
            }

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string ITEMCODE = "";

            string INV = "";

            string P1 = "";


            for (int i = 1; i <= iRowCnt; i++)
            {



                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                ITEMCODE = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                INV = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 18]);
                P1 = range.Text.ToString().Trim();



                //深圳巨航機保
                if (!String.IsNullOrEmpty(ITEMCODE))
                {
                    if (!String.IsNullOrEmpty(INV))
                    {

                        if (!String.IsNullOrEmpty(P1))
                        {
                            System.Data.DataTable  G5 = Get6(ITEMCODE, INV);
                            if (G5.Rows.Count > 0)
                            {
                                string ID = G5.Rows[0][0].ToString();
                                UPWH_WEBSTOCK3(P1+" ", ID);
                            }


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
        public void AddAUOGD4(string DOCTYPE, string ITEMCODE, string DOCDATE, int QTY, string INVOICE, string COMPANY, string P1, string P2)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_WEBSTOCK(DOCTYPE,ITEMCODE,DOCDATE,QTY,INVOICE,COMPANY,P1,P2) values(@DOCTYPE,@ITEMCODE,@DOCDATE,@QTY,@INVOICE,@COMPANY,@P1,@P2)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@P1", P1));
            command.Parameters.Add(new SqlParameter("@P2", P2));
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
        public void AddWH_WEBSTOCK2(string COMPANY, string FILENAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_WEBSTOCK2(COMPANY,FILENAME) values(@COMPANY,@FILENAME)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
 
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

        public void TRUNWH_WEBSTOCK2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE WH_WEBSTOCK2", connection);
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
        public void UPWH_WEBSTOCK2(string FILENAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_WEBSTOCK2 SET STATUS='已完成' WHERE FILENAME=@FILENAME ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));

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
        public void UPWH_WEBSTOCK3(string P1, string ID)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_WEBSTOCK SET P1=P1+@P1 WHERE ID=@ID ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@P1", P1));
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
        public void UPWH_WEBSTOCK4()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_WEBSTOCK SET P1='' WHERE  DOCTYPE='深圳巨航機保' ", connection);
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
        public void TRUNCATE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE WH_WEBSTOCK ", connection);
            command.CommandType = CommandType.Text;


            //command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            //command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));

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
        public void DELETE(string DOCTYPE, string ITEMCODE, string COMPANY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE WH_WEBSTOCK WHERE DOCTYPE=@DOCTYPE AND ITEMCODE=@ITEMCODE AND COMPANY=@COMPANY", connection);
            command.CommandType = CommandType.Text;



            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
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
        private void WH_WEBSTOCK_Load(object sender, EventArgs e)
        {
                    string d = @"\\acmesrv01\Public\進出貨序號\倉庫庫存表";
            dataGridView1.DataSource = Get3();
            DD2(d,0);
            dataGridView3.DataSource = TempDt2;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("儲位數量", typeof(string));
            dt.Columns.Add("ERP數量", typeof(string));
            return dt;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Get3();
        }


      
    }
}

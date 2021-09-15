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
using System.Net.Mail;
using System.Reflection;
using System.Collections;
using System.Net.Mime;

namespace ACME
{
    public partial class CHECKPACK : Form
    {
        string OBJ = "15";
        int s1 = 0;
        int inint = 0;
        int H = 0;
        int HS = 0;
        int H2 = 0;
        private System.Data.DataTable TempDt;
        private System.Data.DataTable TempDt2;
        private System.Data.DataTable TempDtS;
        public CHECKPACK()
        {
            InitializeComponent();
        }

        private System.Data.DataTable MakeTableS()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("Mo", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Mo"];
            dt.PrimaryKey = colPk;

            return dt;
        }

        private System.Data.DataTable Getdata2(string SHIPPINGCODE, string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT GRADE,INVOICE FROM AcmeSqlSP.DBO.WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT GRADE,INVOICE FROM AcmeSqlSPINFINITE.DBO.WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT GRADE,INVOICE FROM AcmeSqlSPCHOICE.DBO.WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETWHNO1(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" 			 SELECT U_TMODEL MODEL,U_PARTNO PARTNO,U_TMODEL+'.'+U_VERSION  MODEL2  FROM ACMESQL02.DBO.OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GETWHNO2(string CARTON_NO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MODEL_NO MODEL,PART_NO PARTNO,INVOICE_NO INVOICE FROM WH_AUO WHERE CARTON_NO =@CARTON_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARTON_NO", CARTON_NO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETWHNO2P(string PALLET_NO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MODEL_NO MODEL,PART_NO PARTNO,INVOICE_NO INVOICE FROM WH_AUO WHERE PALLET_NO =@PALLET_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PALLET_NO", PALLET_NO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETMODEL(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_TMODEL+'.'+U_VERSION  FROM ACMESQL02.DBO.OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETWHNO3(string U_ACME_INV)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE FROM PDN1 WHERE DOCENTRY IN (SELECT DOCENTRY  FROM OPDN WHERE U_ACME_INV=@U_ACME_INV) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETWHQTY(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(CAST(QUANTITY AS INT)),0) QTY FROM WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GETWHQTYIP(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(CAST(QUANTITY AS INT)),0) QTY FROM AcmeSqlSPINFINITE.DBO.WH_ITEM  WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GETWHQTYCC(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(CAST(QUANTITY AS INT)),0) QTY FROM AcmeSqlSPCHOICE.DBO.WH_ITEM  WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata21(string SHIPPINGCODE, string ITEMCODE, int quantity)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT INVOICE FROM WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE and quantity=@quantity ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@quantity", quantity));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable Getdata3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE FROM WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata3IP(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE FROM AcmeSqlSPINFINITE.DBO.WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable Getdata3CC(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE FROM AcmeSqlSPCHOICE.DBO.WH_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇要檢查檔案");
                return;
            }
       
            DD();


            MessageBox.Show("檢查完成");

        }

        private void GetExcelProduct(string ExcelFile, string FILE, string TYPE, string DIRNAME,int FLAG,string  FILENAME)
        {
            int CHECK = 0;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false ;
            excelApp.DisplayAlerts = false;

            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCntS = 0;
            int iRowCntE = 0;
            if (FLAG == 0)
            {
                iRowCntS = 2;
                iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;
            }
            if (FLAG == 1)
            {
                iRowCntS = 2;
                iRowCntE = H - 1;
            }
            if (FLAG == 2)
            {
                if (H2 == 3)
                {
                    iRowCntS = H;
                    iRowCntE = HS - 1;
                }
                else
                {
                        iRowCntS = H;
                        iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;  
                }
            }
            if (FLAG == 3)
            {
                iRowCntS = HS;
                iRowCntE = excelSheet.UsedRange.Cells.Rows.Count;
            }

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string WHNO = "";
                string ITEMCODE;
                string DOCDATE = "";
                string GRADE = "";
                string ITEMCODE2 = "";
                string ITEMCODE3 = "";
                string PACKNO = "";
                string PACKNO2 = "";
                string SERNO = "";
                string PARTNO = "";
                string DPARTNO = "";
                string PART = "";
                string PART2 = "";
                string PART3 = "";
                string QTY = "";
                string QTYOUT = "";
                string INV = "";
                string YEAR = "";
                string MON = "";
                string id1x = "";
                string id2x = "";
                string id3x = "";
                string CARTNO = "";
                string CARTNO2 = "";
                string PLT = "";

                string PQTY = "";
                string KQTY = "";
                string KQTY2 = "";
                string lOUT = "";
                string lOUB = "";
                string lOUD = "";
                string CLOC = "";
                string CT = "";
                string CBF = "";
                string CGADF = "";
                string lOUF = "";
                string CGAB = "";

                
                int CHECK1 = 0;
                DataRow dr;
                DataRow drS;

                int O1 = 0;
                int O2 = 0;
                int O3 = 0;
                int O4 = 0;
                string PACKNO3 = "";
                int n;

                int FS = 0;
                if (globals.GroupID.ToString().Trim() != "EEP")
                {
                    DELINVIN(FILE);
                }
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 12]);
                                WHNO = range.Text.ToString().Trim();


                                if (WHNO.ToUpper().IndexOf("WH") != -1)
                                {
                                    TYPE = "內";
                                }
                if (TYPE == "內")
                {

                    DOCDATE = FILE.Substring(1, 7);
                    YEAR = FILE.Substring(1, 3);
                    MON = FILE.Substring(4, 4);
                    int D1 = Convert.ToInt32(DOCDATE);
                    if (D1 > 1100131)
                    {
                        FS = 1;
                    }
                    DOCDATE = (Convert.ToInt16(YEAR) + 1911).ToString() + MON;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 11 + FS]);
                    if (FLAG == 2)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[H, 11 + FS]);
                    }
                    if (FLAG == 3)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[HS, 11 + FS]);
                    }
                    // range.Select();
                    WHNO = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 12 + FS]);
                    PQTY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 13 + FS]);
                    KQTY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 14 + FS]);
                    lOUT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 15 + FS]);
                    lOUD = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 16 + FS]);
                    lOUB = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 17 + FS]);
                    KQTY2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 18 + FS]);
                    CLOC = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 19 + FS]);
                    CT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 20 + FS]);
                    CBF = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 21 + FS]);
                    CGADF = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 22 + FS]);
                    lOUF = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 23 + FS]);
                    CGAB = range.Text.ToString().Trim();

                    if (KQTY == "" || KQTY == "0")
                    {
                        KQTY = KQTY2;
                    }

                    if (CLOC == "0")
                    {
                        CLOC = "";
                    }
                    UPDATEFEE(PQTY, KQTY, lOUT, lOUD, lOUB, CLOC, CT, CBF, CGADF, lOUF, CGAB, KQTY2,WHNO);
                }
                if (TYPE == "外")
                {
                    DOCDATE = FILE.Substring(1, 8);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 7]);
                    // range.Select();
                    WHNO = range.Text.ToString().Trim();
                }
                if (!String.IsNullOrEmpty(WHNO))
                {
                    //第一行要
                    for (int iRecord = iRowCntS; iRecord <= iRowCntE; iRecord++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        // range.Select();
                        ITEMCODE = range.Text.ToString().Trim();

                        if (!String.IsNullOrEmpty(ITEMCODE))
                        {
                            if (ITEMCODE.Length > 12)
                            {
                                PART = ITEMCODE.Substring(11, 3);
                            }
                        }



                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            PARTNO = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                            QTYOUT = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            INV = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                            CARTNO = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8 + FS]);
                            CARTNO2 = range.Text.ToString().Trim();


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10 + FS]);
                            QTY = range.Text.ToString().Trim();

                            if (FS == 1)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                                PLT = range.Text.ToString().Trim();
                            }

                        }
                        if (TYPE == "外")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                            PARTNO = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            INV = range.Text.ToString().Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                            QTY = range.Text.ToString().Trim();
                        }

                        if (ITEMCODE != "")
                        {
                            id1x = ITEMCODE;
                        }
                        if (ITEMCODE == "")
                        {
                            ITEMCODE = id1x;
                        }

                        if (PARTNO != "")
                        {
                            id2x = PARTNO;
                        }
                        if (PARTNO == "")
                        {
                            PARTNO = id2x;
                        }


                        if (INV != "")
                        {
                            id3x = INV;
                        }
                        if (INV == "")
                        {
                            INV = id3x;
                        }

                        if (!String.IsNullOrEmpty(PARTNO))
                        {
                            if (PARTNO.Length > 11)
                            {
                                PART2 = PARTNO.Substring(9, 3);
                                PART3 = PARTNO.Substring(10, 2);
                            }

                        }



                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8 + FS]);
                            PACKNO = range.Text.ToString().Trim().Replace("*", "").ToUpper();
                        }
                        if (TYPE == "外")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                            PACKNO = range.Text.ToString().Trim().Replace("*", "").ToUpper();
                        }
           
                        if (!String.IsNullOrEmpty(PACKNO))
                        {
                            if (PACKNO.Length > 6)
                            {
                                PACKNO2 = PACKNO.Substring(4, 2);

                            }
                            DPARTNO = PACKNO;
                        }


                        if (TYPE == "內")
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9 + FS]);
                            // range.Select();
                            SERNO = range.Text.ToString().Trim();
                        }


                        string CARDNAME = "";
                        string WAREHOUSE = "";
                        System.Data.DataTable t1 = Getdata2(WHNO);
                        if (t1.Rows.Count > 0)
                        {

                            CARDNAME = t1.Rows[0]["CARDNAME"].ToString();
                            WAREHOUSE = t1.Rows[0]["WAREHOUSE"].ToString();
                        }

                        if (!String.IsNullOrEmpty(QTY))
                        {
                            if (PACKNO.Trim() != "片數總計" && SERNO.Trim() != "片數總計" && PACKNO.Trim() != "片數總計:")
                            {
                                string DDATE = textBox1.Text.Substring(0, 6);
                                O4 += Convert.ToInt32(QTY);
                                //if (globals.GroupID.ToString().Trim() != "EEP")
                                //{
                                    if (String.IsNullOrEmpty(PACKNO))
                                    {
                                        if (String.IsNullOrEmpty(PLT))
                                        {
                                            PACKNO = DPARTNO;
                                        }
                                    }

                                    if (!String.IsNullOrEmpty(ITEMCODE) && (!String.IsNullOrEmpty(PACKNO)||!String.IsNullOrEmpty(PLT)))
                                    {

                                        //12345
                                        string AUOMODEL = "";
                                        string AUOPARTNO = "";
                                        string INVOICE = "";
                                        string SAPMODEL = "";
                                        string SAPMODEL2 = "";
                                        string SAPPARTNO = "";
                                        int G2 = 0;
                                        System.Data.DataTable K1 = GETWHNO1(ITEMCODE);
                                        if (K1.Rows.Count > 0)
                                        {
                                            SAPMODEL = K1.Rows[0][0].ToString();
                                            SAPPARTNO = K1.Rows[0][1].ToString();
                                            SAPMODEL2 = K1.Rows[0][2].ToString();
                                            System.Data.DataTable K2 = null;
                                            if (!String.IsNullOrEmpty(PACKNO))
                                            {
                                                K2 = GETWHNO2(PACKNO);
                                            }
                                            else
                                            {
                                                K2 = GETWHNO2P(PLT);
                                            }
                                            
                                            if (K2.Rows.Count > 0)
                                            {
                                                AUOMODEL = K2.Rows[0][0].ToString();
                                                AUOPARTNO = K2.Rows[0][1].ToString();
                                                INVOICE = K2.Rows[0][2].ToString();


                                            if (!String.IsNullOrEmpty(SAPPARTNO) && !String.IsNullOrEmpty(AUOPARTNO) && !String.IsNullOrEmpty(AUOMODEL))
                                            {
                                                if (SAPPARTNO != AUOPARTNO)
                                                {
                                                    string G1 = ITEMCODE.Substring(0, 1).ToUpper() + AUOMODEL.Substring(0, 1).ToUpper();
                                                    string G1S = SAPPARTNO.Substring(0, 2).ToUpper() + AUOPARTNO.Substring(0, 2).ToUpper();

                                                    if ((G1 == "OM" || G1 == "MO" || G1 == "OO") && (G1S == "9197" || G1S == "9791"))
                                                    {

                                                    }

                                                   else  if ((G1 == "OM" || G1 == "MO" || G1 == "OO") && (G1S == "9197" || G1S == "9791"))
                                                    {

                                                    }
                                                    else
                                                    {
                                                        dr = TempDt.NewRow();
                                                        dr["倉庫"] = DIRNAME;
                                                        dr["EXCEL"] = FILE;
                                                        dr["工單號碼"] = WHNO;
                                                        dr["檢查結果"] = "跟原廠比對PARTNO有誤 原廠 " + AUOPARTNO + "倉庫 " + SAPPARTNO;
                                                        dr["A"] = excelFile;
                                                        dr["B"] = FILENAME;
                                                        TempDt.Rows.Add(dr);
                                                        CHECK = 1;
                                                    }
                                                }
                                                else
                                                {
                                                    if (!String.IsNullOrEmpty(AUOMODEL))
                                                    {
                                                        string G1 = ITEMCODE.Substring(0, 1).ToUpper() + AUOMODEL.Substring(0, 1).ToUpper();
                                                        if (G1 == "OM" || G1 == "MO")
                                                        {
                                                            G2 = 1;
                                                        }
                                                    }
                                                }
                                            }                                                if (String.IsNullOrEmpty(AUOPARTNO))
                                                {
                                                    if (!String.IsNullOrEmpty(AUOMODEL))
                                                    {
                                                        string G1 = ITEMCODE.Substring(0, 1).ToUpper() + AUOMODEL.Substring(0, 1).ToUpper();
                                                        if (G1 == "OM" || G1 == "MO")
                                                        {
                                                            System.Data.DataTable K3 = GETWHNO3(INVOICE);

                                                            if (K3.Rows.Count > 0)
                                                            {
                                                                if (ITEMCODE == K3.Rows[0][0].ToString())
                                                                {
                                                                    G2 = 1;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                //型號應為M270HTN02.5 比對後為M270DTN01.100  S/N:NZ19001721504057


                                                if (G2 == 0)
                                                {
                                                    if (!String.IsNullOrEmpty(SAPMODEL) && !String.IsNullOrEmpty(AUOMODEL))
                                                    {
                                                   
                                                        int TT2 = AUOMODEL.IndexOf(".");
                                                        if (TT2 != -1)
                                                        {
                                                            SAPMODEL = SAPMODEL2;
                                                        }
                                                        string F1 = ITEMCODE.Substring(0, 1).ToUpper() + AUOMODEL.Substring(0, 1).ToUpper();
                                                        if (F1 == "TO" || F1 == "OT" || F1 == "OP" || F1 == "OM" || F1 == "MO")
                                                    {
                                                            SAPMODEL = SAPMODEL.Substring(1, SAPMODEL.Length - 1);
                                                        }
                                                        int TT = AUOMODEL.IndexOf(SAPMODEL);
                                                        if (TT == -1)
                                                        {


                                                            dr = TempDt.NewRow();
                                                            dr["倉庫"] = DIRNAME;
                                                            dr["EXCEL"] = FILE;
                                                            dr["工單號碼"] = WHNO;
                                                            dr["檢查結果"] = "型號應為 " + SAPMODEL + " 比對後為 " + AUOMODEL + " S/N:" + PACKNO;
                                                            dr["A"] = excelFile;
                                                            dr["B"] = FILENAME;
                                                            TempDt.Rows.Add(dr);
                                                            CHECK = 1;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (CHECK != 1)
                                        {
                                            UPDATEAUO(ITEMCODE, PACKNO);
                                            System.Data.DataTable GM = GETMODEL(ITEMCODE);
                                            if (GM.Rows.Count > 0)
                                            {
                                                UPDATEAUO2(GM.Rows[0][0].ToString(), PACKNO);
                                            }
                                        }
                                    }


                                    AddINVIN(ITEMCODE, FILE, PARTNO, INV, PACKNO, SERNO, QTY, CARDNAME, "", DOCDATE, DateTime.Now.ToString("yyyyMMdd"), WAREHOUSE, WHNO, DDATE, TYPE, fmLogin.LoginID.ToString(), PLT);
                        //   }
                            }



                            int B1 = PARTNO.IndexOf("總計");
                            int B2 = PARTNO.IndexOf("共計");
                            int B3 = SERNO.IndexOf("總計");
                            int B4 = SERNO.IndexOf("共計");
                            int B7 = SERNO.IndexOf("片數");
                            int B5 = PACKNO.IndexOf("總計");
                            int B6 = PACKNO.IndexOf("共計");
                            if (B1 != -1 || B2 != -1 || B3 != -1 || B4 != -1 || B5 != -1 || B6 != -1 || B7 != -1)
                            {
                                CHECK1 = 1;
                            }

                            System.Data.DataTable H1 = Getdata2(WHNO, ITEMCODE);
                            System.Data.DataTable H3 = null;
                          
                            if (WHNO.IndexOf("IP") != -1)
                            {
                                H3 = Getdata3IP(WHNO);
                            }
                            else if (WHNO.IndexOf("CC") != -1)
                            {
                                H3 = Getdata3CC(WHNO);
                            }
                            else
                            {
                                H3 = Getdata3(WHNO);
                            }
                                if (H3.Rows.Count > 0)
                            {
                                ITEMCODE3 = GRADE;
                                if (!String.IsNullOrEmpty(ITEMCODE))
                                {

                                    if (H1.Rows.Count == 0)
                                    {


                                        dr = TempDt.NewRow();
                                        dr["倉庫"] = DIRNAME;
                                        dr["EXCEL"] = FILE;
                                        dr["工單號碼"] = WHNO;
                                        dr["檢查結果"] = "工單沒有此料號" + ITEMCODE;
                                        dr["A"] = excelFile;
                                        dr["B"] = FILENAME;
                                        TempDt.Rows.Add(dr);
                                        CHECK = 1;
                                    }
                                    else
                                    {
                                        GRADE = H1.Rows[0][0].ToString().Trim();

                                        if (GRADE == "NN")
                                        {
                                            GRADE = "N";
                                        }
                                        ITEMCODE2 = GRADE;
                                        ITEMCODE3 = GRADE;
                                        if (globals.DBNAME != "宇豐")
                                        {
                                            string K = ITEMCODE.Substring(0, 1).ToUpper();
                                            if (K != "K" && K != "A")
                                            {
                                                if (PART != PART2)
                                                {
                                                    if (PACKNO != "9AQ23807-011")
                                                    {
                                                        if (CHECK1 == 0)
                                                        {
                                                            dr = TempDt.NewRow();
                                                            dr["倉庫"] = DIRNAME;
                                                            dr["EXCEL"] = FILE;
                                                            dr["工單號碼"] = WHNO;
                                                            dr["檢查結果"] = "小料異常 : " + PARTNO;
                                                            dr["A"] = excelFile;
                                                            dr["B"] = FILENAME;
                                                            TempDt.Rows.Add(dr);
                                                            CHECK = 1;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                                         string K2 = ITEMCODE.Substring(0, 1).ToUpper();
                                         if (K2 != "K" && K2 != "A")
                                         {
                                             if (!String.IsNullOrEmpty(PACKNO))
                                             {
                                                 string PACKNO1 = PACKNO.Substring(0, 1);
                                                 if (ITEMCODE == "TPAU31901.00001")
                                                 {
                                                     PACKNO1 = PACKNO.Substring(PACKNO.Length - 5, 1);
                                                 }
                                                 if (PACKNO.Length > 1)
                                                 {
                                                     string PACKNO12 = PACKNO.Substring(1, 1);
                                                     if (PACKNO1.ToUpper() != GRADE.ToUpper())
                                                     {
                                                         int PP1 = PACKNO.IndexOf("-");
                                                         if (PP1 == -1)
                                                         {
                                                             if (PACKNO != "条码扫取不了")
                                                             {
                                                    
                                                                 if (CHECK1 == 0)
                                                                 {
                                                                     dr = TempDt.NewRow();
                                                                     dr["倉庫"] = DIRNAME;
                                                                     dr["EXCEL"] = FILE;
                                                                     dr["工單號碼"] = WHNO;
                                                                     dr["檢查結果"] = "箱序異常 : " + PACKNO + " EXCEL " + PACKNO1 + " 規 SAP " + GRADE + " 規";
                                                                     dr["A"] = excelFile;
                                                                     dr["B"] = FILENAME;
                                                                     TempDt.Rows.Add(dr);
                                                                     CHECK = 1;
                                                                 }
                                                             }
                                                         }
                                                     }

                                                     if (WHNO != "WH20210322001X-IP")
                                                     {
                                                         if (PACKNO12.ToUpper() == "R")
                                                         {
                                                             dr = TempDt.NewRow();
                                                             dr["倉庫"] = DIRNAME;
                                                             dr["EXCEL"] = FILE;
                                                             dr["工單號碼"] = WHNO;
                                                             dr["檢查結果"] = "箱序第二碼為R : " + PACKNO;
                                                             dr["A"] = excelFile;
                                                             dr["B"] = FILENAME;
                                                             TempDt.Rows.Add(dr);
                                                             CHECK = 1;
                                                         }
                                                     }
                                                     if (PART3 != PACKNO2)
                                                     {
                                                         int PP1 = PACKNO.IndexOf("-");
                                                         if (TYPE == "外")
                                                         {
                                                             PP1 = -1;
                                                         }
                                                         if (PP1 == -1)
                                                         {
                                                             if (CHECK1 == 0)
                                                             {
                                                                 dr = TempDt.NewRow();
                                                                 dr["倉庫"] = DIRNAME;
                                                                 dr["EXCEL"] = FILE;
                                                                 dr["工單號碼"] = WHNO;
                                                                 dr["檢查結果"] = "小料比對箱序異常 : " + PACKNO;
                                                                 dr["A"] = excelFile;
                                                                 dr["B"] = FILENAME;
                                                                 TempDt.Rows.Add(dr);
                                                                 CHECK = 1;
                                                             }
                                                         }
                                                     }
                                                 }
                                                 else
                                                 {
                                                     dr = TempDt.NewRow();
                                                     dr["倉庫"] = DIRNAME;
                                                     dr["EXCEL"] = FILE;
                                                     dr["工單號碼"] = WHNO;
                                                     dr["檢查結果"] = "箱序異常 : " + PACKNO;
                                                     dr["A"] = excelFile;
                                                     dr["B"] = FILENAME;
                                                     TempDt.Rows.Add(dr);
                                                     CHECK = 1;
                                                 }
                                             }
                                         }

                                if (TYPE == "內")
                                {
                                    if (PACKNO.Trim() != "片數總計" && SERNO.Trim() != "片數總計")
                                    {
                                        if (!String.IsNullOrEmpty(PACKNO) || !String.IsNullOrEmpty(SERNO))
                                        {

                                            if(SERNO.Length >11)
                                            {
                                                if (!String.IsNullOrEmpty(SERNO))
                                                {
                                                    if (int.TryParse(QTY, out n))
                                                    {
                                                        O1 += Convert.ToInt16(QTY);
                                                    }

                                                    if (int.TryParse(QTY, out n))
                                                    {
                                                        O3 += Convert.ToInt16(QTY);
                                                    }
                                                    if (!String.IsNullOrEmpty(PACKNO))
                                                    {
                                                        string PACKNO1 = PACKNO.Substring(0, 1);
                                                        if (PACKNO1 == "Z" || PACKNO1 == "P" || PACKNO1 == "N")
                                                        {
                                                            if (int.TryParse(QTYOUT, out n))
                                                            {
                                                                O2 += Convert.ToInt16(QTYOUT);
                                                            }
                                                        }

                                                        string PP = SERNO.Substring(12, 1);
                                                        string PP2 = SERNO.Substring(6, 1);
                                                        if (PP != PACKNO1)
                                                        {
                                                            if (PP2 != "Z"&& PP2 != "P")
                                                            {
                                                                //箱/片序等級不同
                                                                dr = TempDt.NewRow();
                                                                dr["倉庫"] = DIRNAME;
                                                                dr["EXCEL"] = FILE;
                                                                dr["工單號碼"] = WHNO;
                                                                dr["檢查結果"] = "箱/片序等級不同 片序號碼 " + SERNO;
                                                                dr["A"] = excelFile;
                                                                dr["B"] = FILENAME;
                                                                TempDt.Rows.Add(dr);
                                                                CHECK = 1;
                                                            }
                                                        }

                                                    }
                                                }
                                            }

                                            else
                                            {
                                                string PACKNO1 = PACKNO.Substring(0, 1);
                                                if (PACKNO1 == "Z" || PACKNO1 == "P" || PACKNO1 == "N")
                                                {
                                                    if (int.TryParse(QTY, out n))
                                                    {
                                                        O1 += Convert.ToInt16(QTY);

                                                    }

                                                    if (int.TryParse(QTY, out n))
                                                    {
                                                        O3 += Convert.ToInt16(QTY);
                                                    }
                                                    if (int.TryParse(QTYOUT, out n))
                                                    {
                                                        O2 += Convert.ToInt16(QTYOUT);

                                                    }
                                                    PACKNO3 = PACKNO1;
                                                }

                                                if (!String.IsNullOrEmpty(PACKNO3))
                                                {
                                                    if (PACKNO3 != PACKNO1)
                                                    {
                                                        if (O2 > 0)
                                                        {
                                                            if (O1 != O2)
                                                            {
                                                                if (O1 != O3)
                                                                {

                                                                    if (CHECK != 1)
                                                                    {
                                                                        dr = TempDt.NewRow();
                                                                        dr["倉庫"] = DIRNAME;
                                                                        dr["EXCEL"] = FILE;
                                                                        dr["工單號碼"] = WHNO;
                                                                        dr["檢查結果"] = "Z/P數量異常 : " + PACKNO3 + "規 出貨數量 " + O2 + "倉庫數量 " + O1;
                                                                        dr["A"] = excelFile;
                                                                        dr["B"] = FILENAME;
                                                                        TempDt.Rows.Add(dr);
                                                                        CHECK = 1;
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        O1 = 0;
                                                        O2 = 0;
                                                    }
                                                }


                                            }
                                        }
                                        else
                                        {
                                            if (!String.IsNullOrEmpty(PACKNO3))
                                            {

                                                if (O1 != O2)
                                                {
                                                    if (O1 != O3)
                                                    {
                                                        if (O2 != 0)
                                                        {
                                                            if (CHECK != 1)
                                                            {
                                                                dr = TempDt.NewRow();
                                                                dr["倉庫"] = DIRNAME;
                                                                dr["EXCEL"] = FILE;
                                                                dr["工單號碼"] = WHNO;
                                                                dr["檢查結果"] = "Z/P數量異常 : " + PACKNO3 + "規 出貨數量 " + O2 + "倉庫數量 " + O1;
                                                                dr["A"] = excelFile;
                                                                dr["B"] = FILENAME;
                                                                TempDt.Rows.Add(dr);
                                                                CHECK = 1;
                                                            }
                                                        }
                                                    }
                                                }


                                            }
                                        }
                                    }
                                }
                                if (TYPE == "外")
                                {
                                    string S = "0";
                                    string INV2 = "";
                                    for (int i = 0; i <= H1.Rows.Count - 1; i++)
                                    {
                                         INV2 = H1.Rows[i][1].ToString().Trim();

                                         int IN1 = INV2.ToUpper().IndexOf(INV.ToUpper());
                                         if (IN1==-1)
                                        {
                                            if (S != "1")
                                            {
                                                S = "2";
                                            }
                                        }
                                        else
                                        {
                                            S = "1";
                                        }
                                    }
                                    if (S == "2")
                                    {
                                        if (CHECK1 == 0)
                                        {
                                            dr = TempDt.NewRow();
                                            dr["倉庫"] = DIRNAME;
                                            dr["EXCEL"] = FILE;
                                            dr["工單號碼"] = WHNO;
                                            //原廠INV#異常Z451658419 應為INV#Z191701612
                                            dr["檢查結果"] = "原廠INV#異常" + INV + " 應為INV#" + INV2;
                                            dr["A"] = excelFile;
                                            dr["B"] = FILENAME;
                                            TempDt.Rows.Add(dr);
                                            CHECK = 1;
                                        }
                                    }

                                }
                                if (TYPE == "內")
                                {
                                    if (!String.IsNullOrEmpty(SERNO))
                                    {
                                        string GG = ITEMCODE;
                                        GRADE = GRADE.ToUpper();
                                        int U = SERNO.ToUpper().IndexOf(GRADE);
                                        if (U == -1)
                                        {

                                            if (CHECK1 == 0)
                                            {
                                                dr = TempDt.NewRow();
                                                dr["倉庫"] = DIRNAME;
                                                dr["EXCEL"] = FILE;
                                                dr["工單號碼"] = WHNO;
                                                dr["檢查結果"] = "片序異常 : " + SERNO;
                                                dr["A"] = excelFile;
                                                dr["B"] = FILENAME;
                                                TempDt.Rows.Add(dr);
                                                CHECK = 1;
                                            }
                                        }



                                    }

                                    if (!String.IsNullOrEmpty(CARTNO2))
                                    {
                                        if (CARTNO2 != "片數總計")
                                        {
                                            drS = TempDtS.NewRow();

                                            drS["Mo"] = CARTNO2;

                                            try
                                            {
                                                TempDtS.Rows.Add(drS);
                                            }
                                            catch
                                            {

                                                dr = TempDt.NewRow();
                                                dr["倉庫"] = DIRNAME;
                                                dr["EXCEL"] = FILE;
                                                dr["工單號碼"] = WHNO;
                                                dr["檢查結果"] = "箱號 " + CARTNO + "箱序重複 " + CARTNO2;
                                                dr["A"] = excelFile;
                                                dr["B"] = FILENAME;
                                                TempDt.Rows.Add(dr);
                                                CHECK = 1;

                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                CHECK = 2;
                            }

                        }
                    }

                    System.Data.DataTable GWHQTY = null;
                    if (WHNO.IndexOf("IP") != -1)
                    {
                        GWHQTY = GETWHQTYIP(WHNO);
                    }
                    else if (WHNO.IndexOf("CC") != -1)
                    {
                        GWHQTY = GETWHQTYCC(WHNO);
                    }

                    else
                    {
                        GWHQTY = GETWHQTY(WHNO);
                    }
                    int QW = Convert.ToInt32(GWHQTY.Rows[0][0]);
                    if (QW != O4 && WHNO.Length == 14)
                    {
                        dr = TempDt.NewRow();
                        dr["倉庫"] = DIRNAME;
                        dr["EXCEL"] = FILE;
                        dr["工單號碼"] = WHNO;
                        dr["檢查結果"] = "數量不符 備貨單數量,請檢察工單號碼" + QW.ToString() + "序號數量" + O4.ToString();
                        dr["A"] = excelFile;
                        dr["B"] = FILENAME;
                        TempDt.Rows.Add(dr);
                        CHECK = 1;

                    }
                    else if(WHNO.Length != 14 )
                    {
                        dr = TempDt.NewRow();
                        dr["倉庫"] = DIRNAME;
                        dr["EXCEL"] = FILE;
                        dr["工單號碼"] = WHNO;
                        dr["檢查結果"] = "工單號碼不完整";
                        dr["A"] = excelFile;
                        dr["B"] = FILENAME;
                        TempDt.Rows.Add(dr);
                        CHECK = 1;
                    }

                    if (globals.GroupID.ToString().Trim() != "EEP")
                    {
                        EE(WHNO);
                    }
                }
                else
                {
                    CHECK = 2;
                }

                if (CHECK == 2)
                {
                    dr = TempDt.NewRow();
                    dr["倉庫"] = DIRNAME;
                    dr["EXCEL"] = FILE;
                    dr["工單號碼"] = "";
                    dr["檢查結果"] = "沒有工單號碼";
                    dr["A"] = excelFile;
                    dr["B"] = FILENAME;
                    TempDt.Rows.Add(dr);
                }
                if (CHECK == 0)
                {
                    dr = TempDt.NewRow();
                    dr["倉庫"] = DIRNAME;
                    dr["EXCEL"] = FILE;
                    dr["工單號碼"] = WHNO;
                    dr["檢查結果"] = "沒問題";
                    dr["A"] = excelFile;
                    dr["B"] = FILENAME;
                    TempDt.Rows.Add(dr);


                    if (globals.GroupID.ToString().Trim() != "EEP")
                    {
                        INS(excelFile, WHNO, FILENAME);
                    }
                   
                }


                dataGridView1.DataSource = TempDt;


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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

           
            }
 

            //  dataGridView1.DataSource = TempDt;
        }

        private void INS(string ExcelFile, string WHNO, string FILENAME)
        {

            System.Data.DataTable F1 = GEWHMAIN(WHNO);
            if (F1.Rows.Count > 0)
            {
                string FNAME = FILENAME.Replace("序", "序" + WHNO + "-");
                FNAME = FNAME.Replace(" ", "");
                string DIR = "//acmesrv01//Public//進出貨序號//理貨資料//全部//" + FNAME;
                System.IO.File.Copy(ExcelFile, DIR, true);
            }

            System.Data.DataTable F2 = GEOWRT1(WHNO);

            if (F2.Rows.Count > 0)
            {
                for (int i = 0; i <= F2.Rows.Count - 1; i++)
                {
                    string DOC = F2.Rows[i][0].ToString();
                    string FNAME = FILENAME.Replace("序", "調" + DOC + "-");

                    string DIR = "//acmesrv01//Public//進出貨序號//理貨資料//全部//" + FNAME;
                    System.IO.File.Copy(ExcelFile, DIR, true);

                    System.Data.DataTable F3 = GEOWRT2(WHNO);
                    string OWINV = F3.Rows[0][0].ToString();
                    if (!String.IsNullOrEmpty(OWINV))
                    {

                        UPOWTR(OWINV, DOC);

                    }
                }
            }
        }

        public void UPDATEFEE(string PQTY, string KQTY, string lOUT, string lOUD, string lOUB, string CLOC, string CT, string CBF, string CGADF, string lOUF, string CGAB, string KQTY3, string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_FEE SET PQTY=@PQTY,KQTY=@KQTY,lOUT=@lOUT,lOUD=@lOUD,lOUB=@lOUB,CLOC=@CLOC,CT=@CT,CBF=@CBF,CGADF=@CGADF,lOUF=@lOUF,CGAB=@CGAB,KQTY3=@KQTY3 WHERE SHIPPINGCODE=@SHIPPINGCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PQTY", PQTY));
            command.Parameters.Add(new SqlParameter("@KQTY", KQTY));
            command.Parameters.Add(new SqlParameter("@lOUT", lOUT));
            command.Parameters.Add(new SqlParameter("@lOUD", lOUD));
            command.Parameters.Add(new SqlParameter("@lOUB", lOUB));
            command.Parameters.Add(new SqlParameter("@CLOC", CLOC));
            command.Parameters.Add(new SqlParameter("@CT", CT));
            command.Parameters.Add(new SqlParameter("@CBF", CBF));
            command.Parameters.Add(new SqlParameter("@CGADF", CGADF));
            command.Parameters.Add(new SqlParameter("@lOUF", lOUF));
            command.Parameters.Add(new SqlParameter("@CGAB", CGAB));
            command.Parameters.Add(new SqlParameter("@KQTY3", KQTY3));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            
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

  

        private System.Data.DataTable GetMAXOCLG()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(CLGCODE)+1 ID FROM OCLG ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



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


        public void UPOWTR(string U_ACME_INV, string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE OWTR SET U_ACME_INV=@U_ACME_INV where DOCENTRY=@DOCENTRY ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
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
 
        public void UPONNM(int AUTOKEY, string ObjectCode)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE ONNM SET AUTOKEY=@AUTOKEY WHERE ObjectCode=@ObjectCode", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AUTOKEY", AUTOKEY));
            command.Parameters.Add(new SqlParameter("@ObjectCode", ObjectCode));

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
        private System.Data.DataTable GETSS(string AA)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ABSENTRY FROM ATC1 where [FILENAME]=@AA ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", AA));


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
        private System.Data.DataTable GEWHMAIN(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND ForecastDay ='銷售訂單' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


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

        private System.Data.DataTable GEOWRT1(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT DISTINCT T1.DOCENTRY  PINO FROM WH_MAIN T0 LEFT JOIN WH_ITEM4 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0. SHIPPINGCODE=@SHIPPINGCODE AND ForecastDay ='庫存調撥-撥倉'");
        //    sb.Append(" SELECT PINO FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND ForecastDay ='庫存調撥-撥倉'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


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

        private System.Data.DataTable GEOWRT2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @name2 varchar(200) ");
            sb.Append(" select @name2 =SUBSTRING(COALESCE(@name2 + '/',''),0,200) + INV ");
            sb.Append(" from   (");
            sb.Append(" SELECT DISTINCT  INV  FROM ACMESQLSP.DBO.AP_INVOICEIN WHERE WHNO=@SHIPPINGCODE) PV");
            sb.Append(" SELECT ISNULL(@name2,'') INV");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


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
        private void GetExcelProduct2(string ExcelFile, string FILE, string TYPE, string DIRNAME)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);


            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string WHNO = "";

                int FS = 0;
 
                if (TYPE == "內")
                {
                   string DOCDATE = FILE.Substring(1, 7);

                    int D1 = Convert.ToInt32(DOCDATE);
                    if (D1 > 1100131)
                    {
                        FS = 1;
                    }
  
                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11 + FS]);
                        WHNO = range.Text.ToString().Trim();
  
                  
                        int F1 = WHNO.ToUpper().IndexOf("WH");
                        if (F1 != -1)
                        {

                            WHNO = range.Text.ToString().Trim();
                            if (WHNO.Length == 14)
                            {
                                H2++;
                                if (H2 == 2)
                                {
                                    H = iRecord;
                                }
                                if (H2 == 3)
                                {
                                    HS = iRecord;
                                }

                            }
                        }
                    }
                    // range.Select();
              
                }
         


      
                dataGridView1.DataSource = TempDt;


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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


            //  dataGridView1.DataSource = TempDt;
        }
   
        private void CHECKPACK_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "國內";
            textBox1.Text = GetMenu.Day();

        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("EXCEL", typeof(string));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("檢查結果", typeof(string));
            dt.Columns.Add("A", typeof(string));
            dt.Columns.Add("B", typeof(string));


            return dt;
        }

        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("fie", typeof(string));
            dt.Columns.Add("PanelName", typeof(string));
            dt.Columns.Add("IN", typeof(string));
            dt.Columns.Add("DIRNAME", typeof(string));
            dt.Columns.Add("FILENAME", typeof(string));

            return dt;
        }
        public void DD()
        {
            TempDtS = MakeTableS();
            TempDt = MakeTable();

            for (int h = dataGridView2.SelectedRows.Count - 1; h >= 0; h--)
            {
                TempDtS.Clear();
                DataGridViewRow row;
                row = dataGridView2.SelectedRows[h];


                string fie = row.Cells["fie"].Value.ToString();
                string PanelName = row.Cells["PanelName"].Value.ToString();
                string IN = row.Cells["IN"].Value.ToString();
                string DIRNAME = row.Cells["DIRNAME"].Value.ToString();
                string FILENAME = row.Cells["FILENAME"].Value.ToString();
                if (IN == "內")
                {
                    H2 = 0;
                    GetExcelProduct2(fie, PanelName, IN, DIRNAME);
                    if (H2 > 1)
                    {
                        if (H2 == 2)
                        {
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 1, FILENAME);
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 2, FILENAME);
                        }
                        if (H2 == 3)
                        {
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 1, FILENAME);
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 2, FILENAME);
                            GetExcelProduct(fie, PanelName, IN, DIRNAME, 3, FILENAME);
                        }
                    }
                    else
                    {
                        GetExcelProduct(fie, PanelName, IN, DIRNAME, 0, FILENAME);
                    }
                }
                else
                {
                    GetExcelProduct(fie, PanelName, IN, DIRNAME, 0, FILENAME);
                }

            }
            
        }

        public void DD2(string PATH)
        {


            string[] filebBrand = Directory.GetDirectories(PATH);
            foreach (string fileabBrand in filebBrand)
            {
                DirectoryInfo DIRINFO = new DirectoryInfo(fileabBrand);

                string DIRNAME = DIRINFO.Name.ToString();
                string IN = DIRNAME.Substring(0, 1);
                string IN2 = comboBox1.Text.Substring(1, 1);
                if (IN == IN2)
                {
                    string YEAR = DIRNAME.Substring(1, 4);
                    string YEAR2 = textBox1.Text.Substring(0, 4);
                    if (YEAR == YEAR2)
                    {
                        string[] fileccVer = Directory.GetDirectories(fileabBrand);
                        foreach (string filee in fileccVer)
                        {
                            DirectoryInfo DIRINFO2 = new DirectoryInfo(filee);
                            string DIRNAME2 = DIRINFO2.Name.ToString();
                            int G1 = DIRNAME2.IndexOf("月");
                            string MONTH = DIRNAME2.Substring(0, G1);
                            string MONTH2 = Convert.ToInt16(textBox1.Text.Substring(4, 2)).ToString();
                            if (MONTH == MONTH2)
                            {
                                string[] filecSize = Directory.GetFiles(filee);
                                foreach (string fie in filecSize)
                                {
                                    int aa = fie.LastIndexOf(".");
                                    string Type;
                                    Type = fileabBrand.Replace(PATH, "");
                                    DataRow dr;
                                    FileInfo filess = new FileInfo(fie);
                                    string dd = filess.Name.ToString();

                                    int ad = dd.LastIndexOf(".");

                                    string size = filess.Length.ToString();
                                    string FileDate = "";
                                    int P1 = dd.IndexOf("序");
                                    if (ad != -1)
                                    {
                                        if (P1 != -1)
                                        {

                                            if (IN == "內")
                                            {
                                                int F1 = Convert.ToInt16(dd.Substring(P1 + 1, 3)) + 1911;
                                                FileDate = F1.ToString() + dd.Substring(P1 + 4, 4);
                                            }

                                            if (IN == "外")
                                            {
                                                FileDate = dd.Substring(P1 + 1, 8);
                                            }
                                            if (FileDate == textBox1.Text)
                                            {
                                                string PanelName = dd.Substring(0, ad).ToString();


                                                if (PanelName != "Thumbs")
                                                {
                                                    if (PanelName != "序1031001宜春_平鎮")
                                                    {
                                                        dr = TempDt2.NewRow();
                                                        dr["fie"] = fie;
                                                        dr["PanelName"] = PanelName;
                                                        dr["IN"] = IN;
                                                        dr["DIRNAME"] = DIRNAME;
                                                        dr["FILENAME"] = dd;
                                                        TempDt2.Rows.Add(dr);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }

                            }
                        }
                    }
                }
            }


        }
        public void EE(string WHNO)
        {
            StringBuilder sb2 = new StringBuilder();
            System.Data.DataTable dt2 = GetAUINV3(WHNO);
            if (dt2.Rows.Count > 0)
            {
                for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                {

                    DataRow dd = dt2.Rows[i];


                    sb2.Append(dd["INV"].ToString() + "/");


                }

                sb2.Remove(sb2.Length - 1, 1);

                UPDATEWH(sb2.ToString(), WHNO);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TempDt2 = MakeTable2();
            string ACME = "//acmesrv01//Public//進出貨序號//出貨序號";
            string DRS = "//szcnfs01//Public//出货序号";
            string DRS2 = "//szcnfs01//Public//DRS出货序号";
            if (globals.DBNAME == "達睿生")
            {
                DD2(DRS2);
            }
            else
            {

                DD2(ACME);

                if (checkBox1.Checked == false)
                {
                    try
                    {
                        DD2(DRS);
                    }
                    catch
                    {
                        MessageBox.Show("達睿生Public無法連線");
                    }
                }
            }


            dataGridView2.DataSource = TempDt2;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView2.Rows.Count - 2; i++)
            {
                dataGridView2.Rows[i].Selected = true;

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView2.Rows.Count - 2; i++)
            {
                dataGridView2.Rows[i].Selected = false;
            }
        }
        public void AddINVIN(string ITEMCODE, string FILE, string PARTNO, string INV, string CARTON, string PIC, string QTY, string CARD, string SAPDOC, string DOCDATE, string INSERTDATE, string WAREHOUSE, string WHNO, string DDATE, string INOUT, string USERS, string PLT)
        {
       

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into AP_INVOICEIN(ITEMCODE,ITEMNAME,PARTNO,INV,CARTON,PIC,QTY,CARD,SAPDOC,DOCDATE,INSERTDATE,WAREHOUSE,WHNO,DDATE,INOUT,USERS,PLT) values(@ITEMCODE,@ITEMNAME,@PARTNO,@INV,@CARTON,@PIC,@QTY,@CARD,@SAPDOC,@DOCDATE,@INSERTDATE,@WAREHOUSE,@WHNO,@DDATE,@INOUT,@USERS,@PLT)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", FILE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@PIC", PIC));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));
            command.Parameters.Add(new SqlParameter("@SAPDOC", SAPDOC));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@INSERTDATE", INSERTDATE));
            command.Parameters.Add(new SqlParameter("@WAREHOUSE", WAREHOUSE));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@DDATE", DDATE));
            command.Parameters.Add(new SqlParameter("@INOUT", INOUT));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            //USERS
            //DDATE
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + FILE);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void UPDATEAUO(string ITEMCODE,  string CARTON_NO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_AUO SET ITEMCODE=@ITEMCODE   WHERE CARTON_NO  =@CARTON_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            command.Parameters.Add(new SqlParameter("@CARTON_NO", CARTON_NO));
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
        public void UPDATEAUO2(string MODEL_NO, string CARTON_NO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_AUO SET MODEL_NO=@MODEL_NO  WHERE CARTON_NO  =@CARTON_NO ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@CARTON_NO", CARTON_NO));
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
        public void DELINVIN(string FILE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("DELETE AP_INVOICEIN WHERE ITEMNAME=@ITEMNAME", connection);
            command.CommandType = CommandType.Text;
  
            command.Parameters.Add(new SqlParameter("@ITEMNAME", FILE));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + FILE);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void UPDATEWH(string INVOICENO, string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_MAIN SET INVOICENO=@INVOICENO WHERE SHIPPINGCODE=@SHIPPINGCODE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        private System.Data.DataTable Getdata2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME,SHIPPING_OBU WAREHOUSE FROM wh_main WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GetAUINV3(string WHNO)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT distinct INV FROM  AP_INVOICEIN WHERE WHNO=@WHNO";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                string da = dataGridView2.SelectedRows[0].Cells["IN"].Value.ToString();


                System.Diagnostics.Process.Start(da);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "重新上傳")
                {

                    string A = dataGridView1.CurrentRow.Cells["A"].Value.ToString();
                    string B = dataGridView1.CurrentRow.Cells["B"].Value.ToString();
                    string JOBNO = dataGridView1.CurrentRow.Cells["工單號碼"].Value.ToString();

                    INS(A, JOBNO, B);

                    MessageBox.Show("上傳成功");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            System.Data.DataTable DATECODE = GDATECODE();
           
            if (DATECODE.Rows.Count > 0)
            {
                dataGridView3.DataSource = DATECODE;

                ExcelReport.GridViewToExcel(dataGridView3);
            
            }

        }
        private System.Data.DataTable GDATECODE()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '20'+SUBSTRING(CARTON,7,2) +CASE SUBSTRING(CARTON,9,1) WHEN 'A' THEN '10' WHEN 'B' THEN '11' WHEN 'C' THEN '12'  ELSE SUBSTRING(CARTON,9,1) END");
            sb.Append(" +SUBSTRING(CARTON,10,2) DATECODE,SUM(CAST(QTY AS INT)) 數量");
            sb.Append("  FROM AP_INVOICEIN  WHERE WHNO=@WHNO");
            sb.Append("  GROUP BY  '20'+SUBSTRING(CARTON,7,2) +CASE SUBSTRING(CARTON,9,1) WHEN 'A' THEN '10' WHEN 'B' THEN '11' WHEN 'C' THEN '12'  ELSE SUBSTRING(CARTON,9,1) END");
            sb.Append(" +SUBSTRING(CARTON,10,2)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHNO", textBox2.Text));


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

     
    }
}

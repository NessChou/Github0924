using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class ACCCUSTCOMP : Form
    {
        string A1 = "";
        string A2 = "";
        string ss1 = "";
        string ss2 = "";
        string ss3 = "";
        string ss4 = "";
        string ss5 = "";
        string ss6 = "";
        string ss7 = "";
        string ss8 = "";
        string ss9 = "";
        string ss10 = "";
        string ss101 = "";
        string ss11 = "";
        string ss12 = "";

        public ACCCUSTCOMP()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            Category( "", "Account_Temp6");

            A1 = textBox3.Text;
            A2 = textBox4.Text;

            Category("", "Account_Temp666");


            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\客戶交易排行.xls";

            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report
            ExcelReport.ExcelReportOutputCOMP(GetCust(), ExcelTemplate, OutPutFile);
        }
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("收入單據", typeof(string));
            dt.Columns.Add("收入單號", typeof(Int32));

            dt.Columns.Add("成本單據", typeof(string));
            dt.Columns.Add("成本單號", typeof(Int32));

            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("客戶群組", typeof(string));

            //20081008
            //業務員
            dt.Columns.Add("業務員編號", typeof(string));
            dt.Columns.Add("姓名", typeof(string));


            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("數量", typeof(Int32));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("金額", typeof(Int32));

            dt.Columns.Add("項目成本", typeof(Int32));   //有成本時寫入此欄位
            dt.Columns.Add("單號總成本", typeof(Int32)); //有成本時寫入此欄位

            dt.Columns.Add("單號總收入", typeof(Int32));
            dt.Columns.Add("基礎單號", typeof(Int32));
            dt.Columns.Add("基礎列", typeof(Int32));

            dt.Columns.Add("日期", typeof(DateTime));
            dt.Columns.Add("科目代號", typeof(string));




            return dt;
        }
        private void Category(string ff, string TABLE)
        {
            AddAUOGD1(TABLE);

            System.Data.DataTable dt = null;

            dt = GetSAPRevenueTempLED(A1, A2);

            
            System.Data.DataTable dtCost = MakeTableCombine2();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string 單據;
            string 科目代號;

            Int32 單號;
            DateTime 日期;

            Int32 基礎單號;
            Int32 基礎列;

            //20080904
            //宣告 DuplicateKey 來檢查
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                單據 = Convert.ToString(dt.Rows[i]["單別"]);
                單號 = Convert.ToInt32(dt.Rows[i]["DocNum"]);
                日期 = Convert.ToDateTime(dt.Rows[i]["日期"]);
                科目代號 = Convert.ToString(dt.Rows[i]["科目代號"]);
                //if (單號 == 23116)
                //{
                //    MessageBox.Show("A");
                //}


                dtDoc = GetSAPDoc(單據, 單號, 科目代號, A1, A2);


                基礎單號 = -1;
                基礎列 = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();
                    string TT = dtDoc.Rows[j]["Quantity"].ToString();
                    dr["收入單據"] = 單據;
                    dr["收入單號"] = 單號;
                    dr["日期"] = 日期;
                    dr["科目代號"] = 科目代號;
                    dr["客戶編號"] = "'" + Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["客戶名稱"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["產品編號"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["產品名稱"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);
                    
                        dr["客戶群組"] = Convert.ToString(dt.Rows[i]["部門"]);
                    
                  
                    dr["數量"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["單價"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["金額"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]);
                    dr["單號總成本"] = 0;
                    dr["項目成本"] = 0;


                    //20081008
                    //業務員
                    dr["業務員編號"] = Convert.ToString(dt.Rows[i]["業務員編號"]);
                    dr["姓名"] = Convert.ToString(dt.Rows[i]["姓名"]);



                    if (單據 == "AR" || 單據 == "貸項" || 單據 == "AR預")
                    {



                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            基礎單號 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["基礎單號"] = 基礎單號;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            基礎列 = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["基礎列"] = 基礎列;
                        }

                    }

                    //總收入寫在最後一筆
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (單據 == "AR" || 單據 == "AR-服務")
                        {
                            dr["單號總收入"] = Convert.ToInt32(dt.Rows[i]["總成本"]);

                        }
                        else if (單據 == "AR預")
                        {
                            dr["單號總收入"] = Convert.ToInt32(dr["金額"]);

                        }

                        else
                        {

                            dr["單號總收入"] = Convert.ToInt32(dt.Rows[i]["總成本"]) * (-1);
                        }
                    }

                    if (單據 == "貸項" || 單據 == "貸項-服務" || 單據 == "銷退")
                    {
                        dr["金額"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  數量改成 負數
                        dr["數量"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }


                    //依據  基礎單號 & 基礎列 取得成本
                    //如果單據本身沒有基礎單號 & 基礎列就認列成本

                    //20080916 AR 打錯單 造成 成本遺漏
                    if (單據 == "AR" || 單據 == "AR預")
                    {
                        //0303
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (基礎單號.ToString() == "3169" && 單號.ToString() == "3429")
                            {
                                dr["項目成本"] = 0;
                                dr["單號總成本"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (基礎單號.ToString() == "3167" && 單號.ToString() == "3404")
                            {
                                dr["項目成本"] = 0;
                                dr["單號總成本"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }

                            dtDocLine = GetSAPDocByLine("交貨", 基礎單號, 基礎列);

                            dr["成本單據"] = "交貨";
                            dr["成本單號"] = 基礎單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                               * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["單號總成本"] = 0;
                                    if (單號 != DuplicateKey)
                                    {

                                        dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }
                                    DuplicateKey = 單號;
                                }
                                //20091204一對多
                                if (基礎單號.ToString() == "5394" && 單號.ToString() == "5673")
                                {

                                    dr["單號總成本"] = 2111964;
                                    dr["單號總收入"] = 0;

                                }
                                //2010331多對一
                                if (單號.ToString() == "6975")
                                {
                                    dr["單號總成本"] = 0;

                                }
                                //2010409訂單轉AR
                                if (單號.ToString() == "7022")
                                {
                                    dr["單號總成本"] = "5476";
                                    dr["項目成本"] = "5476";

                                }

                                System.Data.DataTable GT = TF(基礎單號.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["序號"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (單號.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["單號總成本"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }

                            //成本必須來自至於分錄

                        }
                        //沒有基礎單號
                        else
                        {
                            //成本資料為自已
                            dr["成本單據"] = 單據;
                            dr["成本單號"] = 單號;


                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            if (dtDocLine != null)
                            {

                                if (dtDocLine.Rows.Count == 1)
                                {
                                    dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                                   * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                        {
                                            dr["單號總成本"] = 0;
                                        }
                                        else
                                        {
                                            //反回去找銷貨成本
                                            System.Data.DataTable dtSalesCost = GetSalesCost(單號.ToString());
                                            try
                                            {
                                                dr["單號總成本"] = Convert.ToInt32(dtSalesCost.Rows[0]["總成本"]);
                                            }
                                            catch
                                            {
                                                dr["單號總成本"] = 0;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    //Rows.Count =0 成本為零
                                    dr["項目成本"] = 0;
                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }

                        }

                    }

                    // 3 月案例沒有來源單號

                    //20081007 增加銷退..成本為負

                    if (單據 == "貸項" || 單據 == "貸項-服務" || 單據 == "銷退")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////要判斷來源單種類
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));

                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            //成本資料為自已

                            dr["成本單據"] = 單據;

                            dr["成本單號"] = 單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (單號 != DuplicateKey)
                                        {

                                            dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                        }
                                        DuplicateKey = 單號;

                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;

                                    //20081231
                                    if (單號 != DuplicateKey)
                                    {

                                        dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }
                                    DuplicateKey = 單號;
                                }
                            }


                        }
                        else
                        {


                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            //成本資料為自已

                            dr["成本單據"] = 單據;

                            dr["成本單號"] = 單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (單號 != DuplicateKey)
                                        {

                                            dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                        }
                                        DuplicateKey = 單號;

                                        // dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }
                        }
                    }

                    dtCost.Rows.Add(dr);


                }
            }


            dataGridView8.DataSource = dtCost;
            for (int i = 0; i <= dataGridView8.Rows.Count - 1; i++)
            {
       

                ss1 = dataGridView8.Rows[i].Cells["客戶編號"].Value.ToString();
                ss2 = dataGridView8.Rows[i].Cells["客戶名稱"].Value.ToString();
                ss3 = dataGridView8.Rows[i].Cells["姓名"].Value.ToString();
                ss4 = dataGridView8.Rows[i].Cells["數量"].Value.ToString();
                ss5 = dataGridView8.Rows[i].Cells["單號總收入"].Value.ToString();
                ss6 = dataGridView8.Rows[i].Cells["單號總成本"].Value.ToString();
                ss7 = dataGridView8.Rows[i].Cells["收入單號"].Value.ToString();
                ss8 = dataGridView8.Rows[i].Cells["科目代號"].Value.ToString();
                ss9 = dataGridView8.Rows[i].Cells["客戶群組"].Value.ToString();
                ss10 = dataGridView8.Rows[i].Cells["項目成本"].Value.ToString();
                ss101 = dataGridView8.Rows[i].Cells["金額"].Value.ToString();
                ss11 = dataGridView8.Rows[i].Cells["產品編號"].Value.ToString();
                ss12 = dataGridView8.Rows[i].Cells["產品名稱"].Value.ToString();
                DateTime dd = Convert.ToDateTime(dataGridView8.Rows[i].Cells["日期"].Value);

                if (String.IsNullOrEmpty(ss6))
                {
                    ss6 = "0";
                }
                if (String.IsNullOrEmpty(ss5))
                {
                    ss5 = "0";
                }
            
                    AddAUOGD(TABLE, ss1, ss2, ss3, ss4, ss5, ss6, dd, ss7, ss8, ss9);
                

            }
        }
        private System.Data.DataTable GetSalesCost(string BaseRef)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (T1.[Debit] - T1.[Credit])  總成本");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" WHERE T0.TransType=13 and  T0.BaseRef=@BaseRef and T1.[Account] like '5110%' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@BaseRef", BaseRef));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        System.Data.DataTable TF(string TRGETENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT RANK() OVER (ORDER BY DOCENTRY DESC) AS 序號,DOCENTRY AR,TRGETENTRY 交貨 FROM  INV1 WHERE  TRGETENTRY IN (");
            sb.Append(" select docentry  from dln1 where BASEtype='13'");
            sb.Append(" GROUP BY DOCENTRY HAVING COUNT (DISTINCT BASEENTRY) >1) AND    DOCENTRY NOT IN (SELECT DOCENTRY FROM OINV WHERE DOCTOTAL=0) AND TRGETENTRY=@TRGETENTRY ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRGETENTRY", TRGETENTRY));
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
        public void AddAUOGD(string TABLE, string CARDCODE, string CARDNAME, string SALES, string GQty, string GTotal, string GValue, DateTime 日期, string DOCENTRY, string ACCOUNT, string CARDGROUP)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into " + TABLE + "(CARDCODE,CARDNAME,SALES,GQty,GTotal,GValue,DDATE,DOCENTRY,ACCOUNT,CARDGROUP) values(@CARDCODE,@CARDNAME,@SALES,@GQty,@GTotal,@GValue,@DDATE,@DOCENTRY,@ACCOUNT,@CARDGROUP)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@GQty", GQty));
            command.Parameters.Add(new SqlParameter("@GTotal", GTotal));
            command.Parameters.Add(new SqlParameter("@GValue", GValue));
            command.Parameters.Add(new SqlParameter("@DDATE", 日期));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", CARDGROUP));

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

        public void AddAUOGD1(string TABLE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table " + TABLE + " ", connection);
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

        //取單據本身所連結的分錄金額
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {


            }
            else if (DocKind == "銷退")
            {


            }
            else if (DocKind == "貸項")
            {

                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");


            }
            else if (DocKind == "AR" || DocKind == "AR-服務" | DocKind == "AR預")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");



            }
            else if (DocKind == "JE")
            {

            }
            else if (DocKind == "貸項-服務")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
                connection.Close();
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum, Int32 LineNum)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {

                sb.Append(" SELECT LINENUM,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append(" T0.[LineTotal], T0.[StockPrice],T2.總成本 ");
                sb.Append(" FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append(" INNER JOIN (SELECT SUM([Debit]-[Credit]) 總成本,TransId FROM JDT1 WHERE [Account]='51100101' GROUP BY TransId) T2 ");
                sb.Append(" ON(T1.TransId=T2.TransId)");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");
                sb.Append("and   T0.LineNum =@LineNum   ");



            }
            else if (DocKind == "銷退")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "貸項")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");

                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT * FROM (SELECT T0.ACCTCODE,SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='I'   ");
                sb.Append("UNION ALL   ");


            }
            else if (DocKind == "AR預")
            {



                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0   and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2  ");
            }
            else if (DocKind == "JE")
            {

                sb.Append("SELECT  T0.Account as CardCode, T0.LineMemo as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price], ");
                sb.Append("  T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice] FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID  ");
                sb.Append("WHERE T1.TransID =@DocEntry   ");

            }
            else if (DocKind == "貸項-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));

            command.Parameters.Add(new SqlParameter("@A1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenueTempLED(string DocDate1, string DocDate2)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%'  ) and isnull(t2.u_acme_arap,'') <> 'xx' ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");

            sb.Append(" union all");
            sb.Append(" SELECT '貸項' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            //貸項服務
            sb.Append(" union all");
            sb.Append(" SELECT '貸項-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%' )  ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as 單別,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T7.AcctCode)  科目代號,");
            sb.Append("              T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append("              0 總金額,");
            sb.Append("             0  總成本,");
            sb.Append("            0  - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("            GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName,T4.ocrname ");

            sb.Append(" union all");
            sb.Append("              SELECT 'AR預' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account 科目代號,");
            sb.Append("              T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T6.[DOCDATE]) 日期");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and ((T1.[Account] = '22610103') OR (T2.DOCENTRY in (10198,24001))) ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            sb.Append(" union all");
            sb.Append("                         SELECT '貸項' as 單別,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("                        MAX(T6.AcctCode)  科目代號,");
            sb.Append("                         T2.SlpCode 業務員編號, T3.SlpName 姓名,T4.ocrname 部門,");
            sb.Append("                         0 總金額,");
            sb.Append("                              0  總成本,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) 總毛利,MAX(T0.[RefDate]) 日期");
            sb.Append("                         FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("             INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("             Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append("                         WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and Convert(varchar(8),T0.RefDate,112) >=@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append("                         AND T2.[DocTotal] = 0 ");
            sb.Append("                       GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName,T4.ocrname ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            string df = textBox3.Text.Substring(0, 6);
            string df1 = textBox3.Text.Substring(4, 2);
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPDoc(string DocKind, Int32 DocNum, string AcctCode, string A1, string A2)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");


            }
            else if (DocKind == "銷退")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "貸項")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='I'   ");



            }
            else if (DocKind == "AR預")
            {

                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 ");
                sb.Append("INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append("INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)");
                sb.Append("LEFT JOIN ODLN T4 ON (T3.DOCENTRY=T4.DOCENTRY)");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   T1.UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0 and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2     ");


            }
            else if (DocKind == "JE")
            {

                sb.Append("SELECT  T0.Account as CardCode, T0.LineMemo as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price], ");
                sb.Append("  T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice],'' GROUPCODE FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID  ");
                sb.Append("WHERE T1.TransID =@DocEntry   ");

            }
            else if (DocKind == "貸項-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //加入基礎單號 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //加入基礎單號 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            //20081009 增加 科目代號
            sb.Append("AND  T0.AcctCode =@AcctCode   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            //20081009 增加 科目代號
            command.Parameters.Add(new SqlParameter("@AcctCode", AcctCode));
            command.Parameters.Add(new SqlParameter("@A1", A1));
            command.Parameters.Add(new SqlParameter("@A2", A2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }

        private void ACCCUSTCOMP_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            textBox3.Text = GetMenu.DFirst();
            textBox4.Text = GetMenu.DLast();
        }

        public System.Data.DataTable GetCust()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT T0.CARDCODE 客戶編號,T0.CARDGROUP BU,T0.CARDNAME 客戶名稱,'%' A1,日期1=@DATE1,日期2=@DATE2,日期3=@DATE3,日期4=@DATE4,");
            //sb.Append(" SUM(T0.GQTY) 數量,  SUM(T0.GTOTAL) 總實銷金額, SUM(T0.GVALUE) 總實銷成本,  SUM(T0.GTOTAL)-SUM(T0.GVALUE) 銷售利潤,CASE SUM(T0.GTOTAL)  WHEN 0 THEN 0 ELSE CAST((SUM(T0.GTOTAL)-SUM(T0.GVALUE))/SUM(T0.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比,");
            //sb.Append(" MAX(T1.GQTY) 數量3,MAX(T1.GTOTAL) 總實銷金額3,MAX(T1.GVALUE) 總實銷成本3,MAX(T1.GTOTAL)-MAX(T1.GVALUE) 銷售利潤3,CASE MAX(T1.GTOTAL) WHEN 0 THEN 0 ELSE CAST((MAX(T1.GTOTAL)-MAX(T1.GVALUE))/MAX(T1.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比3,");
            //sb.Append(" SUM(T0.GQTY)-MAX(T1.GQTY) 數量2,SUM(T0.GTOTAL)-MAX(T1.GTOTAL) 總實銷金額2,SUM(T0.GVALUE)-MAX(T1.GVALUE)  總實銷成本2,");
            //sb.Append(" (SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)) 銷售利潤2,CASE (SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) WHEN 0 THEN 0 ELSE CAST(((SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)))/(SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) *100 AS DECIMAL(18,2)) END 利潤比2");
            //sb.Append("  FROM Account_Temp6 T0");
            //sb.Append(" LEFT JOIN (SELECT CARDCODE,SUM(GVALUE) GVALUE,SUM(GQTY) GQTY,SUM(GTOTAL)  GTOTAL FROM Account_Temp666  GROUP BY CARDCODE) T1  ON(T0.CARDCODE=T1.CARDCODE)");
            //sb.Append(" GROUP BY T0.CARDCODE,T0.CARDNAME,T0.CARDGROUP");
            //sb.Append(" ORDER BY  SUM(T0.GTOTAL)-SUM(T0.GVALUE) DESC");

            sb.Append("     SELECT '1' A, T0.CARDGROUP BU,T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,'%' A1,日期1=@DATE1,日期2=@DATE2,日期3=@DATE3,日期4=@DATE4, ");
            sb.Append("               SUM(T0.GQTY) 數量,  SUM(T0.GTOTAL) 總實銷金額, SUM(T0.GVALUE) 總實銷成本,  SUM(T0.GTOTAL)-SUM(T0.GVALUE) 銷售利潤,CASE SUM(T0.GTOTAL)  WHEN 0 THEN 0 ELSE CAST((SUM(T0.GTOTAL)-SUM(T0.GVALUE))/SUM(T0.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比, ");
            sb.Append("               MAX(T1.GQTY) 數量3,MAX(T1.GTOTAL) 總實銷金額3,MAX(T1.GVALUE) 總實銷成本3,MAX(T1.GTOTAL)-MAX(T1.GVALUE) 銷售利潤3,CASE MAX(T1.GTOTAL) WHEN 0 THEN 0 ELSE CAST((MAX(T1.GTOTAL)-MAX(T1.GVALUE))/MAX(T1.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比3, ");
            sb.Append("               SUM(T0.GQTY)-MAX(T1.GQTY) 數量2,SUM(T0.GTOTAL)-MAX(T1.GTOTAL) 總實銷金額2,SUM(T0.GVALUE)-MAX(T1.GVALUE)  總實銷成本2, ");
            sb.Append("               (SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)) 銷售利潤2,CASE (SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) WHEN 0 THEN 0 ELSE CAST(((SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)))/(SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) *100 AS DECIMAL(18,2)) END 利潤比2 ");
            sb.Append("                FROM Account_Temp6 T0 ");
            sb.Append("               LEFT JOIN (SELECT CARDCODE,SUM(GVALUE) GVALUE,SUM(GQTY) GQTY,SUM(GTOTAL)  GTOTAL FROM Account_Temp666  GROUP BY CARDCODE) T1  ON(T0.CARDCODE=T1.CARDCODE) ");
            sb.Append("               GROUP BY T0.CARDCODE,T0.CARDNAME,T0.CARDGROUP ");
            sb.Append("        UNION ALL");
            sb.Append("     SELECT '2' A,'','','','','','','','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0");
            sb.Append("    UNION ALL");
            sb.Append("     SELECT '3' A, T0.CARDGROUP BU,'' 客戶編號,'' 客戶名稱,'%' A1,日期1=@DATE1,日期2=@DATE2,日期3=@DATE3,日期4=@DATE4, ");
            sb.Append("               SUM(T0.GQTY) 數量,  SUM(T0.GTOTAL) 總實銷金額, SUM(T0.GVALUE) 總實銷成本,  SUM(T0.GTOTAL)-SUM(T0.GVALUE) 銷售利潤,CASE SUM(T0.GTOTAL)  WHEN 0 THEN 0 ELSE CAST((SUM(T0.GTOTAL)-SUM(T0.GVALUE))/SUM(T0.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比, ");
            sb.Append("               MAX(T1.GQTY) 數量3,MAX(T1.GTOTAL) 總實銷金額3,MAX(T1.GVALUE) 總實銷成本3,MAX(T1.GTOTAL)-MAX(T1.GVALUE) 銷售利潤3,CASE MAX(T1.GTOTAL) WHEN 0 THEN 0 ELSE CAST((MAX(T1.GTOTAL)-MAX(T1.GVALUE))/MAX(T1.GTOTAL)*100 AS DECIMAL(18,2)) END 利潤比3, ");
            sb.Append("               SUM(T0.GQTY)-MAX(T1.GQTY) 數量2,SUM(T0.GTOTAL)-MAX(T1.GTOTAL) 總實銷金額2,SUM(T0.GVALUE)-MAX(T1.GVALUE)  總實銷成本2, ");
            sb.Append("               (SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)) 銷售利潤2,CASE (SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) WHEN 0 THEN 0 ELSE CAST(((SUM(T0.GTOTAL)-SUM(T0.GVALUE))-(MAX(T1.GTOTAL)-MAX(T1.GVALUE)))/(SUM(T0.GTOTAL)-MAX(T1.GTOTAL)) *100 AS DECIMAL(18,2)) END 利潤比2 ");
            sb.Append("                FROM Account_Temp6 T0 ");
            sb.Append("               LEFT JOIN (SELECT CARDCODE,SUM(GVALUE) GVALUE,SUM(GQTY) GQTY,SUM(GTOTAL)  GTOTAL FROM Account_Temp666  GROUP BY CARDCODE) T1  ON(T0.CARDCODE=T1.CARDCODE) ");
            sb.Append("               GROUP BY T0.CARDGROUP ");
            sb.Append("       ORDER BY  A,SUM(T0.GTOTAL)-SUM(T0.GVALUE) DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@DATE3", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DATE4", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["data"];
        }
    }
}

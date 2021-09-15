using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.IO;
using Microsoft.Office.Interop.Excel;


//20200110 AUO 上傳路徑
//\\Acmesrv01\public\TFT廠商提供INVOICE-箱序對應片序\AU每月出貨片序號

//SH20191129003X
//select carton,pic from AP_INVOICEIN
//where inv in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')

//ClosedXML.Excel

//20191225 常見服務
//1.請輸入工單號碼
//2.勾選指定工單號碼重送
//3.按下 自動化


//相關資料表
//[ACME_CUS_TIME] -> 設定下一次讀取區間 StartName = "SN"
//ACME_CUS_SET -> 指定客戶
//ACME_ARES_MAIL where SysCode='SerialNo' -> 郵件設定
//ACME_MAIL_LOG ->系統 LOG  DocType = "SN"


//AP_INVOICEIN -> Sunny 倉庫上傳
//WH_AUO ->採購 AUO 提供


//觸發點  AP_INVOICEIN 資料異動

//可參考框架
//https://github.com/quozd/awesome-dotnet
//https://github.com/thangchung/awesome-dotnet-core
///([a-zA-Z0-9\s_\\.\-\(\):])+(.pdf)$/
///.*\.PDF
///"ShipDoc_CSD1573140_Pk.pdf" "ShipDoc_CSD1568824_Pk.pdf" "ShipDoc_CSD1573140_Inv.pdf" 
namespace ACME
{
    public partial class fmAcmeCusSerial : Form
    {

        private StreamWriter sw;

        private string ShipConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";

        private string TestFlag = "N";
        private string TestRec = "N";
        private string ErrorFlag = "N";

        public fmAcmeCusSerial()
        {
            InitializeComponent();
            //不支援路徑寫法
            //自動 Append
            sw = new StreamWriter("log.txt", true, Encoding.UTF8);//creating html file
        }

        public System.Data.DataTable GetData(string Sql)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(Sql);
            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }

        private void SizeAllColumns(Object sender, EventArgs e)
        {
            dgCustomer.AutoResizeColumns(
                DataGridViewAutoSizeColumnsMode.AllCells);
        }


        private void GridViewAutoSize(DataGridView dgv)
        {

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                //dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader;
               // dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                //DisplayedCells
               // dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }

            //dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                int colw = dgv.Columns[i].Width;
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgv.Columns[i].Width = colw;
            }
        }


        private string GetCartonNo(string SerialNo)
        {
            string Sql = "select CARTON_NO from wh_auo where shipping_no='{0}'";
            Sql = string.Format(Sql, SerialNo);
            System.Data.DataTable dt = GetData(Sql);
            try
            {
                return Convert.ToString(dt.Rows[0][0]);
            }
            catch
            {
                return "";
            }
        }

        public System.Data.DataTable UpdateBlankRow(System.Data.DataTable dt)
        {
            DataRow dr;
            string F01;
            string KeyValue = ""; ;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];
                try
                {
                    F01 = Convert.ToString(dr[0]);
                }
                catch
                {
                    F01 = "";
                }
                if (F01 == "")
                {
                    dr.BeginEdit();
                    dr[0] = KeyValue;
                    dr.EndEdit();
                }
                else
                {
                    KeyValue = F01;
                }

            }
            return dt;
        }


        public System.Data.DataTable UpdateBlankCartonRow(System.Data.DataTable dt)
        {
            DataRow dr;
            string F01;
            string KeyValue = ""; ;
            string SerialNo="";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];
                SerialNo = Convert.ToString(dr["序號"]);
                try
                {
                    F01 = Convert.ToString(dr["箱號"]);
                    
                }
                catch
                {
                    F01 = "";
                }

                if (F01 == "")
                {
                    KeyValue = GetCartonNo(SerialNo);
                    dr.BeginEdit();
                    dr[0] = KeyValue;
                    dr.EndEdit();
                }
                

            }
            return dt;
        }

        private Int32 GetDataCount(string sql)
        {

            SqlConnection connection = new SqlConnection(ShipConnectiongString);

            StringBuilder sb = new StringBuilder();

            // sb.Append("SELECT Count(*) As RecCount FROM gb_pick2 Where BillNo=@BillNo");

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            //
            // command.Parameters.Add(new SqlParameter("@BillNo", BillNo));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            Int32 RecCount = 0;
            try
            {
                connection.Open();
                //command.Prepare();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    RecCount = (Int32)reader[0];
                }

                // return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }

            return RecCount;

        }


        private System.Data.DataTable GetSerial(string WhNo ,ref string Msg)
        {

            string Sql = @"select distinct carton 箱號,SHIPPING_NO 序號 from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where whno='{0}'
and pic ='') t1 on t1.CARTON=t0.CARTON_NO
union all
select distinct carton 箱號,pic 序號 from AP_INVOICEIN
where whno='{0}'
and pic <>''
";
            

            Sql = string.Format(Sql, WhNo);

            System.Data.DataTable dt = GetData(Sql);

           // dt = UpdateBlankRow(dt);

            //20191022 補箱號
            dt = UpdateBlankCartonRow(dt);


            // AddTextlvw("核對數量",4);
            //核對數量
            //string Sql2 = "select isnull(Sum(convert(int,Quantity)),0) Qty from WH_Item where ShippingCode='{0}'";
            //Sql2 = string.Format(Sql2, WhNo);
            //Int32 recCount = GetDataCount(Sql2);
            //// //MessageBox.Show(recCount.ToString());
            //if (recCount != dt.Rows.Count)
            //{
            //    Msg = "資料不符合->工單數量={0} 箱序數量={1}";
            //    Msg = string.Format(Msg, recCount.ToString(), dt.Rows.Count.ToString());
            //}

            return dt;
        }



        private System.Data.DataTable GetSerialInvoice(string WhNo, ref string Msg)
        {
            //20191021 distinct
            string Sql = @"select distinct carton 箱號,SHIPPING_NO 序號,INVOICE_NO 發票號碼,
isnull(case when CHARINDEX('.',MODEL_NO) >0 then
 SUBSTRING(MODEL_NO,1,CHARINDEX('.',MODEL_NO)+1) 
 else
 MODEL_NO
 end,'') 型號 from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where whno='{0}'
and pic ='') t1 on t1.CARTON=t0.CARTON_NO
union all
select  distinct carton 箱號,pic 序號,INV as 發票號碼,
isnull(case when CHARINDEX('.',t3.MODEL_NO) >0 then
 SUBSTRING(t3.MODEL_NO,1,CHARINDEX('.',t3.MODEL_NO)+1) 
 else
 t3.MODEL_NO
 end ,'')
from AP_INVOICEIN t2
left join WH_AUO t3 on t3.INVOICE_NO=t2.INV  and t3.SHIPPING_NO=t2.PIC
where whno='{0}'
and pic <>''
";


            Sql = string.Format(Sql, WhNo);

            System.Data.DataTable dt = GetData(Sql);

            //dt = UpdateBlankRow(dt);

            //20191022 補箱號
            dt = UpdateBlankCartonRow(dt);


            // AddTextlvw("核對數量",4);
            //核對數量
            string Sql2 = "select isnull(Sum(convert(int,Quantity)),0) Qty from WH_Item where ShippingCode='{0}'";
            Sql2 = string.Format(Sql2, WhNo);
            Int32 recCount = GetDataCount(Sql2);
            // //MessageBox.Show(recCount.ToString());
            if (recCount != dt.Rows.Count)
            {
                Msg = "資料不符合->工單數量={0} 箱序數量={1}";
                Msg = string.Format(Msg, recCount.ToString(), dt.Rows.Count.ToString());

                //20191022增加 檢核異常
                //來源比對
//select Carton,Pic,Qty from AP_INVOICEIN
//where whno='WH20191021008X'
//and Pic=''
//order by CARTON
            }

            return dt;
        }


        private void button14_Click(object sender, EventArgs e)
        {
            button4.Enabled = true;
            button15.Enabled = false;
            
            //WH20191223014X 台亞 Shipping三角 連結 倉庫工單
            
            ErrorFlag = "N";

            if (TestFlag == "N")
            {
                lvwUpdate.Items.Clear();
            }

            //AddTextlvw("查詢開始");

            //20191021 //加入 distinct

            string Sql = @"select distinct T0.CARTON_NO  箱號,SHIPPING_NO 序號,STOCK_IN_WEEK WC from WH_AUO t0
INNER JOIN AP_INVOICEIN T1 ON (T0.PALLET_NO=T1.PLT)
where T1.WHNO='{0}'  AND T1.PLT <> ''
union all
select distinct carton 箱號,SHIPPING_NO 序號,STOCK_IN_WEEK WC from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where whno='{0}'
and pic ='' AND carton <> '') t1 on t1.CARTON=t0.CARTON_NO
union all
select distinct carton 箱號,pic 序號,'' WC from AP_INVOICEIN
where whno='{0}'
and pic <>''
";
            string WhNo = txtWhNo.Text;

            Sql = string.Format(Sql,WhNo);

            System.Data.DataTable dt = GetData(Sql);

            //dt = UpdateBlankRow(dt);

            //20191022 補箱號
            dt = UpdateBlankCartonRow(dt);


            dgData.DataSource = dt;

            GridViewAutoSize(dgData);

            

           // AddTextlvw("核對數量",4);
            //核對數量
            string Sql2 = "select isnull(Sum(convert(int,Quantity)),0) Qty from WH_Item where ShippingCode='{0}'";
            Sql2 = string.Format(Sql2, WhNo);
            Int32 recCount = GetDataCount(Sql2);
           // //MessageBox.Show(recCount.ToString());
            if (recCount != dt.Rows.Count)
            {
                if (TestFlag == "N")
                {
                    //MessageBox.Show("資料筆數不符合");
                }
                AddTextlvw(WhNo+"-核對數量-資料筆數不符合", 1);

                ErrorFlag = "Y";
                //異常分析
                button7_Click(sender, e);
            }

            //AddTextlvw(TestRec+"完成");
            //System.Diagnostics.Process.Start(OutPutFile);
        }
        private System.Data.DataTable GETA1(string PARAM_NO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT PARAM_DESC   FROM PARAMS WHERE PARAM_KIND ='WHLOCATION' AND PARAM_NO =@PARAM_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));


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
        public void SqlPost(string sql)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            try
            {
                connection.Open();
                command.ExecuteNonQuery();
                //  txtMsg.Text += sql + "\r\n";
            }
            catch (System.Exception ex)
            {
                //textBox1.Text += ex.Message + "\r\n";
            }
            finally
            {
                connection.Close();
            }
        }


        private void fmAcmeCusSerial_Load(object sender, EventArgs e)
        {
            
            txtStartDate.Text = DateTime.Now.ToString("yyyyMMdd");
            txtEndDate.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox2.Text = DateTime.Now.ToString("yyyyMM") + "01";
            textBox3.Text = DateTime.Now.ToString("yyyyMMdd");
            button4.Enabled = false;
            button15.Enabled = false;

            //開始點---------------------------------------
            //if (DateTime.Now.Hour > 20)
            //{
            //    Close();
            //    return;
            //}
            //else
            //{
            //    button2_Click(sender, e);
            //    Close();
            //}
            //結束---------------------------------------


            //
            //string Sql = "select CardName 客戶 from [ACME_CUS_SET]";
            //System.Data.DataTable dt = GetData(Sql);
            //dgCustomer.DataSource = dt;
            //dgCustomer.Columns[0].Width = 200;
            //GridViewAutoSize(dgCustomer);

            tabControl1.TabPages.Remove(tabPage3);
             tabControl1.TabPages.Remove(tabPage1);
             tabControl1.TabPages.Remove(tabPage7);
             tabControl1.TabPages.Remove(tabPage6);
             tabControl1.TabPages.Remove(tabPage4);
            

            //Packing List
            //http://www.torus.com.tw/product/et61/manual_/et611.htm

            //AP2AP電子對帳
            //http://web.goodservice.com.tw/b61.php



            

           
        }

        

        //2 Accept
        private void AddTextlvw(string sText, int iconIdx = 2)
        {
            lvwUpdate.Items.Add(sText, iconIdx);
            lvwUpdate.Items[lvwUpdate.Items.Count - 1].Selected = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
            
            TestFlag = "Y";
            string Sql = @"SELECT top 100 ShippingCode FROM WH_main WHERE ([WH_main].[closeDay] LIKE '%20190%') AND ([WH_main].[createName] LIKE '%sunny')";


            System.Data.DataTable dt = GetData(Sql);

            string WhNo = "";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                WhNo = Convert.ToString(dt.Rows[i][0]);
                txtWhNo.Text = WhNo;
                TestRec = (i + 1).ToString() + "/" + dt.Rows.Count.ToString();
                button14_Click(sender, e);
                lvwUpdate.Refresh();
            }

            TestFlag = "N";

            //Excel Automation
            //https://stackoverflow.com/questions/2431840/excel-automation-using-c-sharp

            //Zapier
            //var urlFields = []

//    urlFields[ urlFields.length ] = "barcodeNumber";
//    urlFields[ urlFields.length ] = "articleNumber";
//    urlFields[ urlFields.length ] = "licence";
//    urlFields[ urlFields.length ] = "description";
//    urlFields[ urlFields.length ] = "purchaseOrderNumber";
//    urlFields[ urlFields.length ] = "quantity";
////	urlFields[ urlFields.length ] = "grossWeight";
////	urlFields[ urlFields.length ] = "netWeight";
////	urlFields[ urlFields.length ] = "cartonMeasurement";
////	urlFields[ urlFields.length ] = "cargoTrackingNoteNumber";
//    urlFields[ urlFields.length ] = "comment";
            //view-source:http://www.dreispur.de/smg/index.html //PDFJS

    //            Purchaser’s name and/or logo;
    //Product reference and/or order number;
    //Country of destination (for example, a forwarder might have to distinguish some cartons for Canada and others for the US);
    //Other relevant information about the products: season, size, color, or breakdown of the different types of goods inside a particular carton;
    //Net weight & gross weight of carton;
    //Dimensions of the carton;
    //Number of cartons (example: 1/230; 2/230; 3/230…); Individual numbers are written in marker

            //https://help.shipstation.com/hc/en-us/articles/360026190932-V3-Create-Shipments-Labels

        }


        private void SendMail(string WhNo, string CardName)
        {
            DataRow dr;

            string strSubject;
            string UserCode;
            string UserMail;
            string MailContent = "";
            string MailDate = DateTime.Now.ToString("yyyyMMdd");

            string DocType = "SN";

            System.Data.DataTable dt = GetACME_MAILLIST();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                //產生檔案

                //發送郵件
                strSubject = string.Format("[工單] {0} 箱片序號 - 沒有資料", WhNo);

                if (cb.Checked)
                {
                    strSubject = string.Format("[工單重送] {0} 箱片序號 - 沒有資料", WhNo);
                }

                dr = dt.Rows[i];
                UserCode = Convert.ToString(dr["UserCode"]);

                SetMsg("[郵寄] " + UserCode);

                UserMail = Convert.ToString(dt.Rows[i]["UserMail"]);

                if (string.IsNullOrEmpty(UserMail))
                {
                    UserMail = "terrylee@acmepoint.com";
                }

                MailContent = string.Format("客戶:{0} <br> 工單:{1} <br>", CardName, WhNo);

                

                MailTest(strSubject, UserCode, UserMail, MailContent);


                //Stage
                AddACME_MAIL_LOG(DocType, MailDate, UserCode,WhNo);

            }
        }

        private void SendMail(string FileName, string FileName2,string  WhNo, string CardName, string CheckMsg)
        {
            DataRow dr;
           
            string strSubject;
            string UserCode;
            string UserMail;
            string MailContent ="";
            string MailDate = DateTime.Now.ToString("yyyyMMdd");

            string DocType = "SN";

            System.Data.DataTable dt = GetACME_MAILLIST();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                //產生檔案

                //發送郵件
                strSubject = string.Format("[工單] {0} 箱片序號檔案上傳結果訊息通知", WhNo);
                if (CheckMsg != "")
                {
                    strSubject = string.Format("[工單] {0} 箱片序號檔案上傳結果[異常]通知", WhNo);
                }


                if (cb.Checked)
                {
                    strSubject = string.Format("[工單重送] {0} 箱片序號檔案上傳結果訊息通知", WhNo);
                    if (CheckMsg != "")
                    {
                        strSubject = string.Format("[工單重送] {0} 箱片序號檔案上傳結果[異常]通知", WhNo);
                    }
                }

                dr = dt.Rows[i];
                UserCode = Convert.ToString(dr["UserCode"]);

                SetMsg("[郵寄] " + UserCode);

                UserMail = Convert.ToString(dt.Rows[i]["UserMail"]);

                if (string.IsNullOrEmpty(UserMail))
                {
                    UserMail = "terrylee@acmepoint.com";
                }

                MailContent = string.Format("客戶:{0} <br> 工單:{1} <br>", CardName,WhNo);

                if (CheckMsg == "")
                {
                    
                }
                else
                {
                    MailContent += string.Format("異常訊息:{0} <br>", CheckMsg);
                    //
                    button7_Click(null,null);
                    ErrorFlag = "Y";

                    MailContent += "<br>" +
                        "[異常資料]<br>" + htmlMessageBody(dgError);

                }

                MailTest(strSubject, UserCode, UserMail, MailContent, FileName, FileName2);


                //Stage
                AddACME_MAIL_LOG(DocType, MailDate, UserCode,WhNo);

            }
        }

        private System.Data.DataTable GetNext()
        {
            string Sql = "select MaxID,StartDate from acme_cus_time where StartName='SN'";

            System.Data.DataTable dt = GetData(Sql);

            return dt;
        }


       

        private System.Data.DataTable UpdateNext(string MaxID,string UpdateDate)
        {
            string Sql = "update acme_cus_time set MaxID='{0}',StartDate='{1}',UpdateDate='{2}',UpdateTime='{3}'   where StartName='SN'";
            Sql = string.Format(Sql,MaxID,UpdateDate,
                DateTime.Now.ToString("yyyyMMdd"),
                DateTime.Now.ToString("HHmmss"));

            System.Data.DataTable dt = GetData(Sql);

            return dt;
        }

        private System.Data.DataTable UpdateNextNoData(string MaxID)
        {
            string Sql = "update acme_cus_time set MaxID='{0}',UpdateDate='{1}',UpdateTime='{2}'   where StartName='SN'";
            Sql = string.Format(Sql, MaxID, 
                DateTime.Now.ToString("yyyyMMdd"),
                DateTime.Now.ToString("HHmmss"));

            System.Data.DataTable dt = GetData(Sql);

            return dt;
        }

        //DataTable oTable = GetDataTable("cars");
        //DataTable nTable = DataTableFilterSort(oTable, "SPEED='10'", "DIST asc");
        private System.Data.DataTable DataTableFilter(System.Data.DataTable oTable, string filterExpression)
        {
            DataView dv = new DataView();
            dv.Table = oTable;
            dv.RowFilter = filterExpression;
           // dv.Sort = sortExpression;
            System.Data.DataTable nTable = dv.ToTable();
            return nTable;
        }

        private System.Data.DataTable DataTableFilterSort2(System.Data.DataTable oTable, string filterExpression, string sortExpression)
        {
            System.Data.DataTable nTable = oTable.Select(filterExpression, sortExpression).CopyToDataTable();
            return nTable;
        }

        //只取符合設定的客戶
//select count(*) from AP_INVOICEIN t0
//inner join ACME_CUS_SET t1 on t1.CardName=t0.CARD

        private void button2_Click(object sender, EventArgs e)
        {
            lblMsg.Text = "Starting";
            SetMsg("Starting");
            //CardName //inner join ACME_CUS_SET t1 on t1.CardName=t0.CARD
            //INSERTDATE >'{0}' 
            string Sql = @"select max(t0.id) MaxID, t0.whno,t0.Card from AP_INVOICEIN t0
inner join ACME_CUS_SET t1 on substring(t1.CardName,1,4)=substring(t0.CARD,1,4)
where t0.INSERTDATE >='{0}'
and t0.ID > {1}
group by t0.whno,t0.Card
order by MaxID 
";

            //[ACME_CUS_TIME]

            string SqlMaxID = "select max(id) MaxID from AP_INVOICEIN";
            System.Data.DataTable dtMaxID = GetData(SqlMaxID);

            string DoDate = txtDate.Text;
            string MaxID = "123";
            System.Data.DataTable dtGetNext = GetNext();
            DoDate = Convert.ToString(dtGetNext.Rows[0]["StartDate"]);
            MaxID = Convert.ToString(dtGetNext.Rows[0]["MaxID"]);

            Sql = string.Format(Sql, DoDate,MaxID);

            if (cb.Checked)
            {
                Sql = @"select max(t0.id) MaxID, t0.whno,t0.Card from AP_INVOICEIN t0
inner join ACME_CUS_SET t1 on substring(t1.CardName,1,4)=substring(t0.CARD,1,4)
where t0.WHNO='{0}'
group by t0.whno,t0.Card
order by MaxID ";

                Sql = string.Format(Sql, txtWhNo.Text);
            }

            System.Data.DataTable dt = GetData(Sql);
            dgData.DataSource = dt;
           // GridViewAutoSize(dgData);

            //記錄 最後處理的日期及ID

            DataRow dr;
            string WhNo = "";
            string FileName="";

            string CardName;
            string CheckMsg="";
            string OutFileName = "";
            string OutFileName2 = ""; 


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                ErrorFlag = "N";

                CheckMsg = "";
                //產生檔案
                WhNo = Convert.ToString(dt.Rows[i]["whno"]);

                if (WhNo =="WH20191224011X") continue;

                txtWhNo.Text = WhNo;

                SetMsg("[工單] "+WhNo);

                CardName = Convert.ToString(dt.Rows[i]["Card"]);
                MaxID = Convert.ToString(dt.Rows[i]["MaxID"]);

                System.Data.DataTable dtData = GetSerial(WhNo, ref CheckMsg);

                if (dtData == null || dtData.Rows.Count == 0)
                {
                    // //MessageBox.Show("沒有資料");
                    // return;

                    //20191016 沒有資料先不發送
                    lblMsg.Text = "";
                   // string strSubject = string.Format("[工單] {0} 箱片序號 - 沒有資料", WhNo);

                    //MailTest(strSubject, "terrylee",
                    //    "terrylee@acmepoint.com", WhNo);

                    SendMail(WhNo, CardName);


                    continue;
                }
                else
                {
                    //OutFileName = GetExePath() + "\\" + WhNo + ".xlsx";
                    OutFileName = GetExePath() + "\\" + WhNo + ".xls";
                    WriteDataTableToExcel(dtData, "箱序號", OutFileName, "");
                }

                dtData = GetSerialInvoice(WhNo, ref CheckMsg);

                if (dtData == null || dtData.Rows.Count == 0)
                {
                    // //MessageBox.Show("沒有資料");
                    // return;
                    //20191016 沒有資料先不發送
                    continue;
                }
                else
                {
                    //OutFileName2 = GetExePath() + "\\" + WhNo + "_Inv.xlsx";
                    OutFileName2 = GetExePath() + "\\" + WhNo + "_Inv.xls";

                   // System.Data.DataTable dtGroup = dtDistinct(dtData,"型號");
                    WriteDataTableToExcel2(dtData, "箱序號", OutFileName2, "");
                }

                


                //發送郵件
               SendMail(OutFileName, OutFileName2, WhNo, CardName, CheckMsg);
               
            }


            if (cb.Checked)
            {
                return;
            }

            if (dt.Rows.Count > 0)
            {
                string sqlDate = "select INSERTDATE from AP_INVOICEIN where id='{0}'";
                sqlDate = string.Format(sqlDate, MaxID); ;
                System.Data.DataTable dtDate = GetData(sqlDate);
                string UpdateDate = Convert.ToString(dtDate.Rows[0][0]);

                UpdateNext(MaxID, UpdateDate);
            }
            else
            {

                MaxID = Convert.ToString(dtMaxID.Rows[0][0]);
                UpdateNextNoData(MaxID);
            }

            lblMsg.Text ="處理筆數:"+ dt.Rows.Count.ToString();
        }

        public void AddACME_MAIL_LOG(string DocType, string MailDate, string UserCode,string Msg)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            SqlCommand command = new SqlCommand("Insert into ACME_MAIL_LOG(DocType,MailDate,UserCode,Msg) values(@DocType,@MailDate,@UserCode,@Msg)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocType", DocType));
            command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            command.Parameters.Add(new SqlParameter("@Msg", Msg));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }


        private void dgData_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void SetMsg(string Msg)
        {

            WriteToLog(sw, string.Format("[{0}] {1}\r\n",
                DateTime.Now.ToString("yyyyMMdd HHmmss"),
                 Msg  ));
        }

        private void WriteToLog(StreamWriter sw, string Msg)
        {
            // StreamWriter sw = new StreamWriter("file.html", true, Encoding.UTF8);//creating html file
            sw.Write(Msg);
            // sw.Close();
        }

        //private  void ExportToExcelFile(DataGridView dGV, string filename, string tabName)
        //{
        //    //Creating DataTable
        //    System.Data.DataTable dt = new System.Data.DataTable();

        //    //Adding the Columns
        //    foreach (DataGridViewColumn column in dGV.Columns)
        //    {
        //        dt.Columns.Add(column.HeaderText, column.ValueType);
        //    }

        //    //Adding the Rows
        //    foreach (DataGridViewRow row in dGV.Rows)
        //    {
        //        dt.Rows.Add();
        //        foreach (DataGridViewCell cell in row.Cells)
        //        {
        //            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
        //        }
        //    }

        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        wb.Worksheets.Add(dt, tabName);
        //        wb.SaveAs(filename);
        //    }
        //}  
        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent, string FileName, string FileName2)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));

            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\Report.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;

            if (!string.IsNullOrEmpty(FileName))
            {
                message.Attachments.Add(new Attachment(FileName));
            }

            if (!string.IsNullOrEmpty(FileName2))
            {
                message.Attachments.Add(new Attachment(FileName2));
            }

            //message.Attachments.Add(new Attachment(Chart));

            //bettytseng@acmepoint.com
            //davidhuang@acmepoint.com
            //20191008
            if (SlpName.ToLower() == "sunny")
            {
                message.CC.Add(new MailAddress("bettytseng@acmepoint.com"));
                message.CC.Add(new MailAddress("davidhuang@acmepoint.com"));
            }


            SmtpClient client = new SmtpClient();
            client.Host = "smtp.acmepoint.com";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";
            //group-acmepoint@acmepoint.com
            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

            try
            {
                client.Send(message);
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        SetMsg(String.Format("Failed to deliver message to {0}",
                           ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                       ex.ToString()));
            }

        }

        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));



            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\Report.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "博豐光電");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;

            //message.Attachments.Add(new Attachment(Xls));
            //message.Attachments.Add(new Attachment(Chart));

            SmtpClient client = new SmtpClient();
            client.Host = "smtp.acmepoint.com";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";
            //group-acmepoint@acmepoint.com
            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

            try
            {
                client.Send(message);
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                       // SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        //SetMsg(String.Format("Failed to deliver message to {0}",
                        //    ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                //        ex.ToString()));
            }

        }

        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }


        public bool WriteDataTableToExcel
(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation, string ReporType)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
           // Microsoft.Office.Interop.Excel.Range excelCellrange;

            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            try
            {
               

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = worksheetName;
                WriteDataTableToSheetByArray(dataTable, excelSheet);
                //now save the workbook and exit Excel
                //excelworkBook.SaveAs(saveAsLocation);
                excelworkBook.SaveAs(saveAsLocation, XlFileFormat.xlWorkbookNormal,
                      "", "", Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlNoChange, 
                    1, false, Type.Missing, Type.Missing, Type.Missing);

                excelworkBook.Close();
               
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }
        }

        public bool WriteDataTableToExcel2
(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation, string ReporType)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook= null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            // Microsoft.Office.Interop.Excel.Range excelCellrange;

            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            try
            {
               

                excelworkBook = null;

                System.Data.DataTable dtGroup = dtDistinct(dataTable, "型號");
                excelworkBook = excel.Workbooks.Add(Type.Missing);
                //預設值 一個 or 3
                int wbCount = excelworkBook.Worksheets.Count;
                //workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                //for (int i = 1; i <= wbCount; i++)
                //{
                // 不能全部刪除 //至少要有一個
                //    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i];
                //    excelSheet.Delete();
                //}

                string ModelNo = "";
                for (int i = 0; i <= dtGroup.Rows.Count - 1; i++)
                {
                    
                    try
                    {
                        ModelNo = Convert.ToString(dtGroup.Rows[i][0]);
                        if (ModelNo == "") ModelNo = "無型號";
                    }
                    catch
                    {
                        ModelNo = "無型號";
                    }
                    // Creation a new Workbook
                    
                    // Workk sheet
                    if (i >= excelworkBook.Worksheets.Count)
                    {
                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i+1];
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    excelSheet.Name = ModelNo;
                    //   WriteDataTableToSheetByArray(dataTable, excelSheet);

                    string filterExpression = "";
                    try
                    {
                        filterExpression = string.Format("型號='{0}'", Convert.ToString(dtGroup.Rows[i][0]));
                        if (ModelNo == "無型號")
                        {
                            filterExpression = "Isnull(型號,'') = ''";
                        }
                    }
                    catch
                    {
                        filterExpression = "Isnull(型號,'') = ''";
                    
                    }

                    System.Data.DataTable dtFilter = DataTableFilter(dataTable, filterExpression);
                    WriteDataTableToSheetByArray(dtFilter, excelSheet);
                    
                }

                //// Creation a new Workbook
                //excelworkBook = excel.Workbooks.Add(Type.Missing);

                //// Workk sheet
                //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                //excelSheet.Name = worksheetName;
             //   WriteDataTableToSheetByArray(dataTable, excelSheet);
                //now save the workbook and exit Excel
               // excelworkBook.SaveAs(saveAsLocation); ;

                excelworkBook.SaveAs(saveAsLocation, XlFileFormat.xlWorkbookNormal,
                    "", "", Type.Missing, Type.Missing,
                  XlSaveAsAccessMode.xlNoChange,
                  1, false, Type.Missing, Type.Missing, Type.Missing);


                excelworkBook.Close();
               
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {

                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                
            }
        }

        private static void WriteDataTableToSheetByArray(System.Data.DataTable dataTable,
            Worksheet worksheet)
        {
            int rows = dataTable.Rows.Count + 1;
            int columns = dataTable.Columns.Count;

            var data = new object[rows, columns];

            int rowcount = 0;
            for (int i = 1; i <= dataTable.Columns.Count; i++)
            {
                data[rowcount, i - 1] = dataTable.Columns[i - 1].ColumnName;
            }

            rowcount += 1;
            foreach (DataRow datarow in dataTable.Rows)
            {
                for (int i = 1; i <= dataTable.Columns.Count; i++)
                {

                    // Filling the excel file 
                    data[rowcount, i - 1] = datarow[i - 1].ToString();
                }

                rowcount += 1;
            }

            var startCell = (Range)worksheet.Cells[1, 1];
            var endCell = (Range)worksheet.Cells[rows, columns];
            var writeRange = worksheet.Range[startCell, endCell];

            //aRange.Columns.AutoFit();

            writeRange.Value2 = data;

            writeRange.Columns.AutoFit();
        }


        public System.Data.DataTable GetACME_MAILLIST()
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            //string sql = "SELECT UserCode,UserMail FROM ACME_ARES_MAIL where SysCode='SerialNo' and  Active='Y' and UserCode='terrylee'";
            string sql = "SELECT UserCode,UserMail FROM ACME_ARES_MAIL where SysCode='SerialNo' and  Active='Y' ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            //command.Parameters.Add(new SqlParameter("@DocType", DocType));
            //command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MAIL");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MAIL"];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetACME_MAILLIST();
            dgData.DataSource = dt;

            GridViewAutoSize(dgData);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = dgData.DataSource as System.Data.DataTable;

            if (dt == null || dt.Rows.Count == 0)
            {
                //MessageBox.Show("沒有資料");
                return;
            }

           

            //string FileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
            string FileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xls";

            if (Environment.UserName.ToLower() == "terrylee")
            {
                FileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
            }

            //ExportToExcelFile(dgData, FileName, "Acme");
            WriteDataTableToExcel(dt, "箱序號", FileName, "");
            //WriteDataTableToExcel(dt, "箱序號", FileName, "");
            System.Diagnostics.Process.Start(FileName);
        }

        private System.Data.DataTable dtDistinct(System.Data.DataTable dt,string FieldName)
        {
            System.Data.DataTable dtGroup;

            DataView dv = dt.DefaultView;
            //distinct
            //dtGroup = dv.ToTable(true, "型號");
            dtGroup = dv.ToTable(true, FieldName);
            return dtGroup;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            button15.Enabled = true;

            if (TestFlag == "N")
            {
                lvwUpdate.Items.Clear();
            }

            //AddTextlvw("查詢開始");

//            string Sql = @"select distinct carton 箱號,SHIPPING_NO 序號,INVOICE_NO 發票號碼,
//isnull(case when CHARINDEX('.',MODEL_NO) >0 then
// SUBSTRING(MODEL_NO,1,CHARINDEX('.',MODEL_NO)+1) 
// else
// MODEL_NO
// end,'') 型號 from WH_AUO t0
//inner join (
//select carton,pic from AP_INVOICEIN
//where whno='{0}'
//and pic ='') t1 on t1.CARTON=t0.CARTON_NO
//union all
//select carton 箱號,pic 序號,INV as 發票號碼,
//isnull(case when CHARINDEX('.',t3.MODEL_NO) >0 then
// SUBSTRING(t3.MODEL_NO,1,CHARINDEX('.',t3.MODEL_NO)+1) 
// else
// t3.MODEL_NO
// end ,'')
//from AP_INVOICEIN t2
//left join WH_AUO t3 on t3.INVOICE_NO=t2.INV  and t3.SHIPPING_NO=t2.PIC
//where whno='{0}'
//and pic <>''
//";
           string WhNo = txtWhNo.Text;

//            Sql = string.Format(Sql, WhNo);

//            System.Data.DataTable dt = GetData(Sql);

//           // dt = UpdateBlankRow(dt);

//            //20191022 補箱號
//            dt = UpdateBlankCartonRow(dt);

            string Msg="";
            System.Data.DataTable dt = GetSerialInvoice(WhNo, ref  Msg);
            dgData.DataSource = dt;

            GridViewAutoSize(dgData);



            // AddTextlvw("核對數量",4);
            //核對數量
            string Sql2 = "select isnull(Sum(convert(int,Quantity)),0) Qty from WH_Item where ShippingCode='{0}'";
            Sql2 = string.Format(Sql2, WhNo);
            Int32 recCount = GetDataCount(Sql2);
            // //MessageBox.Show(recCount.ToString());
            if (recCount != dt.Rows.Count)
            {
                if (TestFlag == "N")
                {
                    //MessageBox.Show("資料筆數不符合");
                }
                AddTextlvw(WhNo + "-核對數量-資料筆數不符合", 1);
            }



            //string filterExpression = "Isnull(型號,'') = ''";
                    
            //  System.Data.DataTable dtFilter = DataTableFilter(dt, filterExpression);
            //  dgData.DataSource = dtFilter;

            //System.Data.DataTable dtGroup;

            //DataView dv = dt.DefaultView;
            ////distinct
            //dtGroup = dv.ToTable(true, "型號");
            //dgData.DataSource = dtGroup;
        }

        private void fmAcmeCusSerial_FormClosed(object sender, FormClosedEventArgs e)
        {
            sw.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string Sql = "select CardName 客戶 from [ACME_CUS_SET]";
            System.Data.DataTable dt = GetData(Sql);
            dgCustomer.DataSource = dt;
            dgCustomer.Columns[0].Width = 200;
            GridViewAutoSize(dgCustomer);
            tabControl1.SelectedTab = tabPage1;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //20200102 增加 Invoice

            //個案分析
//工單:WH20191023022X 
//異常訊息:資料不符合->工單數量=182 箱序數量=185 
            //PZ190B1962801870 -> Qty = 2 (滿箱15片)

              string WhNo = txtWhNo.Text;

              string Sql1 = @"select count(distinct SHIPPING_NO)  from WH_AUO where CARTON_NO='{0}'";

            string Sql2 = @"select Inv Invoice,carton 箱號,pic 序號,qty 數量, 0 as 比對數量, '' 結果 from AP_INVOICEIN
where whno='{0}' order by  carton,pic";


            string Sql3 = @"select carton 箱號,pic 序號,qty 數量 from AP_INVOICEIN
where whno='{0}' order by  carton,pic";

//            string Sql3 = @"select carton 箱號,pic 序號,qty 數量 from AP_INVOICEIN
//where whno='{0}' order by  carton,pic";

            //Check AP_INVOICEIN Qty in WH_AUO
            Sql2 = string.Format(Sql2,WhNo);

            Sql3 = string.Format(Sql3, WhNo);
                
            System.Data.DataTable dt = GetData(Sql2);
            dgError.DataSource = dt;

            //System.Data.DataTable dtOrigin = dt.DefaultView.ToTable();
            System.Data.DataTable dt3 = GetData(Sql3);
            dgCarton.DataSource = dt3;

            DataRow dr;
            string Carton;
            string SerialNo;
            string Sql = "";
            string Qty="";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];
                Carton = Convert.ToString(dr["箱號"]);
                SerialNo = Convert.ToString(dr["序號"]);

                Qty= Convert.ToString(dr["數量"]).Trim();
                if (string.IsNullOrEmpty(Carton) || !string.IsNullOrEmpty(SerialNo)) continue;

                Sql = string.Format(Sql1, Carton);
                System.Data.DataTable dtCount = GetData(Sql);
                Int32 RowCount = Convert.ToInt32(dtCount.Rows[0][0]);

                dr.BeginEdit();
                dr["比對數量"] = RowCount;

                if (RowCount.ToString() != Qty)
                {
                    dr["結果"] = "異常";
                }
                dr.EndEdit();

            }

            DataView dv = dt.DefaultView;

            dv.RowFilter = "結果='異常'";
            dv.Sort = "結果 desc";

            GridViewAutoSize(dgError);
        }

        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  查無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }

            //create html & table
            //strB.AppendLine("<html><body><center><table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border-collapse:collapse;'>");
            strB.AppendLine("<tr class='HeaderBorder'>");
            //cteate table header
            for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
            {
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                {

                    //if (i != 0)
                    //{
                    //    strB.AppendLine("<tr class='HeaderBorder'>");
                    //    for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                    //    {
                    //        strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                    //    }
                    //    strB.AppendLine("</tr>");
                    //}

                    //處理鍵值




                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;
                }
                else
                {
                    tmpKeyValue = "";
                }


                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }



                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                }


                for (int d = 1; d <= dg.Rows[i].Cells.Count - 1; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    // HttpUtility.HtmlDecode("&nbsp;&nbsp;&nbsp;")

                    if (dgvc.ValueType == typeof(Int32))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Int32 x = Convert.ToInt32(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }

                    else if (dgvc.ValueType == typeof(Decimal) || dgvc.ValueType == typeof(Double))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Decimal x = Convert.ToDecimal(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0.00") + "</td>");
                        }


                    }
                    else
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {

                            if (dg.Columns[dgvc.ColumnIndex].HeaderText.IndexOf("日期") >= 0)
                            {
                                string sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                             dgvc.Value.ToString().Substring(4, 2) + "/" +
                                             dgvc.Value.ToString().Substring(6, 2);


                                strB.AppendLine("<td>" + sDate + "</td>");
                            }
                            else
                            {
                                strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                            }
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }


        private void button8_Click(object sender, EventArgs e)
        {
            //System.Data.DataTable dt = GetNext();
            string sql ="select * from ACME_CUS_TIME";
            System.Data.DataTable dt = GetData(sql);
            dgData.DataSource = dt;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //ACME_MAIL_LOG //沒有 
            string sql = "select * from ACME_MAIL_LOG where DocType = 'SN' and MailDate>='{0}' and MailDate<='{1}'";
            sql = string.Format(sql,txtStartDate.Text ,txtEndDate.Text);
            System.Data.DataTable dt = GetData(sql);
            dgData.DataSource = dt;

            GridViewAutoSize(dgData);
        }

        private void button10_Click(object sender, EventArgs e)
        {

            button4.Enabled = true;
            button15.Enabled = false;

            //MSZ1579571/W041580606(沒有資料)

            string Sql = @"select distinct carton 箱號,SHIPPING_NO 序號,STOCK_IN_WEEK WC from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where INV='{0}'
and pic ='') t1 on t1.CARTON=t0.CARTON_NO
union all
select distinct carton 箱號,pic 序號,'' WC from AP_INVOICEIN
where INV='{0}'
and pic <>''
";
            string WhNo = txtInvoiceNo.Text;

            Sql = string.Format(Sql, WhNo);

            System.Data.DataTable dt = GetData(Sql);

            //dt = UpdateBlankRow(dt);

            //20191022 補箱號
            dt = UpdateBlankCartonRow(dt);


            dgData.DataSource = dt;

            GridViewAutoSize(dgData);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            button4.Enabled = true;
            button15.Enabled = false;

            string Sql = @"select distinct carton_no 箱號,SHIPPING_NO 序號,STOCK_IN_WEEK WC from WH_AUO
where Invoice_NO='{0}'
";
            string WhNo = txtInvoiceNo.Text;

            Sql = string.Format(Sql, WhNo);

            System.Data.DataTable dt = GetData(Sql);

            //dt = UpdateBlankRow(dt);

            //20191022 補箱號
            dt = UpdateBlankCartonRow(dt);


            dgData.DataSource = dt;

            GridViewAutoSize(dgData);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            //記錄數量不符的工單號碼
            //追踨是否有處理
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
//            select distinct carton 箱號,SHIPPING_NO 序號 from WH_AUO t0
//inner join (
//select carton,pic from AP_INVOICEIN
//where INV in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')
//and pic ='') t1 on t1.CARTON=t0.CARTON_NO
//union all
//select distinct carton 箱號,pic 序號 from AP_INVOICEIN
//where INV in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')
//and pic <>'' 



//            select distinct carton,pic from AP_INVOICEIN
//where inv in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')

//select distinct shipping_no,carton_no from WH_AUO
//where INVOICE_NO in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')

            //取得 工單號 WH20191129040X -> 3200 
//            select whno, carton,pic,inv from AP_INVOICEIN
//where inv in ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')
        //WH20191129040X
        //  ('MSZ1602246','MSZ1608243','Z191608325','Z191608330','Z191609507')
            //箱 ->Qty 3240 PCS
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //WH20200103019X
//            select * from AP_INVOICEIN
//where whno='WH20200103019X'
        }

        private void button15_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtData = dgData.DataSource as System.Data.DataTable;

            if (dtData == null || dtData.Rows.Count == 0)
            {
                //MessageBox.Show("沒有資料");
                return;
            }

            string WhNo = txtWhNo.Text;

            //string FileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
            //string FileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xls";

            string OutFileName = GetExePath() + "\\" + WhNo + ".xls";
            WriteDataTableToExcel(dtData, "箱序號", OutFileName, "");

            string CheckMsg = "";
            dtData = GetSerialInvoice(WhNo, ref CheckMsg);

            if (dtData == null || dtData.Rows.Count == 0)
            {
                // //MessageBox.Show("沒有資料");
                // return;
                //20191016 沒有資料先不發送
                
            }
            else
            {
                //OutFileName2 = GetExePath() + "\\" + WhNo + "_Inv.xlsx";
                string OutFileName2 = GetExePath() + "\\" + WhNo + "_Inv.xls";

                // System.Data.DataTable dtGroup = dtDistinct(dtData,"型號");
                WriteDataTableToExcel2(dtData, "箱序號", OutFileName2, "");
            }
            //if (Environment.UserName.ToLower() == "terrylee")
            //{
            //    OutFileName = GetExePath() + "\\" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
            //}

           
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dgData);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("請輸入客戶");
                return;
            }
            System.Data.DataTable T1 = GETT1(textBox2.Text, textBox3.Text, textBox1.Text);
            dataGridView1.DataSource = T1;
        }

        public static System.Data.DataTable GETT1(string DOCDATE1, string DOCDATE2, string INV)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("      select distinct carton 箱號,SHIPPING_NO 序號,INV INVOICE,CARD 客戶,ITEMCODES 型號,STOCK_IN_WEEK   from WH_AUO t0");
            sb.Append(" inner join (");
            sb.Append(" select carton,pic,INV,CARD,ITEMCODE ITEMCODES from AP_INVOICEIN");
            sb.Append(" where CARD  LIKE '%" + INV + "%'");
            sb.Append(" AND DOCDATE BETWEEN @DOCDATE1 AND @DOCDATE2");
            sb.Append(" and pic ='') t1 on t1.CARTON=t0.CARTON_NO");
            sb.Append(" union all");
            sb.Append(" select distinct carton 箱號,pic 序號,INV,CARD,ITEMCODE ITEMCODES,''   from AP_INVOICEIN");
            sb.Append(" where CARD LIKE '%" + INV + "%'");
            sb.Append(" AND DOCDATE BETWEEN @DOCDATE1 AND @DOCDATE2");
            sb.Append(" and pic <>''");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE1", DOCDATE1));
            command.Parameters.Add(new SqlParameter("@DOCDATE2", DOCDATE2));
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

        private void button18_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToCSV(dataGridView1, Environment.CurrentDirectory + @"\SHIPPING" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button19_Click(object sender, EventArgs e)
        {
            button4.Enabled = true;
            button15.Enabled = false;

            //WH20191223014X 台亞 Shipping三角 連結 倉庫工單

            ErrorFlag = "N";

            if (TestFlag == "N")
            {
                lvwUpdate.Items.Clear();
            }

            //AddTextlvw("查詢開始");

            //20191021 //加入 distinct
            string Sql = @"select distinct  PALLET_NO,CARTON_NO ,SHIPPING_NO from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where whno='{0}'
and pic ='') t1 on t1.CARTON=t0.CARTON_NO
union all
select distinct  PALLET_NO,CARTON_NO ,SHIPPING_NO from WH_AUO t0
inner join (
select carton,pic from AP_INVOICEIN
where whno='{0}'
and pic ='') t1 on t1.pic=t0.SHIPPING_NO
";
            string WhNo = txtWhNo.Text;

            Sql = string.Format(Sql, WhNo);

            System.Data.DataTable dt = GetData(Sql);

            //dt = UpdateBlankRow(dt);


            dgData.DataSource = dt;

            GridViewAutoSize(dgData);



       

            //AddTextlvw(TestRec+"完成");
            //System.Diagnostics.Process.Start(OutPutFile);
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {

        }

        public System.Data.DataTable GETF1(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT DISTINCT SHIPPING_NO,MAX(SHIPPING_TIME) SHIPPING_TIME,MAX(INVOICE_NO) INVOICE_NO,MAX(CARTON_NO) CARTON_NO,MAX(PALLET_NO) PALLET_NO,MAX(MODEL_NO) MODEL_NO,MAX(FINAL_GRADE) FINAL_GRADE,MAX(PRODUCT_TYPE) PRODUCT_TYPE,MAX(STOCK_IN_WEEK)  STOCK_IN_WEEK FROM ACMESQLSP.DBO.WH_AUO T0 ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.AP_INVOICEIN T1 ON ( T0.CARTON_NO =T1.CARTON)    ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.WH_Item4 T2 ON (T1.WHNO =T2.ShippingCode AND T1.ITEMCODE =T2.ItemCode)    ");
            sb.Append("          WHERE  T2.DOCENTRY=50295 AND T1.ITEMCODE='M238HVN01.000E2' AND INVOICE_NO='Z191840510' AND T1.QTY <> 1 ");
            sb.Append("		  AND T0.CARTON_NO <> ''");
            sb.Append("          GROUP BY SHIPPING_NO ");
            sb.Append("          UNION ALL  ");
            sb.Append("          SELECT DISTINCT  SHIPPING_NO,MAX(SHIPPING_TIME) SHIPPING_TIME,MAX(INVOICE_NO) INVOICE_NO,MAX(CARTON_NO) CARTON_NO,MAX(PALLET_NO) PALLET_NO,MAX(MODEL_NO) MODEL_NO,MAX(FINAL_GRADE) FINAL_GRADE,MAX(PRODUCT_TYPE) PRODUCT_TYPE,MAX(STOCK_IN_WEEK)  STOCK_IN_WEEK FROM ACMESQLSP.DBO.WH_AUO  T0 ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.AP_INVOICEIN T1 ON ( T0.SHIPPING_NO  =T1.PIC )    ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.WH_Item4 T2 ON (T1.WHNO =T2.ShippingCode AND T1.ITEMCODE =T2.ItemCode)    ");
            sb.Append("          WHERE  T2.DOCENTRY=50295 AND T1.ITEMCODE='M238HVN01.000E2' AND INVOICE_NO='Z191840510'  AND T1.QTY=1 ");
            sb.Append("          GROUP BY SHIPPING_NO ");
            sb.Append("          UNION ALL  ");
            sb.Append("          SELECT DISTINCT SHIPPING_NO,MAX(SHIPPING_TIME) SHIPPING_TIME,MAX(INVOICE_NO) INVOICE_NO,MAX(CARTON_NO) CARTON_NO,MAX(PALLET_NO) PALLET_NO,MAX(MODEL_NO) MODEL_NO,MAX(FINAL_GRADE) FINAL_GRADE,MAX(PRODUCT_TYPE) PRODUCT_TYPE,MAX(STOCK_IN_WEEK)  STOCK_IN_WEEK FROM ACMESQLSP.DBO.WH_AUO T0 ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.AP_INVOICEIN T1  ON (T0.PALLET_NO =T1.PLT)       ");
            sb.Append("          LEFT JOIN ACMESQLSP.DBO.WH_Item4 T2 ON (T1.WHNO =T2.ShippingCode AND T1.ITEMCODE =T2.ItemCode)    ");
            sb.Append("          WHERE T2.DOCENTRY=50295 AND T1.ITEMCODE='M238HVN01.000E2' AND INVOICE_NO='Z191840510'  AND T1.QTY <> 1 ");
            sb.Append("		  AND ISNULL(T0.PALLET_NO,'')<>''");
            sb.Append("          GROUP BY SHIPPING_NO ");



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
    }
}

//工單
//SELECT COUNT(*) FROM WH_main WHERE ([WH_main].[closeDay] LIKE '%20190%') AND ([WH_main].[createName] LIKE '%sunny')
//select Sum(convert(int,Quantity)) Qty from WH_Item where ShippingCode='WH20190102004X'

//AutoHotkey
//https://sites.google.com/view/ahktool/


//lleyton Web 
//http://wf.acmepoint.net/RMA/JS/FAE2.aspx
//exec sp_executesql N' SELECT  DISTINCT SHIPPING_TIME,SHIPPING_NO,INVOICE_NO,CARTON_NO,PALLET_NO,MODEL_NO,FINAL_GRADE,PRODUCT_TYPE,STOCK_IN_WEEK
//FROM ACMESQLSP.DBO.WH_AUO  
//WHERE INVOICE_NO=@INVOICE AND ITEMCODE=@ITEMCODE 
//',N'@INVOICE nvarchar(10),@ITEMCODE nvarchar(15)',@INVOICE=N'MSZ1550372',
//@ITEMCODE=N'G101UAN01.00002'


 //using (XLWorkbook wb = new XLWorkbook())
 //                   {
 //                       wb.Worksheets.Add(dt);

//20191023 Sikuli
//http://puremonkey2010.blogspot.com/2014/01/ui-automation-sikuli.html
//type(u"中文")
//while not exists
//popup(u"下線了")
//findAll(…)
//x=find().right().find()
//for i in range(3):
　//　click(x)
//if …: con1 else: con2
//App.focus
//App.close
//wait() and waitVanish() 
//https://subscription.packtpub.com/book/application_development/9781782167877/1/ch01lvl1sec06/top-13-features-you-need-to-know-about
//highlight()
//region -> r.wait()
//Screen() capture
//findAll() # find all matches
//mm = SCREEN.getLastMatches()
//from guide import *
//paste()

//# First, check if image exists for n seconds
//# Pattern and similar, set how much similar your image will be.
//if exists(Pattern("GoogleSearch.png").similar(0.8), time_in_seconds):
//    if exists(Pattern("FeelingLucky.png").similar(0.6), time_in_seconds):
//        click(Pattern("FeelingLucky.png").similar(0.6))
//C:\Windows\explorer.exe \\acmesrv01\Public\AUO_出貨文件
//http://wyj-learning.blogspot.com/2018/06/sikuli_30.html

//password = Do.input("please enter your secret", "Secret", "defaultSecret", True, 10)
//# the dialog's input field displays the text as dots per character
//if not password:
//  # password is empty or dialog autoclosed
//  print "not allowed - exiting"
//  exit(1)
//# we can proceed

//where = Region(0,0,300,300)
//result = Do.input("please fill in", "A filename", "someImage.png", where)
//# the dialog will display somewhere in the upper left of the screen
//# with a box title as "A filename"
//# and a preset input field containing "someImage.png"
//if not result:
//  # input field was left empty
//  print "we will use a default file name"
//else:
//  print "we will use as filename: " + result

//cmd = r'c:\Program Files\myapp.exe -x "c:\Some Place\some.txt" >..\log.txt'
//openApp(cmd)

//# using an existing window if possible
//myApp = App("Firefox")
//if not myApp.window(): # no window(0) - Firefox not open
//        App.open("c:\\Program Files\\Mozilla Firefox\\Firefox.exe")
//        wait(2)
//myApp.focus()
//wait(1)
//type("l", KEY_CTRL) # switch to address field
//setShowActions(True)
//myApp=App("Firefox")
//myApp.focus()
//type(" ", KEY_ALT) # Open windows sizing control
//wait(1)
//type("x") # Maximize the window
    //搜尋：
    //    find(圖片)：在搜尋範圍內找出最佳比對結果
    //    findAll(圖片)：在搜尋範圍內找出所有比對結果
    //    wait(圖片, [time out])：在範圍內等待圖片出現，最多等待 [time out] 秒
    //    waitVanish(圖片, [time out])：在範圍內等待圖片消失，最多等待 [time out] 秒
    //    exists(圖片)：檢查範圍內是否存在目標，若否回傳 None，不會跳例外處理(與find() 不同)
    //滑鼠動作：
    //    click(圖片)：滑鼠左鍵單擊
    //    doubleClick(圖片)：滑鼠左鍵雙擊
    //    rightClick(圖片)：滑鼠右鍵單擊
    //    hover(圖片)：拖移滑鼠
    //    dragDrop(圖片A, 圖片B )：將 圖片A拖到 圖片B放開(例如：將檔案拖移至垃圾桶)
    //鍵盤動作：
    //    type(text)：於視窗中鍵盤輸入 text
    //    type(圖片, text)：點擊圖片取得焦點後，鍵盤輸入 text
    //    pasts(text)：透過剪貼簿，將text貼入畫面中
    //    pasts(圖片, text)：點擊圖片取得焦點後，透過剪貼簿，將text貼入畫面中
//https://wenku.baidu.com/view/2b4da4b5960590c69ec3766d.html

//https://chromium.googlesource.com/external/dart/+/532c5163f199b4b20d641211eeac6dfae2472ce9/dart/editor/ft/sikuli/util.sikuli/util.py

//def _key_cmd(key):
//  "Send meta-key to the editor."
//  if is_OSX():
//    type(key, KeyModifier.META)
//  else:
//    type(key, KeyModifier.CTRL)

// xlWorkSheet.Shapes.AddPicture(@"C:\pic.JPG", MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 50, 300, 45);

////Gathering information of an old image
//IPictureShape oldImage = worksheet.Pictures[0];
//int leftPosition = oldImage.Left;
//int topPosition = oldImage.Top;
//int height = oldImage.Height;
//int width = oldImage.Width;



//Removing the old image
//worksheet.Pictures[0].Remove();

////Create image 
//Image image = Image.FromStream(imageStream);
  
////Replace the image with new one and assigning its bounds
//IPictureShape newImage = worksheet.Pictures.AddPicture(image,"New.png",ExcelImageFormat.Png);                newImage.Left = leftPosition;
//newImage.Top = topPosition;
//newImage.Height = height;
//newImage.Width = width;

//  ActiveSheet.ListObjects("Table1").Resize Range("$A$1:$B$33")
// ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B$183"), , xlYes).Name = _
    //    "表格3"
    //Range("表格3[#All]").Select
    //ActiveSheet.ListObjects("表格3").TableStyle = "TableStyleMedium13"
 //Excel.ListObject tbl = (Excel.ListObject)WSheet.ListObjects.AddEx(
 //       SourceType: Excel.XlListObjectSourceType.xlSrcRange,
 //       Source: range,
 //       XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes);
// _excelSheetItems = excelSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, _cells.GetRange(1, TittleRow, ResultColumn, TittleRow + 1), XlListObjectHasHeaders: XlYesNoGuess.xlYes);
 //excelSheet.Cells.Columns.AutoFit();
 //           excelSheet.Cells.Rows.AutoFit();


//Sub UnMergeFill()

//Dim cell As Range, joinedCells As Range

//For Each cell In ThisWorkbook.ActiveSheet.UsedRange
//    If cell.MergeCells Then
//        Set joinedCells = cell.MergeArea
//        cell.MergeCells = False
//        joinedCells.Value = cell.Value
//    End If
//Next

//End Sub
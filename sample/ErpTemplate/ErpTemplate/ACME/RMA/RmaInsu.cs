using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Transactions;
using System.Configuration;
using System.Net;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Web.UI;
using System.Collections;


namespace ACME
{
    public partial class RmaInsu : Form
    {
        public RmaInsu()
        {
            InitializeComponent();
        }
        private System.Data.DataTable OrderData;
        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\INSU.xls";

                string RATE = "30.5";
                System.Data.DataTable T1 = GetRATE();
                if (T1.Rows.Count > 0)
                {

                    RATE = T1.Rows[0][0].ToString();
                }

                OrderData = GetOrderData6(RATE);
                
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable GetOrderData()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.shippingCode JOBNO,tradeCondition 貿易條件,receiveDay 運送方式,boardCountNo 貿易形式,");
            sb.Append(" quantity 出口日期,SEQNO LINE,INDESCRIPTION 品名,MARKNOS MODEL,InvoiceNo_seq VER,");
            sb.Append(" InQty 數量,RMANO,VENDERNO,CodeName 客戶簡稱 FROM RMA_MAIN T0 LEFT JOIN  RMA_invoiced T1 ON (T0.shippingCode=T1.shippingCode)");
        //    sb.Append(" WHERE T0.SHIPPINGCODE='RMA20170301001X'");

            sb.Append("  where T0.SHIPPINGCODE in ( " + textBox2.Text + ") ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData6(string RATE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select a.shipToDate+'%'  保險費率,case when isnull(a.QUANTITY,'')=''  then a.arriveDay else a.QUANTITY end COLLATE  Chinese_Taiwan_Stroke_CI_AS  日期,a.shippingcode 工單號碼,a.tradeCondition 貿易條件 ,a.boardcountno 貿易類別 ,b.INDescription 品名");
            sb.Append("              ,b.InQty 數量,a.receiveday 運送方式,a.cardname 客戶,b.amount 價格,");
            sb.Append("              保險費=case a.receiveday ");
            sb.Append("              when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("              when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("              when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end");
            sb.Append("              ,匯率=@RATE from RMA_main a ");
            sb.Append("              left join RMA_invoiced b on(a.shippingcode=b.shippingcode)");
            sb.Append("               WHERE A.CFS='checked'");
            sb.Append("              AND (a.QUANTITY between @startday and @endday or arriveDay between @startday and @endday or CLOSEDAY between @startday and @endday)");
            sb.Append("              UNION ALL");
            sb.Append("                select a.shipToDate+'%'  保險費率,REPLACE(REPLACE(REPLACE(a.createDate,'.',''),'/',''),'-','') COLLATE  Chinese_Taiwan_Stroke_CI_AS 日期,a.shippingcode 工單號碼,a.tradeCondition 貿易條件 ,a.boardcountno 貿易類別 ,b.INDescription 品名");
            sb.Append("              ,b.InQty 數量,a.receiveday 運送方式,a.cardname 客戶,b.amount 價格,");
            sb.Append("              保險費=case a.receiveday ");
            sb.Append("              when 'SEA'  then case a.dollarsKind when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("              when 'AIR'  then case a.dollarsKind when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("              when 'TRUCK'  then case a.dollarsKind when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end");
            sb.Append("              ,匯率=@RATE from RMA_mainSZ a ");
            sb.Append("              left join RMA_invoicedSZ b on(a.shippingcode=b.shippingcode)");
            sb.Append("               WHERE A.dollarsKind='checked'");
            sb.Append("              AND REPLACE(REPLACE(REPLACE(a.createDate,'.',''),'/',''),'-','') between @startday and @endday");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startday", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@endday", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@RATE", RATE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRATE()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT RATE FROM ORTT where Convert(varchar(10),ratedate,112)=@RATE and currency='USD' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RATE", GetMenu.Day()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void RmaInsu_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyy");
            textBox3.Text = DateTime.Now.ToString("yyyy");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\DHL快遞費用.xls";

  

                OrderData = GetOrderData();

                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
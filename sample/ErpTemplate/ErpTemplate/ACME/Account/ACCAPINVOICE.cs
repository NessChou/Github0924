﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Management;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net.Mime;
using System.Web.UI;

namespace ACME
{
    public partial class ACCAPINVOICE : Form
    {
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP_TEST;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string FA = "acmesql98";
        public ACCAPINVOICE()
        {

            InitializeComponent();
            FormLoad();
        }
        private void FormLoad() 
        {
            UtilSimple.SetLookupBinding(cmbBU, GetMenu.MoneyBU("ACCAPINVOICE"), "DataTEXT", "DataValue");
            if (globals.DBNAME == "進金生")
            {
                strCn98 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                FA = "acmesql02";
            }
            SetDateTimeTextBox();//起訖設當月第一天跟最後一天,過帳日期為當日

            SetdgvInvoiceTrack();

            cmbBU.SelectedValue = "ALL";


        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
      
        private void btnQuery_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetAccAPInvoice();
            CheckInvoiceTrack(ref dt);//友達要依據發票號碼填入對應統編
            //System.Data.DataTable dtData = CombineDataTable(dt);//合併por1

            dgvAccApInvoice.DataSource = dt;

            System.Data.DataTable dtAPInvoice = GetAPInvoice();//又突然要加新功能顯示ap發票
            
            dgvApinvoice.DataSource = dtAPInvoice;

            //dgvAccApInvoice.Rows[dgvAccApInvoice.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Yellow;
            dgvApinvoice.AllowUserToAddRows = false;
           
        }
        private System.Data.DataTable GetAccAPInvoice()
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT DISTINCT( t0.docdate ),t0.docdate por1DocDate,t0.docentry OPDNDocEntry ,");
            sb.Append("             STUFF((SELECT '/' + CAST (T22.baseref AS NVARCHAR) FROM POR1 T22 LEFT JOIN pdn1 t11 on t11.baseentry = T22.docentry and  t11.baseline = T22.linenum  AND t11.BASETYPE = 22 WHERE (t11.DOCENTRY = T1.DOCENTRY) GROUP BY T22.BaseRef  FOR XML PATH('')),1,1,'') AS por1BaseEntry, ");
            sb.Append("             STUFF((SELECT '/'+ CAST (T22.DOCENTRY AS NVARCHAR) FROM POR1 T22 LEFT JOIN pdn1 t11 on t11.baseentry = T22.docentry and  t11.baseline = T22.linenum  AND t11.BASETYPE = 22 WHERE (t11.docentry = T1.DOCENTRY) GROUP BY T22.DocEntry FOR XML PATH('')),1,1,'') AS por1Docentry, ");
            sb.Append("             t3.cardcode ,t3.cardname ,  ");
            sb.Append("             (select sum(t1.quantity) from pdn1 t1 where t0.docentry = t1.DocEntry) Quantity,  ");
            sb.Append("             (cast(t0.doctotalsy as int) -cast(t0.VatSumSy as int)) UnTax, t0.VatSumSy ,t0.DocTotalSy ,t0.U_ACME_INV,t0.U_ACME_Invoice shipdate,t3.doccur Currency,'' OriCurrencyAmount,t0.U_ACME_RATE1,t0.u_acme_lc,t0.U_PC_BSDAT u_acme_shipday,t0.U_PC_BSINV ,t4.LicTradNum TaxIdNumber  ,T4.U_PC_BSTY1 InvoiceType,'' Signoff ");
            sb.Append("             FROM opdn t0  ");
            sb.Append("             LEFT JOIN pdn1 t1 on t0.DocEntry =t1.docentry ");
            sb.Append("             LEFT JOIN por1 t2  on (t1.baseentry=T2.docentry and  t1.baseline=t2.linenum  AND t1.BASETYPE=22)  ");
            sb.Append("             LEFT JOIN opor t3  on t2.docentry = t3.docentry");
            sb.Append("             LEFT JOIN OCRD t4  on t3.CARDCODE = t4.cardcode");
            sb.Append("             LEFT JOIN OITM T11 ON t2.ITEMCODE = T11.ITEMCODE  ");

            sb.Append("             WHERE substring(t0.cardcode,1,1) in ('S','U')  And t1.itemcode not in (select itemcode from oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))  And  ");
            sb.Append("              substring(t1.itemcode,1,2) <> 'ZR'  AND T0.DOCSTATUS='O'  ");
            if (txbShipDateStart.Text != "" || txbShipDateEnd.Text != "")
            {
                sb.Append("            and Convert(varchar(10),t0.DocDate,112) between '" +txbShipDateStart.Text + "' and '" + txbShipDateEnd.Text + "'");
            }


            if (txbCardCode.Text != "")
            {
                string[] cardcode = txbCardCode.Text.Split('、');
                for (int i = 0; i < cardcode.Length; i++)
                {
                    if (i == 0)
                    {
                        sb.Append("           and  t3.cardname like '%" + cardcode[i] + "%'");
                    }
                    else
                    {
                        sb.Append("             or t3.cardname like '%" + cardcode[i] + "%'");
                    }

                }

            }
            


            if (cmbBU.Text != "")
            {
                switch (cmbBU.SelectedValue.ToString()) 
                {
                    case "ALL":
                        break;

                    case "ADP+AUO全部":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");
                        break;

                    case "AUO全部":
                    case "ADP全部":
                        sb.Append(" and T0.CARDCODE like '%" + cmbBU.SelectedValue.ToString() + "%'  ");
                        break;

                    case "Non AUO/ADP":
                    case "NON AUO/ADP":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5) NOT IN ('S0001','S0623') ");
                        break;

                }
             
            }
            sb.Append("  ORDER BY  t0.docentry ASC");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "AccApInvoice");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetAPInvoice()
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT  distinct(T1.DocEntry) APInvoiceDocentry,(Convert(varchar(10),T1.DocDate,112)) AP過帳日期,Convert(varchar(10),T1.DocDate,112) AP發票日期,CAST(T1.DocEntry as nvarchar) AP發票號碼,Convert(varchar(10),T2.DocDate,112) 收踩過帳日期,");
            sb.Append("             STUFF((SELECT '/' + CAST (T22.BaseRef AS NVARCHAR) FROM PCH1 T22 WHERE (T22.DOCENTRY = T2.DOCENTRY and t22.linenum = t2.linenum) GROUP BY T22.BaseRef  FOR XML PATH('')),1,1,'') AS 收貨採購單單號, ");
            sb.Append("             T1.CardCode 廠商編號,T1.CardName 廠商名稱, ");
            sb.Append("             (select CONVERT(NVARCHAR(20), CAST(sum(t2.quantity) AS Money),1) from PCH1 t2 where t1.docentry = t2.DocEntry) as 數量,  ");
            sb.Append("             CONVERT(NVARCHAR(20),CAST((cast(T1.DocTotal as int)  - cast(T1.VatSum as int)) AS Money),1)  未稅總計, ");
            sb.Append("             CONVERT(NVARCHAR(20),CAST(T1.VatSum AS Money),1) 稅額,CONVERT(NVARCHAR(20),CAST(T1.DocTotal AS Money),1) 總計,T1.U_ACME_INV InvoiceNo,Convert(varchar(10),T1.U_ACME_Invoice,112) Invoice日期,  ");
            sb.Append("             T3.U_ACME_GradePrice 原始幣別,CONVERT(NVARCHAR(20),CAST((cast(T1.DocTotal as float) /cast(T1.U_ACME_rate1 as float)) AS Money),1) 原幣金額,T1.U_ACME_rate1 匯率,T1.u_acme_lc LC,T1.U_PC_BSINV AP發票發票號碼,T1.U_PC_BSNOT 統一編號  ");
            sb.Append("             FROM OPCH T1  ");
            sb.Append("             LEFT JOIN PCH1 T2 ON  T1.DocEntry = T2.DocEntry ");
            sb.Append("             LEFT JOIN OPDN T3 ON T2.BaseRef = T3.DocEntry ");

            sb.Append("             WHERE T2.BaseType = '20' ");
            if (txbAPDate.Text != "" )
            {
                sb.Append("            and Convert(varchar(10),T1.DocDate,112) = '" + txbAPDate.Text + "'");
            }


            if (txbCardCode.Text != "")
            {
                string[] cardcode = txbCardCode.Text.Split('、');
                for (int i = 0; i < cardcode.Length; i++)
                {
                    if (i == 0)
                    {
                        sb.Append("           and  T1.cardname like '%" + cardcode[i] + "%'");
                    }
                    else
                    {
                        sb.Append("             or T1.cardname like '%" + cardcode[i] + "%'");
                    }

                }

            }
            else
            {
                //只抓S開頭的公司
                sb.Append("           and  t3.CardCode like 'S%'");
            }

            if (cmbBU.Text != "")
            {
                switch (cmbBU.SelectedValue.ToString())
                {
                    case "ALL":
                        break;

                    case "ADP+AUO全部":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");
                        break;

                    case "AUO全部":
                    case "ADP全部":
                        sb.Append(" and T0.CARDCODE like '%" + cmbBU.SelectedValue.ToString() + "%'  ");
                        break;

                    case "Non AUO/ADP":
                    case "NON AUO/ADP":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5) NOT IN ('S0001','S0623') ");
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5) NOT IN ('S0001','S0623') ");
                        break;

                }

            }
            /*
            if (txbDocDate.Text != "")
            {
                sb.Append("            and Convert(varchar(10),T1.DocDate,112) = '" + txbDocDate.Text + "'");
            }
            */
            sb.Append("            UNION");
            //最後一行總計
            sb.Append("             SELECT '','','','','總計','','','', ");
            sb.Append("             CONVERT(NVARCHAR(20), CAST(sum(t2.quantity) AS Money),1)as 數量, ");
            sb.Append("             CONVERT(NVARCHAR(20), CAST(sum(cast(T1.DocTotal as decimal) - cast(T1.VatSum as decimal)) AS Money),1)as 未稅總計, ");
            sb.Append("             CONVERT(NVARCHAR(20), CAST(sum(T1.VatSum) AS Money),1)as 稅額,  ");
            sb.Append("             CONVERT(NVARCHAR(20), CAST(sum(T1.DocTotal) AS Money),1)as 總計,'','','',  ");
            sb.Append("             CONVERT(NVARCHAR(20), CAST(sum(cast(T1.DocTotal as float)/cast(T1.U_ACME_rate1 as float)) AS Money),1)as 原幣金額,'','','',''  ");
            sb.Append("             FROM OPCH T1  ");
            sb.Append("             LEFT JOIN PCH1 T2 ON  T1.DocEntry = T2.DocEntry ");
            sb.Append("             LEFT JOIN OPDN T3 ON T2.BaseRef = T3.DocEntry ");

            sb.Append("             WHERE T2.BaseType = '20' ");
            if (txbShipDateStart.Text != "" || txbShipDateEnd.Text != "")
            {
                sb.Append("            and Convert(varchar(10),T1.DocDate,112) between '" + txbShipDateStart.Text + "' and '" + txbShipDateEnd.Text + "'");
            }


            if (txbCardCode.Text != "")
            {
                string[] cardcode = txbCardCode.Text.Split('、');
                for (int i = 0; i < cardcode.Length; i++)
                {
                    if (i == 0)
                    {
                        sb.Append("           and  T1.cardname like '%" + cardcode[i] + "%'");
                    }
                    else
                    {
                        sb.Append("             or T1.cardname like '%" + cardcode[i] + "%'");
                    }

                }

            }
            else
            {
                //只抓S開頭的公司
                sb.Append("           and  t3.CardCode like 'S%'");
            }

            if (cmbBU.Text != "")
            {
                switch (cmbBU.SelectedValue.ToString())
                {
                    case "ALL":
                        break;

                    case "ADP+AUO全部":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");
                        break;

                    case "AUO全部":
                    case "ADP全部":
                        sb.Append(" and T0.CARDCODE like '%" + cmbBU.SelectedValue.ToString() + "%'  ");
                        break;

                    case "Non AUO/ADP":
                    case "NON AUO/ADP":
                        sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5) NOT IN ('S0001','S0623') ");
                        break;

                }

            }
            /*
            if (txbDocDate.Text != "")
            {
                sb.Append("            and Convert(varchar(10),T1.DocDate,112) = '" + txbDocDate.Text + "'");
            }
            */


            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "AccApInvoice");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void CheckInvoiceTrack(ref System.Data.DataTable dt) 
        {
            DateTime DocDate = new DateTime();
            string error = "";
            if (txbDocDate.Text != "" && txbDocDate.Text.Length == 8)
            {
                int year = Convert.ToInt32(txbDocDate.Text.Substring(0, 4));
                int month = Convert.ToInt32(txbDocDate.Text.Substring(4, 2));
                int date = Convert.ToInt32(txbDocDate.Text.Substring(6, 2));

                DocDate = new DateTime(year, month, date);
            }
            else 
            {
                MessageBox.Show("過帳日期應為8碼");
                return;
            }

            foreach (DataRow row in dt.Rows) 
            {
                string DocEntry = row["OPDNDocEntry"].ToString();
                string U_Acme_Inv = row["U_PC_BSINV"].ToString();
                string U_Track = U_Acme_Inv != "" ? U_Acme_Inv.Substring(0, 2):""; 
                string Time = row["shipdate"].ToString().Substring(0, 6);
                string U_Year = Time.Substring(0, 4);
                string U_Period = GetPeriod(Time.Substring(4, 2)) ;
                string InvoiceType = row["InvoiceType"].ToString();//憑證類別
                if (row["OriCurrencyAmount"].ToString() == "") 
                {
                    //填入原幣金額
                    row.BeginEdit();
                    row["OriCurrencyAmount"] = (Convert.ToDouble(row["DocTotalSy"]) / Convert.ToDouble(row["U_ACME_RATE1"])).ToString("N2");
                    row.EndEdit();
                }
                if (Convert.ToString(row["TaxIdNumber"])== "84149738" || Convert.ToString(row["TaxIdNumber"]) == "16130599" || Convert.ToString(row["TaxIdNumber"]) == "")
                {
                 //友達有兩種統編要用發票號碼判斷是否為正確統編 
                    
                    string Track = "";
                    string Number = "";
                   
                    if (U_Acme_Inv.Length == 10 && U_Acme_Inv != "__________") 
                    {
                        row.BeginEdit();
                        Track = U_Acme_Inv.Substring(0, 2);
                        Number = U_Acme_Inv.Substring(2, 8);
                        
                        
                        row["TaxIdNumber"] = GetInvoiceTrack(Track, Number, Time).Rows[0]["TaxIdNum"].ToString(); 
                        row.EndEdit();
                    }
                }
               
                if (Convert.ToDecimal(row["VatSumSy"]) == 0) 
                {
                    //憑證類別當稅額等於0時捉『免用統一發票/收據』，其餘捉業夥伴主檔的憑證類別。
                    row.BeginEdit();
                    row["InvoiceType"] = "4";

                    row["TaxIdNumber"] = "";
                    row.EndEdit();
                }
                if (row["u_acme_shipday"].ToString() == "") 
                {
                    row.BeginEdit();

                    row["u_acme_shipday"] = row["shipdate"];

                    row.EndEdit();
                }
                if (U_Acme_Inv != "__________" && U_Acme_Inv !="" &&  CheckYearTrack(U_Track, U_Year, U_Period, InvoiceType) == false )
                {
                    error += DocEntry + ",";
                }
                if (txbDocDate.Text != "") 
                {
                    row.BeginEdit();
                    if (Convert.ToDateTime(row["DocDate"]).Month != DocDate.Month) 
                    {
                        //AP過帳月份 若與收採月份不同，要再備註打上原因才能產生
                        //不要備註了 AP過帳月分與採收月份相異 手動去sap上傳 他們說的
                        //row["Comments"] = "AP過帳月分與採收月份相異";

                    }
                    row["DocDate"] = DocDate;//過帳日期為設定收採明細匯入AP發票之過帳日期

                    row.EndEdit();

                }
                if (row["Signoff"] == "") 
                {
                    System.Data.DataTable dtSignoff = GetSignoff(row["por1Docentry"].ToString());
                    if (dtSignoff.Rows.Count > 0)
                    {
                        row.BeginEdit();
                        if (dtSignoff.Rows.Count > 1)
                        {
                            string Signoff = "";
                            foreach (DataRow rw in dtSignoff.Rows) 
                            {
                                Signoff += rw["SIGNOFF"].ToString() + "/";
                            }
                            Signoff = Signoff.Trim('/');
                            row["SIGNOFF"] = Signoff;
                        }
                        else 
                        {
                            string Signoff = dtSignoff.Rows[0]["SIGNOFF"].ToString();
                            row["SIGNOFF"] = Signoff;

                        }

                    }
                }
                

            }
            if (error != "") 
            {
                MessageBox.Show(error.TrimEnd(',') + "與發票字軌相異");
            }

        }
        
            private string GetPeriod(string Month) 
        {
            string U_Perid = "";
            switch (Month) 
            {
                case "01":
                case "02":
                    U_Perid = "1-2";
                    break;
                case "03":
                case "04":
                    U_Perid = "3-4";
                    break;
                case "05":
                case "06":
                    U_Perid = "5-6";
                    break;
                case "07":
                case "08":
                    U_Perid = "7-8";
                    break;
                case "09":
                case "10":
                    U_Perid = "9-10";
                    break;
                case "11":
                case "12":
                    U_Perid = "11-12";
                    break;
            }
            return U_Perid;
        }
        private System.Data.DataTable CombineDataTable(System.Data.DataTable dt) 
        {
            System.Data.DataTable dtData = new System.Data.DataTable();

            foreach (DataRow row in dt.Rows) 
            {

            }




            return dtData;

        }
        private System.Data.DataTable GetPDN1(string Docentry)
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT * FROM PDN1");
            sb.Append("            WHERE Docentry = @Docentry  ");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PDN1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOPDN(string Docentry)
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT * FROM OPDN");
            sb.Append("            WHERE Docentry = @Docentry  ");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PDN1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetInvoiceTrack(string Track,string Number,string Time)
        {
            string FD = "";
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT * FROM InvoiceTrack");
            sb.Append("            WHERE Track = @Track and (Numstart < @Number and Numend > @Number) and Time = @Time ");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Track", Track));
            command.Parameters.Add(new SqlParameter("@Number", Number));
            command.Parameters.Add(new SqlParameter("@Time", Time));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PDN1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private bool CheckYearTrack(string U_Track,string U_Year,string U_Period,string InvoiceType) 
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT Code, U_Year, U_Period, U_Track, U_InvType FROM [@CADMEN_YEARTRACK] ");
            sb.Append("  WHERE (U_Year = @U_Year ) AND (U_Period = @U_Period ) and (U_Track = @U_Track) ");
            if (InvoiceType == "0" || InvoiceType == "1")
            {
                sb.Append("  AND (U_InvType = '0') ");
            }
            if (InvoiceType == "2" )
            {
                sb.Append("  AND (U_InvType IN ('1','2')) ");
            }
            else if (InvoiceType == "8") 
            {
                sb.Append("  AND (U_InvType = '3') ");
            }
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_Track", U_Track));
            command.Parameters.Add(new SqlParameter("@U_Period", U_Period));
            command.Parameters.Add(new SqlParameter("@U_Year", U_Year));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PDN1");
            }
            finally
            {
                connection.Close();
            }
            if (ds.Tables[0].Rows.Count > 0) 
            {
                return true;
            }
            return false;
        }
        private System.Data.DataTable GetMaxOPCH()
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  MAX(DOCENTRY) DocEntry FROM OPCH");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPCH");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSignoff(string Docentry)
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            if (Docentry.Contains("/"))
            {
                string[] DocNum = Docentry.Split('/');
                sb.Append(" (");
                for (int i = 0; i < DocNum.Length; i++)
                {

                    if (i == DocNum.Length - 1)
                    {
                        sb.Append(" SELECT  distinct(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE  ,U1.SAPDOC, ");
                        sb.Append(" (USERNAME+' '+(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))+' '+ UPDATE_TIME) SIGNOFF  ");
                        sb.Append(" FROM  AcmeSqlEEP.dbo.SYS_TODOHIS U0   ");
                        sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 U1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=U1.ID)   ");
                        sb.Append(" where ISNULL(U1.SAPDOC,'') <> '' AND U0.flow_desc = ('採購單簽核流程(TFT)' ) AND U0.STATUS NOT IN ('NR','NF')AND  ( U1.SAPDOC like '"+ DocNum [i]+ "%'  ) AND   ");
                        sb.Append(" (CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112)) =  (SELECT  MAX(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE FROM  AcmeSqlEEP.dbo.SYS_TODOHIS V0  ");
                        sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 V1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=V1.ID) ");
                        sb.Append(" where ISNULL(V1.SAPDOC,'') <> '' AND V0.flow_desc = ('採購單簽核流程(TFT)' ) AND V0.STATUS NOT IN ('NR','NF') AND V1.SAPDOC =U1.SAPDOC ) ");
                    }
                    else
                    {
                        sb.Append(" SELECT  distinct(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE  ,U1.SAPDOC, ");
                        sb.Append(" (USERNAME+' '+(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))+' '+ UPDATE_TIME) SIGNOFF  ");
                        sb.Append(" FROM  AcmeSqlEEP.dbo.SYS_TODOHIS U0   ");
                        sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 U1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=U1.ID)   ");
                        sb.Append(" where ISNULL(U1.SAPDOC,'') <> '' AND U0.flow_desc = ('採購單簽核流程(TFT)' ) AND U0.STATUS NOT IN ('NR','NF')AND  ( U1.SAPDOC like '" + DocNum[i] + "%'  ) AND   ");
                        sb.Append(" (CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112)) =  (SELECT  MAX(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE FROM  AcmeSqlEEP.dbo.SYS_TODOHIS V0  ");
                        sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 V1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=V1.ID) ");
                        sb.Append(" where ISNULL(V1.SAPDOC,'') <> '' AND V0.flow_desc = ('採購單簽核流程(TFT)' ) AND V0.STATUS NOT IN ('NR','NF') AND V1.SAPDOC =U1.SAPDOC ) ");
                        sb.Append(" union ");
                    }

                }
                sb.Append(" )");
            }
            else
            {
                sb.Append(" SELECT  distinct(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE  ,U1.SAPDOC, ");
                sb.Append(" (USERNAME+' '+(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))+' '+ UPDATE_TIME) SIGNOFF  ");
                sb.Append(" FROM  AcmeSqlEEP.dbo.SYS_TODOHIS U0   ");
                sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 U1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=U1.ID)   ");
                sb.Append(" where ISNULL(U1.SAPDOC,'') <> '' AND U0.flow_desc = ('採購單簽核流程(TFT)' ) AND U0.STATUS NOT IN ('NR','NF')AND  ( U1.SAPDOC like '" + Docentry + "%'  ) AND   ");
                sb.Append(" (CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112)) =  (SELECT  MAX(CONVERT(varchar(12),CAST(UPDATE_DATE AS DATETIME),112))UPDATE_DATE FROM  AcmeSqlEEP.dbo.SYS_TODOHIS V0  ");
                sb.Append(" LEFT JOIN  AcmeSqlEEP.dbo.ACME_POR1 V1 ON (REPLACE(substring(form_presentation,6,DataLength(form_presentation)-6),'''','')=V1.ID) ");
                sb.Append(" where ISNULL(V1.SAPDOC,'') <> '' AND V0.flow_desc = ('採購單簽核流程(TFT)' ) AND V0.STATUS NOT IN ('NR','NF') AND V1.SAPDOC =U1.SAPDOC ) ");
            }
          
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Signoff");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void SetDateTimeTextBox() 
        {
            DateTime Now = DateTime.Now;
            string Days = DateTime.DaysInMonth(Now.Year, Now.Month).ToString();
            string Month = Now.Month.ToString().Length == 1 ? "0" + Now.Month : Now.Month.ToString();
            string Year = Now.Year.ToString();
            txbShipDateEnd.Text = Year + Month + Days;
            txbShipDateStart.Text = Year + Month + "01";
            txbDocDate.Text = Now.ToString("yyyyMMdd");
            txbAPDate.Text = Now.ToString("yyyyMMdd");
        }
        private void SetdgvInvoiceTrack() 
        {
           
            System.Data.DataTable dt = GetInvoiceTrack();
            dgvInvoiceTrack.DataSource = dt;
        }
       
        private System.Data.DataTable GetInvoiceTrack()
        {
            DateTime LastYear = new DateTime(DateTime.Now.AddYears(-1).Year, 1,1);
            DateTime ThisYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime NextYear = new DateTime(DateTime.Now.AddYears(1).Year,1,1);
            string thisyear = ThisYear.ToString("yyyyMM");
            string lastyear = LastYear.ToString("yyyyMM");
            string nextyear = NextYear.ToString("yyyyMM");
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT * FROM InvoiceTrack ");
            sb.Append("  WHERE Time >= @ThisYear and Time < @NextYear ");
            sb.Append("  UNION ");
            sb.Append("  SELECT * FROM InvoiceTrack ");
            sb.Append("  WHERE Time >= @LastYear and Time < @ThisYear ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@LastYear", lastyear));
            command.Parameters.Add(new SqlParameter("@ThisYear", thisyear));
            command.Parameters.Add(new SqlParameter("@NextYear", nextyear));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "AccApInvoice");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void btnCustNumber_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();
            if (LookupValues != null)
            {
                if (txbCardCode.Text == "") 
                {
                    txbCardCode.Text = Convert.ToString(LookupValues[1]); 
                }
                else
                {
                    txbCardCode.Text += "、" + Convert.ToString(LookupValues[1]);
                }
                    
            }
            txbCardCode.Text = txbCardCode.Text.TrimEnd('、');
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (dgvAccApInvoice.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            if (dgvAccApInvoice.GetCellCount(DataGridViewElementStates.Selected) == 0) 
            {
                MessageBox.Show("未選取上傳的行");
                return;
            }
            if (globals.UserID == "nesschou")
            {
                MessageBox.Show("確認是否為測試區");
            }
            foreach (DataGridViewRow row in dgvAccApInvoice.Rows)
            {
                if (Convert.ToBoolean(row.Cells["ColCheck"].Value) == false) 
                {
                    continue;
                }
                /*AP過帳月分與採收月份相異 手動去sap上傳 他們說的
                if (row.Cells["Comments"].Value.ToString() == "AP過帳月分與採收月份相異")
                {
                    DialogResult DialogComments = MessageBox.Show("AP過帳月分與採收月份相異,請填入備註,目前為預設是否繼續？", "提示", MessageBoxButtons.YesNo);
                    if (DialogComments == DialogResult.Yes) continue;
                    
                    return;
                }*/
            }
            DialogResult Dialog = MessageBox.Show("當前環境為" + globals.DBNAME + "是否繼續？","提示",MessageBoxButtons.YesNo);
            if (Dialog == DialogResult.Yes)  
            {
                try
                {
                    SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                    oCompany = new SAPbobsCOM.Company();

                    oCompany.Server = "acmesap";
                    oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                    oCompany.UseTrusted = false;
                    oCompany.DbUserName = "sapdbo";
                    oCompany.DbPassword = "@rmas";
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                    int i = 0; //  to be used as an indexfmship

                    oCompany.CompanyDB = FA;
                    oCompany.UserName = "A02";
                    oCompany.Password = "6500";
                    int result = oCompany.Connect();
                    if (result == 0)
                    {
                        System.Data.DataTable dt = dgvAccApInvoice.DataSource as System.Data.DataTable;

                        foreach (DataGridViewRow row in dgvAccApInvoice.Rows)
                        {
                            if (Convert.ToBoolean(row.Cells["ColCheck"].Value) == false)
                            {
                                continue;
                            }
                            if (Convert.ToInt32(row.Cells["DocTotalSy"].Value) == 0) 
                            {
                                UpdateOpdnStatus(Convert.ToString(row.Cells["OPDNDocEntry"].Value) ,"C");//O開啟,C關閉
                                continue;
                            }

                            SAPbobsCOM.Documents oPURCHINV = null;
                            oPURCHINV = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

                            string OPDNDocentry = row.Cells["OPDNDocentry"].Value.ToString();
                            System.Data.DataTable dtOPDN = GetOPDN(OPDNDocentry);


                            oPURCHINV.CardCode = row.Cells["CardCode"].Value.ToString();
                            oPURCHINV.CardName = row.Cells["CardName"].Value.ToString();
                            oPURCHINV.DocDate = Convert.ToDateTime(row.Cells["DocDate"].Value);
                            oPURCHINV.DocTotal = Convert.ToDouble(row.Cells["DocTotalSy"].Value);
                            oPURCHINV.TaxDate = Convert.ToDateTime(row.Cells["DocDate"].Value);
                            oPURCHINV.DiscountPercent = 0;
                            //oPURCHINV.Comments = row.Cells["Comments"].Value.ToString();
                            


                            //下面這些在sap要用add on才看的到
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSNOT").Value = row.Cells["TaxIdNumber"].Value;
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAMN").Value = Convert.ToDouble(row.Cells["UnTax"].Value);//未稅金額
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSTAX").Value = Convert.ToDouble(row.Cells["VatSumSy"].Value);//稅額
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAMT").Value = Convert.ToDouble(row.Cells["DocTotalSy"].Value) ;//含稅總額

                        

                            int Year = Convert.ToInt32((row.Cells["shipdate"].Value).ToString().Substring(0, 4));
                            int Month = Convert.ToInt32((row.Cells["shipdate"].Value).ToString().Substring(4, 2));
                            int Day = Convert.ToInt32((row.Cells["shipdate"].Value).ToString().Substring(6, 2));
                            DateTime InvoiceTime = new DateTime(Year,Month,Day);
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSDAT").Value = InvoiceTime;//發票日期
                            oPURCHINV.UserFields.Fields.Item("U_ACME_Invoice").Value = InvoiceTime;//發票日期

                            oPURCHINV.UserFields.Fields.Item("U_PC_BSTY1").Value = row.Cells["InvoiceType"].Value;//憑證類別
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSINV").Value = row.Cells["U_PC_BSINV"].Value;//發票號碼
                            
                            DateTime U_PC_BSAPP = new DateTime(InvoiceTime.AddMonths(1).Year, InvoiceTime.AddMonths(1).Month, 15);//次月15
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAPP").Value = U_PC_BSAPP;//申報年月
                            /*
                            if (row.Cells["DocDate"].Value.ToString() != row.Cells["por1DocDate"].Value.ToString()) 
                            {
                                //過帳日期相異
                                oPURCHINV.UserFields.Fields.Item("U_FOC").Value = row.Cells["Comments"].Value.ToString();
                            }*/

                            if (dtOPDN.Rows.Count > 0) 
                            {

                                oPURCHINV.Address = dtOPDN.Rows[0]["Address"].ToString();

                            } 

                            
                            System.Data.DataTable dtPDN1 = GetPDN1(OPDNDocentry);
                            int BaseLine = 0;
                            foreach (DataRow RowPDN1 in dtPDN1.Rows)
                            {
                                oPURCHINV.Lines.ItemCode = RowPDN1["ItemCode"].ToString();
                                oPURCHINV.Lines.ItemDescription = RowPDN1["Dscription"].ToString();
                                oPURCHINV.Lines.Quantity = Convert.ToDouble(RowPDN1["Quantity"]);
                                oPURCHINV.Lines.ShipDate = Convert.ToDateTime(RowPDN1["ShipDate"]);
                                oPURCHINV.Lines.Price = Convert.ToDouble(RowPDN1["Price"]);
                                oPURCHINV.Lines.Currency = RowPDN1["Currency"].ToString();

                                oPURCHINV.Lines.WarehouseCode = RowPDN1["WhsCode"].ToString();
                                oPURCHINV.Lines.Address = RowPDN1["Address"].ToString();
                                oPURCHINV.Lines.ShipToDescription = RowPDN1["ShipToDesc"].ToString();
                                oPURCHINV.Lines.BaseEntry = Convert.ToInt32(RowPDN1["Docentry"]);
                                oPURCHINV.Lines.BaseLine = BaseLine;
                                oPURCHINV.Lines.BaseType = 20;
                                oPURCHINV.Lines.DiscountPercent = 0;

                                BaseLine += 1;

                                oPURCHINV.Lines.Add();
                            }
                            int res = oPURCHINV.Add();
                            if (res != 0)
                            {
                                string error = oCompany.GetLastErrorDescription();
                                MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                                return;
                            }
                            else
                            {
                                System.Data.DataTable G4 = GetMaxOPCH();
                                string OWTR = G4.Rows[0][0].ToString();
                                //MessageBox.Show("上傳成功  AP發票單號 : " + OWTR);
                                string TaxIdNumber = row.Cells["TaxIdNumber"].Value.ToString();
                                string InvoiceType = row.Cells["InvoiceType"].Value.ToString();
                                if (InvoiceType != "4")
                                {
                                    UpdateOPCH(OWTR, TaxIdNumber, InvoiceType, U_PC_BSAPP);
                                }



                            }


                        }
                    }
                }
                catch (Exception ex) 
                {
                    MessageBox.Show("上傳資料有問題");
                    return;
                }
               

                
            }
            MessageBox.Show("上傳完成");
           

        }

        private void btnCancelCardCode_Click(object sender, EventArgs e)
        {
            string[] CardCode = txbCardCode.Text.Split('、');
            int count = txbCardCode.Text.IndexOf('、', CardCode.Length - 1);//最後一個出現頓號的位置

            txbCardCode.Text = txbCardCode.Text.Substring(0, count);

            

        }
        private void UpdateOPCH(string DocEntry,string TaxIdNumber,string InvoiceType,DateTime U_PC_BSAPP) 
        {

            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE OPCH SET LicTradNum = @TaxIdNumber,U_PC_BSTY1 = @InvoiceType,U_PC_BSAPP = @U_PC_BSAPP WHERE Docentry = @DocEntry ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@TaxIdNumber", TaxIdNumber));
            command.Parameters.Add(new SqlParameter("@InvoiceType", InvoiceType));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAPP", U_PC_BSAPP));
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
        private void UpdateOpdnStatus(string DocEntry,string DOCSTATUS)
        {

            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE OPDN SET DOCSTATUS = @DOCSTATUS WHERE Docentry = @DocEntry ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry)); 
            command.Parameters.Add(new SqlParameter("@DOCSTATUS", DOCSTATUS));


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


        private void btnExcel_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            if (this.tabControl1.SelectedTab.Text == "收採發票")
            {
                GetExcelDataTabledgvAccApInvoice(ref dt);
            }
            else if (this.tabControl1.SelectedTab.Text == "AP發票明細") 
            {
                GetExcelDataTabledgvApinvoice(ref dt);
            }
          
            string location = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Excel\\temp\\AP發票.xls";
            if (dt.Rows.Count > 0)
            {
                string Template = System.Environment.CurrentDirectory + "\\Excel\\" + "AP發票.xls";

                WriteDataTableToExcel(dt, Template, location);

                Process.Start(location);
            }
            else 
            {
                MessageBox.Show("請選擇行");
            }
            


        }
        private void GetExcelDataTabledgvAccApInvoice(ref System.Data.DataTable dt) 
        {
            for (int count = 0; count < dgvAccApInvoice.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgvAccApInvoice.Columns[count].Name);
                dt.Columns.Add(dc);
            }

            int x = 0;
            foreach (DataGridViewRow rows in dgvAccApInvoice.Rows) 
            {
                if (Convert.ToBoolean(rows.Cells["ColCheck"].EditedFormattedValue) == false) continue;

                // 循環行

                DataRow dr = dt.NewRow();
                for (int count = 0; count < dgvAccApInvoice.Columns.Count; count++)
                {
                    dr[count] = Convert.ToString(dgvAccApInvoice.Rows[x].Cells[count].FormattedValue);
                }
                dt.Rows.Add(dr);
                x++;
            }
        }
        private void GetExcelDataTabledgvApinvoice(ref System.Data.DataTable dt)
        {
            for (int count = 0; count < dgvApinvoice.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgvApinvoice.Columns[count].Name);
                dt.Columns.Add(dc);
            }

            int x = 0;
            foreach (DataGridViewRow rows in dgvApinvoice.Rows)
            {

                // 循環行

                DataRow dr = dt.NewRow();
                for (int count = 0; count < dgvApinvoice.Columns.Count; count++)
                {
                    dr[count] = Convert.ToString(dgvApinvoice.Rows[x].Cells[count].FormattedValue);
                }
                dt.Rows.Add(dr);
                x++;
            }
        }
        public bool WriteDataTableToExcel
(System.Data.DataTable dataTable, string Template,string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            // Microsoft.Office.Interop.Excel.Range excelCellrange;
            object oMissing = System.Reflection.Missing.Value;


            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            excel.DisplayAlerts = false;

            try
            {


                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Sheets.get_Item(1);

                // Workk sheet
                SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                excelSheet.Name = "1.明細";
                WriteDataTableToSheetByArray(dataTable, excelSheet);
                //now save the workbook and exit Excel
                //excelworkBook.SaveAs(saveAsLocation);
                SheetTemplate.Delete();

                Microsoft.Office.Interop.Excel.Worksheet sheet = excelworkBook.Application.Sheets["2.發票字軌"] as Microsoft.Office.Interop.Excel.Worksheet;
                sheet.Move(Type.Missing, excelworkBook.Application.Sheets[2]);
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
        private void WriteDataTableToSheetByArray(System.Data.DataTable dataTable,
            Worksheet worksheet)
        {
            try
            {
                int rows = dataTable.Rows.Count + 1;
                int columns = dataTable.Columns.Count  ;


                var data = new object[rows, columns];

                int rowcount = 0;

                for (int i = 1; i <= columns; i++)
                {

                    //表頭欄位
                    if (i == 1) data[rowcount, i - 1] = "No.";
                    else
                    {
                        data[rowcount, i - 1] = this.dgvAccApInvoice.Columns[i - 1].HeaderText;
                    }


                }

                rowcount += 1;
                foreach (DataRow datarow in dataTable.Rows)
                {
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // Filling the excel file 
                        if (i == 1) data[rowcount, i - 1] = rowcount.ToString();
                        else
                        {
                            data[rowcount, i - 1] = datarow[i - 1].ToString();
                        }

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
            catch (Exception ex) 
            {
            }
           
        }
        private System.Data.DataTable MakeDtAccApInvoice() 
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("No.", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));
            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["格式"];
            //dt.PrimaryKey = colPk;
            dt.TableName = "AccApInvoice";

            return dt;
        }

        private void dgvAccApInvoice_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            
            if (this.dgvAccApInvoice.CurrentCell.ColumnIndex == 0 && e.RowIndex != -1) 
            {
                DataGridViewCheckBoxCell dgvCheck = (DataGridViewCheckBoxCell)(this.dgvAccApInvoice.Rows[this.dgvAccApInvoice.CurrentCell.RowIndex].Cells[0]);
                
                if (Convert.ToBoolean(dgvCheck.EditedFormattedValue) == false)
                {
                    decimal x = txbOriCurrencyAmount.Text == "" ? 0 : Convert.ToDecimal(txbOriCurrencyAmount.Text);
                    decimal y = dgvAccApInvoice.Rows[e.RowIndex].Cells["OriCurrencyAmount"].Value == null ? 0 : Convert.ToDecimal(dgvAccApInvoice.Rows[e.RowIndex].Cells["OriCurrencyAmount"].Value);

                    txbOriCurrencyAmount.Text = (x + y).ToString("N2");

                    decimal a = txbAccountAmount.Text == "" ? 0 : Convert.ToDecimal(txbAccountAmount.Text);
                    decimal b = dgvAccApInvoice.Rows[e.RowIndex].Cells["DocTotalSy"].Value == null ? 0 : Convert.ToDecimal(dgvAccApInvoice.Rows[e.RowIndex].Cells["DocTotalSy"].Value);

                    txbAccountAmount.Text = (a + b).ToString("N2");

                    dgvCheck.Value = true;
                    dgvAccApInvoice.Rows[e.RowIndex].Cells["ColCheck"].Value = true;
                }
                else
                {
                    //若CheckBox已经被勾上

                    decimal x = txbOriCurrencyAmount.Text == "" ? 0 : Convert.ToDecimal(txbOriCurrencyAmount.Text);
                    decimal y = dgvAccApInvoice.Rows[e.RowIndex].Cells["OriCurrencyAmount"].Value == null ? 0 : Convert.ToDecimal(dgvAccApInvoice.Rows[e.RowIndex].Cells["OriCurrencyAmount"].Value);

                    txbOriCurrencyAmount.Text = (x - y).ToString("N2");

                    decimal a = txbAccountAmount.Text == "" ? 0 : Convert.ToDecimal(txbAccountAmount.Text);
                    decimal b = dgvAccApInvoice.Rows[e.RowIndex].Cells["DocTotalSy"].Value == null ? 0 : Convert.ToDecimal(dgvAccApInvoice.Rows[e.RowIndex].Cells["DocTotalSy"].Value);

                    txbAccountAmount.Text = (a - b).ToString("N2");

                    dgvCheck.Value = false;
                    dgvAccApInvoice.Rows[e.RowIndex].Cells["ColCheck"].Value = false;
                }

                dgvAccApInvoice.EndEdit();
            }
            //CountOriSum();
        }

        private void dgvAccApInvoice_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
            dgvApinvoice.Rows[dgvApinvoice.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Yellow;
        }
        private void btnCheckCheckBox_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button button = sender as System.Windows.Forms.Button;
            if (button.Text == "全部")
            {
                foreach (DataGridViewRow row in dgvAccApInvoice.Rows)
                {
                    row.Cells["ColCheck"].Value = true;
                }

            }
            else if (button.Text == "部分")
            {
                int i = 0;
                foreach (DataGridViewRow row in dgvAccApInvoice.SelectedRows) 
                {
                    row.Cells["ColCheck"].Value = true;
                    i++;
                }
                if (i == 0) 
                {
                    MessageBox.Show("請先反白選取行");
                }
            }
            CountOriSum();

        }
        private void CountOriSum() 
        {
            decimal Sum = 0;
            decimal Count = 0;
            foreach (DataGridViewRow row in dgvAccApInvoice.Rows) 
            {
                if (Convert.ToBoolean(row.Cells["ColCheck"].Value)  == true)
                {
                    Sum += row.Cells["OriCurrencyAmount"].Value == null ? 0 : Convert.ToDecimal(row.Cells["OriCurrencyAmount"].Value);
                    Count += row.Cells["DocTotalSy"].Value == null ? 0 : Convert.ToDecimal(row.Cells["DocTotalSy"].Value);
                }
            }
            txbOriCurrencyAmount.Text = Sum.ToString("N2");

            txbAccountAmount.Text = Count.ToString("N2");


        }

        private void dgvAccApInvoice_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            //顯示HeaderCell
            for (int i = 0; i < dgvAccApInvoice.Rows.Count; i++) 
            {
                DataGridViewRow r = this.dgvAccApInvoice.Rows[i];
                r.HeaderCell.Value = string.Format("{0}", i + 1);

            }
            this.dgvAccApInvoice.Refresh();
        }

        private void dgvAccApInvoice_RowPostPaint_1(object sender, DataGridViewRowPostPaintEventArgs e)
        {

        }

        private void btnImportInvoiceTrack_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {
                string file = dialog.FileName;
                if (file == null) { MessageBox.Show("請選擇檔案"); return; } 

                WriteExcelToInvoiceTrack(file);
                MessageBox.Show("上傳完成,請重按查詢以更新資料");

            }
        }
        private void WriteExcelToInvoiceTrack(string FileName) 
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            object oMissing = System.Reflection.Missing.Value;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCount = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCount = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            try
            {
                string TaxIdNum = "";
                string Track = "";
                string NumStart = "";
                string NumEnd = "";
                string Time = "";

                for (int i = 1; i <= iRowCount; i++) 
                {
                   //EXCEL要從1開始計數
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    TaxIdNum = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    Track = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    NumStart = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    range.Select();
                    NumEnd = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    Time = range.Text.ToString().Trim();

                    if (CheckInvoiceTrack(Time, TaxIdNum) == true) 
                    {
                        MessageBox.Show("已有資料");
                        return;

                    }
                    InsertInvoiceTrack(TaxIdNum, Track, NumStart, NumEnd, Time);
                }

            }
            catch (Exception ex) 
            {
                
            }
            finally 
            { 
                //Quit
                excelApp.Quit();
            }
        }
        private bool CheckInvoiceTrack(string Time,string TaxIdNum) 
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("SELECT * FROM InvoiceTrack WHERE Time = @Time and TaxIdNum = @TaxIdNum", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Time", Time));
            command.Parameters.Add(new SqlParameter("@TaxIdNum", TaxIdNum));
            command.CommandType = CommandType.Text;

            //command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "InvoiceTrack");
            }
            finally
            {
                connection.Close();
            }
            if (ds.Tables[0].Rows.Count > 0) 
            {

                return true;
            }
            return false;
        }
        private void InsertInvoiceTrack(string TaxIdNum,string Track,string NumStart, string NumEnd,string Time) 
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into InvoiceTrack(TaxIdNum,Track,NumStart,NumEnd,Time) VALUES(@TaxIdNum,@Track,@NumStart,@NumEnd,@Time)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TaxIdNum", TaxIdNum));
            command.Parameters.Add(new SqlParameter("@Track", Track));
            command.Parameters.Add(new SqlParameter("@NumStart", NumStart));
            command.Parameters.Add(new SqlParameter("@NumEnd", NumEnd));
            command.Parameters.Add(new SqlParameter("@Time", Time));

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

        private void btnOpenExample_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(System.Environment.CurrentDirectory + "\\Excel\\" + "InvoiceTrack.xlsx");
        }
    }
}

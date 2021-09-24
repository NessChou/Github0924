using System;
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
            SetDateTimeTextBox();//起訖設當月第一天跟最後一天
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetAccAPInvoice();
            CheckInvoiceTrack(ref dt);//友達要依據發票號碼填入對應統編
            System.Data.DataTable dtData = CombineDataTable(dt);//合併por1
            dgvAccApInvoice.DataSource = dt;
        }
        private System.Data.DataTable GetAccAPInvoice()
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT DISTINCT( t0.docdate ),t0.docentry OPDNDocEntry ,t2.BaseEntry pdn1BaseEntry,");
            sb.Append("             STUFF((SELECT '/'+ CAST (T22.DOCENTRY AS NVARCHAR) FROM POR1 T22 LEFT JOIN pdn1 t11 on t11.baseentry = T22.docentry and  t11.baseline = T22.linenum  AND t11.BASETYPE = 22  ");
            sb.Append("             WHERE(t11.docentry = T1.DOCENTRY)  GROUP BY T22.DocEntry FOR XML PATH('')),1,1,'') AS por1Docentry,");
            sb.Append("             t3.Docentry por1Docentry ,t3.cardcode ,t3.cardname ,  ");
            sb.Append("             (select sum(t1.quantity) from pdn1 t1 where t0.docentry = t1.DocEntry) Quantity,  ");
            sb.Append("             (cast(t0.doctotalsy as int) -cast(t0.VatSumSy as int)) UnTax, t0.VatSumSy ,t0.DocTotalSy ,t0.U_acme_inv,t2.shipdate,t0.U_PC_BSINV,t1.u_acme_shipday,t2.currency,t0.U_ACME_RATE1 ,'' OriCurrencyAmount,t0.u_acme_lc,t4.LicTradNum TaxIdNumber,T4.U_IN_BSTY1 InvoiceType ,t0.U_ACME_Invoice  ");
            sb.Append("             FROM opdn t0  ");
            sb.Append("             LEFT JOIN pdn1 t1 on t0.DocEntry =t1.docentry ");
            sb.Append("             LEFT JOIN por1 t2  on (t1.baseentry=T2.docentry and  t1.baseline=t2.linenum  AND t1.BASETYPE=22)  ");
            sb.Append("             LEFT JOIN opor t3  on t2.docentry = t3.docentry");
            sb.Append("             LEFT JOIN OCRD t4  on t3.CARDCODE = t4.cardcode");
            sb.Append("             LEFT JOIN OITM T11 ON t2.ITEMCODE = T11.ITEMCODE  ");

            sb.Append("             WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' and t0.DocStatus <> 'C' ");
            if (txbShipDateStart.Text != "" || txbShipDateEnd.Text != "")
            {
                sb.Append("            and Convert(varchar(10),t0.DocDate,112) between '" +txbShipDateStart.Text + "' and '" + txbShipDateEnd.Text + "'");
            }
            
            if (txbDocDate.Text != "")
            {
                sb.Append("            and t0.DocDate = '" + txbDocDate.Text + "'");
            }
            if (txbCardCode.Text != "")
            {
                string[] cardcode = txbCardCode.Text.Split('、');
                for (int i = 0; i < cardcode.Length; i++) 
                {
                    if (i == 0)
                    {
                        sb.Append("           and  t3.cardname like '%" + txbShipDateStart.Text + "%'");
                    }
                    else 
                    {
                        sb.Append("             or t3.cardname like '%" + txbShipDateStart.Text + "%'");
                    }
                   
                }
                
            }
            if (cmbBU.Text != "")
            {
                if (cmbBU.SelectedValue.ToString() == "ADP+AUO全部")
                {
                    sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");

                }
                else
                {
                    sb.Append(" and T0.CARDCODE like '%" + cmbBU.SelectedValue.ToString() + "%'  ");
                }
            }
            if (txbOriCurrencyAmount.Text != "") 
            {
                sb.Append("            and t0.DocTotalSy = '" + txbDocDate.Text + "'");
            }

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
            foreach (DataRow row in dt.Rows) 
            {
                if (row["OriCurrencyAmount"].ToString() == "") 
                {
                    //填入原幣金額
                    row.BeginEdit();
                    row["OriCurrencyAmount"] = Convert.ToDouble(row["DocTotalSy"]) * Convert.ToDouble(row["U_ACME_RATE1"]);
                    row.EndEdit();
                }
                if (row["TaxIdNumber"].ToString() == "84149738" || row["TaxIdNumber"].ToString() == "16130599" || row["TaxIdNumber"].ToString() == "")
                {
                 //友達有兩種統編要用發票號碼判斷是否為正確統編 
                    string U_Acme_Inv = row["U_acme_inv"].ToString();
                    string Track = "";
                    string Number = "";
                    if (U_Acme_Inv.Length == 8) 
                    {
                        row.BeginEdit();
                        Track = U_Acme_Inv.Substring(0, 2);
                        Number = U_Acme_Inv.Substring(2, 6);
                        
                        row["TaxIdNumber"] = GetInvoiceTrack(Number, Track).Rows[0]["TaxIdNumber"].ToString(); 
                        row.EndEdit();
                    }
                }
                if (row["VatSumSy"].ToString() == "0") 
                {
                    //憑證類別當稅額等於0時捉『免用統一發票/收據』，其餘捉業夥伴主檔的憑證類別。
                    row.BeginEdit();
                    row["InvoiceType"] = "4";

                    row.EndEdit();
                }
            }

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
        private System.Data.DataTable GetInvoiceTrack(string Track,string Number)
        {
            string FD = "";
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT * FROM InvoiceTrack");
            sb.Append("            WHERE Track = @Track and (Numstart < @Number and Numstart > @Number) ");
            //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Track", Track));
            command.Parameters.Add(new SqlParameter("@Number", Number));
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
        private void SetDateTimeTextBox() 
        {
            DateTime Now = DateTime.Now;
            string Days = DateTime.DaysInMonth(Now.Year, Now.Month).ToString();
            string Month = Now.Month.ToString().Length == 1 ? "0" + Now.Month : Now.Month.ToString();
            string Year = Now.Year.ToString();
            txbShipDateEnd.Text = Year + Month + Days;
            txbShipDateStart.Text = Year + Month + "01";
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
            if (globals.UserID == "nesschou")
            {
                MessageBox.Show("確認是否為測試區");
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

                    int i = 0; //  to be used as an index

                    oCompany.CompanyDB = FA;
                    oCompany.UserName = "A02";
                    oCompany.Password = "6500";
                    int result = oCompany.Connect();
                    if (result == 0)
                    {
                        System.Data.DataTable dt = dgvAccApInvoice.DataSource as System.Data.DataTable;
                        foreach (DataRow row in dt.Rows)
                        {
                            SAPbobsCOM.Documents oPURCHINV = null;
                            oPURCHINV = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

                            string OPDNDocentry = row["OPDNDocentry"].ToString();
                            System.Data.DataTable dtOPDN = GetOPDN(OPDNDocentry);


                            oPURCHINV.CardCode = row["CardCode"].ToString();
                            oPURCHINV.CardName = row["CardName"].ToString();
                            oPURCHINV.DocDate = Convert.ToDateTime(row["DocDate"]);
                            oPURCHINV.DocTotal = Convert.ToDouble(row["DocTotalSy"]);
                            oPURCHINV.TaxDate = Convert.ToDateTime(row["DocDate"]);


                            //下面這些在sap要用add on才看的到
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSNOT").Value = row["TaxIdNumber"];
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAMN").Value = Convert.ToDouble(row["UnTax"]);//未稅金額
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSTAX").Value = Convert.ToDouble(row["VatSumSy"]);//稅額
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAMT").Value = Convert.ToDouble(row["DocTotalSy"]) ;//含稅總額

                            oPURCHINV.UserFields.Fields.Item("U_PC_BSTY1").Value = row["InvoiceType"];//憑證類別
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSINV").Value = row["U_acme_inv"];//發票號碼
                            DateTime invoiceTime = Convert.ToDateTime(row["U_ACME_Invoice"]);
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSDAT").Value = invoiceTime;//發票日期
                            DateTime datetime = new DateTime(invoiceTime.Year, invoiceTime.AddMonths(1).Month, 15);//次月15
                            oPURCHINV.UserFields.Fields.Item("U_PC_BSAPP").Value = datetime;//申報年月



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


                               
                                BaseLine += 1;

                                oPURCHINV.Lines.Add();


                            }
                            int res = oPURCHINV.Add();
                            if (res != 0)
                            {
                                string error = oCompany.GetLastErrorDescription();
                                //MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                            }
                            else
                            {
                                System.Data.DataTable G4 = GetMaxOPCH();
                                string OWTR = G4.Rows[0][0].ToString();
                                //MessageBox.Show("上傳成功 採購報價單號 : " + OWTR);
                                string TaxIdNumber = row["TaxIdNumber"].ToString();
                                string InvoiceType = row["InvoiceType"].ToString();
                                if (InvoiceType != "4")
                                {
                                    UpdateOPCH(OWTR, TaxIdNumber, InvoiceType, datetime);
                                }



                            }


                        }
                    }
                }
                catch (Exception ex) 
                {

                }
               

                
            }
                
           

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
    }
}

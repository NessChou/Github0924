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
using System.Net.Mime;

namespace ACME
{
    public partial class APLCK : ACME.fmBase7
    {
        System.Net.Mail.Attachment data = null;
        public string dtd = "";
        public string PublicString;
        string STATUS;
        private System.Data.DataTable OrderData;
        public APLCK()
        {
            InitializeComponent();
        }

  
        private void CalcTotals()
        {
            //資料瀏覽時不計算
            if (MyTableStatus == "0" || String.IsNullOrEmpty(MyTableStatus.ToString()))
            {
                return;
            }


           decimal iTotal = 0;
           decimal iamt = 0;

            int i = this.pLC1DataGridView.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToDecimal(pLC1DataGridView.Rows[iRecs].Cells["Amt2"].Value);

            }
            iamt = Convert.ToDecimal(lcAmtTextBox.Text);
            decimal aa = iamt - iTotal;
            aa = Math.Round(aa, 2, MidpointRounding.AwayFromZero);
            lcTotalTextBox.Text = aa.ToString();
           


        }

        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                lC.APLC.RejectChanges();
                lC.APLC1.RejectChanges();
                lC.AP_Download.RejectChanges();
            }
            catch
            {
            }

            return true;

        }
        public override void AfterEndEdit()
        {
 
            try
            {
                CalcTotals();
                WW();
            }
            catch (Exception ex)
            {
              //  MessageBox.Show(ex.Message);
            }
            updateNameTextBox.Text = fmLogin.LoginID.ToString();
            aPLCBindingSource.EndEdit();
            aPLCTableAdapter.Update(lC.APLC);

           

        }
        public override void STOP()
        {
            if (lcNoTextBox.Text == "")
            {
                MessageBox.Show("LCNO必須輸入");
                this.SSTOPID = "1";
                lcNoTextBox.Focus();
                return ;
            }

            decimal iTotal = 0;
            decimal iamt = 0;

            int i = this.pLC1DataGridView.Rows.Count - 1;
            if (pLC1DataGridView.Rows.Count > 1)
            {
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    iTotal += Convert.ToDecimal(pLC1DataGridView.Rows[iRecs].Cells["Amt2"].Value);

                }
            }
            if (String.IsNullOrEmpty(lcAmtTextBox.Text))
            {
                iamt = 0;
            }
            else
            {
                iamt = Convert.ToDecimal(lcAmtTextBox.Text);
            }
            decimal aa = iamt - iTotal;
            aa = Math.Round(aa, 2, MidpointRounding.AwayFromZero);
            if (aa < 0)
            {
                MessageBox.Show("沖銷金額不能小於零");
                this.SSTOPID = "1";
                return ;

            }


            System.Data.DataTable dt1 = download2(lcNoTextBox.Text);
            System.Data.DataTable dt2 = download3(docNumTextBox.Text);
            if (dt1.Rows.Count > 0)
            {


                if (STATUS == "INSERT")
                {
                    MessageBox.Show("LCNO重複!!!");
                    this.SSTOPID = "1";
                    return ;

                }
                else
                {
                    if (dt2.Rows.Count > 0)
                    {
                        string LC = dt2.Rows[0][0].ToString();

                        if (LC != lcNoTextBox.Text)
                        {
                            MessageBox.Show("LCNO重複!!!");
                            this.SSTOPID = "1";
                            return ;

                        }


                    }
                }
            }
            try
            {
                if (bankCodeTextBox.Text.Trim() == "USD")
                {


                    decimal iTotals = 0;
                    decimal iamts = 0;

                    int iS = this.pLC1DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= iS; iRecs++)
                    {
                        iTotals += Convert.ToDecimal(pLC1DataGridView.Rows[iRecs].Cells["Amt2"].Value);

                    }
                    iamts = Convert.ToDecimal(lcAmtTextBox.Text);
                    decimal aas = iamts - iTotals;
                    aas = Math.Round(aas, 2, MidpointRounding.AwayFromZero);
                    decimal T1 = Convert.ToDecimal(aas);

                    if (T1 < 10000)
                    {
                        DialogResult result;
                        result = MessageBox.Show("請檢查出貨/押匯時間 確認完請按'是'  尚未確認請按'否' ", "YES/NO", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No)
                        {
                            this.SSTOPID = "1";

                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
     

        public override void AfterCancelEdit()
        {
            WW();
            pLC1DataGridView.Enabled = false;
        }
        public override void AfterEdit()
        {
            pLC1DataGridView.Enabled = true;
            bankCodeTextBox.ReadOnly = true;
            lcTotalTextBox.ReadOnly = true;
        }
 
        public override void EndEdit()
        {
            WW();
            STATUS = "";

        }
        private void WW()
        {
            button2.Enabled = true;
            button5.Enabled = true;
            button8.Enabled = true;
            button3.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;
            pLC1DataGridView.Enabled = true;

            lcTotalTextBox.ReadOnly = true;
            docNumTextBox.ReadOnly = true;
            bankCodeTextBox.ReadOnly = true;

            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
        }
    
        public override void query()
        {
            docNumTextBox.ReadOnly = false;


        }
        public override void AfterAddNew()
        {
            WW();
 
        }



        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            aPLCTableAdapter.Connection = MyConnection;
            pLC1TableAdapter.Connection = MyConnection;
            aP_DownloadTableAdapter.Connection = MyConnection;
        }
        public override void SetInit()
        {

            MyBS = aPLCBindingSource;
            MyTableName = "APLC";
            MyIDFieldName = "DocNum";

        
            UtilSimple.SetLookupBinding(bankNameComboBox, "BankName", aPLCBindingSource, "BankName");

           
        }
        public override void FillData()
        {
            try
            {

                if (!String.IsNullOrEmpty(PublicString))
                {
                    MyID = PublicString.Trim();
                }

                aPLCTableAdapter.Fill(lC.APLC,MyID);
                pLC1TableAdapter.Fill(lC.PLC1,MyID);
                aP_DownloadTableAdapter.Fill(lC.AP_Download, MyID);

              
                DataGridViewRow row;
                for (int i = pLC1DataGridView.Rows.Count - 2; i >= 0; i--)
                {

                    row = pLC1DataGridView.Rows[i];

                    string d3 = row.Cells[2].Value.ToString();
                    string d22 = row.Cells[21].Value.ToString();
                    System.Data.DataTable sa = Getpor1(d3, d22);
                    if (sa.Rows.Count > 0)
                    {
                        DataRow drw = sa.Rows[0];
                        if (drw["targettype"].ToString() == "20")
                        {
                            row.Cells[23].Value = "checked";
                        }
                    }
                }
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {

                Validate();

                aPLCTableAdapter.Connection.Open();


                aPLCBindingSource.EndEdit();
                pLC1BindingSource1.EndEdit();
                aP_DownloadBindingSource.EndEdit();

                tx = aPLCTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(aPLCTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(pLC1TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(aP_DownloadTableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;

                aPLCTableAdapter.Update(lC.APLC);
                lC.APLC.AcceptChanges();
                pLC1TableAdapter.Update(lC.PLC1);
                lC.PLC1.AcceptChanges();
                aP_DownloadTableAdapter.Update(lC.AP_Download);
                lC.AP_Download.AcceptChanges();

                tx.Commit();
                this.MyID = this.docNumTextBox.Text;

                UpdateData = true;
            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }

                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                aPLCTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        public override void SetDefaultValue()
        {
            STATUS = "INSERT";
      

            string NumberName = "AL" + DateTime.Now.ToString("yyyyMMdd");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.docNumTextBox.Text = NumberName + AutoNum;
            createNameTextBox.Text = fmLogin.LoginID.ToString();
            lcCloseCheckBox.Checked=false;
            fCHECKCheckBox.Checked = false;

            lCTYPETextBox.Text = "達擎";
            this.aPLCBindingSource.EndEdit();
        }


        private void pLC1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {

            e.Row.Cells["SendDate"].Value ="2008";
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (bankCodeTextBox.Text == "")
            {
                MessageBox.Show("請輸入幣別");
                
            }
            else
            {

                object[] LookupValues = GetMenu.GetMenuListS();

                if (LookupValues != null)
                {
                    cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                    cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
        }

       

        private void pLC1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (bankCodeTextBox.Text == "NTD")
                {

                    if (this.pLC1DataGridView.Rows[e.RowIndex].Cells["TaxCode"].Value.ToString() == "@")
                    {
                        if (pLC1DataGridView.Columns[e.ColumnIndex].Name == "Qty" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "Price" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "TaxCode" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "Tax" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "LcNo")
                        {

                            decimal iQuantity = 0;
                            decimal iUnitPrice = 0;
                            decimal itax = 0;
                            decimal LcNo = 0;
                            iQuantity = Convert.ToInt32(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Qty"].Value);
                            iUnitPrice = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value);
                            itax = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Tax"].Value);
                            LcNo = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["LcNo"].Value);
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Amt2"].Value = (iQuantity * iUnitPrice * LcNo + itax).ToString();

                        }
                    }
                    else
                    {
                        if (pLC1DataGridView.Columns[e.ColumnIndex].Name == "Qty" ||
                     pLC1DataGridView.Columns[e.ColumnIndex].Name == "TaxCode" ||
                     pLC1DataGridView.Columns[e.ColumnIndex].Name == "LcNo" ||
                            pLC1DataGridView.Columns[e.ColumnIndex].Name == "comments")
                        {

                            decimal iQuantity = 0;
                      
                            decimal orgUnitPrice = 0;
                            decimal itax = 0;
                            decimal LcNo = 0;
                            decimal itax2 = 0;
                            decimal itax3 = 0;
                            iQuantity = Convert.ToInt32(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Qty"].Value);
                            itax = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["TaxCode"].Value);
                            orgUnitPrice = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["comments"].Value);
                            itax2 = itax / 100 + 1;
                            itax3 = itax / 100;
                            LcNo = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["LcNo"].Value);
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value = (orgUnitPrice * LcNo).ToString();
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Tax"].Value = (iQuantity * orgUnitPrice * LcNo * itax3).ToString();
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Amt2"].Value = (iQuantity * orgUnitPrice * LcNo + iQuantity * orgUnitPrice * LcNo * itax3).ToString();


                        }
                    }
                }
                if (bankCodeTextBox.Text == "USD")
                {
                    if (this.pLC1DataGridView.Rows[e.RowIndex].Cells["TaxCode"].Value.ToString() == "@")
                    {
                        if (pLC1DataGridView.Columns[e.ColumnIndex].Name == "Qty" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "Price" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "TaxCode" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "Tax" ||
                           pLC1DataGridView.Columns[e.ColumnIndex].Name == "LcNo")
                        {

                            decimal iQuantity = 0;
                            decimal iUnitPrice = 0;
                            decimal itax = 0;

                            iQuantity = Convert.ToInt32(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Qty"].Value);
                            iUnitPrice = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value);
                            itax = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Tax"].Value);

                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Amt2"].Value = (iQuantity * iUnitPrice + itax).ToString();
                        }
                    }
                    else
                    {
                        if (pLC1DataGridView.Columns[e.ColumnIndex].Name == "Qty" ||
                          pLC1DataGridView.Columns[e.ColumnIndex].Name == "Price" ||
                          pLC1DataGridView.Columns[e.ColumnIndex].Name == "TaxCode" ||
                          pLC1DataGridView.Columns[e.ColumnIndex].Name == "LcNo")
                        {
                            decimal iQuantity = 0;
                            decimal iUnitPrice = 0;
                            decimal itax = 0;
                            decimal itax2 = 0;
                            decimal itax3 = 0;

                            iQuantity = Convert.ToInt32(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Qty"].Value);
                            iUnitPrice = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value);

                            itax = Convert.ToDecimal(this.pLC1DataGridView.Rows[e.RowIndex].Cells["TaxCode"].Value);
                            itax2 = itax / 100 + 1;
                            itax3 = itax / 100;
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Tax"].Value = (iQuantity * iUnitPrice * itax3).ToString();
                            this.pLC1DataGridView.Rows[e.RowIndex].Cells["Amt2"].Value = (iQuantity * iUnitPrice * itax2).ToString();
                        }
                         
                    }

         
                }
                if (pLC1DataGridView.Columns[e.ColumnIndex].Name == "Amt2")
                {
                    decimal iTotal = 0;
                    decimal iamt = 0;

                    int i = this.pLC1DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= i; iRecs++)
                    {
                        iTotal += Convert.ToDecimal(pLC1DataGridView.Rows[iRecs].Cells["Amt2"].Value);

                    }
                  
                    iamt = Convert.ToDecimal(lcAmtTextBox.Text);
                    decimal aa = iamt - iTotal;
                    aa = Math.Round(aa, 2, MidpointRounding.AwayFromZero);
                    lcTotalTextBox.Text = aa.ToString();
                }
            }
            catch (Exception ex)
            {
             
            }

        }


        private void button2_Click(object sender, EventArgs e)
        {

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\OPCH\\LC.xls";

            System.Data.DataTable T1 = ExecuteQuery();
                        //Excel的樣版檔
            string ExcelTemplate = FileName;

                        //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                        //產生 Excel Report
            ExcelReport.ExcelReportOutput(T1, ExcelTemplate, OutPutFile, "N");
          
        }

        private System.Data.DataTable ExecuteQuery()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select t0.bankName +'  L/C NO:'+t0.lcNo +'  '+bankCode+CAST(T0.LCAMT AS VARCHAR) LC,T1.DonNo 單號,T1.ItemName 品名,T1.Qty 數量,T1.Price 單價,T1.Tax 稅額,T1.AMT 金額,T0.lcTotal 未沖");
            sb.Append("                         , T1.CargoDate 出貨時間,T1.CargoDate2 押匯時間 ,T1.CardName 公司,T1.InvoceNo INVOICE,bankCode");
            sb.Append("                         ,'開狀日: '+lcDate +'  最後交貨日:'+lastDate+'   L/C有效期限:'+expDate LC2  from APLC T0 LEFT JOIN PLC1 T1 ON (T0.DocNum=T1.DocNum) ");
            sb.Append(" where t0.docNum=@DocNum ");
            sb.Append(" ORDER BY ISNULL(T1.CargoDate2,'A')");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable ExecuteQuery2(string INVO, string aa, string GH, string AMT, string AUO, string AUO2, string AUO3, string AUO4, string AUO5, string COMPANY, string DATENAME, string APMEMO, string QTY)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                 select case substring(T1.CargoDate2,7,2) ");
            sb.Append("                         when '01' then '1'  when '02' then '2' when '03' then '3' ");
            sb.Append("                         when '04' then '4'  when '05' then '5' when '06' then '6' ");
            sb.Append("                         when '07' then '7'  when '08' then '8' when '09' then '9'");
            sb.Append("                         else   substring(T1.CargoDate2,7,2)  end +");
            sb.Append("                        case substring(T1.CargoDate2,7,2) ");
            sb.Append("  when '01' then 'st' when '02' then 'nd' when '03' then 'rd'");
            sb.Append("                                     when '21' then 'st' when '22' then 'nd' when '23' then 'rd' ");
            sb.Append("                                     when '31' then 'st' ");
            sb.Append("                                     else  'th' ");
            sb.Append("                        end");
            sb.Append("                        +' '+");
            sb.Append("                        case substring(T1.CargoDate2,5,2) ");
            sb.Append("                        when '01' then 'JAN' when '02' then 'FEB' when '03' then 'MAR' ");
            sb.Append("                        when '04' then 'APR' when '05' then 'MAY' when '06' then 'JUN' ");
            sb.Append("                        when '07' then 'JUL' when '08' then 'AUG' when '09' then 'SEP' ");
            sb.Append("                        when '10' then 'OCT' when '11' then 'NOV' when '12' then 'DEC' ");
            sb.Append("                        END+'. '+ substring(T1.CargoDate2,1,4)  AS DATE,");
            sb.Append("                         @aa+T0.LCNO+' DATE: '+case substring(T0.lcDate,7,2) ");
            sb.Append("                         when '01' then '1'  when '02' then '2' when '03' then '3' ");
            sb.Append("                         when '04' then '4'  when '05' then '5' when '06' then '6' ");
            sb.Append("                         when '07' then '7'  when '08' then '8' when '09' then '9'");
            sb.Append("                         else   substring(T0.lcDate,7,2)  end");
            sb.Append("                         +case substring(T0.lcDate,7,2) ");
            sb.Append("  when '01' then 'st' when '02' then 'nd' when '03' then 'rd'");
            sb.Append("                                     when '21' then 'st' when '22' then 'nd' when '23' then 'rd' ");
            sb.Append("                                     when '31' then 'st' ");
            sb.Append("                                     else  'th' ");
            sb.Append("                        end");
            sb.Append("                        +' '+");
            sb.Append("                        case substring(T0.lcDate,5,2) ");
            sb.Append("                        when '01' then 'JAN' when '02' then 'FEB' when '03' then 'MAR' ");
            sb.Append("                        when '04' then 'APR' when '05' then 'MAY' when '06' then 'JUN' ");
            sb.Append("                        when '07' then 'JUL' when '08' then 'AUG' when '09' then 'SEP' ");
            sb.Append("                        when '10' then 'OCT' when '11' then 'NOV' when '12' then 'DEC' ");
            sb.Append("                        END+'. '+ substring(T0.lcDate,1,4)+' IN GOOD ORDER AND CONDITION ON'  AS LC,");
            sb.Append("                        T2.BANKNAME BANK,T2.DESCRIPTION DES,T2.GOODS GOODS,T2.FOB FOB,T0.BANKCODE USD,t0.incoiveno INV,AMT=@AMT,INVNO=@INVNO,GH=@GH,AUO=@AUO,AUO2=@AUO2,AUO3=@AUO3,AUO4=@AUO4,AUO5=@AUO5,COMPANY=@COMPANY,DATENAME=@DATENAME,APMEMO=@APMEMO,QTY=@QTY");
            sb.Append("                        from APLC T0");
            sb.Append("                        LEFT JOIN PLC1 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append("                          LEFT JOIN AP_BANK T2 ON (T0.BANKNAME=T2.BANKCODE) ");
            sb.Append("                        WHERE (isnull(T1.STATUS,'') = '' OR isnull(T1.STATUS,'') = 'False') and t1.invoceno is not NULL ");
            sb.Append("                        and t0.docNum=@DocNum ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));
            command.Parameters.Add(new SqlParameter("@INVNO", INVO));
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@GH", GH));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@AUO", AUO));
            command.Parameters.Add(new SqlParameter("@AUO2", AUO2));
            command.Parameters.Add(new SqlParameter("@AUO3", AUO3));
            command.Parameters.Add(new SqlParameter("@AUO4", AUO4));
            command.Parameters.Add(new SqlParameter("@AUO5", AUO5));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@DATENAME", DATENAME));
            command.Parameters.Add(new SqlParameter("@APMEMO", APMEMO));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable ExecuteQuery2CARG(string INVO, string aa, string GH, string AMT, string AUO, string AUO2, string AUO3, string AUO4, string AUO5, string COMPANY, string DATENAME, string APMEMO, string QTY, string CargoDate2)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                 select case substring(T1.CargoDate2,7,2) ");
            sb.Append("                         when '01' then '1'  when '02' then '2' when '03' then '3' ");
            sb.Append("                         when '04' then '4'  when '05' then '5' when '06' then '6' ");
            sb.Append("                         when '07' then '7'  when '08' then '8' when '09' then '9'");
            sb.Append("                         else   substring(T1.CargoDate2,7,2)  end +");
            sb.Append("                        case substring(T1.CargoDate2,7,2) ");
            sb.Append("  when '01' then 'st' when '02' then 'nd' when '03' then 'rd'");
            sb.Append("                                     when '21' then 'st' when '22' then 'nd' when '23' then 'rd' ");
            sb.Append("                                     when '31' then 'st' ");
            sb.Append("                                     else  'th' ");
            sb.Append("                        end");
            sb.Append("                        +' '+");
            sb.Append("                        case substring(T1.CargoDate2,5,2) ");
            sb.Append("                        when '01' then 'JAN' when '02' then 'FEB' when '03' then 'MAR' ");
            sb.Append("                        when '04' then 'APR' when '05' then 'MAY' when '06' then 'JUN' ");
            sb.Append("                        when '07' then 'JUL' when '08' then 'AUG' when '09' then 'SEP' ");
            sb.Append("                        when '10' then 'OCT' when '11' then 'NOV' when '12' then 'DEC' ");
            sb.Append("                        END+'. '+ substring(T1.CargoDate2,1,4)  AS DATE,");
            sb.Append("                         @aa+T0.LCNO+' DATE: '+case substring(T0.lcDate,7,2) ");
            sb.Append("                         when '01' then '1'  when '02' then '2' when '03' then '3' ");
            sb.Append("                         when '04' then '4'  when '05' then '5' when '06' then '6' ");
            sb.Append("                         when '07' then '7'  when '08' then '8' when '09' then '9'");
            sb.Append("                         else   substring(T0.lcDate,7,2)  end");
            sb.Append("                         +case substring(T0.lcDate,7,2) ");
            sb.Append("  when '01' then 'st' when '02' then 'nd' when '03' then 'rd'");
            sb.Append("                                     when '21' then 'st' when '22' then 'nd' when '23' then 'rd' ");
            sb.Append("                                     when '31' then 'st' ");
            sb.Append("                                     else  'th' ");
            sb.Append("                        end");
            sb.Append("                        +' '+");
            sb.Append("                        case substring(T0.lcDate,5,2) ");
            sb.Append("                        when '01' then 'JAN' when '02' then 'FEB' when '03' then 'MAR' ");
            sb.Append("                        when '04' then 'APR' when '05' then 'MAY' when '06' then 'JUN' ");
            sb.Append("                        when '07' then 'JUL' when '08' then 'AUG' when '09' then 'SEP' ");
            sb.Append("                        when '10' then 'OCT' when '11' then 'NOV' when '12' then 'DEC' ");
            sb.Append("                        END+'. '+ substring(T0.lcDate,1,4)+' IN GOOD ORDER AND CONDITION ON'  AS LC,");
            sb.Append("                        T2.BANKNAME BANK,T2.DESCRIPTION DES,T2.GOODS GOODS,T2.FOB FOB,T0.BANKCODE USD,t0.incoiveno INV,AMT=@AMT,INVNO=@INVNO,GH=@GH,AUO=@AUO,AUO2=@AUO2,AUO3=@AUO3,AUO4=@AUO4,AUO5=@AUO5,COMPANY=@COMPANY,DATENAME=@DATENAME,APMEMO=@APMEMO,QTY=@QTY");
            sb.Append("                        from APLC T0");
            sb.Append("                        LEFT JOIN PLC1 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append("                          LEFT JOIN AP_BANK T2 ON (T0.BANKNAME=T2.BANKCODE) ");
            sb.Append("                        WHERE t0.docNum=@DocNum AND  CargoDate2 =@CargoDate2");

            sb.Append("    AND  (isnull(T1.STATUS,'') = '' OR isnull(T1.STATUS,'') = 'False') ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));
            command.Parameters.Add(new SqlParameter("@INVNO", INVO));
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@GH", GH));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@AUO", AUO));
            command.Parameters.Add(new SqlParameter("@AUO2", AUO2));
            command.Parameters.Add(new SqlParameter("@AUO3", AUO3));
            command.Parameters.Add(new SqlParameter("@AUO4", AUO4));
            command.Parameters.Add(new SqlParameter("@AUO5", AUO5));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@DATENAME", DATENAME));
            command.Parameters.Add(new SqlParameter("@APMEMO", APMEMO));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CargoDate2", CargoDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable AMT()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(T1.AMT) AMT,SUM(T1.qty) QTY from APLC T0");
            sb.Append(" LEFT JOIN PLC1 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append(" WHERE (isnull(T1.STATUS,'') = '' OR isnull(T1.STATUS,'') = 'False')  and t1.invoceno is not NULL ");
            sb.Append(" and t0.docNum=@DocNum");
         
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable AMTCARG(string CargoDate2)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(T1.AMT) AMT,SUM(T1.qty) QTY from APLC T0");
            sb.Append(" LEFT JOIN PLC1 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append(" WHERE (isnull(T1.STATUS,'') = '' OR isnull(T1.STATUS,'') = 'False')  and t1.invoceno is not NULL ");
            sb.Append(" and t0.docNum=@DocNum AND  CargoDate2 =@CargoDate2 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));
            command.Parameters.Add(new SqlParameter("@CargoDate2", CargoDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button4_Click(object sender, EventArgs e)
        {
            
            AP frm1 = new AP();
            frm1.cardcode = cardCodeTextBox.Text;
            frm1.usd = bankCodeTextBox.Text;
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (frm1.c == "1")
                    {


                        System.Data.DataTable dt1 = GetMenu.GetAR2(frm1.a);
                        System.Data.DataTable dt11 = GetMenu.GetAR3(frm1.a);
                        System.Data.DataTable dt2 = lC.PLC1;

                        if (bankCodeTextBox.Text == "NTD")
                        {
                            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt1.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "AP發票";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                decimal taxx3 = 1 + taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                            
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx3), 0, MidpointRounding.AwayFromZero).ToString();
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = cardCodeTextBox.Text;
                                drw2["CardName2"] = cardNameTextBox.Text;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }
                        else
                        {
                            for (int i = 0; i <= dt11.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt11.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "AP發票";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["LcNo"] = drw["匯率"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                decimal taxx3 = 1 + taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx3), 2, MidpointRounding.AwayFromZero).ToString();
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = cardCodeTextBox.Text;
                                drw2["CardName2"] = cardNameTextBox.Text;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }
                            
                    }
                    if (frm1.c == "2")
                    {
                        try
                        {

                            System.Data.DataTable dt1 = GetMenu.GetOP2(frm1.a);

                            System.Data.DataTable dt2 = lC.PLC1;

                            if (bankCodeTextBox.Text == "NTD")
                                {
                                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                    {
                                        string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                        DataRow drw = dt1.Rows[i];
                                        DataRow drw2 = dt2.NewRow();
                                        drw2["DocNum"] = docNumTextBox.Text;
                                        drw2["LcNo"] = drw["rate"];
                                        drw2["PKind"] = "採購單";
                                        drw2["DonNo"] = drw["DocNum"];
                                        drw2["ChNo"] = drw["U_CHI_NO"];
                                        drw2["ItemCode"] = drw["ItemCode"];
                                        drw2["ItemName"] = drw["Dscription"];
                                        drw2["Comments"] = drw["Price"];
                                        string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                        decimal price = Convert.ToDecimal(drw["Price"]) * Convert.ToDecimal(drw["rate"]);
                                        decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                        decimal taxx2 = taxx / 100;
                                        decimal taxx3 = 1 + taxx / 100;
                                        drw2["Qty"] = qry;
                                        drw2["Price"] = price;
                                        drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);

                                        drw2["Amt"] =
                                         ((Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])) * (Convert.ToDecimal(drw["Price"]) * Convert.ToDecimal(drw["rate"])) + Convert.ToDecimal(drw["VatSumsy"])).ToString("0");
                                        string amt = drw2["Amt"].ToString();
                                        drw2["Tax"] = drw["VatSumsy"];
                                        drw2["CardCode"] = cardCodeTextBox.Text;
                                        drw2["CardName2"] = cardNameTextBox.Text;
                                        drw2["LineNum"] = drw["LineNum"];
                                        dt2.Rows.Add(drw2);

                                    }
                                }
                                else
                                {
                                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                    {
                                        string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                        DataRow drw = dt1.Rows[i];
                                        DataRow drw2 = dt2.NewRow();
                                        drw2["DocNum"] = docNumTextBox.Text;
                                        drw2["LcNo"] = drw["rate"];
                                        drw2["PKind"] = "採購單";
                                        drw2["DonNo"] = drw["DocNum"];
                                        drw2["ChNo"] = drw["U_CHI_NO"];
                                        drw2["ItemCode"] = drw["ItemCode"];
                                        drw2["ItemName"] = drw["Dscription"];
                                        drw2["Comments"] = drw["Price"];
                                        decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                        decimal taxx2 = taxx / 100;
                                        string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                        string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                        drw2["Qty"] = qry;
                                        drw2["Price"] = drw["Price"];
                                        drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                 
                                        drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) + Convert.ToDecimal(tax)), 2, MidpointRounding.AwayFromZero).ToString();
            
                                        drw2["Tax"] = tax;
                                        drw2["CardCode"] = cardCodeTextBox.Text;
                                        drw2["CardName2"] = cardNameTextBox.Text;
                                        drw2["LineNum"] = drw["LineNum"];
                                        dt2.Rows.Add(drw2);

                                    }

                            }
               


                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }


                    }
                    if (frm1.c == "4")
                    {
                        try
                        {

                            System.Data.DataTable dt1 = GetMenu.GetOP2Q(frm1.a);

                            System.Data.DataTable dt2 = lC.PLC1;

                            if (bankCodeTextBox.Text == "NTD")
                            {
                                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                {
                                    string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                    DataRow drw = dt1.Rows[i];
                                    DataRow drw2 = dt2.NewRow();
                                    drw2["DocNum"] = docNumTextBox.Text;
                                    drw2["LcNo"] = drw["rate"];
                                    drw2["PKind"] = "採購報價";
                                    drw2["DonNo"] = drw["DocNum"];
                                    drw2["ChNo"] = drw["U_CHI_NO"];
                                    drw2["ItemCode"] = drw["ItemCode"];
                                    drw2["ItemName"] = drw["Dscription"];
                                    drw2["Comments"] = drw["Price"];
                                    string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                    decimal price = Convert.ToDecimal(drw["Price"]) * Convert.ToDecimal(drw["rate"]);
                                    decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                    decimal taxx2 = taxx / 100;
                                    decimal taxx3 = 1 + taxx / 100;
                                    drw2["Qty"] = qry;
                                    drw2["Price"] = price;
                                    drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);

                                    drw2["Amt"] =
                                     ((Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])) * (Convert.ToDecimal(drw["Price"]) * Convert.ToDecimal(drw["rate"])) + Convert.ToDecimal(drw["VatSumsy"])).ToString("0");
                                    string amt = drw2["Amt"].ToString();
                                    drw2["Tax"] = drw["VatSumsy"];
                                    drw2["CardCode"] = cardCodeTextBox.Text;
                                    drw2["CardName2"] = cardNameTextBox.Text;
                                    drw2["LineNum"] = drw["LineNum"];
                                    dt2.Rows.Add(drw2);

                                }
                            }
                            else
                            {
                                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                {
                                    string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                    DataRow drw = dt1.Rows[i];
                                    DataRow drw2 = dt2.NewRow();
                                    drw2["DocNum"] = docNumTextBox.Text;
                                    drw2["LcNo"] = drw["rate"];
                                    drw2["PKind"] = "採購報價";
                                    drw2["DonNo"] = drw["DocNum"];
                                    drw2["ChNo"] = drw["U_CHI_NO"];
                                    drw2["ItemCode"] = drw["ItemCode"];
                                    drw2["ItemName"] = drw["Dscription"];
                                    drw2["Comments"] = drw["Price"];
                                    decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                    decimal taxx2 = taxx / 100;
                                    string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                    string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                    drw2["Qty"] = qry;
                                    drw2["Price"] = drw["Price"];
                                    drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);

                                    drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) + Convert.ToDecimal(tax)), 2, MidpointRounding.AwayFromZero).ToString();

                                    drw2["Tax"] = tax;
                                    drw2["CardCode"] = cardCodeTextBox.Text;
                                    drw2["CardName2"] = cardNameTextBox.Text;
                                    drw2["LineNum"] = drw["LineNum"];
                                    dt2.Rows.Add(drw2);

                                }

                            }



                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }


                    }
                    if (frm1.c == "3")
                    {


                        System.Data.DataTable dt1 = GetMenu.GetAR22(frm1.a);
                        System.Data.DataTable dt11 = GetMenu.GetAR32(frm1.a);
                        System.Data.DataTable dt2 = lC.PLC1;

                        if (bankCodeTextBox.Text == "NTD")
                        {
                            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt1.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "收貨採購";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                drw2["Amt"] = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) + Convert.ToDecimal(tax)).ToString();
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = cardCodeTextBox.Text;
                                drw2["CardName2"] = cardNameTextBox.Text;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }
                        else
                        {
                            for (int i = 0; i <= dt11.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt11.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "收貨採購";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["LcNo"] = drw["匯率"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                drw2["Amt"] = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) + Convert.ToDecimal(tax)).ToString();
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = cardCodeTextBox.Text;
                                drw2["CardName2"] = cardNameTextBox.Text;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }

                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                try
                {
                    decimal iTotal = 0;
                    decimal iamt = 0;

                    int x = this.pLC1DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= x; iRecs++)
                    {
                        iTotal += Convert.ToDecimal(pLC1DataGridView.Rows[iRecs].Cells["Amt2"].Value);

                    }

                    iamt = Convert.ToDecimal(lcAmtTextBox.Text);
                    decimal aa = iamt - iTotal;
                    aa = Math.Round(aa, 2, MidpointRounding.AwayFromZero);
                    lcTotalTextBox.Text = aa.ToString();
                }
                catch { }
            }

        }



       
        private void APLC_Load(object sender, EventArgs e)
        {
            STATUS = "";
            textBox6.Text = GetMenu.DFirst();
            textBox7.Text = GetMenu.DLast();
            WW();
            pLC1DataGridView.Enabled = false;
            //
            DataGridViewLinkColumn column = new DataGridViewLinkColumn();
            column.Name = "Link";
            column.UseColumnTextForLinkValue = true;
            column.Text = "讀取檔案";
            column.LinkBehavior = LinkBehavior.HoverUnderline;
            column.TrackVisitedState = true;
            aP_DownloadDataGridView.Columns.Add(column);

        }

        public static System.Data.DataTable Getinvoice(string docnum)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select invoceno from plc1 where docnum=@docnum and status is null and invoceno is not null";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docnum", docnum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }

        public static System.Data.DataTable Getpor1(string docno, string linenum)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql = "select targettype targettype from por1 where docentry=@docno and  linenum=@linenum ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docno", docno));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
        public static System.Data.DataTable Getinvoice1(string docnum)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select docnum,sum(amt) as aa from plc1 where docnum=@docnum and status is null and invoceno is not NULL group by docnum";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docnum", docnum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
   



        private void button5_Click(object sender, EventArgs e)
        {
     
            try
            {

               
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    try
                    {
            
                            FileName = lsAppDir + "\\Excel\\OPCH\\" + bankNameComboBox.Text + "NTD.xls";


                            if (dCHECKTextBox.Text.Trim() == "Checked")
                        {
                            FileName = lsAppDir + "\\Excel\\OPCH\\達運NTD.xls";
                        }
                        if (!String.IsNullOrEmpty(FileName))
                        {
                            System.Data.DataTable T1 = GETNTD(docNumTextBox.Text);

                            string PKIND = "";
                            if (T1.Rows.Count > 0)
                            {
                                PKIND = T1.Rows[0][0].ToString();


                                if (PKIND == "採購單")
                                {
                                    StringBuilder sb2 = new StringBuilder();
                                    System.Data.DataTable T2 = GETINV(docNumTextBox.Text);
                                    if (T2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i <= T2.Rows.Count - 1; i++)
                                        {
                                            string QTY = T2.Rows[i]["QTY"].ToString();
                                            string INV = T2.Rows[i]["INV"].ToString();
                                            System.Data.DataTable T3 = GETINV2(INV);
                                            string QTY2 = T3.Rows[0]["QTY"].ToString();
                                            if (QTY2 != QTY)
                                            {
                                                sb2.Append(INV + Environment.NewLine);
                                            }
                                        }

                                        if (!String.IsNullOrEmpty(sb2.ToString()))
                                        {
                                            string INVD = "數量不符INVOICE" + Environment.NewLine + sb2.ToString();
                                            MessageBox.Show(INVD);
                                            DialogResult result;
                                            result = MessageBox.Show("請確定是否要匯出CR", "YesNo", MessageBoxButtons.YesNo);
                                            if (result == DialogResult.No)
                                            {
                                                return;
                                            }

                                        }
                                    }

                                }
                                if (PKIND == "收貨採購")
                                {
                                    T1 = GETNTD2(docNumTextBox.Text);
                                }

                                string ExcelTemplate = FileName;

                                //輸出檔
                                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                                //產生 Excel Report
                                ExcelReport.ExcelReportOutput(T1, ExcelTemplate, OutPutFile, "N");
                            }
                        }
                    }
                    catch { }
                    try
                    {
                        System.Data.DataTable dt1 = Getinvoiceno(docNumTextBox.Text);
                        if (dt1.Rows.Count > 0)
                        {
                            string COMPANY = GETLCCOMP(bankNameComboBox.Text, "COMPANY").Rows[0][0].ToString();
                            string AUO3 = GETLCCOMP(bankNameComboBox.Text, "AUO3").Rows[0][0].ToString();
                            string AUO4 = GETLCCOMP(bankNameComboBox.Text, "AUO4").Rows[0][0].ToString();
                            string AUO5 = GETLCCOMP(bankNameComboBox.Text, "AUO5").Rows[0][0].ToString();
                  
                            FileName = lsAppDir + "\\Excel\\OPCH\\兆豐.xls";
                            if (bankNameComboBox.Text == "元大")
                            {
                                FileName = lsAppDir + "\\Excel\\OPCH\\元大.xls";
                            }
   
                            System.Data.DataTable dt5 = AP_ANK2();
                            StringBuilder sb2 = new StringBuilder();
                            for (int i = 0; i <= dt5.Rows.Count - 1; i++)
                            {
                                DataRow dd = dt5.Rows[i];



                                sb2.Append(dd["InvoceNo"].ToString() + "、");


                            }
                            sb2.Remove(sb2.Length - 1, 1);
                            string ef = sb2.ToString();


                            string lcno;
                            if (bankNameComboBox.Text == "國泰" || bankNameComboBox.Text == "ICBC")
                            {
                                lcno = "L/C NBR.";

                            }
                            else
                            {
                                lcno = "L/C NO.";
                            }

                            CalcTotals2();
                            System.Data.DataTable FG = AMT();
                            DataRow drw = FG.Rows[0];
                            string F = drw["AMT"].ToString();
                            string GH = drw["QTY"].ToString();
    
                  
                            string AUO = "";
                            string AUO2 = "";
                            string DATENAME = "";
                            string APMEMO = "";
                
                                if (bankNameComboBox.Text == "一銀")
                                {
                                    FileName = lsAppDir + "\\Excel\\OPCH\\一銀.xls";
                                }
                                if (bankNameComboBox.Text == "國泰")
                                {
                                    AUO = "TO:     AU OPTRONICS CORPORATION" + "\r\n" + "          NO.1,LI-HSIN ROAD 2,  SCIENCE-" + "\r\n" + "          BASED INDUSTRIAL PARK  HSIN-" + "\r\n" + "          CHU TAIWAN" + "\r\n" + "           TEL : 886-3-563-2899";
               
                                }
                                else
                                {
                                    AUO = "TO:     AU OPTRONICS CORPORATION NO.1," + "\r\n" + "           LI-HSIN ROAD 2,  SCIENCE-BASED" + "\r\n" + "           INDUSTRIAL PARK  HSIN-CHU CITY," + "\r\n" + "           TAIWAN" + "\r\n" + "           TEL : 886-3-563-2899";
                         
                                }
               
                                DATENAME = "ISSUING DATE:";

                                 AUO2 = "HAS COMPLIED WITH THE P/O AND L/C TERMS FROM AU OPTRONICS CORPORATION  UNDER";
                         



                            OrderData = ExecuteQuery2(ef, lcno, GH, F, AUO, AUO2, AUO3, AUO4, AUO5, COMPANY, DATENAME, APMEMO, GH);
                            string NewFileName2 = lsAppDir + "\\Excel\\temp\\" +
DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                            ExcelReport.ExcelReportOutput(OrderData, FileName, NewFileName2, "N");
                          //  GetExcelProduct(FileName);
                        }
                        else
                        {
                            MessageBox.Show("Cargo 沒有資料");
                        }
                        // }
                    }
                    catch (Exception ex)
                    {
                    }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void A1(string CARGODATE,string LC)
        {
            try
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
        
                    string COMPANY = GETLCCOMP(bankNameComboBox.Text, "COMPANY").Rows[0][0].ToString();
                    string AUO3 = GETLCCOMP(bankNameComboBox.Text, "AUO3").Rows[0][0].ToString();
                    string AUO4 = GETLCCOMP(bankNameComboBox.Text, "AUO4").Rows[0][0].ToString();
                    string AUO5 = GETLCCOMP(bankNameComboBox.Text, "AUO5").Rows[0][0].ToString();

                    FileName = lsAppDir + "\\Excel\\OPCH\\兆豐.xls";
                    if (bankNameComboBox.Text == "元大")
                    {
                        FileName = lsAppDir + "\\Excel\\OPCH\\元大.xls";
                    }

                    System.Data.DataTable dt5 = AP_ANK2CARG(CARGODATE);
                    StringBuilder sb2 = new StringBuilder();
                    for (int i = 0; i <= dt5.Rows.Count - 1; i++)
                    {
                        DataRow dd = dt5.Rows[i];



                        sb2.Append(dd["InvoceNo"].ToString() + "、");


                    }
                    sb2.Remove(sb2.Length - 1, 1);
                    string ef = sb2.ToString();


                    string lcno;
                    if (bankNameComboBox.Text == "國泰" || bankNameComboBox.Text == "ICBC")
                    {
                        lcno = "L/C NBR.";

                    }
                    else
                    {
                        lcno = "L/C NO.";
                    }

                    CalcTotals2();
                    System.Data.DataTable FG = AMTCARG(CARGODATE);
                    DataRow drw = FG.Rows[0];
                    string F = drw["AMT"].ToString();
                    string GH = drw["QTY"].ToString();


                    string AUO = "";
                    string AUO2 = "";
                    string DATENAME = "";
                    string APMEMO = "";
            
                        if (bankNameComboBox.Text == "一銀")
                        {
                            FileName = lsAppDir + "\\Excel\\OPCH\\一銀.xls";
                        }
                        if (bankNameComboBox.Text == "國泰")
                        {
                            AUO = "TO:     AU OPTRONICS CORPORATION" + "\r\n" + "          NO.1,LI-HSIN ROAD 2,  SCIENCE-" + "\r\n" + "          BASED INDUSTRIAL PARK  HSIN-" + "\r\n" + "          CHU TAIWAN" + "\r\n" + "           TEL : 886-3-563-2899";

                        }
                        else
                        {
                            AUO = "TO:     AU OPTRONICS CORPORATION NO.1," + "\r\n" + "           LI-HSIN ROAD 2,  SCIENCE-BASED" + "\r\n" + "           INDUSTRIAL PARK  HSIN-CHU CITY," + "\r\n" + "           TAIWAN" + "\r\n" + "           TEL : 886-3-563-2899";

                        }

                        DATENAME = "ISSUING DATE:";

                        AUO2 = "HAS COMPLIED WITH THE P/O AND L/C TERMS FROM AU OPTRONICS CORPORATION  UNDER";
              



                    OrderData = ExecuteQuery2CARG(ef, lcno, GH, F, AUO, AUO2, AUO3, AUO4, AUO5, COMPANY, DATENAME, APMEMO, GH, CARGODATE);
                    if (OrderData.Rows.Count > 0)
                    {
                        string S1 = lcNoTextBox.Text.Replace("/", "") + "_CR_USD" + LC;
                        string NewFileName2 = lsAppDir + "\\Excel\\RMA\\temp\\" +
        S1 + ".xls";


                        ExcelReport.ExcelReportOutputLEMON(OrderData, FileName, NewFileName2, "N");
                    }
        
            }
            catch (Exception ex)
            {
            }
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("PO No.", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(int));
            dt.Columns.Add("單價USD", typeof(decimal));
            dt.Columns.Add("5%", typeof(decimal));
            dt.Columns.Add("金額USD", typeof(decimal));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("Invoice No.", typeof(string));
            dt.Columns.Add("出貨時間", typeof(string));
            dt.Columns.Add("押匯時間", typeof(string));

            return dt;
        }
        private void GetExcelProduct(string ExcelFile)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            string aa = lcNoTextBox.Text.Replace("/", "-");
            aa = aa.Trim();
            string bb = bankNameComboBox.Text.Trim();
           // excelSheet.Name = bankNameComboBox.Text.Substring(0, 4) ;
                //aa.Substring(0,6);

            excelSheet.Name = bb+aa;
            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                string FieldValue2 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 21;
                int DetailRow2 = 25;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    //progressBar1.Value = iRecord;
                    //progressBar1.PerformStep();


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                       
                            break;
                        }


                    }

                }

      

      
     

            }
            finally
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
  DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
               
                }
                catch
                {

                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                string Msg = string.Empty;
                System.Diagnostics.Process.Start(NewFileName);
           

            }
        }



        private bool CheckSerial(string sData, ref string FieldValue)
        {

            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }
        private bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }

            //}
            return false;
        }
        private void SetRow(int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[iRow][FieldName]);
            }

        }
        private System.Data.DataTable AP_ANK1()
        {
            
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select description as DES from ap_ank1 ");
            sb.Append(" where bankcode=@bankcode  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@bankcode",bankNameComboBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ap_ank1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable DIST()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT bankName 銀行,T0.lcNo LC,T0.lastDate 最後交貨日 ");
            sb.Append("         ,expDate 有效期限,bankCode 幣別,");
            sb.Append("         lcAmt 開狀金額,LCAMT-LCTOTAL 已使用金額,LCTOTAL 未沖金額,memo 結案原因 ");
            sb.Append("         from APLC T0 ");
            sb.Append("         where isnull(lcClose,'') <> 'checked'  and  isnull(FCHECK,'') <> 'checked'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ap_ank1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable AP_ANK2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

    
            sb.Append("     Select   case when CHARINDEX('(', InvoceNo) > 0 then  substring(InvoceNo,0,CHARINDEX('(', InvoceNo)) else InvoceNo end InvoceNo from plc1  ");
            sb.Append(" WHERE (isnull(STATUS,'') = '' OR isnull(STATUS,'') = 'False')  and isnull(invoceno,'')  <> '' and isnull(CargoDate2,'') <> '' ");
            sb.Append("              AND  docNum=@DocNum ");
            sb.Append("			  GROUP BY case when CHARINDEX('(', InvoceNo) > 0 then  substring(InvoceNo,0,CHARINDEX('(', InvoceNo)) else InvoceNo end");
            sb.Append("			  order by MAX(DOCENTRY1) ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable AP_ANK2CARG(string CargoDate2)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" Select distinct case when CHARINDEX('(', InvoceNo) > 0 then  substring(InvoceNo,0,CHARINDEX('(', InvoceNo)) else InvoceNo end InvoceNo from plc1 ");
            sb.Append(" WHERE docNum=@DocNum AND  CargoDate2 =@CargoDate2  order by InvoceNo ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", docNumTextBox.Text));
            command.Parameters.Add(new SqlParameter("@CargoDate2", CargoDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                string server = "//acmesrv01//SAP_Share//LC//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = download(filename);
                if (dt2.Rows.Count > 0)
                {
                    MessageBox.Show("檔案名稱重複,請修改檔名");
                }
                else
                {
                    if (result == DialogResult.OK)


                        MessageBox.Show(Path.GetFileName(opdf.FileName));
                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }
                    System.Data.DataTable dt1 = lC.AP_Download;

                    DataRow drw = dt1.NewRow();
                    drw["docNum"] = docNumTextBox.Text;
                    drw["seq"] = (aP_DownloadDataGridView.Rows.Count).ToString();
                    drw["filename"] = filename;
                    drw["path"] = @"\\acmesrv01\SAP_Share\LC\" + filename;
                    dt1.Rows.Add(drw);
                    this.aP_DownloadBindingSource.EndEdit();
                    this.aP_DownloadTableAdapter.Update(lC.AP_Download);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable download(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from AP_download where [filename] = @DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " AP_download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" AP_download "];
        }

        private void aP_DownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;


                if (dgv.Columns[e.ColumnIndex].Name == "Link")
                {
                    System.Data.DataTable dt1 = lC.AP_Download;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];

                    string aa = drw["path"].ToString();


                    System.Diagnostics.Process.Start(aa);

                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

 

        private void button8_Click(object sender, EventArgs e)
        {
            CalcTotals2();
        }
        private void CalcTotals2()
        {
        

            Int32 iTotal = 0;
            decimal iVatSum = 0;
            decimal iVatSum2 = 0;

            int i = this.pLC1DataGridView.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(pLC1DataGridView.SelectedRows[iRecs].Cells["Qty"].Value);
                iVatSum += Convert.ToDecimal(pLC1DataGridView.SelectedRows[iRecs].Cells["Tax"].Value);
                iVatSum2 += Convert.ToDecimal(pLC1DataGridView.SelectedRows[iRecs].Cells["Amt2"].Value);
            }

            textBox2.Text = iTotal.ToString("0");
            textBox3.Text = iVatSum.ToString();
            textBox4.Text = iVatSum2.ToString();
        
           

        }



        public static System.Data.DataTable download2(string lcno)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select lcNo from APLC where   lcno=@lcno";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@lcno", lcno));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " APLC ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" APLC "];
        }
        public static System.Data.DataTable download3(string DOCNUM)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select lcNo from APLC where   DOCNUM=@DOCNUM";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " APLC ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" APLC "];
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            System.Data.DataTable dt2 = lC.PLC1;
            DataRow newCustomersRow = dt2.NewRow();

            int i = pLC1DataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["DocNum"] = drw["DocNum"];

            newCustomersRow["DOCENTRY"] = AutoNum;
            newCustomersRow["Pkind"] = drw["Pkind"];
            newCustomersRow["DonNo"] = drw["DonNo"];
            newCustomersRow["ChNo"] = drw["ChNo"];
            newCustomersRow["ItemCode"] = drw["ItemCode"];
            newCustomersRow["ItemName"] = drw["ItemName"];
            newCustomersRow["Qty"] = drw["Qty"];
            newCustomersRow["Price"] = drw["Price"];
            newCustomersRow["TaxCode"] = drw["TaxCode"];
            newCustomersRow["Tax"] = drw["Tax"];
            newCustomersRow["Amt"] = drw["Amt"];
            newCustomersRow["CardName"] = drw["CardName"];
            newCustomersRow["InvoceNo"] = drw["InvoceNo"];
            newCustomersRow["CargoDate"] = drw["CargoDate"];

            newCustomersRow["CargoDate2"] = drw["CargoDate2"];
            newCustomersRow["SendDate"] = drw["SendDate"];
            newCustomersRow["status"] = drw["status"];
            newCustomersRow["Comments"] = drw["Comments"];
            newCustomersRow["CardCode"] = drw["CardCode"];
            newCustomersRow["CardName2"] = drw["CardName2"];
            newCustomersRow["LineNum"] = drw["LineNum"];

            try
            {
                dt2.Rows.InsertAt(newCustomersRow, pLC1DataGridView.Rows.Count);
                pLC1BindingSource1.DataSource = dt2;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
        public static System.Data.DataTable GetAPLLC(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT * FROM APLC where docnum=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public static System.Data.DataTable Getinvoiceno(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT * FROM PLC1 where docnum=@shippingcode and (isnull(STATUS,'') = '' OR isnull(STATUS,'') = 'False')  and isnull(invoceno,'')  <> '' and isnull(CargoDate2,'') <> ''  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }

        public static System.Data.DataTable GETA11(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT CARGODATE2 FROM PLC1 WHERE DOCNUM=@shippingcode AND ISNULL(CARGODATE2,'') <>'' AND (isnull(STATUS,'') = '' OR isnull(STATUS,'') = 'False') ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }

        public static System.Data.DataTable GETA12(string DOCNUM, string CargoDate2)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DONNO 'PO No.',ItemName 品名,Qty 數量,price '單價USD',TAX '5%',AMT '金額USD',CardName 廠商名稱,INVOCENO 'Invoice No.',CargoDate 出貨時間,CargoDate2 押匯時間 FROM PLC1 WHERE DOCNUM=@DOCNUM   AND  CargoDate2 =@CargoDate2  AND (isnull(STATUS,'') = '' OR isnull(STATUS,'') = 'False')  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            command.Parameters.Add(new SqlParameter("@CargoDate2", CargoDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static System.Data.DataTable GETA13(string DOCNUM, string CargoDate2)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SUM(AMT) '金額USD' FROM PLC1 WHERE DOCNUM=@DOCNUM   AND  CargoDate2 =@CargoDate2   AND (isnull(STATUS,'') = '' OR isnull(STATUS,'') = 'False')  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            command.Parameters.Add(new SqlParameter("@CargoDate2", CargoDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static System.Data.DataTable GETLCCOMP(string LCCOMPANY, string CTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT MEMO FROM AP_LCCOMPANY WHERE LCCOMPANY=@LCCOMPANY AND CTYPE=@CTYPE ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCCOMPANY", LCCOMPANY));
            command.Parameters.Add(new SqlParameter("@CTYPE", CTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static System.Data.DataTable GETQTY(string DOCNUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY) QTY  FROM PLC1 WHERE DOCNUM=@DOCNUM ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static System.Data.DataTable GETNTD(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                           SELECT MAX(PKIND) PKIND,LCAMT AMT,''''+T0.lcNo LC,SUBSTRING(T0.LCDATE,1,4)+'/'+SUBSTRING(T0.LCDATE,5,2)++'/'+SUBSTRING(T0.LCDATE,7,2) DATE, ");
            sb.Append("                           SUBSTRING(expDate,1,4)+'/'+SUBSTRING(expDate,5,2)++'/'+SUBSTRING(expDate,7,2) ");
            sb.Append("                           ENDDATE,T1.DonNo DOCENTRY,INVOCENO INV,U_PC_BSINV SAPINV,CONVERT(varchar(12),U_PC_BSDAT, 111) INVDATE, ");
            sb.Append("                           T3.DOCENTRY 收採,T4.DOCENTRY 採購單,T1.AMT AMT2,SUBSTRING(T0.lastDate,1,4)+'/'+SUBSTRING(T0.lastDate,5,2)++'/'+SUBSTRING(T0.lastDate,7,2) lastDate");
            sb.Append(" ,'民國 '+CAST(MAX(SUBSTRING(T0.LCDATE,1,4)-1911) AS VARCHAR) +' 年'+MAX(SUBSTRING(T0.LCDATE,5,2))+' 月'+ MAX(SUBSTRING(T0.LCDATE,7,2))+ ' 日' LCCDATE ,'民國 '+CAST(MAX(SUBSTRING(T0.expDate,1,4)-1911) AS VARCHAR) +' 年'+MAX(SUBSTRING(T0.expDate,5,2))+' 月'+ MAX(SUBSTRING(T0.expDate,7,2))+ ' 日' ENDDATE2 FROM APLC T0 ");
            sb.Append("             LEFT JOIN PLC1 T1  ON(T0.DOCNUM=T1.DOCNUM)");
            sb.Append("             LEFT JOIN ACMESQL02.DBO.OPCH T2 ON(T1.DONNO=T2.DOCENTRY)");
            sb.Append("             LEFT JOIN (SELECT DISTINCT DOCENTRY,TRGETENTRY FROM ACMESQL02.DBO.PDN1 WHERE TARGETTYPE=18 ) T3 ON(T2.DOCENTRY=T3.TRGETENTRY)");
            sb.Append("             LEFT JOIN (SELECT DISTINCT DOCENTRY,TRGETENTRY FROM ACMESQL02.DBO.POR1 WHERE TARGETTYPE=20 ) T4 ON(T3.DOCENTRY=T4.TRGETENTRY)");
            sb.Append("  where T0.docnum=@shippingcode AND (isnull(T1.status,'')) <> 'True' ");
            sb.Append(" GROUP BY LCAMT,T0.lcNo ,SUBSTRING(T0.LCDATE,1,4)+'/'+SUBSTRING(T0.LCDATE,5,2)++'/'+SUBSTRING(T0.LCDATE,7,2) ,");
            sb.Append("             SUBSTRING(expDate,1,4)+'/'+SUBSTRING(expDate,5,2)+'/'+SUBSTRING(expDate,7,2)");
            sb.Append("             ,T1.DonNo ,INVOCENO ,U_PC_BSINV,CONVERT(varchar(12),U_PC_BSDAT, 111) ,");
            sb.Append("             T3.DOCENTRY ,T4.DOCENTRY,T1.AMT,SUBSTRING(T0.lastDate,1,4)+'/'+SUBSTRING(T0.lastDate,5,2)++'/'+SUBSTRING(T0.lastDate,7,2)  ");
     

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }

        public static System.Data.DataTable GETINV(string DOCNUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT SUM(QTY) QTY,InvoceNo INV  FROM PLC1 WHERE DOCNUM=@DOCNUM AND ISNULL(InvoceNo,'') <> ''　GROUP BY InvoceNo ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", DOCNUM));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }


        public static System.Data.DataTable GETINV2(string U_ACME_INV)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT SUM(CAST(T1.Quantity AS INT)) QTY FROM OPDN T0 LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)　WHERE T0.U_ACME_INV=@U_ACME_INV ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static System.Data.DataTable GETNTD2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                           SELECT DISTINCT MAX(PKIND) PKIND,LCAMT AMT,''''+T0.lcNo LC,SUBSTRING(T0.LCDATE,1,4)+'/'+SUBSTRING(T0.LCDATE,5,2)++'/'+SUBSTRING(T0.LCDATE,7,2) DATE, ");
            sb.Append("                           SUBSTRING(expDate,1,4)+'/'+SUBSTRING(expDate,5,2)++'/'+SUBSTRING(expDate,7,2) ");
            sb.Append("                           ENDDATE,T1.DonNo DOCENTRY,INVOCENO INV,T5.U_PC_BSINV SAPINV,CONVERT(varchar(12),T5.U_PC_BSDAT, 111) INVDATE, ");
            sb.Append("                           T2.DOCENTRY 收採,T4.DOCENTRY 採購單,SUBSTRING(T0.lastDate,1,4)+'/'+SUBSTRING(T0.lastDate,5,2)++'/'+SUBSTRING(T0.lastDate,7,2) lastDate,T2.DOCTOTAL AMT2,'民國 '+CAST(MAX(SUBSTRING(T0.LCDATE,1,4)-1911) AS VARCHAR) +' 年'+MAX(SUBSTRING(T0.LCDATE,5,2))+' 月'+ MAX(SUBSTRING(T0.LCDATE,7,2))+ ' 日' LCCDATE,'民國 '+CAST(MAX(SUBSTRING(T0.expDate,1,4)-1911) AS VARCHAR) +' 年'+MAX(SUBSTRING(T0.expDate,5,2))+' 月'+ MAX(SUBSTRING(T0.expDate,7,2))+ ' 日' ENDDATE2 FROM APLC T0 ");
            sb.Append("                           LEFT JOIN PLC1 T1  ON(T0.DOCNUM=T1.DOCNUM) ");
            sb.Append("                           LEFT JOIN ACMESQL02.DBO.OPDN T2 ON(T1.DONNO=T2.DOCENTRY) ");
            sb.Append("                           LEFT JOIN (SELECT DISTINCT DOCENTRY,TRGETENTRY FROM ACMESQL02.DBO.POR1 WHERE TARGETTYPE=20 ) T4 ON(T2.DOCENTRY=T4.TRGETENTRY) ");
            sb.Append("                           LEFT JOIN (SELECT DISTINCT T1.U_PC_BSINV,BASEENTRY,U_PC_BSDAT FROM ACMESQL02.DBO.PCH1 T0 LEFT JOIN ACMESQL02.DBO.OPCH T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE BASETYPE=20 ) T5 ON(T2.DOCENTRY=T5.BASEENTRY)               ");
            sb.Append("  where T0.docnum=@shippingcode AND (isnull(T1.status,'')) <> 'True' ");
            sb.Append("               GROUP BY LCAMT,T0.lcNo ,SUBSTRING(T0.LCDATE,1,4)+'/'+SUBSTRING(T0.LCDATE,5,2)++'/'+SUBSTRING(T0.LCDATE,7,2) , ");
            sb.Append("                           SUBSTRING(expDate,1,4)+'/'+SUBSTRING(expDate,5,2)+'/'+SUBSTRING(expDate,7,2) ");
            sb.Append("                           ,T1.DonNo ,INVOCENO ,T5.U_PC_BSINV,CONVERT(varchar(12),T5.U_PC_BSDAT, 111) , ");
            sb.Append("                           T2.DOCENTRY ,T4.DOCENTRY,T1.AMT,SUBSTRING(T0.lastDate,1,4)+'/'+SUBSTRING(T0.lastDate,5,2)++'/'+SUBSTRING(T0.lastDate,7,2)  ,T2.DOCTOTAL");


         
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }



        private void button3_Click(object sender, EventArgs e)
        {
  


            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\OPCH\\LCBANK.xls";

            System.Data.DataTable T1 = DIST();
            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report
            ExcelReport.ExcelReportOutput(T1, ExcelTemplate, OutPutFile, "N");
        }

        private void pLC1DataGridView_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            try
            {
                CalcTotals();


            }
            catch 
            {
               
            }

   
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bankCodeTextBox.Text = comboBox1.Text;
        }

      

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataGridViewRow row;

            row = pLC1DataGridView.Rows[0];
            string a0 = row.Cells["CargoDate2"].Value.ToString();
            if (!String.IsNullOrEmpty(a0))
            {
                if (comboBox2.Text == "帶")
                {
                    for (int i = pLC1DataGridView.Rows.Count - 2; i >= 0; i--)
                    {
                        row = pLC1DataGridView.Rows[i];
                        row.Cells["CargoDate2"].Value = a0;
                    }
                }

                if (comboBox2.Text == "不帶")
                {
                    for (int i = pLC1DataGridView.Rows.Count - 2; i >= 1; i--)
                    {
                        row = pLC1DataGridView.Rows[i];
                        row.Cells["CargoDate2"].Value = "";
                    }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = pLC1DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = pLC1DataGridView.SelectedRows[i];

                row.Cells[17].Value = "True";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = pLC1DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = pLC1DataGridView.SelectedRows[i];

                row.Cells[17].Value = "False";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = pLC1DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = pLC1DataGridView.SelectedRows[i];

                row.Cells["CargoDate2"].Value = textBox1.Text;
            }

        }


        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            DataGridViewRow row;

         
            string a0 = textBox5.Text;
          

                for (int i = pLC1DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {
                    row = pLC1DataGridView.SelectedRows[i];
                        row.Cells["CardName"].Value = a0;
                    }
                

         
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = openFileDialog1.FileName;


                WriteExcelGBPICK3(FileName);

            }


        }
        private void WriteExcelGBPICK3(string ExcelFile)
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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                System.Data.DataTable dt4 = lC.PLC1;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 1]);
                range.Select();
                lcNoTextBox.Text = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 2]);
                range.Select();
                DateTime g1 = Convert.ToDateTime(range.Text.ToString().Trim());
                DateTime g1S = g1.AddDays(1);
                lcDateTextBox.Text = g1.ToString("yyyyMMdd");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 3]);
                range.Select();
                bankNameComboBox.Text = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 4]);
                range.Select();
                bankCodeTextBox.Text = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 5]);
                range.Select();
                lcAmtTextBox.Text = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 6]);
                range.Select();
                DateTime g2 = Convert.ToDateTime(range.Text.ToString().Trim());
                lastDateTextBox.Text = g2.ToString("yyyyMMdd");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, 7]);
                range.Select();
                DateTime g3 = Convert.ToDateTime(range.Text.ToString().Trim());
                expDateTextBox.Text = g3.ToString("yyyyMMdd");
                string DOCENTRY;
                for (int iRecord = 8; iRecord <= iRowCnt - 1; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCENTRY = range.Text.ToString().Trim();

                    System.Data.DataTable dtDOC = GetMenu.Get12(DOCENTRY);
                    if (dtDOC.Rows.Count > 0)
                    {
                        string CARDNAME = dtDOC.Rows[0]["客戶"].ToString();
                        string CARD = dtDOC.Rows[0]["客戶編號"].ToString();
                        string CARNAME = dtDOC.Rows[0]["客戶名稱"].ToString();
                        StringBuilder sb = new StringBuilder();
                        for (int i = 0; i <= dtDOC.Rows.Count - 1; i++)
                        {
                            string DOC = dtDOC.Rows[i]["Docentry"].ToString();
                            string LINENUM = dtDOC.Rows[i]["LINENUM"].ToString();
                            sb.Append("'" + DOC + " " + LINENUM + "',");
                        }

                        sb.Remove(sb.Length - 1, 1);

                        //                DateTime g1S = g1.AddDays(1);
                        System.Data.DataTable dt1 = GetMenu.GetAR22(sb.ToString());
                        System.Data.DataTable dt11 = GetMenu.GetAR32(sb.ToString());
                        System.Data.DataTable dt2 = lC.PLC1;

                        if (bankCodeTextBox.Text == "NTD")
                        {
                            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt1.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "收貨採購";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                drw2["CardName"] = CARDNAME;
                                drw2["CargoDate2"] = g1S.ToString("yyyyMMdd");
                                //CargoDate2
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                decimal taxx3 = 1 + taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                             //   drw2["Amt"] = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) + Convert.ToDecimal(tax)).ToString();
                                drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx3), 0, MidpointRounding.AwayFromZero).ToString();
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = CARD;
                                drw2["CardName2"] = CARNAME;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }
                        else
                        {
                            for (int i = 0; i <= dt11.Rows.Count - 1; i++)
                            {
                                string NumberName = "AS" + DateTime.Now.ToString("yyyyMMdd");
                                DataRow drw = dt11.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                drw2["DocNum"] = docNumTextBox.Text;
                                drw2["LcNo"] = drw["rate"];
                                drw2["PKind"] = "收貨採購";
                                drw2["DonNo"] = drw["DocNum"];
                                drw2["ChNo"] = drw["U_CHI_NO"];
                                drw2["ItemCode"] = drw["ItemCode"];
                                drw2["ItemName"] = drw["Dscription"];
                                drw2["Comments"] = drw["Price"];
                                drw2["LcNo"] = drw["匯率"];
                                drw2["InvoceNo"] = drw["inv"];
                                drw2["CargoDate"] = drw["日期"];
                                drw2["CardName"] = CARDNAME;
                                drw2["CargoDate2"] = g1S.ToString("yyyyMMdd");
                                decimal taxx = Convert.ToDecimal((drw["Vatprcnt"]));
                                decimal taxx2 = taxx / 100;
                                decimal taxx3 = 1 + taxx / 100;
                                string qry = (Convert.ToInt64(drw["Quantity"]) - Convert.ToInt64(drw["QTY"])).ToString();
                                string tax = (Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx2).ToString();
                                drw2["Qty"] = qry;
                                drw2["Price"] = drw["Price"];
                                drw2["TaxCode"] = Convert.ToInt64(drw["Vatprcnt"]);
                                drw2["Amt"] = Math.Round((Convert.ToInt64(qry) * Convert.ToDecimal(drw["Price"]) * taxx3), 2, MidpointRounding.AwayFromZero).ToString();
                                drw2["Tax"] = tax;
                                drw2["Tax"] = tax;
                                drw2["CardCode"] = CARD;
                                drw2["CardName2"] = CARNAME;
                                drw2["LineNum"] = drw["LineNum"];
                                dt2.Rows.Add(drw2);
                            }
                        }
                    }
                }




            }
            finally
            {



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

        private void button11_Click(object sender, EventArgs e)
        {

            System.Data.DataTable GG1 = GETOPEN();

            dataGridView1.DataSource = GG1;
            ExcelReport.GridViewToExcel(dataGridView1);

        }
        private System.Data.DataTable GETOPEN()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.LCNO,lcDate 開狀日期,T0.memo 結案原因,T1.ItemCode 料號,T1.ItemName 品名,T1.QTY 數量,PRICE 價格,TaxCode 稅碼,TAX 稅額");
            sb.Append(" ,T1.LcNo 匯率,Amt 金額,T1.CardName 客戶,T1.InvoceNo,T1.CargoDate 出貨時間,T1.CargoDate2 押匯時間,T1.SendDate 寄出時間,T1.Comments 美金價格 FROM APLC T0");
            sb.Append(" LEFT JOIN PLC1 T1 ON (T0.docNum=T1.docNum)");
            sb.Append(" WHERE lcDate BETWEEN @lcDate1 AND @lcDate2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@lcDate1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@lcDate2", textBox7.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ap_ank1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button12_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            System.Data.DataTable dtCost = MakeTableCombine();
            DataRow dr = null;
            System.Data.DataTable D1 = GETA11(docNumTextBox.Text);
            if (D1.Rows.Count > 0)
            {
                for (int i = 0; i <= D1.Rows.Count - 1; i++)
                {
                    string CARGODATE = D1.Rows[i][0].ToString();
                    System.Data.DataTable dt = GETA12(docNumTextBox.Text, CARGODATE);

                    System.Data.DataTable dt2 = GETA13(docNumTextBox.Text, CARGODATE);
                    string LC = dt2.Rows[0][0].ToString();
                    A1(CARGODATE, LC);

           

                    for (int s = 0; s <= dt.Rows.Count - 1; s++)
                    {
                        DataRow dd = dt.Rows[s];
                        dr = dtCost.NewRow();
                        dr["PO No."] = dd["PO No."].ToString();
                        dr["品名"] = dd["品名"].ToString();
                        dr["數量"] = Convert.ToInt32(dd["數量"]);
                        dr["單價USD"] = Convert.ToDecimal(dd["單價USD"]);
                        dr["5%"] = Convert.ToDecimal(dd["5%"]);
                        dr["金額USD"] = Convert.ToDecimal(dd["金額USD"]);
                        dr["廠商名稱"] = dd["廠商名稱"].ToString();
                        dr["Invoice No."] = dd["Invoice No."].ToString();
                        dr["出貨時間"] = dd["出貨時間"].ToString();
                        dr["押匯時間"] = dd["押匯時間"].ToString();
                        dtCost.Rows.Add(dr);
                    }
                    dr = dtCost.NewRow();
                    dr["PO No."] = "";
                    dr["品名"] = "";
                    dr["數量"] = "0";
                    dr["單價USD"] = "0";
                    dr["5%"] = "0";
                    dr["金額USD"] = Convert.ToDecimal(dt2.Rows[0][0]);
                    dr["廠商名稱"] = "";
                    dr["Invoice No."] = "";
                    dr["出貨時間"] = "";
                    dr["押匯時間"] = "";
                    dtCost.Rows.Add(dr);
                }
                dataGridView7.DataSource = dtCost;



                try
                {


                    DialogResult result;


                    string template;
                    StreamReader objReader;
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\MailTemplates\\LC.html";
                    objReader = new StreamReader(FileName);

                    template = objReader.ReadToEnd();
                    objReader.Close();
                    objReader.Dispose();
                    string F1 = bankNameComboBox.Text + " L/C NO:" + lcNoTextBox.Text + " USD" + lcAmtTextBox.Text;
                    string F2 = "開狀日: " + lcDateTextBox.Text + " 最後交貨日:" + lastDateTextBox.Text + " L/C有效期限:" + expDateTextBox.Text;
                    StringWriter writer = new StringWriter();
                    HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

                    string Html = htmlMessageBody(dataGridView7).ToString();
                    template = template.Replace("##AA##", Html);
                    template = template.Replace("##F1##", F1);
                    template = template.Replace("##F2##", F2);
                    MailMessage message = new MailMessage();


                    message.To.Add(fmLogin.LoginID.ToString() + "@ACMEPOINT.COM");




                    message.Subject = "進貨開狀CR";
                    message.Body = template;
                    string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                    string[] filenames = Directory.GetFiles(OutPutFile);
                    foreach (string file in filenames)
                    {

                        string m_File = "";

                        m_File = file;
                        data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                        //附件资料
                        ContentDisposition disposition = data.ContentDisposition;


                        // 加入邮件附件
                        message.Attachments.Add(data);


                    }


                    message.IsBodyHtml = true;

                    SmtpClient client = new SmtpClient();
                    client.Send(message);
                    data.Dispose();
                    message.Attachments.Dispose();


                    MessageBox.Show("寄信成功");


                }
                catch (Exception ex)
                {
                    DELETEFILE();
                    MessageBox.Show(ex.Message);
                }
            }





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
                //if (iCol == 4)
                //{
                //    strB.AppendLine("<th class='style2'>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
                //else if ( iCol == 5)
                //{
                //    strB.AppendLine("<th class='style3'>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
                //else
                //{
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                //if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                //{

             

                //    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                //    tmpKeyValue = KeyValue;
                //}
                //else
                //{
                //    tmpKeyValue = "";
                //}


                KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                tmpKeyValue = KeyValue;

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

                    string w1 = dg.Rows[i].Cells[1].Value.ToString();
                    string w4 = dg.Rows[i].Cells[4].Value.ToString();
                    if (String.IsNullOrEmpty(w1))
                    {
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()) || dgvc.Value.ToString()=="0")
                        {
                            if (d == 4)
                            {
                                strB.AppendLine("<td align='right'>Total</td>");
                            }
                            else
                            {
                                strB.AppendLine("<td>&nbsp;</td>");
                            }

              
                        }
               
                        else
                        {
                            strB.AppendLine("<td><font color='red'>" + dgvc.Value.ToString() + "</font></td>");
                        }
                    }
                    else
                    {

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
                                    if (dgvc.Value.ToString() == "0")
                                    {
                                        strB.AppendLine("<td>&nbsp;</td>");
                                    }
                                    else
                                    {

                                        string sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                     dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                     dgvc.Value.ToString().Substring(6, 2);


                                        strB.AppendLine("<td>" + sDate + "</td>");
                                    }
                                }
                                else
                                {
                                    strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                                }
                            }
                        }
                    }
                    //}

                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }

        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        }


    }



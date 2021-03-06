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
using Microsoft.VisualBasic.Devices;
using System.Diagnostics;
using System.Net.Mime;
using System.Text.RegularExpressions;

namespace ACME
{
    public partial class ADShip : ACME.fmBase1
    {
        int CON = 0;
        StringBuilder sbS = new StringBuilder();
        Attachment data = null;
        int f2 = 0;
        int f3 = 0;
        string mail = "";
        int CHO1 = 0;
        int CHO2 = 0;
        int CHO3 = 0;
        int COPY = 0;
        int SOL = 0;
        string COMPANY = "";
            string 付款 = "";
            string 離倉日期 = "";
            string 特殊嘜頭 = "";
            string 注意事項 = "";
            string FORWARDER = "";
            string 運輸方式 = "";
            string 貿易條件 = "";
            string shipform = "";
            string shipto = ""; 
            string 付款方式 = "";
            string DIR = "";
            string PATH = "";
            string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string hh = "";
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string DRS = "Data Source=202.3.189.166\\ACMEEB03,50001;Initial Catalog=AcmeSHIDRS;User ID=lleytonchen;Password=aey9919";


        public ADShip()
        {
            InitializeComponent();


        }
        public string PublicString;
        public string PublicString2;
        public string PublicString3;
        
        public override void query()
        {
            button5.Visible = true;
            button5.Enabled = true;
 
            shippingCodeTextBox.ReadOnly = false;
            cardCodeTextBox.ReadOnly = false;
            cardNameTextBox.ReadOnly = false;
            createNameTextBox.ReadOnly = false;
            modifyNameTextBox.ReadOnly = false;
            receivePlaceTextBox.ReadOnly=false;
            goalPlaceTextBox.ReadOnly=false;
            shipmentTextBox.ReadOnly=false;
            unloadCargoTextBox.ReadOnly = false;

 
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                ship.Shipping_main.RejectChanges();
                ship.Shipping_Item.RejectChanges();
                ship.InvoiceM.RejectChanges();
                ship.InvoiceD.RejectChanges();
                ship.PackingListM.RejectChanges();
                ship.PackingListD.RejectChanges();
                ship.LADINGM.RejectChanges();
                ship.LADINGD.RejectChanges();
                ship.Mark.RejectChanges();
                ship.LcInstro.RejectChanges();
                ship.LcInstro1.RejectChanges();
                ship.CFS.RejectChanges();
                ship.Download.RejectChanges();
                ship.Download2.RejectChanges();
                ship.Download3.RejectChanges();
        
            }
            catch
            { 
            }
            return true;
        }

        public override void SAVE()
        {
            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();
            packingListMBindingSource.EndEdit();
            packingListDBindingSource.EndEdit();
            downloadBindingSource.EndEdit();
            download2BindingSource.EndEdit();
            download3BindingSource.EndEdit();
            invoiceMBindingSource.EndEdit();
            invoiceDBindingSource.EndEdit();
            lADINGMBindingSource.EndEdit();
            lADINGDBindingSource.EndEdit();
            cFSBindingSource.EndEdit();
            markBindingSource.EndEdit();
            lcInstroBindingSource.EndEdit();
            lcInstro1BindingSource.EndEdit();


            shipping_mainTableAdapter.Update(ship.Shipping_main);
            shipping_ItemTableAdapter.Update(ship.Shipping_Item);
            invoiceMTableAdapter.Update(ship.InvoiceM);
            invoiceDTableAdapter.Update(ship.InvoiceD);
            packingListMTableAdapter.Update(ship.PackingListM);
            packingListDTableAdapter.Update(ship.PackingListD);
            lADINGMTableAdapter.Update(ship.LADINGM);
            lADINGDTableAdapter.Update(ship.LADINGD);
            cFSTableAdapter.Update(ship.CFS);
            markTableAdapter.Update(ship.Mark);
            downloadTableAdapter.Update(ship.Download);
            download2TableAdapter.Update(ship.Download2);
            download3TableAdapter.Update(ship.Download3);
            lcInstroTableAdapter.Update(ship.LcInstro);
            lcInstro1TableAdapter.Update(ship.LcInstro1);

            ship.Shipping_main.AcceptChanges();
            ship.Shipping_Item.AcceptChanges();
            ship.InvoiceM.AcceptChanges();
            ship.InvoiceD.AcceptChanges();
            ship.PackingListM.AcceptChanges();
            ship.PackingListD.AcceptChanges();
            ship.LADINGM.AcceptChanges();
            ship.LADINGD.AcceptChanges();
            ship.CFS.AcceptChanges();
            ship.Mark.AcceptChanges();
            ship.Download.AcceptChanges();
            ship.Download2.AcceptChanges();
            ship.Download3.AcceptChanges();
            ship.LcInstro.AcceptChanges();
            ship.LcInstro1.AcceptChanges();


            MessageBox.Show("儲存成功");

            付款 = "";
            離倉日期 = "";
            特殊嘜頭 = "";
            注意事項 = "";
            FORWARDER = "";
            運輸方式 = "";
            貿易條件 = "";
            shipform = "";
            shipto = "";
            付款方式 = "";
        }
        public override void EndEdit()
        {
            textBox1.ReadOnly = false;
            button7.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button12.Enabled = true;
            button2.Enabled = true;
            button19.Enabled = true;
            contextMenuStrip2.Enabled = false;
            contextMenuStrip3.Enabled = false;
            contextMenuStrip4.Enabled = false;
            contextMenuStrip5.Enabled = false;
            add7TextBox.ReadOnly = true;
            付款 = "";
            離倉日期 = "";
            特殊嘜頭 = "";
            注意事項 = "";
            FORWARDER = "";
            運輸方式 = "";
            貿易條件 = "";
            shipform = "";
            shipto = "";
            付款方式 = "";

        }
        public override void EndEdit2()
        {
            GetMenu.DELETELOGIN(shippingCodeTextBox.Text);
        }

        public override void STOP2()
        {
            if (globals.GroupID.ToString().Trim() != "EEP")
            {
                System.Data.DataTable L1 = GetLOGIN();
                if (L1.Rows.Count > 0)
                {
                    string H1 = L1.Rows[0][0].ToString();
                    MessageBox.Show("此工單" + H1 + "修改中");
                    this.SSTOPID2 = "1";

                    return;
                }
            }
        }
        public override void STOP()
        {


            if (receiveDayTextBox.Text == "")
            {
                MessageBox.Show("請輸入運送方式");
                this.SSTOPID = "1";
                receiveDayTextBox.Focus();
                return;
            }
            if (boardCountNoTextBox.Text == "")
            {
                MessageBox.Show("請輸入貿易形式");
                this.SSTOPID = "1";
                boardCountNoTextBox.Focus();
                return;

            }

            if (GetDOWNLOAD2().Rows.Count > 0)
            {
                if (GetDOWNLOAD22().Rows.Count == 0)
                {
                    this.SSTOPID = "1";
                    MessageBox.Show("報單號碼與上傳檔案不一致");
                    return;
                }
            }
            CalAMTINVOICE("A");
            if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1349-00")
            {
                if (shipping_ItemDataGridView.Rows.Count > 2)
                {
                    if (CHO1 == 1)
                    {
                 
                        this.SSTOPID = "1";
                        MessageBox.Show("SAP與正航訂單料號不一致");
                        return;
                    }
                    if (CHO2 == 1)
                    {
                        DialogResult result;
                        result = MessageBox.Show("相同料號有兩個單價，請確認SA金額是否正確", "請確認", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No)
                        {
                            this.SSTOPID = "1";
                            return;
                        }
                    }

                    if (CHO3 == 1)
                    {
                        MessageBox.Show("SAP與正航訂單數量不一致");
                        this.SSTOPID = "1";
                        return;
                    }
                }

            }
            if (cardCodeTextBox.Text == "")
            {
                MessageBox.Show("請輸入客戶編號");
                this.SSTOPID = "1";
                cardCodeTextBox.Focus();
                return;
            }
            string S = cardCodeTextBox.Text.Substring(0, 1);

            int n;
            if (int.TryParse(S, out n))
            {
                if (bRANDTextBox.Text == "")
                {
                    MessageBox.Show("請輸入BRAND");
                    this.SSTOPID = "1";
                    bRANDTextBox.Focus();
                    return;
                }
                if (globals.GroupID.ToString().Trim() == "SHI")
                {
                    if (cFSCheckBox.Checked == false)
                    {

                        DialogResult result;
                        result = MessageBox.Show("請確認需不需要保險", "YES/NO", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            this.SSTOPID = "1";
                            cFSCheckBox.Focus();
                            return;
                        }
                    }
                }
            }
        }
        public override void AfterEdit()
        {
          
            bindingNavigator1.Enabled = true;
            bindingNavigator3.Enabled = true;
            bindingNavigator4.Enabled = true;
            bindingNavigator6.Enabled = true;
            button5.Visible = false;


            contextMenuStrip2.Enabled = true;
            contextMenuStrip3.Enabled = true;
            contextMenuStrip4.Enabled = true;
            contextMenuStrip5.Enabled = true;

            shippingCodeTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;

            receivePlaceTextBox.ReadOnly = true;
            goalPlaceTextBox.ReadOnly = true;
            shipmentTextBox.ReadOnly = true;
            unloadCargoTextBox.ReadOnly = true;
            add7TextBox.ReadOnly = true;
            invoiceNoTextBox.ReadOnly = true;
            invoiceNo_seqTextBox.ReadOnly = true;
            pLNoTextBox.ReadOnly = true;
            seqMNoTextBox.ReadOnly = true;

            receiveDayTextBox.ReadOnly = true;
            quantityTextBox.ReadOnly = true;
            boardCountNoTextBox.ReadOnly = true;
            dOCTYPETextBox.ReadOnly = true;
            付款 = "";
            離倉日期 = "";
            特殊嘜頭 = "";
            注意事項 = "";
            FORWARDER = "";
            運輸方式 = "";
            貿易條件 = "";
            shipform = "";
            shipto = "";
            付款方式 = "";

            modifyNameTextBox.Text = fmLogin.LoginID.ToString();
            if (globals.UserID.ToString().Trim().ToUpper() == "KIKILEE")
            {
                modifyNameTextBox.Text = "LilyLee";
            }

            if (globals.GroupID.ToString().Trim() != "EEP")
            {
         

                if (shipping_ItemDataGridView.Rows.Count > 1)
                {
                    shipping_ItemDataGridView.Columns["CURRENCY"].ReadOnly = true;
                    invoiceDDataGridView.Columns["CURRENCY2"].ReadOnly = true;
                    string REMARK = shipping_ItemDataGridView.Rows[0].Cells["ItemRemark"].Value.ToString();
                    int t1 = dOCTYPETextBox.Text.IndexOf("調撥");
                    if (t1 == -1)
                    {
                        if (REMARK == "銷售訂單" || REMARK == "Choice" || REMARK == "Infinite" || REMARK == "TOP GARDEN")
                        {


                            invoiceDDataGridView.Columns["dataGridViewTextBoxColumn34"].ReadOnly = true;
                            shipping_ItemDataGridView.Columns["ItemPrice"].ReadOnly = true;
                            shipping_ItemDataGridView.Columns["CHOPrice1"].ReadOnly = true;
                            invoiceDDataGridView.Columns["CHOPrice"].ReadOnly = true;


                        }
                        else
                        {
                            invoiceDDataGridView.Columns["dataGridViewTextBoxColumn34"].ReadOnly = false;

                            shipping_ItemDataGridView.Columns["ItemPrice"].ReadOnly = false;
                        }
                    }
                }
            }
     


        
        }

        public override void AfterAddNew()
        {
        
            bindingNavigator1.Enabled = true;
            bindingNavigator3.Enabled = true;
            bindingNavigator4.Enabled = true;
            bindingNavigator6.Enabled = true;
  
            nTDollarsTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");
            textBox1.ReadOnly = false;
            button7.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button2.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button12.Enabled = true;
            button19.Enabled = true;
            shippingCodeTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            receivePlaceTextBox.ReadOnly = true;
            goalPlaceTextBox.ReadOnly = true;
            shipmentTextBox.ReadOnly = true;
            unloadCargoTextBox.ReadOnly = true;
            receiveDayTextBox.ReadOnly = true;
            quantityTextBox.ReadOnly = true;
            boardCountNoTextBox.ReadOnly = true;
            dOCTYPETextBox.ReadOnly = true;
            invoiceNoTextBox.ReadOnly = true;
            invoiceNo_seqTextBox.ReadOnly = true;
            pLNoTextBox.ReadOnly = true;
            seqMNoTextBox.ReadOnly = true;
            add7TextBox.ReadOnly = true;

            tabControl1.SelectedIndex = 0;
        }

        public override void AfterCancelEdit()
        {

            GetMenu.DELETELOGIN(shippingCodeTextBox.Text);

            bindingNavigator1.Enabled = false;
            bindingNavigator3.Enabled = false;
            bindingNavigator4.Enabled = false;
            bindingNavigator6.Enabled = false;
            button5.Visible = false;


            contextMenuStrip2.Enabled = false;
            contextMenuStrip3.Enabled = false;
            contextMenuStrip4.Enabled = false;
            contextMenuStrip5.Enabled = false;


            shippingCodeTextBox.ReadOnly = true;

            receiveDayTextBox.ReadOnly = true;
            quantityTextBox.ReadOnly = true;
            boardCountNoTextBox.ReadOnly = true;
            dOCTYPETextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            receivePlaceTextBox.ReadOnly = true;
            goalPlaceTextBox.ReadOnly = true;
            shipmentTextBox.ReadOnly = true;
            unloadCargoTextBox.ReadOnly = true;

            cardCodeTextBox.ReadOnly = true;
            cardNameTextBox.ReadOnly = true;
            invoiceNoTextBox.ReadOnly = true;
            invoiceNo_seqTextBox.ReadOnly = true;
            pLNoTextBox.ReadOnly = true;
            seqMNoTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            nTDollarsTextBox.ReadOnly = true;
            dollarsKindTextBox.ReadOnly = true;
            add7TextBox.ReadOnly = true;


            button7.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;

            button12.Enabled = true;
            textBox1.ReadOnly = false;
            button2.Enabled = true;
            button19.Enabled = true;
            付款 = "";
            離倉日期 = "";
            特殊嘜頭 = "";
            注意事項 = "";
            FORWARDER = "";
            運輸方式 = "";
            貿易條件 = "";
            shipform = "";
            shipto = "";
            付款方式 = "";

        }
        public override void AfterEndEdit()
        {
            COPY = 0;

            bindingNavigator1.Enabled = false;
            bindingNavigator3.Enabled = false;
            bindingNavigator4.Enabled = false;
            bindingNavigator6.Enabled = false;
            button5.Visible = false;

         

        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
          
            shipping_mainTableAdapter.Connection = MyConnection;
            shipping_ItemTableAdapter.Connection = MyConnection;
            cFSTableAdapter.Connection=MyConnection;
            markTableAdapter.Connection = MyConnection;
            invoiceDTableAdapter.Connection =MyConnection;
            invoiceMTableAdapter.Connection = MyConnection;
            lADINGMTableAdapter.Connection = MyConnection;
            lADINGDTableAdapter.Connection = MyConnection;
            packingListMTableAdapter.Connection = MyConnection;
            packingListDTableAdapter.Connection = MyConnection;
            lcInstroTableAdapter.Connection = MyConnection;
            lcInstro1TableAdapter.Connection = MyConnection;
            downloadTableAdapter.Connection = MyConnection;
            download2TableAdapter.Connection = MyConnection;
            download3TableAdapter.Connection = MyConnection;
        }
        public override void SetInit()
        {

            MyBS = shipping_mainBindingSource;
            MyTableName = "Shipping_main";
            MyIDFieldName = "ShippingCode";

           UtilSimple.SetLookupBinding(shipToDateComboBox, "shipToDate", shipping_mainBindingSource, "shipToDate");
          // UtilSimple.SetLookupBinding(add7ComboBox, "add7", shipping_mainBindingSource, "add7");

           //處理複製
           MasterTable = ship.Shipping_main;
           DetailTables = new System.Data.DataTable[] { ship.Shipping_Item };
           DetailBindingSources = new BindingSource[] { shipping_ItemBindingSource };
        }

        public override void FillData()
        {
            try
            {
                if (!String.IsNullOrEmpty(PublicString))
                {
                    MyID = PublicString.Trim();
                }

                if (!String.IsNullOrEmpty(PublicString2))
                {
                    MyID = PublicString2.Trim();
                    tabControl1.SelectedTab = 可下載檔案;
                }
                shipping_mainTableAdapter.Fill(ship.Shipping_main, MyID);
                shipping_ItemTableAdapter.Fill(ship.Shipping_Item, MyID);
                cFSTableAdapter.Fill(ship.CFS, MyID);
                markTableAdapter.Fill(ship.Mark,MyID);
            
                invoiceMTableAdapter.Fill(ship.InvoiceM, MyID);
                invoiceDTableAdapter.Fill(ship.InvoiceD, MyID);
                lADINGMTableAdapter.Fill(ship.LADINGM, MyID);
                lADINGDTableAdapter.Fill(ship.LADINGD, MyID);
                packingListMTableAdapter.Fill(ship.PackingListM, MyID);
                packingListDTableAdapter.Fill(ship.PackingListD, MyID);
                downloadTableAdapter.Fill(ship.Download, MyID);
                download2TableAdapter.Fill(ship.Download2, MyID);
                download3TableAdapter.Fill(ship.Download3, MyID);
                lcInstroTableAdapter.Fill(ship.LcInstro, MyID);
                lcInstro1TableAdapter.Fill(ship.LcInstro1, MyID);

 

                System.Data.DataTable K2 = GetFEE(shippingCodeTextBox.Text);
                System.Data.DataTable K3 = GetSHPCAR(shippingCodeTextBox.Text);
                System.Data.DataTable K4 = GetSHPCAR2(shippingCodeTextBox.Text);
                System.Data.DataTable K5 = GetSHPCAR3(shippingCodeTextBox.Text);
                dataGridView2.DataSource = K2;
                dataGridView3.DataSource = K3;
                dataGridView4.DataSource = K4;
                dataGridView5.DataSource = K5;
                System.Data.DataTable INVO = GetINVO(shippingCodeTextBox.Text);
                textBox3.Text = INVO.Compute("Sum(AMOUNT)", null).ToString();
                 System.Data.DataTable INVO2 = Getfee(shippingCodeTextBox.Text);
                 textBox6.Text = INVO2.Compute("sum(Amount)", null).ToString();
                System.Data.DataTable AP = GetAP(shippingCodeTextBox.Text);

                StringBuilder sb2 = new StringBuilder();

                textBox4.Text = "";
                if (AP.Rows.Count > 0)
                {
                    for (int i = 0; i <= AP.Rows.Count - 1; i++)
                    {
                        DataRow dd = AP.Rows[i];
                        sb2.Append(dd["SHIP"].ToString() + System.Environment.NewLine);
                    }
                    textBox4.Text = sb2.ToString();
                }

                textBox5.Text = "";
                if (ship.Shipping_Item.Rows.Count > 0)
                {
                    if (ship.Shipping_Item.Rows[0]["ItemRemark"].ToString() == "採購訂單")
                    {
                        System.Data.DataTable SALES = GetSALES(ship.Shipping_Item.Rows[0]["Docentry"].ToString());
                        if (SALES.Rows.Count > 0)
                        {

                            textBox5.Text = SALES.Rows[0]["業務"].ToString();
                        }

                        
                        
                    }
                }

                SHIPNO();
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
                try
                {

                    shipping_ItemBindingSource.MoveFirst();

                    for (int i = 0; i <= shipping_ItemBindingSource.Count - 1; i++)
                    {
                        DataRowView row3 = (DataRowView)shipping_ItemBindingSource.Current;

                        row3["SeqNo"] = i;

                        shipping_ItemBindingSource.EndEdit();

                        shipping_ItemBindingSource.MoveNext();
                    }

                    UPINVOICE();

                    UPPACK();
                  

                    downloadBindingSource.MoveFirst();

                    for (int i = 0; i <= downloadBindingSource.Count - 1; i++)
                    {
                        DataRowView rowd = (DataRowView)downloadBindingSource.Current;

                        rowd["seq"] = i;



                        downloadBindingSource.EndEdit();

                        downloadBindingSource.MoveNext();
                    }


                    download2BindingSource.MoveFirst();

                    for (int i = 0; i <= download2BindingSource.Count - 1; i++)
                    {
                        DataRowView row1 = (DataRowView)download2BindingSource.Current;

                        row1["seq"] = i;



                        download2BindingSource.EndEdit();

                        download2BindingSource.MoveNext();
                    }

                    download3BindingSource.MoveFirst();

                    for (int i = 0; i <= download3BindingSource.Count - 1; i++)
                    {
                        DataRowView row1 = (DataRowView)download3BindingSource.Current;

                        row1["seq"] = i;



                        download3BindingSource.EndEdit();

                        download3BindingSource.MoveNext();
                    }

                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);

                }


                shipping_mainTableAdapter.Connection.Open();


                shipping_mainBindingSource.EndEdit();
                shipping_ItemBindingSource.EndEdit();
                packingListMBindingSource.EndEdit();
                packingListDBindingSource.EndEdit();
                downloadBindingSource.EndEdit();
                download2BindingSource.EndEdit();
                download3BindingSource.EndEdit();
                invoiceMBindingSource.EndEdit();
                invoiceDBindingSource.EndEdit();
                lADINGMBindingSource.EndEdit();
                lADINGDBindingSource.EndEdit();
                cFSBindingSource.EndEdit();
                markBindingSource.EndEdit();
                lcInstroBindingSource.EndEdit();
                lcInstro1BindingSource.EndEdit();

                tx = shipping_mainTableAdapter.Connection.BeginTransaction();

                SqlDataAdapter Adapter = util.GetAdapter(shipping_mainTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(shipping_ItemTableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(invoiceMTableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter3 = util.GetAdapter(invoiceDTableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter4 = util.GetAdapter(packingListMTableAdapter);
                Adapter4.UpdateCommand.Transaction = tx;
                Adapter4.InsertCommand.Transaction = tx;
                Adapter4.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter5 = util.GetAdapter(packingListDTableAdapter);
                Adapter5.UpdateCommand.Transaction = tx;
                Adapter5.InsertCommand.Transaction = tx;
                Adapter5.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter6 = util.GetAdapter(lADINGMTableAdapter);
                Adapter6.UpdateCommand.Transaction = tx;
                Adapter6.InsertCommand.Transaction = tx;
                Adapter6.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter7 = util.GetAdapter(lADINGDTableAdapter);
                Adapter7.UpdateCommand.Transaction = tx;
                Adapter7.InsertCommand.Transaction = tx;
                Adapter7.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter8 = util.GetAdapter(cFSTableAdapter);
                Adapter8.UpdateCommand.Transaction = tx;
                Adapter8.InsertCommand.Transaction = tx;
                Adapter8.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter9 = util.GetAdapter(markTableAdapter);
                Adapter9.UpdateCommand.Transaction = tx;
                Adapter9.InsertCommand.Transaction = tx;
                Adapter9.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter10 = util.GetAdapter(downloadTableAdapter);
                Adapter10.UpdateCommand.Transaction = tx;
                Adapter10.InsertCommand.Transaction = tx;
                Adapter10.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter11 = util.GetAdapter(download2TableAdapter);
                Adapter11.UpdateCommand.Transaction = tx;
                Adapter11.InsertCommand.Transaction = tx;
                Adapter11.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter12 = util.GetAdapter(lcInstroTableAdapter);
                Adapter12.UpdateCommand.Transaction = tx;
                Adapter12.InsertCommand.Transaction = tx;
                Adapter12.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter13 = util.GetAdapter(lcInstro1TableAdapter);
                Adapter13.UpdateCommand.Transaction = tx;
                Adapter13.InsertCommand.Transaction = tx;
                Adapter13.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter14 = util.GetAdapter(download3TableAdapter);
                Adapter14.UpdateCommand.Transaction = tx;
                Adapter14.InsertCommand.Transaction = tx;
                Adapter14.DeleteCommand.Transaction = tx;

                shipping_mainTableAdapter.Update(ship.Shipping_main);
                shipping_ItemTableAdapter.Update(ship.Shipping_Item);
                invoiceMTableAdapter.Update(ship.InvoiceM);
                invoiceDTableAdapter.Update(ship.InvoiceD);
                packingListMTableAdapter.Update(ship.PackingListM);
                packingListDTableAdapter.Update(ship.PackingListD);
                lADINGMTableAdapter.Update(ship.LADINGM);
                lADINGDTableAdapter.Update(ship.LADINGD);
                cFSTableAdapter.Update(ship.CFS);
                markTableAdapter.Update(ship.Mark);
                downloadTableAdapter.Update(ship.Download);
                download2TableAdapter.Update(ship.Download2);
                download3TableAdapter.Update(ship.Download3);
                lcInstroTableAdapter.Update(ship.LcInstro);
                lcInstro1TableAdapter.Update(ship.LcInstro1);

                ship.Shipping_main.AcceptChanges();
                ship.Shipping_Item.AcceptChanges();
                ship.InvoiceM.AcceptChanges();
                ship.InvoiceD.AcceptChanges();
                ship.PackingListM.AcceptChanges();
                ship.PackingListD.AcceptChanges();
                ship.LADINGM.AcceptChanges();
                ship.LADINGD.AcceptChanges();
                ship.CFS.AcceptChanges();
                ship.Mark.AcceptChanges();
                ship.Download.AcceptChanges();
                ship.Download2.AcceptChanges();
                ship.Download3.AcceptChanges();
                ship.LcInstro.AcceptChanges();
                ship.LcInstro1.AcceptChanges();
                
                this.MyID = this.shippingCodeTextBox.Text;
                 tx.Commit();


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
                this.shipping_mainTableAdapter.Connection.Close();
               
            }
            return UpdateData;
        }


        public override void SetDefaultValue()
        {
            
            if (kyes == null)
            {

                string NumberName = "SH" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

                kyes = NumberName + AutoNum + "L";
            }

   
            
            this.shippingCodeTextBox.Text = kyes;
            this.nTDollarsTextBox.Text = "0";
            kyes = this.shippingCodeTextBox.Text;

            createNameTextBox.Text = fmLogin.LoginID.ToString().Trim();
        
            System.Data.DataTable J1 = GETOHEM(fmLogin.LoginID.ToString().Trim());
            if (J1.Rows.Count > 0)
            {
                add7TextBox.Text = J1.Rows[0][0].ToString();
            }
            if (globals.UserID.ToString().Trim().ToUpper() == "KIKILEE")
            {
                createNameTextBox.Text = "LilyLee";
                add7TextBox.Text = "LilyLee";
            }
            iTEMSCheckBox.Checked = false;
           cFSCheckBox.Checked = false;
           iNSUCHECKCheckBox.Checked = false;
           buCardcodeCheckBox.Checked = false;
           add10CheckBox.Checked = false;
           tAXCHECKCheckBox.Checked = false;
            this.shipping_mainBindingSource.EndEdit();
            kyes = null;
            quantityTextBox.Text = "未結";


        }

        public override void AfterCopy()
        {
            
            if (kyes == null)
            {
                string NumberName = "SH" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.shippingCodeTextBox.Text = NumberName + AutoNum + "L";
                kyes = this.shippingCodeTextBox.Text;
            }
        }
        public override void AfterCopy2()
        {
            COPY = 1;

            tabControl1.SelectedIndex = 0;
            add2TextBox.Text = "";
            add6TextBox.Text = "";
            closeDayTextBox.Text ="";
            forecastDayTextBox.Text ="";
            arriveDayTextBox.Text = "";
            boatNameTextBox.Text = "";
            boatCompanyTextBox.Text = "";
            boardCountTextBox.Text = "";
            boardDeliverTextBox.Text = "";
            sendGoodsTextBox.Text = "";
            modifyDateTextBox.Text = "";
            cFSCheckBox.Text = "";
            buCardnameTextBox.Text = "";
            soNoTextBox.Text = "";
            add9TextBox.Text = "";
            shipping_OBUTextBox.Text = "";
            shipToDateComboBox.Text = "";
            System.Data.DataTable J1 = GETOHEM(fmLogin.LoginID.ToString().Trim());
            if (J1.Rows.Count > 0)
            {
                add7TextBox.Text = J1.Rows[0][0].ToString();
             
            }
            createNameTextBox.Text = fmLogin.LoginID.ToString().Trim();
            if (globals.UserID.ToString().Trim().ToUpper() == "KIKILEE")
            {
                createNameTextBox.Text = "LilyLee";
                add7TextBox.Text = "LilyLee";
            }
            modifyNameTextBox.Text = "";
            nTDollarsTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");

            buCardcodeCheckBox.Checked = false;
            quantityTextBox.Text = "未結";
            iTEMSCheckBox.Checked = false;
            add10CheckBox.Checked = false;

        }
        private void shipping_ItemDataGridView_DefaultValuesNeeded_1(object sender, DataGridViewRowEventArgs e)
        {

            int iRecs;

            iRecs = shipping_ItemDataGridView.Rows.Count - 1;
            e.Row.Cells["SeqNo"].Value = iRecs.ToString();
            e.Row.Cells["ItemPrice"].Value = 1;
            e.Row.Cells["Quantity"].Value = 0;
            e.Row.Cells["ItemAmount"].Value = 0;
            e.Row.Cells["CHOPrice1"].Value = 0;
            e.Row.Cells["CHOAmount1"].Value = 0;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            LookupValues = GetMenu.GetCHI5();

            if (LookupValues != null)
            {
                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);
            }
           
        }

        private void shipping_ItemDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
                   shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemPrice")
                {
                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;

                    iQuantity = Convert.ToInt32(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    iUnitPrice = Convert.ToDecimal(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemPrice"].Value);
                    this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemAmount"].Value = (iQuantity * iUnitPrice).ToString();

                }

                if (shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "CHOPrice1" ||
                    shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Quantity")
                {
                    decimal iQuantity = 0;
                    decimal CHOPrice1 = 0;

                    iQuantity = Convert.ToInt32(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    CHOPrice1 = Convert.ToDecimal(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["CHOPrice1"].Value);
                    this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["CHOAmount1"].Value = (iQuantity * CHOPrice1).ToString();

                }



                if (shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemCode" ||
       shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Dscription" ||
                       shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
        shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemPrice" ||
                        shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemAmount" ||
                        shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Remark"
       )
                {
                    this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["RED2"].Value = "True";
                }
            }
            catch { 
            
            }

        }


        private void button5_Click(object sender, EventArgs e)
        {
           

            fmInvo frm1 = new fmInvo();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                shippingCodeTextBox.Text = frm1.Invo;
            }
        }

        private void fmShip_Load(object sender, EventArgs e)
        {
            string USER = fmLogin.LoginID.ToString();
      
            bindingNavigator1.Enabled = false;
            bindingNavigator3.Enabled = false;
            bindingNavigator4.Enabled = false;
            bindingNavigator6.Enabled = false;
            bindingNavigator1.Visible = true;
            bindingNavigator3.Visible= true;
            bindingNavigator4.Visible = true;
            bindingNavigator6.Visible = true;

          contextMenuStrip2.Enabled = false;
          contextMenuStrip3.Enabled = false;
          contextMenuStrip4.Enabled = false;
          contextMenuStrip5.Enabled = false;
          button7.Enabled = true;
          button13.Enabled = true;
          button11.Enabled = true;
          checkBox3.Enabled = true;
          checkBox5.Enabled = true;
          button10.Enabled = true;
          button3.Enabled = true;
          button12.Enabled = true;
          textBox1.ReadOnly = false;
          button2.Enabled = true;
          add6TextBox.ReadOnly = true;
          button19.Enabled = true;
          textBox2.Text = USER + "@acmepoint.com";
          textBox1.Text = USER + "@acmepoint.com";
    
 
          ExcelReport.DELETEFILE();
          ExcelReport.DELETEFOLDER();
          string GROUP = globals.GroupID.ToString().Trim();
          if (globals.UserID.ToString().Trim().ToUpper() == "KIKILEE")
          {
              textBox1.Text = "lilylee@adlab.com.tw";
              textBox2.Text = "lilylee@adlab.com.tw";
          }
          if (GROUP != "EEP" && GROUP != "SHI" && GROUP != "ShipBuy" && GROUP != "WH")
          {   
              lcInstro1DataGridView.Columns["LPRICE"].Visible = false;
              lcInstro1DataGridView.Columns["LAMT"].Visible = false;
          }
          if (GROUP != "EEP")
          {
              textBox9.Visible = false;
              textBox10.Visible = false;
              textBox11.Visible = false;
              textBox12.Visible = false;
          }

    
          DIR = "//acmesrv01//SAP_Share//shipping宇豐//";
          PATH = @"\\acmesrv01\SAP_Share\shipping宇豐\";

          //shippingAD
          shipping_ItemDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;



            System.Data.DataTable T1 = GetWHSA();
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }



        }


        private void bindingNavigatorAddNewItem3_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dtm = GetMenu.getaa(shippingCodeTextBox.Text);

            if (dtm.Rows.Count.ToString() == "0")
            {
                MessageBox.Show("請先儲存主檔");
                bindingNavigator3.Enabled = false;
            }

            System.Data.DataTable dt1 = GetMenu.Getinvoicem(shippingCodeTextBox.Text);
             System.Data.DataTable dtt = GetMenu.GetPacking(shippingCodeTextBox.Text);
             System.Data.DataTable dtt2 = GetMenu.GetPacking2(shippingCodeTextBox.Text);
            try
            {
                int INVO = dt1.Rows.Count;
                int PACK = dtt.Rows.Count;
                if (INVO < 1)
                {
                    packingListMBindingSource.RemoveCurrent();
                    packingListMBindingSource.EndEdit();
                    MessageBox.Show("請輸入invoice單號");
                    return;
                }

                if (INVO == PACK)
                {
                    packingListMBindingSource.RemoveCurrent();
                    packingListMBindingSource.EndEdit();
                    MessageBox.Show("請輸入invoice單號");
                    return;
                  
                }

                int i = dtt2.Rows.Count;
                DataRow drw = dt1.Rows[i];

                pDateTextBox.Text = DateTime.Now.ToString("yyyyMMdd");

                string invoiceno = drw["InvoiceNo"].ToString();
                string InvoiceNo_seq = drw["InvoiceNo_seq"].ToString();
                string aa = invoiceno + "-" + InvoiceNo_seq;


                pLNoTextBox.Text = aa;

                        bill_ToTextBox.Text = drw["BillTo"].ToString();
                        shippedByTextBox.Text = drw["ShipTo"].ToString();
                        oBUBillToTextBox.Text = drw["OBUBillTo"].ToString();
                        oBUShipToTextBox.Text = drw["OBUShipTo"].ToString();
                  
                        if (shipmentTextBox.Text != "")
                        {
                            shipping_FromTextBox.Text = shipmentTextBox.Text;
                        }

                        if (receiveDayTextBox.Text != "")
                        {
                            forAccountTextBox.Text = receiveDayTextBox.Text;
                        }
                     
                    
                        if (unloadCargoTextBox.Text != "")
                        {
                            shipping_ToTextBox.Text = unloadCargoTextBox.Text;
                        }

                        if (boatNameTextBox.Text != "")
                        {
                            shipping_PerTextBox.Text = boatNameTextBox.Text;
                        }

                        if (closeDayTextBox.Text != "")
                        {
                            shippedOnTextBox.Text = closeDayTextBox.Text;
                        }

                            string DOCTYPE = dOCTYPETextBox.Text;
                string OUTTYPE = boardCountNoTextBox.Text;

                if (((DOCTYPE == "銷售" && OUTTYPE == "出口") || (DOCTYPE == "銷售" && OUTTYPE == "三角") || DOCTYPE == "調撥單" || DOCTYPE == "發貨單") && mEMO3TextBox.Text != "")
                {

                    string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                    StringBuilder sb = new StringBuilder();

                    string ss = "";
                    foreach (string S in arrurl)
                    {


                        System.Data.DataTable DF1 = GetMenu.getCAR(S);
                        if (DF1.Rows.Count > 0)
                        {
                            SHIPCAR frm1 = new SHIPCAR();
                            frm1.SHIPPINGCODE = S;
                            if (frm1.ShowDialog() == DialogResult.OK)
                            {
                                ss = frm1.a.ToString();

                            }
                        }
                    }
                    GETPACK(InvoiceNo_seq, ss, toolStripComboBox1.Text);

                }
                else
                {
                    System.Data.DataTable dt3 = GetMenu.Getshipinvo(shippingCodeTextBox.Text, invoiceno, InvoiceNo_seq);

                    System.Data.DataTable dt4 = ship.PackingListD;

                    if (invoiceDDataGridView.Rows.Count > 1 && packingListDDataGridView.Rows.Count < 2)
                    {

                        for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                        {
                            DataRow drw3 = dt3.Rows[j];
                            DataRow drw2 = dt4.NewRow();
                            drw2["ShippingCode"] = shippingCodeTextBox.Text;
                            drw2["plno"] = pLNoTextBox.Text;
                            drw2["seqno"] = j;
                            drw2["DescGoods"] = drw3["INDescription"];
                            drw2["TREETYPE"] = drw3["TREETYPE"];
                            drw2["VISORDER"] = drw3["VISORDER"];
                            drw2["SOID"] = drw3["SOID"];

                            dt4.Rows.Add(drw2);
                        }
                    }

                }
                            try
                            {

                                this.packingListMBindingSource.EndEdit();
                                this.packingListMTableAdapter.Update(ship.PackingListM);
                                ship.PackingListM.AcceptChanges();

                                this.packingListDBindingSource.EndEdit();
                                this.packingListDTableAdapter.Update(ship.PackingListD);
                                ship.PackingListD.AcceptChanges();
                                
                            }
                            catch (Exception ex)
                            {

                                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "PackingTran1", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);

                      
                            }
                          

       
            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Pack新增", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

            }
           
        }
        public void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }
        public void DELPACK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete SHIPPING_PACK where users=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        private void WriteExcelPACK(string SEQ, string CHE, string CAR, string CHOSHIP)
        {

            DELPACK();

            int SQ = Convert.ToInt16(SEQ);
            int GS = mEMO3TextBox.Text.IndexOf(",");
            string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
            StringBuilder sb = new StringBuilder();
            string SHIPPINGCODE = "";
            int M1 = 0;
            foreach (string i in arrurl)
            {
                if (GS != -1)
                {
                    M1++;
                    if (M1 == SQ)
                    {
                        SHIPPINGCODE = i;
                    }
                }
                else
                {
                    SHIPPINGCODE = i;
                }


                sb.Append("'" + i + "',");
            }
            sb.Remove(sb.Length - 1, 1);
            string BLC = "";
            if (toolStripComboBox1.Text != "")
            {
                SHIPPINGCODE = toolStripComboBox1.Text;
            }
            System.Data.DataTable dtBLC = GetWHPACKBLC(SHIPPINGCODE, SQ);
            if (dtBLC.Rows.Count > 0)
            {
                BLC = dtBLC.Rows[0][0].ToString();
            }
            System.Data.DataTable dt3 = GetWHPACK(SHIPPINGCODE, BLC, CHE, sb.ToString(), CAR);
            if (dt3.Rows.Count == 0)
            {
                MessageBox.Show("包裝明細無資料");
                return;
            }
            if (dt3.Rows.Count > 0)
            {
                string PLATENO;
                string PLATENOE1 = "";
                string PLATENOE2 = "";
                string CARTONNO;
                string ITEMCODE;
                string QTY;
                string CARTONQTY;
                string NW;
                string GW;
                string L;
                string W;
                string H;
                string LOACTION;
                string PLATENO2 = "";
                string L2 = "";
                string W2 = "";
                string H2 = "";
                string MITEM = "";
                string GW2 = "";
                string INVOICE = "";
                string ITEMNAME = "";
                string WHNO = "";
                string WHNOD = "";
                int SER = 0;
                int SER2 = 0;
                string SERX = "";
                int CARTONNO2 = 0;
                string ES;
                string CARTONNO3 = "";
                string CARTONNO5 = "";
                int SER3 = 0;

                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];


                    WHNO = drw3["SHIPPINGCODE"].ToString().Trim();
                    PLATENO = drw3["PLATENO"].ToString().Trim();
                    PLATENO2 = drw3["PLATENO2"].ToString().Trim();
                    CARTONNO = drw3["CARTONNO"].ToString().Trim();
                    ITEMCODE = drw3["ITEMCODE"].ToString().Trim();
                    ITEMNAME = drw3["ITEMNAME"].ToString().Trim();
                    QTY = drw3["QTY"].ToString().Trim();
                    CARTONQTY = drw3["CARTONQTY"].ToString().Trim();
                    NW = drw3["NW"].ToString().Trim();
                    GW = drw3["GW2"].ToString().Trim();
                    L = drw3["L"].ToString().Trim();
                    W = drw3["W"].ToString().Trim();
                    H = drw3["H"].ToString().Trim();

                    if (j == 0)
                    {
                        PLATENOE1 = PLATENO;
                    }
                    if (j == 1)
                    {
                        PLATENOE2 = PLATENO;
                    }
                    CARTONNO5 = CARTONNO;
                    if (!String.IsNullOrEmpty(PLATENO2))
                    {
                        System.Data.DataTable H1 = GetSHIPPACK9(WHNO, PLATENO2);
                        if (H1.Rows.Count > 0)
                        {
                            CARTONNO5 = H1.Rows[0][0].ToString();
                        }
                    }
                    LOACTION = drw3["LOACTION"].ToString().Trim();
                    INVOICE = drw3["AUNO"].ToString().Trim();
                    ES = drw3["ES"].ToString().Trim();
                    if (QTY == "空箱")
                    {
                        QTY = "0";
                        ITEMCODE = "空箱";
                    }

                    int CARTONNO4 = 0;
                    if (WHNOD != WHNO)
                    {
                        CARTONNO2 = 0;
                    }
                    if (cardCodeTextBox.Text == "1362-00")
                    {
                        if (PLATENOE1 == "1" && PLATENOE2 == "1")
                        {

                        }
                        else
                        {
                            if (PLATENO == "1")
                            {
                                CARTONNO2 = 0;
                            }
                        }
                    }
                    if (!String.IsNullOrEmpty(ITEMCODE))
                    {
                        if (ITEMCODE != "空箱")
                        {
                            CARTONNO4 = CARTONNO2 + 1;
                            if (String.IsNullOrEmpty(CARTONNO))
                            {
                                CARTONNO = "0";
                            }
                            CARTONNO2 += Convert.ToInt16(CARTONNO);
                        }

                        //if (ITEMCODE == "M270DAN02.55QA2")
                        //{
                        //    MessageBox.Show("A");
                        //}

                        string F1 = CARTONNO5 + ITEMCODE + GW + L + W + H + QTY;
                        if ((CARTONNO5 + ITEMCODE + GW + L + W + H + QTY != MITEM) || (String.IsNullOrEmpty(L)))
                        {
                            SERX = SER2.ToString();
                            SER3 = 0;
                        }
                        else
                        {
                            if (SER3 == 0)
                            {
                                SER++;
                                SERX = SER.ToString();
                                UPPACKS(SERX);
                                SER3 = 1;
                            }

                        }
                        if (!String.IsNullOrEmpty(PLATENO2))
                        {
                            MITEM = CARTONNO5 + ITEMCODE + GW + L + W + H + QTY;
                        }
                        else
                        {
                            MITEM = CARTONNO + ITEMCODE + GW + L + W + H + QTY;
                        }


                        CARTONNO3 = CARTONNO4.ToString().Trim() + "~" + CARTONNO2.ToString().Trim();
                        if (CARTONNO == "1")
                        {
                            CARTONNO3 = CARTONNO4.ToString();
                        }
                        if (CARTONNO == "0")
                        {
                            CARTONNO3 = "";
                        }
                        if (String.IsNullOrEmpty(NW))
                        {
                            System.Data.DataTable G1 = GetSHIPPACK6(ITEMCODE, QTY);
                            if (G1.Rows.Count > 0)
                            {
                                NW = G1.Rows[0][0].ToString();
                            }
                            else
                            {
                                System.Data.DataTable G2 = GetSHIPPACK7(ITEMCODE);
                                if (G2.Rows.Count > 0)
                                {
                                    string PAL_NW = G2.Rows[0]["PAL_NW"].ToString();
                                    string PAL_QTY = G2.Rows[0]["PAL_QTY"].ToString();

                                    decimal n;
                                    if (decimal.TryParse(PAL_NW, out n) && decimal.TryParse(PAL_QTY, out n) && decimal.TryParse(QTY, out n))
                                    {
                                        NW = ((Convert.ToDecimal(PAL_NW) / Convert.ToDecimal(PAL_QTY)) * Convert.ToDecimal(QTY)).ToString("#,##0.000");
                                    }
                                }
                            }
                        }
                        else
                        {
                            NW = Convert.ToDecimal(NW).ToString("0.000");
                        }
                        AddPACK(PLATENO, CARTONNO, ITEMCODE, QTY, CARTONQTY, NW, GW, L, W, H, LOACTION, SERX, CARTONNO3, INVOICE, ITEMNAME, WHNO, ES);
                    }

                    WHNOD = WHNO;
                }
            }

            if (SEQ == "01")
            {
                System.Data.DataTable dt3H = GetWHPACKH(sbS.ToString());

                if (dt3H.Rows.Count > 0)
                {
                    DELMARK();

                    for (int j = 0; j <= dt3H.Rows.Count - 1; j++)
                    {
                        string MARK = dt3H.Rows[j][0].ToString();
                        InsertMARK(shippingCodeTextBox.Text, j.ToString(), MARK);
                    }

                    this.markBindingSource.EndEdit();
                    this.markTableAdapter.Update(ship.Mark);
                    ship.Mark.AcceptChanges();
                    markTableAdapter.Fill(ship.Mark, MyID);
                }
            }

        }
        public void AddPACK(string PLATENO, string CARTONNO, string ITEMCODE, string QTY, string CARTONQTY, string NW, string GW, string L, string W, string H, string LOACTION, string SER, string CARTONNO2, string INVOICE, string ITEMNAME, string WHNO, string ES)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" Insert into SHIPPING_PACK(PLATENO,CARTONNO,ITEMCODE,QTY,CARTONQTY,NW,GW,L,W,H,LOACTION,USERS,SER,CARTONNO2,INVOICE,ITEMNAME,WHNO,ES) values(@PLATENO,@CARTONNO,@ITEMCODE,@QTY,@CARTONQTY,@NW,@GW,@L,@W,@H,@LOACTION,@USERS,@SER,@CARTONNO2,@INVOICE,@ITEMNAME,@WHNO,@ES)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PLATENO", PLATENO));
            command.Parameters.Add(new SqlParameter("@CARTONNO", CARTONNO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARTONQTY", CARTONQTY));
            command.Parameters.Add(new SqlParameter("@NW", NW));
            command.Parameters.Add(new SqlParameter("@GW", GW));
            command.Parameters.Add(new SqlParameter("@L", L));
            command.Parameters.Add(new SqlParameter("@W", W));
            command.Parameters.Add(new SqlParameter("@H", H));
            command.Parameters.Add(new SqlParameter("@LOACTION", LOACTION));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@CARTONNO2", CARTONNO2));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@ES", ES));
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

        public void UPPACKS(string SER)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE SHIPPING_PACK  SET SER=@SER  WHERE ID=(SELECT MAX(ID) FROM SHIPPING_PACK WHERE USERS=@USERS) ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public void DELMARK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete MARK where SHIPPINGCODE=@SHIPPINGCODE ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
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
        public System.Data.DataTable GetSHIPPACK6(string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  PAL_NW  FROM ACMESQL02.DBO.OITM  T1 ");
            sb.Append(" INNER JOIN CART T2 ON (T1.U_TMODEL=T2.MODEL_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS");
            sb.Append("  AND T1.U_VERSION =T2.MODEL_Ver COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE T1.ITEMCODE=@ITEMCODE  AND T2.PAL_QTY =@QTY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public System.Data.DataTable GetSHIPPACK7(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  PAL_NW,PAL_QTY  FROM ACMESQL02.DBO.OITM  T1  ");
            sb.Append(" INNER JOIN CART T2 ON (T1.U_TMODEL=T2.MODEL_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS ");
            sb.Append(" AND T1.U_VERSION =T2.MODEL_Ver COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" WHERE T1.ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetSHIPPACK9(string SHIPPINGCODE, string PLATENO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SUM(CAST(CARTONNO AS INT)) FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND PLATENO2=@PLATENO GROUP BY PLATENO2  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLATENO", PLATENO));
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
        public System.Data.DataTable GetWHPACKBLC(string SHIPPINGCODE, int SEQ)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT BLC  FROM WH_PACK4 WHERE SHIPPINGCODE =@SHIPPINGCODE  AND SEQ=@SEQ AND ISNULL(BLC,'') <> ''  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@SEQ", SEQ));
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
        private void InsertMARK(string ShippingCode, string seq, string Mark)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO MARK (ShippingCode,seq,Mark) VALUES(@ShippingCode,@seq,@Mark)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@Mark", Mark));



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
        public System.Data.DataTable GetWHPACKH(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT MARK  FROM WH_PACK3 WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        private void GETPACK(string SEQ, string CAR, string CHOSHIP)
        {
            Clear(sbS);
            string CHE = "";
            if (checkBox6.Checked)
            {
                CHE = "TRUE";
            }
            CON = 0;
            SBS();
            DELPACK4();
            WrPACK4();
            WriteExcelPACK(SEQ, CHE, CAR, CHOSHIP);

            System.Data.DataTable dt3 = GetSHIPPACK();

            System.Data.DataTable dt4 = ship.PackingListD;
            string DPLATENO = "";
            if (dt3.Rows.Count > 0 && packingListDDataGridView.Rows.Count < 2)
            {

                string DESED = "";
                int GV = 0;
                string SERS = "";
                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];
                    DataRow drw2 = dt4.NewRow();
                    string QQ = drw3["QTY"].ToString();
                    string SER = drw3["SER"].ToString();
                    string ES = drw3["ES"].ToString();
                    if (SERS != SER)
                    {
                        GV = 0;
                    }
                    SERS = drw3["SER"].ToString();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["plno"] = pLNoTextBox.Text;
                    drw2["seqno"] = j;
                    drw2["WHNO"] = drw3["WHNO"].ToString();
                    drw2["LOCATION"] = drw3["LOCATION"].ToString();
                    string PLATENO = drw3["PLATENO"].ToString().Trim();
                    //if (PLATENO == "12")
                    //{
                    //    MessageBox.Show("A");
                    //}
                    string ITEMCODE = drw3["ITEMCODE"].ToString().Trim();
                    string ITEMNAME = drw3["ITEMNAME"].ToString().Trim();
                    if (ITEMCODE == "")
                    {

                    }
                    if (ITEMCODE != "空箱")
                    {
                        drw2["PACKMARK"] = "True";
                    }
                    else
                    {
                        int CON1 = Convert.ToInt16(drw3["CARTONNO"]);
                        CON += CON1;
                        drw2["PACKMARK"] = "";
                    }

                    //if (ITEMCODE == "ACMERMA01.RMA01")
                    //{
                    //    MessageBox.Show("A");
                    //}

                    System.Data.DataTable H1S = GetSHIPPACK4S(shippingCodeTextBox.Text, ITEMCODE, ITEMNAME);

                    System.Data.DataTable H1S2 = GetSHIPPACK4S2(sbS.ToString(), ITEMNAME, ITEMCODE);
                    if (H1S.Rows.Count > 0)
                    {

                        string S1 = H1S.Rows[0][0].ToString().Trim();
                        drw2["DescGoods"] = S1;

                        if (H1S.Rows.Count > 1)
                        {

                            System.Data.DataTable H1SQ = GetSHIPPACK4SQTY(shippingCodeTextBox.Text, ITEMCODE, QQ);
                            if (H1SQ.Rows.Count > 0)
                            {
                                string S1Q = H1SQ.Rows[0][0].ToString().Trim();
                                drw2["DescGoods"] = S1Q;
                            }
                        }

                    }

                    if (SER.Trim() != "0")
                    {

                        GV++;
                        if (GV == 1)
                        {
                            System.Data.DataTable dt31 = GetSHIPPACK2(SER);
                            if (dt31.Rows.Count > 0)
                            {
                                string PACKAGE = dt31.Rows[0][0].ToString().Trim();
                                if (PACKAGE == "0-0")
                                {
                                    PACKAGE = "";
                                }
                                drw2["PackageNo"] = PACKAGE;
                                drw2["CNo"] = dt31.Rows[0][1].ToString().Trim();
                                drw2["Quantity"] = "'@" + drw3["CARTONQTY"];
                                drw2["Net"] = "'@" + Convert.ToDecimal(drw3["NW"]).ToString("0.000");
                                drw2["Gross"] = "'@" + drw3["GW"];
                                drw2["MeasurmentCM"] = "'@" + drw3["L"] + "x" + drw3["W"] + "x" + drw3["H"];
                            }
                        }
                        if (GV == 2)
                        {
                            System.Data.DataTable dt31 = GetSHIPPACK5(SER);
                            if (dt31.Rows.Count > 0)
                            {
                                drw2["Quantity"] = dt31.Rows[0][0].ToString().Trim();
                                drw2["Gross"] = dt31.Rows[0][1].ToString().Trim();
                                drw2["Net"] = dt31.Rows[0][2].ToString().Trim();
                            }
                            drw2["DescGoods"] = "";
                            drw2["PACKMARK"] = "";
                        }
                    }
                    else
                    {
                        GV = 0;

                        if (drw3["ITEMCODE"].ToString().Trim() == "空箱")
                        {

                            drw2["DescGoods"] = "(THIS PALLET INCLUDED " + drw3["CARTONNO"].ToString().Trim() + " EMPTY CARTONS.)";
                            drw2["PackageNo"] = "";
                            drw2["CNo"] = "";
                            drw2["Quantity"] = "";
                            drw2["Net"] = "";
                            drw2["Gross"] = "";
                            drw2["MeasurmentCM"] = "";
                        }
                        else
                        {
                            string PACK = drw3["PLATENO"].ToString().Trim();
                            string CNo = drw3["CARTONNO2"].ToString().Trim();
                            drw2["PackageNo"] = drw3["PLATENO"].ToString().Trim();
                            drw2["CNo"] = drw3["CARTONNO2"].ToString().Trim();
                            drw2["Quantity"] = drw3["CARTONQTY"].ToString().Trim();
                            drw2["Net"] = Convert.ToDecimal(drw3["NW"]).ToString("#,##0.000");
                            drw2["Gross"] = Convert.ToDecimal(drw3["GW"]).ToString("#,##0.00");
                            if (!String.IsNullOrEmpty(drw3["L"].ToString()))
                            {
                                drw2["MeasurmentCM"] = drw3["L"] + "x" + drw3["W"] + "x" + drw3["H"];
                            }
                        }
                    }
                    int ACME = ITEMCODE.IndexOf("ACME");
                    string DESE = drw2["DescGoods"].ToString();
                    int n;
                    if (int.TryParse(drw2["Quantity"].ToString(), out n) && int.TryParse(drw3["QTY"].ToString(), out n))
                    {
                        if (DESE != DESED && ACME == -1)
                        {
                            int QTY = Convert.ToInt16(drw2["Quantity"]);
                            int QTY2 = Convert.ToInt16(drw3["QTY"]);
                            if (QTY >= QTY2)
                            {
                                drw2["PALQTY"] = drw3["QTY"].ToString();
                            }

                            //20180604
                            System.Data.DataTable G11 = GetSHIPPACKQTY(ITEMCODE);
                            if (G11.Rows.Count > 0)
                            {
                                drw2["PALQTY"] = G11.Rows[0][0].ToString();
                            }
                        }
                    }
                    if (GV == 1)
                    {
                        if (DESE != DESED)
                        {
                            drw2["PALQTY"] = drw3["QTY"].ToString();
                        }
                    }
                    if (GV == 2)
                    {
                        drw2["PALQTY"] = "";
                    }
                    DESED = DESE;
                    drw2["SeqNo2"] = "";
                    drw2["TREETYPE"] = "";
                    drw2["VISORDER"] = 0;
                    drw2["SOID"] = "";
                    //if (!checkBox6.Checked)
                    //{
                    //    if (DPLATENO == PLATENO)
                    //    {
                    //        drw2["PackageNo"] = "";
                    //    }
                    //}

                    if (!String.IsNullOrEmpty(PLATENO))
                    {
                        DPLATENO = PLATENO;
                    }
                    if (GV <= 2)
                    {
                        dt4.Rows.Add(drw2);
                    }


                }

            }
            userNameTextBox.Text = CON.ToString();



        }

        public System.Data.DataTable GetSHIPPACK()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT *,CASE LOACTION WHEN N'中国' THEN 'CHINA' ELSE LOACTION END  LOCATION FROM SHIPPING_PACK where users=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetSHIPPACK4S2(string SHIPPINGCODE, string ITEMNAME, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT ITEMCODE  FROM WH_PACK2  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )  AND ITEMNAME=@ITEMNAME AND ITEMCODE <>@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetSHIPPACK4S(string ShippingCode, string ITEMCODE, string ITEMNAME)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND ITEMCODE=@ITEMCODE  ");
            if (ITEMCODE == "ACMERMA01.RMA01" && ITEMNAME.Length > 4)
            {
                string IM = ITEMNAME.Substring(3, 2);
                sb.Append(" AND INDescription  LIKE '%" + IM + "%' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetSHIPPACK4SQTY(string ShippingCode, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND ITEMCODE=@ITEMCODE AND INQTY=@QTY  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public System.Data.DataTable GetSHIPPACK5(string SER)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUM(CAST(CARTONQTY AS INT))　QTY, SUM(CAST(GW AS DECIMAL(10,2)))　GW, CAST(SUM(CAST(NW AS DECIMAL(10,4))) AS DECIMAL(10,3))　NW   FROM SHIPPING_PACK　WHERE SER=@SER AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetSHIPPACKQTY(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT QTY FROM SHIPPING_PACK where users=@USERS  AND ITEMCODE=@ITEMCODE ORDER BY CAST(QTY AS INT) DESC   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetWHPACK(string SHIPPINGCODE, string BLC, string CHE, string SB, string CAR)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (CHE == "TRUE")
            {
                sb.Append(" SELECT *,CAST(GW AS DECIMAL(10,2)) GW2 FROM WH_PACK2 WHERE SHIPPINGCODE IN (" + SB + "  ) ");
            }
            else
            {
                sb.Append(" SELECT *,CAST(GW AS DECIMAL(10,2)) GW2 FROM WH_PACK2 WHERE SHIPPINGCODE =@SHIPPINGCODE ");
            }
            if (!String.IsNullOrEmpty(BLC))
            {
                sb.Append(" AND BLC =@BLC ");
            }

            if (!String.IsNullOrEmpty(CAR))
            {
                sb.Append(" AND FLAG1 IN (" + CAR + "  ) ");
            }

            sb.Append("    ORDER BY SHIPPINGCODE,ID");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BLC", BLC));
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
        public System.Data.DataTable GetSHIPPACK2(string SER)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CAST(MIN(CAST(PLATENO AS INT)) AS VARCHAR) +'-'+CAST(MAX(CAST(PLATENO AS INT)) AS VARCHAR) PLATENO,CAST(MIN(CAST(SUBSTRING(CARTONNO2,1, (CASE CHARINDEX('~', CARTONNO2) WHEN 0 THEN 10 ELSE CHARINDEX('~', CARTONNO2)  END) -1) AS INT)) AS VARCHAR)+'~'+CAST(MAX(CAST(SUBSTRING(CARTONNO2, CHARINDEX('~', CARTONNO2)+1,5) AS INT)) AS VARCHAR) 　CARTONNO  FROM SHIPPING_PACK WHERE SER=@SER AND　USERS =@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

        private void bindingNavigatorAddNewItem5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetMenu.Getinvoicem(shippingCodeTextBox.Text);
           
            try
            {
                if (dt1.Rows.Count<1)
                {

                    MessageBox.Show("請輸入invoice單號");
                    return;
                }
               
            
                    DataRow drw = dt1.Rows[0];

                    string NumberName = "la" + DateTime.Now.ToString("yyyyMMdd");
                    string AutoNum = util.GetAutoNumber1(MyConnection, NumberName);

                    this.seqMNoTextBox.Text = AutoNum;
          
              

                    if (shipmentTextBox.Text != "")
                    {
                        loadingTextBox.Text = shipmentTextBox.Text;
                    }
                    if (receivePlaceTextBox.Text != "")
                    {
                        ladingTextBox.Text = receivePlaceTextBox.Text;
                    }
                    consigneeTextBox.Text = drw["shipTo"].ToString();
                    notifyPartTextBox.Text =  drw["billTo"].ToString();
                    shipperTextBox.Text = "ADVANCED DISPLAY LAB INC.";
                    if (boatNameTextBox.Text != "")
                    {
                        oceanVesselTextBox.Text = boatNameTextBox.Text;
                    }

                    if (unloadCargoTextBox.Text != "")
                    {
                        dischargeTextBox.Text = unloadCargoTextBox.Text;
                    }

                    lADINGMBindingSource.EndEdit();
                }
                
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }        
     
    

        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            try
            {

 
                System.Data.DataTable SE1 = GetINVSEQ(shippingCodeTextBox.Text);

                string SQE1 = SE1.Rows[0]["COUN"].ToString();
 
                    string NumberName = "I" + DateTime.Now.ToString("yyyyMMdd");
                    string AutoNum = "";
                    string SEQ = "";
                    if (SQE1 == "1")
                    {
                        AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                        SEQ = "01";
                        invoiceNoTextBox.Text = NumberName + AutoNum;
                    }
                    else
                    {
                        AutoNum = SE1.Rows[0]["INVOICENO"].ToString();

                        int f = Convert.ToInt16(SQE1);

                        if (f > 9)
                        {
                            SEQ = f.ToString();
                        }
                        else
                        {
                            SEQ = "0" + SQE1;
                        }
                        invoiceNoTextBox.Text = AutoNum;
                    }

                    invoiceNo_seqTextBox.Text = SEQ;




                    DataGridViewRow rowt;
                    rowt = shipping_ItemDataGridView.Rows[0];
                    string aas = rowt.Cells["ItemRemark"].Value.ToString();

                    string aa = rowt.Cells["Docentry"].Value.ToString();
                    if (shipmentTextBox.Text != "")
                    {
                        invoiceFromTextBox.Text = shipmentTextBox.Text;
                    }
                    if (unloadCargoTextBox.Text != "")
                    {
                        invoiceToTextBox.Text = unloadCargoTextBox.Text;
                    }


                    invoiceShipTextBox.Text = receiveDayTextBox.Text;
                    invoice_memoTextBox.Text = shippingCodeTextBox.Text;



                    System.Data.DataTable dtPI = GetMenu.GetPI(shippingCodeTextBox.Text);

                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i <= dtPI.Rows.Count - 1; i++)
                    {

                        DataRow dd = dtPI.Rows[i];


                        sb.Append(dd["docentry"].ToString() + ",");


                    }

                    sb.Remove(sb.Length - 1, 1);
                    pInoTextBox1.Text = sb.ToString();



                    string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

           
                        System.Data.DataTable dt1CHO = GetCHO3(shipping_OBUTextBox.Text, strCn);

                        System.Data.DataTable dt2CHO = GetCHO2(cardCodeTextBox.Text, strCn);
                        if (dt1CHO.Rows.Count > 0)
                        {

                            DataRow drw = dt1CHO.Rows[0];

                            oBUShipToTextBox1.Text = drw["shipbuilding"].ToString() +
                                    Environment.NewLine + drw["shipstreet"].ToString() +
                                    Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                                    Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                                    Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();
                        }

                        if (dt2CHO.Rows.Count > 0)
                        {

                            DataRow drw = dt2CHO.Rows[0];

                            oBUBillToTextBox1.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();
                        }
                
                            oBUBillToTextBox1.Text = "";
                            oBUShipToTextBox1.Text = "";
                            System.Data.DataTable dt1CHOAD = GetCHO3(pinoTextBox.Text, strCn);

                            System.Data.DataTable dt2CHOAD = GetCHO2(cardCodeTextBox.Text, strCn);
                            if (dt1CHOAD.Rows.Count > 0)
                            {

                                DataRow drw = dt1CHOAD.Rows[0];

                                shipToTextBox.Text = drw["shipbuilding"].ToString() +
                                        Environment.NewLine + drw["shipstreet"].ToString() +
                                        Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                                        Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                                        Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();
                            }

                            if (dt2CHOAD.Rows.Count > 0)
                            {

                                DataRow drw = dt2CHOAD.Rows[0];

                                billToTextBox.Text = drw["billbuilding"].ToString() +
                                Environment.NewLine + drw["billstreet"].ToString() +
                                Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                                Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                                Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();
                            }
                        
                    

       

                    if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1349-00")
                    {
                        if (add2TextBox.Text != "")
                        {
                            oBUShipToTextBox1.Text = shipToTextBox.Text;
                        }
                    }
                //    invoiceDDataGridView.Enabled = false;
                
            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Invoice新增", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                MessageBox.Show(ex.Message);
            }


            try
            {



                System.Data.DataTable dt3 = Getshipitem(shippingCodeTextBox.Text, 1, "");

                System.Data.DataTable dt4 = ship.InvoiceD;

                if (shipping_ItemDataGridView.Rows.Count > 1 && invoiceDDataGridView.Rows.Count < 2)
                {

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt3.Rows[i];
                        DataRow drw2 = dt4.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["InvoiceNo"] = invoiceNoTextBox.Text;
                        drw2["InvoiceNo_seq"] = invoiceNo_seqTextBox.Text;
                        drw2["SeqNo"] = i.ToString();
                        string ITEMCODE = drw["itemcode"].ToString();
                        System.Data.DataTable TT1 = GetADINVOPROD(ITEMCODE);
                        if (TT1.Rows.Count > 0)
                        {
                            drw2["INDescription"] = TT1.Rows[0][0].ToString();
                        }
                        else
                        {
                            drw2["INDescription"] = drw["bb"];
                        }
                        drw2["InQty"] = drw["Quantity"];
                        drw2["UnitPrice"] = drw["ItemPrice"];
                        drw2["CURRENCY"] = drw["CURRENCY"];
                        drw2["RATE"] = drw["RATE"];
                        drw2["RATEUSD"] = drw["RATEUSD"];
                        string TYPE = drw["OLDORDER"].ToString();

                        drw2["amount"] = drw["ItemAmount"];
                  
                        
                
                        drw2["LINENUM"] = drw["linenum"];
                    

                        drw2["CHOPrice"] = drw["CHOPrice"];
                        drw2["CHOAmount"] = drw["CHOAmount"];
                        drw2["TREETYPE"] = TYPE;
                        drw2["VISORDER"] = drw["VISORDER"];
                        dt4.Rows.Add(drw2);

                    }

                }

                try
                {

                    this.invoiceMBindingSource.EndEdit();
                    this.invoiceMTableAdapter.Update(ship.InvoiceM);
                    ship.InvoiceM.AcceptChanges();

                    this.invoiceDBindingSource.EndEdit();
                    this.invoiceDTableAdapter.Update(ship.InvoiceD);
                    ship.InvoiceD.AcceptChanges();

                }
                catch (Exception ex)
                {

                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "InvoiceTran1", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void lADINGDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                int iRecs;

                iRecs = lADINGDDataGridView.Rows.Count - 1;
                e.Row.Cells["dataGridViewTextBoxColumn54"].Value = iRecs.ToString();

                System.Data.DataTable dt1 = GetMenu.Getgross(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = GetOrderData4();
                
   
                if (dt1.Rows.Count > 0)
                {
         
                    DataRow drw = dt1.Rows[0];
                    e.Row.Cells["dataGridViewTextBoxColumn58"].Value = drw["gross"].ToString();
                }

                if (dt2.Rows.Count > 0)
                {
                    DataRow drw2 = dt2.Rows[0];
                    e.Row.Cells["dataGridViewTextBoxColumn56"].Value = drw2["packageno"].ToString() + "PLTS";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void invoiceDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = invoiceDDataGridView.Rows.Count - 1;
            e.Row.Cells["dataGridViewTextBoxColumn30"].Value = iRecs.ToString();
            e.Row.Cells["dataGridViewTextBoxColumn33"].Value = "0";
            e.Row.Cells["dataGridViewTextBoxColumn34"].Value = "0";
            e.Row.Cells["dataGridViewTextBoxColumn35"].Value = "0";
        }

        private void invoiceDDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (invoiceDDataGridView.Columns[e.ColumnIndex].Name == "dataGridViewTextBoxColumn33")
                {
                    int iQuantity = 0;
                    decimal iUnitPrice = 0;
                    if (!String.IsNullOrEmpty(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn34"].Value.ToString()))
                    {
                        iUnitPrice = Convert.ToDecimal(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn34"].Value);
                    }
              
                    iQuantity = Convert.ToInt32(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn33"].Value);
                    this.invoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn35"].Value = (iQuantity * iUnitPrice).ToString();



                    int iQuantity1 = 0;
                    decimal iUnitPrice1 = 0;
                    if (!String.IsNullOrEmpty(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["CHOPrice"].Value.ToString()))
                    {
                        iUnitPrice1 = Convert.ToDecimal(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["CHOPrice"].Value);
                    }
                    iQuantity1 = Convert.ToInt32(this.invoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn33"].Value);

                    this.invoiceDDataGridView.Rows[e.RowIndex].Cells["CHOAmount"].Value = (iQuantity1 * iUnitPrice1).ToString();
                }
            }
            catch { }

        }

        private void packingListDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
          

            int iRecs;

            iRecs = packingListDDataGridView.Rows.Count-1;
            e.Row.Cells["dataGridViewTextBoxColumn44"].Value = iRecs.ToString();
           
        }

 

        private void 儲存SToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {

                this.Validate();

                if (invoiceNo_seqTextBox.Text == "")
                {

                    MessageBox.Show("請輸入invoice代號");
                    return;
                }
                if (invoiceDDataGridView.Rows.Count > 1)
                {


                    UPINVOICE();
                }


                    
                    try
                    {

                        this.invoiceMBindingSource.EndEdit();
                        this.invoiceMTableAdapter.Update(ship.InvoiceM);
                        ship.InvoiceM.AcceptChanges();
                        this.invoiceDBindingSource.EndEdit();
                        this.invoiceDTableAdapter.Update(ship.InvoiceD);
                        ship.InvoiceD.AcceptChanges();

                        MessageBox.Show("儲存成功");
                    }
                    catch (Exception ex)
                    {

                        GetMenu.InsertLog(fmLogin.LoginID.ToString(), "InvoiceTran2", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                        MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                    }
             


                    }
               
                  
                
      
            
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "invoice儲存異常", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));
                MessageBox.Show(ex.Message);
            }

        }

        private void CalcTotals1()
        {
            try
            {

                decimal AMT = 0;
                string CURRENCY2 = "";

       
                int i = this.invoiceDDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    if (!String.IsNullOrEmpty(invoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn35"].Value.ToString().Trim()))
                    {
                        AMT += Convert.ToDecimal(invoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn35"].Value.ToString().Trim());
                    }

                    CURRENCY2 = invoiceDDataGridView.Rows[iRecs].Cells["CURRENCY2"].Value.ToString().Trim();
                }

                System.Data.DataTable t1 = HF1(AMT.ToString());
                int G=Convert.ToInt32(AMT.ToString("###0"));
                double AMT1 = Convert.ToDouble(AMT);
            
                if (G != 0)
                {
                    if (CURRENCY2 == "RMB")
                    {
                        
                        amountTotalEngTextBox.Text = "SAY TOTAL : RMB DOLLARS " + new Class1().NumberToString(AMT1);
                        amountTotalTextBox.Text = "RMB"+t1.Rows[0][0].ToString();
                    }
                    else
                    {
                        amountTotalEngTextBox.Text = "SAY TOTAL : US DOLLARS " + new Class1().NumberToString(AMT1);
                        amountTotalTextBox.Text =  t1.Rows[0][0].ToString();
                    }
                }
                else
                {
                    amountTotalEngTextBox.Text = "";
                }
                this.invoiceMBindingSource.EndEdit();
                this.invoiceMTableAdapter.Update(ship.InvoiceM);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public System.Data.DataTable HF1(string SQL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select CONVERT(NVARCHAR(20),CAST(@SQL AS Money),1) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SQL", SQL));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void CalcTotals1C()
        {
            try
            {

                double AMT = 0;



                int i = this.invoiceDDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    if (!String.IsNullOrEmpty(invoiceDDataGridView.Rows[iRecs].Cells["CHOAmount"].Value.ToString().Trim()))
                    {
                        AMT += Convert.ToDouble(invoiceDDataGridView.Rows[iRecs].Cells["CHOAmount"].Value.ToString().Trim());
                    }



                }
                int G = Convert.ToInt32(AMT.ToString("###0"));
                double AMT1 = Convert.ToDouble(AMT);
                if (G != 0)
                {
                    amountTotalEngTextBox.Text = "SAY TOTAL : US DOLLARS " + new Class1().NumberToString(AMT1);
                }
                else
                {
                    amountTotalEngTextBox.Text = "";
                }
                this.invoiceMBindingSource.EndEdit();
                this.invoiceMTableAdapter.Update(ship.InvoiceM);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void CalAMTINVOICE(string TYPE)
        {
            try
            {

                //int AMT = 0;
                //int s = this.shipping_ItemDataGridView.Rows.Count - 2;
                //for (int iRecs = 0; iRecs <= s; iRecs++)
                //{
                //    string F1 = shipping_ItemDataGridView.Rows[iRecs].Cells["ItemAmount"].Value.ToString().Trim();
                //    if (!String.IsNullOrEmpty(shipping_ItemDataGridView.Rows[iRecs].Cells["ItemAmount"].Value.ToString().Trim()))
                //    {
                //        AMT += Convert.ToInt32(shipping_ItemDataGridView.Rows[iRecs].Cells["ItemAmount"].Value.ToString().Trim());
                //    }
                //}

                //int AMTI = 0;
                //int i = this.invoiceDDataGridView.Rows.Count - 2;
                //for (int iRecs = 0; iRecs <= i; iRecs++)
                //{
                //    if (!String.IsNullOrEmpty(invoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn35"].Value.ToString().Trim()))
                //    {
                //        AMTI += Convert.ToInt32(invoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn35"].Value.ToString().Trim());
                //    }
                //}

                //if (TYPE == "A")
                //{

                //    if (AMT != AMTI)
                //    {
                //        MessageBox.Show("項目金額跟Invoice金額不同");
                //    }
                //}

                //if (TYPE == "B")
                //{
                //    if (AMT != AMTI)
                //    {
                //        SHIPAMT frm1 = new SHIPAMT();
                //        frm1.JOBNO = shippingCodeTextBox.Text;
                //        if (frm1.ShowDialog() == DialogResult.OK)
                //        {
                //            int  ss = Convert.ToInt32(frm1.b.ToString());

                //            if (AMTI != ss)
                //            {
                //                MessageBox.Show("輸入金額跟Invoice金額不同");
                //                return;
                               
                //            }


                //        }
                //    }
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void CalcTotals2()
        {
            string QTY = "";
            string NET2 = "";
            string GROSS2 = "";
            try
            {

                Int32 Quantity = 0;
                decimal NET = 0;
                decimal GROSS = 0;

                int i = this.packingListDDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    QTY= packingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn48"].Value.ToString();
                    NET2 = packingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn49"].Value.ToString();
                    GROSS2 = packingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn50"].Value.ToString();
                  
                   if (!String.IsNullOrEmpty(QTY))
                    {

                        int g = QTY.LastIndexOf("@");
                        if (g == -1)
                        {
                            Quantity += Convert.ToInt32(QTY.Trim());

                        }
                    }
                   if (!String.IsNullOrEmpty(NET2))
                    {
                        int U = NET2.LastIndexOf("@");
                        if (U == -1)
                        {

                            NET += Convert.ToDecimal(NET2.Trim());
                        }
                    }

                   if (!String.IsNullOrEmpty(GROSS2))
                    {

                        int V = GROSS2.LastIndexOf("@");
                        if (V == -1)
                        {
                            GROSS += Convert.ToDecimal(GROSS2.Trim());
                        }
                    }


                }

                quantityTextBox1.Text = Quantity.ToString();
                netTextBox.Text = NET.ToString();
                grossTextBox.Text = GROSS.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
      
        private void 儲存SToolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                this.Validate();
        
                if (packingListDDataGridView.Rows.Count > 1)
                {

                    string d;
                    string f;
                    CalcTotals2();


                    int k = packingListDDataGridView.Rows.Count - 2;

                    DataGridViewRow rowP;

                    rowP = packingListDDataGridView.Rows[k];
                    string a0 = rowP.Cells["PackageNo"].Value.ToString();
                    string a1 = rowP.Cells["dataGridViewTextBoxColumn46"].Value.ToString();
                    
                    int g = a0.LastIndexOf("-");
                    if (g == 0)
                    {
                        f = a0;
                    }
                    else
                    {
                        f = a0.Substring(g + 1);
                    }
                    sayTotalTextBox.Text = f;

                    int ss = a1.LastIndexOf("~");
                    if (ss == 0)
                    {
                        d = a1;
                    }
                    else
                    {
                        d = a1.Substring(ss + 1);
                    }


                    if (f != "")
                    {
                        int amountText = Convert.ToInt32(f);
                        string s = f;
                        columnTotalTextBox.Text = new Class1().NumberToString2(amountText, s, d);
                    }



                    UPPACK();


                }


                try
                {

                    this.packingListMBindingSource.EndEdit();
                    this.packingListMTableAdapter.Update(ship.PackingListM);
                    ship.PackingListM.AcceptChanges();

                    this.packingListDBindingSource.EndEdit();
                    this.packingListDTableAdapter.Update(ship.PackingListD);
                    ship.PackingListD.AcceptChanges();
      

                    MessageBox.Show("儲存成功");

                }
                catch (Exception ex)
                {

                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "PackingTran2", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                }
       

           

            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Packing儲存", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));
                MessageBox.Show(ex.Message);
            }
          
        }
 
        private void 儲存SToolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                 this.Validate();
               
                try
                {
                    lADINGDBindingSource.MoveFirst();

                    for (int i = 0; i <= lADINGDBindingSource.Count - 1; i++)
                    {
                        DataRowView row = (DataRowView)lADINGDBindingSource.Current;

                        row["seqno"] = i;



                        lADINGDBindingSource.EndEdit();

                        lADINGDBindingSource.MoveNext();
                    }


                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);

                }

                SqlTransaction tx = null;
                try
                {

         
                    this.lADINGMBindingSource.EndEdit();
                    this.lADINGDBindingSource.EndEdit();


                    this.lADINGMTableAdapter.Update(ship.LADINGM);
                    this.lADINGDTableAdapter.Update(ship.LADINGD);

                    ship.PackingListM.AcceptChanges();
                    ship.PackingListD.AcceptChanges();

                    MessageBox.Show("儲存成功");

                }
                catch (Exception ex)
                {
                  

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                }
         

    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        private void cFSDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["dataGridViewTextBoxColumn13"].Value = util.GetSeqNo(2, cFSDataGridView);
        }

        private void button3_Click(object sender, EventArgs e)
        {
        
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            FileName = lsAppDir + "\\Excel\\PACKAD.xls";
            GetExcelProduct(FileName, GetOrderData3(), "Y");


         
        }



   
        private void downloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK2")
                {
                 
                        System.Data.DataTable dt1 = ship.Download;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string aa = drw["path"].ToString();
                        string filename = drw["filename"].ToString();
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
                        string SA = sATextBox.Text.ToUpper();
                        string ID = drw["id"].ToString();

                        System.IO.File.Copy(aa, NewFileName, true);
                        System.Diagnostics.Process.Start(NewFileName);

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

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
            string f = "c";
            string[] filebType = Directory.GetDirectories(DIR);
            string dd = DateTime.Now.ToString("yyyyMM");
            string tt = DIR + dd;
            foreach (string fileaSize in filebType)
            {

                if (fileaSize == tt)
                {
                    f = "d";

                }

            }
            if (f == "c")
            {
                Directory.CreateDirectory(tt);
            }
         
                string server = DIR + dd + "//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download(filename);

           
                    if (result == DialogResult.OK)
                    {
               
                        string file = opdf.FileName;
                        bool F1 = getrma.UploadFile(file, server, false);
                        if (F1 == false)
                        {
                            return;
                        }
                        System.Data.DataTable dt1 = ship.Download;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (downloadDataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = PATH+ de + filename;
             
     
                        dt1.Rows.Add(drw);

                        downloadBindingSource.MoveFirst();

                        for (int i = 0; i <= downloadBindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)downloadBindingSource.Current;

                            rowd["seq"] = i;

                            downloadBindingSource.EndEdit();
                            downloadBindingSource.MoveNext();
                        }

                        this.downloadBindingSource.EndEdit();
                        this.downloadTableAdapter.Update(ship.Download);
                        this.ship.Download.AcceptChanges();
                  

                        MessageBox.Show("上傳成功");
                    }


                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button12_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            

            FileName = lsAppDir + "\\Excel\\ADBook3.xls";


            GetExcelProduct3(FileName,GetOrderData());
            dollarsKindTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        private System.Data.DataTable GetOrderData()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  select isnull(d.SoNo,'') SoNo,isnull(a.ShippingCode,'') ShippingCode,isnull(a.shipper,'') shipper,isnull(a.Consignee,'') Consignee,isnull(b.Cargo,'') TARE,isnull(a.NotifyPart,'') NotifyPart,");
            sb.Append(" isnull(d.receivePlace,'') receivePlace,isnull(a.OceanVessel,'') OceanVessel,isnull(a.Discharge,'') Discharge,isnull(a.Delivery,'') Delivery, ");
            sb.Append(" isnull(b.ContainerSeals,'') ContainerSeals,isnull(b.Packages,'') Packages,isnull(b.Description,'') Description,isnull(b.Measurement,'') Measurement,isnull(a.freightPaid,'') freightPaid,a.loading shipment  from ladingm a left join ladingd b on(a.shippingcode=b.shippingcode and a.seqmno=b.seqmno)");
            sb.Append("  left join shipping_main  d on(a.ShippingCode=d.ShippingCode) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.SeqMNo=@SeqMNo ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@SeqMNo", seqMNoTextBox.Text));
           
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
        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
           
                sb.Append(" SELECT a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期");
                sb.Append(" ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],b.[MarkNos],");
                if (GetINVMARK().Rows.Count == 0)
                {
                    sb.Append(" cast(seqno2+1 as varchar)+')'+b.[INDescription] as INDescription");
                }
                else
                {
                    sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(seqno2+1 as varchar)+')'+b.[INDescription] END INDescription ");
                }
                sb.Append(" ,b.[InQty]  ,c.brand +' BRAND' as BRAND,c.TradeCondition as Trade,");
                sb.Append(" CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[UnitPrice] AS Money),1) UnitPrice");
                sb.Append(",CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[Amount] AS Money),1) Amount,kPIYESNO PRJ");
                sb.Append(" FROM [InvoiceM] as a ");
                sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)");
                sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
                sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");
            
         


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo", invoiceNoTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo_seq ", invoiceNo_seqTextBox.Text));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData2DRS(string CARD, string ADD, string TEL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期");
            sb.Append(" ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],b.[MarkNos],");
            if (GetINVMARK().Rows.Count == 0)
            {
                sb.Append(" cast(seqno2+1 as varchar)+')'+b.[INDescription] as INDescription");
            }
            else
            {
                sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(seqno2+1 as varchar)+')'+b.[INDescription] END INDescription ");
            }
            sb.Append(" ,b.[InQty] ,b.[UnitPrice]  ,b.[Amount],c.brand +' BRAND' as BRAND,c.TradeCondition as Trade,CARD=@CARD,[ADD]=@ADD,TEL=@TEL FROM [InvoiceM] as a");
            sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo", invoiceNoTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo_seq ", invoiceNo_seqTextBox.Text));
            command.Parameters.Add(new SqlParameter("@CARD ", CARD));
            command.Parameters.Add(new SqlParameter("@ADD ", ADD));
            command.Parameters.Add(new SqlParameter("@TEL ", TEL));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData2DRS2()
        {

            SqlConnection connection = globals.shipConnection ;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T2.building CARD,T2.street+ISNULL(T2.COUNTY,'') [ADD],");
            sb.Append(" CASE WHEN ISNULL(T2.block,'') <> '' THEN  'TEL: '+block ELSE '' END +CASE WHEN ISNULL(T2.city,'') <> '' THEN  'FAX: '+city ELSE '' END TEL FROM OPOR T0 ");
            sb.Append("   LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docnum");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docnum", pinoTextBox.Text));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GetINVMARK()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT shippingcode FROM INVOICED a WHERE  a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq and marknos='true'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo", invoiceNoTextBox.Text));
            command.Parameters.Add(new SqlParameter("@InvoiceNo_seq ", invoiceNo_seqTextBox.Text));

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
        public System.Data.DataTable GetSA(string PINO)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管");
            sb.Append(" FROM ORDR T0 ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode");
            sb.Append(" INNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" WHERE    T0.DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", PINO));

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

        public System.Data.DataTable GetINVPACK()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT shippingcode FROM PackingListD a WHERE  a.shippingcode=@shippingcode and a.PLNo=@PLNo  and PACKMARK='true'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));


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

        public System.Data.DataTable GetDOWNLOADSA()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [PATH]   FROM DOWNLOAD WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(DLCHECK,'')='True'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

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
        public System.Data.DataTable GetDOWNLOADSAT1()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.SHIPPINGCODE,DOCENTRY  FROM SHIPPING_ITEM T0 ");
            sb.Append(" WHERE ITEMREMARK='銷售訂單'  AND T0.SHIPPINGCODE IN (");
            sb.Append(" SELECT DISTINCT SHIPPINGCODE FROM DOWNLOAD WHERE ISNULL(SA,'') = '')");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

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
        public System.Data.DataTable GetDOWNLOADSA2()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SA,SALES   FROM DOWNLOAD WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(DLCHECK,'')='True' AND ISNULL(SA,'') <> ''  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

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
        public System.Data.DataTable GetDOWNLOADSA3(string SA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT EMAIL FROM OHEM WHERE ([lastName]+[firstName])=@SA  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SA", SA));

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

        public System.Data.DataTable GetDOWNLOADSA4(string SLPNAME)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_EMAIL FROM OSLP WHERE SLPNAME=@SLPNAME  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SLPNAME", SLPNAME));

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


        private System.Data.DataTable GetDOWNLOAD2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                      SELECT * FROM DOWNLOAD2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND MARK='1' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetDOWNLOAD22()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM DOWNLOAD2 WHERE MARK='1' AND REPLACE([FILENAME],' ','') LIKE '%" + add9TextBox.Text.ToString().Replace(" ","") + "%'  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    

        private System.Data.DataTable GetSALES(string DOCENTRY)
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" select CASE WHEN SUBSTRING(T0.CARDCODE,1,1)='S' THEN T3.[lastName]+T3.[firstName] ELSE T2.SLPNAME END 業務 from opor T0  ");
            sb.Append(" LEFT  JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode");
            sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  WHERE T0.DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSame()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT a.SayTotal+' PLTS' as cc,a.[PLNo]  invoiceNo,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[OBUShipTo] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計'");
            sb.Append(" ,a.[Net] as '耐特',a.[Gross] as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[OBUBillTo] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]");
            sb.Append(" ,a.[Quantity] as '總數',a.[Net],a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo],substring(seqno,2,1)+')'+b.[DescGoods] as DescGoods");
            sb.Append(" ,b.[Quantity] as Quantity ,b.[Net] as Ne ,cast(b.[Gross] as varchar) as Go ,b.[MeasurmentCM] FROM [PackingListM] as a");
            sb.Append(" left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.PLNo=@PLNo order by cast(seqno as int)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void SBS()
        {
            string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS.Append("'" + i + "',");
            }
            sbS.Remove(sbS.Length - 1, 1);
        }
        private System.Data.DataTable GetOrderData3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT a.SayTotal+' PLTS' as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計' ");
            sb.Append("               ,a.[Net] as '耐特',a.[Gross] as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo] ");
            sb.Append("               ,a.[Quantity] as '總數',a.[Net],a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo], ");
            if (GetINVPACK().Rows.Count == 0)
            {
                sb.Append("               CAST(seqno2+1 AS VARCHAR)+')'+CASE ISNULL(TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods] END as DescGoods ");
            }
            else
            {

                sb.Append(" CASE WHEN ISNULL(PACKMARK,'') <> 'True' THEN '' ELSE cast(seqno2+1 as varchar)+')' END+CASE ISNULL(TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods]  END DescGoods ");
            }

          
            sb.Append("               ,b.[Quantity] as Quantity ,b.[Net] as Ne ,cast(b.[Gross] as varchar) as Go ,b.[MeasurmentCM] FROM [PackingListM] as a ");
            sb.Append("               left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo) ");
            sb.Append("               left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.PLNo=@PLNo order by cast(seqno as int) ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3DRS(string CARD,string ADD,string TEL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT a.SayTotal+' PLTS' as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計' ");
            sb.Append("               ,a.[Net] as '耐特',a.[Gross] as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo] ");
            sb.Append("               ,a.[Quantity] as '總數',a.[Net],a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo], ");
            if (GetINVPACK().Rows.Count == 0)
            {
                sb.Append("               CAST(seqno2+1 AS VARCHAR)+')'+CASE ISNULL(TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods] END as DescGoods ");
            }
            else
            {

                sb.Append(" CASE WHEN ISNULL(PACKMARK,'') <> 'True' THEN '' ELSE cast(seqno2+1 as varchar)+')' END+CASE ISNULL(TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods]  END DescGoods ");
            }


            sb.Append("               ,b.[Quantity] as Quantity ,b.[Net] as Ne ,cast(b.[Gross] as varchar) as Go ,b.[MeasurmentCM],CARD=@CARD,[ADD]=@ADD,TEL=@TEL FROM [PackingListM] as a ");
            sb.Append("               left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo) ");
            sb.Append("               left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.PLNo=@PLNo order by cast(seqno as int) ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));
            command.Parameters.Add(new SqlParameter("@ADD", ADD));
            command.Parameters.Add(new SqlParameter("@TEL", TEL));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSHIPMARK()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT MARK FROM mark WHERE SHIPPINGCODE=@SHIPPINGCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSHIPOHEM(string USER)
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT HOMETEL FROM OHEM WHERE WORKCOUNTR='CN' AND HOMETEL=@HOMETEL AND ISNULL(TERMDATE,'') =''   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", USER));
        

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSHIPEXSIT()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SHIPPINGCODE  FROM  SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
       

      
        private System.Data.DataTable GetOrderData4()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select top 1 packageno,seqno,cno from PackingListd");
            sb.Append(" where shippingcode=@shippingcode order by cast(seqno as int) desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListd");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void GetExcelProduct(string ExcelFile,System.Data.DataTable dt,string FLAG)
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
            excelSheet.Name = shippingCodeTextBox.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string B2 = "//acmew08r2ap//table//SIGN//USER//";
                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                if (FLAG == "Y")
                {

                    excelSheet.Shapes.AddPicture(B2 + createNameTextBox.Text.Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoTrue, 350, 650, 200, 80);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue,dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 10;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue,dt);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }

                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.Getmark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {

                        //最後一筆不作




                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 6]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }

                }
                       
            }
            finally
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
                             shippingCodeTextBox.Text + "-PACKING LIST.xls";

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
        private void GetExcelProductBOM(string ExcelFile, System.Data.DataTable dt)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false ;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Name = shippingCodeTextBox.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue, dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 10;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue, dt);

                            range.Value2 = FieldValue;

                            if (iField == 8)
                            {
                                if (FieldValue == "1")
                                {
                                    for (int L = 1; L <= 6; L++)
                                    {
                                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, L]);
                                        range.Select();
                                        range.Font.Bold = true;

                                    }
                                }
                            }
                        }

                        DetailRow++;
                    }

                }

              
                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.Getmark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {

                        //最後一筆不作




                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 6]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 8]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);
                }


       
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 7]);
                    range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();

  


                   
                     
                  
                    


                

            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);

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

        private void GetExcelProduct2(string ExcelFile,System.Data.DataTable dt,string FLAG)
        {
       
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
       
            excelApp.Visible = false;
  
            //Interop params
            object oMissing = System.Reflection.Missing.Value;
     
            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
           // excelApp.ActiveWindow.Zoom = 95;
            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Name = shippingCodeTextBox.Text;
           
           

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                string B2 = "//acmew08r2ap//table//SIGN//USER//";
                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                if (FLAG == "Y")
                {
    //                excelSheet.Shapes.AddPicture(B2 + fmLogin.LoginID.ToString().Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
    //Microsoft.Office.Core.MsoTriState.msoTrue, Convert.ToInt16(textBox9.Text), Convert.ToInt16(textBox10.Text), Convert.ToInt16(textBox11.Text), Convert.ToInt16(textBox12.Text));

                    excelSheet.Shapes.AddPicture(B2 + createNameTextBox.Text.Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoTrue, 410, 682, 200, 80);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue,dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 23;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue,dt);

                       
                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }



                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.Getmark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count != 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {

                        //最後一筆不作




                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }

                }
            }
            finally
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
            shippingCodeTextBox.Text + "-INVOICE.xls";

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

                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);

            }
        }
        private void GetObuInvoExcel(string ExcelFile,System.Data.DataTable dt)
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
            excelSheet.Name = shippingCodeTextBox.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


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
 
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
         
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

                        if (CheckSerial(sTemp, ref FieldValue,dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 22;
                    
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue,dt);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }

        

                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.Getmark(shippingCodeTextBox.Text);
                //增加另一talbe處理

                for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                {

                    //最後一筆不作

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                    // range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    string FieldName = "mark";

                    FieldValue1 = "";
                    FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                    range.Value2 = FieldValue1;

                    DetailRow1++;
                }
          

            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);

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

                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);

            }
        }

        private void GetExcelProduct3(string ExcelFile,System.Data.DataTable dt)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false ;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Name = shippingCodeTextBox.Text;
            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                string FieldValue2 = string.Empty;
                string FieldValue3 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
       
                int DetailRow3 = 0;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue,dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                          
                            DetailRow3 = 25;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue, dt);

                            int H1 = FieldValue.IndexOf("=");
                            FieldValue = FieldValue.Replace("=", "");
                            if (H1 != -1)
                            {
                                range.Value2 = " =" + FieldValue.ToString();
                            }
                            else
                            {
                                range.Value2 = FieldValue.ToString();
                            }


                        }

                        DetailRow++;
                    }

                }

              

                //增加另一talbe處理

                
                System.Data.DataTable mark = GetMenu.Getmark2(shippingCodeTextBox.Text);
                if (mark.Rows.Count > 0)
                {
                    for (int a3Row = 0; a3Row <= mark.Rows.Count - 1; a3Row++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3, 1]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName2 = "ContainerSeals";

                        FieldValue3 = "";
                        FieldValue3 = Convert.ToString(mark.Rows[a3Row][FieldName2]);

                        range.Value2 = FieldValue3;
                        DetailRow3++;
                    }
                }

            
            }

            finally
            {

                string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);

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
        private void GetExcelinsu(string ExcelFile,System.Data.DataTable dt)
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

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


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
                string FieldValue3 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                int DetailRow2 = 0;
                int DetailRow3 = 0;
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

                        if (CheckSerial(sTemp, ref FieldValue,dt))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 24;
                            DetailRow2 = 24;
                            DetailRow3 = 24;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue,dt);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }

                for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != dt.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        FieldValue = "";
                        SetRow(aRow, sTemp, ref FieldValue,dt);

                        range.Value2 = FieldValue;


                    }

                    DetailRow++;
                }

                
            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);

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

                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);

            }
        }
        private void SetRow(int iRow, string sData, ref string FieldValue,System.Data.DataTable dt)
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
                FieldValue = Convert.ToString(dt.Rows[iRow][FieldName]);
            }

        }

        private void SetRow1(int iRow, string sData, ref string FieldValue, System.Data.DataTable dt)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(dt.Rows[iRow][FieldName]);
            }

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


        private bool CheckSerial(string sData, ref string FieldValue, System.Data.DataTable dt)
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
                FieldValue = Convert.ToString(dt.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }

        private void button10_Click(object sender, EventArgs e)
        {

            CalAMTINVOICE("B");
            CalcTotals1();
           
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            FileName = lsAppDir + "\\Excel\\INVO2AD.xls";
            GetExcelProduct2(FileName, GetOrderData2(), "Y");

       
        }

        private void markDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {

            e.Row.Cells["dataGridViewTextBoxColumn10"].Value = util.GetSeqNo(2, markDataGridView);
        }

        private void tabPage9_Enter(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetMenu.Getinvoicem(shippingCodeTextBox.Text);


            if (dt1.Rows.Count <= 0)
            {

                bindingNavigator6.Enabled = false;
                MessageBox.Show("請輸入invoice單號");

                tabControl1.SelectedIndex = 1;
            }    
        }

        private void tabPage8_Enter(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetMenu.Getinvoicem(shippingCodeTextBox.Text);


            if (dt1.Rows.Count <= 0 )
            {

                MessageBox.Show("請新增invoice單號");

    
                tabControl1.SelectedIndex = 1;
            }
            else if (receiveDayTextBox.Text == "")
            {
                MessageBox.Show("請輸入運送方式");

                tabControl1.SelectedIndex = 1;
            }
     }

    
        private void 明細插入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.InvoiceD;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["InvoiceNo"] = invoiceNoTextBox.Text;
         
            newCustomersRow["InvoiceNo_seq"] = invoiceNo_seqTextBox.Text;

            newCustomersRow["amount"] = 0;

            try
            {

                dt2.Rows.InsertAt(newCustomersRow, invoiceDDataGridView.CurrentRow.Index);

  
                for (int j = 0; j <= invoiceDDataGridView.Rows.Count - 2; j++)
                {
                    invoiceDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }


                this.invoiceDBindingSource.EndEdit();
                this.invoiceDTableAdapter.Update(ship.InvoiceD);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("收件人地址為" + textBox2.Text + "是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
              

                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\Report.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                string Html = GetTODO_USERDataSource2();

                template = template.Replace("##shippingCode##", "JOB NO: " + shippingCodeTextBox.Text);
                template = template.Replace("##soNo##", "Shipping Order No : " + soNoTextBox.Text);
                template = template.Replace("##tradeCondition##", "貿易條件 : " + tradeConditionTextBox.Text);
                template = template.Replace("##closeDay##", "結關日 : " + closeDayTextBox.Text);
                template = template.Replace("##forecastDay##", "預計開航日 : " + forecastDayTextBox.Text);
                template = template.Replace("##arriveDay##", "預計抵達日 : " + arriveDayTextBox.Text);

                template = template.Replace("##receivePlace##", "收貨地 : " + receivePlaceTextBox.Text);
                template = template.Replace("##goalPlace##", "目的地 : " + goalPlaceTextBox.Text);
                template = template.Replace("##shipment##", "裝船港 : " + shipmentTextBox.Text);
                template = template.Replace("##boatName##", "港名/航次 : " + boatNameTextBox.Text);
                template = template.Replace("##boatCompany##", "船公司 : " + boatCompanyTextBox.Text);
                template = template.Replace("##unloadCargo##", "卸貨港 : " + unloadCargoTextBox.Text);
                template = template.Replace("##boardCount##", "20呎 : " + boardCountTextBox.Text);
                template = template.Replace("##boardDeliver##", "40呎 : " + boardDeliverTextBox.Text);
                template = template.Replace("##sendGoods##", "併櫃/CBM : " + sendGoodsTextBox.Text);
                template = template.Replace("##quantity##", "報單號碼 : " + add9TextBox.Text);
                template = template.Replace("##receiveDay##", "運送方式 : " + receiveDayTextBox.Text);
                template = template.Replace("##boardCountNo##", "貿易形式 : " + boardCountNoTextBox.Text);
                template = template.Replace("##Content##", Html);


                MailMessage message = new MailMessage();

                string aa = textBox2.Text;

                message.To.Add(new MailAddress(aa));
       
                message.Subject = "ShippingOrder";
                message.Body = template;

                //格式為 Html
                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                try
                {
                    client.Send(message);

                    MessageBox.Show("寄信成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }

            
            }

           
      
        }

        private string GetTODO_USERDataSource2()
        {
            System.Data.DataTable dtEvent = GetMenu.GetMail(shippingCodeTextBox.Text);

            string html = string.Empty;
            string DateGroup = string.Empty;
            int f1 = 0;
     
            foreach (DataRow row in dtEvent.Rows)
            {
                f1++;
                string Docentry = Convert.ToString(row["Docentry"]);
                string Dscription = Convert.ToString(row["Dscription"]);
                string Quantity = Convert.ToString(row["Quantity"]);
                string LC = Convert.ToString(row["LC"]);
                string seqno = Convert.ToString(row["seqno"]);
                string Checked = Convert.ToString(row["Checked"]).Trim();
                string shippingcode = Convert.ToString(row["shippingcode"]);
                string RED = Convert.ToString(row["RED"]).Trim();
                System.Data.DataTable K1 = GetINSHIP(shippingcode);
                StringBuilder sb = new StringBuilder();

                if (K1.Rows.Count > 0)
                {
                    if (f1 == 2 && Checked == "True")
                    {
                        f2 = 1;

                    }
                }
                if (RED == "True")
                {
                    f3 = 1;
                }
        
                sb.Append(" 		<tr height='24' style='mso-height-source:userset;height:18.0pt'>");

                if ((Checked != "True" && shippingcode != "1"&&f2==1)||(RED=="True"))
                {
                    sb.Append(" 								<td height='24' width='62' style='height: 18.0pt; width: 47pt;color: #FF0000;' class='style586x'>");
                    sb.Append(" 									" + seqno + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td colspan='2' width='105' style='width: 79pt;color: #FF0000;'class='style587x'>");
                    sb.Append(" 								" + Docentry + "</td>");
                    sb.Append(" 								<td colspan='4' width='286' style='width: 215pt;color: #FF0000;' class='style588x'>");
                    sb.Append(" 								" + Dscription + "</td>");
                    sb.Append(" 								<td width='112' style='width: 84pt;color: #FF0000;' class='style586x'>");
                    sb.Append(" 								" + Quantity + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td width='90' style='width: 68pt;color: #FF0000;' class='style587x'>");
                    sb.Append(" 								" + LC + "</td></tr>");
                }
                else
                {
                    sb.Append(" 								<td height='24' width='62' style='height: 18.0pt; width: 47pt' class='style586x'>");
                    sb.Append(" 									" + seqno + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td colspan='2' width='105' style='width: 79pt'class='style587x'>");
                    sb.Append(" 								" + Docentry + "</td>");
                    sb.Append(" 								<td colspan='4' width='286' style='width: 215pt' class='style588x'>");
                    sb.Append(" 								" + Dscription + "</td>");
                    sb.Append(" 								<td width='112' style='width: 84pt' class='style586x'>");
                    sb.Append(" 								" + Quantity + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td width='90' style='width: 68pt' class='style587x'>");
                    sb.Append(" 								" + LC + "</td></tr>");
                }


                UPDATEINMAIL(shippingCodeTextBox.Text, Convert.ToString(row["DocNum"]), seqno);
                UPDATEINMAIL2(shippingCodeTextBox.Text, Convert.ToString(row["DocNum"]), seqno);
                UPDATEINMAIL3(shippingCodeTextBox.Text);
                html = html + sb.ToString();
            }
            return html;
        }

        private string GetTODO_USERDataSource2SA()
        {
            System.Data.DataTable dtEvent = GetMenu.GetMail(shippingCodeTextBox.Text);

            string html = string.Empty;
            string DateGroup = string.Empty;
            int f1 = 0;
    
            foreach (DataRow row in dtEvent.Rows)
            {
                f1++;
                string Docentry = Convert.ToString(row["Docentry"]);
                string Dscription = Convert.ToString(row["Dscription"]);
                string Quantity = Convert.ToString(row["Quantity"]);
                string LC = Convert.ToString(row["LC"]);
                string seqno = Convert.ToString(row["seqno"]).Trim();
                string Checked = Convert.ToString(row["Checked"]).Trim();
                string shippingcode = Convert.ToString(row["shippingcode"]);
                string RED = Convert.ToString(row["RED"]).Trim();
                System.Data.DataTable K1 = GetINSHIP(shippingcode);
                string CARD = "";
                if (seqno.Trim() != "項次")
                {
                    System.Data.DataTable dtEventSA = GetMenu.GetMailSA(shippingCodeTextBox.Text, seqno);
                    if (dtEventSA.Rows.Count > 0)
                    {
                        CARD = dtEventSA.Rows[0][0].ToString();
                    }
                }
                else
                {
                    CARD = "客戶資料";
                }

                StringBuilder sb = new StringBuilder();

                if (K1.Rows.Count > 0)
                {
                    if (f1 == 2 && Checked == "True")
                    {
                        f2 = 1;

                    }
                }
                if (RED == "True")
                {
                    f3 = 1;
                }

                sb.Append(" 		<tr height='24' style='mso-height-source:userset;height:18.0pt'>");
                if ((Checked != "True" && shippingcode != "1" && f2 == 1)||(RED=="True"))
                {
                    sb.Append(" 								<td height='24' width='62' style='height: 18.0pt; width: 47pt;color: #FF0000;' class='style586x'>");
                    sb.Append(" 									" + seqno + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td colspan='2' width='105' style='width: 79pt;color: #FF0000;'class='style587x'>");
                    sb.Append(" 								" + Docentry + "</td>");
                    sb.Append(" 								<td colspan='4' width='286' style='width: 215pt;color: #FF0000;' class='style588x'>");
                    sb.Append(" 								" + Dscription + "</td>");
                    sb.Append(" 								<td width='112' style='width: 84pt;color: #FF0000;' class='style586x'>");
                    sb.Append(" 								" + Quantity + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td width='90' style='width: 68pt;color: #FF0000;' class='style587x'>");
                    sb.Append(" 								" + LC + "</td>");
                    sb.Append(" 								<td width='90' style='width: 68pt;color: #FF0000;' class='style587x'>");
                    sb.Append(" 								" + CARD + "</td></tr>");
                }
                else {
                    sb.Append(" 								<td height='24' width='62' style='height: 18.0pt; width: 47pt' class='style586x'>");
                    sb.Append(" 									" + seqno + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td colspan='2' width='105' style='width: 79pt'class='style587x'>");
                    sb.Append(" 								" + Docentry + "</td>");
                    sb.Append(" 								<td colspan='4' width='286' style='width: 215pt' class='style588x'>");
                    sb.Append(" 								" + Dscription + "</td>");
                    sb.Append(" 								<td width='112' style='width: 84pt' class='style586x'>");
                    sb.Append(" 								" + Quantity + "<font class='font6'><span style='mso-spacerun:yes'>&nbsp;</span></font></td>");
                    sb.Append(" 								<td width='90' style='width: 68pt' class='style587x'>");
                    sb.Append(" 								" + LC + "</td>");
                    sb.Append(" 								<td width='90' style='width: 68pt' class='style587x'>");
                    sb.Append(" 								" + CARD + "</td></tr>");
                
                }
                html = html + sb.ToString();

                UPDATEINMAIL(shippingCodeTextBox.Text, Convert.ToString(row["DocNum"]), seqno);
                UPDATEINMAIL2(shippingCodeTextBox.Text, Convert.ToString(row["DocNum"]), seqno);
                UPDATEINMAIL3(shippingCodeTextBox.Text);
            }
            return html;
        }
    
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
          
            System.Data.DataTable dt2 = ship.InvoiceD;
            DataRow newCustomersRow = dt2.NewRow();

            int i = invoiceDDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["InvoiceNo"] = invoiceNoTextBox.Text;
            newCustomersRow["InvoiceNo_seq"] = invoiceNo_seqTextBox.Text;
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["INDescription"] = drw["INDescription"];
            newCustomersRow["MarkNos"] = drw["MarkNos"];
            newCustomersRow["InQty"] = drw["InQty"];
            newCustomersRow["UnitPrice"] = drw["UnitPrice"];
            newCustomersRow["Amount"] = drw["Amount"];
            newCustomersRow["CHOPrice"] = drw["CHOPrice"];
            newCustomersRow["CHOAmount"] = drw["CHOAmount"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            newCustomersRow["VISORDER"] = drw["VISORDER"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, invoiceDDataGridView.Rows.Count);

                UPINVOICE();

                this.invoiceDBindingSource.EndEdit();
                this.invoiceDTableAdapter.Update(ship.InvoiceD);
           }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
   

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.PackingListD;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["pLNo"] = pLNoTextBox.Text;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, packingListDDataGridView.CurrentRow.Index);
                UPPACK();
                this.packingListDBindingSource.EndEdit();
                this.packingListDTableAdapter.Update(ship.PackingListD);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.PackingListD;
            DataRow newCustomersRow = dt2.NewRow();

            int i = packingListDDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["pLNo"] = pLNoTextBox.Text;

            newCustomersRow["CNo"] = drw["CNo"];
            newCustomersRow["DescGoods"] = drw["DescGoods"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Net"] = drw["Net"];
            newCustomersRow["Gross"] = drw["Gross"];
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["MeasurmentCM"] = drw["MeasurmentCM"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            newCustomersRow["VISORDER"] = drw["VISORDER"];
            newCustomersRow["SOID"] = drw["SOID"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, packingListDDataGridView.Rows.Count);

                UPPACK();

                this.packingListDBindingSource.EndEdit();
                this.packingListDTableAdapter.Update(ship.PackingListD);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt2 = ship.LADINGD;
            DataRow newCustomersRow = dt2.NewRow();

            int i = invoiceDDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["seqMNo"] = seqMNoTextBox.Text;

            newCustomersRow["ContainerSeals"] = drw["ContainerSeals"];
            newCustomersRow["Packages"] = drw["Packages"];
            newCustomersRow["Description"] = drw["Description"];
            newCustomersRow["Cargo"] = drw["Cargo"];
            newCustomersRow["Measurement"] = drw["Measurement"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            newCustomersRow["SeqNo"] = 100;
          
            try
            {
                dt2.Rows.InsertAt(newCustomersRow, lADINGDDataGridView.Rows.Count);

                for (int j = 0; j <= lADINGDDataGridView.Rows.Count - 2; j++)
                {
                    lADINGDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }


                this.lADINGDBindingSource.EndEdit();
                this.lADINGDTableAdapter.Update(ship.LADINGD);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.LADINGD;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["SeqNo"] = 100;
            newCustomersRow["seqMNo"] = seqMNoTextBox.Text;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, lADINGDDataGridView.CurrentRow.Index);

                for (int j = 0; j <= lADINGDDataGridView.Rows.Count - 2; j++)
                {
                    lADINGDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }


                this.lADINGDBindingSource.EndEdit();
                this.lADINGDTableAdapter.Update(ship.LADINGD);
         
           }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }




        private void receivePlaceTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (Validate() && (invoiceMBindingSource != null))
            {
                // Flag
                bool deleteRow = true;

                // Get row to be deleted
                DataRowView rowView = invoiceMBindingSource.Current as DataRowView;
                if (rowView == null)
                {
                    return;
                }
                ACMEDataSet.ship.InvoiceMRow row =
                   rowView.Row as ACMEDataSet.ship.InvoiceMRow;


                // Check for child rows
                ACMEDataSet.ship.InvoiceDRow[] childRows = row.GetInvoiceDRows();
                if (childRows.Length > 0)
                {
                    DialogResult userChoice = MessageBox.Show("刪除了invoice主資料明細檔也會刪除,確定要刪除?", "Warning", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {
                        // Delete row and child rows
                        foreach (ACMEDataSet.ship.InvoiceDRow childStory in childRows)
                        {

                            childStory.Delete();
                        }
                    }
                    else
                    {
                        deleteRow = false;
                    }
                }
                else
                {
                    DialogResult userChoice = MessageBox.Show("確定要刪除invoice主資料?", "Warning", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {

                    }

                }

                // Delete row?
                if (deleteRow)
                {
                    invoiceMBindingSource.RemoveCurrent();
            
                    try
                    {

                        this.invoiceMBindingSource.EndEdit();
                        this.invoiceDBindingSource.EndEdit();

                        this.invoiceMTableAdapter.Update(ship.InvoiceM);
                        this.invoiceDTableAdapter.Update(ship.InvoiceD);

                        ship.InvoiceM.AcceptChanges();
                        ship.InvoiceD.AcceptChanges();

                        MessageBox.Show("刪除成功");

                    }
                    catch (Exception ex)
                    {

                        GetMenu.InsertLog(fmLogin.LoginID.ToString(), "InvoiceTran3", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                        MessageBox.Show(ex.Message, "刪除錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                    }
                
                }
            }
            
        }

        private void SENDMAIL(string TEMP, string MAIL, int f2, int f3)
        {
            try
            {
                System.Data.DataTable dtdet = GetMenu.Getdet(wHSCODETextBox.Text);
                if (dtdet.Rows.Count > 0)
                {
                    string TM = dtdet.Rows[0][0].ToString();
                    if (!String.IsNullOrEmpty(TM))
                    {
                        int TNN = goalPlaceTextBox.Text.IndexOf(TM);

                        if (TNN == -1)
                        {
                            MessageBox.Show("“目的地” vs “倉庫名稱”不同");
                            return;
                        }
                    }
                }
                string WARR = "";
                if (MAIL == "A")
                {
                    WARR = textBox1.Text;
                }
                if (MAIL == "B")
                {
                    WARR = mail;
                }

               System.Data.DataTable K1 = GetMenu.GetSAME(forecastDayTextBox.Text.Trim(), receivePlaceTextBox.Text.Trim(), goalPlaceTextBox.Text.Trim(), tradeConditionTextBox.Text.Trim(), shippingCodeTextBox.Text);

           

                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();
                    System.Data.DataTable dtPI = GetMenu.GetIN(shippingCodeTextBox.Text);
                    System.Data.DataTable dtPI1 = GetMenu.GetTIFF2AD(shippingCodeTextBox.Text);


                    for (int i = 0; i <= dtPI.Rows.Count - 1; i++)
                    {

                        DataRow dd = dtPI.Rows[i];
                        string docentry = dd["docentry"].ToString();
                        sb.Append(docentry + "/");



                    }
                    for (int i = 0; i <= dtPI1.Rows.Count - 1; i++)
                    {

                        DataRow dd = dtPI1.Rows[i];
                        string ItemCode = dd["MODEL"].ToString();
                        string QTY = dd["QTY"].ToString().Trim();
                        sb2.Append(ItemCode + "*" + QTY + "pcs/");


                    }

                    sb.Remove(sb.Length - 1, 1);
                    sb2.Remove(sb2.Length - 1, 1);

                    string a = sb.ToString();
                    string shi = sb2.ToString();



                    string template;
                    StreamReader objReader;
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    if (boardCountNoTextBox.Text == "進口")
                    {
                        if (receiveDayTextBox.Text.Trim().ToUpper() == "SEA" || receiveDayTextBox.Text.Trim().ToUpper() == "AIR")
                        {
                            if (cardCodeTextBox.Text.Trim() == "S0062" || cardCodeTextBox.Text.Trim() == "S0233")
                            {
                                FileName = lsAppDir + "\\MailTemplates\\進口2.htm";
                            }
                            if (cardCodeTextBox.Text.Trim() == "S0001-AV" || cardCodeTextBox.Text.Trim() == "S0001-CSD" || cardCodeTextBox.Text.Trim() == "S0001-DD")
                            {
                                FileName = lsAppDir + "\\MailTemplates\\進口3.htm";
                            }
                            else
                            {
                                FileName = lsAppDir + "\\MailTemplates\\進口.htm";
                            }
                        }
                        else
                        {
                            FileName = lsAppDir + "\\MailTemplates\\進口.htm";
                        }
                    }
                    else if (boardCountNoTextBox.Text == "出口")
                    {
                        FileName = lsAppDir + "\\MailTemplates\\出口.htm";
                    }
                    else
                    {
                        FileName = lsAppDir + "\\MailTemplates\\SI2.htm";
                    }

          
                    objReader = new StreamReader(FileName);

                    template = objReader.ReadToEnd();
                    objReader.Close();
                    objReader.Dispose();

                    StringWriter writer = new StringWriter();
                    HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

                    StringBuilder se = new StringBuilder();

                    if (K1.Rows.Count > 0)
                    {

                        for (int j = 0; j <= K1.Rows.Count - 1; j++)
                        {

                            DataRow d = K1.Rows[j];

                            string DH = d["倉庫"].ToString();
                            if (!String.IsNullOrEmpty(DH))
                            {
                                se.Append(d["BU"].ToString().Trim() + "_" + d["JOBNO"].ToString() + "_" + DH + "/");
                            }
                            else
                            {
                                se.Append(d["BU"].ToString().Trim() + "_" + d["JOBNO"].ToString() + "/");
                            }


                        }

                        se.Remove(se.Length - 1, 1);

                        MessageBox.Show(se.ToString() + "可併櫃回台");
                        if (MAIL == "B")
                        {
                            template = template.Replace("##anpei2##", "請上排程");
                            template = template.Replace("##anpei3##", "");
                            template = template.Replace("##anpei##", "");
                            template = template.Replace("##anpei4##", "Dear SA,");
                            template = template.Replace("##anpei5##", "");
                            template = template.Replace("##anpei6##", "");
                        }
                        else
                        {
                            template = template.Replace("##anpei2##", "請安排以下出貨，");
                            if (add10CheckBox.Checked)
                            {
                                template = template.Replace("##anpei5##", "本票請申請 AUO貨代免倉期10天");
                            }
                            else
                            {
                                template = template.Replace("##anpei5##", "");
                            }
                            template = template.Replace("##anpei6##", "隨貨請勿放出貨文件。");
                            template = template.Replace("##anpei##", "請安排與" + se + "一起併貨回台");
                            template = template.Replace("##anpei3##", "如有問題，煩請告知。");
                            template = template.Replace("##anpei4##", "Dear,");
                        }

                    }
                    else
                    {
                        if (MAIL == "B")
                        {
                            template = template.Replace("##anpei2##", "請上排程");
                            template = template.Replace("##anpei3##", "");
                            template = template.Replace("##anpei##", "");
                            template = template.Replace("##anpei4##", "Dear SA,");
                            template = template.Replace("##anpei5##", "");
                            template = template.Replace("##anpei6##", "");
                        }
                        else
                        {

                            template = template.Replace("##anpei2##", "請安排以下出貨，");
                            if (add10CheckBox.Checked)
                            {
                                template = template.Replace("##anpei5##", "本票請申請 AUO貨代免倉期10天");
                            }
                            else
                            {
                                template = template.Replace("##anpei5##", "");
                            }
                            template = template.Replace("##anpei6##", "隨貨請勿放出貨文件。");
                            template = template.Replace("##anpei##", "如有問題，煩請告知。");
                            template = template.Replace("##anpei3##", "");
                            template = template.Replace("##anpei4##", "Dear,");
                        }
                    }

               

                    string Html = TEMP;
                //123456
                    System.Data.DataTable dtEvent = GetMenu.Getmark(shippingCodeTextBox.Text);
          
       
                                 StringBuilder sb9 = new StringBuilder();
               

                                 if (dtEvent.Rows.Count > 0)
                                 {
                                     for (int i = 0; i <= dtEvent.Rows.Count - 1; i++)
                                     {

                                         DataRow dd = dtEvent.Rows[i];
                                         sb9.Append(dd["mark"].ToString() + "<br>");
                                     }
                                 }
                
                    template = template.Replace("##PayWay##", "付款方式: " + add3TextBox.Text);
                    template = template.Replace("##date##", "出貨日期: " + lCNOTextBox.Text);
                    template = template.Replace("##Content##", Html);
                    template = template.Replace("##mark##", sb9.ToString());
                    template = template.Replace("##markname##", "嘜頭請加註如下:");
                    template = template.Replace("##G1##", "貨代資料/備註");
                    template = template.Replace("##shippingCode##", "SI NO: " + shippingCodeTextBox.Text);
                    template = template.Replace("##tradeCondition##", "貿易條件 : " + tradeConditionTextBox.Text + " by " + receiveDayTextBox.Text.Trim() + " (" + boardCountNoTextBox.Text.Trim() + ")");
                    template = template.Replace("##closeDay##", "結關日 : " + closeDayTextBox.Text);
                    if (forecastDayTextBox.Text == "")
                    {
                        template = template.Replace("##forecastDay##", "");
                    }
                    else
                    {

                        template = template.Replace("##forecastDay##", "預計開航日 : " + forecastDayTextBox.Text);
                    }

                    if (arriveDayTextBox.Text == "")
                    {
                        template = template.Replace("##arriveDay##", "");
                    }
                    else
                    {

                        template = template.Replace("##arriveDay##", "預計抵達日 : " + arriveDayTextBox.Text);

                    }

                    if (receivePlaceTextBox.Text == "")
                    {
                        template = template.Replace("##receivePlace##", "");
                    }
                    else
                    {

                        template = template.Replace("##receivePlace##", "取貨地 : " + receivePlaceTextBox.Text);

                    }


                    if (goalPlaceTextBox.Text == "")
                    {
                        template = template.Replace("##goalPlace##", "");
                    }
                    else
                    {

                        template = template.Replace("##goalPlace##", "目的地 : " + goalPlaceTextBox.Text);

                    }

                    if (shipmentTextBox.Text == "")
                    {

                        template = template.Replace("##shipment##", "");
                    }
                    else
                    {

                        template = template.Replace("##shipment##", "裝船港 : " + shipmentTextBox.Text);

                    }


                    if (unloadCargoTextBox.Text == "")
                    {
                        template = template.Replace("##unloadCargo##", "");
                    }
                    else
                    {
                        template = template.Replace("##unloadCargo##", "卸貨港 : " + unloadCargoTextBox.Text);

                    }

                    if (boardCountTextBox.Text == "")
                    {
                        template = template.Replace("##boardCount##", "");
                    }
                    else
                    {
                        template = template.Replace("##boardCount##", "20呎 : " + boardCountTextBox.Text);

                    }
                    if (boardDeliverTextBox.Text == "")
                    {
                        template = template.Replace("##boardDeliver##", "");
                    }
                    else
                    {
                        template = template.Replace("##boardDeliver##", "40呎 : " + boardDeliverTextBox.Text);

                    }
                    if (sendGoodsTextBox.Text == "")
                    {
                        template = template.Replace("##sendGoods##", "");
                    }
                    else
                    {
                        template = template.Replace("##sendGoods##", "併櫃/CBM : " + sendGoodsTextBox.Text);

                    }
                    if (pLTSTextBox.Text == "")
                    {
                        template = template.Replace("##PLTS##", "");
                    }
                    else
                    {
                        template = template.Replace("##PLTS##", " = " + pLTSTextBox.Text + "PLTS");

                    }

                    template = template.Replace("##receiveDay##", "運送方式 : " + receiveDayTextBox.Text);
                    template = template.Replace("##boardCountNo##", "貿易形式 : " + boardCountNoTextBox.Text);
                    template = template.Replace("##memo##", memoTextBox1.Text.Replace(System.Environment.NewLine, "<br>"));

                    string h = fmLogin.LoginID.ToString();
                    System.Data.DataTable dt1 = Getemployee(h);

                    if ((dt1.Rows.Count) > 0)
                    {
                        DataRow drw = dt1.Rows[0];
                        string a1 = drw["mobile"].ToString();
                        string a2 = drw["OFFICEEXT"].ToString();
                        template = template.Replace("##eng##", a1);
                        template = template.Replace("##name##", "#" + a2);

                    }
                    template = template.Replace("##eng##", "Lily Lee");
                    template = template.Replace("##TEL##", "02-8791-8368");
                    MailMessage message = new MailMessage();

                 

                    System.Data.DataTable T1 = GetMenu.GetTIFF(shippingCodeTextBox.Text);
                    string COUNTRY = T1.Rows[0][0].ToString();
                    string CITY = T1.Rows[0][1].ToString();
                    int D = tradeConditionTextBox.Text.ToUpper().IndexOf("FOB");
                    int D2 = tradeConditionTextBox.Text.ToUpper().IndexOf("FCA");
                    string sd = "";
                    if (cardNameTextBox.Text.Length > 1)
                    {
                        sd = cardNameTextBox.Text.Substring(0, 2);
                    }
                    string h1 = "";
                    if (D != -1)
                    {
                        h1 = " to " + COUNTRY;
                    }
                    else if (D2 != -1)
                    {
                        h1 = " to " + CITY;
                    }


                    if (MAIL == "A")
                    {
                        message.To.Add(new MailAddress(textBox1.Text));
                    }
                    else
                    {
                        string[] arrurl = mail.Split(new Char[] { ';' });

                        foreach (string i in arrurl)
                        {

                            message.To.Add(i);

                        }
                    }

                    string df = "";
                    if (boardCountNoTextBox.Text == "出口" && sd == "友達")
                    {

                        df = shippingCodeTextBox.Text + "(出口_一般出口)_" + tradeConditionTextBox.Text.Trim() + h1
                        + " by " + receiveDayTextBox.Text + "_PO#" + a + " " + shi;

                    }
                    else
                    {

                        df = shippingCodeTextBox.Text + "(" + boardCountNoTextBox.Text + ")_" + tradeConditionTextBox.Text.Trim() + h1
      + " by " + receiveDayTextBox.Text + "_PO#" + a + " " + shi;


                    }

                    string DF2 = "";
           

                    if (MAIL == "B")
                    {
                        StringBuilder sb3 = new StringBuilder();
                        System.Data.DataTable dtREMARKSA = GetREMARKSA();
                        if (dtREMARKSA.Rows.Count > 0)
                        {
                            for (int i = 0; i <= dtREMARKSA.Rows.Count - 1; i++)
                            {

                                DataRow dd = dtREMARKSA.Rows[i];
                                sb3.Append(dd["REMARK"].ToString() + "/");
                            }

                            sb3.Remove(sb3.Length - 1, 1);
                            DF2 = "for " + sb3.ToString() + "_";
                        }
                    }

                    string DF3 = "";

                    if (add10CheckBox.Checked)
                    {
                        DF3 = "本票請申請 AUO貨代免倉期10天，";
                    }
                    string RED = "";
                    if (f2 == 1 || f3 == 1)
                    {
                        RED = "(REV#紅字處)";
                    }
             
                    message.Subject = RED + DF3 + DF2 + df;
                    message.Body = template;
                    message.IsBodyHtml = true;

                    SmtpClient client = new SmtpClient();
                    client.Send(message);
                    MessageBox.Show("寄信成功");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void button13_Click(object sender, EventArgs e)
        {
            f2 = 0;
            f3 = 0;
            if (add10TextBox.Text != "Checked")
            {
                add10CheckBox.Checked = false;
            }

            SENDMAIL(GetTODO_USERDataSource2(), "A", f2, f3);

            lcInstro1TableAdapter.Fill(ship.LcInstro1, MyID);
        }
        public System.Data.DataTable Getemployee(string name)
        {
            SqlConnection connection = new SqlConnection(strCn02);

            string sql = "select * from acmesql02.dbo.ohem where hometel=@name";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@name", name));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "employee");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["employee"];
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {

            if (Validate() && (packingListMBindingSource != null))
            {
                // Flag
                bool deleteRow = true;

                // Get row to be deleted
                DataRowView rowView = packingListMBindingSource.Current as DataRowView;
                if (rowView == null)
                {
                    return;
                }
                ACMEDataSet.ship.PackingListMRow row =
                   rowView.Row as ACMEDataSet.ship.PackingListMRow;


                // Check for child rows
                ACMEDataSet.ship.PackingListDRow[] childRows = row.GetPackingListDRows();
                if (childRows.Length > 0)
                {
                    DialogResult userChoice = MessageBox.Show("刪除了packing主資料明細檔也會刪除,確定要刪除?", "Warning", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {
                        // Delete row and child rows
                        foreach (ACMEDataSet.ship.PackingListDRow childStory in childRows)
                        {

                            childStory.Delete();
                        }
                    }
                    else
                    {
                        deleteRow = false;
                    }
                }
              

                // Delete row?
                if (deleteRow)
                {

           
               
                    try
                    {
                        packingListMBindingSource.RemoveCurrent();

    
                        this.packingListMBindingSource.EndEdit();
                        this.packingListDBindingSource.EndEdit();


                        this.packingListMTableAdapter.Update(ship.PackingListM);
                        this.packingListDTableAdapter.Update(ship.PackingListD);

                        ship.PackingListM.AcceptChanges();
                        ship.PackingListD.AcceptChanges();

                        MessageBox.Show("刪除成功");

                    }
                    catch (Exception ex)
                    {

                        GetMenu.InsertLog(fmLogin.LoginID.ToString(), "PackingTran3", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                        MessageBox.Show(ex.Message, "刪除錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                    }
                
                  
             
                }

              
            }

       }

        private void button18_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {
                add2TextBox.Text = Convert.ToString(LookupValues[0]);
                add6TextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void 儲存SToolStripButton1_Click(object sender, EventArgs e)
        {
      
            try
            {

                this.Validate();
                this.lcInstroBindingSource.EndEdit();
                this.lcInstro1BindingSource.EndEdit();



                this.lcInstroTableAdapter.Update(ship.LcInstro);
                this.lcInstro1TableAdapter.Update(ship.LcInstro1);

                ship.LcInstro.AcceptChanges();
                ship.LcInstro1.AcceptChanges();

                MessageBox.Show("儲存成功");

            }
            catch (Exception ex)
            {
               
                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


            }
    

    
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (Validate() && (lcInstroBindingSource != null))
            {
                // Flag
                bool deleteRow = true;

                // Get row to be deleted
                DataRowView rowView = lcInstroBindingSource.Current as DataRowView;
                if (rowView == null)
                {
                    return;
                }
                ACMEDataSet.ship.LcInstroRow row =
                   rowView.Row as ACMEDataSet.ship.LcInstroRow;


                // Check for child rows
            //      ship.InvoiceDRow[] childRows = row.GetInvoiceDRows();
                ACMEDataSet.ship.LcInstro1Row[] childRows = row.GetLcInstro1Rows();
                if (childRows.Length > 0)
                {
                    DialogResult userChoice = MessageBox.Show("刪除採購Instruction主資料明細檔也會刪除,確定要刪除?", "Warning", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {
                        // Delete row and child rows
                        foreach (ACMEDataSet.ship.LcInstro1Row childStory in childRows)
                        {

                            childStory.Delete();
                        }
                    }
                    else
                    {
                        deleteRow = false;
                    }
                }
                else
                {
                    DialogResult userChoice = MessageBox.Show("確定要刪除採購Instruction主資料?", "Warning", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {

                    }

                }

                // Delete row?
                if (deleteRow)
                {
                    lcInstroBindingSource.RemoveCurrent();
                    lcInstroBindingSource.EndEdit();
                }
            }
            
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetMenu.GetLcInstro1(shippingCodeTextBox.Text);

            try
            {

                if (shipping_ItemDataGridView.Rows.Count < 1)
                {

                    MessageBox.Show("項目/料號為空值");

                    return;

                }
                iTEMSCheckBox.Checked = false;
                add10CheckBox.Checked = false;
                    string NumberName = "LI" + DateTime.Now.ToString("yyyyMMdd");
                    string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                    docNumTextBox.Text = NumberName + AutoNum;
                    
                    lcInstroBindingSource.EndEdit();
                    System.Data.DataTable dt3 = Getshipitem(shippingCodeTextBox.Text, 1, "");

                    System.Data.DataTable dt4 = ship.LcInstro1;

                    if (shipping_ItemDataGridView.Rows.Count > 1 && lcInstro1DataGridView.Rows.Count < 2)
                    {
                        DataGridViewRow row;
                        for (int i = 0; i <= shipping_ItemDataGridView.Rows.Count - 2; i++)
                        {
                            row = shipping_ItemDataGridView.Rows[i];
                            DataRow drw2 = dt4.NewRow();
                            string LINE = row.Cells["linenum"].Value.ToString();
                            if (String.IsNullOrEmpty(LINE))
                            {
                                LINE = "0";
                            }
                            drw2["ShippingCode"] = shippingCodeTextBox.Text;
                            drw2["DocNum"] = docNumTextBox.Text;
                            drw2["SeqNo"] = i.ToString();
                            drw2["Docentry"] = row.Cells["Docentry"].Value.ToString();
                            drw2["linenum"] = LINE;
                            drw2["ItemCode"] = row.Cells["ItemCode"].Value.ToString(); 
                            drw2["Dscription"] = row.Cells["Dscription"].Value.ToString(); 
                            drw2["Quantity"] = row.Cells["Quantity"].Value.ToString(); 
                            drw2["ItemPrice"] = row.Cells["ItemPrice"].Value.ToString(); 
                            drw2["ItemAmount"] = row.Cells["ItemAmount"].Value.ToString(); 
                            dt4.Rows.Add(drw2);

                        }

                    }


            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message);
            }
           
        }



       

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.LcInstro1;
            DataRow newCustomersRow = dt2.NewRow();

            int i = lcInstro1DataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["docnum"] = drw["docnum"];
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["Docentry"] = drw["Docentry"];
            newCustomersRow["ItemCode"] = drw["ItemCode"];
            newCustomersRow["Dscription"] = drw["Dscription"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["ItemPrice"] = drw["ItemPrice"];
            newCustomersRow["ItemAmount"] = drw["ItemAmount"];
            newCustomersRow["LC"] = drw["LC"];
            try
            {
                dt2.Rows.InsertAt(newCustomersRow, lcInstro1DataGridView.Rows.Count);


                this.lcInstro1BindingSource.EndEdit();
                this.lcInstro1TableAdapter.Update(ship.LcInstro1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

  

        private void goalPlaceTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }
        }

        private void shipmentTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }
        }

        private void unloadCargoTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }
        }

        private void closeDayTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }
        }

        private void tradeConditionTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && e.Shift)
            {
                UserValueDialog aForm = new UserValueDialog();
                aForm.FormID1 = this.GetType().ToString();
                aForm.ObjID1 = ((System.Windows.Forms.TextBox)sender).Name;
                if (aForm.ShowDialog() == DialogResult.OK)
                {
                    ((System.Windows.Forms.TextBox)sender).Text = aForm.SelectValue;
                }
            }   
        }

  


     


      

   


     
  
     
        private void buCardcodeCheckBox_Click(object sender, EventArgs e)
        {
            if (buCardcodeCheckBox.Checked)
            {

                buCardnameTextBox.Text = DateTime.Now.ToString("yyyyMMdd");
                quantityTextBox.Text = "已結";
            }
            else
            {
                buCardnameTextBox.Text = "";
                quantityTextBox.Text = "未結";
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (Validate() && (lcInstroBindingSource != null))
            {
                // Flag
                bool deleteRow = true;

                DataRowView rowView = lcInstroBindingSource.Current as DataRowView;
                if (rowView == null)
                {
                    return;
                }
                ACMEDataSet.ship.LcInstroRow row =
                   rowView.Row as ACMEDataSet.ship.LcInstroRow;


                // Check for child rows
                //      ship.InvoiceDRow[] childRows = row.GetInvoiceDRows();
                ACMEDataSet.ship.LcInstro1Row[] childRows = row.GetLcInstro1Rows();
                if (childRows.Length > 0)
                {
                
                        // Delete row and child rows
                        foreach (ACMEDataSet.ship.LcInstro1Row childStory in childRows)
                        {

                            childStory.Delete();
                        }
                    
                  
                }
                else
                {
                   
                }

              
            }
            System.Data.DataTable dt3 = GetMenu.GetLcInstro1(shippingCodeTextBox.Text);

            System.Data.DataTable dt4 = ship.LcInstro1;

    
                
                for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                {
                    DataRow drw = dt3.Rows[i];
                    DataRow drw2 = dt4.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["DocNum"] = docNumTextBox.Text;
                    drw2["SeqNo"] = i.ToString();
                    drw2["Docentry"] = drw["Docentry"].ToString();
                    drw2["ItemCode"] = drw["ItemCode"].ToString();
                    drw2["linenum"] = drw["linenum"];
                    drw2["Dscription"] = drw["Dscription"].ToString();
                    drw2["Quantity"] = drw["Quantity"].ToString();
                    drw2["ItemPrice"] = drw["ItemPrice"].ToString();
                    drw2["ItemAmount"] = drw["ItemAmount"].ToString();
                    drw2["RED"] = drw["RED"].ToString();
                    dt4.Rows.Add(drw2);

                

            }
            
        }



   

   







        private void packingListDDataGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                System.Data.DataTable f = ship.PackingListD;

                DataRowView row = (DataRowView)packingListDBindingSource.Current;
                if (String.IsNullOrEmpty(row["shippingcode"].ToString()))
                {
                    MessageBox.Show("系統異常，請離開程式通知MIS處理");
                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Packing增加一列Current", "單號" + shippingCodeTextBox.Text, DateTime.Now.ToString("yyyyMMddHHmmss"));
                    UpdatePacking();
                    return;

                }

           
                try
                {

               
                    this.packingListDBindingSource.EndEdit();
                    this.packingListDTableAdapter.Update(ship.PackingListD);
                    ship.PackingListD.AcceptChanges();
             

                

                }
                catch (Exception ex)
                {

                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "PackingTran3", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                }
    


   
            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Packing增加一列", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));
            }
        }



        private void InsertPacking(string ShippingCode,string PLNo,string SeqNo,string PackageNo,string CNo,string DescGoods,string Quantity,string Net,string Gross,string MeasurmentCM)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO packinglistd (ShippingCode,PLNo,SeqNo,PackageNo,CNo,DescGoods,Quantity,Net,Gross,MeasurmentCM) VALUES(@ShippingCode,@PLNo,@SeqNo,@PackageNo,@CNo,@DescGoods,@Quantity,@Net,@Gross,@MeasurmentCM)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@PackageNo", PackageNo));
            command.Parameters.Add(new SqlParameter("@CNo", CNo));
            command.Parameters.Add(new SqlParameter("@DescGoods", DescGoods));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Net", Net));
            command.Parameters.Add(new SqlParameter("@Gross", Gross));
            command.Parameters.Add(new SqlParameter("@MeasurmentCM", MeasurmentCM));

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

        private void DeletePacking(string ShippingCode, string PLNo)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" delete packinglistd where ShippingCode=@ShippingCode and PLNo=@PLNo ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));

      
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
       
        private void button16_Click(object sender, EventArgs e)
        {

            try
            {
     

                Validate();

                packingListMBindingSource.EndEdit();
                packingListMTableAdapter.Update(ship.PackingListM);
                ship.PackingListM.AcceptChanges();

                
                MessageBox.Show("儲存成功");

     
            }

            catch (Exception ex)
            {

                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Packing修改總計", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));
        
            }
     
        }



        private void UpdatePacking()
        {
            DeletePacking(shippingCodeTextBox.Text, pLNoTextBox.Text);
            string seqno;
            string PackageNo;
            string CNo;
            string desc;
            string qty;
            string net;
            string gro;
            string mea;
            for (int i = 0; i <= packingListDDataGridView.Rows.Count - 2; i++)
            {


                seqno = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn44"].Value.ToString();
                PackageNo = packingListDDataGridView.Rows[i].Cells["PackageNo"].Value.ToString();
                CNo = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn46"].Value.ToString();
                desc = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn47"].Value.ToString();
                qty = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn48"].Value.ToString();
                net = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn49"].Value.ToString();
                gro = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn50"].Value.ToString();
                mea = packingListDDataGridView.Rows[i].Cells["dataGridViewTextBoxColumn51"].Value.ToString();
                InsertPacking(shippingCodeTextBox.Text, pLNoTextBox.Text, seqno, PackageNo, CNo, desc, qty, net, gro, mea);

            }
        }

        private void invoiceDDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void packingListDDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void lADINGDDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void shipping_ItemDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
               try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        public static System.Data.DataTable GetSHPCAR(string JOBNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
  
            sb.Append(" SELECT DISTINCT [FILENAME] CARNAME,[PATH] CARPATH   FROM shipping_CAR T0");
            sb.Append(" LEFT JOIN shipping_CAR2 T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" LEFT JOIN shipping_CARdownload T2 ON(T0.SHIPPINGCODE=T2.SHIPPINGCODE)");
            sb.Append("  WHERE ISNULL([FILENAME] ,'') <> '' AND JOBNO=@JOBNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JOBNO", JOBNO));
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
        public static System.Data.DataTable GetSHPCAR2(string JOBNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT T0.CARTYPE 類型,T0.SHIPPINGCODE 併單工單,T2.CARSIZE 車型,T2.CARSIZEL 長,T2.CARSIZEW 寬,T2.CARSIZEH 高,T2.CARTYPE 廠商    FROM shipping_CAR T0 ");
            sb.Append("               LEFT JOIN shipping_CAR2 T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("               LEFT JOIN shipping_CAR4 T2 ON(T0.SHIPPINGCODE=T2.SHIPPINGCODE) ");
            sb.Append("  WHERE T1.JOBNO=@JOBNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JOBNO", JOBNO));
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

        public static System.Data.DataTable GetSHPCAR3(string JOBNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                          SELECT DISTINCT T0.SHIPPINGCODE 併單工單,T0.CARTYPE 類型   FROM shipping_CAR T0  ");
            sb.Append("                             LEFT JOIN shipping_CAR2 T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append(" WHERE JOBNO=@JOBNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JOBNO", JOBNO));
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
       
        public static System.Data.DataTable GetFEE(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CardName 供應商,SubCompany 子公司,DocDate 日期,SAP SAP單號,ITEM 費用名稱,Amount 金額,DocCur 幣別,DocCur1 匯率,FeeCheck,ID FROM dbo.Shipping_Fee   where ShippingCode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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
        public static System.Data.DataTable GetAP(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT 'quantity :'+isnull(cast(I.quantity as varchar),'')+' net :'+isnull(cast(I.net as varchar),'')+' gross :'+isnull(cast(I.gross as varchar),'')+' package :'+isnull(cast(I.saytotal as varchar),'')+' 20呎 :'+isnull(cast(b.boardCount as varchar),'')+' 40呎 :'+isnull(cast(b.boardDeliver as varchar),'')+' 併櫃/CBM :'+isnull(b.sendGoods,'') SHIP");
            sb.Append("              FROM PackingListM AS i ");
            sb.Append("              left join shipping_main b on (i.shippingcode=b.shippingcode)");
            sb.Append("               where i.[shippingcode]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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

        public static System.Data.DataTable GetINVO(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT AMOUNT  FROM INVOICEM T0 ");
    sb.Append(" LEFT JOIN INVOICED T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.INVOICENO=T1.INVOICENO AND T0.INVOICENO_SEQ=T1.INVOICENO_SEQ) ");
    sb.Append("  where T0.[shippingcode]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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

        public static System.Data.DataTable Getfee(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT cast(AMOUNT AS DECIMAL(15,5)) AMOUNT FROM shipping_fee T0 where T0.[shippingcode]=@shippingcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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


        public static System.Data.DataTable RETAB()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CELLNAME FROM USERSSHIP WHERE USERID='EEP' ");
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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

        public static System.Data.DataTable RETAB2(string USERID, string CELLNAME)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CELLNAME FROM USERSSHIP WHERE USERID='EEP' AND CELLNAME NOT IN (SELECT CELLNAME  FROM USERSSHIP WHERE USERID=@USERID ) AND CELLNAME=@CELLNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERID", USERID));
            command.Parameters.Add(new SqlParameter("@CELLNAME", CELLNAME));
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



        private void download2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
          
                try
                {

                    DataGridView dgv = (DataGridView)sender;

                    if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                    {

                        System.Data.DataTable dt1 = ship.Download2;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];

                        string aa = drw["path"].ToString();
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = drw["filename"].ToString();
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;
           
                        System.IO.File.Copy(aa, NewFileName, true);
                        System.Diagnostics.Process.Start(NewFileName);
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

        private void button25_Click_1(object sender, EventArgs e)
        {
            try
            {
            string f = "c";
            string[] filebType = Directory.GetDirectories(DIR);
            string dd = DateTime.Now.ToString("yyyyMM");
            string tt = DIR + dd;
            foreach (string fileaSize in filebType)
            {

                if (fileaSize == tt)
                {
                    f = "d";

                }

            }
            if (f == "c")
            {
                Directory.CreateDirectory(tt);
            }
       
                string server = DIR + dd + "//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download2(filename);

                if (dt2.Rows.Count > 0)
                {
                      string G1 = dt2.Rows[0]["filename"].ToString().Replace(" ", "").ToUpper().Trim();
                      string BAU = add9TextBox.Text.Replace(" ", "").ToUpper().Trim();
                      int F1  = G1.IndexOf(BAU);
                      if (F1 == -1)
                      {

                          MessageBox.Show("檔案名稱重複,請修改檔名");
                      }
                      else
                      {
                          if (result == DialogResult.OK)
                          {

                              string file = opdf.FileName;
                              bool FF1 = getrma.UploadFile(file, server, false);
                              if (FF1 == false)
                              {
                                  return;
                              }

                              System.Data.DataTable dt1 = ship.Download2;

                              DataRow drw = dt1.NewRow();
                              drw["ShippingCode"] = shippingCodeTextBox.Text;
                              drw["seq"] = (download2DataGridView.Rows.Count).ToString();
                              drw["filename"] = filename;
                              string de = DateTime.Now.ToString("yyyyMM") + "\\";

                              drw["path"] = PATH + de + filename;


                              dt1.Rows.Add(drw);

                              download2BindingSource.MoveFirst();

                              for (int i = 0; i <= download2BindingSource.Count - 1; i++)
                              {
                                  DataRowView rowd = (DataRowView)download2BindingSource.Current;

                                  rowd["seq"] = i;

                                  download2BindingSource.EndEdit();

                                  download2BindingSource.MoveNext();
                              }

                              this.download2BindingSource.EndEdit();
                              this.download2TableAdapter.Update(ship.Download2);
                              ship.Download2.AcceptChanges();

                              MessageBox.Show("上傳成功");
                          }
                      }
                }
                else
                {
                    if (result == DialogResult.OK)
                    {

                        string file = opdf.FileName;
                        bool FF1 = getrma.UploadFile(file, server, false);
                        if (FF1 == false)
                        {
                            return;
                        }
                        System.Data.DataTable dt1 = ship.Download2;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (download2DataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
            
                        drw["path"] = PATH + de + filename;

      
                        dt1.Rows.Add(drw);

                        download2BindingSource.MoveFirst();

                        for (int i = 0; i <= download2BindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)download2BindingSource.Current;

                            rowd["seq"] = i;



                            download2BindingSource.EndEdit();

                            download2BindingSource.MoveNext();
                        }

                        this.download2BindingSource.EndEdit();
                        this.download2TableAdapter.Update(ship.Download2);
                        ship.Download2.AcceptChanges();

                        MessageBox.Show("上傳成功");
                    }

                }
            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "不可下載檔案上傳", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));
                if (fmLogin.LoginID.ToString().ToUpper() != "EVAHSU")
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }



        private void tabPage6_Enter_1(object sender, EventArgs e)
        {
            System.Data.DataTable dtm = GetMenu.getaa(shippingCodeTextBox.Text);

            if (dtm.Rows.Count == 0 || shipping_ItemDataGridView.Rows.Count == 1 )
            {
                MessageBox.Show("請先儲存主檔或項目/料號沒資料");
             
                tabControl1.SelectedIndex = 0;

            }
        }


        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (Validate() && (lADINGMBindingSource != null))
            {
                // Flag
                bool deleteRow = true;

                // Get row to be deleted
                DataRowView rowView = lADINGMBindingSource.Current as DataRowView;
                if (rowView == null)
                {
                    return;
                }
                ACMEDataSet.ship.LADINGMRow row =
                   rowView.Row as ACMEDataSet.ship.LADINGMRow;


                // Check for child rows
                ACMEDataSet.ship.LADINGDRow[] childRows = row.GetLADINGDRows();
                if (childRows.Length > 0)
                {
                    DialogResult userChoice = MessageBox.Show("刪除了LADING主資料明細檔也會刪除,確定要刪除?", "Warning", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (userChoice == DialogResult.Yes)
                    {
                        // Delete row and child rows
                        foreach (ACMEDataSet.ship.LADINGDRow childStory in childRows)
                        {

                            childStory.Delete();
                        }
                    }
                    else
                    {
                        deleteRow = false;
                    }
                }


                if (deleteRow)
                {


                    SqlTransaction tx = null;
                    try
                    {
                        lADINGMBindingSource.RemoveCurrent();

                   
                        this.lADINGMBindingSource.EndEdit();
                        this.lADINGDBindingSource.EndEdit();

                    

                        this.lADINGMTableAdapter.Update(ship.LADINGM);
                        this.lADINGDTableAdapter.Update(ship.LADINGD);

                 

                        MessageBox.Show("刪除成功");

                    }
                    catch (Exception ex)
                    {
                       

                        MessageBox.Show(ex.Message, "刪除錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);


                    }
                  

                }


            }
        }

        private void shipping_ItemDataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            shipping_ItemDataGridView.ImeMode = ImeMode.Off;
        }

        private void invoiceDDataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            invoiceDDataGridView.ImeMode = ImeMode.Off;
        }

        private void packingListDDataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            packingListDDataGridView.ImeMode = ImeMode.Off;
        }

        private void lcInstro1DataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            lcInstro1DataGridView.ImeMode = ImeMode.Off;
        }



        private object[] GetCardList(string aa, string dg,string dg2,string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "u_acme_workday", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "銷售單號", "倉庫名稱", "排程日期", "工作天數", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT distinct cast(T0.docentry as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,u_acme_workday,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','')+isnull(u_acme_workday,'')  'KEY',T0.docentry 單號 ");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM ORDR T0 inner join rdr1 T1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            sb.Append(" LEFT JOIN ( SELECT T1.序號,T0.ADDRESS FROM CRD1 T0");
            sb.Append("         LEFT JOIN ( SELECT  RANK() OVER (ORDER BY ADDRESS ) AS 序號,ADDRESS,CARDCODE FROM CRD1 T0");
            sb.Append("        WHERE T0.ADRESTYPE='S' AND T0.CARDCODE='" + aa + "' ) T1 ");
            sb.Append("            ON ( T0.ADDRESS=T1.ADDRESS ) ");
            sb.Append("      WHERE T0.ADRESTYPE='S'  AND T0.CARDCODE='" + aa + "'  ) T3");
            sb.Append(" ON (  T3.ADDRESS= T0.SHIPTOCODE)");
            sb.Append("  WHERE 1=1 "); 
            
            if (dg == "")
            {
                sb.Append(" AND t1.linestatus='O' ");

            }

            if (dg2 != "check")
            {
                sb.Append(" AND T0.cardcode='" + aa + "'  ");

            }
            else
            {
                sb.Append(" AND T0.cardname like '%" + bb + "%'  ");

            }

            sb.Append(" order by t0.docentry desc ");

      

            MultiValueDialog2 dialog = new MultiValueDialog2();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            if (SOL == 1)
            {
                dialog.KeyFieldName = "銷售單號";
            }
            else
            {
                dialog.KeyFieldName = "KEY";
            }
            dialog.SqlScript = sb.ToString();

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;
                    hh = dialog.qg;
                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }
    
        private object[] GetCardListORIN(string aa, string dg2, string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "AR貸項單號", "倉庫名稱", "排程日期", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct cast(T0.docentry as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') 'KEY',T0.docentry 單號");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM ORIN T0 inner join RIN1 T1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            sb.Append(" LEFT JOIN ( SELECT T1.序號,T0.ADDRESS FROM CRD1 T0");
            sb.Append("         LEFT JOIN ( SELECT  RANK() OVER (ORDER BY ADDRESS ) AS 序號,ADDRESS,CARDCODE FROM CRD1 T0");
            sb.Append("        WHERE T0.ADRESTYPE='S' AND T0.CARDCODE='" + aa + "' ) T1 ");
            sb.Append("            ON ( T0.ADDRESS=T1.ADDRESS ) ");
            sb.Append("      WHERE T0.ADRESTYPE='S'  AND T0.CARDCODE='" + aa + "'  ) T3");
            sb.Append(" ON (  T3.ADDRESS= T0.SHIPTOCODE)");
            sb.Append("  WHERE 1=1 ");


            if (dg2 != "check")
            {
                sb.Append(" AND T0.cardcode='" + aa + "'  ");

            }
            else
            {
                sb.Append(" AND T0.cardname like '%" + bb + "%'  ");

            }

            sb.Append(" order by t0.docentry desc ");



            MultiValueDialog2 dialog = new MultiValueDialog2();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            dialog.KeyFieldName = "KEY";
            dialog.SqlScript = sb.ToString();

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;
                    hh = dialog.qg;
                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }
        private object[] GetCardListORINT(string aa, string dg2, string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "AR貸項草稿單號", "倉庫名稱", "排程日期", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct cast(T0.DOCNUM as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') 'KEY',T0.docentry 單號");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM ODRF T0 inner join DRF1 T1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            sb.Append(" LEFT JOIN ( SELECT T1.序號,T0.ADDRESS FROM CRD1 T0");
            sb.Append("         LEFT JOIN ( SELECT  RANK() OVER (ORDER BY ADDRESS ) AS 序號,ADDRESS,CARDCODE FROM CRD1 T0");
            sb.Append("        WHERE T0.ADRESTYPE='S' AND T0.CARDCODE='" + aa + "' ) T1 ");
            sb.Append("            ON ( T0.ADDRESS=T1.ADDRESS ) ");
            sb.Append("      WHERE T0.ADRESTYPE='S'  AND T0.CARDCODE='" + aa + "'  ) T3");
            sb.Append(" ON (  T3.ADDRESS= T0.SHIPTOCODE)");
            sb.Append("  WHERE 1=1 AND T0.OBJTYPE=14 ");


            if (dg2 != "check")
            {
                sb.Append(" AND T0.cardcode='" + aa + "'  ");

            }
            else
            {
                sb.Append(" AND T0.cardname like '%" + bb + "%'  ");

            }

            sb.Append(" order by t0.docentry desc ");



            MultiValueDialog2 dialog = new MultiValueDialog2();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            dialog.KeyFieldName = "KEY";
            dialog.SqlScript = sb.ToString();

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;
                    hh = dialog.qg;
                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }
        private object[] GetCardListORPC(string aa, string dg2, string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "AP貸項單號", "倉庫名稱", "排程日期", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct cast(T0.docentry as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') 'KEY',T0.docentry 單號");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM ORPC T0 inner join RPC1 T1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            sb.Append(" LEFT JOIN ( SELECT T1.序號,T0.ADDRESS FROM CRD1 T0");
            sb.Append("         LEFT JOIN ( SELECT  RANK() OVER (ORDER BY ADDRESS ) AS 序號,ADDRESS,CARDCODE FROM CRD1 T0");
            sb.Append("        WHERE T0.ADRESTYPE='S' AND T0.CARDCODE='" + aa + "' ) T1 ");
            sb.Append("            ON ( T0.ADDRESS=T1.ADDRESS ) ");
            sb.Append("      WHERE T0.ADRESTYPE='S'  AND T0.CARDCODE='" + aa + "'  ) T3");
            sb.Append(" ON (  T3.ADDRESS= T0.SHIPTOCODE)");
            sb.Append("  WHERE 1=1 ");


            if (dg2 != "check")
            {
                sb.Append(" AND T0.cardcode='" + aa + "'  ");

            }
            else
            {
                sb.Append(" AND T0.cardname like '%" + bb + "%'  ");

            }

            sb.Append(" order by t0.docentry desc ");



            MultiValueDialog2 dialog = new MultiValueDialog2();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            dialog.KeyFieldName = "KEY";
            dialog.SqlScript = sb.ToString();

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;
                    hh = dialog.qg;
                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }


        private object[] GetCardListORPD(string aa, string dg2, string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "採購退貨單號", "倉庫名稱", "排程日期", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct cast(T0.docentry as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') 'KEY',T0.docentry 單號");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM ORPD T0 inner join RPD1 T1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            sb.Append(" LEFT JOIN ( SELECT T1.序號,T0.ADDRESS FROM CRD1 T0");
            sb.Append("         LEFT JOIN ( SELECT  RANK() OVER (ORDER BY ADDRESS ) AS 序號,ADDRESS,CARDCODE FROM CRD1 T0");
            sb.Append("        WHERE T0.ADRESTYPE='S' AND T0.CARDCODE='" + aa + "' ) T1 ");
            sb.Append("            ON ( T0.ADDRESS=T1.ADDRESS ) ");
            sb.Append("      WHERE T0.ADRESTYPE='S'  AND T0.CARDCODE='" + aa + "'  ) T3");
            sb.Append(" ON (  T3.ADDRESS= T0.SHIPTOCODE)");
            sb.Append("  WHERE 1=1 ");


            if (dg2 != "check")
            {
                sb.Append(" AND T0.cardcode='" + aa + "'  ");

            }
            else
            {
                sb.Append(" AND T0.cardname like '%" + bb + "%'  ");

            }

            sb.Append(" order by t0.docentry desc ");



            MultiValueDialog2 dialog = new MultiValueDialog2();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            dialog.KeyFieldName = "KEY";
            dialog.SqlScript = sb.ToString();

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;
                    hh = dialog.qg;
                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }
      
        private void button27_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetJoy1("收貨地");

            if (LookupValues != null)
            {
                receivePlaceTextBox.Text = Convert.ToString(LookupValues[0]);
               

            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetJoy1("目的地");

            if (LookupValues != null)
            {
                goalPlaceTextBox.Text = Convert.ToString(LookupValues[0]);


            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetJoy1("卸貨港");

            if (LookupValues != null)
            {
                unloadCargoTextBox.Text = Convert.ToString(LookupValues[0]);


            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetJoy1("裝船港");

            if (LookupValues != null)
            {
                shipmentTextBox.Text = Convert.ToString(LookupValues[0]);


            }
        }



        private void tabPage4_Enter(object sender, EventArgs e)
        {
            if (COPY == 0)
            {
                System.Data.DataTable dtm = GetMenu.getaa(shippingCodeTextBox.Text);

                if (dtm.Rows.Count.ToString() == "0" || shipping_ItemDataGridView.Rows.Count == 1)
                {
                    MessageBox.Show("請先儲存主檔或項目/料號沒資料");

                    tabControl1.SelectedIndex = 0;
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            AP_WHS_List form = new AP_WHS_List();


            if (form.ShowDialog() == DialogResult.OK)
            {
        

            }
        }




        public static System.Data.DataTable GetShipping_WHS()
        {
            SqlConnection con = globals.Connection;
            string sql = "SELECT WHSCODE DataValue FROM Shipping_WHS";
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            wHSCODETextBox.Text = comboBox2.Text;
            if (GetWH().Rows.Count > 0)
            {
                memoTextBox1.Text = GetWH().Rows[0][0].ToString();
            }
        }



        private System.Data.DataTable GetWHSTOCK()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE JOBTITLE='船務倉管'  AND ISNULL(TERMDATE,'') =''   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", comboBox2.Text));


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
        private System.Data.DataTable GetWHSA()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT HOMETEL FROM OHEM WHERE JOBTITLE='業助' AND ISNULL(TERMDATE,'') = '' ORDER BY HOMETEL ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", comboBox2.Text));


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
        private System.Data.DataTable GetWHSHIP()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE JOBTITLE='船務'  AND ISNULL(TERMDATE,'') =''   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", comboBox2.Text));


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
       
        private System.Data.DataTable GetWH()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT DESCRIPTION  FROM Shipping_WHS WHERE WHSCODE=@AA ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", comboBox2.Text));


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
        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {

            System.Data.DataTable dt4 = GetShipping_WHS();


            comboBox2.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }
        private void button15F(object sender, EventArgs e)
        {



            tabControl1.SelectedIndex = 0;


            System.Data.DataTable dt1 = null;
            dt1 = GetSTKBILLNO(pinoTextBox.Text);


            System.Data.DataTable dt2 = ship.Shipping_Item;


            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["seqNo"] = "0";
                drw2["Docentry"] = drw["Docnum"];
                drw2["ItemCode"] = drw["ItemCode"];
                drw2["Dscription"] = drw["Dscription"];
                drw2["PiNo"] = "";
                drw2["ItemRemark"] = "進貨憑單";
                drw2["Quantity"] = drw["數量"];
                drw2["CHOPrice"] = drw["單價"];
                drw2["linenum"] = drw["ROWNO"];
                drw2["CHOAmount"] = drw["金額"];
                //  drw2["CHOMemo"] = drw["備註"];

                shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                dt2.Rows.Add(drw2);


                sAMEMOTextBox.Text = drw["備註"].ToString();
            }


            for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
            {
                shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
            }

            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();



        }
        private void button15F2(object sender, EventArgs e)
        {



            tabControl1.SelectedIndex = 0;


            System.Data.DataTable dt1 = null;
            dt1 = GetSTKBILLNO2(pinoTextBox.Text);


            System.Data.DataTable dt2 = ship.Shipping_Item;


            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["seqNo"] = "0";
                drw2["Docentry"] = drw["Docnum"];
                drw2["ItemCode"] = drw["ItemCode"];
                drw2["Dscription"] = drw["Dscription"];
                drw2["PiNo"] = "";
                drw2["ItemRemark"] = "銷貨憑單";
                drw2["Quantity"] = drw["數量"];
                drw2["CHOPrice"] = drw["單價"];
                drw2["linenum"] = drw["ROWNO"];
                drw2["CHOAmount"] = drw["金額"];
                //  drw2["CHOMemo"] = drw["備註"];

                shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                dt2.Rows.Add(drw2);


                sAMEMOTextBox.Text = drw["備註"].ToString();
            }


            for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
            {
                shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
            }

            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();



        }
        private void button15_Click(object sender, EventArgs e,string FLAG)
        {
            if (cardCodeTextBox.Text == "")
            {
                MessageBox.Show("請輸入客戶編號");

                return;
            }
            string dg = "";
            if (checkBox1.Checked)
            {
                dg = "check";
            }
            else
            {
                dg = "0";
            }
            object[] LookupValues = null;
            LookupValues = GetMenu.GetowtrCHOAD(cardCodeTextBox.Text, dg, FLAG);
            if (LookupValues != null)
            {
                tabControl1.SelectedIndex = 0;

                string docentry = Convert.ToString(LookupValues[0]);
                pinoTextBox.Text = docentry;

                System.Data.DataTable dt1 = null;

                dt1 = GetCHO(docentry,FLAG);

           

                System.Data.DataTable dt2 = ship.Shipping_Item;
             

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["seqNo"] = "0";
                    drw2["Docentry"] = drw["Docnum"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["PiNo"] = "";
                    if (FLAG == "2")
                    {
                        drw2["ItemRemark"] = "宇豐銷售";
                    }
                    if (FLAG == "4")
                    {
                        drw2["ItemRemark"] = "宇豐採購";
                    }
                    drw2["Quantity"] = drw["數量"];
                    drw2["CHOPrice"] = drw["單價"];
                    drw2["linenum"] = drw["ROWNO"];
                    drw2["CHOAmount"] = drw["金額"];
               
                    drw2["ItemPrice"] = drw["單價"];
                    drw2["ItemAmount"] = drw["金額"];
            
                    shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                    dt2.Rows.Add(drw2);


                    sAMEMOTextBox.Text = drw["備註"].ToString();
                    kPIYESNOTextBox.Text = drw["PRJ"].ToString();
                }


                for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
                {
                    shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
                }

                shipping_mainBindingSource.EndEdit();
                shipping_ItemBindingSource.EndEdit();
            }

            
        }
        private void button15TIOJEN_Click(object sender, EventArgs e)
        {



            tabControl1.SelectedIndex = 0;


            System.Data.DataTable dt1 = null;
            dt1 = GetCHOTIOJEN(pinoTextBox.Text);


            System.Data.DataTable dt2 = ship.Shipping_Item;


            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["seqNo"] = "0";
                drw2["Docentry"] = drw["Docnum"];
                drw2["ItemCode"] = drw["ItemCode"];
                drw2["Dscription"] = drw["Dscription"];
                drw2["PiNo"] = "";
                drw2["ItemRemark"] = "宇豐調整";
                drw2["Quantity"] = drw["數量"];
                drw2["CHOPrice"] = drw["單價"];
                drw2["linenum"] = drw["ROWNO"];
                drw2["CHOAmount"] = drw["金額"];
                //  drw2["CHOMemo"] = drw["備註"];

                shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                dt2.Rows.Add(drw2);


                sAMEMOTextBox.Text = drw["備註"].ToString();
            }


            for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
            {
                shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
            }

            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();



        }
        private void button15DIAOBO_Click(object sender, EventArgs e)
        {
   
 
       
                tabControl1.SelectedIndex = 0;


                System.Data.DataTable dt1 = null;
                dt1 = GetCHODIAOBO(pinoTextBox.Text);
       

                System.Data.DataTable dt2 = ship.Shipping_Item;


                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["seqNo"] = "0";
                    drw2["Docentry"] = drw["Docnum"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["PiNo"] = "";
                    drw2["ItemRemark"] = "宇豐調撥";
                    drw2["Quantity"] = drw["數量"];
                    drw2["CHOPrice"] = drw["單價"];
                    drw2["linenum"] = drw["ROWNO"];
                    drw2["CHOAmount"] = drw["金額"];
                  //  drw2["CHOMemo"] = drw["備註"];

                    shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                    dt2.Rows.Add(drw2);


                    sAMEMOTextBox.Text = drw["備註"].ToString();
                }


                for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
                {
                    shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
                }

                shipping_mainBindingSource.EndEdit();
                shipping_ItemBindingSource.EndEdit();
            


        }
        private void buttonRETURN_Click(object sender, EventArgs e)
        {



            tabControl1.SelectedIndex = 0;


            System.Data.DataTable dt1 = null;
            dt1 = GetRETURN(pinoTextBox.Text);


            System.Data.DataTable dt2 = ship.Shipping_Item;


            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["seqNo"] = "0";
                drw2["Docentry"] = drw["Docnum"];
                drw2["ItemCode"] = drw["ItemCode"];
                drw2["Dscription"] = drw["Dscription"];
                drw2["PiNo"] = "";
                drw2["ItemRemark"] = "借入還出";
                drw2["Quantity"] = drw["數量"];
                drw2["CHOPrice"] = drw["單價"];
                drw2["linenum"] = drw["ROWNO"];
                drw2["CHOAmount"] = drw["金額"];
                // drw2["CHOMemo"] = drw["備註"];

                shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                dt2.Rows.Add(drw2);


                sAMEMOTextBox.Text = drw["備註"].ToString();
            }


            for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
            {
                shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
            }

            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();



        }
        private void buttonGIATO_Click(object sender, EventArgs e)
        {



            tabControl1.SelectedIndex = 0;


            System.Data.DataTable dt1 = null;
            dt1 = GetGIATO(pinoTextBox.Text);


            System.Data.DataTable dt2 = ship.Shipping_Item;


            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["seqNo"] = "0";
                drw2["Docentry"] = drw["Docnum"];
                drw2["ItemCode"] = drw["ItemCode"];
                drw2["Dscription"] = drw["Dscription"];
                drw2["PiNo"] = "";
                drw2["ItemRemark"] = "宇豐借出" ;
                drw2["Quantity"] = drw["數量"];
                drw2["CHOPrice"] = drw["單價"];
                drw2["linenum"] = drw["ROWNO"];
                drw2["CHOAmount"] = drw["金額"];
               // drw2["CHOMemo"] = drw["備註"];

                shipping_OBUTextBox.Text = drw["Docnum"].ToString();



                dt2.Rows.Add(drw2);


                sAMEMOTextBox.Text = drw["備註"].ToString();
            }


            for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
            {
                shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
            }

            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();



        }
        public System.Data.DataTable GetCHO(string DocEntry,string FLAG)
        {
            SqlConnection connection = new SqlConnection(strCn16);
         
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT T0.BillNO Docnum,T1.ProdID ItemCode,J.InvoProdName Dscription,T1.Quantity 數量,T1.Price 單價,T1.Amount 金額,T1.ROWNO,T0.REMARK 備註,T0.ProjectID PRJ FROM OrdBillMain T0");
            sb.Append("                      Inner Join OrdBillSub T1 On T0.Flag=T1.Flag And T0.BillNO=T1.BillNO   Inner Join comProduct J On T1.ProdID =J.ProdID    ");
            sb.Append("                       where T0.BillNO=@BillNO and T0.Flag =@Flag");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", DocEntry));
            command.Parameters.Add(new SqlParameter("@Flag", FLAG));
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
        public System.Data.DataTable GetCHODIAOBO(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);
      
  
            StringBuilder sb = new StringBuilder();
            sb.Append("                            SELECT  A.BillNO Docnum,A.ProdID ItemCode,A.ProdName Dscription,A.Quantity 數量,A.Price 單價,A.Amount 金額,A.ROWNO,S.REMARK 備註  From comProdRec A ");
            sb.Append(" Inner Join stkMoveSub G On G.Flag=A.Flag And G.MoveNO=A.BillNO And G.RowNo=A.RowNO ");
            sb.Append(" Inner Join stkMoveMAIN S ON (G.MoveNO =S.MoveNO AND G.Flag =S.Flag)");

            sb.Append("                       where A.BillNO=@BillNO  ");
            
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", DocEntry));

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

        public System.Data.DataTable GetSTKBILLNO(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);


            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT  A.BillNO Docnum,A.ProdID ItemCode,A.ProdName Dscription,A.Quantity 數量,A.Price 單價,A.Amount 金額,A.ROWNO,S.Remark  備註  From comProdRec A   ");
            sb.Append("                                           INNER join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =100)");
            sb.Append("                       where A.BillNO=@BillNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", DocEntry));

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

        public System.Data.DataTable GetSTKBILLNO2(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);


            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT  A.BillNO Docnum,A.ProdID ItemCode,A.ProdName Dscription,A.Quantity 數量,A.Price 單價,A.Amount 金額,A.ROWNO,S.Remark  備註  From comProdRec A   ");
            sb.Append("                                           INNER join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)");
            sb.Append("                       where A.BillNO=@BillNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", DocEntry));

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
        public System.Data.DataTable GetCHOTIOJEN(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);


            StringBuilder sb = new StringBuilder();
            sb.Append("                                          SELECT  A.BillNO Docnum,A.ProdID ItemCode,A.ProdName Dscription,A.Quantity 數量,A.Price 單價,A.Amount 金額,A.ROWNO,S.REMARK 備註  From comProdRec A  ");
            sb.Append("               Inner Join StkAdjustSub G On G.Flag=A.Flag And G.AdjustNO=A.BillNO And G.RowNo=A.RowNO  ");
            sb.Append("               Inner Join StkAdjustMain S ON (G.AdjustNO =S.AdjustNO AND G.Flag =S.Flag) ");

            sb.Append("                       where A.BillNO=@BillNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", DocEntry));

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
        public System.Data.DataTable GetGIATO(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);


            StringBuilder sb = new StringBuilder();
            sb.Append("                                                           SELECT  G.BorrowNO Docnum,G.ProdID ItemCode,G.ProdName Dscription,G.Quantity 數量,G.Price 單價,G.Amount 金額,G.ROWNO,S.REMARK 備註  ");
            sb.Append("                                                                                      From  stkBorrowSub G ");
            sb.Append("               Inner Join StkBorrowMain S ON (G.BorrowNO =S.BorrowNO AND G.Flag =S.Flag) ");
            sb.Append("                                     where G.BorrowNO=@BorrowNO   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BorrowNO", DocEntry));

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
        public System.Data.DataTable GetRETURN(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn16);


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   G.Flag,G.ReturnNO  Docnum,G.ProdID ItemCode,G.ProdName Dscription,G.Quantity 數量,0 單價,0 金額,G.ROWNO,S.REMARK 備註   ");
            sb.Append(" From  stkReturnSub G  ");
            sb.Append(" Inner Join stkReturnMain S ON (G.ReturnNO  =S.ReturnNO  AND G.Flag =S.Flag)  ");
            sb.Append(" where G.ReturnNO=@ReturnNO   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ReturnNO", DocEntry));

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
        public System.Data.DataTable GetADINVOPROD(string ProdID)
        {
            SqlConnection connection = new SqlConnection(strCn16);

            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT InvoProdName  FROM comProduct where ProdID=@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

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


        public System.Data.DataTable GetOrdrship1(string Doc_no)
        {
            
            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.totalfrgn,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號 ");
            sb.Append("               ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER,U_SHIPPRICE");
            sb.Append(" ,T0.DOCCUR,T0.DOCRATE,Convert(varchar(8),T0.DOCDATE,112) DOCDATE    from rdr1 t1 ");
            sb.Append("               left join ordr t0 on (t1.docentry=t0.docentry)   where  1=1   ");
            if (SOL != 1)
            {
                sb.Append(" AND T1.TREETYPE <> 'I'  ");
                sb.Append(" AND  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','')+isnull(u_acme_workday,'')  in (N" + Doc_no + ") order by t0.Docnum,visorder ");
            }
            else
            {
                sb.Append(" AND  ISNULL(cast(T0.Docnum as nvarchar),'')  in (N" + Doc_no + ") order by t0.Docnum,visorder ");
            }
          
        
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
        
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

        public System.Data.DataTable GetOrdrshipORINT(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.LINETOTAL,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER   from DRF1 t1");
            sb.Append(" left join ODRF t0 on (t1.docentry=t0.docentry) where T1.OBJTYPE=14 AND  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') in (" + Doc_no + ") order by t0.Docnum,visorder ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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

        public System.Data.DataTable GetOrdrshipORIN(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.LINETOTAL,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER   from RIN1 t1");
            sb.Append(" left join ORIN t0 on (t1.docentry=t0.docentry) where  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') in (" + Doc_no + ") order by t0.Docnum,visorder ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GetOrdrshipORPC(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.LINETOTAL,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER   from RPC1 t1");
            sb.Append(" left join ORPC t0 on (t1.docentry=t0.docentry) where  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') in (" + Doc_no + ") order by t0.Docnum,visorder ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GetOrdrshipORPD(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.LINETOTAL,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER   from RPD1 t1");
            sb.Append(" left join ORPD t0 on (t1.docentry=t0.docentry) where  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') in (" + Doc_no + ") order by t0.Docnum,visorder ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GetINSHIP(string shippingcode)
        {

            SqlConnection MyConnection;

            MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT *  FROM lcInstro1 where shippingcode=@shippingcode and isnull(checked,'')=''");
         


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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
        public System.Data.DataTable Getshipitem(string shippingcode,int  TYPE,string DOC)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();


            sb.Append("     select t1.itemcode,Dscription,Quantity,bb=     ");
            sb.Append("  Dscription, ");

            sb.Append("  ItemPrice ");

                sb.Append("                            ,t1.Docentry,linenum,CHOAmount,CHOPrice,OLDORDER,VISORDER,T1.CURRENCY,T1.RATE,T1.RATEUSD,T1.ItemAmount  from shipping_item T1 ");
            sb.Append("               LEFT JOIN  ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            if (TYPE == 1)
            {
                sb.Append(" WHERE T1.SHIPPINGCODE=@shippingcode");
            }
            if (TYPE == 2)
            {
                sb.Append(" WHERE T1.DOCENTRY1 IN (" + DOC + ") ");
            }
            if (checkBox4.Checked)
            {
                sb.Append(" AND  OldOrder <> 'I' ");
            }
      

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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


        public System.Data.DataTable GetshipTYPE()
        {
            SqlConnection MyConnection = globals.Connection;
           
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SHIPPINGCODE FROM SHIPPING_ITEM WHERE ITEMREMARK='採購訂單' AND SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
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
        public System.Data.DataTable GetINVSEQ(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select COUNT(*)+1 COUN,MAX(INVOICENO) INVOICENO FROM INVOICEM WHERE SHIPPINGCODE=@SHIPPINGCODE");

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

        public System.Data.DataTable GetCHO2(string ID, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T2.MEMO billbuilding,T2.[Address] billstreet,T2.Telephone billblock,T2.FaxNo billcity,T2.LinkMan billzipcode  FROM   comCustDesc T0");
            sb.Append(" LEFT JOIN comCustAddress T2 ON (T0.EngAddrID=T2.AddrID and T0.ID=T2.ID  )");
            sb.Append(" WHERE T0.ID=@ID AND T0.Flag=1 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
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
        public System.Data.DataTable GetCHO22(string BillNO, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select CustomerID  from ordBillMain where BillNO=@BillNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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
        public System.Data.DataTable GetCHO3(string BILLNO, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select   T1.MEMO shipbuilding,T1.[Address] shipstreet,T1.Telephone shipblock,T1.FaxNo shipcity,T1.LinkMan shipzipcode from OrdBillMain T0");
            sb.Append("   LEFT JOIN comCustAddress T1 ON (T0.AddressID=T1.AddrID AND T0.CustomerID=T1.ID )");
            sb.Append("  where BILLNO=@BILLNO and t0.Flag=2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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
        public System.Data.DataTable GetREMARKSA()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT REMARK FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
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

        public System.Data.DataTable GetDOCCUR(string DOCENTRY,string ORDR)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCCUR,DOCRATE,Convert(varchar(8),DOCDATE,112) DOCDATE  from " + ORDR + " WHERE DOCENTRY=@DOCENTRY  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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
        public System.Data.DataTable GetDOCCUR2(string DOCDATE)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT RATE  FROM ORTT WHERE CURRENCY='USD' AND Convert(varchar(8),RATEDATE,112)=@DOCDATE  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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
        public System.Data.DataTable GetCHOICE(string BILLNO, string ProdID, int Quantity, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select Price,Amount,CUSTOMERID,A.REMARK,G.ProdID,CAST(G.Quantity AS INT)  from  OrdBillMain A ");
            sb.Append(" Inner Join OrdBillSub G  On (G.Flag=A.Flag  And G.BillNO=A.BillNO)");
            sb.Append(" where a.flag=2 AND A.BILLNO=@BILLNO AND G.ProdID=@ProdID AND CAST(G.Quantity AS INT)=@Quantity ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@strCn1", strCn1));
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
        public System.Data.DataTable GetCHOICEIT(string BILLNO, string ProdID, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select Price,Amount,CUSTOMERID,A.REMARK,G.ProdID,CAST(G.Quantity AS INT)  from  OrdBillMain A ");
            sb.Append(" Inner Join OrdBillSub G  On (G.Flag=A.Flag  And G.BillNO=A.BillNO)");
            sb.Append(" where a.flag=2 AND A.BILLNO=@BILLNO AND G.ProdID=@ProdID  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@strCn1", strCn1));
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
        public System.Data.DataTable GetLOGIN()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LOGINID FROM SHIPPING_LOGIN WHERE SHIPPINGCODE=@SHIPPINGCODE   ");
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

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

        public System.Data.DataTable GetLOGIN2(string LOGINID)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LOGINID FROM SHIPPING_LOGIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND LOGINID=@LOGINID   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@LOGINID", LOGINID));
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
        public System.Data.DataTable GetCHOICECHECK(string BILLNO, string ProdID, int Quantity, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();


            sb.Append(" select CAST(Quantity AS INT) QTY from OrdBillSub T0");
            sb.Append(" WHERE BILLNO+ProdID IN (");
            sb.Append("                          select DISTINCT  BillNO+ProdID from (");
            sb.Append("                                                     select g.BILLNO,G.ProdID,Price   from  OrdBillSub g");
            sb.Append("                                           where g.flag=2 and BillNO+ProdID in (    ");
            sb.Append("                                                      select g.BILLNO+G.ProdID  from  OrdBillSub g");
            sb.Append("                                           where g.flag=2  and Price <> 0 ");
            sb.Append("                                           GROUP BY g.BILLNO,G.ProdID");
            sb.Append("                                           HAVING COUNT(*) > 1 )");
            sb.Append("                                           GROUP BY g.BILLNO,G.ProdID,Price");
            sb.Append("                                           HAVING COUNT(*) = 1 ) as a  where  BILLNO=@BILLNO AND ProdID=@ProdID) AND CAST(Quantity AS INT)=@Quantity");
      
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@strCn1", strCn1));
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

        private void button34_Click(object sender, EventArgs e)
        {
              try
            {

                if (download2DataGridView.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇檔案");

                    return;
                }

                DialogResult result;
                result = MessageBox.Show("請確定是否要將檔案移到可下載區", "YesNo", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    if (download2DataGridView.SelectedRows.Count > 0)
                    {

                        DataGridViewRow row;
                        for (int i = download2DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                        {

                            row = download2DataGridView.SelectedRows[i];

                            string T1 = row.Cells["seq"].Value.ToString();

                            System.Data.DataTable N1 = GetMenu.GetDOWNLOAD2(shippingCodeTextBox.Text, T1);

                            string T2 = N1.Rows[0][0].ToString();
                            string T3 = N1.Rows[0][1].ToString();


                            int J = downloadDataGridView.Rows.Count;

                            System.Data.DataTable dth = ship.Download;
                            //row1
                            DataRow drw2 = dth.NewRow();

                            drw2["Seq"] = J.ToString();
                            drw2["ShippingCode"] = shippingCodeTextBox.Text;
                            drw2["filename"] = T2;
                            drw2["path"] = T3;
                            if (dOCTYPETextBox.Text == "銷售")
                            {
                                System.Data.DataTable G1 = GetSA(pinoTextBox.Text.Trim());
                                if (G1.Rows.Count > 0)
                                {
                                    drw2["SA"] = G1.Rows[0]["業管"].ToString();
                                    drw2["SALES"] = G1.Rows[0]["業務"].ToString();
                                }

                            }
                            dth.Rows.Add(drw2);
                        }

                    }

                    int iSelectRowCount = download2DataGridView.SelectedRows.Count;


                    //判斷是否是選擇了行
                    if (iSelectRowCount > 0)
                    {
                        //循環刪除行
                        foreach (DataGridViewRow dgvRow in download2DataGridView.SelectedRows)
                        {
                            download2DataGridView.Rows.Remove(dgvRow);
                        }

                    }

       
                }

                download2BindingSource.MoveFirst();

                for (int i = 0; i <= download2BindingSource.Count - 1; i++)
                {
                    DataRowView row1 = (DataRowView)download2BindingSource.Current;

                    row1["seq"] = i;



                    download2BindingSource.EndEdit();

                    download2BindingSource.MoveNext();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            add3TextBox.Text = comboBox3.Text;
         
      
        }

        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = GetMenu.Getfee("add3");


            comboBox3.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }



    
        public void AddDRS(string JOB,string INV,string QTY,string MODEL)
        {
            SqlConnection connection = new SqlConnection(DRS);
            SqlCommand command = new SqlCommand(" Insert into Shipping_DRS(JOB,INV,QTY,MODEL) values(@JOB,@INV,@QTY,@MODEL)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JOB", JOB));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));

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


        private void T1()
        {
            System.Data.DataTable dt = GetMenu.GETDRSINV(shippingCodeTextBox.Text);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string 工單號碼 = dt.Rows[i]["工單號碼"].ToString();
                string INV = dt.Rows[i]["INV"].ToString();
                string 數量 = dt.Rows[i]["數量"].ToString();
                System.Data.DataTable dt2 = GetMenu.GETDRSINV2(shippingCodeTextBox.Text);
                for (int j = 0; j <= dt2.Rows.Count - 1; j++)
                {
                    sb.Append(dt2.Rows[i]["MODEL"].ToString() + "/");

                }
                sb.Remove(sb.Length - 1, 1);
                AddDRS(工單號碼, INV, 數量, sb.ToString());
            }
        }

        public System.Data.DataTable GetDRS(string JOB)
        {

            SqlConnection MyConnection = new SqlConnection(DRS);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM Shipping_DRS WHERE JOB=@JOB");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@JOB", JOB));
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



        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                add7TextBox.Text = comboBox5.Text;
                System.Data.DataTable O1 = GetSHIPEXSIT();
                System.Data.DataTable O2 = GetSHIPOHEM(comboBox5.Text);
                if (O1.Rows.Count == 0)
                {
                    if (O2.Rows.Count > 0)
                    {
                        shippingCodeTextBox.Text = shippingCodeTextBox.Text.Replace("X", "D");

                        DataGridViewRow row;
                        for (int i = shipping_ItemDataGridView.Rows.Count - 1; i >= 0; i--)
                        {
                            row = shipping_ItemDataGridView.Rows[i];

                            row.Cells[0].Value = shippingCodeTextBox.Text;
                        }
                        this.shipping_ItemBindingSource.EndEdit();
                    }
                    else
                    {
                        shippingCodeTextBox.Text = shippingCodeTextBox.Text.Replace("D", "X");

                        DataGridViewRow row;
                        for (int i = shipping_ItemDataGridView.Rows.Count - 1; i >= 0; i--)
                        {
                            row = shipping_ItemDataGridView.Rows[i];

                            row.Cells[0].Value = shippingCodeTextBox.Text;
                        }
                        this.shipping_ItemBindingSource.EndEdit();
                    }

                    this.shipping_mainBindingSource.EndEdit();

                }
            }
            catch
            { }


        }

        private void comboBox5_MouseClick(object sender, MouseEventArgs e)
        {
            //WH
            System.Data.DataTable dt3 = null;
            if (globals.GroupID.ToString().Trim() == "SHI" || globals.GroupID.ToString().Trim() == "EEP")
            {
                dt3 = GetOHEMSHIP1();
            }
            if (globals.GroupID.ToString().Trim() == "ShipBuy" || globals.GroupID.ToString().Trim() == "WH")
            {
                dt3 = GetOHEMSHIP2();
            }

                comboBox5.Items.Clear();


                for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                {
                    comboBox5.Items.Add(Convert.ToString(dt3.Rows[i][0]));
                }
            

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

            quantityTextBox.Text = comboBox6.Text;

            if (quantityTextBox.Text == "已結")
            {
                buCardnameTextBox.Text = GetMenu.Day();
                buCardcodeCheckBox.Checked = true;
            }
            else if (quantityTextBox.Text == "取消")
            {
                buCardnameTextBox.Text = GetMenu.Day();
                buCardcodeCheckBox.Checked = false;
            }
            else
            {
                buCardnameTextBox.Text = "";
                buCardcodeCheckBox.Checked = false;
            }
        }

        private void comboBox6_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SHIPSTATUS");

            comboBox6.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox6.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            receiveDayTextBox.Text = comboBox7.Text;

        }

        private void comboBox7_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("receiveDay");

            comboBox7.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox7.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox8_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("boardCountNo");

            comboBox8.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox8.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            boardCountNoTextBox.Text = comboBox8.Text;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            f2 = 0;
            f3 = 0;
            if (add10TextBox.Text != "Checked")
            {
                add10CheckBox.Checked = false;
            }

            StringBuilder ss = new StringBuilder();
            if (listBox1.SelectedItems.Count != 0)
            {


                ArrayList al = new ArrayList();
                for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                {
                    string f = listBox1.SelectedItems[i].ToString();
                    al.Add(listBox1.SelectedItems[i].ToString());
                }


            
                foreach (string v in al)
                {
                    ss.Append("" + v + "@acmepoint.com;");
                }
            }

            if (checkBox3.Checked)
            {
                System.Data.DataTable  SHIPSTOCCK = GetWHSHIP();
                if (SHIPSTOCCK.Rows.Count > 0)
                {
                    for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                    {
                        DataRow dd = SHIPSTOCCK.Rows[i];
                        ss.Append(dd["EMAIL"].ToString() + ";");
                    }
                }
            }

            if (checkBox5.Checked)
            {
                System.Data.DataTable SHIPSTOCCK = GetWHSTOCK();
                if (SHIPSTOCCK.Rows.Count > 0)
                {
                    for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                    {
                        DataRow dd = SHIPSTOCCK.Rows[i];
                        ss.Append(dd["EMAIL"].ToString() + ";");
                    }
                }
            }

            if (ss.Length > 5)
            {
                ss.Remove(ss.Length - 1, 1);
                mail = ss.ToString();
                if (globals.GroupID.ToString().Trim() == "EEP")
                {
                    mail = "lleytonchen@acmepoint.com";
                }
                SENDMAIL(GetTODO_USERDataSource2SA(), "B", f2, f3);
            }
            else
            {
                MessageBox.Show("請選擇收件者");
            }
      
        }


        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (dataGridView2.Columns[e.ColumnIndex].Name == "FeeCheck")
                {

                    string T1 = this.dataGridView2.Rows[e.RowIndex].Cells["FeeCheck"].Value.ToString();
                    string ID = this.dataGridView2.Rows[e.RowIndex].Cells["ID"].Value.ToString();
                    if (T1 == "True")
                    {
                        UPDATESAP("True", ID);
                    }
                    else
                    {
                        UPDATESAP("False", ID);
                    }

          
                }
            }
            catch { }
        }
        private void UPINVOICE()
        {
            try
            {
                this.invoiceMBindingSource.EndEdit();
                this.invoiceMTableAdapter.Update(ship.InvoiceM);
                ship.InvoiceM.AcceptChanges();
                this.invoiceDBindingSource.EndEdit();
                this.invoiceDTableAdapter.Update(ship.InvoiceD);
                ship.InvoiceD.AcceptChanges();

                invoiceDBindingSource.MoveFirst();
                int s = 0;
                for (int i = 0; i <= invoiceDBindingSource.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)invoiceDBindingSource.Current;
                    row["seqno"] = i;
                    if (GetINVMARK().Rows.Count == 0)
                    {

                        row["seqno2"] = i;
                    }
                    else
                    {

                        string MarkNos = row["MarkNos"].ToString();
                        if (MarkNos == "True")
                        {
                            row["seqno2"] = s++;
                        }
                        else
                        {
                            row["seqno2"] = "";
                        }
                    }

                    invoiceDBindingSource.EndEdit();

                    invoiceDBindingSource.MoveNext();
                }

              
            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);

            }
        }
        private void UPPACK()
        {
            try
            {
                this.packingListMBindingSource.EndEdit();
                this.packingListMTableAdapter.Update(ship.PackingListM);
                ship.PackingListM.AcceptChanges();
                this.packingListDBindingSource.EndEdit();
                this.packingListDTableAdapter.Update(ship.PackingListD);
                ship.PackingListD.AcceptChanges();

                packingListDBindingSource.MoveFirst();
                int s = 0;
                for (int i = 0; i <= packingListDBindingSource.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)packingListDBindingSource.Current;
                    row["seqno"] = i;
                    if (GetINVPACK().Rows.Count == 0)
                    {

                        row["seqno2"] = i;
                    }
                    else
                    {

                        string PACKMARK = row["PACKMARK"].ToString();
                        if (PACKMARK == "True")
                        {
                            row["seqno2"] = s++;
                        }
                        else
                        {
                            row["seqno2"] = "";
                        }
                    }

                    packingListDBindingSource.EndEdit();

                    packingListDBindingSource.MoveNext();
                }


            }

            catch (Exception ex)
            {

                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Packing儲存1", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

            }

        }
        public void UPDATEINMAIL(string ShippingCode, string DocNum, string seqno)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE lcInstro1 SET Checked='True'  where ShippingCode=@ShippingCode and DocNum=@DocNum and seqno=@seqno ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));


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

        public void UPDATEINMAIL2(string ShippingCode, string DocNum, string seqno)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE lcInstro1 SET RED=''  where ShippingCode=@ShippingCode and DocNum=@DocNum and seqno=@seqno ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));


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
        public void UPDATEINMAIL3(string ShippingCode)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE SHIPPING_ITEM SET RED=''  where ShippingCode=@ShippingCode ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));



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

        public void UPDATESAP(string CHECK,string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE Shipping_Fee SET FeeCheck=@CHECK  where ID=@ID ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@CHECK", CHECK));



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
        public void UPDATEDOWNSA(string SA, string SALES,string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE DOWNLOAD SET SA=@SA,SALES=@SALES WHERE SHIPPINGCODE=@SHIPPINGCODE ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SA", SA));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
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


        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
            string f = "c";
            string[] filebType = Directory.GetDirectories("//acmesrv01//SAP_Share//shipping//");
            string dd = DateTime.Now.ToString("yyyyMM");
            string tt = "//acmesrv01//SAP_Share//shipping//" + dd;
            foreach (string fileaSize in filebType)
            {

                if (fileaSize == tt)
                {
                    f = "d";

                }

            }
            if (f == "c")
            {
                Directory.CreateDirectory(tt);
            }
          
                string server = "//acmesrv01//SAP_Share//shipping//" + dd + "//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download3(filename);

                if (dt2.Rows.Count > 0)
                {
                    string G1 = dt2.Rows[0]["filename"].ToString().Replace(" ", "").ToUpper().Trim();
                    string BAU = add9TextBox.Text.Replace(" ", "").ToUpper().Trim();
                    int F1 = G1.IndexOf(BAU);
                    if (F1 == -1)
                    {

                             MessageBox.Show("檔案名稱重複,請修改檔名");
                         }
                }
                else
                {
                    if (result == DialogResult.OK)
                    {

                        string file = opdf.FileName;
                        bool F1 = getrma.UploadFile(file, server, false);
                        if (F1 == false)
                        {
                            return;
                        }
                        System.Data.DataTable dt1 = ship.Download3;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (download3DataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = @"\\acmesrv01\SAP_Share\shipping\" + de + filename;
                        dt1.Rows.Add(drw);

                        download3BindingSource.MoveFirst();

                        for (int i = 0; i <= download3BindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)download3BindingSource.Current;

                            rowd["seq"] = i;

                            download3BindingSource.EndEdit();
                            download3BindingSource.MoveNext();
                        }

                        this.download3BindingSource.EndEdit();
                        this.download3TableAdapter.Update(ship.Download3);
                        this.ship.Download3.AcceptChanges();


                        MessageBox.Show("上傳成功");
                    }


                }
            }
            catch (Exception ex)
            {
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "可下載檔案上傳", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                MessageBox.Show(ex.Message);
            }
        }


        private void download3DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK3")
                {

                    System.Data.DataTable dt1 = ship.Download3;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];
                    string aa = drw["path"].ToString();
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    string filename = drw["filename"].ToString();
                    string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                    System.IO.File.Copy(aa, NewFileName, true);
                    System.Diagnostics.Process.Start(NewFileName);

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


        private void button126_Click(object sender, EventArgs e)
        {
            if (dOCTYPETextBox.Text == "")
            {
                MessageBox.Show("請選擇單據");
                return;
            }

            if (dOCTYPETextBox.Text == "銷售")
            {
                SOL = 0;
                button15_Click(sender, e, "2");
           
                
            }
     
            if (dOCTYPETextBox.Text == "採購")
            {
                button15_Click(sender, e, "4");
            }


            if (dOCTYPETextBox.Text == "調撥")
            {
                button15DIAOBO_Click(sender, e);
            }

            if (dOCTYPETextBox.Text == "進貨憑單")
            {
                button15F(sender, e);

            }


            if (dOCTYPETextBox.Text == "進貨憑單")
            {
                button15F(sender, e);

            }

            if (dOCTYPETextBox.Text == "調整")
            {
                button15TIOJEN_Click(sender, e);
            }

            if (dOCTYPETextBox.Text == "借出")
            {
                buttonGIATO_Click(sender, e);

            }

            if (dOCTYPETextBox.Text == "借入還出")
            {
                buttonRETURN_Click(sender, e);

            }
          
      
            
            
        }

        private void 全選ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in invoiceDDataGridView.Rows)
                dr.Cells[0].Value = "True";


        }

        private void 顯示編號全取消ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in invoiceDDataGridView.Rows)
                dr.Cells[0].Value = "False";
        }

        private void 顯示編號全選ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in packingListDDataGridView.Rows)
                dr.Cells[0].Value = "True";
        }

        private void 顯示編號全取消ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in packingListDDataGridView.Rows)
                dr.Cells[0].Value = "False";
        }

 

        private void button19_Click(object sender, EventArgs e)
        {

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\SHIPMARK.xls";

            System.Data.DataTable K1 = GetSHIPMARK();
            if (K1.Rows.Count > 0)
            {
                GetExcelinsu(FileName, K1);
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        

        private void lcInstro1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LC" ||
                lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LDOCENTRY" ||
                                lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LITEMCODE" ||
                                lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LDESC" ||
                 lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LQTY" ||
                                 lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LPRICE" ||
                                 lcInstro1DataGridView.Columns[e.ColumnIndex].Name == "LAMT"
                )
            {
                this.lcInstro1DataGridView.Rows[e.RowIndex].Cells["RED"].Value = "True";
            }


        }

        private void lcInstro1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = lcInstro1DataGridView.Rows.Count - 1;
            e.Row.Cells["LSEQNO"].Value = iRecs.ToString();
        }

        private void lcInstro1DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= lcInstro1DataGridView.Rows.Count - 1)
                return;
            DataGridViewRow dgr = lcInstro1DataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["RED"].Value.ToString().Trim() == "True")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }
           

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }



        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            iNSSHIPWAYTextBox.Text = comboBox10.Text;
        }

        private void comboBox10_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("receiveDay");

            comboBox10.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox10.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }





        private void fmShip_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Data.DataTable L1 = GetLOGIN2(fmLogin.LoginID.ToString());
            if (L1.Rows.Count > 0)
            {
                string H1 = L1.Rows[0][0].ToString();
                MessageBox.Show("此工單" + H1 + "修改中，無法按Ｘ ");
                e.Cancel = true;
            }

 
        }

        private void button48_Click(object sender, EventArgs e)
        {
            if (invoiceNoTextBox.Text == "")
            {
                MessageBox.Show("請先新增Invoice");
                return;
            }

            if (shipping_ItemDataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow row;
                StringBuilder sb = new StringBuilder();

                for (int i = shipping_ItemDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = shipping_ItemDataGridView.SelectedRows[i];

                    sb.Append("'" + row.Cells["Docentry1"].Value.ToString() + "',");
                }
                sb.Remove(sb.Length - 1, 1);


                System.Data.DataTable dt3 = Getshipitem(shippingCodeTextBox.Text, 2, sb.ToString());

                System.Data.DataTable dt4 = ship.InvoiceD;

                int G1 = invoiceDDataGridView.Rows.Count - 1;
                if (shipping_ItemDataGridView.Rows.Count > 1)
                {

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt3.Rows[i];
                        DataRow drw2 = dt4.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["InvoiceNo"] = invoiceNoTextBox.Text;
                        drw2["InvoiceNo_seq"] = invoiceNo_seqTextBox.Text;
                        drw2["SeqNo"] = (G1 + i).ToString();
                        string ITEMCODE = drw["itemcode"].ToString();
                        System.Data.DataTable TT1 = GetADINVOPROD(ITEMCODE);
                        if (TT1.Rows.Count > 0)
                        {
                            drw2["INDescription"] = TT1.Rows[0][0].ToString();
                        }
                        else
                        {
                            drw2["INDescription"] = drw["bb"];
                        }
                        drw2["InQty"] = drw["Quantity"];
                        drw2["UnitPrice"] = drw["ItemPrice"];
          
                        string TYPE = drw["OLDORDER"].ToString();
         

         


                        drw2["LINENUM"] = drw["linenum"];


                        drw2["CHOPrice"] = drw["CHOPrice"];
                        drw2["CHOAmount"] = drw["CHOAmount"];
                        drw2["TREETYPE"] = TYPE;
                        drw2["VISORDER"] = drw["VISORDER"];
                        dt4.Rows.Add(drw2);

                    }

                }

                try
                {

                    this.invoiceMBindingSource.EndEdit();
                    this.invoiceMTableAdapter.Update(ship.InvoiceM);
                    ship.InvoiceM.AcceptChanges();

                    this.invoiceDBindingSource.EndEdit();
                    this.invoiceDTableAdapter.Update(ship.InvoiceD);
                    ship.InvoiceD.AcceptChanges();

                }
                catch (Exception ex)
                {

                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "InvoiceTran1", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            else
            {
                MessageBox.Show("請選擇項目");
            }
        }


        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINKC")
                {


                        string CARNAME = dataGridView3.CurrentRow.Cells["CARNAME"].Value.ToString();
                        string aa = dataGridView3.CurrentRow.Cells["CARPATH"].Value.ToString();

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + CARNAME;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);



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

        private void comboBox11_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = null;
            dt3 = GetMenu.GetBUGB("SHIPAD");

            comboBox11.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox11.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            dOCTYPETextBox.Text = comboBox11.Text;
        }

     


        private void UpdateTTUSD(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update download set STATUS='Y' where ID=@ID");

//            select T0.SHIPPINGCODE,T0.SEQ,[FILENAME],CREATENAME 製單人,MODIFYNAME 修改者  from download2 T0
//LEFT JOIN SHIPPING_MAIN T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE)
// WHERE SUBSTRING(T0.SHIPPINGCODE,3,4)='2016' AND ISNULL(STATUS,'') = 'Y'
//ORDER BY T0.SHIPPINGCODE,SEQ

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
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

        private System.Data.DataTable GETOHEM(string HOMETEL)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = new SqlConnection(strCn02);
            sb.Append(" SELECT CASE HOMETEL WHEN 'EvaHsu' THEN  'EvaHsuS' ELSE HOMETEL END HOMETEL  FROM OHEM WHERE HOMETEL=@HOMETEL ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", HOMETEL));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }


        public  System.Data.DataTable GetOHEMSHIP1()
        {
            SqlConnection con = new SqlConnection(strCn02);

            string sql = "SELECT HOMETEL FROM OHEM WHERE DEPT IN (7) AND ISNULL(TERMDATE,'') ='' ORDER BY HOMETEL";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }
        public  System.Data.DataTable GetOHEMSHIP2()
        {
            SqlConnection con = new SqlConnection(strCn02);

            string sql = "SELECT HOMETEL FROM OHEM WHERE DEPT IN (5,6) AND ISNULL(TERMDATE,'') ='' ORDER BY HOMETEL";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }

        private void dataGridView5_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView5.SelectedRows.Count > 0)
            {

                string da = dataGridView5.SelectedRows[0].Cells["併單工單"].Value.ToString();

                SHICAR a = new SHICAR();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

 

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {

                DialogResult result;
                result = MessageBox.Show("是否要寄出", "Close", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {

                        this.downloadBindingSource.EndEdit();
                        this.downloadTableAdapter.Update(ship.Download);
                        ship.Download.AcceptChanges();

                    }
                    catch (Exception ex)
                    {

                        GetMenu.InsertLog(fmLogin.LoginID.ToString(), "Download", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                        MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }

                    System.Data.DataTable M1 = GetDOWNLOADSA();
                    if (M1.Rows.Count == 0)
                    {
                        MessageBox.Show("請選擇要上傳檔案");
                        return;
                    }
                    System.Data.DataTable M2 = GetDOWNLOADSA2();
                    if (M2.Rows.Count == 0)
                    {
                        MessageBox.Show("沒有SA資料");
                        return;
                    }

                    string SA = M2.Rows[0]["SA"].ToString();
                    System.Data.DataTable M3 = GetDOWNLOADSA3(SA);
                    if (M3.Rows.Count > 0)
                    {

                        string template;
                        StreamReader objReader;
                        string FileName = string.Empty;
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


                        FileName = lsAppDir + "\\MailTemplates\\SHIPD.htm";

                        objReader = new StreamReader(FileName);

                        template = objReader.ReadToEnd();
                        objReader.Close();
                        objReader.Dispose();



                        StringWriter writer = new StringWriter();
                        HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                        template = template.Replace("##ETC##", "ETD : " + closeDayTextBox.Text);
                        template = template.Replace("##ETA##", "ETA : " + arriveDayTextBox.Text);
                        if (receiveDayTextBox.Text.Trim().ToUpper() != "TRUCK")
                        {
                            template = template.Replace("##HANZ##", "港名/航次 : " + boatNameTextBox.Text);
                            template = template.Replace("##SHIPNO##", "Shipping Order No : " + soNoTextBox.Text);

                        }
                        else
                        {
                            template = template.Replace("##HANZ##", "");
                            template = template.Replace("##SHIPNO##", "");
                        }


                        MailMessage message = new MailMessage();
                        string SALES = M2.Rows[0]["SALES"].ToString();

                        string MSA = M3.Rows[0][0].ToString();
                        MSA = "JOYCHEN@ACMEPOINT.COM";
                        message.To.Add(MSA);


                        if (!String.IsNullOrEmpty(SALES))
                        {
                            System.Data.DataTable M4 = GetDOWNLOADSA4(SALES);
                            if (M4.Rows.Count > 0)
                            {
                                string MSALES = M4.Rows[0][0].ToString();
                                message.To.Add(MSA);
                            }

                        }

                        message.CC.Add(fmLogin.LoginID.ToString());
                        string CARDNAME = cardNameTextBox.Text;
                        if (CARDNAME != "")
                        {
                            Regex rex = new Regex(@"^[A-Za-z0-9]+$");
                            string ENG = CARDNAME.Substring(0, 1);
                            Match ma = rex.Match(ENG);
                            if (ma.Success)
                            {
                                int t1 = CARDNAME.IndexOf(" ");
                                if (t1 != -1)
                                {
                                    CARDNAME = CARDNAME.Substring(0, t1);
                                }
                            }
                            else
                            {
                                CARDNAME = CARDNAME.Substring(0, 5);
                            }
                        }

                        message.Subject = "Shipping Doc_" + CARDNAME + "_" + receiveDayTextBox.Text + "_" + shippingCodeTextBox.Text;
                        message.Body = template;

                        for (int i = 0; i <= M1.Rows.Count - 1; i++)
                        {
                            string m_File = M1.Rows[i][0].ToString();
                            data = new Attachment(m_File, MediaTypeNames.Application.Octet);

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
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click_2(object sender, EventArgs e)
        {



            System.Data.DataTable M1 = GetDOWNLOADSAT1();
            for (int i = 0; i <= M1.Rows.Count - 1; i++)
            {
                string SHIPNO = M1.Rows[i][0].ToString();
                string DOCNO = M1.Rows[i][1].ToString();
                System.Data.DataTable G1 = GetSA(DOCNO);
                string SA = "";
                string SALES = "";
                if (G1.Rows.Count > 0)
                {
                    SA = G1.Rows[0]["業管"].ToString();
                    SALES = G1.Rows[0]["業務"].ToString();

                    UPDATEDOWNSA(SA, SALES, SHIPNO);
                }
            }
        }

        private void downloadDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (downloadDataGridView.Columns[e.ColumnIndex].Name == "DLCHECK")
                {

                    string S1 = Convert.ToString(this.downloadDataGridView.Rows[e.RowIndex].Cells["DLCHECK"].Value);
                    if (S1 == "True")
                    {
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE"].Value = GetMenu.Day();

                    }
                    else
                    {
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE"].Value = "";
                    }
                }
            }
            catch { }
        }
        public System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM SHIPPING_ITEM T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMREMARK ='宇豐銷售' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        public void SHIPNO()
        {
            if (dOCTYPETextBox.Text == "銷售")
            {
                mEMO3TextBox.Text = "";
                System.Data.DataTable dt3 = GetSHIP(shippingCodeTextBox.Text);
                if (dt3.Rows.Count > 0)
                {
                    string ITEMREMARK = dt3.Rows[0]["ITEMREMARK"].ToString();
                    if (ITEMREMARK == "宇豐銷售")
                    {

                        StringBuilder sb2 = new StringBuilder();
                        StringBuilder sb3 = new StringBuilder();
                        for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                        {
                            string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString();
                            string LINENUM = dt3.Rows[i]["LINENUM"].ToString();
                            sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                        }
                        sb2.Remove(sb2.Length - 1, 1);
                        string A = sb2.ToString();
                        System.Data.DataTable SS = GetSH(A, "銷售單");
                        if (SS.Rows.Count > 0)
                        {
                            for (int i = 0; i <= SS.Rows.Count - 1; i++)
                            {
                                string CODE = SS.Rows[i]["CODE"].ToString();

                                sb3.Append(CODE + ",");

                            }
                            sb3.Remove(sb3.Length - 1, 1);
                            mEMO3TextBox.Text = sb3.ToString();
                        }

                    }
                }
            }


            shipping_mainBindingSource.EndEdit();
            shipping_mainTableAdapter.Update(ship.Shipping_main);
            ship.Shipping_main.AcceptChanges();

        }
        private System.Data.DataTable GetSH(string DocEntry, string ITEMREMARK)
        {

            SqlConnection connection = null;
            connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from WH_item4 T0 left join wh_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK=@ITEMREMARK ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMREMARK", ITEMREMARK));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private void mEMO3TextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {


                string MEMOT = mEMO3TextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("WH201");
                string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                if (G1 != -1)
                {
                    string[] arrurl = H1.Split(new Char[] { ';' });

                    foreach (string i in arrurl)
                    {
                        MEMO = i.Substring(0, 17);
                        WH_main a = new WH_main();
                        a.PublicString = MEMO;
                        a.Show();
                    }

                }
            }
            catch { }
        }

        private void mEMO3TextBox_TextChanged(object sender, EventArgs e)
        {
            toolStripComboBox1.Items.Clear();
            int G1 = mEMO3TextBox.Text.IndexOf(",");
            if (G1 != -1)
            {
                string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                toolStripComboBox1.Items.Add("");
                foreach (string i in arrurl)
                {
                    toolStripComboBox1.Items.Add(i);

                }
            }
        }

        private void WrPACK4()
        {

            string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });

            int M1 = 0;
            foreach (string i in arrurl)
            {
                System.Data.DataTable WH4 = GetWHPACK6(i);
                if (WH4.Rows.Count > 0)
                {
                    for (int j = 0; j <= WH4.Rows.Count - 1; j++)
                    {
                        M1++;
                        AddWHPACK4(M1, i, WH4.Rows[j][0].ToString());
                    }
                }

            }
        }
        public void AddWHPACK4(int SEQ, string SHIPPINGCODE, string BLC)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" Insert into WH_PACK4(SEQ,SHIPPINGCODE,BLC,USERS) values(@SEQ,@SHIPPINGCODE,@BLC,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SEQ", SEQ));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BLC", BLC));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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
        public System.Data.DataTable GetWHPACK6(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT DISTINCT BLC,MAX(ID)  FROM WH_PACK2 WHERE SHIPPINGCODE =@SHIPPINGCODE AND ISNULL(BLC,'') <> ''  GROUP BY BLC ORDER BY MAX(ID)  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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

        public void DELPACK4()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete WH_PACK4 where users=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable Getshipitem08()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT * FROM ACMESQLSP.DBO.PackingListD WHERE SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", textBox13.Text.Trim()));
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
        public System.Data.DataTable GetshipitemM()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT MEMO FROM ACMESQLSP.DBO.PackingListM WHERE SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", textBox13.Text.Trim()));
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
        public System.Data.DataTable GetWHPACKH2()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT * FROM ACMESQLSP.DBO.MARK WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", textBox13.Text));
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
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {

                DeletePacking(shippingCodeTextBox.Text, pLNoTextBox.Text);
                packingListDTableAdapter.Fill(ship.PackingListD, MyID);
                markTableAdapter.Fill(ship.Mark, MyID);
                System.Data.DataTable dt3 = Getshipitem08();
                System.Data.DataTable dt5 = GetshipitemM();
                System.Data.DataTable dt4 = ship.PackingListD;
                if (dt5.Rows.Count > 0)
                {
                    memoTextBox.Text = dt5.Rows[0][0].ToString();
                }
                if (dt3.Rows.Count > 0)
                {

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt3.Rows[i];
                        DataRow drw2 = dt4.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["PLNo"] = pLNoTextBox.Text;
                        drw2["Doctentry"] = drw["Doctentry"];
                        drw2["PackageNo"] = drw["PackageNo"];
                        drw2["CNo"] = drw["CNo"];
                        drw2["DescGoods"] = drw["DescGoods"];
                        drw2["Quantity"] = drw["Quantity"];
                        drw2["Net"] = drw["Net"];
                        drw2["Gross"] = drw["Gross"];
                        drw2["MeasurmentCM"] = drw["MeasurmentCM"];
                        drw2["TREETYPE"] = drw["TREETYPE"];
                        drw2["VISORDER"] = drw["VISORDER"];
                        drw2["SOID"] = drw["SOID"];
                        drw2["PACKMARK"] = drw["PACKMARK"];

                        drw2["SeqNo2"] = drw["SeqNo2"];
                        drw2["CURRENCY"] = drw["CURRENCY"];
                        drw2["RATE"] = drw["RATE"];
                        drw2["AMTF"] = drw["AMTF"];
                        drw2["RATEUSD"] = drw["RATEUSD"];
                        drw2["WHNO"] = drw["WHNO"];
                        drw2["LOCATION"] = drw["LOCATION"];
                        drw2["PALQTY"] = drw["PALQTY"];
                        drw2["ITEMCODE"] = drw["ITEMCODE"];



                        dt4.Rows.Add(drw2);

                    }

                }

                System.Data.DataTable dt3H = GetWHPACKH2();

                if (dt3H.Rows.Count > 0)
                {

                    System.Data.DataTable dth = ship.Mark;
                    for (int j = 0; j <= dt3H.Rows.Count - 1; j++)
                    {

                        DataRow drw2 = dth.NewRow();

                        drw2["Seq"] = dt3H.Rows[j]["Seq"].ToString();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["Mark"] = dt3H.Rows[j]["Mark"].ToString();
                        dth.Rows.Add(drw2);

                    }

                }

                try
                {


                    this.packingListMBindingSource.EndEdit();
                    this.packingListMTableAdapter.Update(ship.PackingListM);
                    ship.PackingListM.AcceptChanges();

                    this.packingListDBindingSource.EndEdit();
                    this.packingListDTableAdapter.Update(ship.PackingListD);
                    ship.PackingListD.AcceptChanges();

                    this.markBindingSource.EndEdit();
                    this.markTableAdapter.Update(ship.Mark);
                    ship.Mark.AcceptChanges();

                }
                catch (Exception ex)
                {

                    GetMenu.InsertLog(fmLogin.LoginID.ToString(), "InvoiceTran1", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

}


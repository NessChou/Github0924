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
using System.Globalization;

namespace ACME
{
    public partial class fmShip : ACME.fmBase4
    {

        string OBJ = "15";
        int s1 = 0;
        int inint = 0;
        string DRS = "";
        int CON = 0;
        Attachment data = null;
        int f2 = 0;
        int f3 = 0;
        string mail = "";
        int CHO1 = 0;
        int CHO2 = 0;
        int CHO3 = 0;
        int COPY = 0;
        StringBuilder DF3F = new StringBuilder();
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
        string strINF = "Data Source=acmesap;Initial Catalog=AcmeSqlSPINFINITE;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCHO = "Data Source=acmesap;Initial Catalog=AcmeSqlSPCHOICE;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strTOP = "Data Source=acmesap;Initial Catalog=AcmeSqlSPTOPGARDEN;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string hh = "";
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCnSP = "Data Source=acmesap;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

        StringBuilder sbS = new StringBuilder();

        public fmShip()
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
            receivePlaceTextBox.ReadOnly = false;
            goalPlaceTextBox.ReadOnly = false;
            shipmentTextBox.ReadOnly = false;
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
                ship.SHIP_FEE.RejectChanges();

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
            sHIP_FEEBindingSource.EndEdit();

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
            sHIP_FEETableAdapter.Update(ship.SHIP_FEE);

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
            ship.SHIP_FEE.AcceptChanges();


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
            button14.Enabled = true;
            button50.Enabled = true;
            button51.Enabled = true;
            button17.Enabled = true;
            button7.Enabled = true;
            button43.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button12.Enabled = true;
            comboBox1.Enabled = true;
            button2.Enabled = true;
            button19.Enabled = true;
            contextMenuStrip2.Enabled = false;
            contextMenuStrip3.Enabled = false;
            contextMenuStrip4.Enabled = false;
            contextMenuStrip5.Enabled = false;
            add1TextBox.ReadOnly = true;
            add7TextBox.ReadOnly = true;
            button40.Enabled = true;
            button41.Enabled = true;
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

        public void NOTEXPORT()
        {
            if (globals.DBNAME == "進金生")
            {

                System.Data.DataTable H1 = RETAB();
                for (int i = 0; i <= H1.Rows.Count - 1; i++)
                {
                    string CELLNAME = H1.Rows[i]["CELLNAME"].ToString();
                    string GROUP = globals.GroupID.ToString().Trim();



                    System.Data.DataTable H2 = RETAB2(GROUP, CELLNAME);
                    if (H2.Rows.Count > 0)
                    {
                        tabControl1.TabPages.Remove(tabControl1.TabPages[CELLNAME]);
                    }

                }
            }


        }

        public override void AfterLoad()
        {
            //Visible = false;
            //Text = shippingCodeTextBox.Text;
            //Visible = true;

        }

        public override void STOP()
        {

            if (globals.DBNAME == "宇豐")
            {
                MessageBox.Show("登入選單錯誤");
                this.SSTOPID = "1";
                return;
            }
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
                    if (cFSCheckBox.Checked == false && eNDCHECKCheckBox.Checked == false)
                    {

                        DialogResult result;
                        result = MessageBox.Show("請確認需不需要保險", "YES/NO", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            this.SSTOPID = "1";
                            cFSCheckBox.Focus();
                            return;
                        }
                        else
                        {
                            eNDCHECKCheckBox.Checked = true;
                        }
                    }
                }

            }

            if (createNameTextBox.Text.ToLower() == "veraliu" || createNameTextBox.Text.ToUpper() == "BELLAZHANG")
            {
                System.Data.DataTable JK1 = GetITEMINVOICE();
                if (JK1.Rows.Count > 0)
                {
                    string SQTY = JK1.Rows[0]["數量"].ToString();
                    string IQTY = JK1.Rows[1]["數量"].ToString();
                    string SAMT = JK1.Rows[0]["金額"].ToString();
                    string IAMT = JK1.Rows[1]["金額"].ToString();

                    if (SQTY != "0" && IQTY != "0")
                    {
                        if (SQTY != IQTY)
                        {
                            DialogResult result;
                            result = MessageBox.Show("INVOICE數量跟主資料不一致，請確認是否要存檔", "YES/NO", MessageBoxButtons.YesNo);
                            if (result == DialogResult.No)
                            {
                                this.SSTOPID = "1";
                                return;
                            }
                        }
                    }

                    if (SSTOPID != "1")
                    {
                        if (SAMT != "0.0000" && IAMT != "0.0000")
                        {
                            if (SAMT != IAMT)
                            {
                                DialogResult result;
                                result = MessageBox.Show("INVOICE金額跟主資料不一致，請確認是否要存檔", "YES/NO", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No)
                                {
                                    this.SSTOPID = "1";
                                    return;
                                }
                            }
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
            add1TextBox.ReadOnly = true;
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


            if (download2DataGridView.Rows.Count > 1)
            {
                download2DataGridView.Columns["filename"].ReadOnly = true;
            }

            if (downloadDataGridView.Rows.Count > 1)
            {
                downloadDataGridView.Columns["DOCDATE"].ReadOnly = true;
                downloadDataGridView.Columns["SA"].ReadOnly = true;
                downloadDataGridView.Columns["SALES"].ReadOnly = true;
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
            button14.Enabled = true;
            button50.Enabled = true;
            button51.Enabled = true;
            button17.Enabled = true;
            button40.Enabled = true;
            button41.Enabled = true;
            button7.Enabled = true;
            button43.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button2.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button12.Enabled = true;
            comboBox1.Enabled = true;

            button19.Enabled = true;
            shippingCodeTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            receivePlaceTextBox.ReadOnly = true;
            goalPlaceTextBox.ReadOnly = true;
            shipmentTextBox.ReadOnly = true;
            unloadCargoTextBox.ReadOnly = true;
            add1TextBox.ReadOnly = true;
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
            add1TextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            nTDollarsTextBox.ReadOnly = true;
            dollarsKindTextBox.ReadOnly = true;
            add7TextBox.ReadOnly = true;

            button14.Enabled = true;
            button50.Enabled = true;
            button51.Enabled = true;
            button17.Enabled = true;
            button40.Enabled = true;
            button41.Enabled = true;
            button7.Enabled = true;
            button43.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button12.Enabled = true;
            comboBox1.Enabled = true;
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
            cFSTableAdapter.Connection = MyConnection;
            markTableAdapter.Connection = MyConnection;
            dataTable2TableAdapter.Connection = MyConnection;
            invoiceDTableAdapter.Connection = MyConnection;
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
            sHIP_FEETableAdapter.Connection = MyConnection;
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
                markTableAdapter.Fill(ship.Mark, MyID);
                try
                {
                    dataTable2TableAdapter.Fill(ship.DataTable2, MyID);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex + "請輸入保險費率");
                }
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
                sHIP_FEETableAdapter.Fill(ship.SHIP_FEE, MyID);

                if (globals.DBNAME == "進金生")
                {
                    System.Data.DataTable K1 = GetOPDN(shippingCodeTextBox.Text);
                    if (K1.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = K1;
                    }
                    else
                    {
                        dataGridView1.DataSource = GetOPDN("1234");
                    }
                }

                SHIPOWTR();

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

                if (globals.DBNAME == "達睿生")
                {

                    shipping_ItemDataGridView.Columns[12].HeaderText = "acme單價";
                    shipping_ItemDataGridView.Columns[13].HeaderText = "acme金額";
                }

                SHIPNO();


                System.Data.DataTable G1 = GetFEE3();
                if (G1.Rows.Count == 0)
                {
                    AddFEE();
                }

                DMARK();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private System.Data.DataTable GetFEE3()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select * from SHIP_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

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
                sHIP_FEEBindingSource.EndEdit();

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


                SqlDataAdapter Adapter15 = util.GetAdapter(sHIP_FEETableAdapter);
                Adapter15.UpdateCommand.Transaction = tx;
                Adapter15.InsertCommand.Transaction = tx;
                Adapter15.DeleteCommand.Transaction = tx;

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
                sHIP_FEETableAdapter.Update(ship.SHIP_FEE);

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
                ship.SHIP_FEE.AcceptChanges();

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
                System.Data.DataTable T1 = GetHR(fmLogin.LoginID.ToString().Trim());
                if (globals.DBNAME == "達睿生")
                {
                    kyes = "DRS" + NumberName + AutoNum + "X";
                }

                else
                {
                    if (T1.Rows.Count > 0)
                    {
                        COMPANY = "達睿生";
                        kyes = NumberName + AutoNum + "D";
                    }
                    else
                    {
                        kyes = NumberName + AutoNum + "X";
                    }
                }


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
            if (fmLogin.LoginID.ToString().Trim().ToUpper() == "MAGGIEWENG")
            {
                add7TextBox.Text = "MaggieWengP";
            }

            iTEMSCheckBox.Checked = false;
            rUSHCheckBox.Checked = false;
            cFSCheckBox.Checked = false;
            kPIYESNOCheckBox.Checked = false;
            iNSUCHECKCheckBox.Checked = false;
            buCardcodeCheckBox.Checked = false;
            add10CheckBox.Checked = false;
            tAXCHECKCheckBox.Checked = false;
            eNDCHECKCheckBox.Checked = false;

            createDateCheckBox.Checked = false;
            this.shipping_mainBindingSource.EndEdit();
            kyes = null;
            quantityTextBox.Text = "未結";
            add1TextBox.Text = "SAP系統";

            if (globals.DBNAME == "達睿生")
            {
                notifyMemoTextBox.Text = "送货地址：" +
                     Environment.NewLine + "公司名稱：" +
                     Environment.NewLine + "英文名稱：" +
                     Environment.NewLine + "公司地址：" +
                     Environment.NewLine + "英文地址：" +
                     Environment.NewLine + "联系方式：" +
                     Environment.NewLine + "PS：" +
                     Environment.NewLine + "" +
                     Environment.NewLine + "" +
                     Environment.NewLine + "" +
                     Environment.NewLine + "仓管备货" +
                     Environment.NewLine + "船务理货" +
                     Environment.NewLine + "外仓提供理货资料" +
                     Environment.NewLine + "仓管放货" +
                     Environment.NewLine + "提供出货文件给外仓" +
                     Environment.NewLine + "提供DRS影本清关文件给报关行报关" +
                     Environment.NewLine + "订车" +
                     Environment.NewLine + "申报，出税单，税金支付成功" +
                     Environment.NewLine + "海关放行" +
                     Environment.NewLine + "送货司机资料：" +
                     Environment.NewLine + "货物送达客人工厂" +
                     Environment.NewLine + "回传签收单";
            }


        }

        public override void AfterCopy()
        {

            if (kyes == null)
            {
                string NumberName = "SH" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                kyes = this.shippingCodeTextBox.Text;
            }
        }
        public override void AfterCopy2()
        {
            COPY = 1;

            tabControl1.SelectedIndex = 0;
            add2TextBox.Text = "";
            add6TextBox.Text = "";
            closeDayTextBox.Text = "";
            forecastDayTextBox.Text = "";
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

            if (fmLogin.LoginID.ToString().Trim().ToUpper() == "MAGGIEWENG")
            {
                add7TextBox.Text = "MaggieWengP";
            }
            createNameTextBox.Text = fmLogin.LoginID.ToString().Trim();

            modifyNameTextBox.Text = "";
            nTDollarsTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");

            buCardcodeCheckBox.Checked = false;
            eNDCHECKCheckBox.Checked = false;
            iNSUPRICECheckBox.Checked = false;
            quantityTextBox.Text = "未結";
            iTEMSCheckBox.Checked = false;
            rUSHCheckBox.Checked = false;
            add10CheckBox.Checked = false;
            createDateCheckBox.Checked = false;
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
            this.shipping_ItemBindingSource.EndEdit();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            if (add1TextBox.Text == "正航系統CHOICE")
            {
                LookupValues = GetMenu.GetCHI();
            }
            else if (add1TextBox.Text == "正航系統INFINITE")
            {
                LookupValues = GetMenu.GetCHI2();
            }
            else if (add1TextBox.Text == "正航系統TOP GARDEN")
            {
                LookupValues = GetMenu.GetCHI4();
            }


            else
            {
                LookupValues = GetMenu.GetMenuList();
            }

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
            catch
            {

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
            bindingNavigator3.Visible = true;
            bindingNavigator4.Visible = true;
            bindingNavigator6.Visible = true;
            button24.Enabled = true;
            button43.Enabled = true;
            contextMenuStrip2.Enabled = false;
            contextMenuStrip3.Enabled = false;
            contextMenuStrip4.Enabled = false;
            contextMenuStrip5.Enabled = false;
            button7.Enabled = true;
            button43.Enabled = true;
            button13.Enabled = true;
            button11.Enabled = true;
            checkBox3.Enabled = true;
            checkBox5.Enabled = true;
            button10.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button12.Enabled = true;
            comboBox1.Enabled = true;
            button14.Enabled = true;
            button50.Enabled = true;
            button51.Enabled = true;
            button17.Enabled = true;
            button40.Enabled = true;
            button41.Enabled = true;
            textBox1.ReadOnly = false;
            button2.Enabled = true;
            add6TextBox.ReadOnly = true;
            button19.Enabled = true;
            textBox2.Text = USER + "@acmepoint.com";
            textBox1.Text = USER + "@acmepoint.com";

            if (USER.ToUpper() == "EVAHSU" || USER.ToUpper() == "LLEYTONCHEN" || USER.ToUpper() == "VERALIU" || USER.ToUpper() == "DANALUO")
            {
                //  checkBox6.Visible = true;
                createDateCheckBox.Visible = true;
            }

            ExcelReport.DELETEFILE();
            ExcelReport.DELETEFOLDER();
            string GROUP = globals.GroupID.ToString().Trim();

            if (GROUP != "EEP" && GROUP != "SHI" && GROUP != "ShipBuy" && GROUP != "WH")
            {
                lcInstro1DataGridView.Columns["LPRICE"].Visible = false;
                lcInstro1DataGridView.Columns["LAMT"].Visible = false;
            }
            if (GROUP != "EEP")
            {
                button45.Visible = false;
                textBox8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                textBox11.Visible = false;
                textBox12.Visible = false;

                button47.Visible = false;
            }


            if (GROUP == "EEP" || GROUP == "ShipBuy")
            {
                comboBox9.Visible = true;

            }
            System.Data.DataTable O2 = GetMenu.GetSHIPOHEM(USER);
            if (O2.Rows.Count > 0)
            {
                DRS = "DRS";

                button8.Visible = true;
            }

            if (globals.DBNAME == "船務測試區")
            {
                DIR = "//acmesrv01//SAP_Share//shipping測試區//";
                PATH = @"\\acmesrv01\SAP_Share\shipping測試區\";
            }
            else if (globals.DBNAME == "達睿生")
            {
                DIR = "//acmesrv01//SAP_Share//shipping達睿生//";
                PATH = @"\\acmesrv01\SAP_Share\shipping達睿生\";
            }

            else
            {
                DIR = "//acmesrv01//SAP_Share//shipping//";
                PATH = @"\\acmesrv01\SAP_Share\shipping\";
            }

            //shippingAD
            shipping_ItemDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;

            if (GROUP != "SHI" && GROUP != "EEP" && GROUP != "ShipBuy" && GROUP != "BOSS" && USER.ToUpper() != "APPLECHEN")
            {
                NOTEXPORT();
            }


            System.Data.DataTable T1 = GetMenu.GetWHSA();
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }



        }
        public void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }

        private void bindingNavigatorAddNewItem3_Click(object sender, EventArgs e)
        {
            createDateCheckBox.Checked = false;
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

                if (INVO == PACK )
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

                if (((DOCTYPE == "銷售訂單" && OUTTYPE == "出口") || (DOCTYPE == "銷售訂單" && OUTTYPE == "三角") || DOCTYPE == "調撥單" || DOCTYPE == "發貨單") && mEMO3TextBox.Text != "")
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


                    this.markBindingSource.EndEdit();
                    this.markTableAdapter.Update(ship.Mark);
                    ship.Mark.AcceptChanges();

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




        private void bindingNavigatorAddNewItem5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetMenu.Getinvoicem(shippingCodeTextBox.Text);

            try
            {
                if (dt1.Rows.Count < 1)
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
                notifyPartTextBox.Text = drw["billTo"].ToString();
                shipperTextBox.Text = "ACMEPOINT  TECHNOLOGY CO., LTD.";
                if (boatNameTextBox.Text != "")
                {
                    oceanVesselTextBox.Text = boatNameTextBox.Text;
                }

                if (unloadCargoTextBox.Text != "")
                {
                    dischargeTextBox.Text = unloadCargoTextBox.Text;
                }

                lADINGMBindingSource.EndEdit();

                System.Data.DataTable dt4 = ship.LADINGD;

                DataRow drw2 = dt4.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["SeqMNo"] = AutoNum;
                drw2["SeqNo"] = "0";
                System.Data.DataTable dt1B = GetMenu.Getgross(shippingCodeTextBox.Text);
                if (dt1B.Rows.Count > 0)
                {
                    drw2["Cargo"] = dt1B.Rows[0]["KGS"].ToString();
                    drw2["Packages"] = dt1B.Rows[0]["PLTS"].ToString();
                    if (receiveDayTextBox.Text.ToUpper() == "AIR")
                    {
                        System.Data.DataTable dt1C = GetPACKCM(shippingCodeTextBox.Text);
                        if (dt1C.Rows.Count > 0)
                        {
                            drw2["Measurement"] = dt1C.Rows[0][0].ToString();
                        }
                    }
                    else
                    {
                        drw2["Measurement"] = dt1B.Rows[0]["CBM"].ToString();
                    }
                }


                System.Data.DataTable GINV = GetINVOM(shippingCodeTextBox.Text);
                if (GINV.Rows.Count > 0)
                {
                    drw2["Description"] = "INV NO: " + GINV.Rows[0][0].ToString();

                }
                dt4.Rows.Add(drw2);

                drw2 = dt4.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["SeqMNo"] = AutoNum;
                drw2["SeqNo"] = "1";

                if (GINV.Rows.Count > 0)
                {
                    drw2["Description"] = "JOB NO: " + shippingCodeTextBox.Text;

                }

                if (receiveDayTextBox.Text.ToUpper() == "AIR")
                {
                    System.Data.DataTable dt1C = GetPACKCM(shippingCodeTextBox.Text);
                    if (dt1C.Rows.Count > 1)
                    {
                        drw2["Measurement"] = dt1C.Rows[1][0].ToString();
                    }
                }
                dt4.Rows.Add(drw2);

                System.Data.DataTable dt1C3 = GetPACKCM(shippingCodeTextBox.Text);
                if (dt1C3.Rows.Count > 2)
                {

                    drw2 = dt4.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqMNo"] = AutoNum;
                    drw2["SeqNo"] = "2";
                    drw2["Description"] = "";
                    drw2["Measurement"] = dt1C3.Rows[2][0].ToString();
                    dt4.Rows.Add(drw2);
                }

                lADINGDBindingSource.EndEdit();
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            string CONN = "";
            try
            {


                if (add1TextBox.Text == "正航系統CHOICE")
                {
                    CONN = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                }

                if (add1TextBox.Text == "正航系統INFINITE")
                {
                    CONN = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";

                }

                if (add1TextBox.Text == "正航系統TOP GARDEN")
                {
                    CONN = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }

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


                if (add1TextBox.Text == "正航系統CHOICE" || add1TextBox.Text == "正航系統INFINITE" || add1TextBox.Text == "正航系統TOP GARDEN")
                {



                    System.Data.DataTable dt1CHO = GetCHO3(shipping_OBUTextBox.Text, CONN);

                    System.Data.DataTable dt2CHO = GetCHO2(cardCodeTextBox.Text, CONN);
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

                }
                else if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1349-00")
                {



                    System.Data.DataTable dt1 = GetMenu.Getaddress2(add2TextBox.Text);
                    if (dt1.Rows.Count > 0)
                    {
                        DataRow drw = dt1.Rows[0];
                        oBUBillToTextBox1.Text = drw["公司全稱"].ToString() + "\r\n" + drw["地址"].ToString() + "\r\n" + "TEL:" + drw["電話"].ToString() + "\r\n" + "FAX:" + drw["傳真"].ToString() + "\r\n" + "ATTN:" + drw["大名"].ToString();
                    }




                    StringBuilder sb2 = new StringBuilder();

                    System.Data.DataTable dtPO = GetMenu.GetPO(shippingCodeTextBox.Text);
                    if (dtPO.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dtPO.Rows.Count - 1; i++)
                        {

                            DataRow dd = dtPO.Rows[i];


                            sb2.Append(dd["pino"].ToString() + ",");


                        }

                        sb2.Remove(sb2.Length - 1, 1);
                        pOnoTextBox.Text = sb2.ToString();
                    }


                    System.Data.DataTable dt1s = GetMenu.Getocrdnew1(aa, aas);
                    if (dt1s.Rows.Count > 0)
                    {
                        DataRow drw3 = dt1s.Rows[0];

                        shipToTextBox.Text = drw3["shipbuilding"].ToString() +
                        Environment.NewLine + drw3["shipstreet"].ToString() +
                        Environment.NewLine + "TEL:" + drw3["shipblock"].ToString() +
                        Environment.NewLine + "FAX:" + drw3["shipcity"].ToString() +
                        Environment.NewLine + "ATTN:" + drw3["shipzipcode"].ToString();


                        billToTextBox.Text = drw3["billbuilding"].ToString() +
                        Environment.NewLine + drw3["billstreet"].ToString() +
                        Environment.NewLine + "TEL:" + drw3["billblock"].ToString() +
                        Environment.NewLine + "FAX:" + drw3["billcity"].ToString() +
                        Environment.NewLine + "ATTN:" + drw3["billzipcode"].ToString();
                    }
                }

                else
                {


                    if (aas == "發貨單" || aas == "調撥單")
                    {
                        System.Data.DataTable dt1G = GetMenu.Getaddress2(cardCodeTextBox.Text);
                        if (dt1G.Rows.Count > 0)
                        {

                            DataRow drw = dt1G.Rows[0];
                            //DataRow drw1 = dt1G.Rows[1];
                            billToTextBox.Text = drw["公司全稱"].ToString() + "\r\n" + drw["地址"].ToString() + "\r\n" + "TEL:" + drw["電話"].ToString() + "\r\n" + "FAX:" + drw["傳真"].ToString() + "\r\n" + "ATTN:" + drw["大名"].ToString();
                        }
                        //  shipToTextBox.Text = drw1["公司全稱"].ToString() + "\r\n" + drw1["地址"].ToString() + "\r\n" + "TEL:" + drw1["電話"].ToString() + "\r\n" + "FAX:" + drw1["傳真"].ToString() + "\r\n" + "ATTN:" + drw1["大名"].ToString();
                    }
                    else if (aas == "銷售訂單" || aas == "AR貸項")
                    {
                        StringBuilder sb2 = new StringBuilder();

                        System.Data.DataTable dtPO = GetMenu.GetPO(shippingCodeTextBox.Text);
                        if (dtPO.Rows.Count > 0)
                        {
                            for (int i = 0; i <= dtPO.Rows.Count - 1; i++)
                            {

                                DataRow dd = dtPO.Rows[i];


                                sb2.Append(dd["pino"].ToString() + ",");


                            }

                            sb2.Remove(sb2.Length - 1, 1);
                            pOnoTextBox.Text = sb2.ToString();
                        }


                        System.Data.DataTable dt1 = GetMenu.Getocrdnew1(aa, aas);

                        if (dt1.Rows.Count > 0)
                        {

                            DataRow drw = dt1.Rows[0];


                            shipToTextBox.Text = drw["shipbuilding"].ToString() +
                            Environment.NewLine + drw["shipstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();


                            billToTextBox.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();
                        }

                    }
                    else if (aas == "採購退貨")
                    {
                        StringBuilder sb2 = new StringBuilder();

                        System.Data.DataTable dtPO = GetMenu.GetPO(shippingCodeTextBox.Text);
                        if (dtPO.Rows.Count > 0)
                        {
                            for (int i = 0; i <= dtPO.Rows.Count - 1; i++)
                            {

                                DataRow dd = dtPO.Rows[i];


                                sb2.Append(dd["pino"].ToString() + ",");


                            }

                            sb2.Remove(sb2.Length - 1, 1);
                            pOnoTextBox.Text = sb2.ToString();
                        }


                        System.Data.DataTable dt1 = GetMenu.Getocrdnew1(aa, aas);

                        if (dt1.Rows.Count > 0)
                        {

                            DataRow drw = dt1.Rows[0];

                            billToTextBox.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();

                            shipToTextBox.Text = billToTextBox.Text;
                        }

                    }
                    else if (aas == "AP貸項")
                    {
                        StringBuilder sb2 = new StringBuilder();

                        System.Data.DataTable dtPO = GetMenu.GetPO(shippingCodeTextBox.Text);
                        if (dtPO.Rows.Count > 0)
                        {
                            for (int i = 0; i <= dtPO.Rows.Count - 1; i++)
                            {

                                DataRow dd = dtPO.Rows[i];


                                sb2.Append(dd["pino"].ToString() + ",");


                            }

                            sb2.Remove(sb2.Length - 1, 1);
                            pOnoTextBox.Text = sb2.ToString();
                        }


                        System.Data.DataTable dt1 = GetMenu.Getocrdnew1(aa, aas);

                        if (dt1.Rows.Count > 0)
                        {

                            DataRow drw = dt1.Rows[0];

                            billToTextBox.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();

                            shipToTextBox.Text = billToTextBox.Text;
                        }

                    }
                    else if (aas == "採購訂單")
                    {
                        if (globals.DBNAME == "達睿生")
                        {
                            string T1 = "DRS Tech.(Shenzhen)Ltd." +
                                      Environment.NewLine + "达睿生科技发展(深圳)有限公司" +
                                      Environment.NewLine + "RM2102,YIFANG Building,Number 315,Shuang Ming Avenue,Guang Ming District,ShenZhen,China (P.C.518107)" +
                                      Environment.NewLine + "DANZHUTOU INDUSTRIAL PARK，LONGGANG DISTRICT，" +
                                      Environment.NewLine + "Shenzhen，China." +
                                      Environment.NewLine + " 深圳市光明区光明街道东周社区双明大道315号易方大厦2102室" +
                                      Environment.NewLine + "TEL:+86-755-25911195" +
                                      Environment.NewLine + "FAX:+86-755-25911201" +
                                      Environment.NewLine + "统一信用代码：91440300564218558D";
                            billToTextBox.Text = T1;
                            shipToTextBox.Text = T1;
                        }
                        else
                        {
                            System.Data.DataTable dt1 = GetMenu.Getocrdnew1(aa, aas);

                            if (dt1.Rows.Count > 0)
                            {

                                billToTextBox.Text = dt1.Rows[0][0].ToString().Replace("\r", System.Environment.NewLine);
                                shipToTextBox.Text = dt1.Rows[0][1].ToString().Replace("\r", System.Environment.NewLine);
                            }

                        }
                    }

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
                        string PRODID = drw["bb"].ToString();
                        string ITEMCODE = drw["ITEMCODE"].ToString();
                        string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                        StringBuilder sb = new StringBuilder();
                        foreach (string ESi in arrurl)
                        {
                            sb.Append("'" + ESi + "',");
                        }
                        sb.Remove(sb.Length - 1, 1);
                        if (cardCodeTextBox.Text == "1279-03")
                        {
                            System.Data.DataTable OI1 = GetWHPACK2ES2(sb.ToString(), ITEMCODE);
                            if (OI1.Rows.Count > 0)
                            {
                                string OINAME = "";
                                string MODEL = OI1.Rows[0][0].ToString();
                                string GRADE = OI1.Rows[0][1].ToString();

                                string OIES = OI1.Rows[0][2].ToString();
                                string TMODEL = OI1.Rows[0][3].ToString();
                                if (kPIYESNOCheckBox.Checked)
                                {
                                    OINAME = MODEL + GRADE;
                                }
                                else
                                {
                                    OINAME = MODEL;

                                }

                                System.Data.DataTable OI2 = GetSHIPOITM2(TMODEL);
                                if (OI2.Rows.Count > 0)
                                {
                                    System.Data.DataTable OI3 = GetSHIPOITM4(ITEMCODE);
                                    if (OI3.Rows.Count > 0)
                                    {
                                        OINAME = OINAME + OI3.Rows[0][0].ToString();
                                    }
                                }

                                if (!String.IsNullOrEmpty(OIES))
                                {
                                    OIES = " (" + OIES + ")";

                                }

                                PRODID = OINAME + OIES;
                            }
                            else
                            {
                                System.Data.DataTable OI2 = GetWHPACK2ES3(ITEMCODE);
                                if (OI2.Rows.Count > 0)
                                {
                                    string MODEL = OI2.Rows[0][0].ToString();
                                    string GRADE = OI2.Rows[0][1].ToString();
                                    if (kPIYESNOCheckBox.Checked)
                                    {
                                        PRODID = MODEL + GRADE;
                                    }
                                }
                            }
                        }



                        if (String.IsNullOrEmpty(PRODID))
                        {
                            if (add1TextBox.Text == "正航系統CHOICE" || add1TextBox.Text == "正航系統INFINITE" || add1TextBox.Text == "正航系統TOP GARDEN")
                            {
                                System.Data.DataTable J1 = GetCHOITEM(drw["itemcode"].ToString(), CONN);
                                if (J1.Rows.Count > 0)
                                {
                                    PRODID = J1.Rows[0][0].ToString();
                                }

                            }
                        }
                        drw2["INDescription"] = PRODID;
                        drw2["InQty"] = drw["Quantity"];
                        drw2["UnitPrice"] = drw["ItemPrice"];
                        drw2["CURRENCY"] = drw["CURRENCY"];
                        drw2["RATE"] = drw["RATE"];
                        drw2["RATEUSD"] = drw["RATEUSD"];
                        drw2["ITEMCODE"] = ITEMCODE;

                        string TYPE = drw["OLDORDER"].ToString();
                        int T1 = add1TextBox.Text.IndexOf("正航系統");
                        if (T1 == -1)
                        {

                            drw2["amount"] = 1;

                            drw2["SOID"] = drw["Docentry"];

                        }
                        else
                        {
                            drw2["amount"] = drw["ItemAmount"];
                        }


                        if (DRS == "DRS")
                        {
                            drw2["amount"] = drw["ItemAmount"];
                        }


                        drw2["LINENUM"] = drw["linenum"];


                        drw2["CHOPrice"] = drw["CHOPrice"];
                        drw2["CHOAmount"] = drw["CHOAmount"];
                        drw2["TREETYPE"] = TYPE;
                        drw2["VISORDER"] = drw["VISORDER"];

                        Clear(sbS);
                        SBS();

                        System.Data.DataTable dtF = GetWHLOCATION(sbS.ToString(), drw["ITEMCODE"].ToString());
                        if (dtF.Rows.Count > 0)
                        {
                            drw2["LOCATION"] = dtF.Rows[0][0].ToString();

                        }

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

            iRecs = packingListDDataGridView.Rows.Count - 1;
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
                int G = Convert.ToInt32(AMT.ToString("###0"));
                double AMT1 = Convert.ToDouble(AMT);

                if (G != 0)
                {
                    if (CURRENCY2 == "RMB")
                    {

                        amountTotalEngTextBox.Text = "SAY TOTAL : RMB DOLLARS " + new Class1().NumberToString(AMT1);
                        amountTotalTextBox.Text = "RMB" + t1.Rows[0][0].ToString();
                    }
                    else
                    {
                        amountTotalEngTextBox.Text = "SAY TOTAL : US DOLLARS " + new Class1().NumberToString(AMT1);
                        amountTotalTextBox.Text = "US" + t1.Rows[0][0].ToString();
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
        private System.Data.DataTable checkCn21(string Docentry)
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("select G.ProdID  AS INDescription,CAST(G.Quantity AS INT)  AS InQty,Price as UnitPrice,Amount ");
            sb.Append(" from  OrdBillMain A   ");
            sb.Append(" Inner Join OrdBillSub G  On (G.Flag=A.Flag  And G.BillNO=A.BillNO)    ");

            sb.Append("where a.flag=2 AND A.BILLNO=@DOCENTRY ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", Docentry));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicePart");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable checkPartNumber(string DocEntry)
        {
            SqlConnection connection = new SqlConnection(strCn02);
            StringBuilder sb = new StringBuilder();
            string[] DocentryCount = DocEntry.Split(',');
            sb.Append("SELECT CONCAT('  '+ B.U_ITEMNAME +' '+ B.U_MODEL + ' * ',CAST(Quantity AS decimal(16,0)),'PCS') AS INDescription,'' AS InQty,'' AS UnitPrice,'' AS Amount  ");
            sb.Append("FROM RDR1 AS A ");
            sb.Append("LEFT JOIN OITM B ON A.ITEMCODE = B.ITEMCODE  WHERE  A.TreeType = 'I' and");
            sb.Append("(");
            for (int i = 0; i < DocentryCount.Length; i++)
            {
                if (i == 0)
                {
                    sb.Append(" A.Docentry = @Docentry" + i);

                }
                else
                {
                    sb.Append(" or A.Docentry = @Docentry" + i);
                }
            }
            sb.Append(")");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            for (int i = 0; i < DocentryCount.Length; i++)
            {
                command.Parameters.Add(new SqlParameter("@DOCENTRY" + i, DocentryCount[i]));
            }


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicePart");
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
                    QTY = packingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn48"].Value.ToString();
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
                    string f = "";
                    string f2 = "";
                    CalcTotals2();

                    string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                    StringBuilder sb = new StringBuilder();
                    foreach (string i in arrurl)
                    {
                        sb.Append("'" + i + "',");
                    }
                    sb.Remove(sb.Length - 1, 1);
                    System.Data.DataTable dt3 = GetWHPACK2(sb.ToString());

                    int k = packingListDDataGridView.Rows.Count - 2;

                    DataGridViewRow rowP;

                    rowP = packingListDDataGridView.Rows[k];
                    string a0 = rowP.Cells["PackageNo"].Value.ToString().Trim();
                    string a1 = rowP.Cells["dataGridViewTextBoxColumn46"].Value.ToString().Trim();
                    System.Data.DataTable G1 = GetSHIPPACK8(shippingCodeTextBox.Text, pLNoTextBox.Text);
                    if (G1.Rows.Count > 0)
                    {
                        f = G1.Rows[0][0].ToString();
                        f2 = G1.Rows[0][1].ToString();
                    }

                    //if(dOCTYPETextBox.Text =="銷售")
                    //{
                    System.Data.DataTable B1 = GetB1(shippingCodeTextBox.Text, pLNoTextBox.Text);
                    if (B1.Rows.Count > 1)
                    {
                        int M1 = 0;
                        int M2 = 0;
                        int B3 = 0;
                        for (int i = 0; i <= B1.Rows.Count - 1; i++)
                        {
                            string WHNO = B1.Rows[i][0].ToString();
                            System.Data.DataTable B2 = GetB2(WHNO);
                            if (B2.Rows.Count > 0)
                            {
                                int n;
                                string B2S = B2.Rows[0][0].ToString();
                                if (int.TryParse(B2S, out n))
                                {
                                    M1 += Convert.ToInt16(B2S);
                                }
                            }
                            if (i == B1.Rows.Count - 1)
                            {
                                System.Data.DataTable BB3 = GetB3(WHNO);
                                if (BB3.Rows.Count > 0)
                                {
                                    string PALENO = BB3.Rows[0][0].ToString().Trim();
                                    if (String.IsNullOrEmpty(PALENO))
                                    {
                                        PALENO = "1";
                                    }
                                    if (PALENO == B1.Rows.Count.ToString())
                                    {
                                        PALENO = "1";
                                    }
                                    if (PALENO != "1")
                                    {
                                        B3 = 1;
                                    }
                                }
                            }
                            System.Data.DataTable B2C = GetB2CNO(WHNO);
                            if (B2C.Rows.Count > 0)
                            {
                                int n;
                                string B2S = B2C.Rows[0][0].ToString();
                                if (int.TryParse(B2S, out n))
                                {
                                    M2 += Convert.ToInt16(B2S);
                                }
                            }
                        }
                        if (B3 == 0)
                        {
                            f = M1.ToString();
                            f2 = M2.ToString();
                        }

                    }
                    else
                    {

                        System.Data.DataTable B2 = GetB2S(shippingCodeTextBox.Text, pLNoTextBox.Text);
                        if (B2.Rows.Count > 1)
                        {
                            int PACKS = 0;

                            int PACK = 0;
                            int P3 = 0;
                            int PACKD = 0;

                            int GGG = 0;
                            int SPACK = 0;

                            for (int i = 0; i <= B2.Rows.Count - 1; i++)
                            {


                                string DESC = B2.Rows[i][1].ToString();
                                PACK = Convert.ToInt16(B2.Rows[i][0]);
                                P3 = Convert.ToInt16(B2.Rows[i][2]);

                                if (i == 0 && P3 != 1 && dOCTYPETextBox.Text == "銷售訂單")
                                {
                                    SPACK = PACK - 1;
                                }

                                if (PACK <= PACKD)
                                {
                                    PACKS += PACKD;
                                }
                                else if (boardCountNoTextBox.Text == "進口" && dOCTYPETextBox.Text == "採購")
                                {
                                    PACKS += PACKD;
                                }
                                if (i == B2.Rows.Count - 1)
                                {
                                    PACKS += PACK;
                                    if (DESC == "")
                                    {
                                        GGG = 1;
                                    }
                                }
                                PACKD = PACK;
                                if (boardCountNoTextBox.Text == "出口" && dOCTYPETextBox.Text == "採購")
                                {
                                    PACKS += PACK;
                                }
                            }
                            if (GGG == 0 && SPACK != 0)
                            {
                                f = (PACKS - SPACK).ToString();
                            }

                            if (dt3.Rows.Count == 0)
                            {
                                if (boardCountNoTextBox.Text == "出口" && dOCTYPETextBox.Text == "採購")
                                {
                                    f = PACKS.ToString();
                                }
                            }
                        }

                        System.Data.DataTable B2S = GetB2S2(shippingCodeTextBox.Text, pLNoTextBox.Text);
                        if (B2S.Rows.Count > 1)
                        {
                            int GGG = 0;
                            int CNOS = 0;
                            int CNOS2 = 0;
                            int CNOD = 0;
                            int CNO = 0;

                            for (int i = 0; i <= B2S.Rows.Count - 1; i++)
                            {
                                string DESC = B2S.Rows[i][1].ToString();
                                CNO = Convert.ToInt16(B2S.Rows[i][0]);

                                if (CNO <= CNOD)
                                {
                                    CNOS += CNOD;
                                }
                                else if (boardCountNoTextBox.Text == "進口" && dOCTYPETextBox.Text == "採購")
                                {
                                    CNOS += CNOD;
                                }
                                //else if (cardCodeTextBox.Text == "1362-00")
                                //{
                                //    CNOS += CNOD;
                                //}
                                if (i == B2S.Rows.Count - 1)
                                {
                                    CNOS += CNO;

                                    if (DESC == "")
                                    {
                                        GGG = 1;
                                    }
                                }
                                CNOD = CNO;

                                if (boardCountNoTextBox.Text == "出口" && dOCTYPETextBox.Text == "採購")
                                {
                                    CNOS2 += CNO;
                                }
                            }
                            if (GGG == 0)
                            {
                                f2 = CNOS.ToString();
                            }
                            if (boardCountNoTextBox.Text == "出口" && dOCTYPETextBox.Text == "採購")
                            {
                                f2 = CNOS2.ToString();
                            }
                        }

                        //if (cardCodeTextBox.Text == "1362-00")
                        //{
                        //    System.Data.DataTable B2SF = GetSUMPACK();
                        //    if (B2SF.Rows.Count > 0)
                        //    {
                        //        f2 = B2SF.Rows[0][0].ToString();
                        //    }
                        //}
                    }



                    if (dt3.Rows.Count == 2)
                    {
                        System.Data.DataTable dt4 = GetWHPACK3(sb.ToString());
                        int SEQNO = 0;

                        int PLATENOD = 0;
                        string MATERIALD = "";
                        for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                        {
                            string PLATENO = dt4.Rows[i]["PLATENO"].ToString();
                            string MATERIAL = dt4.Rows[i]["MATERIAL"].ToString();
                            string ID = dt4.Rows[i]["ID"].ToString();
                            if (MATERIALD != MATERIAL)
                            {
                                SEQNO++;
                            }
                            if (!String.IsNullOrEmpty(PLATENO))
                            {
                                int PLATENOI = Convert.ToInt16(PLATENO);

                                if (PLATENOI < PLATENOD)
                                {
                                    SEQNO++;
                                }

                                PLATENOD = Convert.ToInt16(PLATENO);
                            }
                            UPWHPACK(SEQNO.ToString(), ID);

                            MATERIALD = MATERIAL;
                        }

                        //System.Data.DataTable H1 = GetWHPACK5N(sb.ToString());
                        //int PLATENODD = 0;
                        //int PLATENODD2 = 0;
                        //if (H1.Rows.Count > 0)
                        //{

                        //    for (int i = 0; i <= H1.Rows.Count - 1; i++)
                        //    {
                        //        int PLATENO = Convert.ToInt16(H1.Rows[i][0]);

                        //        if (PLATENO > PLATENODD)
                        //        {
                        //            PLATENODD2 = 0;
                        //        }
                        //        else
                        //        {
                        //            PLATENODD2 += PLATENODD;
                        //        }

                        //        if (i == H1.Rows.Count - 1)
                        //        {
                        //            PLATENODD2 += PLATENO;
                        //        }

                        //        PLATENODD = PLATENO;
                        //    }

                        //    f = PLATENODD2.ToString();
                        //}
                    }
                    if (shippingCodeTextBox.Text == "SH20200513009X")
                    {
                        f = "4";
                    }

                    sayTotalTextBox.Text = f;
                    //if (f == "0")
                    //{
                    //    sayTotalTextBox.Text = f2;
                    //}
                    sayCTNTextBox.Text = f2;

                    string f3 = "";
                    if (userNameTextBox.Text != "0" && userNameTextBox.Text != "")
                    {
                        f3 = " + " + userNameTextBox.Text + " EMPTY CTNS";
                    }
                    if (f != "")
                    {
                        int amountText = Convert.ToInt32(f);
                        string s = f;
                        if (createDateCheckBox.Checked)
                        {
                        }
                        else
                        {
                            columnTotalTextBox.Text = new Class1().NumberToString2(amountText, s, f2) + f3 + " ONLY.";
                        }
                    }



                    UPPACK();

                    if (sendGoodsTextBox.Text == "" || sendGoodsTextBox.Text == "0.00")
                    {
                        GETCBM("1");

                    }

                    if (boardCountNoTextBox.Text != "進口")
                    {
                        GETCBM("1");
                    }


                    if (cBMTextBox.Text == "" || cBMTextBox.Text == "0.00")
                    {
                        GETCBM("2");

                    }

                    if (boardCountNoTextBox.Text != "進口")
                    {
                        GETCBM("2");
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

                    markTableAdapter.Fill(ship.Mark, MyID);
                    MessageBox.Show("儲存成功");

                    System.Data.DataTable SHIPPCAK = GetSHIPPCAK();
                    decimal n;

                    if (SHIPPCAK.Rows.Count > 0)
                    {

                        for (int i = 0; i <= SHIPPCAK.Rows.Count - 1; i++)
                        {
                            decimal N1 = 0;
                            decimal G1 = 0;
                            string NET = SHIPPCAK.Rows[i]["NET"].ToString();
                            string GROSS = SHIPPCAK.Rows[i]["GROSS"].ToString();
                            if (decimal.TryParse(NET, out n))
                            {
                                N1 = Convert.ToDecimal(NET);
                            }

                            if (decimal.TryParse(GROSS, out n))
                            {
                                G1 = Convert.ToDecimal(GROSS);
                            }

                            if (N1 > G1)
                            {
                                string MESS = "Net " + N1 + "大於" + "Gross " + GROSS;
                                MessageBox.Show(MESS);
                            }

                        }



                    }

                    decimal N2 = 0;
                    decimal G2 = 0;
                    string NET2 = netTextBox.Text;
                    string GROSS2 = grossTextBox.Text;
                    if (decimal.TryParse(NET2, out n))
                    {
                        N2 = Convert.ToDecimal(NET2);
                    }

                    if (decimal.TryParse(GROSS2, out n))
                    {
                        G2 = Convert.ToDecimal(GROSS2);
                    }

                    if (N2 > G2)
                    {
                        string MESS = "總計 Net " + NET2 + "大於" + "Gross " + GROSS2;
                        MessageBox.Show(MESS);
                    }
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

        private void GETCBM(string DTYPE)
        {
            System.Data.DataTable K2 = GetWHPACKCBM2(shippingCodeTextBox.Text, DTYPE);
            decimal CBB = 0;
            if (K2.Rows.Count > 0)
            {
                for (int i = 0; i <= K2.Rows.Count - 1; i++)
                {
                    string CM = K2.Rows[i][0].ToString();
                    string[] CMS2 = CM.ToUpper().Split(new Char[] { '/' });
                    foreach (string F2 in CMS2)
                    {
                        string[] CMS = F2.ToUpper().Split(new Char[] { 'X' });
                        StringBuilder sbS = new StringBuilder();
                        int M1 = 0;
                        string L = "";
                        string W = "";
                        string H = "";
                        int T2 = -1;
                        string DD = "1";
                        foreach (string F in CMS)
                        {
                            M1++;
                            if (M1 == 1)
                            {
                                L = F;
                            }
                            if (M1 == 2)
                            {
                                W = F;
                            }
                            if (M1 == 3)
                            {
                                T2 = F.IndexOf("*");
                                if (T2 != -1)
                                {
                                    DD = F.Substring(T2 + 1, F.Length - T2 - 1);
                                    H = F.Substring(0, T2);
                                }
                                else
                                {
                                    H = F;
                                }
                            }
                        }
                        decimal n;

                        decimal GA = 1000000;
                        L = L.Replace("'@", "");
                        L = L.Replace("@", "");
                        int DD1 = 0;
                        int DF1 = CM.IndexOf("*");
                        if (DF1 == -1 && boardCountNoTextBox.Text != "進口")
                        {
                            if (i > 0)
                            {
                                int P1 = Convert.ToInt16(K2.Rows[i][1]);
                                int P2 = Convert.ToInt16(K2.Rows[i - 1][1]);
                                string PACK2 = K2.Rows[i][2].ToString();
                                DD1 = Convert.ToInt16(K2.Rows[i][1]) - Convert.ToInt16(K2.Rows[i - 1][1]);
                                //if (DD1 < 0)
                                //{
                                int F1 = PACK2.IndexOf("-");

                                if (F1 != -1)
                                {
                                    string[] arrurl = PACK2.Split(new Char[] { '-' });
                                    string G1 = "";
                                    string G2 = "";
                                    int L1 = 0;
                                    foreach (string F in arrurl)
                                    {
                                        L1++;

                                        if (L1 == 1)
                                        {
                                            G1 = F;
                                        }
                                        if (L1 == 2)
                                        {
                                            G2 = F;
                                        }
                                    }
                                    DD1 = Convert.ToInt16(G2) - Convert.ToInt16(G1) + 1;

                                    // DD = "2";
                                }
                                else
                                {
                                    DD1 = 1;
                                }
                                //}
                                //else
                                //{
                                //    DD1 = 1;
                                //}
                            }
                            else
                            {
                                DD1 = Convert.ToInt16(K2.Rows[i][1]);
                            }
                        }
                        if (DF1 == -1 && dOCTYPETextBox.Text == "調撥單")
                        {
                            string PACK = K2.Rows[i][2].ToString();

                            int F1 = PACK.IndexOf("-");

                            if (F1 != -1)
                            {
                                string[] arrurl = PACK.Split(new Char[] { '-' });
                                string G1 = "";
                                string G2 = "";
                                int L1 = 0;
                                foreach (string F in arrurl)
                                {
                                    L1++;

                                    if (L1 == 1)
                                    {
                                        G1 = F;
                                    }
                                    if (L1 == 2)
                                    {
                                        G2 = F;
                                    }
                                }
                                DD1 = Convert.ToInt16(G2) - Convert.ToInt16(G1) + 1;
                            }

                            else
                            {
                                DD1 = 1;
                            }
                        }
                        if (DD == "1")
                        {
                            DD = DD1.ToString();
                        }
                        if (decimal.TryParse(L, out n) && decimal.TryParse(W, out n) && decimal.TryParse(H, out n) && decimal.TryParse(DD, out n))
                        {
                            if (DD == "0")
                            {
                                DD = "1";
                            }
                            decimal ff3 = (Convert.ToDecimal(L) * Convert.ToDecimal(W) * Convert.ToDecimal(H)) * Convert.ToDecimal(DD);
                            CBB += ff3 / GA;
                        }
                    }
                }

                if (DTYPE == "1")
                {
                    sendGoodsTextBox.Text = CBB.ToString("0.00");
                }

                if (DTYPE == "2")
                {
                    cBMTextBox.Text = CBB.ToString("0.00");
                }
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
            UPPACK();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            if (globals.DBNAME == "達睿生")
            {
                System.Data.DataTable G1 = GetshipTYPE();
                if (G1.Rows.Count > 0)
                {

                    int T2 = cardNameTextBox.Text.IndexOf("進金生");
                    int T3 = cardNameTextBox.Text.IndexOf("进金生");
                    int T4 = cardNameTextBox.Text.IndexOf("友達");
                    string CARD = "";
                    string ADD = "";
                    string TEL = "";
                    if (T2 != -1 || T3 != -1)
                    {
                        FileName = lsAppDir + "\\Excel\\DRS\\PACKDRSAP2.xls";
                    }
                    else if (T4 != -1)
                    {
                        FileName = lsAppDir + "\\Excel\\DRS\\PACKDRSAP.xls";
                    }
                    else
                    {
                        FileName = lsAppDir + "\\Excel\\DRS\\PACKDRSAP3.xls";
                        System.Data.DataTable K1 = GetOrderData2DRS2();
                        if (K1.Rows.Count > 0)
                        {
                            CARD = K1.Rows[0]["CARD"].ToString();
                            ADD = K1.Rows[0]["ADD"].ToString();
                            TEL = K1.Rows[0]["TEL"].ToString();
                        }
                    }
                    GetExcelProduct(FileName, GetOrderData3DRS(CARD, ADD, TEL), "N", "N");

                }
                else
                {


                    FileName = lsAppDir + "\\Excel\\DRS\\PACKDRS.xls";
                    GetExcelProduct(FileName, GetOrderData3(), "N", "N");
                }
            }

            else
            {
                string OHEM = fmLogin.LoginID.ToString().ToUpper();
                if (DRS == "DRS")
                {
                    FileName = lsAppDir + "\\Excel\\DRS\\PACKDRSACMEDRS.xls";
                    GetExcelProduct(FileName, GetOrderData3(), "N", "N");
                }
                else if (OHEM == "EVAHSU" || OHEM == "BETTYTSENG")
                {
                    FileName = lsAppDir + "\\Excel\\DRS\\PACKDRSACME.xls";
                    GetExcelProduct(FileName, GetOrderData3(), "Y", "N");
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\PACK.xls";
                    GetExcelProduct(FileName, GetOrderData3(), "N", "N");
                }

            }


        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (cardCodeTextBox.Text == "")
            {
                MessageBox.Show("請輸入客戶編號");
                return;
            }

            try
            {
                SHIPAP frm1 = new SHIPAP();
                frm1.cardcode = cardCodeTextBox.Text;
                frm1.CLOSE = checkBox1.CheckState.ToString();
                if (frm1.ShowDialog() == DialogResult.OK)
                {
                    string ss = frm1.a.ToString();

                    tabControl1.SelectedIndex = 0;
                    System.Data.DataTable dt1 = null;
                    string NAME = globals.DBNAME;
                    if (NAME == "進金生")
                    {
                        dt1 = GetMenu.GetOrdr2(ss);
                    }
                    else
                    {
                        dt1 = GetMenu.GetOrdr2DRS(ss);
                    }


                    System.Data.DataTable dt2 = ship.Shipping_Item;

                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {

                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["seqNo"] = "0";
                        string DOC = drw["Docnum"].ToString();
                        drw2["Docentry"] = DOC;
                        drw2["ItemCode"] = drw["ItemCode"];
                        drw2["Dscription"] = drw["Dscription"];
                        drw2["ItemRemark"] = "採購訂單";
                        drw2["Quantity"] = drw["QTY"];
                        drw2["ItemPrice"] = drw["Price"];
                        drw2["linenum"] = drw["linenum"];
                        drw2["ItemAmount"] = drw["totalfrgn"];
                        drw2["VISORDER"] = drw["VISORDER"];


                        System.Data.DataTable B1 = GetDOCCUR(DOC, "OPOR");
                        if (B1.Rows.Count > 0)
                        {
                            drw2["CURRENCY"] = B1.Rows[0][0].ToString();
                        }
                        dt2.Rows.Add(drw2);


                    }

                    pinoTextBox.Text = dt1.Rows[0][0].ToString();



                    for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
                    {
                        shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
                    }


                    shipping_mainBindingSource.EndEdit();
                    shipping_ItemBindingSource.EndEdit();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {


                object[] LookupValues = GetMenu.GetMenuOw();

                if (LookupValues != null)
                {
                    tabControl1.SelectedIndex = 0;

                    string pino = pinoTextBox.Text;
                    pinoTextBox.Text = Convert.ToString(LookupValues[0]);
                    string docentry = pinoTextBox.Text;

                    System.Data.DataTable dt1 = GetMenu.GetOwtr(docentry);

                    System.Data.DataTable dt2 = ship.Shipping_Item;


                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["Docentry"] = drw["Docnum"];
                        drw2["seqNo"] = "0";
                        drw2["linenum"] = drw["linenum"];
                        drw2["ItemCode"] = drw["ItemCode"];
                        drw2["Dscription"] = drw["Dscription"];
                        drw2["WHSCODE"] = drw["Filler"];
                        string QTY = drw["Quantity"].ToString();
                        int R = QTY.IndexOf(".");
                        if (R != -1)
                        {
                            QTY = QTY.Substring(0, R);
                        }
                        drw2["Quantity"] = QTY;
                        drw2["ItemPrice"] = "0";
                        drw2["ItemAmount"] = "0";
                        drw2["ItemRemark"] = "調撥單";
                        drw2["VISORDER"] = drw["VISORDER"];
                        drw2["Remark"] = drw["comments"];
                        dt2.Rows.Add(drw2);

                    }


                    for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
                    {
                        shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
                    }
                    shipping_mainBindingSource.EndEdit();
                    shipping_ItemBindingSource.EndEdit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

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

                if (dt2.Rows.Count > 0)
                {
                    string G1 = dt2.Rows[0]["filename"].ToString().Replace(" ", "").ToUpper().Trim();
                    string BAU = add9TextBox.Text.Replace(" ", "").ToUpper().Trim();
                    int F1 = G1.IndexOf(BAU);
                    if (F1 == -1)
                    {

                        MessageBox.Show("檔案名稱重複,請修改檔名");
                        return;
                    }
                }

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
                    drw["path"] = PATH + de + filename;
                    if (globals.DBNAME == "進金生" || globals.DBNAME == "測試區98")
                    {


                        int T1 = add1TextBox.Text.IndexOf("正航系統");
                        if (T1 == -1)
                        {
                            if (shipping_ItemDataGridView.Rows.Count > 1)
                            {
                                DataGridViewRow row;
                                for (int i = 0; i <= shipping_ItemDataGridView.Rows.Count - 2; i++)
                                {
                                    row = shipping_ItemDataGridView.Rows[i];

                                    string ItemRemark = row.Cells["ItemRemark"].Value.ToString();
                                    string Docentry = row.Cells["Docentry"].Value.ToString();
                                    if (ItemRemark == "銷售訂單")
                                    {
                                        System.Data.DataTable G1 = GetMenu.GetSA(Docentry);
                                        if (G1.Rows.Count > 0)
                                        {
                                            drw["SA"] = G1.Rows[0]["業管"].ToString();
                                            drw["SALES"] = G1.Rows[0]["業務"].ToString();
                                        }
                                    }
                                    if (String.IsNullOrEmpty(drw["SA"].ToString()))
                                    {
                                        if (ItemRemark == "調撥單")
                                        {
                                            System.Data.DataTable G1 = GetSAOWTR(Docentry);
                                            if (G1.Rows.Count > 0)
                                            {
                                                string SALES = G1.Rows[0]["業務"].ToString();
                                                drw["SALES"] = SALES;
                                                System.Data.DataTable G2 = GetSA2();
                                                if (G2.Rows.Count > 0)
                                                {
                                                    drw["SA"] = G2.Rows[0][0].ToString();
                                                }
                                            }

                                        }
                                    }
                                }
                            }


                        }
                        else
                        {
                            if (dOCTYPETextBox.Text == "銷售訂單")
                            {

                                string STRN = "";
                                if (add1TextBox.Text == "正航系統CHOICE")
                                {
                                    STRN = strCn;
                                }
                                else if (add1TextBox.Text == "正航系統INFINITE")
                                {
                                    STRN = strCn22;
                                }
                                else if (add1TextBox.Text == "正航系統TOP GARDEN")
                                {
                                    STRN = strCn20;
                                }
                                System.Data.DataTable G2 = GetCHOF(pinoTextBox.Text.Trim(), STRN);
                                if (G2.Rows.Count > 0)
                                {

                                    string SA = G2.Rows[0]["SA"].ToString();
                                    string SALES = G2.Rows[0]["SALES"].ToString();

                                    System.Data.DataTable G21 = GetOSLP(SALES);
                                    if (G21.Rows.Count > 0)
                                    {
                                        drw["SALES"] = G21.Rows[0]["SALES"].ToString();
                                    }
                                    System.Data.DataTable G22 = GetOHEM(SA);
                                    if (G22.Rows.Count > 0)
                                    {
                                        drw["SA"] = G22.Rows[0]["SA"].ToString();
                                    }
                                }
                            }
                        }

                    }

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
                GetMenu.InsertLog(fmLogin.LoginID.ToString(), "可下載檔案上傳", "單號" + shippingCodeTextBox.Text + ex.Message, DateTime.Now.ToString("yyyyMMddHHmmss"));

                if (fmLogin.LoginID.ToString().ToUpper() != "EVAHSU")
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        public void UPOCLG(int atcentry, string DOCTYPE, string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE OCLG SET atcentry=@atcentry where DOCTYPE=@DOCTYPE AND DOCENTRY=@DOCENTRY ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@atcentry", atcentry));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
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
        private System.Data.DataTable GETODLNS(string U_WH_NO)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY DOC  FROM ODLN WHERE U_Shipping_no like '%" + U_WH_NO + "%' AND U_Shipping_no LIKE '%SH%' ");

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


        private System.Data.DataTable GetMAXOATC()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(ABSENTRY)+1 ID FROM OATC ");

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
        private System.Data.DataTable GetODLN(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CARDCODE,TEL1,ISNULL(T2.CNTCTCODE,0) CNTCTCODE,SLPCODE FROM ODLN T0 ");
            sb.Append(" LEFT JOIN OCPR T2 ON (T0.CNTCTCODE=T2.CNTCTCODE)");
            sb.Append(" WHERE T0.DOCENTRY=@DOC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));


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
        private System.Data.DataTable GetATC1(string ABSENTRY)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT max(LINE)+1 FROM ATC1 WHERE ABSENTRY=@ABSENTRY   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ABSENTRY", ABSENTRY));


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
        public void AddOACT(int AbsEntry)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into OATC(AbsEntry) values(@AbsEntry)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));

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
        private System.Data.DataTable GetATC1S(string ABSENTRY)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ABSENTRY FROM ATC1 WHERE ABSENTRY=@ABSENTRY   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ABSENTRY", ABSENTRY));


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
        private System.Data.DataTable GetMAXOCLG2(string DOCTYPE, string DOCENTRY)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM OCLG WHERE DOCTYPE=@DOCTYPE AND DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        public void AddATC1(int AbsEntry, int Line, string srcPath, string trgtPath, string FileName, string FileExt, DateTime Date, int UsrID, string Copied, string Override)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into ATC1(AbsEntry,Line,srcPath,trgtPath,FileName,FileExt,Date,UsrID,Copied,Override) values(@AbsEntry,@Line,@srcPath,@trgtPath,@FileName,@FileExt,@Date,@UsrID,@Copied,@Override)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));
            command.Parameters.Add(new SqlParameter("@Line", Line));
            command.Parameters.Add(new SqlParameter("@srcPath", srcPath));
            command.Parameters.Add(new SqlParameter("@trgtPath", trgtPath));
            command.Parameters.Add(new SqlParameter("@FileName", FileName));
            command.Parameters.Add(new SqlParameter("@FileExt", FileExt));
            command.Parameters.Add(new SqlParameter("@Date", Date));
            command.Parameters.Add(new SqlParameter("@UsrID", UsrID));
            command.Parameters.Add(new SqlParameter("@Copied", Copied));
            command.Parameters.Add(new SqlParameter("@Override", Override));
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
        public System.Data.DataTable GetCHOF(string DocEntry, string STRN)
        {
            SqlConnection connection = null;


            connection = new SqlConnection(STRN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" select  P.PersonName SALES,T0.Maker SA from ordBillMain T0   Left join comPerson P ON (T0.Salesman=P.PersonID)   ");
            sb.Append(" WHERE  T0.BillNO =@BillNO AND T0.FLAG=2");
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

        private void button12_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\Book3.xls";


            GetExcelProduct3(FileName, GetOrderData());
            dollarsKindTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        private System.Data.DataTable GetOrderData()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  select isnull(d.SoNo,'') SoNo,isnull(a.ShippingCode,'') ShippingCode,isnull(a.shipper,'') shipper,isnull(a.Consignee,'') Consignee,isnull(b.Cargo,'') TARE,isnull(a.NotifyPart,'') NotifyPart,");
            sb.Append(" isnull(d.receivePlace,'') receivePlace,isnull(a.OceanVessel,'') OceanVessel,isnull(a.Discharge,'') Discharge,isnull(a.Delivery,'') Delivery, ");
            sb.Append(" isnull(b.ContainerSeals,'') ContainerSeals,isnull(b.Packages,'') Packages,isnull(b.Description,'') Description,isnull(b.Measurement,'') Measurement,isnull(a.freightPaid,'') freightPaid,a.loading shipment,lANDTYPE BL  from ladingm a left join ladingd b on(a.shippingcode=b.shippingcode and a.seqmno=b.seqmno)");
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
            if (globals.DBNAME == "達睿生")
            {


                sb.Append(" SELECT a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,''''+a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期");
                sb.Append(" ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],b.[MarkNos],");
                if (GetINVMARK().Rows.Count == 0)
                {
                    sb.Append(" cast(seqno2+1 as varchar)+')'+b.[INDescription] as INDescription");
                }
                else
                {
                    sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(seqno2+1 as varchar)+')'+b.[INDescription] END INDescription ");
                }
                sb.Append(" ,b.[InQty] ,b.[UnitPrice],b.[Amount],c.brand +' BRAND' as BRAND,c.TradeCondition as Trade FROM [InvoiceM] as a");
                sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)");
                sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
                sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");
            }
            else
            {
                sb.Append(" SELECT a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,''''+a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期");
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
                sb.Append(" CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[UnitPrice] AS Money),2) UnitPrice");
                sb.Append(",CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[Amount] AS Money),2) Amount");
                sb.Append(" FROM [InvoiceM] as a ");
                sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)");
                sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode) ");
                sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");
            }



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

            SqlConnection connection = globals.shipConnection;

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

        public System.Data.DataTable GetSA2()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 (T3.[lastName]+T3.[firstName]) 業管   FROM ORDR T0  ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" INNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append(" WHERE    T0.CARDCODE=@CARDCODE ORDER BY DOCENTRY DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", cardCodeTextBox.Text));

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
        public System.Data.DataTable GetSAOWTR(string PINO)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (T2.[SlpName]) 業務 ");
            sb.Append(" FROM OWTR T0  ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" WHERE    CAST(T0.DOCENTRY AS VARCHAR)=@DOCENTRY");
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

        public System.Data.DataTable GetOSLP(string SALES)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SLPNAME SALES FROM OSLP WHERE SlpName like '%" + SALES + "%'  ");

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
        public System.Data.DataTable GetOHEM(string HOMETEL)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [lastName]+[firstName]  姓名 FROM OHEM WHERE HOMETEL=@HOMETEL ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", HOMETEL));
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
            sb.Append(" SELECT [PATH]   FROM DOWNLOAD  WHERE　(ISNULL(DLCHECK,'')='True') AND SHIPPINGCODE=@SHIPPINGCODE  ");

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
        public System.Data.DataTable GetDOWNLOADSAW()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            string ADD9 = add9TextBox.Text;
            sb.Append("  SELECT [PATH]   FROM DOWNLOAD WHERE　(ISNULL(DLCHECK2,'')='True') AND  SHIPPINGCODE=@SHIPPINGCODE");
            if (!String.IsNullOrEmpty(ADD9))
            {
                sb.Append("  UNION  ALL");
                sb.Append("   SELECT [PATH]   FROM DOWNLOAD2 WHERE　  [filename]　 like '%" + ADD9 + "%'   AND  SHIPPINGCODE=@SHIPPINGCODE");
            }
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

        public System.Data.DataTable GetDOWNLOADSAW2()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            string ADD9 = add9TextBox.Text;
            sb.Append("  SELECT [PATH]   FROM DOWNLOAD WHERE　(ISNULL(DLCHECK2,'')='True') AND  SHIPPINGCODE=@SHIPPINGCODE");

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
        public System.Data.DataTable GetDOWNLOADWH(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY 單號,ITEMCODE 產品編號,Quantity 數量 FROM Shipping_Item WHERE SHIPPINGCODE=@SHIPPINGCODE ");

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

        private System.Data.DataTable GetITEMINVOICE()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(QUANTITY),0) 數量,ISNULL(SUM(ITEMAMOUNT),0) 金額  FROM Shipping_Item WHERE SHIPPINGCODE =@SHIPPINGCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT ISNULL(SUM(INQTY),0), ISNULL(SUM(AMOUNT),0)   FROM InvoiceD  WHERE SHIPPINGCODE =@SHIPPINGCODE");

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
            sb.Append(" SELECT * FROM DOWNLOAD2 WHERE MARK='1' AND REPLACE([FILENAME],' ','') LIKE '%" + add9TextBox.Text.ToString().Replace(" ", "") + "%'  ");



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
        private System.Data.DataTable GetTT()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT B.VISORDER,a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,a.[PIno],a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期 ");
            sb.Append("               ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],b.[MarkNos], ");
            if (GetINVMARK().Rows.Count == 0)
            {
                sb.Append(" cast(seqno2+1 as varchar)+')'+b.[INDescription] as INDescription");
            }
            else
            {
                sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(seqno2+1 as varchar)+')'+b.[INDescription] END INDescription ");
            }
            sb.Append("               ,cast(b.[InQty] as varchar) InQty,CAST(b.[UnitPrice] AS VARCHAR) UnitPrice,CAST(b.[Amount] AS VARCHAR) Amount,c.brand +' BRAND' as BRAND,c.TradeCondition as Trade FROM [InvoiceM] as a ");
            sb.Append("               left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq) ");
            sb.Append("               left join shipping_main as c on (a.shippingcode=c.shippingcode)  ");
            sb.Append("          WHERE  a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq   ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT B.VISORDER,'','','','','','','','','',''");
            sb.Append("               ,'','','','','','   '+KIT+'*'+QTY+'PC' ");
            sb.Append("               ,'','','','','' FROM INVOICEDKIT A");
            sb.Append(" LEFT JOIN [InvoiceD] B ON (A.shippingcode=B.shippingcode AND A.InvoiceNo=B.InvoiceNo AND A.InvoiceNo_seq=B.InvoiceNo_seq AND B.TREETYPE='S' AND A.ITEMNAME=B.INDescription)");
            sb.Append(" WHERE  a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq   ");
            //sb.Append(" ORDER BY VISORDER,a.[PIno] DESC");





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
        private System.Data.DataTable GetObuInvo()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,a.[PIno],a.[POno] as pono,'BILL TO:'+a.[obubillTo] as billTo,'SHIP TO:'+a.[obushipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期");
            sb.Append(" ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],");
            if (GetINVMARK().Rows.Count == 0)
            {
                sb.Append(" cast(seqno2+1 as varchar)+')'+b.[INDescription] as INDescription");
            }
            else
            {
                sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(seqno2+1 as varchar)+')'+b.[INDescription] END INDescription ");
            }

            sb.Append(" ,b.[InQty] ,b.[CHOPrice]  UnitPrice,b.[CHOAmount] Amount,c.brand +' BRAND' as BRAND,c.TradeCondition as Trade FROM [InvoiceM] as a");
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
        private System.Data.DataTable GetOBUPack()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            System.Data.DataTable T1 = GetINVPACK();
            sb.Append(" SELECT a.SayTotal+' PLTS' as cc,a.[PLNo]  invoiceNo,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[OBUShipTo] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計'");
            sb.Append(" ,a.[Net] as '耐特',a.[Gross] as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[OBUBillTo] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]");
            sb.Append(" ,a.[Quantity] as '總數',a.[Net],a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo],");
            if (T1.Rows.Count == 0)
            {
                sb.Append("  CAST(seqno2+1 AS VARCHAR)+')'+ b.[DescGoods]   DescGoods ");
            }
            else
            {

                sb.Append(" CASE WHEN ISNULL(PACKMARK,'') <> 'True' THEN b.[DescGoods]  ELSE cast(seqno2+1 as varchar)+')'+b.[DescGoods] END DescGoods ");
            }

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
        private System.Data.DataTable GetOrderData3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT CASE WHEN a.SayTotal='0' THEN CASE WHEN ISNULL(A.SAYCTN,'0')='0' THEN  CAST(T0.CNO AS VARCHAR) +'CTNS' ELSE A.SAYCTN +'CTNS' END ELSE  a.SayTotal+' PLTS' END as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計' ");
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
            sb.Append(" 				LEFT JOIN (SELECT ShippingCode,PLNo,MAX(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,3) END AS INT))  CNO  FROM PackingListD  GROUP BY ShippingCode,PLNo) T0 ON (T0.ShippingCode=A.ShippingCode and T0.PLNo=A.PLNo)");
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
        private System.Data.DataTable GetOrderData3DRS(string CARD, string ADD, string TEL)
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


        private System.Data.DataTable GetOrderData3BOM()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT B.VISORDER,a.SayTotal+' PLTS' as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計'    ");
            sb.Append("                                                       ,'' '耐特','' '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]    ");
            sb.Append("                                                       ,'' as '總數','' Net,'' Gross,a.[SayTotal],'' PackageNo,'' CNo,CAST(B.seqno+1 AS VARCHAR)+')'+b.[DescGoods]+ ' Attachment List'  as DescGoods    ");
            sb.Append("                                                       ,'' Quantity ,''  Ne ,''  Go ,'' MeasurmentCM,1 A FROM [PackingListM] as a    ");
            sb.Append("                                                       left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)    ");
            sb.Append("                                         where a.shippingcode=@shippingcode and a.PLNo=@PLNo AND TREETYPE='S' ");
            sb.Append(" UNION ALL");
            sb.Append("                         SELECT B.VISORDER,a.SayTotal+' PLTS' as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計'    ");
            sb.Append("                                                       ,cast(D.[Net] as varchar) as '耐特',cast(D.[Gross] as varchaR) as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]    ");
            sb.Append("                                                       ,CAST(D.[QTY] AS VARCHAR) as '總數',CAST(d.[Net] AS VARCHAR) Net,CAST(d.[Gross] AS VARCHAR) Gross,a.[SayTotal],b.[PackageNo],D.[CNo],E.ENGLISH+')'+D.[KIT] as DescGoods    ");
            sb.Append("                                                       ,CAST(D.[QTY] AS VARCHAR)  Quantity ,D.[Net]  Ne ,D.[Gross]  Go ,'' MeasurmentCM,2 A  FROM [PackingListM] as a    ");
            sb.Append("                                                       left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)    ");
            sb.Append("                               left join  [PackingListDKIT] as D on (B.ShippingCode=D.ShippingCode and B.PLNo=D.PLNo AND B.DescGoods=D.ITEMNAME)    ");
            sb.Append("                                                       left join shipping_main as c on (a.shippingcode=c.shippingcode)     ");
            sb.Append("                  left join  [PackingListDKITENGLISH] as E on (D.SEQNO=E.LINENUM)   ");
            sb.Append("                                         where a.shippingcode=@shippingcode and a.PLNo=@PLNo AND TREETYPE='S' ");
            sb.Append(" ORDER BY VISORDER,A ");
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


        private void GetExcelProduct(string ExcelFile, System.Data.DataTable dt, string FLAG, string FLAG2)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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
                    string OHEM = fmLogin.LoginID.ToString().ToUpper();
                    if (OHEM == "EVAHSU")
                    {
                        excelSheet.Shapes.AddPicture(B2 + "EVAHSU.JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 350, 640, 200, 80);
                    }
                    else
                    {
                        excelSheet.Shapes.AddPicture(B2 + createNameTextBox.Text.Trim().ToUpper() + ".JPG", Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue, 350, 650, 200, 80);
                    }
                }
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

                if (dOCTYPETextBox.Text == "銷售訂單")
                {
                    if (boardCountNoTextBox.Text == "三角" || boardCountNoTextBox.Text == "出口")
                    {
                        if (mEMO3TextBox.Text != "")
                        {
                            StringBuilder sbs = new StringBuilder();
                            StringBuilder sbs2 = new StringBuilder();
                            StringBuilder sbs3 = new StringBuilder();
                            string MAT = "";
                            string MAT2 = "";
                            string MAT3 = "";
                            string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                            StringBuilder sb = new StringBuilder();
                            foreach (string i in arrurl)
                            {
                                sb.Append("'" + i + "',");
                            }
                            sb.Remove(sb.Length - 1, 1);

                            System.Data.DataTable dt3 = GetWHPACK2(sb.ToString());


                            if (dt3.Rows.Count == 2)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 31, 1]);
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                                for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                                {
                                    string MATERIAL = dt3.Rows[i]["MATERIAL"].ToString();
                                    System.Data.DataTable dt5 = GetWHPACK5(sb.ToString(), MATERIAL);
                                    if (dt5.Rows.Count > 0)
                                    {
                                        for (int s = 0; s <= dt5.Rows.Count - 1; s++)
                                        {
                                            string PLATENO = dt5.Rows[s]["PLATENO"].ToString();

                                            if (i == 0)
                                            {
                                                sbs.Append(PLATENO + " &");
                                            }

                                            if (i == 1)
                                            {
                                                sbs2.Append(PLATENO + " &");
                                            }

                                        }
                                        //塑料棧板
                                        int H1 = MATERIAL.ToUpper().IndexOf("IPPC");
                                        int H2 = MATERIAL.IndexOf("料");
                                        int H4 = MATERIAL.IndexOf("卡");
                                        int H5 = MATERIAL.IndexOf("塑膠板");
                                        int H6 = MATERIAL.IndexOf("塑料棧板");
                                        int H3 = MATERIAL.IndexOf("合");
                                        if (H1 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": IPPC STAMPED WOODEN PALLETS";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": IPPC STAMPED WOODEN PALLETS";
                                            }
                                        }
                                        if (H2 != -1 || H4 != -1 || H5 != -1 || H6 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": NON-WOODEN PACKAGING MATERIAL";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": NON-WOODEN PACKAGING MATERIAL";
                                            }
                                        }
                                        if (H3 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": PLYWOOD PALLET FOR PACKAGING MATERIAL";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": PLYWOOD PALLET FOR PACKAGING MATERIAL";
                                            }
                                        }
                                    }
                                }
                                if (sbs.Length != 0)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 31, 2]);
                                    range.Value2 = "(PALLET NO. " + sbs.Remove(sbs.Length - 1, 1) + MAT + ")";
                                    range.Font.Size = 10;
                                    range.Font.Name = "Arial";
                                }

                                if (sbs2.Length != 0)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 32, 2]);
                                    range.Value2 = "(PALLET NO. " + sbs2.Remove(sbs2.Length - 1, 1) + MAT2 + ")";
                                    range.Font.Size = 10;
                                    range.Font.Name = "Arial";
                                }
                            }


                            if (dt3.Rows.Count == 3)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 31, 1]);
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 31, 1]);
                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                                for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                                {
                                    string MATERIAL = dt3.Rows[i]["MATERIAL"].ToString();
                                    System.Data.DataTable dt5 = GetWHPACK5(sb.ToString(), MATERIAL);
                                    if (dt5.Rows.Count > 0)
                                    {
                                        for (int s = 0; s <= dt5.Rows.Count - 1; s++)
                                        {
                                            string PLATENO = dt5.Rows[s]["PLATENO"].ToString();

                                            if (i == 0)
                                            {
                                                sbs.Append(PLATENO + " &");
                                            }

                                            if (i == 1)
                                            {
                                                sbs2.Append(PLATENO + " &");
                                            }

                                            if (i == 2)
                                            {
                                                sbs3.Append(PLATENO + " &");
                                            }
                                        }
                                        //塑料棧板
                                        int H1 = MATERIAL.ToUpper().IndexOf("IPPC");
                                        int H2 = MATERIAL.IndexOf("料");
                                        int H4 = MATERIAL.IndexOf("卡");
                                        int H5 = MATERIAL.IndexOf("塑膠板");
                                        int H6 = MATERIAL.IndexOf("塑料棧板");
                                        int H3 = MATERIAL.IndexOf("合");
                                        if (H1 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": IPPC STAMPED WOODEN PALLETS";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": IPPC STAMPED WOODEN PALLETS";
                                            }
                                            if (i == 2)
                                            {
                                                MAT3 = ": IPPC STAMPED WOODEN PALLETS";
                                            }
                                        }
                                        if (H2 != -1 || H4 != -1 || H5 != -1 || H6 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": NON-WOODEN PACKAGING MATERIAL";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": NON-WOODEN PACKAGING MATERIAL";
                                            }
                                            if (i == 2)
                                            {
                                                MAT3 = ": NON-WOODEN PACKAGING MATERIAL";
                                            }
                                        }
                                        if (H3 != -1)
                                        {
                                            if (i == 0)
                                            {
                                                MAT = ": PLYWOOD PALLET FOR PACKAGING MATERIAL";
                                            }
                                            if (i == 1)
                                            {
                                                MAT2 = ": PLYWOOD PALLET FOR PACKAGING MATERIAL";
                                            }
                                            if (i == 2)
                                            {
                                                MAT3 = ": PLYWOOD PALLET FOR PACKAGING MATERIAL";
                                            }
                                        }
                                    }
                                }
                                if (sbs.Length != 0)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 31, 2]);
                                    range.Value2 = "(PALLET NO. " + sbs.Remove(sbs.Length - 1, 1) + MAT + ")";
                                    range.Font.Size = 10;
                                    range.Font.Name = "Arial";
                                }

                                if (sbs2.Length != 0)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 32, 2]);
                                    range.Value2 = "(PALLET NO. " + sbs2.Remove(sbs2.Length - 1, 1) + MAT2 + ")";
                                    range.Font.Size = 10;
                                    range.Font.Name = "Arial";
                                }

                                if (sbs3.Length != 0)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 33, 2]);
                                    range.Value2 = "(PALLET NO. " + sbs3.Remove(sbs3.Length - 1, 1) + MAT3 + ")";
                                    range.Font.Size = 10;
                                    range.Font.Name = "Arial";
                                }
                            }

                        }


                    }
                }



                System.Data.DataTable dt4 = GetWHITEM(shippingCodeTextBox.Text, pLNoTextBox.Text);
                if (dt4.Rows.Count > 0)
                {

                    //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 30, 1]);
                    //range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);
                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29 + i, 3]);
                        range.Value2 = dt4.Rows[i][0].ToString();


                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    }

                    int n;
                    string A1 = "";
                    string A2 = "";
                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {
                        if (i == 0)
                        {

                            A1 = dt4.Rows[i][0].ToString();
                            int F1 = A1.IndexOf("~");
                            int F2 = A1.IndexOf(")");
                            if (F1 != -1 && F2 != -1)
                            {
                                A1 = A1.Substring(F1 + 1, F2 - F1 - 1);
                            }
                        }

                        if (i == 1)
                        {
                            A2 = dt4.Rows[i][0].ToString();

                            int F1 = A2.IndexOf("ITEM");
                            int F2 = A2.IndexOf("~");
                            if (F1 != -1 && F2 != -1)
                            {
                                A2 = A2.Substring(F1 + 4, F2 - F1 - 4);
                            }
                        }

                    }

                    if (int.TryParse(A1, out n) && int.TryParse(A2, out n))
                    {

                        int D1 = Convert.ToInt16(A1);
                        int D2 = Convert.ToInt16(A2);

                        if (D1 > D2)
                        {
                            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                            {

                                string LOCATION = dt4.Rows[i]["LOCATION"].ToString();
                                System.Data.DataTable dt4P = GetWHITEMP(shippingCodeTextBox.Text, pLNoTextBox.Text, LOCATION);
                                if (dt4P.Rows.Count > 0)
                                {

                                    //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 30, 1]);
                                    //range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);
                                    StringBuilder sbss = new StringBuilder();
                                    for (int i2 = 0; i2 <= dt4P.Rows.Count - 1; i2++)
                                    {
                                        sbss.Append("" + dt4P.Rows[i2][0].ToString() + ".");
                                    }


                                    sbss.Remove(sbss.Length - 1, 1);

                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + 29 + i, 3]);
                                    range.Value2 = "ITEM" + sbss.ToString() + ")MADE IN " + LOCATION;

                                }

                            }
                        }
                    }


                }

                if (FLAG2 == "Y")
                {
                    System.Data.DataTable dtF = GetCTN(shippingCodeTextBox.Text, pLNoTextBox.Text);
                    if (dtF.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dtF.Rows.Count - 1; i++)
                        {
                            string PACKAGENO = dtF.Rows[i][0].ToString();
                            int SEQNO = Convert.ToInt16(dtF.Rows[i][1]);
                            System.Data.DataTable dtF2 = GetCTN2(shippingCodeTextBox.Text, pLNoTextBox.Text, SEQNO);
                            int SEQNO2 = Convert.ToInt16(dtF2.Rows[0][0]) - 1;
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[SEQNO2 + 26 + i, 1]);
                            // range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[SEQNO2 + 26 + i, 3]);
                            range.Select();
                            range.Value2 = "(" + PACKAGENO + "PCS/ CTN)";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.VerticalAlignment = XlVAlign.xlVAlignBottom;
                            range.Font.Bold = true;
                            range.Font.Size = 8;
                        }
                    }
                }
                else
                {
                    System.Data.DataTable dtF = GetCTN(shippingCodeTextBox.Text, pLNoTextBox.Text);
                    if (dtF.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dtF.Rows.Count - 1; i++)
                        {
                            string PACKAGENO = dtF.Rows[i][0].ToString();
                            int SEQNO = Convert.ToInt16(dtF.Rows[i][1]);
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[SEQNO + 26 + i, 1]);
                            // range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[SEQNO + 26 + i, 3]);
                            range.Select();
                            range.Value2 = "(" + PACKAGENO + "PCS/ CTN)";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.VerticalAlignment = XlVAlign.xlVAlignBottom;
                            range.Font.Bold = true;
                            range.Font.Size = 8;
                        }
                    }
                }

            }
            finally
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\Excel\\temp\\" +
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

        public System.Data.DataTable GetCTN2(string SHIPPINGCODE, string PLNO, int SEQNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT COUNT(*) A FROM (SELECT CAST(B.seqno AS decimal(10,1)) S FROM [PackingListM] as a  ");
            sb.Append(" left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)  ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)  ");
            sb.Append(" LEFT JOIN (SELECT ShippingCode,PLNo,MAX(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,3) END AS INT))  CNO  FROM PackingListD  GROUP BY ShippingCode,PLNo) T0 ON (T0.ShippingCode=A.ShippingCode and T0.PLNo=A.PLNo) ");
            sb.Append(" where a.shippingcode=@SHIPPINGCODE      AND a.PLNO=@PLNO     ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT  CAST(B.seqno+'.1' AS decimal(10,1) )  FROM [PackingListM] as a  ");
            sb.Append(" left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)  ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)  ");
            sb.Append("                   LEFT JOIN INVOICED D ON (D.INDESCRIPTION=B.DESCGOODS AND D.SHIPPINGCODE=B.SHIPPINGCODE) ");
            sb.Append("                   LEFT JOIN ACMESQL02.DBO.OSCN E ON (D.ITEMCODE =E.ItemCode  COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" where a.shippingcode=@SHIPPINGCODE  AND a.PLNO=@PLNO AND ISNULL(b.[CNo],'') <>'') SS");
            sb.Append(" WHERE S<= @SEQNO");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));
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
        private void GetExcelProductBOM(string ExcelFile, System.Data.DataTable dt)
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

        private void GetExcelProduct2(string ExcelFile, System.Data.DataTable dt, string FLAG)
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
                            DetailRow1 = 24;
                            break;
                        }


                    }

                }

                string PINO = dt.Rows[0]["PIno"].ToString().Trim('\'');   //子料號
                System.Data.DataTable dtPINO = null;
                if (PINO.Length == 12)
                {

                    dtPINO = checkCn21(PINO);
                }
                else
                {
                    dtPINO = checkPartNumber(PINO);
                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= dt.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dt.Rows.Count - 1 || dtPINO.Rows.Count > 0)
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


                        }

                        DetailRow++;
                    }

                }


                if (dtPINO.Rows.Count > 0)
                {
                    for (int aRow = 0; aRow <= dtPINO.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != dtPINO.Rows.Count - 1)
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
                            SetRow(aRow, sTemp, ref FieldValue, dtPINO);


                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }


                System.Data.DataTable dt4 = GetWHITEM2(shippingCodeTextBox.Text);
                if (dt4.Rows.Count == 0)
                {
                    System.Data.DataTable dt4S = GetWHITEM2S(shippingCodeTextBox.Text);

                    Clear(sbS);
                    SBS();


                    if (dt4S.Rows.Count > 0)
                    {

                        for (int i = 0; i <= dt4S.Rows.Count - 1; i++)
                        {
                            string ITEMCODE = dt4S.Rows[i]["ITEMCODE"].ToString();
                            string DOCENTRY = dt4S.Rows[i]["DOCENTRY"].ToString();
                            System.Data.DataTable dtF = GetWHLOCATION(sbS.ToString(), ITEMCODE);
                            if (dtF.Rows.Count > 0)
                            {
                                string LOCATION = dtF.Rows[0][0].ToString();
                                UPDATEINVOICED(LOCATION, DOCENTRY);

                            }
                        }
                    }

                }

                System.Data.DataTable dt5 = GetWHITEM2(shippingCodeTextBox.Text);
                if (dt5.Rows.Count > 0)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + dtPINO.Rows.Count + 26, 1]);
                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, oMissing);
                    for (int i = 0; i <= dt5.Rows.Count - 1; i++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[dt.Rows.Count + dtPINO.Rows.Count + 27 + i, 2]);
                        range.Value2 = dt5.Rows[i][0].ToString();
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
        private void GetObuInvoExcel(string ExcelFile, System.Data.DataTable dt)
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
                            SetRow(aRow, sTemp, ref FieldValue, dt);

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

        private void GetExcelProduct3(string ExcelFile, System.Data.DataTable dt)
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

                            //int H1 = FieldValue.IndexOf("=");
                            //FieldValue = FieldValue.Replace("=", "");
                            //if (H1 != -1)
                            //{
                            //    range.Value2 = " =" + FieldValue.ToString();
                            //}
                            //else
                            //{
                            //    range.Value2 = FieldValue.ToString();
                            //}

                            range.Value2 = FieldValue.ToString();
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
        private void GetExcelinsu(string ExcelFile, System.Data.DataTable dt)
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
                            SetRow(aRow, sTemp, ref FieldValue, dt);

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
                        SetRow(aRow, sTemp, ref FieldValue, dt);

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
        private void SetRow(int iRow, string sData, ref string FieldValue, System.Data.DataTable dt)
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

            CalcTotals1();

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            if (comboBox1.Text == "")
            {
                if (globals.DBNAME == "達睿生")
                {
                    System.Data.DataTable G1 = GetshipTYPE();
                    if (G1.Rows.Count > 0)
                    {

                        int T2 = cardNameTextBox.Text.IndexOf("進金生");
                        int T3 = cardNameTextBox.Text.IndexOf("进金生");
                        int T4 = cardNameTextBox.Text.IndexOf("友達");
                        string CARD = "";
                        string ADD = "";
                        string TEL = "";
                        if (T2 != -1 || T3 != -1)
                        {
                            FileName = lsAppDir + "\\Excel\\DRS\\INVODRSAP2.xls";
                        }
                        else if (T4 != -1)
                        {

                            FileName = lsAppDir + "\\Excel\\DRS\\INVODRSAP.xls";
                        }
                        else
                        {
                            FileName = lsAppDir + "\\Excel\\DRS\\INVODRSAP3.xls";
                            System.Data.DataTable K1 = GetOrderData2DRS2();
                            if (K1.Rows.Count > 0)
                            {
                                CARD = K1.Rows[0]["CARD"].ToString();
                                ADD = K1.Rows[0]["ADD"].ToString();
                                TEL = K1.Rows[0]["TEL"].ToString();
                            }

                        }

                        GetExcelProduct2(FileName, GetOrderData2DRS(CARD, ADD, TEL), "N");
                    }
                    else
                    {

                        FileName = lsAppDir + "\\Excel\\DRS\\INVODRS.xls";
                        GetExcelProduct2(FileName, GetOrderData2(), "Y");
                    }

                }

                else
                {
                    string OHEM = fmLogin.LoginID.ToString().ToUpper();

                    if (DRS == "DRS")
                    {
                        FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACMEDRS.xls";
                        GetExcelProduct2(FileName, GetOrderData2(), "N");
                    }
                    else if (OHEM == "NANCYTSAI" || OHEM == "EVAHSU" || OHEM == "BETTYTSENG")
                    {
                        FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACME.xls";
                        GetExcelProduct2(FileName, GetOrderData2(), "Y");
                    }
                    else
                    {

                        FileName = lsAppDir + "\\Excel\\INVO2.xls";
                        GetExcelProduct2(FileName, GetOrderData2(), "N");
                    }
                }


            }
            else if (comboBox1.Text == "UPS")
            {
                FileName = lsAppDir + "\\Excel\\INVO.xls";
                GetExcelProduct2(FileName, GetOrderData2(), "N");
            }
            else if (comboBox1.Text == "HSBC")
            {
                FileName = lsAppDir + "\\Excel\\INVO3.xls";
                GetExcelProduct2(FileName, GetOrderData2(), "N");
            }
            else if (comboBox1.Text == "BOM")
            {
                FileName = lsAppDir + "\\Excel\\INVO2.xls";
                GetExcelProduct2(FileName, GetOrderData2(), "N");
            }


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


            if (dt1.Rows.Count <= 0)
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

                if ((Checked != "True" && shippingcode != "1" && f2 == 1) || (RED == "True"))
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
                if ((Checked != "True" && shippingcode != "1" && f2 == 1) || (RED == "True"))
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
                System.Data.DataTable dtPI1 = GetMenu.GetTIFF2(shippingCodeTextBox.Text);


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

                if (globals.DBNAME == "達睿生")
                {
                    FileName = lsAppDir + "\\MailTemplates\\進口DRS.htm";
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
                        if (add10CheckBox.Checked || rUSHCheckBox.Checked)
                        {
                            SICHECK();
                            template = template.Replace("##anpei5##", DF3F.ToString());
                        }
                        else
                        {
                            template = template.Replace("##anpei5##", "");
                        }
                        template = template.Replace("##anpei6##", "");
                    }
                    else
                    {
                        template = template.Replace("##anpei2##", "請安排以下出貨，");

                        if (add10CheckBox.Checked || rUSHCheckBox.Checked)
                        {
                            SICHECK();
                            template = template.Replace("##anpei5##", DF3F.ToString());
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
                        if (add10CheckBox.Checked || rUSHCheckBox.Checked)
                        {
                            SICHECK();
                            template = template.Replace("##anpei5##", DF3F.ToString());
                        }
                        else
                        {
                            template = template.Replace("##anpei5##", "");
                        }
                        template = template.Replace("##anpei6##", "");
                    }
                    else
                    {

                        template = template.Replace("##anpei2##", "請安排以下出貨，");
                        if (add10CheckBox.Checked || rUSHCheckBox.Checked)
                        {

                            SICHECK();
                            template = template.Replace("##anpei5##", DF3F.ToString());

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
                template = template.Replace("##TEL##", "02-8791-2868");
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


                StringBuilder DF3 = new StringBuilder();
                if (add10CheckBox.Checked || iTEMSCheckBox.Checked || rUSHCheckBox.Checked)
                {
                    if (rUSHCheckBox.Checked)
                    {
                        DF3.Append("急貨+");
                    }
                    if (add10CheckBox.Checked)
                    {
                        DF3.Append("本票請申請 AUO貨代免倉期10天+");
                    }
                    if (iTEMSCheckBox.Checked)
                    {
                        DF3.Append("已確認小料號");
                    }
                }
                string RED = "";
                if (f2 == 1 || f3 == 1)
                {
                    RED = "(REV#紅字處)";
                }

                string SUBJECT = (RED + DF3.ToString() + DF2 + df).Replace("\r", "").Replace("\n", "");
                message.Subject = SUBJECT;
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
            if (rUSHTextBox.Text != "Checked")
            {
                rUSHCheckBox.Checked = false;
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
                rUSHCheckBox.Checked = false;
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


                if (boardCountNoTextBox.Text == "三角" || boardCountNoTextBox.Text == "出口" || boardCountNoTextBox.Text == "內銷")
                {
                    int J = markDataGridView.Rows.Count;
                    int K1 = 0;
                    System.Data.DataTable dth = ship.Mark;
                    //row1
                    DataRow drw2 = dth.NewRow();
                    K1 = J;
                    if (K1 < 10)
                    {
                        drw2["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw2["Seq"] = K1.ToString();
                    }
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["Mark"] = shippingCodeTextBox.Text;
                    dth.Rows.Add(drw2);

                    //row2
                    DataRow drw3 = dth.NewRow();
                    K1 = J + 1;
                    if (K1 < 10)
                    {
                        drw3["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw3["Seq"] = K1.ToString();
                    }
                    drw3["ShippingCode"] = shippingCodeTextBox.Text;

                    int T1 = receivePlaceTextBox.Text.ToUpper().IndexOf("TAIWAN");

                    int T2 = receivePlaceTextBox.Text.ToUpper().IndexOf("CHINA");
                    if (globals.DBNAME == "達睿生")
                    {
                        if (T1 != -1)
                        {
                            drw3["Mark"] = "DRS-T";
                        }
                        else if (T2 != -1)
                        {
                            drw3["Mark"] = "DRS-C";
                        }
                        else
                        {
                            drw3["Mark"] = "DRS";
                        }
                    }
                    else
                    {
                        if (T1 != -1)
                        {
                            drw3["Mark"] = "ACME-T";
                        }
                        else if (T2 != -1)
                        {
                            drw3["Mark"] = "ACME-C";
                        }
                        else
                        {
                            drw3["Mark"] = "ACME";
                        }
                    }
                    dth.Rows.Add(drw3);

                    //row3
                    DataRow drw4 = dth.NewRow();
                    K1 = J + 2;
                    if (K1 < 10)
                    {
                        drw4["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw4["Seq"] = K1.ToString();
                    }
                    drw4["ShippingCode"] = shippingCodeTextBox.Text;
                    drw4["Mark"] = "P/L No.";
                    dth.Rows.Add(drw4);

                    //row4
                    DataRow drw5 = dth.NewRow();
                    K1 = J + 3;
                    if (K1 < 10)
                    {
                        drw5["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw5["Seq"] = K1.ToString();
                    }
                    drw5["ShippingCode"] = shippingCodeTextBox.Text;
                    drw5["Mark"] = "產地: MADE IN XXXXX";
                    dth.Rows.Add(drw5);
                }


                if (boardCountNoTextBox.Text == "進口")
                {
                    int J = markDataGridView.Rows.Count;
                    int K1 = 0;

                    System.Data.DataTable dth = ship.Mark;
                    //row1
                    DataRow drw2 = dth.NewRow();
                    K1 = J;
                    if (K1 < 10)
                    {
                        drw2["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw2["Seq"] = K1.ToString();
                    }
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["Mark"] = shippingCodeTextBox.Text;
                    dth.Rows.Add(drw2);


                    //row4
                    DataRow drw5 = dth.NewRow();
                    K1 = J + 1;
                    if (K1 < 10)
                    {
                        drw5["Seq"] = "0" + K1.ToString();
                    }
                    else
                    {
                        drw5["Seq"] = K1.ToString();
                    }
                    drw5["ShippingCode"] = shippingCodeTextBox.Text;
                    drw5["Mark"] = "產地: MADE IN XXXXX";
                    dth.Rows.Add(drw5);
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

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                CalcTotals1();
                //CalcTotals1C();

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                if (cardCodeTextBox.Text == "0257-00" || add1TextBox.Text == "正航系統TOP GARDEN" || cardCodeTextBox.Text == "S0050")
                {
                    FileName = lsAppDir + "\\Excel\\GARDENINVO.xls";

                }
                else if (cardCodeTextBox.Text == "0511-00" || add1TextBox.Text == "正航系統CHOICE")
                {

                    if (add2TextBox.Text == "0355-01")
                    {
                        FileName = lsAppDir + "\\Excel\\CHOICEINVOMEGA.xls";
                    }
                    else
                    {
                        FileName = lsAppDir + "\\Excel\\CHOICEINVO.xls";
                    }
                }
                else if (cardCodeTextBox.Text == "1349-00" || add1TextBox.Text == "正航系統INFINITE")
                {
                    FileName = lsAppDir + "\\Excel\\INFINITEINVO.xls";
                }

                GetExcelProduct2(FileName, GetObuInvo(), "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                if (cardCodeTextBox.Text == "0257-00" || add1TextBox.Text == "正航系統TOP GARDEN" || cardCodeTextBox.Text == "S0050")
                {
                    FileName = lsAppDir + "\\Excel\\GARDENPACK.xls";

                }
                else if (cardCodeTextBox.Text == "0511-00" || add1TextBox.Text == "正航系統CHOICE")
                {

                    FileName = lsAppDir + "\\Excel\\CHOICEPACK.xls";
                }
                else if (cardCodeTextBox.Text == "1349-00" || add1TextBox.Text == "正航系統INFINITE")
                {
                    FileName = lsAppDir + "\\Excel\\INFINITEPACK.xls";
                }

                GetExcelProduct(FileName, GetOBUPack(), "N", "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }






        private void button22_Click(object sender, EventArgs e)
        {
            if (add2TextBox.Text == "")
            {
                MessageBox.Show("請輸入OBU客戶資訊");
            }
            else
            {


                System.Data.DataTable dt1 = GetMenu.Getaddress2(add2TextBox.Text);
                DataRow drw = dt1.Rows[0];
                oBUBillToTextBox.Text = drw["公司全稱"].ToString() + "\r\n" + drw["地址"].ToString() + "\r\n" + "TEL:" + drw["電話"].ToString() + "\r\n" + "FAX:" + drw["傳真"].ToString() + "\r\n" + "ATTN:" + drw["大名"].ToString();
                oBUShipToTextBox.Text = shippedByTextBox.Text;
            }
        }

        private void button23_Click(object sender, EventArgs e, string TYPE)
        {
            try
            {
                object[] LookupValues = null;

                if (TYPE == "發貨")
                {
                    LookupValues = GetMenu.GetMenuOg();
                }

                if (TYPE == "收貨")
                {
                    LookupValues = GetMenu.GetMenuOgN();
                }

                if (LookupValues != null)
                {
                    tabControl1.SelectedIndex = 0;

                    string pino = pinoTextBox.Text;
                    pinoTextBox.Text = Convert.ToString(LookupValues[0]);
                    string docentry = pinoTextBox.Text;

                    System.Data.DataTable dt1 = null;

                    if (TYPE == "發貨")
                    {
                        dt1 = GetMenu.GetOige(docentry);
                    }

                    if (TYPE == "收貨")
                    {
                        dt1 = GetMenu.GetOigN(docentry);
                    }

                    System.Data.DataTable dt2 = ship.Shipping_Item;


                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["seqNo"] = "0";
                        drw2["Docentry"] = drw["Docnum"];
                        drw2["ItemCode"] = drw["ItemCode"];
                        drw2["linenum"] = drw["linenum"];
                        drw2["Dscription"] = drw["Dscription"];
                        drw2["CURRENCY"] = drw["DOCCUR"];
                        int iQuantity = 0;
                        int iUnitPrice = 0;
                        iQuantity = Convert.ToInt32(drw["Quantity"]);
                        if (drw["delivrdqty"].ToString() != "")
                        {
                            iUnitPrice = Convert.ToInt32(drw["delivrdqty"]);
                        }
                        else
                        {
                            iUnitPrice = 0;
                        }
                        drw2["Quantity"] = (iQuantity - iUnitPrice).ToString();
                        drw2["ItemPrice"] = drw["Price"];

                        drw2["ItemAmount"] = drw["linetotal"];
                        drw2["ItemRemark"] = TYPE + "單";
                        drw2["Remark"] = drw["comments"];
                        drw2["VISORDER"] = drw["VISORDER"];
                        dt2.Rows.Add(drw2);

                    }
                    for (int j = 0; j <= shipping_ItemDataGridView.Rows.Count - 2; j++)
                    {
                        shipping_ItemDataGridView.Rows[j].Cells[1].Value = j.ToString();
                    }
                }
                shipping_mainBindingSource.EndEdit();
                shipping_ItemBindingSource.EndEdit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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





        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\移倉PACK.xls";
                GetExcelProduct(FileName, GetOrderData3(), "N", "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\移倉invo.xls";

                GetExcelProduct2(FileName, GetOrderData2(), "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }



        private void button24_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataTable2DataGridView);
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
        private void InsertAA(string SHIPNO, string SHIPNO2)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO AA (SHIPNO,SHIPNO2) VALUES(@SHIPNO,@SHIPNO2)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPNO", SHIPNO));
            command.Parameters.Add(new SqlParameter("@SHIPNO2", SHIPNO2));



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
        private void InsertDownload2(string ShippingCode, string seq, string filename, string path)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO download2 (ShippingCode,seq,filename,path,MARK,STATUS) VALUES(@ShippingCode,@seq,@filename,@path,@MARK,@STATUS)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@MARK", "1"));
            command.Parameters.Add(new SqlParameter("@STATUS", ""));


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
        private void DELETEDownload2(string ShippingCode, string filename)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE download2 where ShippingCode=@ShippingCode and filename=@filename ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@filename", filename));


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
        private void UPDATEADD9(string ADD9, string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE SHIPPING_MAIN SET ADD9 =@ADD9 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ADD9", ADD9));
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

        private void UPDATEINVOICED(string LOCATION, string DOCENTRY)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE INVOICED SET LOCATION =@LOCATION WHERE  DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
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
        private void InsertPacking(string ShippingCode, string PLNo, string SeqNo, string PackageNo, string CNo, string DescGoods, string Quantity, string Net, string Gross, string MeasurmentCM)
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



        public static System.Data.DataTable GetOPDN(string shippingcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select distinct t4.docentry 收貨採購單號,T10.U_PC_BSINV 進項發票號碼,cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,FILENAME+'.'+Fileext 檔案名稱  from oclg t2 ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY) ");
            sb.Append(" inner join opdn t4 on(t2.docentry=t4.docentry) ");
            sb.Append(" left join PDN1 t5 on (t4.docentry=T5.docentry )");
            sb.Append(" left join PCH1 t12 on (t12.baseentry=T5.docentry and  t12.baseline=t5.linenum and t12.basetype='20'  )");
            sb.Append(" left join OPCH t10 on (T12.DOCENTRY=T10.DOCENTRY )");
            sb.Append(" where  t2.doctype='20' and isnull(t3.[FILENAME],'') <> '' and t4.u_shipping_no=@shippingcode ");

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
        public static System.Data.DataTable GetOWTR(string WH)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select distinct t2.docentry 調撥單號,cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(380) )+'.'+Fileext 路徑,FILENAME+'.'+Fileext 檔案名稱  from oclg t2  ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)  ");
            sb.Append(" where  t2.doctype='67' and isnull(t3.[FILENAME],'') <> '' ");
            sb.Append(" AND T2.DocEntry IN (SELECT DISTINCT T0.DOCENTRY FROM OWTR T0");
            sb.Append(" LEFT JOIN WTR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)　");
            sb.Append(" WHERE T0.Filler ='OT001'　AND T1.WhsCode IN ('TW001','TW012','TW017'))");
            sb.Append(" AND (Attachment  LIKE '%WH%'　AND DOCTYPE='67' AND Attachment  LIKE '%X%')");
            sb.Append(" AND SUBSTRING(Attachment,CHARINDEX('WH', Attachment),14) IN (" + WH + ")");

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

        public static System.Data.DataTable GetHK(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ArriveDay ETA,receiveDay SHIPWAY,boatName 港名,shipment 裝船港,unloadCargo 卸貨港,soNo 提單號,forecastDay ETD FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        public static System.Data.DataTable GetSHPCAR4(string SHIPPINGCODE, string JOBNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT JOBNO FROM shipping_CAR2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND JOBNO <> @JOBNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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

        public static System.Data.DataTable GetDOWNLOAD2SEQ(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SEQ+1  SEQ FROM Download2 WHERE SHIPPINGCODE= @SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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

        public static System.Data.DataTable GetINVOM(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT INVOICENO+'-'+InvoiceNo_seq  FROM INVOICEM　WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingcode));
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
        public static System.Data.DataTable GetPACKCM(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MEASURMENTCM+ CASE WHEN COUNT(*) =1 THEN '' ELSE '*'+ CAST(COUNT(*) AS VARCHAR) END+  ' CM' CM FROM PackingListD WHERE SHIPPINGCODE=@SHIPPINGCODE AND MEASURMENTCM <> '' GROUP BY MEASURMENTCM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingcode));
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
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案名稱")
                {

                    for (int j = 0; j <= 1; j++)
                    {


                        System.Data.DataTable dt1 = GetOPDN(shippingCodeTextBox.Text);
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];


                        string aa = drw["path"].ToString() + "\\" + drw["路徑"].ToString();

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = drw["檔案名稱"].ToString();
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void download2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {

                DataGridView dgv = (DataGridView)sender;
                System.Data.DataTable dt1 = ship.Download2;
                int i = e.RowIndex;
                DataRow drw = dt1.Rows[i];

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {


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
                if (dgv.Columns[e.ColumnIndex].Name == "MARK2")
                {
                    string MARK = download2DataGridView.CurrentRow.Cells["MARK2"].Value.ToString();
                    //  string MARK = drw["MARK"].ToString();
                }
                //MARK2

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
                    int F1 = G1.IndexOf(BAU);
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

            if (dtm.Rows.Count == 0 || shipping_ItemDataGridView.Rows.Count == 1)
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

        private void button26_Click(object sender, EventArgs e, string DTYPES)
        {

            int k1 = 0;
            CHO1 = 0;
            CHO2 = 0;
            CHO3 = 0;

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

            string dg2 = "";
            if (checkBox2.Checked)
            {
                dg2 = "check";
            }

            string ds = "";
            string aa = cardCodeTextBox.Text;
            string bb = cardNameTextBox.Text;
            if (DTYPES == "")
            {
                object[] LookupValues = GetCardList(aa, dg, dg2, bb);

                if (LookupValues != null)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i <= LookupValues.Length - 1; i++)
                    {

                        sb.Append("'" + Convert.ToString(LookupValues[i]) + "',");

                    }
                    sb.Remove(sb.Length - 1, 1);
                    ds = sb.ToString();
                }
            }


            try
            {
                tabControl1.SelectedIndex = 0;

                if (!String.IsNullOrEmpty(hh))
                {
                    pinoTextBox.Text = hh;
                }

                System.Data.DataTable dt1 = GetOrdrship1(ds, DTYPES);

                System.Data.DataTable dt2 = ship.Shipping_Item;

                if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1030-00" || cardCodeTextBox.Text == "1349-00")
                {

                    DataRow dro = dt1.Rows[0];
                    string 最終客戶 = dro["最終客戶"].ToString();
                    add6TextBox.Text = 最終客戶;
                    shipping_OBUTextBox.Text = dro["正航單號"].ToString();

                    System.Data.DataTable dt22 = GetMenu.GetO(最終客戶);
                    if (dt22.Rows.Count > 0)
                    {
                        DataRow dro2 = dt22.Rows[0];
                        add2TextBox.Text = dro2["cardcode"].ToString();
                    }
                }
                else
                {
                    shipping_OBUTextBox.Text = "";
                    add6TextBox.Text = "";
                    add2TextBox.Text = "";
                }


                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    string DOC = drw["Docnum"].ToString();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["seqNo"] = "0";
                    drw2["Docentry"] = DOC;
                    /*string ITEMCODE = drw["ItemCode"].ToString();
                    string Dscription = drw["Dscription"].ToString();
                    string CUSTCODE = drw["客戶料號"].ToString();


                    if (!String.IsNullOrEmpty(CUSTCODE))
                    {
                        Dscription = drw["Dscription"].ToString() + " P/N:" + CUSTCODE;
                    }*/
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["PiNo"] = drw["NumAtCard"];
                    drw2["ItemRemark"] = "銷售訂單";
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["ItemPrice"] = drw["Price"];
                    drw2["linenum"] = drw["linenum"];
                    drw2["ItemAmount"] = drw["totalfrgn"];
                    drw2["SOLARPRICE"] = drw["U_SHIPPRICE"];
                    drw2["STATUS"] = drw["貨況"];
                    drw2["CHOMemo"] = drw["注意事項"].ToString();
                    drw2["OldOrder"] = drw["TREETYPE"].ToString();
                    drw2["VISORDER"] = drw["VISORDER"];
                    drw2["CHODOC"] = drw["正航單號"];
                    drw2["CURRENCY"] = drw["DOCCUR"];
                    drw2["RATE"] = drw["DOCRATE"];
                    drw2["WHSCODE"] = drw["WHSCODE"];
                    string DOCDATE = drw["DOCDATE"].ToString();

                    System.Data.DataTable B2 = GetDOCCUR2(DOCDATE);
                    if (B2.Rows.Count > 0)
                    {
                        drw2["RATEUSD"] = B2.Rows[0][0].ToString();
                    }

                    StringBuilder sb3 = new StringBuilder();

                    //drw["注意事項"].ToString()
                    string gj = "付款: " + drw["付款"].ToString() +
                     Environment.NewLine + "離倉日期: " + drw["離倉日期"].ToString() +
                     Environment.NewLine + "特殊嘜頭: " + drw["特殊嘜頭"].ToString() +
                     Environment.NewLine + "注意事項: " + dt1.Rows[0]["注意事項"].ToString() +
                     Environment.NewLine + "FORWARDER: " + drw["FORWARDER"].ToString() +
                     Environment.NewLine + "運輸方式: " + drw["運輸方式"].ToString() +
                     Environment.NewLine + "貿易條件: " + drw["貿易條件"].ToString() +
                     Environment.NewLine + "SHIP FROM: " + drw["shipform"].ToString() +
                     Environment.NewLine + "SHIP TO: " + drw["shipto"].ToString() +
                     Environment.NewLine + "付款方式: " + drw["付款方式"].ToString();

                    if (shipping_ItemDataGridView.Rows.Count != 1 || sAMEMOTextBox.Text == "")
                    {
                        if (!String.IsNullOrEmpty(付款) && drw["付款"].ToString().Trim().ToUpper() != 付款.Trim().ToUpper())
                        {
                            sb3.Append("付款" + "，");
                        }
                        if (!String.IsNullOrEmpty(離倉日期) && drw["離倉日期"].ToString().Trim().ToUpper() != 離倉日期.Trim().ToUpper())
                        {
                            sb3.Append("離倉日期" + "，");
                        }

                        if (!String.IsNullOrEmpty(特殊嘜頭) && drw["特殊嘜頭"].ToString().Trim().ToUpper() != 特殊嘜頭.Trim().ToUpper())
                        {
                            sb3.Append("特殊嘜頭" + "，");
                        }
                        if (!String.IsNullOrEmpty(注意事項) && drw["注意事項"].ToString().Trim().ToUpper() != 注意事項.Trim().ToUpper())
                        {
                            sb3.Append("注意事項" + "，");
                        }

                        if (!String.IsNullOrEmpty(FORWARDER) && drw["FORWARDER"].ToString().Trim().ToUpper() != FORWARDER.Trim().ToUpper())
                        {
                            sb3.Append("FORWARDER" + "，");
                        }
                        if (!String.IsNullOrEmpty(運輸方式) && drw["運輸方式"].ToString().Trim().ToUpper() != 運輸方式.Trim().ToUpper())
                        {
                            sb3.Append("運輸方式" + "，");
                        }
                        if (!String.IsNullOrEmpty(貿易條件) && drw["貿易條件"].ToString().Trim().ToUpper() != 貿易條件.Trim().ToUpper())
                        {
                            sb3.Append("貿易條件" + "，");
                        }
                        if (!String.IsNullOrEmpty(shipform) && drw["shipform"].ToString().Trim().ToUpper() != shipform.Trim().ToUpper())
                        {
                            sb3.Append("shipform" + "，");
                        }
                        if (!String.IsNullOrEmpty(shipto) && drw["shipto"].ToString().Trim().ToUpper() != shipto.Trim().ToUpper())
                        {
                            sb3.Append("shipto" + "，");
                        }
                        if (!String.IsNullOrEmpty(付款方式) && drw["付款方式"].ToString().Trim().ToUpper() != 付款方式.Trim().ToUpper())
                        {
                            sb3.Append("付款方式" + "，");
                        }

                        if (!String.IsNullOrEmpty(sb3.ToString()) & k1 == 0)
                        {
                            sb3.Remove(sb3.Length - 1, 1);

                            MessageBox.Show(sb3.ToString() + " 內容不一致");
                            k1 = 1;
                        }

                        sAMEMOTextBox.Text = gj;
                        付款 = drw["付款"].ToString().Trim();
                        離倉日期 = drw["離倉日期"].ToString().Trim();
                        特殊嘜頭 = drw["特殊嘜頭"].ToString().Trim();
                        注意事項 = drw["注意事項"].ToString().Trim();
                        FORWARDER = drw["FORWARDER"].ToString().Trim();
                        運輸方式 = drw["運輸方式"].ToString().Trim();
                        貿易條件 = drw["貿易條件"].ToString().Trim();
                        shipform = drw["shipform"].ToString().Trim();
                        shipto = drw["shipto"].ToString().Trim();
                        付款方式 = drw["付款方式"].ToString().Trim();
                    }
                    else
                    {
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




                    try
                    {


                        if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1349-00")
                        {



                            string S1 = drw["ItemCode"].ToString().Trim();
                            string S3 = drw["正航單號"].ToString().Trim();


                            int L1 = Convert.ToInt32(drw["Quantity"]);

                            string strCn = "";

                            if (cardCodeTextBox.Text == "0257-00")
                            {
                                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                            }
                            if (cardCodeTextBox.Text == "0511-00")
                            {
                                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                            }
                            if (cardCodeTextBox.Text == "1349-00")
                            {
                                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                            }
                            System.Data.DataTable FJ = GetCHOICE(S3, S1, L1, strCn);
                            System.Data.DataTable FJ2 = GetCHOICEIT(S3, S1, strCn);
                            System.Data.DataTable CH1 = GetCHOICECHECK(S3, S1, L1, strCn);

                            if (!String.IsNullOrEmpty(S3))
                            {
                                if (FJ2.Rows.Count == 0)
                                {
                                    MessageBox.Show("SAP與正航訂單料號不一致");
                                    CHO1 = 1;
                                }
                                else if (FJ.Rows.Count == 0)
                                {
                                    if (CHO1 == 0)
                                    {
                                        MessageBox.Show("SAP與正航訂單數量不一致");
                                        CHO3 = 1;
                                    }
                                }

                                else
                                {

                                    string P1 = FJ.Rows[0][0].ToString();
                                    string P2 = FJ.Rows[0][1].ToString();
                                    string P3 = FJ.Rows[0][2].ToString();
                                    string P4 = FJ.Rows[0][3].ToString().Substring(0, 250);
                                    string P5 = FJ.Rows[0][4].ToString();
                                    string P6 = FJ.Rows[0][5].ToString();
                                    if (CH1.Rows.Count > 0)
                                    {
                                        string CHJ = FJ.Rows[0][0].ToString();
                                        if (CHJ != "0")
                                        {
                                            CHO2 = 1;
                                            drw2["CHOPrice"] = P1;
                                            drw2["CHOAmount"] = P2;
                                            drw2["CHOMemo"] = P4;
                                            add8TextBox.Text = P3;
                                        }
                                        else
                                        {
                                            drw2["CHOPrice"] = 0;
                                            drw2["CHOAmount"] = 0;
                                            drw2["CHOMemo"] = P4;
                                            add8TextBox.Text = P3;
                                        }
                                    }
                                    else if (FJ.Rows.Count > 0)
                                    {
                                        drw2["CHOPrice"] = P1;
                                        drw2["CHOAmount"] = P2;
                                        drw2["CHOMemo"] = P4;
                                        add8TextBox.Text = P3;
                                    }
                                }
                            }
                        }

                    }
                    catch
                    {

                    }

                    dt2.Rows.Add(drw2);
                }



                shipping_ItemBindingSource.MoveFirst();

                for (int i = 0; i <= shipping_ItemBindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)shipping_ItemBindingSource.Current;

                    row3["SeqNo"] = i;



                    shipping_ItemBindingSource.EndEdit();

                    shipping_ItemBindingSource.MoveNext();
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            shipping_mainBindingSource.EndEdit();
            shipping_ItemBindingSource.EndEdit();


        }
        public static System.Data.DataTable GERCARDCHI(string BillNO)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT T0.CustomerID,T1.FullName FROM OrdBillMain T0 LEFT JOIN comCustomer T1 ON (T0.CustomerID =T1.ID AND T1.Flag=1 ) WHERE T0.FLAG=2 AND BillNO =@BillNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


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

        public static System.Data.DataTable GERCARD(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT CARDCODE,CARDNAME,U_ACME_TARDETERM 貿易條件,U_ACME_SHIPFORM1 收貨地,U_ACME_SHIPTO1 目的地,CASE  WHEN U_ACME_BYAIR IN ('Truck','SEA','AIR') THEN  U_ACME_BYAIR ELSE '' END 運輸方式 FROM ORDR WHERE DOCENTRY=@DOCENTRY ");
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
        public static System.Data.DataTable GERCARDOWTR(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT CARDCODE,CARDNAME,U_ACME_TARDETERM 貿易條件,U_ACME_SHIPFORM1 收貨地,U_ACME_SHIPTO1 目的地,CASE  WHEN U_ACME_BYAIR IN ('Truck','SEA','AIR') THEN  U_ACME_BYAIR ELSE '' END 運輸方式 FROM OWTR WHERE DOCENTRY=@DOCENTRY ");
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
        public static System.Data.DataTable GERCARDD(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE");
            sb.Append("  WHEN isnull(u_acme_workday,'') IN ('出口海運','出口合外','出口空運','急單(含快遞)') THEN '出口'  ");
            sb.Append(" WHEN u_acme_workday IN ('境外海運','境外合外','境外空運','境外陸運') THEN  '三角'  ");
            sb.Append(" WHEN  isnull(u_acme_workday,'') IN ('進口轉內銷','內銷')THEN  '內銷'  END 貿易形式");
            sb.Append("  FROM RDR1 WHERE DOCENTRY=@DOCENTRY  AND ISNULL(U_ACME_WORKDAY,'') <> ''");

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
        public static System.Data.DataTable GERCARD1(string TradeCondition)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DISTINCT TradeCondition FROM   SHIPPING_MAIN WHERE TradeCondition =@TradeCondition ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@TradeCondition", TradeCondition));


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
        public static System.Data.DataTable GERCARD3(string bb, string PORTTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT PORT FROM Account_Temp7 WHERE PORT LIKE '%" + bb + "%'  AND PORTTYPE =@PORTTYPE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PORTTYPE", PORTTYPE));
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
        public System.Data.DataTable GERCARD2(string BillNO, string ITEMREMARK)
        {

            SqlConnection connection = null;
            if (ITEMREMARK == "Choice")
            {
                connection = new SqlConnection(strCn);
            }
            if (ITEMREMARK == "Infinite")
            {
                connection = new SqlConnection(strCn22);
            }
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CustomerID CARDCODE,T1.FullName CARDNAME FROM OrdBillMain T0 LEFT JOIN comCustomer T1 ON (T0.CustomerID =T1.ID AND T1.Flag=1 ) WHERE T0.FLAG=2 AND BillNO =@BillNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private object[] GetCardList(string aa, string dg, string dg2, string bb)
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
            sb.Append("  WHERE  T1.TREETYPE <> 'I'    ");

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
        private void SICHECK()
        {
            DF3F.Length = 0;
            DF3F.Capacity = 0;
            if (rUSHCheckBox.Checked)
            {
                DF3F.Append("急貨");
            }
            if (add10CheckBox.Checked)
            {
                DF3F.Append("本票請申請 AUO貨代免倉期10天");
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
        private object[] GetCardListOPDN(string aa, string dg2, string bb)
        {

            string[] FieldNames = new string[] { "銷售單號", "倉庫名稱", "u_acme_work", "KEY", "單號", "序號" };

            string[] Captions = new string[] { "收貨採購單號", "倉庫名稱", "排程日期", "KEY", "單號", "收貨方" };

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct cast(T0.docentry as varchar) as 銷售單號,T2.WHSNAME as  倉庫名稱,Convert(varchar(8),t1.u_acme_work,112) as  u_acme_work,replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') 'KEY',T0.docentry 單號");
            sb.Append("  ,ISNULL(T3.序號,1) 序號 FROM OPDN T0 inner join PDN1 T1 on (t0.docentry=t1.docentry)  ");
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
            string sql = "SELECT WHSCODE DataValue FROM Shipping_WHS order by WHSCODE ";
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


        private void button15_Click(object sender, EventArgs e, string DOCTYPE, string FLAG, string DTYPES)
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
            if (DTYPES == "")
            {


                if (FLAG == "銷退")
                {
                    if (DOCTYPE == "Choice")
                    {
                        LookupValues = GetMenu.GetowtrTAO(cardCodeTextBox.Text);
                    }
                    if (DOCTYPE == "Infinite")
                    {
                        LookupValues = GetMenu.GetowtrTAO2(cardCodeTextBox.Text);
                    }
                    if (DOCTYPE == "TOP GARDEN")
                    {
                        LookupValues = GetMenu.GetowtrCHO2T(cardCodeTextBox.Text, dg);
                    }
                }
                else
                {
                    if (DOCTYPE == "Choice")
                    {
                        LookupValues = GetMenu.GetowtrCHO(cardCodeTextBox.Text, dg);
                    }
                    if (DOCTYPE == "Infinite")
                    {
                        LookupValues = GetMenu.GetowtrCHO2(cardCodeTextBox.Text, dg);
                    }
                    if (DOCTYPE == "TOP GARDEN")
                    {
                        LookupValues = GetMenu.GetowtrCHO2T(cardCodeTextBox.Text, dg);
                    }
                    if (DOCTYPE == "CHOICE採購")
                    {
                        LookupValues = GetMenu.GetowtrCHOCHO(cardCodeTextBox.Text, dg, FLAG);
                    }
                    if (DOCTYPE == "INFINITE採購")
                    {
                        LookupValues = GetMenu.GetowtrCHOCHO2(cardCodeTextBox.Text, dg, FLAG);
                    }

                }
            }
            if (LookupValues != null || pinoTextBox.Text != "")
            {
                tabControl1.SelectedIndex = 0;

                string docentry = "";

                if (DTYPES == "1")
                {
                    docentry = pinoTextBox.Text;
                }
                else
                {
                    docentry = Convert.ToString(LookupValues[0]);
                }
                pinoTextBox.Text = docentry;

                System.Data.DataTable dt1 = null;
                if (FLAG == "銷退")
                {
                    dt1 = GetSHIOTAO(pinoTextBox.Text, DOCTYPE);
                }
                else
                {
                    dt1 = GetCHO(docentry, DOCTYPE, FLAG);
                }



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
                    drw2["ItemRemark"] = DOCTYPE;
                    drw2["Quantity"] = drw["數量"];
                    drw2["CHOPrice"] = drw["單價"];
                    drw2["linenum"] = drw["ROWNO"];
                    drw2["CHOAmount"] = drw["金額"];

                    if (FLAG != "銷退")
                    {
                        drw2["CHOMemo"] = drw["備註"];

                    }
                    drw2["ItemPrice"] = drw["單價"];
                    drw2["ItemAmount"] = drw["金額"];
                    //if (DOCTYPE == "CHOICE採購" || DOCTYPE == "AD")
                    //{
                    //    drw2["ItemPrice"] = drw["單價"];
                    //    drw2["ItemAmount"] = drw["金額"];
                    //}

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


        }


        private void button15DIAOBO_Click(object sender, EventArgs e, string DOCTYPE, string FLAG)
        {

            object[] LookupValues = null;

            if (DOCTYPE == "Choice")
            {
                LookupValues = GetMenu.GetowtrDIAO();
            }
            if (DOCTYPE == "Infinite")
            {
                LookupValues = GetMenu.GetowtrDIAO2();
            }
            if (DOCTYPE == "TOP GARDEN")
            {
                LookupValues = GetMenu.GetowtrDIAOT();
            }

            if (LookupValues != null)
            {
                tabControl1.SelectedIndex = 0;

                pinoTextBox.Text = Convert.ToString(LookupValues[0]);

                tabControl1.SelectedIndex = 0;


                System.Data.DataTable dt1 = null;

                dt1 = GetCHODIAOBO(pinoTextBox.Text, DOCTYPE);


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
                    drw2["ItemRemark"] = DOCTYPE;
                    drw2["Quantity"] = drw["數量"];
                    drw2["CHOPrice"] = drw["單價"];
                    drw2["linenum"] = drw["ROWNO"];
                    drw2["CHOAmount"] = drw["金額"];
                    drw2["CHOMemo"] = drw["備註"];

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



        }

        public System.Data.DataTable GetCHO(string DocEntry, string CHO, string FLAG)
        {
            SqlConnection connection = null;
            if (CHO == "Choice" || CHO == "CHOICE採購")
            {
                connection = new SqlConnection(strCn);
            }
            if (CHO == "Infinite" || CHO == "INFINITE採購")
            {
                connection = new SqlConnection(strCn22);
            }
            if (CHO == "TOP GARDEN")
            {
                connection = new SqlConnection(strCn20);
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT T0.BillNO Docnum,T1.ProdID ItemCode,T1.ProdName Dscription,T1.Quantity 數量,T1.Price 單價,T1.Amount 金額,T1.ROWNO,T0.REMARK 備註 FROM OrdBillMain T0");
            sb.Append("                      Inner Join OrdBillSub T1 On T0.Flag=T1.Flag And T0.BillNO=T1.BillNO ");
            if (CHO == "AD")
            {
                sb.Append("                       where T0.BillNO=@BillNO and T0.Flag =@Flag");
            }
            else
            {
                if (CHO == "CHOICE採購" || CHO == "INFINITE採購")
                {
                    sb.Append("                       where T0.BillNO=@BillNO and T0.Flag =4");
                }
                else
                {
                    sb.Append("                       where T0.BillNO=@BillNO and T0.Flag =2 ");
                }
            }
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
        public System.Data.DataTable GetCHODIAOBO(string DocEntry, string CHO)
        {
            SqlConnection connection = null;
            if (CHO == "Choice")
            {
                connection = new SqlConnection(strCn);
            }
            if (CHO == "Infinite")
            {
                connection = new SqlConnection(strCn22);
            }
            if (CHO == "TOP GARDEN")
            {
                connection = new SqlConnection(strCn20);
            }


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

        public System.Data.DataTable GetSHIOTAO(string DocEntry, string CHO)
        {
            SqlConnection connection = null;
            if (CHO == "Choice")
            {
                connection = new SqlConnection(strCn);
            }
            if (CHO == "Infinite")
            {
                connection = new SqlConnection(strCn22);
            }
            if (CHO == "TOP GARDEN")
            {
                connection = new SqlConnection(strCn20);
            }


            StringBuilder sb = new StringBuilder();
            sb.Append("                       SELECT T0.FundBillNo Docnum,T1.ProdID ItemCode,T1.ProdName Dscription,T1.Quantity 數量,T1.Price 單價,T1.Amount 金額,T1.ROWNO,T0.REMARK 備註 FROM comBillAccounts T0 ");
            sb.Append("                       LEFT JOIN comProdRec T1 ON (T0.FundBillNo=T1.BillNO AND T0.Flag=T1.Flag)");
            sb.Append("                        WHERE T0.Flag=600 AND  T0.FundBillNo=@BILLNO");

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

        public System.Data.DataTable GetHR(string ENGNAME)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             SELECT ENGNAME FROM HR_CHI where    COMPANY in ('達睿生科技發展深圳有限公司') AND ENGNAME=@ENGNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ENGNAME", ENGNAME));

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

        public System.Data.DataTable GetOrdrship1(string Doc_no, string DRS)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.totalfrgn,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號 ");
            sb.Append("               ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER,U_SHIPPRICE");
            sb.Append(" ,T0.DOCCUR,T0.DOCRATE,Convert(varchar(8),T0.DOCDATE,112) DOCDATE,T1.WHSCODE     from rdr1 t1 ");
            sb.Append("               left join ordr t0 on (t1.docentry=t0.docentry)   where  1=1   ");
            sb.Append(" AND T1.TREETYPE <> 'I'  ");
            if (DRS == "1")
            {
                sb.Append(" AND T0.DOCNUM=@DOCNUM ");
            }
            else
            {
                sb.Append(" AND  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','')+isnull(u_acme_workday,'')  in (N" + Doc_no + ") order by t0.Docnum,visorder ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCNUM", pinoTextBox.Text));
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

        public System.Data.DataTable GetOrdrship2(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.totalfrgn,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號 ");
            sb.Append("               ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER,U_SHIPPRICE");
            sb.Append(" ,T0.DOCCUR,T0.DOCRATE,Convert(varchar(8),T0.DOCDATE,112) DOCDATE    from rdr1 t1 ");
            sb.Append("               left join ordr t0 on (t1.docentry=t0.docentry)   where  1=1   ");
            sb.Append(" AND T1.TREETYPE <> 'I'  ");
            sb.Append(" AND  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','')+isnull(u_acme_workday,'')  in (N" + Doc_no + ") order by t0.Docnum,visorder ");


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
        public System.Data.DataTable GetOrdrshipOPDN(string Doc_no)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.LINETOTAL,t0.u_acme_tardeterm 貿易條件,U_CHI_NO 正航單號");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,Convert(varchar(8),T1.U_ACME_SHIPDAY,112)  離倉日期,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式,TREETYPE,VISORDER   from PDN1 t1");
            sb.Append(" left join OPDN t0 on (t1.docentry=t0.docentry) where  replace(replace(replace(ISNULL(Convert(nvarchar(8),t1.u_acme_work,112),'')+ISNULL(cast(T0.Docnum as nvarchar),'')+ISNULL(T1.WHSCODE,'')+ISNULL(T0.SHIPTOCODE,''),'''',''),' ',''),'.','') in (" + Doc_no + ") order by t0.Docnum,visorder ");


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
        public System.Data.DataTable Getshipitem(string shippingcode, int TYPE, string DOC)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();

            sb.Append(" select t1.itemcode,Dscription,Quantity,bb=      ");
            sb.Append(" case  WHEN T2.U_GROUP='Z&R-費用類群組' then  'FREIGHT'       WHEN substring(t1.itemcode,1,4) ='ACME' then Dscription      ");
            sb.Append(" ELSE   U_ITEMNAME COLLATE  Chinese_Taiwan_Stroke_CI_AS +' '+REPLACE(ISNULL(U_MODEL,''),'NON','')   end,  ");
            sb.Append(" ItemPrice  ");
            sb.Append(" ,t1.Docentry,linenum,CHOAmount,CHOPrice,OLDORDER,VISORDER,T1.CURRENCY,T1.RATE,T1.RATEUSD,T1.ItemAmount  from shipping_item T1  ");
            sb.Append(" LEFT JOIN  ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");

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

        public System.Data.DataTable Getshipitem07()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT * FROM ACMESQLSP_TEST.DBO.INVOICED WHERE SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", textBox8.Text.Trim()));
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




        public System.Data.DataTable GetSHIPPACKINV()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT INVOICE INV  FROM SHIPPING_PACK  WHERE ISNULL(INVOICE,'') <>'' AND users=@USERS ");

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

        public System.Data.DataTable GetSHIPOITM2(string MODEL)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PARAM_DESC  FROM PARAMS WHERE PARAM_KIND ='EVAITEM'　AND PARAM_NO =@MODEL ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
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

        public System.Data.DataTable GetSHIPOITM3(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUBSTRING(ITEMNAME,11,2)     FROM SHIPPING_PACK ");
            sb.Append(" where users=@USERS  AND ITEMCODE=@ITEMCODE   ");
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


        public System.Data.DataTable GetSUMPACK()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(CAST(CARTONNO AS INT)) CARTNO FROM WH_PACK2 WHERE SHIPPINGCODE IN (SELECT DISTINCT WHNO  FROM PackingListD   WHERE SHIPPINGCODE=@SHIPPINGCODE) AND QTY <>'空箱' ");
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

        public System.Data.DataTable GetSHIPOITM4(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  SUBSTRING(U_PARTNO,11,2)      FROM OITM WHERE ITEMCODE=@ITEMCODE  ");
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
        public System.Data.DataTable GetWHPACK2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  T0.MATERIAL,MAX(T1.SEQ) SEQ  FROM WH_PACK2  T0");
            sb.Append(" LEFT JOIN (SELECT MAX(PLATENO) SEQ,MATERIAL  FROM WH_PACK2  T0 WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ") AND QTY <> '空箱' ");
            sb.Append(" GROUP BY MATERIAL ) T1 ON (T0.MATERIAL =T1.MATERIAL)");
            sb.Append("  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ")  AND QTY <> '空箱' ");
            sb.Append(" GROUP BY T0.MATERIAL ORDER BY SEQ");
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

        public System.Data.DataTable GetWHPACK2ES2(string SHIPPINGCODE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_ITEMNAME+' '+U_MODEL MODEL,' ('+CASE WHEN T1.U_GRADE='NN' THEN 'N' ELSE T1.U_GRADE END+' GRADE)' GRADE,T0.ES,T1.U_MODEL TMODEL   FROM WH_PACK2 T0 ");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ") AND T0.ITEMCODE=@ITEMCODE  ");

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

        public System.Data.DataTable GetWHPACK2ES3(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_ITEMNAME+' '+U_MODEL MODEL,' ('+CASE WHEN T1.U_GRADE='NN' THEN 'N' ELSE T1.U_GRADE END+' GRADE)' GRADE  FROM  AcmeSql02.DBO.OITM T1");
            sb.Append(" WHERE  T1.ITEMCODE=@ITEMCODE  ");

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
        public System.Data.DataTable GetWHLOCATION(string SHIPPINGCODE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT LOACTION  FROM WH_PACK2  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  ) AND ITEMCODE=@ITEMCODE");

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
        public System.Data.DataTable GetCTN(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT PALQTY PACKAGENO,SEQNO  FROM PackingListD    WHERE SHIPPINGCODE=@SHIPPINGCODE  AND PLNO=@PLNO AND   ISNULL(PALQTY,'') <> '' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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


        public System.Data.DataTable GetWHITEM(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'ITEM'+CASE WHEN MIN(CAST(seqno2+1 AS INT))=MAX(CAST(seqno2+1 AS INT)) THEN CAST(MAX(CAST(seqno2+1 AS INT)) AS VARCHAR) ELSE ");
            sb.Append("   CAST(MIN(CAST(seqno2+1 AS INT)) AS VARCHAR)+'~'+CAST(MAX(CAST(seqno2+1 AS INT)) AS VARCHAR)  END  + ')MADE IN '+LOCATION  PLATENO,LOCATION FROM PackingListD");
            sb.Append("   WHERE SHIPPINGCODE =@SHIPPINGCODE AND PLNO=@PLNO   AND ISNULL(seqno2,'') <> '' AND ISNULL(LOCATION,'') <> '' ");
            sb.Append("   GROUP BY LOCATION ORDER BY MIN(CAST(seqno2+1 AS INT))");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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

        public System.Data.DataTable GetWHITEMP(string SHIPPINGCODE, string PLNO, string LOCATION)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("       SELECT seqno2+1 SEQ,LOCATION FROM PackingListD WHERE SHIPPINGCODE =@SHIPPINGCODE AND PLNO=@PLNO   AND ISNULL(seqno2,'') <> '' AND ISNULL(LOCATION,'') = @LOCATION ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
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
        public System.Data.DataTable GetWHITEM2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'ITEM'+CASE WHEN MIN(CAST(seqno2+1 AS INT))=MAX(CAST(seqno2+1 AS INT)) THEN CAST(MAX(CAST(seqno2+1 AS INT)) AS VARCHAR) ELSE ");
            sb.Append("   CAST(MIN(CAST(seqno2+1 AS INT)) AS VARCHAR)+'~'+CAST(MAX(CAST(seqno2+1 AS INT)) AS VARCHAR)  END  + ')MADE IN '+LOCATION  PLATENO FROM InvoiceD");
            sb.Append("   WHERE SHIPPINGCODE =@SHIPPINGCODE   AND ISNULL(seqno2,'') <> ''  AND ISNULL(LOCATION,'') <> '' ");
            sb.Append("   GROUP BY LOCATION ORDER BY MIN(CAST(seqno2+1 AS INT))");
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
        public System.Data.DataTable GetWHITEM2S(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" 				SELECT ITEMCODE,DOCENTRY FROM InvoiceD WHERE SHIPPINGCODE =@SHIPPINGCODE  ");
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
        public System.Data.DataTable GetWHPACK3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM WH_PACK2  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )  AND QTY <> '空箱'  ");
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
        public System.Data.DataTable GetSHIPPCAK()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT NET,GROSS  FROM PackingListD  WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNO=@PLNO AND NET NOT LIKE '%@%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNO", pLNoTextBox.Text));
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
        public System.Data.DataTable GetWHPACK5(string SHIPPINGCODE, string MATERIAL)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  CAST(MIN(CAST(PLATENO AS INT)) AS VARCHAR)+'-'+CAST(MAX(CAST(PLATENO AS INT)) AS VARCHAR) PLATENO,MATERIAL   FROM WH_PACK2  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  ) AND ISNULL(PLATENO,'') <> '' AND QTY <>'空箱' AND CAST(MATERIAL AS VARCHAR)    = ('" + MATERIAL + "'  )  GROUP BY SEQNO,MATERIAL  ORDER BY MATERIAL,SEQNO ");


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

        public System.Data.DataTable GetWHPACK5N(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT SUM(CAST(PLATENO AS INT)) PLATE FROM (");
            //sb.Append("  SELECT  CAST(MAX(CAST(PLATENO AS INT)) AS VARCHAR) PLATENO   FROM WH_PACK2 ");
            //sb.Append("  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )  AND ISNULL(PLATENO,'') <> '' ");
            //sb.Append("   GROUP BY SEQNO  ) AS A");



            sb.Append("  SELECT  CAST(MAX(CAST(PLATENO AS INT)) AS VARCHAR) PLATENO   FROM WH_PACK2 ");
            sb.Append("  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )  AND ISNULL(PLATENO,'') <> '' ");
            sb.Append("   GROUP BY SEQNO ");
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

        public System.Data.DataTable GetWHPACKCBM(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(ROUND(SUM(CAST(L AS DECIMAL(10,2))*CAST(W AS DECIMAL(10,2))*CAST(H AS DECIMAL(10,2)))/1000000,2) AS decimal(10,2)) CBM  FROM  WH_PACK2 WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )   AND ISNULL(L,'') <> '' AND ISNULL(W,'') <> ''  AND  ISNULL(H,'') <> ''  AND QTY <>'空箱' ");
            sb.Append("  HAVING ISNULL(CAST(ROUND(SUM(CAST(L AS DECIMAL(10,2))*CAST(W AS DECIMAL(10,2))*CAST(H AS DECIMAL(10,2)))/1000000,2) AS decimal(10,2)),0) <> 0");
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

        public System.Data.DataTable GetWHPACKCBM2(string SHIPPINGCODE, string DTYPE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT MeasurmentCM,(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,2) END AS INT)) PLATENO,PackageNo PACK  FROM PackingListD WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(MeasurmentCM,'') <> ''");
            if (DTYPE == "2")
            {
                sb.Append("AND PLNO=@PLNO ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", pLNoTextBox.Text));
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












        public System.Data.DataTable GetSHIPPACK8(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT MAX(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,2) END AS INT))  PackageNo,MAX(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,3) END AS INT))  CNO  FROM PackingListD  WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNO=@PLNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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



        public System.Data.DataTable GetB1(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT DISTINCT WHNO FROM PackingListD WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNO=@PLNO AND ISNULL(WHNO,'') <> ''  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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

        public System.Data.DataTable GetB2S(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ISNULL(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,3) END as int ),0) PACK,ISNULL(DescGoods,'')+ISNULL(QUANTITY,'')+ISNULL(NET,'')+ISNULL(GROSS,'') DESCS,substring(PackageNo,1,1) P3     ");
            sb.Append("      FROM PackingListD  WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNO=@PLNO ");
            sb.Append(" AND ISNULL(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,3) END as int ),0) <> 0    ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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

        public System.Data.DataTable GetB2S2(string SHIPPINGCODE, string PLNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT   ISNULL(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,4) END AS INT),0) CNO,ISNULL(DescGoods,'')+ISNULL(QUANTITY,'')+ISNULL(NET,'')+ISNULL(GROSS,'') DESCS  ");
            sb.Append("      FROM PackingListD  WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNO=@PLNO ");
            sb.Append(" AND  ISNULL(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,4) END AS INT),0) <> 0    ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));
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
        public System.Data.DataTable GetB2(string WHNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT MAX(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,2) END AS INT))  PackageNo   FROM PackingListD  WHERE WHNO=@WHNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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
        public System.Data.DataTable GetB3(string WHNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT TOP 1 PLATENO   FROM wh_pack2  where shippingcode = @WHNO");
            //   sb.Append("    SELECT MAX(CAST(CASE ISNULL(CHARINDEX('-', PackageNo),0) WHEN 0 THEN PackageNo ELSE SUBSTRING(PackageNo,CHARINDEX('-', PackageNo)+1,2) END AS INT))  PackageNo   FROM PackingListD  WHERE WHNO=@WHNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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
        public System.Data.DataTable GetB2CNO(string WHNO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MAX(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,3) END AS INT))  CNO   FROM PackingListD  WHERE WHNO=@WHNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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
        public System.Data.DataTable GetCHONO()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_CHI_NO FROM ORDR WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", pinoTextBox.Text));
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

        public System.Data.DataTable GetCHOITEM(string ProdID, string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ENGNAME FROM DBO.comProduct   WHERE ProdID=@ProdID AND ISNULL(ENGNAME,'') <>'' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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

        public System.Data.DataTable GetDOCCUR(string DOCENTRY, string ORDR)
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
                                System.Data.DataTable G1 = GetMenu.GetSA(pinoTextBox.Text.Trim());
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

                this.download2BindingSource.EndEdit();
                this.download2TableAdapter.Update(ship.Download2);
                ship.Download2.AcceptChanges();

                this.downloadBindingSource.EndEdit();
                this.downloadTableAdapter.Update(ship.Download);
                ship.Download.AcceptChanges();
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

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            add1TextBox.Text = comboBox4.Text;
        }

        private void comboBox4_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = GetMenu.Getfee("add1");


            comboBox4.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox4.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }


        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                add7TextBox.Text = comboBox5.Text;
                System.Data.DataTable O1 = GetSHIPEXSIT();
                System.Data.DataTable O2 = GetMenu.GetSHIPOHEM2(comboBox5.Text);
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
            if (rUSHTextBox.Text != "Checked")
            {
                rUSHCheckBox.Checked = false;
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
                System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHSHIP();
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
                System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHSTOCK();
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

                this.packingListDBindingSource.EndEdit();
                this.packingListDTableAdapter.Update(ship.PackingListD);
                ship.PackingListD.AcceptChanges();
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
        public void UPDATESHIPWHNO(string WHNO, string WHSCODE, string ShippingCode)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE shipping_item SET WHNO=@WHNO WHERE SHIPPINGCODE=@ShippingCode AND WHSCODE=@WHSCODE  AND  ISNULL(WHNO,'')=''", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public void UPDATESHIPWHNO2(string WHNO, string ShippingCode, string DOCENTRY1)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE shipping_item SET WHNO=@WHNO WHERE SHIPPINGCODE=@ShippingCode AND   DOCENTRY1 IN ( " + DOCENTRY1 + ") ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
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
        public void UPDATESAP(string CHECK, string ID)
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

        private void button36_Click(object sender, EventArgs e, string TYPE)
        {
            int k1 = 0;


            if (cardCodeTextBox.Text == "")
            {
                MessageBox.Show("請輸入客戶編號");

                return;
            }


            string dg2 = "";
            if (checkBox2.Checked)
            {
                dg2 = "check";
            }
            string aa = cardCodeTextBox.Text;
            string bb = cardNameTextBox.Text;
            object[] LookupValues = null;

            if (TYPE == "AR貸項")
            {
                LookupValues = GetCardListORIN(aa, dg2, bb);
            }
            if (TYPE == "AR貸項草稿")
            {
                LookupValues = GetCardListORINT(aa, dg2, bb);
            }
            if (TYPE == "AP貸項")
            {
                LookupValues = GetCardListORPC(aa, dg2, bb);
            }
            if (TYPE == "採購退貨")
            {
                LookupValues = GetCardListORPD(aa, dg2, bb);
            }
            if (TYPE == "收貨採購")
            {
                LookupValues = GetCardListOPDN(aa, dg2, bb);
            }
            if (LookupValues != null)
            {

                StringBuilder sb = new StringBuilder();


                for (int i = 0; i <= LookupValues.Length - 1; i++)
                {

                    sb.Append("'" + Convert.ToString(LookupValues[i]) + "',");

                }
                sb.Remove(sb.Length - 1, 1);
                string ds = sb.ToString();

                try
                {
                    tabControl1.SelectedIndex = 0;


                    pinoTextBox.Text = hh;

                    System.Data.DataTable dt1 = null;
                    if (TYPE == "AR貸項草稿")
                    {
                        dt1 = GetOrdrshipORINT(ds);
                    }
                    if (TYPE == "AR貸項")
                    {
                        dt1 = GetOrdrshipORIN(ds);
                    }
                    if (TYPE == "AP貸項")
                    {
                        dt1 = GetOrdrshipORPC(ds);
                    }
                    if (TYPE == "採購退貨")
                    {
                        dt1 = GetOrdrshipORPD(ds);
                    }
                    if (TYPE == "收貨採購")
                    {
                        dt1 = GetOrdrshipOPDN(ds);
                    }
                    System.Data.DataTable dt2 = ship.Shipping_Item;

                    if (cardCodeTextBox.Text == "0257-00" || cardCodeTextBox.Text == "0511-00" || cardCodeTextBox.Text == "1030-00" || cardCodeTextBox.Text == "1349-00")
                    {

                        DataRow dro = dt1.Rows[0];
                        string 最終客戶 = dro["最終客戶"].ToString();
                        add6TextBox.Text = 最終客戶;
                        shipping_OBUTextBox.Text = dro["正航單號"].ToString();

                        System.Data.DataTable dt22 = GetMenu.GetO(最終客戶);
                        if (dt22.Rows.Count > 0)
                        {
                            DataRow dro2 = dt22.Rows[0];
                            add2TextBox.Text = dro2["cardcode"].ToString();
                        }
                    }
                    else
                    {
                        shipping_OBUTextBox.Text = "";
                        add6TextBox.Text = "";
                        add2TextBox.Text = "";
                    }


                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();

                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["seqNo"] = "0";
                        drw2["Docentry"] = drw["Docnum"];
                        drw2["ItemCode"] = drw["ItemCode"];
                        drw2["Dscription"] = drw["Dscription"];
                        drw2["PiNo"] = drw["NumAtCard"];
                        drw2["ItemRemark"] = TYPE;
                        drw2["Quantity"] = drw["Quantity"];
                        drw2["ItemPrice"] = drw["Price"];
                        drw2["linenum"] = drw["linenum"];
                        drw2["ItemAmount"] = drw["LINETOTAL"];
                        drw2["STATUS"] = drw["貨況"];
                        drw2["CHOMemo"] = drw["注意事項"].ToString();
                        drw2["OldOrder"] = drw["TREETYPE"].ToString();
                        drw2["VISORDER"] = drw["VISORDER"];
                        drw2["CHODOC"] = drw["正航單號"];


                        StringBuilder sb3 = new StringBuilder();


                        string gj = "付款: " + drw["付款"].ToString() +
                         Environment.NewLine + "離倉日期: " + drw["離倉日期"].ToString() +
                         Environment.NewLine + "特殊嘜頭: " + drw["特殊嘜頭"].ToString() +
                         Environment.NewLine + "注意事項: " + drw["注意事項"].ToString() +
                         Environment.NewLine + "FORWARDER: " + drw["FORWARDER"].ToString() +
                         Environment.NewLine + "運輸方式: " + drw["運輸方式"].ToString() +
                         Environment.NewLine + "貿易條件: " + drw["貿易條件"].ToString() +
                         Environment.NewLine + "SHIP FROM: " + drw["shipform"].ToString() +
                         Environment.NewLine + "SHIP TO: " + drw["shipto"].ToString() +
                         Environment.NewLine + "付款方式: " + drw["付款方式"].ToString();

                        if (!String.IsNullOrEmpty(付款) && drw["付款"].ToString().Trim() != 付款.Trim())
                        {
                            sb3.Append("付款" + "，");
                        }
                        if (!String.IsNullOrEmpty(離倉日期) && drw["離倉日期"].ToString().Trim() != 離倉日期.Trim())
                        {
                            sb3.Append("離倉日期" + "，");
                        }

                        if (!String.IsNullOrEmpty(特殊嘜頭) && drw["特殊嘜頭"].ToString().Trim() != 特殊嘜頭.Trim())
                        {
                            sb3.Append("特殊嘜頭" + "，");
                        }
                        if (!String.IsNullOrEmpty(注意事項) && drw["注意事項"].ToString().Trim() != 注意事項.Trim())
                        {
                            sb3.Append("注意事項" + "，");
                        }

                        if (!String.IsNullOrEmpty(FORWARDER) && drw["FORWARDER"].ToString().Trim() != FORWARDER.Trim())
                        {
                            sb3.Append("FORWARDER" + "，");
                        }
                        if (!String.IsNullOrEmpty(運輸方式) && drw["運輸方式"].ToString().Trim() != 運輸方式.Trim())
                        {
                            sb3.Append("運輸方式" + "，");
                        }
                        if (!String.IsNullOrEmpty(貿易條件) && drw["貿易條件"].ToString().Trim() != 貿易條件.Trim())
                        {
                            sb3.Append("貿易條件" + "，");
                        }
                        if (!String.IsNullOrEmpty(shipform) && drw["shipform"].ToString().Trim() != shipform.Trim())
                        {
                            sb3.Append("shipform" + "，");
                        }
                        if (!String.IsNullOrEmpty(shipto) && drw["shipto"].ToString().Trim() != shipto.Trim())
                        {
                            sb3.Append("shipto" + "，");
                        }
                        if (!String.IsNullOrEmpty(付款方式) && drw["付款方式"].ToString().Trim() != 付款方式.Trim())
                        {
                            sb3.Append("付款方式" + "，");
                        }

                        if (!String.IsNullOrEmpty(sb3.ToString()) & k1 == 0)
                        {
                            sb3.Remove(sb3.Length - 1, 1);

                            MessageBox.Show(sb3.ToString() + " 內容不一致");
                            k1 = 1;

                        }

                        sAMEMOTextBox.Text = gj;

                        付款 = drw["付款"].ToString().Trim();
                        離倉日期 = drw["離倉日期"].ToString().Trim();
                        特殊嘜頭 = drw["特殊嘜頭"].ToString().Trim();
                        注意事項 = drw["注意事項"].ToString().Trim();
                        FORWARDER = drw["FORWARDER"].ToString().Trim();
                        運輸方式 = drw["運輸方式"].ToString().Trim();
                        貿易條件 = drw["貿易條件"].ToString().Trim();
                        shipform = drw["shipform"].ToString().Trim();
                        shipto = drw["shipto"].ToString().Trim();
                        付款方式 = drw["付款方式"].ToString().Trim();


                        dt2.Rows.Add(drw2);
                    }



                    shipping_ItemBindingSource.MoveFirst();

                    for (int i = 0; i <= shipping_ItemBindingSource.Count - 1; i++)
                    {
                        DataRowView row3 = (DataRowView)shipping_ItemBindingSource.Current;

                        row3["SeqNo"] = i;



                        shipping_ItemBindingSource.EndEdit();

                        shipping_ItemBindingSource.MoveNext();
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                shipping_mainBindingSource.EndEdit();
                shipping_ItemBindingSource.EndEdit();

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



        private void button38_Click(object sender, EventArgs e)
        {

            KITINVOICE frm1 = new KITINVOICE();
            frm1.q1 = shippingCodeTextBox.Text;
            frm1.q2 = invoiceNoTextBox.Text;
            frm1.q3 = invoiceNo_seqTextBox.Text;
            if (invoiceDDataGridView.SelectedRows.Count > 0)
            {
                frm1.q4 = invoiceDDataGridView.SelectedRows[0].Cells["INDescription"].Value.ToString();
            }
            else
            {
                MessageBox.Show("請選擇要轉出的列");
                return;
            }
            frm1.Show();

        }

        private void button39_Click(object sender, EventArgs e)
        {
            KITPACKING frm1 = new KITPACKING();
            frm1.q1 = shippingCodeTextBox.Text;
            frm1.q2 = pLNoTextBox.Text;
            if (packingListDDataGridView.SelectedRows.Count > 0)
            {
                frm1.q3 = packingListDDataGridView.SelectedRows[0].Cells["dataGridViewTextBoxColumn47"].Value.ToString();
            }
            else
            {
                MessageBox.Show("請選擇要轉出的列");
                return;
            }

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                packingListDTableAdapter.Fill(ship.PackingListD, MyID);
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            System.Data.DataTable H1 = GetOrderData3();
            if (H1.Rows.Count > 0)
            {
                FileName = lsAppDir + "\\Excel\\PACK.xls";
                GetExcelProduct(FileName, GetOrderData3(), "N", "N");
            }

            System.Data.DataTable H2 = GetOrderData3BOM();
            if (H2.Rows.Count > 0)
            {
                FileName = lsAppDir + "\\Excel\\PACKKIT.xls";
                GetExcelProductBOM(FileName, H2);
            }

        }

        private void button41_Click(object sender, EventArgs e)
        {
            CalcTotals1();

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\INVO2.xls";
            GetExcelProduct2(FileName, GetTT(), "N");


        }



        private void button126_Click(object sender, EventArgs e)
        {
            if (dOCTYPETextBox.Text == "")
            {
                MessageBox.Show("請選擇單據");
                return;
            }
            int T1 = add1TextBox.Text.IndexOf("正航系統");
            if (T1 == -1)
            {
                if (dOCTYPETextBox.Text == "銷售訂單")
                {

                    button26_Click(sender, e, "");


                }

                if (dOCTYPETextBox.Text == "採購")
                {
                    button6_Click(sender, e);
                }
                if (dOCTYPETextBox.Text == "調撥單")
                {
                    button8_Click(sender, e);
                }


                if (dOCTYPETextBox.Text == "發貨")
                {
                    button23_Click(sender, e, "發貨");
                }

                if (dOCTYPETextBox.Text == "收貨")
                {
                    button23_Click(sender, e, "收貨");
                }
                if (dOCTYPETextBox.Text == "AR")
                {
                    button36_Click(sender, e, "AR");
                }
                if (dOCTYPETextBox.Text == "AR貸項")
                {
                    button36_Click(sender, e, "AR貸項");
                }
                if (dOCTYPETextBox.Text == "AR貸項草稿")
                {
                    button36_Click(sender, e, "AR貸項草稿");
                }
                if (dOCTYPETextBox.Text == "AP貸項")
                {
                    button36_Click(sender, e, "AP貸項");
                }
                if (dOCTYPETextBox.Text == "採購退貨")
                {
                    button36_Click(sender, e, "採購退貨");
                }
                if (dOCTYPETextBox.Text == "收貨採購")
                {
                    button36_Click(sender, e, "收貨採購");
                }
            }
            else
            {

                if (add1TextBox.Text == "正航系統CHOICE")
                {
                    if (dOCTYPETextBox.Text == "銷售訂單")
                    {
                        button15_Click(sender, e, "Choice", "", "");
                    }
                    if (dOCTYPETextBox.Text == "採購")
                    {
                        button15_Click(sender, e, "CHOICE採購", "4", "");
                    }
                    if (dOCTYPETextBox.Text.Replace("單", "") == "調撥")
                    {
                        button15DIAOBO_Click(sender, e, "Choice", "調撥");
                    }
                    if (dOCTYPETextBox.Text == "銷退")
                    {
                        button15_Click(sender, e, "Choice", "銷退", "");
                    }
                }

                if (add1TextBox.Text == "正航系統INFINITE")
                {
                    if (dOCTYPETextBox.Text == "銷售訂單")
                    {
                        button15_Click(sender, e, "Infinite", "", "");
                    }
                    if (dOCTYPETextBox.Text == "採購")
                    {
                        button15_Click(sender, e, "INFINITE採購", "4", "");
                    }
                    if (dOCTYPETextBox.Text.Replace("單", "") == "調撥")
                    {
                        button15DIAOBO_Click(sender, e, "Infinite", "調撥");
                    }
                    if (dOCTYPETextBox.Text == "銷退")
                    {
                        button15_Click(sender, e, "Infinite", "銷退", "");
                    }
                }
                if (add1TextBox.Text == "正航系統TOP GARDEN")
                {
                    if (dOCTYPETextBox.Text == "銷售訂單")
                    {
                        button15_Click(sender, e, "TOP GARDEN", "", "");
                    }
                    if (dOCTYPETextBox.Text == "採購")
                    {
                        button15_Click(sender, e, "TOP GARDEN採購", "4", "");
                    }
                    if (dOCTYPETextBox.Text.Replace("單", "") == "調撥")
                    {
                        button15DIAOBO_Click(sender, e, "TOP GARDEN", "調撥");
                    }
                    if (dOCTYPETextBox.Text == "銷退")
                    {
                        button15_Click(sender, e, "TOP GARDEN", "銷退", "");
                    }
                }



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

        private void mEMO1TextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {


                string MEMOT = mEMO1TextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("SQ201");
                string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                if (G1 != -1)
                {
                    string[] arrurl = H1.Split(new Char[] { ';' });

                    foreach (string i in arrurl)
                    {
                        MEMO = i.Substring(0, 14);
                        SQUT2 a = new SQUT2();
                        a.PublicString = MEMO;
                        a.Show();
                    }

                }
            }
            catch { }

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

        private void button45_Click(object sender, EventArgs e)
        {

            try
            {

                System.Data.DataTable dt3 = Getshipitem07();

                System.Data.DataTable dt4 = ship.InvoiceD;

                if (dt3.Rows.Count > 0 && invoiceDDataGridView.Rows.Count < 2)
                {

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt3.Rows[i];
                        DataRow drw2 = dt4.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["InvoiceNo"] = invoiceNoTextBox.Text;
                        drw2["InvoiceNo_seq"] = invoiceNo_seqTextBox.Text;
                        drw2["SeqNo"] = i.ToString();
                        drw2["INDescription"] = drw["INDescription"];
                        drw2["InQty"] = drw["InQty"];
                        drw2["UnitPrice"] = drw["UnitPrice"];

                        drw2["SOID"] = drw["SOID"];
                        drw2["amount"] = drw["amount"];

                        drw2["LINENUM"] = drw["LINENUM"];


                        drw2["CHOPrice"] = drw["CHOPrice"];
                        drw2["CHOAmount"] = drw["CHOAmount"];
                        drw2["TREETYPE"] = drw["TREETYPE"]; ;
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



        private void button47_Click(object sender, EventArgs e)
        {
            string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            System.Data.DataTable dt1CHO = GetCHO3(shipping_OBUTextBox.Text, strCn);
            if (dt1CHO.Rows.Count > 0)
            {

                DataRow drw = dt1CHO.Rows[0];

                oBUShipToTextBox1.Text = drw["shipbuilding"].ToString() +
                        Environment.NewLine + drw["shipstreet"].ToString() +
                        Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                        Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                        Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();

            }
            string CHIBILLNO = "";
            System.Data.DataTable DTH = GetCHONO();
            if (DTH.Rows.Count > 0)
            {
                CHIBILLNO = DTH.Rows[0][0].ToString();
            }
            if (!String.IsNullOrEmpty(CHIBILLNO))
            {
                System.Data.DataTable dt2CHO2 = GetCHO22(CHIBILLNO, strCn);
                if (dt2CHO2.Rows.Count > 0)
                {
                    textBox19.Text = dt2CHO2.Rows[0][0].ToString();
                }

            }
            System.Data.DataTable dt2CHO = GetCHO2(textBox19.Text, strCn);
            if (dt2CHO.Rows.Count > 0)
            {

                DataRow drw = dt2CHO.Rows[0];

                oBUBillToTextBox1.Text = drw["billbuilding"].ToString() +
                Environment.NewLine + drw["billstreet"].ToString() +
                Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();
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
        private void SBS()
        {
            string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS.Append("'" + i + "',");
            }
            sbS.Remove(sbS.Length - 1, 1);
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

                        drw2["INDescription"] = drw["bb"];
                        drw2["InQty"] = drw["Quantity"];
                        drw2["UnitPrice"] = drw["ItemPrice"];
                        // drw2["amount"] = drw["ItemAmount"];
                        string TYPE = drw["OLDORDER"].ToString();
                        int T1 = add1TextBox.Text.IndexOf("正航系統");
                        if (T1 == -1)
                        {
                            drw2["amount"] = 1;
                            drw2["SOID"] = drw["Docentry"];
                        }




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

        private void button49_Click(object sender, EventArgs e)
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
            LookupValues = GetMenu.GetowtrCHO2(cardCodeTextBox.Text, dg);


            if (LookupValues != null)
            {
                tabControl1.SelectedIndex = 0;

                string docentry = Convert.ToString(LookupValues[0]);
                pinoTextBox.Text = docentry;

                System.Data.DataTable dt1 = GetCHO(docentry, "TOP GARDEN", "");
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
                    drw2["ItemRemark"] = "TOP GARDEN";
                    drw2["Quantity"] = drw["數量"];
                    drw2["CHOPrice"] = drw["單價"];
                    drw2["linenum"] = drw["ROWNO"];
                    drw2["CHOAmount"] = drw["金額"];
                    drw2["CHOMemo"] = drw["備註"];

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
            int T1 = add1TextBox.Text.IndexOf("正航系統");
            if (T1 == -1)
            {
                dt3 = GetMenu.GetBUGB("SHIPTYPE2");
            }
            else
            {
                dt3 = GetMenu.GetBUGB("SHIPTYPE3");
            }



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


        public System.Data.DataTable GetOHEMSHIP1()
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
        public System.Data.DataTable GetOHEMSHIP2()
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

        private void button8_Click_1(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {

                DialogResult result;
                result = MessageBox.Show("是否要寄出", "Yes/No", MessageBoxButtons.YesNo);
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
                    System.Data.DataTable T1 = GetDOWNLOADSAW2();
                    if (T1.Rows.Count > 0)
                    {
                        M1("WH");
                    }

                    System.Data.DataTable T2 = GetDOWNLOADSA();
                    if (T2.Rows.Count > 0)
                    {
                        M1("SALES");
                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

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

                            strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
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
        private void M1(string TYPE)
        {
            try
            {
                System.Data.DataTable M1 = null;
                if (TYPE == "SALES")
                {
                    M1 = GetDOWNLOADSA();
                }
                if (TYPE == "WH")
                {
                    M1 = GetDOWNLOADSAW();
                }

                System.Data.DataTable M2 = GetDOWNLOADSA2();
                if (TYPE == "SALES")
                {
                    if (M2.Rows.Count == 0)
                    {
                        MessageBox.Show("沒有SA資料");
                        return;
                    }
                }

                string SALES = M2.Rows[0]["SALES"].ToString();
                string SA = M2.Rows[0]["SA"].ToString();
                System.Data.DataTable M3 = GetDOWNLOADSA3(SA);


                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                if (TYPE == "WH")
                {
                    FileName = lsAppDir + "\\MailTemplates\\SHIPDW.htm";
                }
                else
                {
                    FileName = lsAppDir + "\\MailTemplates\\SHIPD.htm";
                }

                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();



                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                template = template.Replace("##ETC##", "ETC(結關日) : " + closeDayTextBox.Text);
                template = template.Replace("##ETD##", "ETD(預計開航日) : " + forecastDayTextBox.Text);
                template = template.Replace("##ETA##", "ETA(預計抵達日) : " + arriveDayTextBox.Text);
                if (receiveDayTextBox.Text.Trim().ToUpper() != "TRUCK")
                {
                    template = template.Replace("##HANZ##", "ocean Vessel(港名/航次) : " + boatNameTextBox.Text);
                    template = template.Replace("##SHIPNO##", "Shipping Order No(提單或訂艙編號) : " + soNoTextBox.Text);

                }
                else
                {
                    template = template.Replace("##HANZ##", "");
                    template = template.Replace("##SHIPNO##", "");
                }
                if (TYPE == "WH")
                {
                    dataGridView7.DataSource = GetDOWNLOADWH(shippingCodeTextBox.Text);
                    string MailContent = htmlMessageBody(dataGridView7).ToString();

                    template = template.Replace("##CC##", MailContent);
                }
                //GetDOWNLOADWH
                MailMessage message = new MailMessage();
                if (TYPE == "SALES")
                {
                    if (M3.Rows.Count > 0)
                    {
                        string MSA = M3.Rows[0][0].ToString();
                        //    MSA = "JOYCHEN@ACMEPOINT.COM";
                        //   MSA = "lleytonchen@ACMEPOINT.COM";
                        message.To.Add(MSA);


                        if (!String.IsNullOrEmpty(SALES))
                        {
                            System.Data.DataTable M4 = GetDOWNLOADSA4(SALES);
                            if (M4.Rows.Count > 0)
                            {
                                string MSALES = M4.Rows[0][0].ToString();
                                message.To.Add(MSALES);
                            }

                        }
                    }
                }
                if (TYPE == "WH")
                {

                    //System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHSTOCK();
                    //if (SHIPSTOCCK.Rows.Count > 0)
                    //{
                    //    for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                    //    {
                    //        string MSALES = SHIPSTOCCK.Rows[i][0].ToString();
                    //        message.To.Add(MSALES);
                    //    }
                    //}

                    string MSA = "JOYCHEN@ACMEPOINT.COM";
                    //   string MSA = "lleytonchen@ACMEPOINT.COM";
                    message.To.Add(MSA);

                }
                //   message.CC.Add(fmLogin.LoginID.ToString());
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
                string SUB = "";
                if (TYPE == "SALES")
                {
                    SUB = "Shipping Doc_" + CARDNAME + "_" + receiveDayTextBox.Text + "_" + shippingCodeTextBox.Text;
                }
                else
                {
                    SUB = "打銷_" + CARDNAME + "_" + shippingCodeTextBox.Text + "_" + mEMO3TextBox.Text;
                }
                message.Subject = SUB;
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
                if (TYPE == "SALES")
                {
                    MessageBox.Show("業務寄信成功");
                }
                else
                {
                    MessageBox.Show("倉庫寄信成功");
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE"].Value = GetMenu.DAYTIME();

                    }
                    else
                    {
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE"].Value = "";
                    }
                }

                if (downloadDataGridView.Columns[e.ColumnIndex].Name == "DLCHECK2")
                {

                    string S1 = Convert.ToString(this.downloadDataGridView.Rows[e.RowIndex].Cells["DLCHECK2"].Value);
                    if (S1 == "True")
                    {
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE2"].Value = GetMenu.DAYTIME();

                    }
                    else
                    {
                        this.downloadDataGridView.Rows[e.RowIndex].Cells["DOCDATE2"].Value = "";
                    }
                }
            }
            catch { }
        }

        private void mEMO3TextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {


                string MEMOT = mEMO3TextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("WH20");
                string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                if (G1 != -1)
                {
                    string[] arrurl = H1.Split(new Char[] { ',' });

                    foreach (string i in arrurl)
                    {
                        MEMO = i.Substring(0, 14);
                        WH_main a = new WH_main();
                        a.PublicString = MEMO;
                        a.Show();
                    }

                }
            }
            catch { }
        }

        public void SHIPOWTR()
        {
            try
            {


                StringBuilder sb = new StringBuilder();
                string MEMOT = mEMO3TextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("WH201");

                if (G1 != -1)
                {
                    string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                    string[] arrurl = H1.Split(new Char[] { ',' });

                    foreach (string i in arrurl)
                    {
                        MEMO = i.Substring(0, 14);
                        sb.Append("'" + MEMO + "',");


                    }
                    sb.Remove(sb.Length - 1, 1);
                    string ds = sb.ToString();
                    System.Data.DataTable GF1 = GetOWTR(ds);
                    if (GF1.Rows.Count > 0)
                    {
                        dataGridView6.DataSource = GF1;
                    }
                    else
                    {
                        dataGridView6.DataSource = GetOWTR("'a1234'");
                    }
                }
                else
                {
                    dataGridView6.DataSource = GetOWTR("'a1234'");
                }
            }
            catch { }
        }
        public void SHIPNO()
        {
            /*
            if (dOCTYPETextBox.Text == "銷售訂單")
            {
                mEMO3TextBox.Text = "";
                System.Data.DataTable dt3 = GetShip(shippingCodeTextBox.Text);
                if (dt3.Rows.Count > 0)
                {
                    string ITEMREMARK = dt3.Rows[0]["ITEMREMARK"].ToString();
                    if (ITEMREMARK == "銷售訂單" || ITEMREMARK == "Infinite" || ITEMREMARK == "Choice" || ITEMREMARK == "TOP GARDEN")
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
                        System.Data.DataTable SS = GetShAll(A);
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

            if (dOCTYPETextBox.Text == "採購" || dOCTYPETextBox.Text == "採購單")
            {
                mEMO3TextBox.Text = "";
                System.Data.DataTable dt3OPCH = GetSHIOPCH(shippingCodeTextBox.Text);
                if (dt3OPCH.Rows.Count > 0)
                {
                    StringBuilder sb2 = new StringBuilder();
                    for (int i = 0; i <= dt3OPCH.Rows.Count - 1; i++)
                    {

                        string DOCENTRY = dt3OPCH.Rows[i]["DOCENTRY"].ToString();

                        sb2.Append("'" + DOCENTRY + "',");
                    }

                    sb2.Remove(sb2.Length - 1, 1);
                    string A = sb2.ToString();

                    System.Data.DataTable SS = GetSH2(A, "收貨採購單");

                    if (SS.Rows.Count > 0)
                    {
                        StringBuilder sb3 = new StringBuilder();
                        for (int j = 0; j <= SS.Rows.Count - 1; j++)
                        {
                            string CODE = SS.Rows[j]["CODE"].ToString();

                            sb3.Append(CODE + ",");

                        }
                        sb3.Remove(sb3.Length - 1, 1);
                        mEMO3TextBox.Text = sb3.ToString();
                    }
                    else
                    {
                        System.Data.DataTable dt3OPCH2 = GetSHIOPCH2(shippingCodeTextBox.Text);
                        if (dt3OPCH2.Rows.Count > 0)
                        {
                            StringBuilder sb3 = new StringBuilder();
                            StringBuilder sb4 = new StringBuilder();
                            for (int i = 0; i <= dt3OPCH2.Rows.Count - 1; i++)
                            {
                                string DOCENTRY = dt3OPCH2.Rows[i]["DOCENTRY"].ToString();
                                string LINENUM = dt3OPCH2.Rows[i]["LINENUM"].ToString();
                                sb3.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                            }
                            sb3.Remove(sb3.Length - 1, 1);
                            string A2 = sb3.ToString();
                            System.Data.DataTable SS2 = GetSH(A2, "採購單");
                            if (SS2.Rows.Count > 0)
                            {
                                for (int i = 0; i <= SS2.Rows.Count - 1; i++)
                                {
                                    string CODE = SS2.Rows[i]["CODE"].ToString();

                                    sb4.Append(CODE + ",");

                                }
                                sb4.Remove(sb4.Length - 1, 1);
                                mEMO3TextBox.Text = sb4.ToString();
                            }

                        }
                    }
                }


            }
            if (dOCTYPETextBox.Text == "調撥" || dOCTYPETextBox.Text == "調撥單" || dOCTYPETextBox.Text == "發貨" || dOCTYPETextBox.Text == "AR貸項")
            {
                if (dOCTYPETextBox.Text == "調撥")
                {
                    dOCTYPETextBox.Text = "調撥單";
                }
                if (dOCTYPETextBox.Text == "發貨")
                {
                    dOCTYPETextBox.Text = "發貨單";
                }
                System.Data.DataTable dt3DIAO = GetShip(shippingCodeTextBox.Text);
                if (dt3DIAO.Rows.Count > 0)
                {
                    StringBuilder sb2 = new StringBuilder();
                    StringBuilder sb3 = new StringBuilder();
                    for (int i = 0; i <= dt3DIAO.Rows.Count - 1; i++)
                    {
                        string DOCENTRY = dt3DIAO.Rows[i]["DOCENTRY"].ToString();
                        string LINENUM = dt3DIAO.Rows[i]["LINENUM"].ToString();
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                    }
                    sb2.Remove(sb2.Length - 1, 1);
                    string A = sb2.ToString();
                    if (dOCTYPETextBox.Text == "AR貸項")
                    {
                        dOCTYPETextBox.Text = "AR貸項通知單";
                    }

                    System.Data.DataTable SS = null;

                    if (dOCTYPETextBox.Text == "調撥單")
                    {
                        SS = GetShAll(A);

                    }
                    else
                    {
                        SS = GetShAll(A);
                    }
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
            */
            //合併上面的
            System.Data.DataTable dt3 = GetShip(shippingCodeTextBox.Text);
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            if (dt3.Rows.Count > 0)
            {
                for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                {
                    string ITEMREMARK = dt3.Rows[i]["ITEMREMARK"].ToString();
                    string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString();
                    string LINENUM = dt3.Rows[i]["LINENUM"].ToString();
                    if (ITEMREMARK == "銷售訂單" || ITEMREMARK == "Infinite" || ITEMREMARK == "Choice" || ITEMREMARK == "TOP GARDEN")
                    {
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + ' ' + "1',");//EX:'79310 0 1'用1記錄為採購訂單 以免後續利用DOCENTRY找尋的時候找到重複單號包含調撥單銷售訂單
                    }
                    else if (ITEMREMARK.Contains("採購"))
                    {
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + ' ' + "2',");
                    }
                    else if (ITEMREMARK.Contains("調撥"))
                    {
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + ' ' + "3',"); //EX:'79310 0 1'用1記錄為調撥單 不用調撥單原因是倉管工單的單據類型可能包含其他種類調撥單
                    }
                }


            }
            sb2.Remove(sb2.Length - 1, 1);
            string A = sb2.ToString();
            System.Data.DataTable SS = GetShAll(A);
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
            shipping_mainBindingSource.EndEdit();
            shipping_mainTableAdapter.Update(ship.Shipping_main);
            ship.Shipping_main.AcceptChanges();
        }
        


            






    public System.Data.DataTable GetSHIPDIAO(string SHIPPINGCODE, string itemremark)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM SHIPPING_ITEM T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND  itemremark = @itemremark ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@itemremark", itemremark));
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
        public System.Data.DataTable GetSHIOPCH(string U_SHIPPING_NO)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "SELECT DOCENTRY FROM OPDN WHERE U_SHIPPING_NO=@U_SHIPPING_NO ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_SHIPPING_NO", U_SHIPPING_NO));
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
        private System.Data.DataTable GetSHIOPCH2(string U_SHIPPING_NO)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT T4.DOCENTRY,T4.LINENUM  FROM OPDN T0");
            sb.Append(" INNER JOIN PDN1 T1 ON (T0.docentry=T1.docentry)");
            sb.Append(" INNER join POR1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum )");
            sb.Append("  WHERE T0.U_SHIPPING_NO=@U_SHIPPING_NO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_SHIPPING_NO", U_SHIPPING_NO));


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
        public System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM SHIPPING_ITEM T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMREMARK in ('銷售訂單','銷售單','TOP GARDEN','Infinite','Choice') ";
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
        public System.Data.DataTable GetShip(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM SHIPPING_ITEM T1 WHERE SHIPPINGCODE=@SHIPPINGCODE";
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

        //strCn22
        private System.Data.DataTable GetSH(string DocEntry, string ITEMREMARK)
        {

            SqlConnection connection = null;

            if (ITEMREMARK == "Infinite")
            {
                connection = new SqlConnection(strINF);
                ITEMREMARK = "銷售單";
            }
            else if (ITEMREMARK == "Choice")
            {
                connection = new SqlConnection(strCHO);
                ITEMREMARK = "銷售單";
            }
            else if (ITEMREMARK == "TOP GARDEN")
            {
                connection = new SqlConnection(strTOP);
                ITEMREMARK = "銷售單";
            }
            else
            {
                connection = globals.Connection;

            }

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from WH_item4 T0 left join wh_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK=@ITEMREMARK AND   ISNULL(soNo,'') <>'Checked'  ");


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
        private System.Data.DataTable GetShAll(string DocEntry)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from WH_item4 T0 left join wh_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where (ItemRemark like '%銷售訂單%' and cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar)+' '+ '1' IN (" + DocEntry + ")) ");
            sb.Append("  or (ItemRemark like '%採購%' and cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar)+' '+ '2' IN (" + DocEntry + ") ) ");
            sb.Append("  or (ItemRemark like '%調撥%' and cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar)+' '+ '3' IN (" + DocEntry + ") )");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



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
        private System.Data.DataTable GetSH3(string DocEntry, string ITEMREMARK)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from WH_item4 T0 left join wh_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK like '%" + ITEMREMARK + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



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
        private System.Data.DataTable GetSH2(string DocEntry, string ITEMREMARK)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from WH_item4 T0 left join wh_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK=@ITEMREMARK ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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



        private void download2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (download2DataGridView.Columns[e.ColumnIndex].Name == "MARK2")
            {
                DataGridViewRow row;
                string G1 = this.download2DataGridView.Rows[e.RowIndex].Cells["MARK2"].Value.ToString();
                string filename = this.download2DataGridView.Rows[e.RowIndex].Cells["filename"].Value.ToString();
                string path = this.download2DataGridView.Rows[e.RowIndex].Cells["path2"].Value.ToString();


                if (G1 == "1")
                {

                    if (!String.IsNullOrEmpty(filename))
                    {
                        int T = filename.IndexOf(".");
                        string a0 = filename.Substring(0, T);
                        if (a0.Length > 26)
                        {
                            string T1 = a0.Substring(13, T - 13);
                            add9TextBox.Text = T1;
                        }
                        else
                        {
                            add9TextBox.Text = a0;
                        }

                        for (int i = 0; i <= dataGridView5.Rows.Count - 1; i++)
                        {
                            row = dataGridView5.Rows[i];
                            string A1 = row.Cells["併單工單"].Value.ToString();
                            string AS = row.Cells["類型"].Value.ToString().Trim();

                            if (AS == "併單")
                            {
                                System.Data.DataTable H1 = GetSHPCAR4(A1, shippingCodeTextBox.Text);

                                if (H1.Rows.Count > 0)
                                {
                                    for (int s = 0; s <= H1.Rows.Count - 1; s++)
                                    {
                                        string JOBNO = H1.Rows[s][0].ToString();
                                        string SEQ = "0";
                                        System.Data.DataTable H2 = GetDOWNLOAD2SEQ(JOBNO);
                                        if (H2.Rows.Count > 0)
                                        {
                                            SEQ = H2.Rows[0][0].ToString();
                                        }

                                        InsertDownload2(JOBNO, SEQ, filename, path);
                                        UPDATEADD9(filename.Substring(0, T), JOBNO);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    add9TextBox.Text = "";


                    for (int i = 0; i <= dataGridView5.Rows.Count - 1; i++)
                    {
                        row = dataGridView5.Rows[i];
                        string A1 = row.Cells["併單工單"].Value.ToString();
                        string AS = row.Cells["類型"].Value.ToString().Trim();

                        if (AS == "併單")
                        {
                            System.Data.DataTable H1 = GetSHPCAR4(A1, shippingCodeTextBox.Text);

                            if (H1.Rows.Count > 0)
                            {
                                for (int s = 0; s <= H1.Rows.Count - 1; s++)
                                {
                                    string JOBNO = H1.Rows[s][0].ToString();

                                    DELETEDownload2(JOBNO, filename);
                                    UPDATEADD9("", JOBNO);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void download2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void cFSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (cFSCheckBox.Checked == false)
            {
                eNDCHECKCheckBox.Checked = false;
            }
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

            System.Data.DataTable dt3 = util.GetSHIPPACK();

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

                    System.Data.DataTable H1S = util.GetSHIPPACK4S(shippingCodeTextBox.Text, ITEMCODE, ITEMNAME);

                    System.Data.DataTable H1S2 = util.GetSHIPPACK4S2(sbS.ToString(), ITEMNAME, ITEMCODE);
                    if (H1S.Rows.Count > 0)
                    {

                        string S1 = H1S.Rows[0][0].ToString().Trim();
                        drw2["DescGoods"] = S1;

                        if (H1S.Rows.Count > 1)
                        {

                            System.Data.DataTable H1SQ = util.GetSHIPPACK4SQTY(shippingCodeTextBox.Text, ITEMCODE, QQ);
                            if (H1SQ.Rows.Count > 0)
                            {
                                string S1Q = H1SQ.Rows[0][0].ToString().Trim();
                                drw2["DescGoods"] = S1Q;
                            }
                        }

                    }

                    int ACME = ITEMCODE.IndexOf("ACME");
                    if (String.IsNullOrEmpty(drw2["DescGoods"].ToString()))
                    {
                        if (ACME != -1)
                        {
                            System.Data.DataTable H2S = util.GetSHIPPS3(shippingCodeTextBox.Text, ITEMCODE);
                            if (H2S.Rows.Count == 0)
                            {
                                System.Data.DataTable H3S = util.GetSHIPPS4(shippingCodeTextBox.Text, ITEMCODE);
                                if (H3S.Rows.Count > 0)
                                {
                                    drw2["DescGoods"] = H3S.Rows[0][0].ToString().Trim();
                                }
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(drw2["DescGoods"].ToString()))
                    {


                        System.Data.DataTable H1 = util.GetSHIPPACK3(ITEMCODE);
                        if (H1.Rows.Count > 0)
                        {
                            string MODE = H1.Rows[0][0].ToString().Trim();
                            string GRADE = H1.Rows[0][1].ToString().Trim();
                            if (MODE.Length > 13)
                            {
                                MODE = MODE.Substring(1, 13);
                            }
                            System.Data.DataTable H2O = util.GetSHIPPACK4O(shippingCodeTextBox.Text, MODE, GRADE);
                            System.Data.DataTable H2 = util.GetSHIPPACK4(shippingCodeTextBox.Text, MODE);
                            if (H2O.Rows.Count > 0)
                            {
                                string DESC = H2O.Rows[0][0].ToString().Trim();

                                drw2["DescGoods"] = DESC;
                            }

                            else
                            {
                                if (H1S2.Rows.Count > 0)
                                {
                                    for (int F = 0; F <= H1S2.Rows.Count - 1; F++)
                                    {
                                        string HS = H1S2.Rows[F][0].ToString();
                                        System.Data.DataTable H1S3 = util.GetSHIPPACK4S(shippingCodeTextBox.Text, HS, ITEMNAME);
                                        if (H1S3.Rows.Count > 0)
                                        {
                                            drw2["DescGoods"] = H1S3.Rows[0][0].ToString().Trim();
                                        }
                                    }
                                }

                            }

                            if (String.IsNullOrEmpty(drw2["DescGoods"].ToString()))
                            {
                                if (H2.Rows.Count > 0)
                                {
                                    drw2["DescGoods"] = H2.Rows[0][0].ToString().Trim();
                                }
                                else
                                {
                                    drw2["DescGoods"] = H1.Rows[0][0].ToString().Trim();
                                }
                            }


                        }
                    }

                    System.Data.DataTable OI12 = util.GetSHIPOITMES();
                    if (OI12.Rows.Count > 0)
                    {
                        System.Data.DataTable OI1 = util.GetSHIPOITM(ITEMCODE);
                        if (OI1.Rows.Count > 0)
                        {
                            string OINAME = "";
                            string MODEL = OI1.Rows[0][0].ToString();
                            string GRADE = OI1.Rows[0][1].ToString();

                            string OIES = OI1.Rows[0][2].ToString();
                            string TMODEL = OI1.Rows[0][3].ToString().Trim();
                            if (kPIYESNOCheckBox.Checked)
                            {
                                OINAME = MODEL + GRADE;
                            }
                            else
                            {
                                OINAME = MODEL;

                            }
                            if (cardCodeTextBox.Text == "1279-03")
                            {
                                System.Data.DataTable OI2 = GetSHIPOITM2(TMODEL);
                                if (OI2.Rows.Count > 0)
                                {
                                    System.Data.DataTable OI3 = GetSHIPOITM3(ITEMCODE);
                                    if (OI3.Rows.Count > 0)
                                    {
                                        OINAME = OINAME + OI3.Rows[0][0].ToString();
                                    }
                                }
                            }
                            if (!String.IsNullOrEmpty(OIES))
                            {
                                OIES = " (" + ES + ")";

                            }
                            drw2["DescGoods"] = OINAME + OIES;
                        }
                    }

                    if (SER.Trim() != "0")
                    {

                        GV++;
                        if (GV == 1)
                        {
                            System.Data.DataTable dt31 = util.GetSHIPPACK2(SER);
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
                            System.Data.DataTable dt31 = util.GetSHIPPACK5(SER);
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
                            System.Data.DataTable G11 = util.GetSHIPPACKQTY(ITEMCODE);
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
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable dt = GetSHIPPACKINV();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    DataRow dd = dt.Rows[i];
                    sb.Append(dd["INV"].ToString() + ",");
                }

                sb.Remove(sb.Length - 1, 1);
                memoTextBox.Text = "REF NO.:" + sb.ToString();
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
        private void WriteExcelPACK(string SEQ, string CHE, string CAR, string CHOSHIP)
        {

            util.DELPACK();

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
            System.Data.DataTable dt3 = util.GetWHPACK(SHIPPINGCODE, BLC, CHE, sb.ToString(), CAR);
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
                        System.Data.DataTable H1 = util.GetSHIPPACK9(WHNO, PLATENO2);
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
                                util.UPPACKS(SERX);
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
                            System.Data.DataTable G1 = util.GetSHIPPACK6(ITEMCODE, QTY);
                            if (G1.Rows.Count > 0)
                            {
                                NW = G1.Rows[0][0].ToString();
                            }
                            else
                            {
                                System.Data.DataTable G2 = util.GetSHIPPACK7(ITEMCODE);
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
                        util.AddPACK(PLATENO, CARTONNO, ITEMCODE, QTY, CARTONQTY, NW, GW, L, W, H, LOACTION, SERX, CARTONNO3, INVOICE, ITEMNAME, WHNO, ES);
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
        public void UPWHPACK(string SEQNO, string ID)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2  SET SEQNO=@SEQNO  WHERE ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));
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


        string HW = "";
        private void button36_Click(object sender, EventArgs e)
        {
            if (HW == "")
            {
                panel3.Hide();
                HW = "1";
            }
            else
            {
                panel3.Show();
                HW = "";
            }

        }

        private void dataGridView6_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案名稱6")
                {


                    string kk = dataGridView6.CurrentRow.Cells["路徑6"].Value.ToString();

                    string aa = dataGridView6.CurrentRow.Cells["path6"].Value.ToString() + "\\" + kk;


                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + kk;

                    int kk1 = kk.ToLower().IndexOf(".msg");
                    if (kk1 != -1)
                    {
                        System.Diagnostics.Process.Start(aa);
                    }
                    else
                    {
                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);
                    }




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

        private void button23_Click_1(object sender, EventArgs e)
        {
            Clear(sbS);
            SBS();
            System.Data.DataTable dt3H = GetWHPACKH(sbS.ToString());

            if (dt3H.Rows.Count > 0)
            {

                System.Data.DataTable dth = ship.Mark;
                for (int j = 0; j <= dt3H.Rows.Count - 1; j++)
                {

                    DataRow drw2 = dth.NewRow();

                    drw2["Seq"] = j.ToString();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["Mark"] = dt3H.Rows[j][0].ToString();
                    dth.Rows.Add(drw2);

                }

            }
            else
            {
                System.Data.DataTable dt3HD = GetWHPACKH2(sbS.ToString());

                if (dt3HD.Rows.Count > 0)
                {
                    System.Data.DataTable dth = ship.Mark;
                    string MARK = dt3HD.Rows[0][0].ToString();
                    string[] NewLine = MARK.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    int G = 0;
                    foreach (string ESi in NewLine)
                    {
                        G++;
                        DataRow drw2 = dth.NewRow();

                        drw2["Seq"] = G.ToString();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["Mark"] = ESi.ToString();
                        dth.Rows.Add(drw2);
                    }
                }
            }
        }
        public System.Data.DataTable GetWHPACKH2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT PACKMEMO  FROM WH_MAIN WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  ) AND PACKMEMO <> '' ");
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



        private void button8_Click_2(object sender, EventArgs e)
        {
            if (dOCTYPETextBox.Text == "銷售訂單")
            {
                System.Data.DataTable G1 = null;

                int T1 = add1TextBox.Text.IndexOf("正航系統");
                if (T1 == -1)
                {

                    if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "船務測試區")
                    {
                        G1 = GERCARD(pinoTextBox.Text.Trim());
                        if (G1.Rows.Count > 0)
                        {
                            cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                            cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                            string 貿易條件 = G1.Rows[0]["貿易條件"].ToString();
                            string 收貨地 = G1.Rows[0]["收貨地"].ToString();
                            string 目的地 = G1.Rows[0]["目的地"].ToString();

                            System.Data.DataTable GET1 = GERCARD1(貿易條件);
                            if (GET1.Rows.Count == 1)
                            {
                                tradeConditionTextBox.Text = GET1.Rows[0][0].ToString();
                            }
                            System.Data.DataTable GET2 = GERCARD3(收貨地, "收貨地");
                            if (GET2.Rows.Count == 1)
                            {
                                receivePlaceTextBox.Text = GET2.Rows[0][0].ToString();
                                shipmentTextBox.Text = GET2.Rows[0][0].ToString();
                            }
                            System.Data.DataTable GET3 = GERCARD3(目的地, "目的地");
                            if (GET3.Rows.Count == 1)
                            {
                                goalPlaceTextBox.Text = GET3.Rows[0][0].ToString();
                                unloadCargoTextBox.Text = GET3.Rows[0][0].ToString();
                            }

                            System.Data.DataTable GET4 = GERCARDD(pinoTextBox.Text.Trim());
                            if (GET4.Rows.Count > 0)
                            {
                                boardCountNoTextBox.Text = GET4.Rows[0][0].ToString();
                            }
                            receiveDayTextBox.Text = G1.Rows[0]["運輸方式"].ToString();



                            dOCTYPETextBox.Text = "銷售";
                            button26_Click(sender, e, "1");
                        }
                    }
                }

                else
                {
                    string ITEMREMARK = "";
                    if (add1TextBox.Text == "正航系統CHOICE")
                    {
                        ITEMREMARK = "Choice";
                    }
                    //正航系統INFINITE
                    if (add1TextBox.Text == "正航系統INFINITE")
                    {
                        ITEMREMARK = "Infinite";
                    }
                    G1 = GERCARD2(pinoTextBox.Text.Trim(), ITEMREMARK);
                    if (G1.Rows.Count > 0)
                    {
                        cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                        cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                        dOCTYPETextBox.Text = "銷售";

                        button15_Click(sender, e, ITEMREMARK, "", "1");
                    }


                }



            }
            else if (dOCTYPETextBox.Text == "調撥單")
            {
                System.Data.DataTable G1 = null;

                int T1 = add1TextBox.Text.IndexOf("正航系統");
                if (T1 == -1)
                {

                    if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "船務測試區")
                    {
                        G1 = GERCARDOWTR(pinoTextBox.Text.Trim());
                        if (G1.Rows.Count > 0)
                        {
                            cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                            cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                            string 貿易條件 = G1.Rows[0]["貿易條件"].ToString();
                            string 收貨地 = G1.Rows[0]["收貨地"].ToString();
                            string 目的地 = G1.Rows[0]["目的地"].ToString();

                            System.Data.DataTable GET1 = GERCARD1(貿易條件);
                            if (GET1.Rows.Count == 1)
                            {
                                tradeConditionTextBox.Text = GET1.Rows[0][0].ToString();
                            }
                            System.Data.DataTable GET2 = GERCARD3(收貨地, "收貨地");
                            if (GET2.Rows.Count == 1)
                            {
                                receivePlaceTextBox.Text = GET2.Rows[0][0].ToString();
                                shipmentTextBox.Text = GET2.Rows[0][0].ToString();
                            }
                            System.Data.DataTable GET3 = GERCARD3(目的地, "目的地");
                            if (GET3.Rows.Count == 1)
                            {
                                goalPlaceTextBox.Text = GET3.Rows[0][0].ToString();
                                unloadCargoTextBox.Text = GET3.Rows[0][0].ToString();
                            }

                            System.Data.DataTable GET4 = GERCARDD(pinoTextBox.Text.Trim());
                            if (GET4.Rows.Count > 0)
                            {
                                boardCountNoTextBox.Text = GET4.Rows[0][0].ToString();
                            }
                            receiveDayTextBox.Text = G1.Rows[0]["運輸方式"].ToString().ToUpper();



                            dOCTYPETextBox.Text = "調撥單";
                            button8_Click(sender, e);
                        }
                    }
                }

                else
                {
                    string ITEMREMARK = "";
                    if (add1TextBox.Text == "正航系統CHOICE")
                    {
                        ITEMREMARK = "Choice";
                    }
                    //正航系統INFINITE
                    if (add1TextBox.Text == "正航系統INFINITE")
                    {
                        ITEMREMARK = "Infinite";
                    }
                    G1 = GERCARD2(pinoTextBox.Text.Trim(), ITEMREMARK);
                    if (G1.Rows.Count > 0)
                    {
                        cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                        cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                        dOCTYPETextBox.Text = "調撥單";

                        button15_Click(sender, e, ITEMREMARK, "", "1");
                    }


                }



            }
        }


        private void sAMEMOTextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {


                string MEMOT = sAMEMOTextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("SI20");
                int G2 = MEMOT.IndexOf("SH20");
                if (G1 != -1 || G2 != -1)
                {
                    string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);

                    string[] arrurl = H1.Split(new Char[] { '+', '、', '/', ',' });

                    foreach (string str in arrurl)
                    {
                        int T1 = str.IndexOf("SH20");
                        int T2 = str.IndexOf("SI20");
                        MEMO = T1 != -1 ? str.Substring(T1, 14) : str.Substring(T2, 14);
                        if (T1 != -1 && (H1.Contains("+") || H1.Contains("、") || H1.Contains("/")))
                        {
                            fmShip a = new fmShip();
                            a.PublicString = MEMO;
                            a.Show();
                        }

                        if (T2 != -1 && (H1.Contains("+") || H1.Contains("、") || H1.Contains("/")))
                        {
                            APShip a = new APShip();
                            a.PublicString = MEMO;
                            a.Show();
                        }
                    }

                }
            }
            catch (Exception ex) { }
        }


        public void AddOCLG(int ClgCode, string CardCode, DateTime CntctDate, int CntctTime, DateTime Recontact, string Closed, string Tel, int CntctSbjct, string Transfered, string DocType, string DocNum, string DocEntry, string Attachment, string DataSource, int AttendUser, int CntctCode, int UserSign, int SlpCode, string Action, int CntctType, int BeginTime, string DurType, string Priority, string Reminder, int RemQty, string RemType, string RemSented, int Instance, string personal, string inactive, string tentative, int AtcEntry, DateTime endDate, int ENDTime)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into OCLG(ClgCode,CardCode,CntctDate,CntctTime,Recontact,Closed,Tel,CntctSbjct,Transfered,DocType,DocNum,DocEntry,Attachment,DataSource,AttendUser,CntctCode,UserSign,SlpCode,Action,CntctType,BeginTime,DurType,Priority,Reminder,RemQty,RemType,RemSented,Instance,personal,inactive,tentative,AtcEntry,duration,endDate,ENDTime) values(@ClgCode,@CardCode,@CntctDate,@CntctTime,@Recontact,@Closed,@Tel,@CntctSbjct,@Transfered,@DocType,@DocNum,@DocEntry,@Attachment,@DataSource,@AttendUser,@CntctCode,@UserSign,@SlpCode,@Action,@CntctType,@BeginTime,@DurType,@Priority,@Reminder,@RemQty,@RemType,@RemSented,@Instance,@personal,@inactive,@tentative,@AtcEntry,@duration,@endDate,@ENDTime)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ClgCode", ClgCode));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@CntctDate", CntctDate));
            command.Parameters.Add(new SqlParameter("@CntctTime", CntctTime));
            command.Parameters.Add(new SqlParameter("@Recontact", Recontact));
            command.Parameters.Add(new SqlParameter("@Closed", Closed));
            command.Parameters.Add(new SqlParameter("@Tel", Tel));
            command.Parameters.Add(new SqlParameter("@CntctSbjct", CntctSbjct));
            command.Parameters.Add(new SqlParameter("@Transfered", Transfered));
            command.Parameters.Add(new SqlParameter("@DocType", DocType));
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@Attachment", Attachment));
            command.Parameters.Add(new SqlParameter("@DataSource", DataSource));
            command.Parameters.Add(new SqlParameter("@AttendUser", AttendUser));
            command.Parameters.Add(new SqlParameter("@CntctCode", CntctCode));
            command.Parameters.Add(new SqlParameter("@UserSign", UserSign));
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            command.Parameters.Add(new SqlParameter("@Action", Action));
            command.Parameters.Add(new SqlParameter("@CntctType", CntctType));
            command.Parameters.Add(new SqlParameter("@BeginTime", BeginTime));
            command.Parameters.Add(new SqlParameter("@DurType", DurType));
            command.Parameters.Add(new SqlParameter("@Priority", Priority));
            command.Parameters.Add(new SqlParameter("@Reminder", Reminder));
            command.Parameters.Add(new SqlParameter("@RemQty", RemQty));
            command.Parameters.Add(new SqlParameter("@RemType", RemType));
            command.Parameters.Add(new SqlParameter("@RemSented", RemSented));
            command.Parameters.Add(new SqlParameter("@Instance", Instance));
            command.Parameters.Add(new SqlParameter("@personal", personal));
            command.Parameters.Add(new SqlParameter("@inactive", inactive));
            command.Parameters.Add(new SqlParameter("@tentative", tentative));
            command.Parameters.Add(new SqlParameter("@AtcEntry", AtcEntry));
            command.Parameters.Add(new SqlParameter("@duration", 1.000000));
            command.Parameters.Add(new SqlParameter("@endDate", endDate));
            command.Parameters.Add(new SqlParameter("@ENDTime", ENDTime));

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

        private void lcInstro1DataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string r = lcInstro1DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            System.Data.DataTable G1 = GetMenu.GetAPLC(r);
            if (G1.Rows.Count > 0)
            {

                APLC a = new APLC();
                a.PublicString = G1.Rows[0][0].ToString();
                a.Show();

            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("收件人地址為" + textBox2.Text + "是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {

                HK();
                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\SHIP.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                template = template.Replace("##A1##", "請依附件做香港出口報關。");

                template = template.Replace("##shippingCode##", "JOB NO: " + shippingCodeTextBox.Text);
                template = template.Replace("##soNo##", "Shipping Order No : " + soNoTextBox.Text);
                template = template.Replace("##tradeCondition##", "貿易條件 : " + tradeConditionTextBox.Text);
                template = template.Replace("##closeDay##", "ETC : " + closeDayTextBox.Text);
                template = template.Replace("##forecastDay##", "ETD : " + forecastDayTextBox.Text);
                template = template.Replace("##arriveDay##", "ETA : " + arriveDayTextBox.Text);

                template = template.Replace("##receivePlace##", "RCCEIPT : " + receivePlaceTextBox.Text);
                template = template.Replace("##goalPlace##", "DESTNATION : " + goalPlaceTextBox.Text);

                template = template.Replace("##boatName##", "FLIGHT NO. : " + boatNameTextBox.Text);
                template = template.Replace("##shipment##", "LOADING PORT : " + shipmentTextBox.Text);
                template = template.Replace("##unloadCargo##", "DISCHARGE PORT : " + unloadCargoTextBox.Text);
                template = template.Replace("##receiveDay##", "SHIPPED VIA : " + receiveDayTextBox.Text);



                MailMessage message = new MailMessage();

                string aa = textBox2.Text;

                message.To.Add(new MailAddress(aa));

                message.Subject = "香港進出口報關通知單" + "-客戶名稱 " + cardNameTextBox.Text + "-數量 " + download22(shippingCodeTextBox.Text).Rows[0][0].ToString() + "-工單號 " + shippingCodeTextBox.Text;
                message.Body = template;

                //格式為 Html
                message.IsBodyHtml = true;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {

                    string m_File = "";

                    m_File = file;
                    data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                    //附件资料
                    ContentDisposition disposition = data.ContentDisposition;


                    // 加入邮件附件
                    message.Attachments.Add(data);

                }
                string F = cardCodeTextBox.Text.Substring(0, 1);

                if (F != "S")
                {
                    System.Data.DataTable GG1 = util.download21(shippingCodeTextBox.Text);
                    if (GG1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= GG1.Rows.Count - 1; i++)
                        {
                            string PATH = GG1.Rows[i][0].ToString();
                            string m_File = "";

                            m_File = PATH;
                            data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                            //附件资料
                            ContentDisposition disposition = data.ContentDisposition;


                            // 加入邮件附件
                            message.Attachments.Add(data);
                        }
                    }
                }

                SmtpClient client = new SmtpClient();
                try
                {
                    client.Send(message);
                    data.Dispose();
                    message.Attachments.Dispose();
                    MessageBox.Show("寄信成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }


            }
        }
        private void UPLOAD(string FILE)
        {
            string PATH = @"\\acmesrv01\SAP_Share\shipping\";
            string DIR = "//acmesrv01//SAP_Share//shipping//";
            string dd = DateTime.Now.ToString("yyyyMM");
            string server = DIR + dd + "//";

            string filename = Path.GetFileName(FILE);
            System.Data.DataTable dt2 = GetMenu.download2(filename);

            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("檔案已上傳過");
            }
            else
            {

                string file = FILE;
                bool FF1 = getrma.UploadFile(file, server, false);
                if (FF1 == false)
                {
                    return;
                }
                System.Data.DataTable GG1 = download2(shippingCodeTextBox.Text);
                string SEQ = GG1.Rows[0][0].ToString();
                string de = DateTime.Now.ToString("yyyyMM") + "\\";
                INSERTDOWNLOAD2(shippingCodeTextBox.Text, SEQ, filename, PATH + de + filename);

            }

        }
        public static System.Data.DataTable download2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select ISNULL(MAX(CAST(SEQ AS INT))+1,0) SEQ  from download2 WHERE SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public void INSERTDOWNLOAD2(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" Insert into Download2(shippingcode,seq,filename,path,STATUS) values(@shippingcode,@seq,@filename,@path,@STATUS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@STATUS", "嘜頭"));

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
        public void AddFEE()
        {
            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into SHIP_FEE(ShippingCode,bSP) values(@ShippingCode,'Unchecked')", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));



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
        public static System.Data.DataTable download22(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SUM(ISNULL(QUANTITY,0)) QTY FROM Shipping_Item WHERE SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void HK()
        {
            DELETEFILE();
            System.Data.DataTable OrderData = GetHK(shippingCodeTextBox.Text);
            if (OrderData.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string F = cardCodeTextBox.Text.Substring(0, 1);

                if (F == "S")
                {

                    FileName = lsAppDir + "\\Excel\\SHIP\\香港進口報關通知單.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\SHIP\\香港出口報關通知單.xlsx";
                }

                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutputLEMONFIT(OrderData, ExcelTemplate, OutPutFile, "N");
                UPLOAD(OutPutFile);
            }
        }



        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            boatCompanyTextBox.Text = comboBox9.Text;

        }

        private void comboBox9_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("REBECCA");

            comboBox9.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox9.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            DDF2(shippingCodeTextBox.Text);
            System.Data.DataTable G1 = util.GETPACL2F(shippingCodeTextBox.Text);
            for (int ii = 0; ii <= G1.Rows.Count - 1; ii++)
            {

                int QQTY = Convert.ToInt32(G1.Rows[ii]["QTY"]);
                decimal NW = Convert.ToDecimal(G1.Rows[ii]["NW"]);
                decimal GW = Convert.ToDecimal(G1.Rows[ii]["GW"]);
                string SAYTOTAL = G1.Rows[ii]["SAYTOTAL"].ToString();
                string CBM = G1.Rows[ii]["CBM"].ToString();
                string PLT = G1.Rows[ii]["PLT"].ToString();
                string CARTON = G1.Rows[ii]["CARTON"].ToString();
                string CBMM = G1.Rows[ii]["CBMM"].ToString();
                string PARTNO = G1.Rows[ii]["PARTNO"].ToString();
                System.Data.DataTable INV2 = util.GetITEMNAME(shippingCodeTextBox.Text, PARTNO);
                string ITEMNAME2 = "";
                if (INV2.Rows.Count > 0)
                {

                    ITEMNAME2 = INV2.Rows[0]["ITEMNAME"].ToString();

                }
                else
                {
                    System.Data.DataTable INV3 = util.GetITEMNAME2(shippingCodeTextBox.Text);
                    if (INV3.Rows.Count > 0)
                    {
                        ITEMNAME2 = INV3.Rows[0]["ITEMNAME"].ToString();
                    }
                }
                string aa2 = invoiceNoTextBox.Text + "-" + invoiceNo_seqTextBox.Text;
                util.AddPACKD(shippingCodeTextBox.Text, aa2, ii.ToString(), PLT, CARTON, ITEMNAME2, QQTY.ToString(), NW.ToString(), GW.ToString(), CBMM, "True");

            }


            memoTextBox.Text = "INV NO.:" + textBox20.Text;
            packingListMTableAdapter.Fill(ship.PackingListM, MyID);
            packingListDTableAdapter.Fill(ship.PackingListD, MyID);
        }

        private void DDF2(string SHIPPINGCODE)
        {

            System.Data.DataTable G1 = util.GETPACL(SHIPPINGCODE);
            if (G1.Rows.Count > 0)
            {
                string FINV = G1.Rows[0]["INVOICENO"].ToString();
                for (int i = 0; i <= G1.Rows.Count - 1; i++)
                {

                    string SIZE = G1.Rows[i]["版數"].ToString();
                    string InvoiceNo = G1.Rows[i]["INVOICENO"].ToString();
                    string LB = InvoiceNo.Substring(0, 2);
                    if (LB == "LB")
                    {
                        System.Data.DataTable M1 = util.GETPACLD(InvoiceNo);

                        if (M1.Rows.Count > 0)
                        {
                            for (int i2 = 0; i2 <= M1.Rows.Count - 1; i2++)
                            {
                                string CCBM = M1.Rows[i2]["CBM"].ToString();

                                string[] sArray = CCBM.Split('*');
                                int F2 = 0;
                                foreach (string F in sArray)
                                {
                                    F2++;
                                }
                                if (F2 > 3)
                                {
                                    int D = CCBM.LastIndexOf("*");
                                    string CC = CCBM.Substring(0, D);
                                    string PLT = sArray[3];


                                    util.UPDATEPACKLB(CC, PLT, InvoiceNo, CCBM);
                                }
                            }
                        }
                        string CBMM = "";
                        System.Data.DataTable GF2 = util.GETPACLS2(InvoiceNo);
                        if (GF2.Rows.Count > 0)
                        {
                            CBMM = GF2.Rows[0][0].ToString();
                        }
                        util.GETCBM(InvoiceNo, CBMM);
                    }
                    else
                    {
                        string[] splitStr = { "CM" };
                        string[] arrurl = SIZE.Split(splitStr, StringSplitOptions.RemoveEmptyEntries);
                        string PLT = "";
                        foreach (string ESi in arrurl)
                        {
                            string[] arrurl2 = ESi.Split(new Char[] { '@' });
                            int F = 0;
                            string PLATENO = "";
                            string CBM = "";

                            foreach (string ESi2 in arrurl2)
                            {
                                F++;
                                string EA = ESi2;
                                if (F == 1)
                                {
                                    PLATENO = EA.Replace(":", "").Replace("No.", "").Trim();
                                }
                                if (F == 2)
                                {
                                    CBM = EA;

                                }
                            }

                            int pall = PLATENO.IndexOf("PALLET");
                            if (pall != -1)
                            {
                                System.Data.DataTable GF1 = util.GETPACLS(PLATENO);
                                if (GF1.Rows.Count > 0)
                                {
                                    PLT = GF1.Rows[0][0].ToString();
                                }
                            }
                            else
                            {
                                PLT = "0";
                            }

                            util.UPDATEPACK(CBM, PLT, InvoiceNo, PLATENO);

                        }
                        string CBMM = "";
                        if (PLT != "0")
                        {
                            System.Data.DataTable GF2 = util.GETPACLS2(InvoiceNo);
                            if (GF2.Rows.Count > 0)
                            {
                                CBMM = GF2.Rows[0][0].ToString();
                            }
                            util.GETCBM(InvoiceNo, CBMM);
                        }
                    }
                }

            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            CalcTotals1();

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\INVO2.xls";
            string OHEM = fmLogin.LoginID.ToString().ToUpper();
            //if (OHEM == "NANCYTSAI")
            //{
            //    FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACME.xls";
            //    GetExcelProduct2(FileName, GetOrderDatCUST(), "Y");
            //}
            //else
            //{
            //    GetExcelProduct2(FileName, GetOrderDatCUST(), "N");
            //}
            FileName = lsAppDir + "\\Excel\\DRS\\INVODRSACME.xls";
            GetExcelProduct2(FileName, GetOrderDatCUST(), "Y");
        }
        private System.Data.DataTable GetOrderDatCUST()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT a.shippingcode JOBNO,a.InvoiceNo+'-'+a.Invoiceno_seq as InvoiceNo,''''+a.[PIno] PIno,''''+a.[POno] as pono,'BILL TO:'+a.[billTo] as billTo,'SHIP TO:'+a.[shipTo] as shipTo,a.[Invoice_memo] as memo,'Ship via : '+a.[InvoiceShip] as InvoiceShip,a.[InvoiceFrom],Convert(varchar(10),Getdate(),111) as 日期 ");
            sb.Append(" ,a.[InvoiceTo],a.[AmountTotal],a.[AmountTotalEng] as AmountTotalEng,b.[SeqNo],b.[MarkNos], ");
            if (GetINVMARK().Rows.Count == 0)
            {
                sb.Append(" cast(ISNULL(seqno2,0)+1 as varchar)+')'+b.[INDescription] as INDescription");
            }
            else
            {
                sb.Append(" CASE WHEN ISNULL(MARKNOS,'') <> 'True' THEN b.[INDescription]  ELSE cast(ISNULL(seqno2,0)+1 as varchar)+')'+b.[INDescription] END INDescription ");
            }
            sb.Append(" ,CAST(b.[InQty] AS VARCHAR)  InQty,c.brand +' BRAND' as BRAND,c.TradeCondition as Trade,");
            sb.Append(" CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[UnitPrice] AS Money),2) UnitPrice ");
            sb.Append(" ,CASE ISNULL(B.CURRENCY,'USD') WHEN 'USD' THEN 'US$'  WHEN '' THEN 'US' ELSE B.CURRENCY END+CONVERT(NVARCHAR(20),CAST(b.[Amount] AS Money),2) Amount ");
            sb.Append(" FROM [InvoiceM] as a  ");
            sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq) ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)  ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSCN D ON (B.ITEMCODE =D.ItemCode  COLLATE  Chinese_Taiwan_Stroke_CI_AS  AND C.CardCode =D.CardCode COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT  distinct '','' InvoiceNo,'' PIno,''pono,'' billTo,'' shipTo,'' memo,'' InvoiceShip,'','' ");
            sb.Append(" ,'', '', '', b.[SeqNo]+'.1', '',  ");
            sb.Append(" CASE WHEN ISNULL(D.Substitute,'')='' THEN  '  P/N:'+F.ITEMCODES ELSE  '  P/N:'+ D.Substitute END COLLATE  Chinese_Taiwan_Stroke_CI_AS INDescription ");
            sb.Append(" ,'', '', '', '','' ");
            sb.Append(" FROM [InvoiceM] as a   ");
            sb.Append(" left join [InvoiceD] as b on(a.shippingcode=b.shippingcode and a.InvoiceNo=b.InvoiceNo and a.InvoiceNo_seq=b.InvoiceNo_seq)  ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSCN D ON (B.ITEMCODE =D.ItemCode  COLLATE  Chinese_Taiwan_Stroke_CI_AS  AND C.CardCode =D.CardCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" LEFT JOIN (");
            sb.Append(" SELECT  SHIPPINGCODE,T0.ITEMCODE,MAX(U_CUSTITEMCODE) ITEMCODES  FROM shipping_ITEM T0 LEFT JOIN AcmeSql02.DBO.RDR1  T1 ON (CAST(T0.Docentry AS varchar) = CAST(T1.DocEntry AS varchar)  AND T0.linenum =T1.LineNum)");
            sb.Append(" WHERE ISNULL(U_CUSTITEMCODE,'') <> ''");
            sb.Append(" GROUP BY SHIPPINGCODE,T0.ITEMCODE ) F ON (B.ITEMCODE =F.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS AND F.ShippingCode =A.ShippingCode)");
            sb.Append(" where a.shippingcode=@shippingcode and a.InvoiceNo=@InvoiceNo and a.InvoiceNo_seq=@InvoiceNo_seq ");
            sb.Append(" ORDER BY b.[SeqNo] ");
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

        private void button51_Click(object sender, EventArgs e)
        {
            UPPACK();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\PACK.xls";

            string OHEM = fmLogin.LoginID.ToString().ToUpper().Trim();
            //if (OHEM == "NANCYTSAI")
            //{
            //    GetExcelProduct(FileName, GetOrderData3CUST(), "Y", "N");
            //}
            //else
            //{
            //    GetExcelProduct(FileName, GetOrderData3CUST(), "N", "N");
            //}
            GetExcelProduct(FileName, GetOrderData3CUST(), "Y", "N");
        }
        private System.Data.DataTable GetOrderData3CUST()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(B.seqno AS decimal(10,1)) S,CASE WHEN a.SayTotal='0' THEN CASE WHEN ISNULL(A.SAYCTN,'0')='0' THEN  CAST(T0.CNO AS VARCHAR) +'CTNS' ELSE A.SAYCTN +'CTNS' END ELSE  a.SayTotal+' PLTS' END as cc,a.[PLNo] ,a.[PDate],a.[ForAccount],'SHIP TO:'+a.[ShippedBy] as ShippedBy,a.[Shipping_From],a.[Shipping_Per] as ShippingPer,Convert(varchar(10),Getdate(),111) as 日期,a.[ColumnTotal] as '欄位統計'   ");
            sb.Append(" ,CAST(a.[Net] AS VARCHAR) as '耐特',CAST(a.[Gross] AS VARCHAR) as '螺絲',a.[Shipping_To],a.[ShippedOn] as ShippedOn,'BILL TO :'+a.[Bill_To] as Bill_To,a.[UserName],a.[CreateDate],a.[Memo]   ");
            sb.Append(" ,a.[Quantity] as '總數',CAST(a.[Net] AS VARCHAR),a.[Gross],a.[SayTotal],b.[SeqNo],b.[PackageNo],b.[CNo],   ");
            sb.Append(" CASE WHEN ISNULL(B.PACKMARK,'') <> 'True' THEN '' ELSE cast(B.seqno2+1 as varchar)+')' END+CASE ISNULL(B.TREETYPE,'') WHEN 'S' THEN b.[DescGoods]+ '(See Attachment List)' ELSE b.[DescGoods]  END DescGoods   ");
            sb.Append(" ,b.[Quantity] as Quantity ,b.[Net] as Ne ,cast(b.[Gross] as varchar) as Go ,b.[MeasurmentCM] FROM [PackingListM] as a   ");
            sb.Append(" left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)   ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)   ");
            sb.Append(" LEFT JOIN (SELECT ShippingCode,PLNo,MAX(CAST(CASE ISNULL(CHARINDEX('~', CNO),0) WHEN 0 THEN CNO ELSE SUBSTRING(CNO,CHARINDEX('~', CNO)+1,3) END AS INT))  CNO  FROM PackingListD  GROUP BY ShippingCode,PLNo) T0 ON (T0.ShippingCode=A.ShippingCode and T0.PLNo=A.PLNo)  ");
            sb.Append(" where a.shippingcode=@shippingcode and a.PLNo=@PLNo  ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT DISTINCT CAST(B.seqno+'.1' AS decimal(10,1) ) ,'' cc,'','','','','','','','' ");
            sb.Append(" ,'','','','','','','','' ");
            sb.Append(" ,'',a.[Net],a.[Gross],'',b.[SeqNo],'','', ");
            sb.Append(" CASE WHEN ISNULL(E.Substitute,'')='' THEN  '  P/N:'+F.ITEMCODES ELSE  '  P/N:'+ E.Substitute END   COLLATE  Chinese_Taiwan_Stroke_CI_AS DescGoods   ");
            sb.Append(" ,'' ,'' ,'' ,'' FROM [PackingListM] as a   ");
            sb.Append(" left join  [PackingListD] as b on (a.ShippingCode=b.ShippingCode and a.PLNo=b.PLNo)   ");
            sb.Append(" left join shipping_main as c on (a.shippingcode=c.shippingcode)   ");
            sb.Append(" LEFT JOIN INVOICED D ON (D.INDESCRIPTION=B.DESCGOODS AND D.SHIPPINGCODE=B.SHIPPINGCODE)  ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSCN E ON (D.ITEMCODE =E.ItemCode  COLLATE  Chinese_Taiwan_Stroke_CI_AS  AND C.CardCode =E.CardCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" LEFT JOIN (");
            sb.Append(" SELECT  SHIPPINGCODE,T0.ITEMCODE,MAX(U_CUSTITEMCODE) ITEMCODES  FROM shipping_ITEM T0 LEFT JOIN AcmeSql02.DBO.RDR1  T1 ON (CAST(T0.Docentry AS varchar) = CAST(T1.DocEntry AS varchar)  AND T0.linenum =T1.LineNum)");
            sb.Append(" WHERE ISNULL(U_CUSTITEMCODE,'') <> ''");
            sb.Append(" GROUP BY SHIPPINGCODE,T0.ITEMCODE ) F ON (D.ITEMCODE =F.ITEMCODE AND F.ShippingCode =A.ShippingCode)");
            sb.Append(" where a.shippingcode=@shippingcode and a.PLNo=@PLNo  ");
            sb.Append(" ORDER BY S ");


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
        public System.Data.DataTable GetBU1()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT REPLACE([WEIGHT],'T','')+'T' T ,CAST(REPLACE([WEIGHT],'T','') AS decimal(10,1)) FROM WH_CARFEE WHERE CARNAME ='友福' ORDER BY CAST(REPLACE([WEIGHT],'T','') AS decimal(10,1)) ";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        public System.Data.DataTable GetFEEH()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT AT,AA,AD,AGA,ASHA,AE,ATIME,AAMT  FROM SHIP_FEE WHERE ABIN=@ABIN AND AAMT <>0");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ABIN", aBINTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        public void UPFEEH(string AT, string AA, string AD, string AGA, string ASHA, string AE, string ATIME, string AAMT)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE SHIP_FEE SET  AT=@AT,AA=@AA,AD=@AD,AGA=@AGA,ASHA=@ASHA,AE=@AE,ATIME=@ATIME,AAMT=@AAMT WHERE  SHIPPINGCODE=@ABIN ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AT", AT));
            command.Parameters.Add(new SqlParameter("@AA", AA));
            command.Parameters.Add(new SqlParameter("@AD", AD));
            command.Parameters.Add(new SqlParameter("@AGA", AGA));
            command.Parameters.Add(new SqlParameter("@ASHA", ASHA));
            command.Parameters.Add(new SqlParameter("@AE", AE));
            command.Parameters.Add(new SqlParameter("@ATIME", ATIME));
            command.Parameters.Add(new SqlParameter("@AAMT", AAMT));
            command.Parameters.Add(new SqlParameter("@ABIN", aBINTextBox.Text));

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
        public System.Data.DataTable GetFEE1(string LOCATION)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT AMT FROM WH_CARFEE  WHERE CARNAME ='友福'  AND LOCATION=@LOCATION AND [WEIGHT] =@WEIGHT");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WEIGHT", aTTextBox.Text));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        public System.Data.DataTable GetFEE2()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT AMT FROM WH_CARFEE  WHERE CARNAME ='友福'  AND LOCATION=@LOCATION AND [WEIGHT] =@WEIGHT");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            string FA = aDTextBox.Text;

            command.Parameters.Add(new SqlParameter("@LOCATION", FA));
            command.Parameters.Add(new SqlParameter("@WEIGHT", aTTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        private void K1()
        {
            try
            {
                double F1 = 0;
                double F2 = 0;
                double F2S = 0;

                //>汐止/基隆/桃園/楊梅櫃場
                if (aATextBox.Text == "內湖(進金生)")
                {
                    //if (aDTextBox.Text != "C  K  S")
                    //{
                    System.Data.DataTable FEE = GetFEE1("內    湖");
                    if (FEE.Rows.Count > 0)
                    {
                        //   F1 = Convert.ToDouble(FEE.Rows[0][0]) * 0.8;
                        F1 = Convert.ToDouble(FEE.Rows[0][0]);
                    }

                    System.Data.DataTable FEE2 = GetFEE2();
                    if (FEE2.Rows.Count > 0)
                    {
                        F2S = Convert.ToDouble(FEE2.Rows[0][0]);
                    }
                    if (F2S > F1)
                    {
                        F1 = F2S;
                    }
                    // }
                }
                else if (aATextBox.Text == "平鎮(博豐)")
                {
                    if (aDTextBox.Text == "汐止櫃場" || aDTextBox.Text == "基隆櫃場" || aDTextBox.Text == "基隆櫃場(長春、台基)")
                    {
                        System.Data.DataTable FEE = GetFEE1("平    鎮");
                        if (FEE.Rows.Count > 0)
                        {
                            F1 = Convert.ToDouble(FEE.Rows[0][0]) * 0.8;


                        }
                    }
                }
                if (aATextBox.Text != "內湖(進金生)")
                {
                    System.Data.DataTable FEE3 = GetFEE2();
                    if (FEE3.Rows.Count > 0)
                    {
                        F2 = Convert.ToDouble(FEE3.Rows[0][0]);
                    }
                }

                decimal aGA = util.CINT(aGATextBox.Text);
                decimal aSH = util.CINT(aSHATextBox.Text);
                decimal aE = util.CINT(aETextBox.Text);
                decimal aTime = util.CINT(aTimeTextBox.Text);


                //bSAMTTextBox
                aAMTTextBox.Text = (Convert.ToDecimal(F1 + F2) + aGA + aSH + aE + aTime).ToString();

            }
            catch { }
        }
        private void K1H()
        {
            try
            {
                double F1 = 0;
                double F2 = 0;


                //>汐止/基隆/桃園/楊梅櫃場
                if (aATextBox.Text == "內湖(進金生)")
                {
                    //if (aDTextBox.Text != "C  K  S")
                    //{
                    System.Data.DataTable FEE = GetFEE1("內    湖");
                    if (FEE.Rows.Count > 0)
                    {
                        //   F1 = Convert.ToDouble(FEE.Rows[0][0]) * 0.8;
                        F1 = Convert.ToDouble(FEE.Rows[0][0]);
                    }

                    System.Data.DataTable FEE2 = GetFEE2();
                    if (FEE2.Rows.Count > 0)
                    {
                        F2 = Convert.ToDouble(FEE2.Rows[0][0]);
                    }
                    if (F2 > F1)
                    {
                        F1 = F2;
                    }
                    // }
                }
                else if (aATextBox.Text == "平鎮(博豐)")
                {
                    if (aDTextBox.Text == "汐止櫃場" || aDTextBox.Text == "基隆櫃場" || aDTextBox.Text == "基隆櫃場(長春、台基)")
                    {
                        System.Data.DataTable FEE = GetFEE1("平    鎮");
                        if (FEE.Rows.Count > 0)
                        {
                            F1 = Convert.ToDouble(FEE.Rows[0][0]) * 0.8;


                        }
                    }
                }
                if (aATextBox.Text != "內湖(進金生)")
                {
                    System.Data.DataTable FEE3 = GetFEE2();
                    if (FEE3.Rows.Count > 0)
                    {
                        F2 = Convert.ToDouble(FEE3.Rows[0][0]);
                    }
                }

                decimal aGA = util.CINT(aGATextBox.Text);
                decimal aSH = util.CINT(aSHATextBox.Text);
                decimal aE = util.CINT(aETextBox.Text);
                decimal aTime = util.CINT(aTimeTextBox.Text);


                //bSAMTTextBox
                aAMTTextBox.Text = (Convert.ToDecimal(F1 + F2) + aGA + aSH + aE + aTime).ToString();

            }
            catch { }
        }

        private void aTComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aAComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aDComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aGATextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aSHATextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aETextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }
        private void aTimeTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }
        private void aBINTextBox_TextChanged(object sender, EventArgs e)
        {
            System.Data.DataTable gg = GetFEEH();
            if (gg.Rows.Count > 0)
            {
                //AT,AA,AD,AGA,ASHA,AE,AAMT
                string AT = gg.Rows[0]["AT"].ToString();
                string AA = gg.Rows[0]["AA"].ToString();
                string AD = gg.Rows[0]["AD"].ToString();
                string AGA = gg.Rows[0]["AGA"].ToString();
                string ASHA = gg.Rows[0]["ASHA"].ToString();
                string AE = gg.Rows[0]["AE"].ToString();
                string ATIME = gg.Rows[0]["ATIME"].ToString();
                string AAMT = gg.Rows[0]["AAMT"].ToString();
                UPFEEH(AT, AA, AD, AGA, ASHA, AE, ATIME, AAMT);
            }

        }




        private void K2()
        {
            try
            {
                double F1 = 0;
                double F2 = 0;

                int G1 = cARDNAME5TextBox.Text.IndexOf("漢翔瑞");
                int G2 = cARDNAME5TextBox.Text.IndexOf("大東飛達");
                int G3 = cARDNAME5TextBox.Text.IndexOf("德迅");
                if (bATextBox.Text == "漢翔瑞" || G1 != -1)
                {

                    F1 = 500;

                }
                if (bATextBox.Text == "大東飛達" || G2 != -1)
                {

                    F1 = 400;

                }
                if (bATextBox.Text == "德迅" || G3 != -1)
                {

                    F1 = 500;

                }
                decimal BAE = util.CINT(bAETextBox.Text);
                bAFTextBox.Text = Convert.ToDecimal(F1).ToString();
                bAAMTTextBox.Text = (Convert.ToDecimal(F1) + BAE).ToString();

            }
            catch { }
        }

        private void K3()
        {
            try
            {

                double F1 = 0;
                double F2 = 0;
                if (bSTextBox.Text == "建新CFS")
                {
                    F1 = 2000;
                }
                if (bSTextBox.Text == "凌凱CFS")
                {
                    F1 = 1950;
                }
                if (bSTextBox.Text == "建新CY")
                {
                    int BSG = Convert.ToInt16(bSGTextBox.Text);
                    if (BSG == 1)
                    {
                        F1 = 2000;
                    }
                    else
                    {
                        F1 = 600 * (BSG - 1) + 2000;
                    }

                }

                if (bSTextBox.Text == "凌凱CY")
                {
                    int BSG = Convert.ToInt16(bSGTextBox.Text);
                    if (BSG == 1)
                    {
                        F1 = 2100;
                    }
                    else
                    {
                        F1 = 500 * (BSG - 1) + 2100;
                    }

                }

                if (bSPCheckBox.Checked)
                {
                    F2 = 500;
                }
                decimal BSE = util.CINT(bSETextBox.Text);
                decimal CSG = util.CINT(cSGTextBox.Text);
                bSFTextBox.Text = Convert.ToDecimal(F1 + F2).ToString();
                bSAMTTextBox.Text = (Convert.ToDecimal(F1 + F2) + BSE + CSG).ToString();

            }
            catch { }
        }


        private void bAETextBox_TextChanged(object sender, EventArgs e)
        {
            K2();
        }

        private void bAComboBox_TextChanged(object sender, EventArgs e)
        {
            K2();
        }

        private void bSComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void bSGTextBox_TextChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void bSPCheckBox_TextChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void bSETextBox_TextChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void K4()
        {
            try
            {

                decimal F1 = 0;
                decimal CAK = util.CINT(cAKTextBox.Text);
                if (CAK <= 300)
                {
                    F1 = CAK * 6;
                }
                else
                {
                    F1 = (CAK - 300) * 2 + 1800;
                }

                if (F1 < 100)
                {
                    F1 = 100;
                }
                cAZTextBox.Text = F1.ToString();


                decimal CAT = util.CINT(cATTextBox.Text);
                decimal CAS = util.CINT(cASTextBox.Text);
                decimal CAX = util.CINT(cAXTextBox.Text);
                cAAMTTextBox.Text = (Convert.ToDecimal(F1) + CAT + CAS + CAX).ToString();

            }
            catch { }
        }

        private void cAKTextBox_TextChanged(object sender, EventArgs e)
        {
            if (cAKTextBox.Text != "") 
            {
                cAXTextBox.Text = (Convert.ToDouble(cAKTextBox.Text) * 2).ToString();
            }
            
            K4();
        }

        private void cATTextBox_TextChanged(object sender, EventArgs e)
        {
            K4();
        }

        private void cAZTextBox_TextChanged(object sender, EventArgs e)
        {
            K4();
        }

        private void cASTextBox_TextChanged(object sender, EventArgs e)
        {
            K4();
        }

        private void K5()
        {
            try
            {
                decimal F1 = 0;
                decimal F2 = 0;
                decimal CSC = util.CINT3(cSCTextBox.Text);
                int result = Convert.ToInt16(Math.Ceiling(CSC));




                if (CSC == 0)
                {
                    F1 = 0;
                }
                else if (CSC < 1)
                {
                    F1 = 380;
                }
                else
                {
                    F1 = CSC * 380;
                }

                F2 = result * 60;
                int RE2 = Convert.ToInt16(Math.Round(F1, 0, MidpointRounding.AwayFromZero));
                if (cSCTextBox.Text != "")
                {
                    cSBTextBox.Text = RE2.ToString();
                }
                cSGTextBox.Text = F2.ToString();


                decimal CST = util.CINT(cSTTextBox.Text);
                decimal CSS = util.CINT(cSSTextBox.Text);
                decimal CSV = util.CINT(cSVTextBox.Text);
                decimal CSD = util.CINT(cSDTextBox.Text);
                decimal CSH = util.CINT(cSHTextBox.Text);
                decimal CSLIU = util.CINT(cSLIUTextBox.Text);
                decimal CSS2 = util.CINT(cSS2TextBox.Text);
                decimal CSS3 = util.CINT(cSS3TextBox.Text);
                //cSAMTTextBox.Text = (Convert.ToDecimal(F2 + RE2) + CST + CSS + CSV + CSD + CSH).ToString();
                cSAMTTextBox.Text = (Convert.ToDecimal(RE2) + CST + CSS + CSV + CSD + CSH + CSLIU + CSS2 + CSS3).ToString();
            }
            catch { }
        }

        private void cSCTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSTTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSBTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSGTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
            K3();
        }

        private void cSSTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSVTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSDTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSHTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void comboBox12_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU1();

            comboBox12.Items.Clear();

            comboBox12.Items.Add("");
            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox12.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            aTTextBox.Text = comboBox12.Text;
        }

        private void comboBox14_SelectedValueChanged(object sender, EventArgs e)
        {
            aATextBox.Text = comboBox14.Text;
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            aDTextBox.Text = comboBox15.Text;
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            bSTextBox.Text = comboBox16.Text;
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            bATextBox.Text = comboBox17.Text;
        }

        private void aATextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aDTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void aTTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void bSTextBox_TextChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void bATextBox_TextChanged(object sender, EventArgs e)
        {
            K2();
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            lANDTYPETextBox.Text = comboBox18.Text;
        }


        private System.Data.DataTable GETLOCATION()
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = new SqlConnection(globals.ConnectionString);
            sb.Append(" SELECT LOCATION FROM WH_DHL_COUNTRY  WHERE ENGNAME=@ENGNAME AND CTYPE ='出口' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ENGNAME", goalPlaceTextBox.Text.Trim()));

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


        private System.Data.DataTable GETFEE(string LOCATION, string KG)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = new SqlConnection(globals.ConnectionString);
            sb.Append("SELECT TOP 1 CASE WHEN KG>30 THEN (ROUND(((FEE)/1.05),0)*@KG) ELSE FEE/1.05 END FEE  FROM [WH_DHL_FEE] WHERE LOCATION =@LOCATION AND CTYPE ='出口' AND KG<=@KG ORDER BY KG DESC "); ;
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@KG", KG));
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

        private System.Data.DataTable GETFIREFEE(string DOCDATE)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = new SqlConnection(globals.ConnectionString);
            sb.Append("SELECT FEE FROM SHIP_FIREFEE WHERE DOCDATE=@DOCDATE"); ;
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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

        private void dHL1TextBox_TextChanged(object sender, EventArgs e)
        {
            KDHL();
        }
        private void KDHL()
        {
            try
            {
                decimal dHL4 = 0;
                decimal dHL5 = 0;
                System.Data.DataTable GDHL1 = GETLOCATION();
                if (GDHL1.Rows.Count > 0)
                {
                    string LOC = GDHL1.Rows[0][0].ToString();
                    if (dHL1TextBox.Text != "")
                    {
                        System.Data.DataTable GDHL2 = GETFEE(LOC, dHL1TextBox.Text);
                        if (GDHL2.Rows.Count > 0)
                        {
                            string DHL = GDHL2.Rows[0][0].ToString();
                            decimal BSE = util.CINT3(DHL);
                            int RE2 = Convert.ToInt16(Math.Round(BSE, 0, MidpointRounding.AwayFromZero));
                            dHL3TextBox.Text = (RE2).ToString();
                        }
                    }
                }

                string gf = DateTime.Now.ToString("yyyyMM");
                System.Data.DataTable GFEE = GETFIREFEE(DateTime.Now.ToString("yyyyMM"));
                if (GFEE.Rows.Count > 0)
                {
                    decimal DD = util.CINT(dHL3TextBox.Text);
                    int RE2 = Convert.ToInt16(Math.Round(DD * Convert.ToDecimal(GFEE.Rows[0][0]), 0, MidpointRounding.AwayFromZero));
                    dHL4 = RE2;
                    dHL4TextBox.Text = (RE2).ToString();
                }
                decimal dHL3 = util.CINT(dHL3TextBox.Text);
                decimal dHL7 = util.CINT(dHL7TextBox.Text);

                if (dHL5CheckBox.Checked)
                {
                    dHL5 = 400;
                }
                //bSAMTTextBox
                dHL2TextBox.Text = (dHL3 + dHL4 + dHL5 + dHL7).ToString();

            }
            catch { }
        }

        private void cSLIUTextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODETextBox1.Text = Convert.ToString(LookupValues[0]);
                cARDNAMETextBox1.Text = Convert.ToString(LookupValues[1]);
            }
        }



        private void button46_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODE2TextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAME2TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODE5TextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAME5TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button49_Click_1(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODE3TextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAME3TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODE4TextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAME4TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();



            if (LookupValues != null)
            {
                cARDCODE6TextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAME6TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }
        private void DMARK()
        {

            if (mEMO3TextBox.Text != "")
            {
                System.Data.DataTable dth = ship.Mark;
                if (dth.Rows.Count == 0)
                {
                    string[] arrurl = mEMO3TextBox.Text.Split(new Char[] { ',' });
                    int G = 0;
                    foreach (string i in arrurl)
                    {
                        System.Data.DataTable dt3HD = GetMenu.GetWHMARK(i);
                        System.Data.DataTable G1 = GetMenu.GetWHPACK4(i);
                        if (G1.Rows.Count > 0)
                        {
                            string FT = G1.Rows[0]["TOTAL"].ToString();

                            if (dt3HD.Rows.Count > 0 && FT != "0")
                            {


                                string MARK = dt3HD.Rows[0][0].ToString();
                                string[] NewLine = MARK.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                                foreach (string ESi in NewLine)
                                {
                                    G++;
                                    DataRow drw2 = dth.NewRow();

                                    drw2["Seq"] = G.ToString();
                                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                                    string EE = ESi.ToString().ToUpper();
                                    int EF = EE.IndexOf("PL/");
                                    int EF2 = EE.IndexOf("NO");
                                    //C/NO
                                    if (EF != -1)
                                    {
                                        string PACK = G1.Rows[0]["PLATENO"].ToString();
                                        if (PACK == "1")
                                        {
                                            drw2["Mark"] = "PL / NO.:1/1";
                                        }
                                        else
                                        {
                                            drw2["Mark"] = "PL / NO.:1/" + PACK + "~" + PACK + "/" + PACK;
                                        }
                                    }
                                    else if (EF2 != -1)
                                    {
                                        string PACK = G1.Rows[0]["CARTONNO"].ToString();
                                        if (PACK == "1")
                                        {
                                            drw2["Mark"] = "C/NO.:1/1";
                                        }
                                        else
                                        {
                                            drw2["Mark"] = "C/NO.:1/" + PACK + "~" + PACK + "/" + PACK;
                                        }
                                    }
                                    else
                                    {
                                        drw2["Mark"] = ESi.ToString();
                                    }
                                    dth.Rows.Add(drw2);

                                }
                            }
                        }
                    }


                    this.markBindingSource.EndEdit();
                    this.markTableAdapter.Update(ship.Mark);
                    ship.Mark.AcceptChanges();


                }



            }


        }

        private void cARDNAMETextBox1_TextChanged(object sender, EventArgs e)
        {
            K2();
        }
        public System.Data.DataTable GetSHIPPCAK2()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" 			  SELECT  BuCntctPrsn  FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(BuCntctPrsn,'') <> ''");
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

        private void cSS3TextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void cSS2TextBox_TextChanged(object sender, EventArgs e)
        {
            K5();
        }

        private void button52_Click(object sender, EventArgs e)
        {
            SHIP_FIREFEE frm1 = new SHIP_FIREFEE();
            frm1.Show();
        }

        //private void dHL6ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (dHL6ComboBox.Text == "2.5公斤以下")
        //    {
        //        dHL7TextBox.Text = "0";
        //    }
        //    if (dHL6ComboBox.Text == "2.51至30公斤")
        //    {
        //        dHL7TextBox.Text = "90";
        //    }
        //    if (dHL6ComboBox.Text == "30.1至70公斤")
        //    {
        //        dHL7TextBox.Text = "550";
        //    }
        //    if (dHL6ComboBox.Text == "70.1至300公斤")
        //    {
        //        dHL7TextBox.Text = "21800";
        //    }
        //    if (dHL6ComboBox.Text == "300.1公斤(含)以上")
        //    {
        //        dHL7TextBox.Text = "7000";
        //    }
        //}

        private void dHL7TextBox_TextChanged(object sender, EventArgs e)
        {
            KDHL();
        }

        private void dHL3TextBox_TextChanged(object sender, EventArgs e)
        {
            KDHL();
        }

        private void dHL4TextBox_TextChanged(object sender, EventArgs e)
        {
            KDHL();
        }

        private void dHL5CheckBox_TextChanged(object sender, EventArgs e)
        {
            KDHL();
        }

        private void dHL5CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            KDHL();
        }

        private void bSFTextBox_TextChanged(object sender, EventArgs e)
        {
            K3();
        }

        private void cSGTextBox_TextChanged_1(object sender, EventArgs e)
        {
            K3();
        }

        private void button53_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            System.Data.DataTable SI1 = GetSI1(shippingCodeTextBox.Text);

            if (SI1.Rows.Count > 0)
            {
                for (int i = 0; i <= SI1.Rows.Count - 1; i++)
                {
                    string WHSCODE = SI1.Rows[i]["WHSCODE"].ToString();
                    System.Data.DataTable SI2 = GetSI2(shippingCodeTextBox.Text, WHSCODE);
                    System.Data.DataTable SI3S = GetSI1S(WHSCODE);
                    string WHSNAMES = SI3S.Rows[0][0].ToString();
                    if (SI2.Rows.Count > 0)
                    {

                        string NumberName = "WH" + DateTime.Now.ToString("yyyyMMdd");
                        SqlConnection Connection = globals.Connection;
                        string AutoNum = util.GetAutoNumber(Connection, NumberName);

                        string KK = NumberName + AutoNum + "X";

                        string username = "";

                        if (SI2.Rows.Count > 0)
                        {

                            System.Data.DataTable ff = GetOHEM(fmLogin.LoginID.ToString());
                            if (ff.Rows.Count > 0)
                            {
                                DataRow drw = ff.Rows[0];
                                username = drw["姓名"].ToString();
                                username = username.Replace("(", "");
                                username = username.Replace(")", "");

                            }
                            string DOC = SI2.Rows[0]["DOCENTRY"].ToString();
                            System.Data.DataTable dt1sar = GetMenu.Getocrdnew2(DOC, "銷售訂單");
                            DataRow drwF = dt1sar.Rows[0];
                            string OBUShip = drwF["shipbuilding"].ToString() +
                                      Environment.NewLine + drwF["shipstreet"].ToString() +
                                      Environment.NewLine + "TEL:" + drwF["shipblock"].ToString() +
                                      Environment.NewLine + "FAX:" + drwF["shipcity"].ToString() +
                                      Environment.NewLine + "ATTN:" + drwF["shipzipcode"].ToString();


                            string OBUBill = drwF["billbuilding"].ToString() +
                     Environment.NewLine + drwF["billstreet"].ToString() +
                     Environment.NewLine + "TEL:" + drwF["billblock"].ToString() +
                     Environment.NewLine + "FAX:" + drwF["billcity"].ToString() +
                     Environment.NewLine + "ATTN:" + drwF["billzipcode"].ToString();
                            //電話號碼

                            System.Data.DataTable D1 = GetOrderDataS(DOC);
                            string 業務 = D1.Rows[0]["業務"].ToString();
                            string 電話號碼 = D1.Rows[0]["電話號碼"].ToString();
                            string 工廠地址 = D1.Rows[0]["工廠地址"].ToString();
                            string 連絡人 = D1.Rows[0]["連絡人"].ToString();

                            int g = 工廠地址.IndexOf("司");


                            if (g != -1)
                            {
                                工廠地址 = 工廠地址.Substring(g + 1).Trim();

                            }

                            string quantity = "";
                            if (dOCTYPETextBox.Text == "銷售訂單" || dOCTYPETextBox.Text == "銷售")
                            {
                                System.Data.DataTable G1 = GetORDR();


                                if (G1.Rows.Count > 0)
                                {
                                    string Doc = G1.Rows[0][0].ToString();
                                    string LINE = G1.Rows[0][1].ToString();
                                    System.Data.DataTable SHIPDATE = GetSHIPDATE(Doc, LINE);
                                    if (SHIPDATE.Rows.Count > 0)
                                    {
                                        quantity = SHIPDATE.Rows[0][2].ToString();
                                    }
                                }
                            }
                            else if (dOCTYPETextBox.Text == "調撥單" || dOCTYPETextBox.Text == "調撥" || dOCTYPETextBox.Text.Contains("調撥"))
                            {
                                System.Data.DataTable K1 = GETARRIVE2(mEMO3TextBox.Text);
                                if (K1.Rows.Count > 0)
                                {
                                    quantity = K1.Rows[0][0].ToString();
                                }
                            }


                            AddSHIPMAIN(KK, cardNameTextBox.Text, cardCodeTextBox.Text, DOC, dOCTYPETextBox.Text, WHSNAMES, 業務, 電話號碼, username, OBUBill, OBUShip, boardCountNoTextBox.Text, quantity);
                            MessageBox.Show("上傳成功 倉管單號 : " + KK);
                            UPDATESHIPWHNO(KK, WHSCODE, shippingCodeTextBox.Text);
                            shipping_ItemTableAdapter.Fill(ship.Shipping_Item, MyID);

                            if (mEMO3TextBox.Text == "")
                            {
                                mEMO3TextBox.Text = KK;

                            }
                            string DOCENTRY = "";
                            int LINENUM = 0;
                            int SeqNo = 0;
                            string ShipDate = "";
                            string U_PAY = "";
                            string U_SHIPDAY = "";
                            string U_SHIPSTATUS = "";
                            string U_MARK = "";
                            string U_MEMO = "";
                            string PO = "";
                            string LOCATION = "";
                            string UNIT = "";
                            int SM = 0;
                            string ITEMCODE = "";
                            string TREETYPE = "";
                            for (int i2 = 0; i2 <= SI2.Rows.Count - 1; i2++)
                            {
                                //Dscription
                                ITEMCODE = SI2.Rows[i2]["ITEMCODE"].ToString();
                                string Dscription = SI2.Rows[i2]["Dscription"].ToString();

                                SeqNo = Convert.ToInt16(SI2.Rows[i2]["SeqNo"]);
                                LINENUM = Convert.ToInt16(SI2.Rows[i2]["LINENUM"]);
                                int QTY = Convert.ToInt16(SI2.Rows[i2]["Quantity"]);


                                DOCENTRY = SI2.Rows[i2]["DOCENTRY"].ToString();
                                System.Data.DataTable SI3 = GetSI3(ITEMCODE, WHSCODE);
                                string PARTNO = "";
                                int ONHAND = 0;
                                string VER = "";
                                string GRADE = "";
                                string FRGNAME = "";
                                string WHSNAME = "";
                                if (SI3.Rows.Count > 0)
                                {
                                    PARTNO = SI3.Rows[0]["PARTNO"].ToString();
                                    ONHAND = Convert.ToInt32(SI3.Rows[0]["ONHAND"]);
                                    VER = SI3.Rows[0]["版本"].ToString();
                                    GRADE = SI3.Rows[0]["等級"].ToString();
                                    FRGNAME = SI3.Rows[0]["品名規格"].ToString();
                                    WHSNAME = SI3.Rows[0]["WHSNAME"].ToString();
                                }

                                System.Data.DataTable SI4 = GetSI4(DOCENTRY, LINENUM);


                                if (SI4.Rows.Count > 0)
                                {

                                    ShipDate = SI4.Rows[0]["排程日期"].ToString();
                                    U_PAY = SI4.Rows[0]["付款"].ToString();
                                    U_SHIPDAY = SI4.Rows[0]["押出貨日"].ToString();
                                    U_SHIPSTATUS = SI4.Rows[0]["貨況"].ToString();
                                    U_MARK = SI4.Rows[0]["特殊嘜頭"].ToString();

                                    U_MEMO = SI4.Rows[0]["注意事項"].ToString();
                                    PO = SI4.Rows[0]["PO"].ToString();
                                    LOCATION = SI4.Rows[0]["產地"].ToString();

                                    UNIT = SI4.Rows[0]["單位"].ToString();
                                }

                                System.Data.DataTable SHIR2 = GetSHIR1(DOCENTRY, LINENUM);

                                if (SHIR2.Rows.Count > 0)
                                {
                                    string P1S = SHIR2.Rows[0][0].ToString();
                                    if (P1S == "S")
                                    {
                                        PARTNO = "母料號";
                                    }
                                    TREETYPE = P1S;
                                }

                                AddSHIPITEM(KK, SeqNo + SM, DOCENTRY, LINENUM, dOCTYPETextBox.Text, ITEMCODE, Dscription, Convert.ToInt32(QTY), PARTNO, ONHAND, VER, GRADE, FRGNAME, WHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, LOCATION, UNIT, TREETYPE);


                                System.Data.DataTable SHIR1 = GetSHIR1(DOCENTRY, LINENUM);
                                string P1 = SHIR1.Rows[0][0].ToString();
                                if (P1 == "S")
                                {
                                    if (SHIR1.Rows.Count > 0)
                                    {
                                        System.Data.DataTable BSHIR2 = GetSHIR2(DOCENTRY);
                                        if (BSHIR2.Rows.Count > 0)
                                        {
                                            for (int Bi2 = 0; Bi2 <= BSHIR2.Rows.Count - 1; Bi2++)
                                            {
                                                SM++;
                                                int SLINENUM = Convert.ToInt16(BSHIR2.Rows[Bi2]["LineNum"]);
                                                string SITEMCODE = BSHIR2.Rows[Bi2]["ItemCode"].ToString();
                                                string SDESC = BSHIR2.Rows[Bi2]["Dscription"].ToString();
                                                string SQTY = BSHIR2.Rows[Bi2]["QTY"].ToString();
                                                System.Data.DataTable BSI3 = GetSI3(SITEMCODE, WHSCODE);

                                                string SPARTNO = "";
                                                int SONHAND = 0;
                                                string SVER = "";
                                                string SGRADE = "";
                                                string SFRGNAME = "";
                                                string SWHSNAME = "";
                                                string SLOCATION = "";
                                                string SUNIT = "";
                                                if (BSI3.Rows.Count > 0)
                                                {
                                                    SUNIT = BSI3.Rows[0]["單位"].ToString();
                                                    SLOCATION = BSI3.Rows[0]["產地"].ToString();
                                                    SONHAND = Convert.ToInt32(BSI3.Rows[0]["ONHAND"]);
                                                    SVER = BSI3.Rows[0]["版本"].ToString();
                                                    SGRADE = BSI3.Rows[0]["等級"].ToString();
                                                    SFRGNAME = BSI3.Rows[0]["品名規格"].ToString();
                                                    SWHSNAME = BSI3.Rows[0]["WHSNAME"].ToString();
                                                }

                                                SPARTNO = ITEMCODE + "-子料號-" + SM.ToString();
                                                AddSHIPITEM(KK, SeqNo + SM, DOCENTRY, SLINENUM, dOCTYPETextBox.Text, SITEMCODE, SDESC, Convert.ToInt32(SQTY), SPARTNO, SONHAND, SVER, SGRADE, SFRGNAME, SWHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, SLOCATION, SUNIT, "I");
                                            }
                                        }
                                    }
                                    //  
                                }
                            }

                        }
                    }
                }
            }
        }
        private System.Data.DataTable GetOrderDataS(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();



            sb.Append("                SELECT m.docnum as 單號,d.linenum as 欄號,Convert(varchar(10),D.U_ACME_SHIPDAY,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,  ");
            sb.Append("                          T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,d.U_CUSTITEMCODE,d.U_CUSTDOCENTRY    ");
            sb.Append("                          ,T1.street+ISNULL(T1.COUNTY,'')  工廠地址,os.slpname as 業務,os.MEMO as 流程,  ");
            sb.Append("                          SALUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,ROHS='ROHS',AU='AUS',  ");
            sb.Append("                          d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,  ");
            sb.Append("                          rtrim(ISNULL(m.numatcard,'')+ISNULL(m.U_ACME_PAYGUI,'')+cast(isnull(m.u_acme_memo,'') as nvarchar(1000))) 備註,oi.usertext 主要描述,oi.U_LOCATION 產地  FROM ordr m  ");
            sb.Append("                          left join rdr1 d on m.docentry=d.docentry  ");
            sb.Append("                          LEFT JOIN  CRD1 T1 ON (M.CARDCODE=T1.CARDCODE AND M.shiptocode=T1.ADDRESS and T1.adrestype='S')    ");
            sb.Append("                          left join oslp os on os.slpcode=m.slpcode  ");
            sb.Append("                          left join oitm oi on oi.itemcode=d.itemcode  ");
            sb.Append("                          where m.DOCENTRY =@DOCENTRY ");
            sb.Append("                          order by m.DOCENTRY,d.visorder  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "new01");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GetORDR()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DOCENTRY,LINENUM FROM WH_ITEM4 WHERE SHIPPINGCODE=@SHIPPINGCODE ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", mEMO3TextBox.Text));
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
        public void AddSHIPMAIN(string ShippingCode, string CardName, string CardCode, string PINO, string forecastDay, string shipping_OBU, string buCntctPrsn, string cFS, string createName, string oBUBillTo, string oBUShipTo, string boardCountNo, string quantity)
        {

            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into WH_MAIN(ShippingCode,CardName,CardCode,PINO,forecastDay,shipping_OBU,buCntctPrsn,cFS,createName,oBUBillTo,oBUShipTo,boardCountNo,Quantity,SONO) values(@ShippingCode,@CardName,@CardCode,@PINO,@forecastDay,@shipping_OBU,@buCntctPrsn,@cFS,@createName,@oBUBillTo,@oBUShipTo,@boardCountNo,@quantity,@SONO)", Connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));

            command.Parameters.Add(new SqlParameter("@PINO", PINO));
            command.Parameters.Add(new SqlParameter("@forecastDay", forecastDay));
            command.Parameters.Add(new SqlParameter("@shipping_OBU", shipping_OBU));

            command.Parameters.Add(new SqlParameter("@buCntctPrsn", buCntctPrsn));
            command.Parameters.Add(new SqlParameter("@cFS", cFS));
            command.Parameters.Add(new SqlParameter("@createName", createName));

            command.Parameters.Add(new SqlParameter("@oBUBillTo", oBUBillTo));
            command.Parameters.Add(new SqlParameter("@oBUShipTo", oBUShipTo));
            command.Parameters.Add(new SqlParameter("@boardCountNo", boardCountNo));
            command.Parameters.Add(new SqlParameter("@quantity", quantity));
            command.Parameters.Add(new SqlParameter("@SONO", "Unchecked"));
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

        public void AddSHIPITEM(string ShippingCode, int SeqNo, string Docentry, int linenum, string ItemRemark, string ItemCode, string Dscription, int Quantity, string PiNo, int NowQty, string Ver, string Grade, string FrgnName, string WHName, string ShipDate, string U_PAY, string U_SHIPDAY, string U_SHIPSTATUS, string U_MARK, string U_MEMO, string PO, string LOCATION, string CARDCODE, string TREETYPE)
        {
            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into WH_ITEM4(ShippingCode,SeqNo,Docentry,linenum,ItemRemark,ItemCode,Dscription,Quantity,PiNo,NowQty,Ver,Grade,FrgnName,WHName,ShipDate,U_PAY,U_SHIPDAY,U_SHIPSTATUS,U_MARK,U_MEMO,PO,LOCATION,CARDCODE,TREETYPE) values(@ShippingCode,@SeqNo,@Docentry,@linenum,@ItemRemark,@ItemCode,@Dscription,@Quantity,@PiNo,@NowQty,@Ver,@Grade,@FrgnName,@WHName,@ShipDate,@U_PAY,@U_SHIPDAY,@U_SHIPSTATUS,@U_MARK,@U_MEMO,@PO,@LOCATION,@CARDCODE,@TREETYPE)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Dscription", Dscription));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@PiNo", PiNo));
            command.Parameters.Add(new SqlParameter("@NowQty", NowQty));
            command.Parameters.Add(new SqlParameter("@Ver", Ver));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@FrgnName", FrgnName));
            command.Parameters.Add(new SqlParameter("@WHName", WHName));

            command.Parameters.Add(new SqlParameter("@ShipDate", ShipDate));
            command.Parameters.Add(new SqlParameter("@U_PAY", U_PAY));
            command.Parameters.Add(new SqlParameter("@U_SHIPDAY", U_SHIPDAY));
            command.Parameters.Add(new SqlParameter("@U_SHIPSTATUS", U_SHIPSTATUS));
            command.Parameters.Add(new SqlParameter("@U_MARK", U_MARK));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@TREETYPE", TREETYPE));


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
        public static System.Data.DataTable GetSHIPDATE(string DOCENTRY, string LINENUM)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              select U_SHIPDAY,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 112 ) SHIPDATE,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 111 ) SHIPDATE2  from RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GetSI1(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT WHSCODE FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND WHSCODE <> '' ");

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
        public static System.Data.DataTable GetSI1S(string WHSCODE)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" 	SELECT WHSNAME FROM OWHS  WHERE WHSCODE=@WHSCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public static System.Data.DataTable GetSI2(string SHIPPINGCODE, string WHSCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT * FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND WHSCODE =@WHSCODE AND WHSCODE like '%tw0%' ");

            sb.Append(" SELECT * FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE AND WHSCODE =@WHSCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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

        public static System.Data.DataTable GetSHIR1(string DOCENTRY, int LINENUM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  TREETYPE FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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

        public static System.Data.DataTable GetSHIR2(string DOCENTRY)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT *,CAST(Quantity AS INT) QTY FROM RDR1 WHERE DOCENTRY=@DOCENTRY  AND TREETYPE='I'");


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
        public static System.Data.DataTable GetSI3(string ITEMCODE, string WHSCODE)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT U_PARTNO PARTNO,T1.ONHAND ONHAND, T0.U_GRADE 等級, 版本='V.'+T0.U_VERSION,T0.frgnname 品名規格,T2.WHSNAME,U_LOCATION 產地,SALUNITMSR 單位  FROM OITM T0");
            sb.Append("  LEFT JOIN OITW T1 ON (T0.ITEMCODE=T1.ITEMCODE) ");
            sb.Append("  LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)   ");
            sb.Append("  WHERE T0.ITEMCODE=@ITEMCODE AND T1.WHSCODE=@WHSCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
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
        public static System.Data.DataTable GetSI4(string DOCENTRY, int LineNum)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("	                             SELECT Convert(varchar(10),d.u_acme_work,111) 排程日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭");
            sb.Append("								 ,d.U_MEMO as 注意事項,m.NUMATCARD as PO,oi.U_LOCATION 產地,d.TREETYPE,oi.SALUNITMSR 單位  FROM ordr m  ");
            sb.Append("                          left join rdr1 d on m.docentry=d.docentry  ");
            sb.Append("                          left join oitm oi on oi.itemcode=d.itemcode  ");
            sb.Append("       WHERE M.DOCENTRY=@DOCENTRY AND D.LineNum =@LineNum");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
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

        private void btnCyberNetExcel_Click(object sender, EventArgs e)
        {
            DataRow dr;
            bool tag = true;
            string WH = mEMO3TextBox.Text;
            string SH = shippingCodeTextBox.Text;
            string Sql = @"declare @ShippingCode  nvarchar(30)
                            set @ShippingCode ='{0}'
                            SELECT  A.DocEntry,
                                    A.LineNum,
                                    '' PN,
                                    A.ItemCode,
                                    C.U_itemname 品名,
                                    Quantity,
                                    DateCode ,
                                    '20200203' as ShipDate,
                                    '26.37' GWeight,
                                     PO,
                                    isnull(LPrint,1) as PrintQty,
                                    '13' 滿箱片數,
                                    isnull(PQty2,'0') 尾箱片數,
                                    '' 板號,
                                    '' 箱號,
                                    '156' 滿板片數,
                                    '' 尾板片數,
                                    B.U_MEMO
                            FROM WH_Item AS A
                            LEFT JOIN AcmeSql02.DBO.OSCN AS B
                            ON A.ItemCode COLLATE Chinese_Taiwan_BOPOMOFO_CI_AI = B.ItemCode 
                            LEFT JOIN AcmeSql02.dbo.oitm as C
                            ON A.ItemCode COLLATE Chinese_Taiwan_BOPOMOFO_CI_AI =  B.ItemCode

                            WHERE ShippingCode = @ShippingCode AND B.cardcode='0206-03' AND C.ItemCode LIKE A.ItemCode COLLATE Chinese_Taiwan_BOPOMOFO_CI_AI
                            ";
            //換算 Week
            string ShipDate = "";
            string ItemCode = "";
            string DocEntry = "";
            string LineNum = "";
            Sql = string.Format(Sql, WH);
            System.Data.DataTable dt = GetData(Sql);
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];
                DocEntry = Convert.ToString(dr["DocEntry"]);
                LineNum = Convert.ToString(dr["LineNum"]);
                ItemCode = Convert.ToString(dr["ItemCode"]);
                dr.BeginEdit();
                try
                {
                    //U_MEMO
                    dr["PN"] = Convert.ToString(dr["U_MEMO"]).Split(';')[0].Replace("MSI P/N:", "").Replace("MIS P/N#", "");
                    dr["滿箱片數"] = Convert.ToString(dr["U_MEMO"]).Split(';')[1].Replace("PCS/CTN", "");
                    dr["GWeight"] = Convert.ToString(dr["U_MEMO"]).Split(';')[2].Replace("G.W.", "").Replace("/CTN", "");
                    dr["滿板片數"] = Convert.ToString(dr["U_MEMO"]).Split(';')[3].Replace("PCS/PLT", "");

                    //找離倉日期 
                    // ShipDate = Convert.ToString(dr["ShipDate"]).Replace("/", "");
                    //  ShipDate = GetShipDateJoy(DocEntry, LineNum);
                    string ShipDocEntry = GetShipDocEntryJoy(WH);
                    ShipDate = GetArriveDayJoy(ShipDocEntry);
                    dr["DateCode"] = ShipDate.Substring(0, 4) + "/" + GetIso8601WeekOfYear(StrToDate(ShipDate));

                    //Model

                    dr["ItemCode"] = GetShipModel(ItemCode);
                    dr["PO"] = GetShipMark(DocEntry, LineNum);
                    // SapItemCode = Convert.ToString(dr["SapItemCode"]);
                    //System.Data.DataTable dtGetOitm = GetOitm(SapItemCode);
                    //if (dtGetOitm.Rows.Count > 0)
                    //{
                    //    string TYPE = "C";//"P" Carton Pallet;
                    //    //U_TMODEL Model,U_VERSION VER Oitm
                    //    string MODEL_NO = Convert.ToString(dtGetOitm.Rows[0]["U_TMODEL"]);
                    //    string U_VERSION = Convert.ToString(dtGetOitm.Rows[0]["U_VERSION"]);
                    //    System.Data.DataTable dtCart = GetCART_Lleyton(MODEL_NO, U_VERSION, TYPE);
                    //    if (dtCart.Rows.Count > 0)
                    //    {
                    //    }
                    //}
                    //                    JOB#WH20200608006X
                    //ACME P/N: M195RTN01.00LF2
                    //出貨數量：405PCS
                    //MSI P/O#  11956732 (SAP系統特殊嘜頭中取得)
                    //MSI P/N: S1J9E1A005AY0, (可由系統料號對應業務伙伴目錄對照號碼之備註MSI P/N)
                    //VENDOR P/N: M195RTN01.0 (可由系統料號對應船務型號)
                    //Description:19.5” TFT LCD MODULE(可由系統料號對應船務品名)
                    //15PCS/箱/G.W.24.00; 每板240pcs
                    //D/C:2020/6/12(請系統轉換成 年/週)

                }
                catch (Exception ex)
                {

                }
                dr.EndEdit();


            }

            Int32 PrintQty = 0;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];
                dr.BeginEdit();
                PrintQty = Convert.ToInt32(Math.Truncate(
                           Convert.ToDouble(Convert.ToInt32(dr["Quantity"]) / Convert.ToInt32(dr["滿箱片數"]))));
                dr["Quantity"] = Convert.ToString(dr["滿箱片數"]);
                dr["PrintQty"] = PrintQty;
                dr.EndEdit();
            }
            string FileNameC = GetExePath() + "\\" + SH + "_Cyber_箱標籤.xls";
            string TemplateC = GetExePath() + "\\Excel\\Mark\\" + "Cyber_箱標籤.xls";
            ExcelCyberBox(dt, TemplateC, FileNameC);

            string FileNameP = GetExePath() + "\\" + SH + "_Cyber_板嘜頭.xls";
            string TemplateP = GetExePath() + "\\Excel\\Mark\\" + "Cyber_板嘜頭.xls";
            ExcelCyberPallet(dt, TemplateP, FileNameP);


            UPLOAD(FileNameC);
            UPLOAD(FileNameP);

            string DIR = "\\\\acmesrv01\\SAP_Share\\shipping\\" + DateTime.Now.ToString("yyyyMM") + "\\";
            string uploadfileC = DIR + SH + "_Cyber_箱標籤.xls";

            string uploadfileP = DIR + SH + "_Cyber_板嘜頭.xls";
            System.Diagnostics.Process.Start(uploadfileC);
            System.Diagnostics.Process.Start(uploadfileP);
        }
        public System.Data.DataTable GetData(string Sql)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
            SqlConnection connection = new SqlConnection(ConnectiongString);//"server=acmesap; pwd=@rmas; uid=sapdbo; database=acmesqlsp"
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(Sql);

            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

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
        public System.Data.DataTable GetData(string ConnectiongString, string Sql)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString); //"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp"
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(Sql);

            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

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
        private string GetShipDocEntryJoy(string DocEntry)
        {
            string Sql = "select BoardCount from WH_main where shippingCode='{0}'";
            Sql = string.Format(Sql, DocEntry);
            System.Data.DataTable dt = GetData(Sql);//"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp"

            if (dt.Rows.Count > 0)
            {
                return Convert.ToString(dt.Rows[0][0]);
            }
            else
            {
                return "";
            }

        }

        private string GetArriveDayJoy(string DocEntry)
        {
            string Sql = "select ArriveDay from SHIPPING_MAIN where shippingCode='{0}'";
            Sql = string.Format(Sql, DocEntry);
            System.Data.DataTable dt = GetData(Sql);//"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp"

            if (dt.Rows.Count > 0)
            {
                return Convert.ToString(dt.Rows[0][0]);
            }
            else
            {
                return "";
            }

        }
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
        public DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }
        private string GetShipModel(string ItemCode)
        {
            string SapConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            string sql = "select U_Model from oitm WHERE ItemCode LIKE '{0}'";
            sql = string.Format(sql, ItemCode);
            System.Data.DataTable dt = GetData(SapConnectiongString, sql);//"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02"

            return Convert.ToString(dt.Rows[0][0]);

        }
        private string GetShipMark(string DocEntry, string LineNum)
        {
            string SapConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            string Sql = "select u_mark from rdr1 where docentry={0} and linenum={1}";
            Sql = string.Format(Sql, DocEntry, LineNum);
            System.Data.DataTable dt = GetData(SapConnectiongString, Sql);//"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02"

            if (dt.Rows.Count > 0)
            {
                return Convert.ToString(dt.Rows[0]["u_mark"]).Replace("MSI PO#", "");
            }
            else
            {
                return "";
            }

        }
        private string GetExePath()
        {
            string path = "\\\\ACMESRV01\\Public\\Users\\NessChou\\AcmeMarkXls\\AcmeBarCodePdf\\bin\\Debug";
            return path;
        }
        public bool ExcelCyberBox(System.Data.DataTable dt,
string Template, string SaveFileName, int Interval = 6, Int32 PageBreak = 4, string EndCell = "A5")
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                //{
                //    MessageBox.Show(nm.NameLocal);
                //}


                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;


                //依  dt 筆數產生分頁
                Int32 PrintQty = 1;

                DataRow dr;

                Microsoft.Office.Interop.Excel.Range cell = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //複製範本
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);

                    //xlSht.Copy(Type.Missing, xlWb.Sheets[xlWb.Sheets.Count]); // copy
                    //xlWb.Sheets[xlWb.Sheets.Count].Name = "NEW SHEET";
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i+1];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    excelSheet.Name = "Q" + (i + 1).ToString();

                    //處理範本
                    string BarcodeFile = "";
                    int shapeCount = excelSheet.Shapes.Count;

                    string[] shapeNames = new string[shapeCount];

                    for (int iShape = 1; iShape <= shapeCount; iShape++)
                    {
                        shapeNames[iShape - 1] = excelSheet.Shapes.Item(iShape).Name;
                    }

                    for (int iShape = 0; iShape <= shapeNames.Length - 1; iShape++)
                    {
                        //  string ShapeName = SheetTemplate.Shapes.Item(iShape).Name;
                        string ShapeName = shapeNames[iShape];
                        if (ShapeName == "Logo" || ShapeName == "Group") continue;
                        //if (ShapeName == "QrCode")
                        //{
                        //    GetQrCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        //    BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        //    UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                        //    continue;
                        //}
                        //更換圖片
                        GetBarCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                    }


                    foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                    {
                        //MessageBox.Show(nm.NameLocal);
                        string FieldName = nm.NameLocal;

                        try
                        {
                            cell = excelSheet.Evaluate(FieldName) as Microsoft.Office.Interop.Excel.Range;
                            if (cell != null) cell.Value = Convert.ToString(dr[FieldName]);
                        }
                        catch
                        {

                        }
                    }


                }


                SheetTemplate.Delete();

                excelworkBook.SaveAs(SaveFileName); ;
                excelworkBook.Close();
                //excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (excelSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(SheetTemplate);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                excelSheet = null;
                SheetTemplate = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
            }
        }
        public bool ExcelCyberPallet(System.Data.DataTable dt, string Template, string SaveFileName,
            int Interval = 6, Int32 PageBreak = 4, string EndCell = "A5")
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;

            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                //{
                //    MessageBox.Show(nm.NameLocal);
                //}


                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;


                //依  dt 筆數產生分頁
                Int32 PrintQty = 1;

                DataRow dr;

                Microsoft.Office.Interop.Excel.Range cell = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //複製範本
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);

                    //xlSht.Copy(Type.Missing, xlWb.Sheets[xlWb.Sheets.Count]); // copy
                    //xlWb.Sheets[xlWb.Sheets.Count].Name = "NEW SHEET";
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i+1];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    excelSheet.Name = "Pallet" + (i + 1).ToString();

                    //處理範本
                    string BarcodeFile = "";
                    int shapeCount = excelSheet.Shapes.Count;

                    string[] shapeNames = new string[shapeCount];

                    for (int iShape = 1; iShape <= shapeCount; iShape++)
                    {
                        shapeNames[iShape - 1] = excelSheet.Shapes.Item(iShape).Name;
                    }

                    for (int iShape = 0; iShape <= shapeNames.Length - 1; iShape++)
                    {
                        //  string ShapeName = SheetTemplate.Shapes.Item(iShape).Name;
                        string ShapeName = shapeNames[iShape];

                        if (ShapeName == "Logo" || ShapeName == "Group")

                            continue;

                        //if (ShapeName == "QrCode")
                        //{
                        //    GetQrCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        //    BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        //    UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                        //    continue;
                        //}
                        //更換圖片
                        GetBarCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                    }


                    foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                    {
                        //MessageBox.Show(nm.NameLocal);
                        string FieldName = nm.NameLocal;

                        try
                        {
                            cell = excelSheet.Evaluate(FieldName) as Microsoft.Office.Interop.Excel.Range;
                            if (cell != null)
                            {
                                cell.Value = Convert.ToString(dr[FieldName]);
                            }
                        }
                        catch
                        {

                        }
                    }


                }


                SheetTemplate.Delete();

                excelworkBook.SaveAs(SaveFileName);
                excelworkBook.Close();
                //excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (excelSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(SheetTemplate);


                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                SheetTemplate = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
            }
        }
        private void GetBarCode(string CodeName, string Data)
        {
            string Url = "https://generator.barcodetools.com/barcode.png?gen=0&data={0}&bcolor=FFFFFF&fcolor=000000&tcolor=000000&fh=14&bred=0&w2n=2.5&xdim=2&w=&h=120&debug=1&btype=7&angle=0&quiet=1&balign=2&talign=2&guarg=1&text=1&tdown=1&stst=1&schk=0&cchk=1&ntxt=1&c128=0";

            Url = string.Format(Url, Data);
            string PicFile = GetExePath() + "\\Output\\" + CodeName + ".jpg";
            GetUrlPicture(Url, PicFile);
        }
        private void UpdatePicture(Microsoft.Office.Interop.Excel.Worksheet excelSheet, string ShapeName, string QrFileName)
        {
            //取得來源圖片的位置大小
            float iLeft = 0;
            float iTop = 0;
            float iWidth = 0;
            float iHeight = 0;

            Shape x = excelSheet.Shapes.Item(ShapeName);

            iLeft = x.Left;
            iTop = x.Top;
            iWidth = x.Width;
            iHeight = x.Height;

            x.Delete();

            x = excelSheet.Shapes.AddPicture(QrFileName,
                Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue, iLeft, iTop, iWidth, iHeight);

            x.Name = ShapeName;
        }
        private void ClearPicture(Microsoft.Office.Interop.Excel.Worksheet excelSheet, string ShapeName)
        {
            //刪掉特定名稱圖片
            Shape x = excelSheet.Shapes.Item(ShapeName);
            x.Delete();
        }
        private void GetUrlPicture(string Url, string PicName)
        {
            WebResponse response = default(WebResponse);
            Stream remoteStream = default(Stream);
            StreamReader readStream = default(StreamReader);
            WebRequest request = WebRequest.Create(Url);
            response = request.GetResponse();
            remoteStream = response.GetResponseStream();
            readStream = new StreamReader(remoteStream);

            System.Drawing.Image img = System.Drawing.Image.FromStream(remoteStream);

            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                byte[] byteImage = ms.ToArray();
                // imgBarCode.ImageUrl = "data:image/png;base64," + Convert.ToBase64String(byteImage);
            }
            img.Save(PicName);

            response.Close();
            remoteStream.Close();
            readStream.Close();
        }

        private void btnYearMonth_Click(object sender, EventArgs e)
        {
            string sYear = txtYear.Text;
            string sMonth = txtMonth.Text;

            //Cyber_年月標籤
            string Dir = GetExePath() + "\\Output\\";
            string DirTemplate = GetExePath() + "\\Excel\\Mark\\";

            if (!Directory.Exists(Dir))
            {
                Directory.CreateDirectory(Dir);
            }

            if (!Directory.Exists(DirTemplate))
            {
                Directory.CreateDirectory(DirTemplate);
            }


            //for (int i = 1; i <= 12; i++)
            // {
            // txtMonth.Text = i.ToString();
            //  sMonth = txtMonth.Text;

            System.Data.DataTable dtData = MakeTableCyber();
            DataRow dr = dtData.NewRow();
            dr["sYear"] = txtYear.Text;
            dr["sMonth"] = txtMonth.Text;
            dtData.Rows.Add(dr);


            string FileName = GetExePath() + "\\" + string.Format("Cyber_{0}年{1}月標籤.xls", sYear, sMonth);
            string Template = GetExePath() + "\\Excel\\Mark\\" + "Cyber_年月標籤1.xls";

            if (sMonth == "7" || sMonth == "8" || sMonth == "9"
                || sMonth == "10" || sMonth == "11" || sMonth == "12")
            {
                Template = GetExePath() + "\\Excel\\Mark\\" + "Cyber_年月標籤2.xls";
            }

            ExcelCyber(dtData, Template, FileName);
        }
        private System.Data.DataTable MakeTableCyber()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("sYear", typeof(string));
            dt.Columns.Add("sMonth", typeof(string));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["格式"];
            //dt.PrimaryKey = colPk;
            dt.TableName = "TmpTable";

            return dt;
        }
        public bool ExcelCyber(System.Data.DataTable dt,
string Template, string SaveFileName, int Interval = 6, Int32 PageBreak = 4, string EndCell = "A5")
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                //{
                //    MessageBox.Show(nm.NameLocal);
                //}

                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;

                //依  dt 筆數產生分頁
                Int32 PrintQty = 1;

                DataRow dr;

                Microsoft.Office.Interop.Excel.Range cell = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //複製範本
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);

                    //xlSht.Copy(Type.Missing, xlWb.Sheets[xlWb.Sheets.Count]); // copy
                    //xlWb.Sheets[xlWb.Sheets.Count].Name = "NEW SHEET";
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i+1];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    excelSheet.Name = "Q" + (i + 1).ToString();

                    //處理範本
                    string BarcodeFile = "";
                    int shapeCount = excelSheet.Shapes.Count;

                    string[] shapeNames = new string[shapeCount];

                    for (int iShape = 1; iShape <= shapeCount; iShape++)
                    {
                        shapeNames[iShape - 1] = excelSheet.Shapes.Item(iShape).Name;
                    }

                    for (int iShape = 0; iShape <= shapeNames.Length - 1; iShape++)
                    {
                        //  string ShapeName = SheetTemplate.Shapes.Item(iShape).Name;
                        string ShapeName = shapeNames[iShape];

                        if (ShapeName == "Logo" || ShapeName == "Group") continue;

                        //if (ShapeName == "QrCode")
                        //{
                        //    GetQrCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        //    BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        //    UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                        //    continue;
                        //}
                        //更換圖片
                        GetBarCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                    }


                    foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                    {
                        //MessageBox.Show(nm.NameLocal);
                        string FieldName = nm.NameLocal;

                        try
                        {
                            cell = excelSheet.Evaluate(FieldName) as Microsoft.Office.Interop.Excel.Range;
                            if (cell != null) cell.Value = Convert.ToString(dr[FieldName]);
                        }
                        catch
                        {

                        }
                    }

                    string sMonth = txtMonth.Text;
                    Color mColor = System.Drawing.Color.FromArgb(105, 63, 35); ;
                    //Change Color
                    //RGB
                    if (sMonth == "1" || sMonth == "7")
                    {
                        //Pantone 469 C ->105 63 35 
                        mColor = System.Drawing.Color.FromArgb(105, 63, 35);
                    }
                    else if (sMonth == "2" || sMonth == "8")
                    {
                        //Pantone Red 032  C -> 239 51 64 
                        mColor = System.Drawing.Color.FromArgb(239, 51, 64);
                    }
                    else if (sMonth == "3" || sMonth == "9")
                    {
                        //Pantone 527 C -> 128 49 167 
                        mColor = System.Drawing.Color.FromArgb(128, 49, 167);
                    }
                    else if (sMonth == "4" || sMonth == "10")
                    {
                        //Pantone Yellow C ->254 221 0 
                        mColor = System.Drawing.Color.FromArgb(254, 221, 0);
                    }
                    else if (sMonth == "5" || sMonth == "11")
                    {
                        //PANTONE 360 C ->108 194 74 
                        mColor = System.Drawing.Color.FromArgb(108, 194, 74);
                    }
                    else if (sMonth == "6" || sMonth == "12")
                    {
                        //Pantone 300C -> 0 94 184
                        mColor = System.Drawing.Color.FromArgb(0, 94, 184);
                    }

                    //Excel.Range rng2 = this.Application.get_Range("A1");
                    //rng2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    //FromArgb
                    //ChartRange.Interior.Color = System.Drawing.Color.FromArgb(255, 0, 0);

                    Range rColor = excelSheet.get_Range("B1", "B10") as Range;
                    rColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(mColor);
                    rColor = excelSheet.get_Range("D1", "D10") as Range;
                    rColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(mColor);

                }


                SheetTemplate.Delete();

                excelworkBook.SaveAs(SaveFileName); ;
                excelworkBook.Close();
                //excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (excelSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(SheetTemplate);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                excelSheet = null;
                SheetTemplate = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
            }
        }
        public System.Data.DataTable GETS1(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(MAX(SEQNO),0) SEQNO FROM WH_ITEM4 WHERE SHIPPINGCODE=@SHIPPINGCODE  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        private void button54_Click(object sender, EventArgs e)
        {
            if (shipping_ItemDataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow row;

                StringBuilder sb = new StringBuilder();
                for (int i = shipping_ItemDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = shipping_ItemDataGridView.SelectedRows[i];

                    string DOCENTRY1 = row.Cells["Docentry1"].Value.ToString();
                    sb.Append("'" + DOCENTRY1 + "',");
                }

                sb.Remove(sb.Length - 1, 1);


                string DOCENTRY = "";
                int LINENUM = 0;
                int SeqNo = 0;
                string ShipDate = "";
                string U_PAY = "";
                string U_SHIPDAY = "";
                string U_SHIPSTATUS = "";
                string U_MARK = "";
                string U_MEMO = "";
                string PO = "";
                string LOCATION = "";
                string UNIT = "";
                string WHNO = "";
                int SM = 0;
                int SM2 = 0;
                string ITEMCODE = "";
                string TREETYPE = "";
                System.Data.DataTable SI2 = GetSIF(shippingCodeTextBox.Text, sb.ToString(), "1");
                if (SI2.Rows.Count > 0)
                {
                    for (int i2 = 0; i2 <= SI2.Rows.Count - 1; i2++)
                    {
                        SM++;
                        ITEMCODE = SI2.Rows[i2]["ITEMCODE"].ToString();
                        string Dscription = SI2.Rows[i2]["Dscription"].ToString();
                        string WHSCODE = SI2.Rows[i2]["WHSCODE"].ToString();
                        SeqNo = Convert.ToInt16(SI2.Rows[i2]["SeqNo"]);
                        LINENUM = Convert.ToInt16(SI2.Rows[i2]["LINENUM"]);
                        int QTY = Convert.ToInt16(SI2.Rows[i2]["Quantity"]);

                        WHNO = SI2.Rows[i2]["WHNO"].ToString();

                        DOCENTRY = SI2.Rows[i2]["DOCENTRY"].ToString();
                        System.Data.DataTable SI3 = GetSI3(ITEMCODE, WHSCODE);
                        string PARTNO = "";
                        int ONHAND = 0;
                        string VER = "";
                        string GRADE = "";
                        string FRGNAME = "";
                        string WHSNAME = "";
                        if (SI3.Rows.Count > 0)
                        {
                            PARTNO = SI3.Rows[0]["PARTNO"].ToString();
                            ONHAND = Convert.ToInt32(SI3.Rows[0]["ONHAND"]);
                            VER = SI3.Rows[0]["版本"].ToString();
                            GRADE = SI3.Rows[0]["等級"].ToString();
                            FRGNAME = SI3.Rows[0]["品名規格"].ToString();
                            WHSNAME = SI3.Rows[0]["WHSNAME"].ToString();
                        }

                        System.Data.DataTable SI4 = GetSI4(DOCENTRY, LINENUM);


                        if (SI4.Rows.Count > 0)
                        {

                            ShipDate = SI4.Rows[0]["排程日期"].ToString();
                            U_PAY = SI4.Rows[0]["付款"].ToString();
                            U_SHIPDAY = SI4.Rows[0]["押出貨日"].ToString();
                            U_SHIPSTATUS = SI4.Rows[0]["貨況"].ToString();
                            U_MARK = SI4.Rows[0]["特殊嘜頭"].ToString();

                            U_MEMO = SI4.Rows[0]["注意事項"].ToString();
                            PO = SI4.Rows[0]["PO"].ToString();
                            LOCATION = SI4.Rows[0]["產地"].ToString();

                            UNIT = SI4.Rows[0]["單位"].ToString();
                        }
                        System.Data.DataTable SHIR2 = GetSHIR1(DOCENTRY, LINENUM);

                        if (SHIR2.Rows.Count > 0)
                        {
                            string P1S = SHIR2.Rows[0][0].ToString();
                            if (P1S == "S")
                            {
                                PARTNO = "母料號";
                            }
                            TREETYPE = P1S;
                        }


                        int GE = Convert.ToInt32(GETS1(WHNO).Rows[0][0]);
                        AddSHIPITEM(WHNO, GE + SM, DOCENTRY, LINENUM, "銷售訂單", ITEMCODE, Dscription, Convert.ToInt32(QTY), PARTNO, ONHAND, VER, GRADE, FRGNAME, WHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, LOCATION, UNIT, TREETYPE);


                        System.Data.DataTable SHIR1 = GetSHIR1(DOCENTRY, LINENUM);
                        string P1 = SHIR1.Rows[0][0].ToString();
                        if (P1 == "S")
                        {
                            if (SHIR1.Rows.Count > 0)
                            {
                                System.Data.DataTable BSHIR2 = GetSHIR2(DOCENTRY);
                                if (BSHIR2.Rows.Count > 0)
                                {
                                    for (int Bi2 = 0; Bi2 <= BSHIR2.Rows.Count - 1; Bi2++)
                                    {
                                        SM2++;
                                        int SLINENUM = Convert.ToInt16(BSHIR2.Rows[Bi2]["LineNum"]);
                                        string SITEMCODE = BSHIR2.Rows[Bi2]["ItemCode"].ToString();
                                        string SDESC = BSHIR2.Rows[Bi2]["Dscription"].ToString();
                                        string SQTY = BSHIR2.Rows[Bi2]["QTY"].ToString();
                                        System.Data.DataTable BSI3 = GetSI3(SITEMCODE, WHSCODE);

                                        string SPARTNO = "";
                                        int SONHAND = 0;
                                        string SVER = "";
                                        string SGRADE = "";
                                        string SFRGNAME = "";
                                        string SWHSNAME = "";
                                        string SLOCATION = "";
                                        string SUNIT = "";
                                        if (BSI3.Rows.Count > 0)
                                        {
                                            SUNIT = BSI3.Rows[0]["單位"].ToString();
                                            SLOCATION = BSI3.Rows[0]["產地"].ToString();
                                            SONHAND = Convert.ToInt32(BSI3.Rows[0]["ONHAND"]);
                                            SVER = BSI3.Rows[0]["版本"].ToString();
                                            SGRADE = BSI3.Rows[0]["等級"].ToString();
                                            SFRGNAME = BSI3.Rows[0]["品名規格"].ToString();
                                            SWHSNAME = BSI3.Rows[0]["WHSNAME"].ToString();
                                        }

                                        SPARTNO = ITEMCODE + "-子料號-" + SM.ToString();
                                        AddSHIPITEM(WHNO, GE + SM + SM2, DOCENTRY, SLINENUM, "銷售訂單", SITEMCODE, SDESC, Convert.ToInt32(SQTY), SPARTNO, SONHAND, SVER, SGRADE, SFRGNAME, SWHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, SLOCATION, SUNIT, "I");
                                    }
                                }
                            }
                        }
                    }
                }



                System.Data.DataTable SI2H = GetSIF(shippingCodeTextBox.Text, sb.ToString(), "2");
                if (SI2H.Rows.Count > 0)
                {
                    string WHSCODE1 = SI2H.Rows[0]["WHSCODE"].ToString();
                    string DOC = SI2H.Rows[0]["DOCENTRY"].ToString();
                    System.Data.DataTable SI3S = GetSI1S(WHSCODE1);
                    string WHSNAMES = SI3S.Rows[0][0].ToString();
                    string NumberName = "WH" + DateTime.Now.ToString("yyyyMMdd");
                    SqlConnection Connection = globals.Connection;
                    string AutoNum = util.GetAutoNumber(Connection, NumberName);

                    string KK = NumberName + AutoNum + "X";

                    string username = "";



                    System.Data.DataTable ff = GetOHEM(fmLogin.LoginID.ToString());
                    if (ff.Rows.Count > 0)
                    {
                        DataRow drw = ff.Rows[0];
                        username = drw["姓名"].ToString();
                        username = username.Replace("(", "");
                        username = username.Replace(")", "");

                    }

                    System.Data.DataTable dt1sar = GetMenu.Getocrdnew2(DOC, "銷售訂單");
                    DataRow drwF = dt1sar.Rows[0];
                    string OBUShip = drwF["shipbuilding"].ToString() +
                              Environment.NewLine + drwF["shipstreet"].ToString() +
                              Environment.NewLine + "TEL:" + drwF["shipblock"].ToString() +
                              Environment.NewLine + "FAX:" + drwF["shipcity"].ToString() +
                              Environment.NewLine + "ATTN:" + drwF["shipzipcode"].ToString();


                    string OBUBill = drwF["billbuilding"].ToString() +
             Environment.NewLine + drwF["billstreet"].ToString() +
             Environment.NewLine + "TEL:" + drwF["billblock"].ToString() +
             Environment.NewLine + "FAX:" + drwF["billcity"].ToString() +
             Environment.NewLine + "ATTN:" + drwF["billzipcode"].ToString();
                    //電話號碼

                    System.Data.DataTable D1 = GetOrderDataS(DOC);
                    string 業務 = D1.Rows[0]["業務"].ToString();
                    string 電話號碼 = D1.Rows[0]["電話號碼"].ToString();
                    string 工廠地址 = D1.Rows[0]["工廠地址"].ToString();
                    string 連絡人 = D1.Rows[0]["連絡人"].ToString();

                    int g = 工廠地址.IndexOf("司");


                    if (g != -1)
                    {
                        工廠地址 = 工廠地址.Substring(g + 1).Trim();

                    }
                    string quantity = "";
                    if (dOCTYPETextBox.Text == "銷售訂單" || dOCTYPETextBox.Text == "銷售")
                    {
                        System.Data.DataTable G1 = GetORDR();


                        if (G1.Rows.Count > 0)
                        {
                            string Doc = G1.Rows[0][0].ToString();
                            string LINE = G1.Rows[0][1].ToString();
                            System.Data.DataTable SHIPDATE = GetSHIPDATE(Doc, LINE);
                            if (SHIPDATE.Rows.Count > 0)
                            {
                                quantity = SHIPDATE.Rows[0][2].ToString();
                            }
                        }
                    }
                    else if (dOCTYPETextBox.Text == "調撥單" || dOCTYPETextBox.Text == "調撥")
                    {
                        System.Data.DataTable K1 = GETARRIVE2(mEMO3TextBox.Text);
                        if (K1.Rows.Count > 0)
                        {
                            quantity = K1.Rows[0][0].ToString();
                        }
                    }





                    AddSHIPMAIN(KK, cardNameTextBox.Text, cardCodeTextBox.Text, DOC, dOCTYPETextBox.Text, WHSNAMES, 業務, 電話號碼, username, OBUBill, OBUShip, boardCountNoTextBox.Text, quantity);
                    MessageBox.Show("上傳成功 倉管單號 : " + KK);

                    UPDATESHIPWHNO2(KK, shippingCodeTextBox.Text, sb.ToString());
                    shipping_ItemTableAdapter.Fill(ship.Shipping_Item, MyID);

                    for (int i2 = 0; i2 <= SI2H.Rows.Count - 1; i2++)
                    {
                        SM++;
                        ITEMCODE = SI2H.Rows[i2]["ITEMCODE"].ToString();
                        string Dscription = SI2H.Rows[i2]["Dscription"].ToString();
                        string WHSCODEF = SI2H.Rows[i2]["WHSCODE"].ToString();
                        SeqNo = Convert.ToInt16(SI2H.Rows[i2]["SeqNo"]);
                        LINENUM = Convert.ToInt16(SI2H.Rows[i2]["LINENUM"]);
                        int QTY = Convert.ToInt16(SI2H.Rows[i2]["Quantity"]);

                        DOCENTRY = SI2H.Rows[i2]["DOCENTRY"].ToString();
                        System.Data.DataTable SI3 = GetSI3(ITEMCODE, WHSCODEF);
                        string PARTNO = "";
                        int ONHAND = 0;
                        string VER = "";
                        string GRADE = "";
                        string FRGNAME = "";
                        string WHSNAME = "";
                        if (SI3.Rows.Count > 0)
                        {
                            PARTNO = SI3.Rows[0]["PARTNO"].ToString();
                            ONHAND = Convert.ToInt32(SI3.Rows[0]["ONHAND"]);
                            VER = SI3.Rows[0]["版本"].ToString();
                            GRADE = SI3.Rows[0]["等級"].ToString();
                            FRGNAME = SI3.Rows[0]["品名規格"].ToString();
                            WHSNAME = SI3.Rows[0]["WHSNAME"].ToString();
                        }

                        System.Data.DataTable SI4 = GetSI4(DOCENTRY, LINENUM);


                        if (SI4.Rows.Count > 0)
                        {

                            ShipDate = SI4.Rows[0]["排程日期"].ToString();
                            U_PAY = SI4.Rows[0]["付款"].ToString();
                            U_SHIPDAY = SI4.Rows[0]["押出貨日"].ToString();
                            U_SHIPSTATUS = SI4.Rows[0]["貨況"].ToString();
                            U_MARK = SI4.Rows[0]["特殊嘜頭"].ToString();

                            U_MEMO = SI4.Rows[0]["注意事項"].ToString();
                            PO = SI4.Rows[0]["PO"].ToString();
                            LOCATION = SI4.Rows[0]["產地"].ToString();

                            UNIT = SI4.Rows[0]["單位"].ToString();
                        }
                        System.Data.DataTable SHIR2 = GetSHIR1(DOCENTRY, LINENUM);

                        if (SHIR2.Rows.Count > 0)
                        {
                            string P1S = SHIR2.Rows[0][0].ToString();
                            if (P1S == "S")
                            {
                                PARTNO = "母料號";
                            }
                            TREETYPE = P1S;
                        }



                        int GE = Convert.ToInt32(GETS1(WHNO).Rows[0][0]);
                        AddSHIPITEM(KK, GE + SM, DOCENTRY, LINENUM, "銷售訂單", ITEMCODE, Dscription, Convert.ToInt32(QTY), PARTNO, ONHAND, VER, GRADE, FRGNAME, WHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, LOCATION, UNIT, TREETYPE);

                        System.Data.DataTable SHIR1 = GetSHIR1(DOCENTRY, LINENUM);
                        string P1 = SHIR1.Rows[0][0].ToString();
                        if (P1 == "S")
                        {
                            if (SHIR1.Rows.Count > 0)
                            {
                                System.Data.DataTable BSHIR2 = GetSHIR2(DOCENTRY);
                                if (BSHIR2.Rows.Count > 0)
                                {
                                    for (int Bi2 = 0; Bi2 <= BSHIR2.Rows.Count - 1; Bi2++)
                                    {
                                        SM++;
                                        int SLINENUM = Convert.ToInt16(BSHIR2.Rows[Bi2]["LineNum"]);
                                        string SITEMCODE = BSHIR2.Rows[Bi2]["ItemCode"].ToString();
                                        string SDESC = BSHIR2.Rows[Bi2]["Dscription"].ToString();
                                        string SQTY = BSHIR2.Rows[Bi2]["QTY"].ToString();
                                        System.Data.DataTable BSI3 = GetSI3(SITEMCODE, WHSCODE1);

                                        string SPARTNO = "";
                                        int SONHAND = 0;
                                        string SVER = "";
                                        string SGRADE = "";
                                        string SFRGNAME = "";
                                        string SWHSNAME = "";
                                        string SLOCATION = "";
                                        string SUNIT = "";
                                        if (BSI3.Rows.Count > 0)
                                        {
                                            SUNIT = BSI3.Rows[0]["單位"].ToString();
                                            SLOCATION = BSI3.Rows[0]["產地"].ToString();
                                            SONHAND = Convert.ToInt32(BSI3.Rows[0]["ONHAND"]);
                                            SVER = BSI3.Rows[0]["版本"].ToString();
                                            SGRADE = BSI3.Rows[0]["等級"].ToString();
                                            SFRGNAME = BSI3.Rows[0]["品名規格"].ToString();
                                            SWHSNAME = BSI3.Rows[0]["WHSNAME"].ToString();
                                        }

                                        SPARTNO = ITEMCODE + "-子料號-" + SM.ToString();
                                        AddSHIPITEM(KK, SeqNo + SM, DOCENTRY, SLINENUM, "銷售訂單", SITEMCODE, SDESC, Convert.ToInt32(SQTY), SPARTNO, SONHAND, SVER, SGRADE, SFRGNAME, SWHSNAME, ShipDate, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, SLOCATION, SUNIT, "I");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        private System.Data.DataTable GETARRIVE2(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT SUBSTRING(ARRIVEDAY,1,4)+'/'+SUBSTRING(ARRIVEDAY,5,2)+'/'+SUBSTRING(ARRIVEDAY,7,2) DDATE FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public static System.Data.DataTable GetSIF(string SHIPPINGCODE, string DOCENTRY1, string PINO)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            if (PINO == "1")
            {

                sb.Append(" AND ISNULL(WHNO,'') <> '' ");
            }
            else
            {
                sb.Append(" AND ISNULL(WHNO,'') = '' ");
            }
            sb.Append("  AND DOCENTRY1 IN ( " + DOCENTRY1 + ") ");
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
        private void sAMEMOTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void shipping_ItemDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void shipping_ItemDataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (shipping_ItemDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Substring(0, 2) == "WH")
            {
                try
                {

                    string MEMOT = shipping_ItemDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    string MEMO = "";
                    int G1 = MEMOT.IndexOf("WH20");
                    string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                    if (G1 != -1)
                    {
                        string[] arrurl = H1.Split(new Char[] { ',' });

                        foreach (string i in arrurl)
                        {
                            MEMO = i.Substring(0, 14);
                            WH_main a = new WH_main();
                            a.PublicString = MEMO;
                            a.Show();
                        }

                    }
                }
                catch { }
            }
        }

        private void boatNameTextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}




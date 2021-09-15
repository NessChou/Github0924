using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;


//atm 需要提供後五碼
namespace ACME
{
    public partial class fmOrgan2Chi : Form
    {

        public static string ConnectiongString = "server=10.10.1.40;pwd=riv@green168;uid=rivagreen;database=CHIComp92";

        
        public static string EEPConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=Acmesqlsp";


        private DataTable dtData;

        string FirstNo;
        string LastNo;

        Int32 gCount = 0;

        public fmOrgan2Chi()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dt = GetPotato();
            dataGridView1.DataSource = dt;
        }

        private DataTable GetPotato()
        {

            SqlConnection connection = new SqlConnection(EEPConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM gb_potato where CreateDate > '20140101' and (ProdID is null or ProdID <>'True' ) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            // command.Parameters.Add(new SqlParameter("@FullName", FullName));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";



            if (globals.GroupID.ToString().Trim() != "EEP")
            {
                ConnectiongString = ConnectiongString.Replace("CHIComp92", "CHIComp02");
            }
          

            //select T0.*,T1.*
            //from Gb_Potato T0
            //inner Join Gb_Potato2 T1 on T0.ID = T1.ID
            //where T0.ID=938

            //成本還沒有取得
            string gOrderNo = "";
            string gInvNo = "";



            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("請選擇資料");
                return;
            }

            if (MessageBox.Show("確定執行嗎？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }



            DataTable dtTest = MakeTable();


            DataTable dtTemp = MakeTableTemp();

            string deptKey = "";
            string dept="";

            DataRow dr;

            Int32 BillNo = 1;

            string ProdID="";
            string ProdName="";
            Int32 Quantity =0 ;
            double Price =0;
            Int32 Amount=0;

            Int32 tmpQuantity = 0;
            double tmpPrice = 0;
            Int32 tmpAmount = 0;


            Int32 dAmount = 0;

                double TaxRate = 0;
                int TaxAmt = 0;

                int Discount = 0;

                Double sPrice = 0;
                int sQuantity = 0;

            Int32 TotalQuantity;
                Int32 TotalAmount;

                int BillDate=0;
                string PreInDate="";

                int RowNO = 0;
                int SerNO = 0;

                string dept_chi="";


                //訂單
                int Flag = 2;

                //誰來處理
                string Maker = "Tiffany";
                string MakerID = "C0007";

                string OrdName = string.Empty;

                string DelAddr = string.Empty;
                string DelMan = string.Empty;
                string DelTel = string.Empty;
                int Qty = 0;
                string PotatoKind;

                string OrderNo;


            //棉花田
                string CustomerID = "TW90144-94";
            

                //預交日期 
                string DelDate = string.Empty;

                //20140901
                //客戶類別
                //特殊通路
                string ClassId = "014";

                //20140912
                string OrderPin = "";

            string BillNO="";

            //業務    //SI16 //柳欣儀

            //預交日期  DAT_REQ

            string BarCodeID;

            DataRow drNew;

            string AddrID = "";
            string CustAddress = DelAddr;
            string LinkMan = DelMan;
            string LinkTelephone = DelTel;
            //取貨日期 -> 視為 出貨日期
            string UserDef1 = string.Empty;
            //指定時段
            string UserDef2 = string.Empty;

            DataTable dtAddress;
            string ZipCode = "";


            //業務人員必須輸入
            //string SalesMan = "SI11"; // 陳那慈
            string SalesMan = "SI18";
            //部門必須輸入 
            string DepartID = "C1"; //生物科技

            string CurrID = "NTD";
            int ExchRate = 1;



            //銷售金額
            int SumBTaxAmt;
            //
            //int TaxType = 0;

            //免稅
            int TaxType = 1;
            //稅
            int SumTax = 0;
            //數量
            int SumQty = Qty;


            //帳月
            int AccMonth ;

            //總計
            int SumAmtATax ;

            //本幣
            int LocalTotal ;
            int LocalTax ;

        
            //結案註記 //已結
            // int BillStatus = 1;
            int BillStatus = 0;

            //客戶訂單編號 //採購單號

            string CustBillNo;


            string Remark = string.Empty;

            //付款方式
            string GatherStyle = string.Empty;
            string GatherOther = "45";

            string TransMark = string.Empty;
            string DueTo;
            string AddressID;

            string dept_chiKey = "";

            string PurchasNo = "";


            //處理空白行

            Int32 DataCount = dtData.Rows.Count-1;
            for (int i = DataCount  ; i>=0 ; i--)
            {
                dr = dtData.Rows[i];
                if (Convert.IsDBNull(dr["採購單號"]))
                {
                    dtData.Rows.Remove(dr);
                }
                else
                {
                    PurchasNo = Convert.ToString(dr["採購單號"]);
                    if (string.IsNullOrEmpty(PurchasNo))
                    {
                        dtData.Rows.Remove(dr);
                    }
                }



            }



            for (int i = 0; i <= dtData.Rows.Count - 1; i++)
            {
                dr = dtData.Rows[i];

                PurchasNo = Convert.ToString(dr["採購單號"]);

                //if (checkBox2.Checked)
                //{
                //    dept = Convert.ToString(dr["需求部門"]);
                //    dept_chi = Convert.ToString(dr["需求部門"]);
                //    BarCodeID = "";
                //    PreInDate = Convert.ToString(dr["交貨日期"]).Replace("/", "");
                //    BillDate = Convert.ToInt32(dr["採購日期"].ToString().Replace("/", ""));
                //    ProdID = "";
                //    ProdName = Convert.ToString(dr["產品名稱(全文)"]);
                //    tmpQuantity = Convert.ToInt32(dr["數量"]);
                //}
                //else
                //{
                    dept = Convert.ToString(dr["出貨部門"]);
                    BarCodeID = Convert.ToString(dr["BarCode"]);
                    PreInDate = Convert.ToString(dr["DAT_REQ"]);
                    dept_chi = Convert.ToString(dr["出貨部門"]);
                    BillDate = Convert.ToInt32(dr["採購日期"]);
                    ProdID = Convert.ToString(dr["barcode"]);
                    ProdName = Convert.ToString(dr["品名"]);
                    tmpQuantity = Convert.ToInt32(dr["採購數量"]);
            //    }


                //新莊新泰門市
                string IndexkeyWord=".";
                if (dept_chi.IndexOf(IndexkeyWord) > 0)
                {
                    dept_chi = dept_chi.Substring(3, dept_chi.Length - 3);
                }


                //產品名稱(全文)

                //換算 ??
                tmpPrice = Convert.ToDouble(dr["單價"]);

                //單筆四捨五入 
                tmpAmount = Convert.ToInt32(dr["金額"]);
                dAmount = Convert.ToInt32( Quantity * Price);

                TaxRate = 0;

                TaxAmt = Convert.ToInt32(Amount * TaxRate);

                Discount = 1;
                //int Flag="";
                //string BillNO="";
                RowNO = RowNO + 1;
                SerNO = SerNO + 1;

                sPrice = Price;
                sQuantity = Quantity;




                if (i == 0)
                {
                    deptKey = dept;
                    dept_chiKey = dept_chi;
                }





                if (i != 0 && dept != deptKey)
                {
                    //Add Master

                    //計算總數量
                    TotalQuantity = Convert.ToInt32(dtTemp.Compute("Sum(Quantity)", null));
                    TotalAmount = Convert.ToInt32(dtTemp.Compute("Sum(Amount)", null));

                    LinkMan = "棉花田-" + dept_chiKey ;
                    dtAddress = GetcomCustAddress(CustomerID, LinkMan);
                    if (dtAddress.Rows.Count > 0)
                    {
                        try
                        {
                            AddrID = Convert.ToString(dtAddress.Rows[0]["AddrID"]);
                            LinkTelephone = Convert.ToString(dtAddress.Rows[0]["Telephone"]);

                            CustAddress = Convert.ToString(dtAddress.Rows[0]["Address"]);
                        }
                        catch
                        {
                            AddrID = "001";
                        }

                    }
                    else
                    {
                        textBox1.Text += LinkMan + "-客戶地址不存在" + "\r\n";
                    }
                    //指定時段
                    UserDef2 = "中午前(9~12小時)";


                    //20140912
                    if (string.IsNullOrEmpty(DelDate))
                    {

                        DelDate = UserDef1;
                    }



                    //TW90144-94
                    //單號
                    BillNO = GetOrderKey(BillDate.ToString());
                    // string CustomerID = "W00002";

                    // 地址
                    AddressID = AddrID;
                    // string ZipCode = "";



                    //業務人員必須輸入
                    //string SalesMan = "SI11"; // 陳那慈
           //         SalesMan = "SI16";
                    //部門必須輸入 
                    DepartID = "C1"; //生物科技

                    CurrID = "NTD";
                    ExchRate = 1;



                    //銷售金額
                    SumBTaxAmt = TotalAmount;
                    //
                    //int TaxType = 0;

                    //免稅
                    TaxType = 1;
                    //稅
                    SumTax = 0;
                    //數量
                    SumQty = TotalQuantity;


                    //帳月
                    AccMonth = Convert.ToInt32(BillDate.ToString().Substring(0, 6));

                    //總計
                    SumAmtATax = SumBTaxAmt + SumTax;

                    //本幣
                    LocalTotal = SumBTaxAmt;
                    LocalTax = SumTax;

                    OrderNo = BillNO;
                    gOrderNo = OrderNo;

                    //結案註記 //已結
                    // int BillStatus = 1;
                    BillStatus = 0;

                    //客戶訂單編號 //採購單號

                    CustBillNo = PurchasNo;


                    Remark = string.Empty;

                    //付款方式
                    GatherStyle = string.Empty;
                    GatherOther = "45";

                    TransMark = string.Empty;

                    GatherStyle = "2";
                    UserDef1 = PreInDate;

                    //帳款歸屬

                    DueTo = CustomerID;

                    Remark += "1.紙箱DM:" + "\r\n";
                    Remark += "2.貨運:" + "\r\n";
                    Remark += "3.實際到貨日:" + "\r\n";
                    Remark += "4.快遞單號:" + "\r\n";
                    Remark += string.Format("5.PO:{0}", CustBillNo) + "\r\n";
                    Remark += "6.付款人:棉花田生機園地股份有限公司" + "\r\n";
                    if (!string.IsNullOrEmpty(OrderPin))
                    {
                        Remark += "7.網訂號碼:" + OrderPin;
                    }
                    Remark += "9.訂購人Email:acmegb-fin@acmegb.com";
                    if (checkBox1.Checked)
                    {
                        AddOrdBillMain(BillDate, CustomerID, AddressID, ZipCode, CustAddress, SalesMan, CurrID, ExchRate, SumBTaxAmt, TaxType, SumTax, SumQty, AccMonth, SumAmtATax, LocalTotal, LocalTax, Flag, BillNO, Maker, MakerID, DepartID,
                            LinkMan, LinkTelephone, CustBillNo, BillStatus,
                             UserDef1, UserDef2, Remark, GatherStyle, GatherOther, DueTo);

                        if (string.IsNullOrEmpty(FirstNo))
                        {
                            FirstNo = BillNO;
                        }
                    }



                    //取得明細資料
                    SerNO = 0;
                    RowNO = 0;
                    for (int k = 0; k <= dtTemp.Rows.Count - 1; k++)
                    {
                        drNew = dtTest.NewRow();
                        drNew["BillNo"] = BillNo.ToString();

                        drNew["ProdID"] = Convert.ToString(dtTemp.Rows[k]["ProdID"]);
                        drNew["ProdName"] = Convert.ToString(dtTemp.Rows[k]["ProdName"]);
                        drNew["Quantity"] = Convert.ToInt32(dtTemp.Rows[k]["Quantity"]);
                        drNew["Price"] = Convert.ToDouble(dtTemp.Rows[k]["Price"]);
                        drNew["Amount"] = Convert.ToInt32(dtTemp.Rows[k]["Amount"]);


                        drNew["TotalQuantity"] = TotalQuantity;
                        drNew["TotalAmount"] = TotalAmount;

                        drNew["Dept"] = deptKey;
                        dtTest.Rows.Add(drNew);


                        ProdID = Convert.ToString(dtTemp.Rows[k]["ProdID"]);

                        ProdName = Convert.ToString(dtTemp.Rows[k]["ProdName"]);

                        Quantity = Convert.ToInt32(dtTemp.Rows[k]["Quantity"]);

                        Price = Convert.ToDouble(dtTemp.Rows[k]["Price"]);

                        Amount = Convert.ToInt32(dtTemp.Rows[k]["Amount"]);
                        dAmount = Convert.ToInt32(Quantity * Price);

                        TaxRate = 0;

                        TaxAmt = Convert.ToInt32(Amount * TaxRate);

                        Discount = 1;
                        //int Flag="";
                        //string BillNO="";
                        RowNO = RowNO + 1;
                        SerNO = SerNO + 1;


                        if (checkBox1.Checked)
                        {
                            //寫入明細
                             AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate);
                        }
                    }





                    //Add Detail
                  //  comboBox1.Items.Add(deptKey);
                    deptKey = dept;
                    dept_chiKey = dept_chi;

                    dtTemp = MakeTableTemp();
                    BillNo++;
                }


                //i == dtData.Rows.Count - 1 -> 最後一筆 ->提前寫入至暫存檔

                //寫入暫存檔-----------------------------------------------------------
                drNew = dtTemp.NewRow();

                //轉換料號
                DataTable dtP = GetProduct(BarCodeID);

                if (dtP.Rows.Count > 0)
                {
                    drNew["ProdID"] = Convert.ToString(dtP.Rows[0]["ProdID"]);
                    drNew["ProdName"] = Convert.ToString(dtP.Rows[0]["ProdName"]);

                }
                else
                {
                    textBox1.Text += BarCodeID + " " + ProdName + "\r\n";

                }




                drNew["Quantity"] = tmpQuantity;
                drNew["Price"] = tmpPrice;
                drNew["Amount"] = tmpAmount;
                dtTemp.Rows.Add(drNew);
                //寫入暫存檔-----------------------------------------------------------

            }

            //Add Master
         //   comboBox1.Items.Add(deptKey);




            TotalQuantity = Convert.ToInt32(dtTemp.Compute("Sum(Quantity)", null));
            TotalAmount = Convert.ToInt32(dtTemp.Compute("Sum(Amount)", null));
            LinkMan = "棉花田-" + dept_chiKey;
            dtAddress = GetcomCustAddress(CustomerID, LinkMan);
            if (dtAddress.Rows.Count > 0)
            {
                try
                {
                    AddrID = Convert.ToString(dtAddress.Rows[0]["AddrID"]);
                    LinkTelephone = Convert.ToString(dtAddress.Rows[0]["Telephone"]);

                    CustAddress = Convert.ToString(dtAddress.Rows[0]["Address"]);
                }
                catch
                {
                    AddrID = "001";
                }

            }
            else
            {
                textBox1.Text += LinkMan + "-客戶地址不存在" + "\r\n";
            }
            //指定時段
            UserDef2 = "中午前(9~12小時)";


            //20140912
            if (string.IsNullOrEmpty(DelDate))
            {

                DelDate = UserDef1;
            }



            //TW90144-94
            //單號
            BillNO = GetOrderKey(BillDate.ToString());
            // string CustomerID = "W00002";

            // 地址
             AddressID = AddrID;
            // string ZipCode = "";



            //業務人員必須輸入
            //string SalesMan = "SI11"; // 陳那慈
             SalesMan = "SI18";
            //部門必須輸入 
            DepartID = "C1"; //生物科技

            CurrID = "NTD";
            ExchRate = 1;



            //銷售金額
            SumBTaxAmt = TotalAmount;
            //
            //int TaxType = 0;

            //免稅
            TaxType = 1;
            //稅
            SumTax = 0;
            //數量
            SumQty = TotalQuantity;


            //帳月
            AccMonth = Convert.ToInt32(BillDate.ToString().Substring(0, 6));

            //總計
            SumAmtATax = SumBTaxAmt + SumTax;

            //本幣
            LocalTotal = SumBTaxAmt;
            LocalTax = SumTax;

            OrderNo = BillNO;
            gOrderNo = OrderNo;

            //結案註記 //已結
            // int BillStatus = 1;
            BillStatus = 0;

            //客戶訂單編號 //採購單號

            CustBillNo = PurchasNo;


            Remark = string.Empty;

            //付款方式
            GatherStyle = string.Empty;
            GatherOther = "45";

            TransMark = string.Empty;

            GatherStyle = "2";
            UserDef1 = PreInDate;

            //帳款歸屬

            DueTo = CustomerID;

            Remark += "1.紙箱DM:" + "\r\n";
            Remark += "2.貨運:" + "\r\n";
            Remark += "3.實際到貨日:" + "\r\n";
            Remark += "4.快遞單號:" + "\r\n";
            Remark += string.Format("5.PO:{0}", CustBillNo) + "\r\n";
            Remark += "6.付款人:棉花田生機園地股份有限公司" + "\r\n";
            if (!string.IsNullOrEmpty(OrderPin))
            {
                Remark += "7.網訂號碼:" + OrderPin;
            }


            if (checkBox1.Checked)
            {
                AddOrdBillMain(BillDate, CustomerID, AddressID, ZipCode, CustAddress, SalesMan, CurrID, ExchRate, SumBTaxAmt, TaxType, SumTax, SumQty, AccMonth, SumAmtATax, LocalTotal, LocalTax, Flag, BillNO, Maker, MakerID, DepartID,
                    LinkMan, LinkTelephone, CustBillNo, BillStatus,
                     UserDef1, UserDef2, Remark, GatherStyle, GatherOther, DueTo);

                LastNo = BillNO;

            }


            //取得明細資料
            SerNO = 0;
            RowNO = 0;
            for (int k = 0; k <= dtTemp.Rows.Count - 1; k++)
            {
                drNew = dtTest.NewRow();
                drNew["BillNo"] = BillNo.ToString();

                drNew["ProdID"] = Convert.ToString(dtTemp.Rows[k]["ProdID"]);
                drNew["ProdName"] = Convert.ToString(dtTemp.Rows[k]["ProdName"]);
                drNew["Quantity"] = Convert.ToInt32(dtTemp.Rows[k]["Quantity"]);
                drNew["Price"] = Convert.ToDouble(dtTemp.Rows[k]["Price"]);
                drNew["Amount"] = Convert.ToInt32(dtTemp.Rows[k]["Amount"]);


                drNew["TotalQuantity"] = TotalQuantity;
                drNew["TotalAmount"] = TotalAmount;

                drNew["Dept"] = deptKey;
                dtTest.Rows.Add(drNew);


                ProdID = Convert.ToString(dtTemp.Rows[k]["ProdID"]);

                ProdName = Convert.ToString(dtTemp.Rows[k]["ProdName"]);

                Quantity = Convert.ToInt32(dtTemp.Rows[k]["Quantity"]);

                Price = Convert.ToDouble(dtTemp.Rows[k]["Price"]);

                Amount = Convert.ToInt32(dtTemp.Rows[k]["Amount"]);
                dAmount = Convert.ToInt32(Quantity * Price);

                TaxRate = 0;

                TaxAmt = Convert.ToInt32(Amount * TaxRate);

                Discount = 1;
                //int Flag="";
                //string BillNO="";
                RowNO = RowNO + 1;
                SerNO = SerNO + 1;


                if (checkBox1.Checked)
                {
                    //寫入明細
                     AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate);
                }
            }


            //Add Detail


            gCount = BillNo;
            dataGridView1.DataSource = dtTest;

            if (checkBox1.Checked == false)
            {

                if (string.IsNullOrEmpty(textBox1.Text))
                {

                    checkBox1.Checked = true;

                    button2.Enabled = true;
                    MessageBox.Show("單數:" + gCount.ToString() + " 檢查作業完成");
                }
                else
                {
                    //tabControl1.TabIndex = 1;
                    tabControl1.SelectedIndex = 1;
                }

            }
            else
            {
                string msg = string.Format("正航單號從 {0}~{1}",FirstNo,LastNo);
                textBox1.Text = msg;
                MessageBox.Show(msg);
            
            }






           // DataGridViewRow row;


           // ///進金生實業股份有限公司
           // //博豐光電股份有限公司
           // //聿豐實業股份有限公司

           // string OrdCom = string.Empty;


           // string OrdName = string.Empty;

           // string DelAddr = string.Empty;
           // string DelMan = string.Empty;
           // string DelTel = string.Empty;
           // int Qty = 0;
           // string PotatoKind;

           // string OrderNo;


           // string CustomerID;
           // int Amount;

           // //預交日期 
           // string DelDate = string.Empty;

           // //20140901
           // //客戶類別
           // //預設網購
           // string ClassId = "009";

           // //20140912
           // string OrderPin = "";

           //// for (int i = 0; i <= dataGridView1.SelectedRows.Count - 1; i++)
           // for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0 ; i--)
           // {
           //     row = dataGridView1.SelectedRows[i];

           //     //MessageBox.Show(Convert.ToString(row.Cells["OrdName"].Value));

           //     //continue;

           //     //label1.Text = i.ToString();
           //     //label1.Refresh();



           //     //訂購人資訊 非收貨人
           //     OrdName = Convert.ToString(row.Cells["OrdName"].Value);

           //     DelAddr = Convert.ToString(row.Cells["DelAddr"].Value);
           //     DelMan = Convert.ToString(row.Cells["DelMan"].Value);
           //     DelTel = Convert.ToString(row.Cells["DelTel"].Value);
           //     DelDate = Convert.ToString(row.Cells["DelDate"].Value);


           //     //20140912
           //     OrderPin = Convert.ToString(row.Cells["OrderPin"].Value); ;


           //     OrdCom = Convert.ToString(row.Cells["OrdCom"].Value);

           //     if (OrdCom == "進金生實業股份有限公司")
           //     {
           //         OrdCom = OrdCom.Substring(0, 3);

           //         OrdName = OrdName + "-" + OrdCom;

           //         ClassId = "010";
           //     }
           //     else if (OrdCom == "博豐光電股份有限公司" || OrdCom == "聿豐實業股份有限公司")
           //     {
           //         OrdCom = OrdCom.Substring(0, 2);

           //         OrdName = OrdName + "-" + OrdCom;

           //         ClassId = "010";
           //     }


           //     string ID = Convert.ToString(row.Cells["ID"].Value);

           //     DataTable dtQ = GetPotatoQty(ID);

           //     try
           //     {
           //         //總數量
           //         Qty = Convert.ToInt32(dtQ.Rows[0]["Qty"]);
           //     }
           //     catch
           //     {

           //     }

           //     //總數量
           //     //Qty = Convert.ToInt32(row.Cells["Qty"].Value);
           //     //總金額
           //     Amount = Convert.ToInt32(row.Cells["Amount"].Value);
           //     //Amount = 530;

           //     PotatoKind = Convert.ToString(row.Cells["PotatoKind"].Value);

           //     //固定的
           //     // CustomerID = "W00006";

           //     //檢查是否存在
           //     DataTable dtGetCustomerByName = GetCustomerByName(OrdName);

           //     ////送貨資訊 - Gb_Friend //20140307 待修正
           //     DataTable dtFriends = GetFriends(ID);


           //     int xFlag = 1;
           //     string xID = "";
           //     string xAddrID = "";
           //     string xAddress = "";
           //     string xLinkMan = "";
           //     string xTelephone = "";


           //     //
           //     string AddrID = "";
           //     string CustAddress = DelAddr;
           //     string LinkMan = DelMan;
           //     string LinkTelephone = DelTel;
           //     //取貨日期 -> 視為 出貨日期
           //     string UserDef1 = string.Empty;
           //     //指定時段
           //     string UserDef2 = string.Empty;


           //     string ZipCode = "";

           //     if (dtFriends.Rows.Count > 0)
           //     {
           //         CustAddress = Convert.ToString(dtFriends.Rows[0]["SAddress"]);

           //         string tmpZipCode = "";
           //         //取郵遞區號
           //         if (CustAddress.Length > 0)
           //         {
           //             for (int w = 0; w <= CustAddress.Length - 1; w++)
           //             {
           //                 if (char.IsDigit(CustAddress[w]))
           //                 {
           //                     tmpZipCode = tmpZipCode + CustAddress[w];
           //                 }
           //                 else
           //                 {
           //                     break;
           //                 }

           //             }

           //             ZipCode = tmpZipCode;

           //             if (!string.IsNullOrEmpty(tmpZipCode))
           //             {
           //                 CustAddress = CustAddress.Replace(tmpZipCode, "");
           //             }
           //         }



           //         LinkMan = Convert.ToString(dtFriends.Rows[0]["SPerson"]);
           //         LinkTelephone = Convert.ToString(dtFriends.Rows[0]["STel"]);

           //         //取貨日期
           //         UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);
           //         //指定時段
           //         UserDef2 = Convert.ToString(dtFriends.Rows[0]["STime"]);
           //     }

           //     //20140912
           //     if (string.IsNullOrEmpty(DelDate))
           //     {

           //         DelDate = UserDef1;
           //     }






           //     if (dtGetCustomerByName.Rows.Count == 0)
           //     {
           //         //新增客戶資料 - 

           //         //全稱 簡稱 整合
           //         CustomerID = GetCustomerKey();
           //         AddcomCustomer(CustomerID, OrdName ,ClassId);
           //         AddcomCustDesc(CustomerID);
           //         AddcomCustTrade(CustomerID);

           //         xFlag = 1;
           //         xID = CustomerID;
           //         //判斷最大一號
           //         //客戶不存在,聯絡資料從 001 三碼起跳
           //         AddrID = "001";

           //         //string xZipCode ="";
           //         //判斷是否為數字
           //         //int i = 0;
           //         //string s = "108";
           //         //bool result = int.TryParse(s, out i); //i now = 108

           //         //char.IsNumber(string s, int index)
           //         //char.IsLetter(string s, int index)





           //         //if (dtFriends.Rows.Count > 0)
           //         //{
           //         //    xAddress = Convert.ToString(dtFriends.Rows[0]["SAddress"]);
           //         //    xLinkMan = Convert.ToString(dtFriends.Rows[0]["SPerson"]);
           //         //    xTelephone = Convert.ToString(dtFriends.Rows[0]["STel"]);

           //         //    ////取貨日期
           //         //    //UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);
           //         //    ////指定時段
           //         //    //UserDef2 = Convert.ToString(dtFriends.Rows[0]["STime"]);
           //         //}


           //         //string xAddress = Convert.ToString(row.Cells["sAddress"].Value); 
           //         //string xLinkMan = Convert.ToString(row.Cells["sPerson"].Value);
           //         //string xTelephone = Convert.ToString(row.Cells["sTel"].Value); ;

           //         //員工 全名 簡稱 先不處理
           //         //寫入地址
           //         AddcomCustAddress(xFlag, xID, AddrID, ZipCode, CustAddress, LinkMan, LinkTelephone);

           //         // MessageBox.Show(ID);
           //         // CustomerID = ID;
           //     }
           //     else
           //     {
           //         CustomerID = Convert.ToString(dtGetCustomerByName.Rows[0]["ID"]);

           //         //判斷 聯絡人 - 地址是否存在

           //         DataTable dtAddress = GetcomCustAddress(CustomerID, LinkMan);

           //         DataTable dtAddrID = GetcomCustAddressID(CustomerID);

           //         try
           //         {
           //             AddrID = (Convert.ToInt32(dtAddrID.Rows[0]["AddrID"]) + 1).ToString("000");
           //         }
           //         catch
           //         {
           //             AddrID = "001";
           //         }

                    
           //         //AddrID = Convert.ToString(dtAddrID.Rows[0]["AddrID"]);


           //         if (dtAddress.Rows.Count == 0)
           //         {
           //             AddcomCustAddress(xFlag, CustomerID, AddrID, ZipCode, CustAddress, LinkMan, LinkTelephone);
           //         }


           //         //AddrID 是否帶入 -> 沒有影響

           //     }


           //     //寫入訂單



           //     //判斷是否有運費
           //     Int32 ShipFee = 0;

           //     try
           //     {
           //         ShipFee = Convert.ToInt32(row.Cells["ShipFee"].Value);
           //     }
           //     catch
           //     {

           //     }



           //     //int BillDate = Convert.ToInt32(DateTime.Now.ToString("yyyyMMdd"));

           //     int BillDate = Convert.ToInt32(row.Cells["CreateDate"].Value);


           //     //單號
           //     string BillNO = GetOrderKey(BillDate.ToString());
           //     // string CustomerID = "W00002";

           //     // 地址
           //     string AddressID = AddrID;
           //     // string ZipCode = "";



           //     //業務人員必須輸入
           //     //string SalesMan = "SI11"; // 陳那慈
           //     string SalesMan = "SI16"; 
           //     //部門必須輸入 
           //     string DepartID = "C1"; //生物科技

           //     string CurrID = "NTD";
           //     int ExchRate = 1;

           //     //20140731
           //     if (ShipFee > 0)
           //     {
           //         Amount = Amount - 7;
           //     }

           //     //銷售金額
           //     int SumBTaxAmt = Amount;
           //     //
           //     //int TaxType = 0;

           //     //免稅
           //     int TaxType = 1;
           //     //稅
           //     int SumTax = 0;
           //     //數量
           //     int SumQty = Qty;


           //     //帳月
           //     int AccMonth = Convert.ToInt32(BillDate.ToString().Substring(0, 6));

           //     //總計
           //     int SumAmtATax = SumBTaxAmt + SumTax;















           //     OrderNo = BillNO;
           //     gOrderNo = OrderNo;


           //     //結案註記 //已結
           //     // int BillStatus = 1;
           //     int BillStatus = 0;

           //     //客戶訂單編號
           //     string CustBillNo = Convert.ToString(row.Cells["ID"].Value);


           //     string Remark = string.Empty;

           //     //付款方式
           //     string GatherStyle = string.Empty;
           //     string GatherOther = string.Empty;

           //     string TransMark = string.Empty;

           //     TransMark = Convert.ToString(row.Cells["TransMark"].Value);

           //     if (TransMark == "貨到付款")
           //     {
           //         GatherStyle = "0";
           //     }
           //     //20140904 月結30day -> 月結30days 
           //     else if (TransMark == "月結30days")
           //     {
           //         GatherStyle = "2";
           //     }
           //    // else if (TransMark == "現金" || TransMark == "電匯" || TransMark == "員工付現")
           //     else if (TransMark == "現金" || TransMark == "員工付現")
           //     {
           //         GatherStyle = "3";
           //         GatherOther = "現金";
           //     }
           //     else if (TransMark == "電匯")
           //     {
           //         GatherStyle = "3";
           //         GatherOther = "匯款";
           //     }
           //     //20140912
           //     else if (TransMark == "信用卡付款")
           //     {
           //             GatherStyle = "3";
           //             GatherOther = "信用卡";

           //     }
           //     ////20140812
           //     //else if (TransMark.Length >= 5)
           //     //{
           //     //    if (TransMark.Substring(0, 5) == "信用卡付款")
           //     //    {
           //     //        GatherStyle = "3";
           //     //        GatherOther = "信用卡";
           //     //    }
           //     //}
           //     //else if (TransMark.ToLower() == "atm")
           //     //{
           //     //    GatherStyle = "3";
           //     //    GatherOther = "ATM";
           //     //}
           //     else if (TransMark == "匯款")
           //     {
           //         GatherStyle = "3";
           //         GatherOther = "匯款";
           //     }
           //     else 
           //     {
           //         GatherStyle = "3";
           //         GatherOther = TransMark;
           //     }

           //     //TransMark
           //     //貨到付款
           //     //FOC
           //     //月結30days
           //     //現金
           //     //電匯
           //     //員工付現
           //     //GatherStyle=0 貨到 1次月 2月結 3其他
           //     //CASE A.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END as 付款







           //     //UserDef1 = Convert.ToString(row.Cells["DelDate"].Value);
           //     //UserDef2 = Convert.ToString(row.Cells["Serv"].Value); 
           //     //20140912
           //     UserDef1 = DateToStr(StrToDate(DelDate).AddDays(-1));

           //     //帳款歸屬

           //     string DueTo = CustomerID;

           //     Remark += "1.紙箱DM:" + "\r\n";
           //     Remark += "2.貨運:" + "\r\n";
           //     Remark += "3.實際到貨日:" + "\r\n";
           //     Remark += "4.快遞單號:" + "\r\n";
           //     Remark += string.Format("5.PO:{0}", ID) + "\r\n";
           //     Remark += "6.付款人:"  + "\r\n";
           //     if (!string.IsNullOrEmpty(OrderPin))
           //     {
           //     Remark += "7.網訂號碼:" + OrderPin;
           //     }
                


           //     AddOrdBillMain(BillDate, CustomerID, AddressID, ZipCode, CustAddress, SalesMan, CurrID, ExchRate, SumBTaxAmt, TaxType, SumTax, SumQty, AccMonth, SumAmtATax, LocalTotal, LocalTax, Flag, BillNO, Maker, MakerID, DepartID,
           //         LinkMan, LinkTelephone, CustBillNo, BillStatus,
           //          UserDef1, UserDef2, Remark, GatherStyle, GatherOther, DueTo);




           //     //讀取明細檔


           //     DataTable dtD = GetPotatoDetail(ID);

           //     //MessageBox.Show(BillNo);

           //     int RowNO = 0;
           //     int SerNO = 0;

           //     string PreInDate = string.Empty;
           //     string ProdID = string.Empty;
           //     string ProdName = string.Empty;
           //     int Quantity = 0;

           //     //int Price = 0;
           //     Double Price = 0;

           //     Double dAmount = 0;

           //     double TaxRate = 0;
           //     int TaxAmt = 0;

           //     int Discount = 0;

           //     Double sPrice = 0;
           //     int sQuantity = 0;

           //     for (int dj = 0; dj <= dtD.Rows.Count - 1; dj++)
           //     {

           //         //單筆範例
           //         ProdID = Convert.ToString(dtD.Rows[dj]["ItemCode"]);
           //         ProdName = Convert.ToString(dtD.Rows[dj]["ItemName"]);


           //         //個別的數量及金額
           //         Quantity = Convert.ToInt16(dtD.Rows[dj]["Qty"]);
           //         Price = Convert.ToInt16(dtD.Rows[dj]["Price"]);

           //         dAmount = Quantity * Price;

           //         TaxRate = 0;

           //         TaxAmt = Convert.ToInt32(Amount * TaxRate);

           //         Discount = 1;
           //         //int Flag="";
           //         //string BillNO="";
           //         RowNO = RowNO + 1;
           //         SerNO = SerNO + 1;

           //         sPrice = Price;
           //         sQuantity = Quantity;

           //         PreInDate = DelDate;

           //         //20140829 - 手動輸入模式

           //         //取貨日期
           //         //UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);

           //         if (string.IsNullOrEmpty(PreInDate))
           //         {
           //             PreInDate = UserDef1;
           //         }

           //         AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate);
           //     } //For

           //     if (ShipFee > 0)
           //     {
           //         RowNO = RowNO + 1;
           //         SerNO = SerNO + 1;

           //         ProdID = "FREIGHT01";
           //         ProdName = "Freight 運費";
           //         Quantity = 1;
           //         // Price = 150;
           //         //Price = 142.8571;
           //         Price = 171;

           //         dAmount = Convert.ToInt32(Quantity * Price);

           //         //TaxRate = 0;
           //         TaxRate = 0.05;

           //         // TaxAmt = Convert.ToInt32(Amount * TaxRate);
           //         TaxAmt = 0;

           //         Discount = 1;


           //         sPrice = Price;
           //         sQuantity = Quantity;


           //         AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate);
           //     }


              
           //     //正式區
           //     if (rbDb02.Checked)
           //     {
           //       //回寫 成立

           //         string memo = string.Format("單號-{0} {1}", BillNO, DateTime.Now.ToString("yyyyMMddHHmmss"));
           //         UpdatePotato(ID, memo);

           //     }

           // } //for


           // MessageBox.Show("己轉入筆數:"+dataGridView1.SelectedRows.Count.ToString());

           // //重新整理
           // DataTable dt = GetPotato();
           // dataGridView1.DataSource = dt;

            
        }

        public string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }
        //字串轉日期
        public DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

      

            return new DateTime(Year, Month, Day, 00, 00, 00);
        }

        private DataTable GetPotatoQty(string ID)
        {

            SqlConnection connection = new SqlConnection(EEPConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Sum(Qty) Qty FROM gb_potato2 where ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        /// <summary>
        ///進金生實業股份有限公司
        //博豐光電股份有限公司
        //聿豐實業股份有限公司
        /// </summary>
        /// <param name="FullName"></param>
        /// <returns></returns>
        private DataTable GetCustomerByName(string FullName)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ID FROM comCustomer ");
            sb.Append("Where FullName =@FullName ");

            sb.Append("and  flag=1 ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@FullName", FullName));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private DataTable GetFriends(string ID)
        {

            SqlConnection connection = new SqlConnection(EEPConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("select T0.* ");
            sb.Append("from Gb_Friend T0  ");
            sb.Append("where T0.DOCID = @ID ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private string GetCustomerKey()
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            //sb.Append("SELECT IsNull(Max(ID),'W00000') FROM comCustomer ");
            //sb.Append("where Flag=1 and SUBSTRING(id,1,1)='W'");

            sb.Append("SELECT IsNull(Max(ID),'00000') FROM comCustomer ");
           // sb.Append("where Flag=1 and SUBSTRING(id,1,8)='TW90143-'");
          //  sb.Append("where Flag=1 and SUBSTRING(id,1,8)='TW90143-' and and LEN(id)=11");
            sb.Append("where Flag=1 and SUBSTRING(id,1,8)='TW90143-' and LEN(id)=11");

            

 


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            // command.Parameters.Add(new SqlParameter("@FullName", FullName));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];

            string ID = Convert.ToString(dt.Rows[0][0]);

            string s = "TW90143-";
            string n = ID.Substring(8, ID.Length - 8);

            ID = s + (Convert.ToInt16(n) + 1).ToString();

            //ID = ID.Substring(1, ID.Length - 1);
            //ID = "W"+(Convert.ToInt16(ID) + 1).ToString("00000");


            return ID;

        }

        /// <summary>
        /// 未稅單價 PriceOfTax = True
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="FullName"></param>
        public void AddcomCustomer(string ID, string FullName, string ClassId)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert into comCustomer(Flag,ID ,FundsAttribution,TransNewID,CurrencyID ,FullName,InvoiceHead,ShortName,ClassId) " +
            "values (@Flag,@ID,@FundsAttribution,@TransNewID,@CurrencyID,@FullName,@InvoiceHead,@ShortName,@ClassId)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@FundsAttribution", ID));
            command.Parameters.Add(new SqlParameter("@TransNewID", ID));
            command.Parameters.Add(new SqlParameter("@CurrencyID", "NTD"));
            command.Parameters.Add(new SqlParameter("@FullName", FullName));
            command.Parameters.Add(new SqlParameter("@InvoiceHead", FullName));
            if (FullName.Length <= 4)
            {
                command.Parameters.Add(new SqlParameter("@ShortName", FullName.Substring(0, FullName.Length)));
            }
            else
            {
                command.Parameters.Add(new SqlParameter("@ShortName", FullName.Substring(0, 4)));
            }

            //ClassId
            command.Parameters.Add(new SqlParameter("@ClassId", ClassId));

            //if (String.IsNullOrEmpty(row.資料行))
            //{
            //    command.Parameters"@資料行".IsNullable = true;
            //    command.Parameters"@資料行".Value = "";
            //}
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

        public void AddcomCustDesc(string ID)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert into comCustDesc(Flag,ID ) " +
            "values (@Flag,@ID)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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

        //  string LinkMan = DelMan;
        //string LinkTelephone = DelTel;


        public void AddOrdBillMain(int BillDate, string CustomerID, string AddressID, string ZipCode, string CustAddress, string SalesMan, string CurrID, int ExchRate, int SumBTaxAmt, int TaxType, int SumTax, int SumQty, int AccMonth, int SumAmtATax, int LocalTotal, int LocalTax, int Flag, string BillNO, string Maker, string MakerID,
            string DepartID, string LinkMan, string LinkTelephone, string CustBillNo, int BillStatus,
            string UserDef1, string UserDef2, string Remark, string GatherStyle, string GatherOther, string DueTo)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert Into OrdBillMain (BillDate,CustomerID,AddressID,ZipCode,CustAddress,SalesMan,CurrID,ExchRate,SumBTaxAmt,TaxType,SumTax,SumQty,AccMonth,SumAmtATax,LocalTotal,LocalTax,Flag,BillNO,Maker,MakerID,DepartID,LinkMan,LinkTelephone,BillStatus,CustBillNo,UserDef1, UserDef2 ,Remark,GatherStyle,GatherOther,DueTo,FormalCust,CheckStyle) " +
            "values (@BillDate,@CustomerID,@AddressID,@ZipCode,@CustAddress,@SalesMan,@CurrID,@ExchRate,@SumBTaxAmt,@TaxType,@SumTax,@SumQty,@AccMonth,@SumAmtATax,@LocalTotal,@LocalTax,@Flag,@BillNO,@Maker,@MakerID,@DepartID,@LinkMan,@LinkTelephone,@BillStatus,@CustBillNo,@UserDef1, @UserDef2 ,@Remark,@GatherStyle,@GatherOther,@DueTo,@FormalCust,1)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillDate", BillDate));
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@AddressID", AddressID));
            command.Parameters.Add(new SqlParameter("@ZipCode", ZipCode));
            command.Parameters.Add(new SqlParameter("@CustAddress", CustAddress));
            command.Parameters.Add(new SqlParameter("@SalesMan", SalesMan));
            command.Parameters.Add(new SqlParameter("@CurrID", CurrID));
            command.Parameters.Add(new SqlParameter("@ExchRate", ExchRate));
            command.Parameters.Add(new SqlParameter("@SumBTaxAmt", SumBTaxAmt));
            command.Parameters.Add(new SqlParameter("@TaxType", TaxType));
            command.Parameters.Add(new SqlParameter("@SumTax", SumTax));
            command.Parameters.Add(new SqlParameter("@SumQty", SumQty));
            command.Parameters.Add(new SqlParameter("@SumAmtATax", SumAmtATax));



            command.Parameters.Add(new SqlParameter("@AccMonth", AccMonth));
            command.Parameters.Add(new SqlParameter("@LocalTotal", LocalTotal));
            command.Parameters.Add(new SqlParameter("@LocalTax", LocalTax));

            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@Maker", Maker));
            command.Parameters.Add(new SqlParameter("@MakerID", MakerID));
            command.Parameters.Add(new SqlParameter("@DepartID", DepartID));

            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
            command.Parameters.Add(new SqlParameter("@LinkTelephone", LinkTelephone));
            command.Parameters.Add(new SqlParameter("@CustBillNo", CustBillNo));


            //直接結案
            // command.Parameters.Add(new SqlParameter("@BillStatus", 1));
            //未結
            command.Parameters.Add(new SqlParameter("@BillStatus", BillStatus));


            command.Parameters.Add(new SqlParameter("@UserDef1", UserDef1));
            command.Parameters.Add(new SqlParameter("@UserDef2", UserDef2));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));

            //GatherStyle
            command.Parameters.Add(new SqlParameter("@GatherStyle", GatherStyle));
            //GatherOther
            command.Parameters.Add(new SqlParameter("@GatherOther", GatherOther));
            //DueTo
            command.Parameters.Add(new SqlParameter("@DueTo", DueTo));

            //FormalCust
            command.Parameters.Add(new SqlParameter("@FormalCust", 1));

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


        public void AddOrdBillSub(int BillDate, int SerNO, string ProdID, string ProdName, int Quantity, Double Price, Double Amount, double TaxRate, int TaxAmt, int Discount,
            int Flag, string BillNO, int RowNO, string PreInDate)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert Into OrdBillSub (BillDate,SerNO,ProdID,ProdName,Quantity,Price,Amount,TaxRate,TaxAmt,Discount,Flag,BillNO,RowNO,sQuantity,sPrice,QtyRemain,PreInDate) " +
            "values (@BillDate,@SerNO,@ProdID,@ProdName,@Quantity,@Price,@Amount,@TaxRate,@TaxAmt,@Discount,@Flag,@BillNO,@RowNO,@sQuantity,@sPrice,@QtyRemain,@PreInDate)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillDate", BillDate));
            command.Parameters.Add(new SqlParameter("@SerNO", SerNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@ProdName", ProdName));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Price", Price));

            command.Parameters.Add(new SqlParameter("@sQuantity", Quantity));
            command.Parameters.Add(new SqlParameter("@sPrice", Price));

            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@TaxRate", TaxRate));
            command.Parameters.Add(new SqlParameter("@TaxAmt", TaxAmt));
            command.Parameters.Add(new SqlParameter("@Discount", Discount));
            command.Parameters.Add(new SqlParameter("@RowNO", RowNO));


            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

            //未出數量
            command.Parameters.Add(new SqlParameter("@QtyRemain", Quantity));

            //PreInDate
            command.Parameters.Add(new SqlParameter("@PreInDate", PreInDate));







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


        //20140902
        //-- 32 -1 1144992 1
        //InvoiceType,TaxKind,AccReceivable,InvoiceStyle
        public void AddcomCustTrade(string ID)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert into comCustTrade(Flag,ID,InvoiceType,TaxKind,AccReceivable,InvoiceStyle ) " +
            "values (@Flag,@ID,@InvoiceType,@TaxKind,@AccReceivable,@InvoiceStyle)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            //20140902
            command.Parameters.Add(new SqlParameter("@InvoiceType", "32"));
            command.Parameters.Add(new SqlParameter("@TaxKind", "1"));
            command.Parameters.Add(new SqlParameter("@AccReceivable", "1144992"));
            command.Parameters.Add(new SqlParameter("@InvoiceStyle", "1"));


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

        /// <summary>
        /// 客戶地址檔 = 1
        /// </summary>
        /// <param name="Flag"></param>
        /// <param name="ID"></param>
        /// <param name="AddrID"></param>
        /// <param name="ZipCode"></param>
        /// <param name="Address"></param>
        /// <param name="LinkMan"></param>
        /// <param name="Telephone"></param>
        public void AddcomCustAddress(int Flag, string ID, string AddrID, string ZipCode, string Address, string LinkMan, string Telephone)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);

            SqlCommand command = new SqlCommand("Insert into comCustAddress(Flag,ID,AddrID,ZipCode,Address,LinkMan,Telephone) values(@Flag,@ID,@AddrID,@ZipCode,@Address,@LinkMan,@Telephone)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@AddrID", AddrID));
            command.Parameters.Add(new SqlParameter("@ZipCode", ZipCode));
            command.Parameters.Add(new SqlParameter("@Address", Address));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
            command.Parameters.Add(new SqlParameter("@Telephone", Telephone));



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

        /// <summary>
        /// 客戶聯絡人 - 依 聯絡人  
        /// </summary>
        /// <returns></returns>
        private DataTable GetcomCustAddress(string ID, string LinkMan)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("select * ");
            sb.Append("from comCustAddress  ");
            sb.Append("where ID = @ID ");
            sb.Append("and  LinkMan = @LinkMan ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }


        private DataTable GetcomCustAddressID(string ID)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("select isnull(Max(AddrID),'000') AddrID ");
            sb.Append("from comCustAddress  ");
            sb.Append("where  flag =1 and ID = @ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@ID", ID));
            //            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        /// <summary>
        /// 格式改為 J 103 03 0002
        /// </summary>
        /// <param name="sDate"></param>
        /// <returns></returns>
        private string GetOrderKey(string sDate)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            //2014 - 1911 =103
            string strFix = "J" + Convert.ToString(Convert.ToInt32(sDate.Substring(0, 4)) - 1911) + sDate.Substring(4, 2);

            sb.Append("SELECT IsNull(Max(BillNo),'0000') FROM OrdBillMain ");
            sb.Append(string.Format("where Flag=2 and SUBSTRING(BillNo,1,6)='{0}'", strFix));


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            // command.Parameters.Add(new SqlParameter("@FullName", FullName));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];

            string ID = Convert.ToString(dt.Rows[0][0]);

            if (ID == "0000")
            {
                ID = strFix + "0001";
            }
            else
            {
                ID = ID.Substring(6, 4);
                ID = strFix + (Convert.ToInt16(ID) + 1).ToString("0000");
            }

            return ID;

        }

        private DataTable GetPotatoDetail(string ID)
        {

            SqlConnection connection = new SqlConnection(EEPConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM gb_potato2 where ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }


        public void UpdatePotato(string ID,string memo)
        {
            SqlConnection connection = new SqlConnection(EEPConnectiongString);
            string sql = "update gb_potato set ProdID='True',memo=@memo where ID=@ID";
            
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@memo", memo));
       
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



        private void fmRg2Chi_Load(object sender, EventArgs e)
        {
            //DataTable dt = GetPotato();
            //dataGridView1.DataSource = dt;
        }

        private  DataSet ImportExcelXLS(string FileName, bool hasHeaders)
        {
            //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;
          //  Extended Properties = "Excel 12.0 Xml;HDR=YES";
            string HDR = hasHeaders ? "Yes" : "No";
             string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"";
             if (checkBox2.Checked)
             {
                 strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=" + HDR + ";IMEX=1\"";
             }
              DataSet output = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                conn.Open();

                System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                foreach (DataRow row in dt.Rows)
                {
                    string sheet = row["TABLE_NAME"].ToString();

                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                    cmd.CommandType = CommandType.Text;

                    System.Data.DataTable outputTable = new System.Data.DataTable(sheet);
                    output.Tables.Add(outputTable);
                    new OleDbDataAdapter(cmd).Fill(outputTable);
                }
            }
            return output;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //要用 UTF-8 格式


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = openFileDialog1.FileName;

                DataSet ds = ImportExcelXLS(FileName, true);

                dtData = ds.Tables[0];


                dataGridView1.DataSource = dtData;



            }
        }


        // public void AddOrdBillSub(int BillDate, int SerNO, string ProdID, string ProdName, int Quantity, Double Price, Double Amount, double TaxRate, int TaxAmt, int Discount,
        //    int Flag, string BillNO, int RowNO, string PreInDate)
        //{

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("BillNO", typeof(string));
            dt.Columns.Add("ProdID", typeof(string));
            dt.Columns.Add("ProdName", typeof(string));
            dt.Columns.Add("Quantity", typeof(Int32));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Amount", typeof(Int32));
            dt.Columns.Add("Dept", typeof(string));

            dt.Columns.Add("TotalQuantity", typeof(Int32));
            dt.Columns.Add("TotalAmount", typeof(Int32));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["部位"];
            //dt.PrimaryKey = colPk;


            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }

        private System.Data.DataTable MakeTableTemp()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

          
            dt.Columns.Add("ProdID", typeof(string));
            dt.Columns.Add("ProdName", typeof(string));
            dt.Columns.Add("Quantity", typeof(Int32));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Amount", typeof(Int32));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["部位"];
            //dt.PrimaryKey = colPk;


            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }

        private DataTable GetProduct(string BarCodeID)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append(" select SuggestPrice price,ProdName ,PackUnit1,ProdID ");
            sb.Append(" From comProduct  ");
            sb.Append(" Where BarCodeID = @BarCodeID ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@BarCodeID", BarCodeID));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            button9_Click(sender, e);

            MessageBox.Show("單數:"+gCount.ToString()+ " 轉入作業完成");
        }
    } //p
} //n 
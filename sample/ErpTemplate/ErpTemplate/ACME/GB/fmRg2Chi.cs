using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;


//atm 需要提供後五碼
namespace ACME
{
    public partial class fmRg2Chi : Form
    {
        private string FileName;
        string strCn2 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn3 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public static string ConnectiongString = "server=10.10.1.40;pwd=riv@green168;uid=rivagreen;database=CHIComp92";

        public fmRg2Chi()
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

            SqlConnection connection = globals.Connection;

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
     
            if (rbDb02.Checked)
            {
                ConnectiongString = ConnectiongString.Replace("CHIComp92", "CHIComp02");
            }

            if (rbDb03.Checked)
            {
                ConnectiongString = ConnectiongString.Replace("CHIComp02", "CHIComp92");
            }



            //成本還沒有取得
            string gOrderNo = "";
            string gInvNo = "";



            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("請查詢及選擇資料");
                return;
            }

            if (MessageBox.Show("確定執行嗎？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            trunOrdBill();
            DataGridViewRow row;


            ///進金生實業股份有限公司
            //博豐光電股份有限公司
            //聿豐實業股份有限公司

            string OrdCom = string.Empty;
            string OrdCom2 = string.Empty;

            string OrdName = string.Empty;

            string DelAddr = string.Empty;
            string DelMan = string.Empty;
            string DelTel = string.Empty;
            string RivaDiscount = string.Empty;
            int Qty = 0;
            string PotatoKind;
            string CUSTTYPE = string.Empty;
            string OrderNo;
            string OrdEmail = "";
            string OrdTEL = "";
            string PROJECT = "";
            string SALES = "";

            string InvNo = "";
            string CustomerID;
            int Amount;
            string RivaNote;
            string UNIT;
            //預交日期 
            string DelDate = string.Empty;

            //20140901
            //客戶類別
            //預設網購
            string ClassId = "009";

            //20140912
            string OrderPin = "";

           // for (int i = 0; i <= dataGridView1.SelectedRows.Count - 1; i++)
            for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0 ; i--)
            {
                row = dataGridView1.SelectedRows[i];

                //MessageBox.Show(Convert.ToString(row.Cells["OrdName"].Value));

                //continue;

                //label1.Text = i.ToString();
                //label1.Refresh();



                //訂購人資訊 非收貨人
                OrdName = Convert.ToString(row.Cells["OrdName"].Value);
    
                DelAddr = Convert.ToString(row.Cells["DelAddr"].Value);
                DelMan = Convert.ToString(row.Cells["DelMan"].Value);
                DelTel = Convert.ToString(row.Cells["DelTel"].Value);
                DelDate = Convert.ToString(row.Cells["DelDate"].Value);
                RivaDiscount = Convert.ToString(row.Cells["RivaDiscount"].Value);
                RivaNote = Convert.ToString(row.Cells["RivaNote"].Value);

                OrdEmail = Convert.ToString(row.Cells["OrdEmail"].Value);
                OrdTEL = Convert.ToString(row.Cells["OrdTEL"].Value);
                OrderPin = Convert.ToString(row.Cells["OrderPin"].Value);

                CUSTTYPE = Convert.ToString(row.Cells["CUSTTYPE"].Value);
                OrdCom = Convert.ToString(row.Cells["OrdCom"].Value);
                OrdCom2 = Convert.ToString(row.Cells["OrdCom"].Value);
                PROJECT = Convert.ToString(row.Cells["PROJECT"].Value);
                SALES = Convert.ToString(row.Cells["SALES"].Value);
                UNIT = Convert.ToString(row.Cells["UNIT"].Value);
                InvNo = Convert.ToString(row.Cells["InvNo"].Value);
                if (OrdCom == "進金生實業股份有限公司")
                {
                    OrdCom = OrdCom.Substring(0, 3);

                    OrdName = OrdName + "-" + OrdCom;

                    ClassId = "010";
                }
                else if (OrdCom == "博豐光電股份有限公司" || OrdCom == "聿豐實業股份有限公司")
                {
                    OrdCom = OrdCom.Substring(0, 2);

                    OrdName = OrdName + "-" + OrdCom;

                    ClassId = "010";
                }


                string ID = Convert.ToString(row.Cells["ID"].Value);

                DataTable dtQ = GetPotatoQty(ID);

                try
                {
                    //總數量
                    Qty = Convert.ToInt32(dtQ.Rows[0]["Qty"]);
                }
                catch
                {

                }


                //總數量
                //Qty = Convert.ToInt32(row.Cells["Qty"].Value);
                //總金額
                Amount = Convert.ToInt32(row.Cells["Amount"].Value);
                //Amount = 530;

                PotatoKind = Convert.ToString(row.Cells["PotatoKind"].Value);

                //固定的
                // CustomerID = "W00006";

                //檢查是否存在
                DataTable dtGetCustomerByName = GetCustomerByName(OrdName,CUSTTYPE);

                ////送貨資訊 - Gb_Friend //20140307 待修正
                DataTable dtFriends = GetFriends(ID);


                int xFlag = 1;
                string xID = "";
                string ADD = "";
               


                //
                string AddrID = "";
                string CustAddress = DelAddr;
                string LinkMan = DelMan;
                string LinkTelephone = DelTel;
                //取貨日期 -> 視為 出貨日期
                string UserDef1 = string.Empty;
                //指定時段
                string UserDef2 = string.Empty;
                string CUSTNO = "";

                string ZipCode = "";

                if (dtFriends.Rows.Count > 0)
                {
                    CustAddress = Convert.ToString(dtFriends.Rows[0]["SAddress"]);

                    string tmpZipCode = "";
                    //取郵遞區號
                    if (CustAddress.Length > 0)
                    {
                        for (int w = 0; w <= CustAddress.Length - 1; w++)
                        {
                            if (char.IsDigit(CustAddress[w]))
                            {
                                tmpZipCode = tmpZipCode + CustAddress[w];
                            }
                            else
                            {
                                break;
                            }

                        }

                        ZipCode = tmpZipCode;

                        if (!string.IsNullOrEmpty(tmpZipCode))
                        {
                            CustAddress = CustAddress.Replace(tmpZipCode, "");
                        }
                    }



                    LinkMan = Convert.ToString(dtFriends.Rows[0]["SPerson"]);
                    LinkTelephone = Convert.ToString(dtFriends.Rows[0]["STel"]);

                    //取貨日期
                    UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);
                    //指定時段
                    UserDef2 = Convert.ToString(dtFriends.Rows[0]["STime"]);
                }

                //20140912
                if (string.IsNullOrEmpty(DelDate))
                {

                    DelDate = UserDef1;
                }






                if (dtGetCustomerByName.Rows.Count == 0)
                {
                    //新增客戶資料 - 

                    //全稱 簡稱 整合
                    CustomerID = GetCustomerKey();

                    AddcomCustomer(CustomerID, OrdName, ClassId, "C0007", OrdEmail, OrdTEL);
                    AddcomCustDesc(CustomerID, PotatoKind);
                        AddcomCustTrade(CustomerID);
                    
                    xFlag = 1;
                    xID = CustomerID;
                    //判斷最大一號
                    //客戶不存在,聯絡資料從 001 三碼起跳
                    AddrID = "001";

                    //string xZipCode ="";
                    //判斷是否為數字
                    //int i = 0;
                    //string s = "108";
                    //bool result = int.TryParse(s, out i); //i now = 108

                    //char.IsNumber(string s, int index)
                    //char.IsLetter(string s, int index)





                    //if (dtFriends.Rows.Count > 0)
                    //{
                    //    xAddress = Convert.ToString(dtFriends.Rows[0]["SAddress"]);
                    //    xLinkMan = Convert.ToString(dtFriends.Rows[0]["SPerson"]);
                    //    xTelephone = Convert.ToString(dtFriends.Rows[0]["STel"]);

                    //    ////取貨日期
                    //    //UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);
                    //    ////指定時段
                    //    //UserDef2 = Convert.ToString(dtFriends.Rows[0]["STime"]);
                    //}


                    //string xAddress = Convert.ToString(row.Cells["sAddress"].Value); 
                    //string xLinkMan = Convert.ToString(row.Cells["sPerson"].Value);
                    //string xTelephone = Convert.ToString(row.Cells["sTel"].Value); ;

                    //員工 全名 簡稱 先不處理
                    //寫入地址
                    if (rbDb02.Checked)
                    {
                        AddcomCustAddress(xFlag, xID, AddrID, ZipCode, CustAddress, LinkMan, LinkTelephone);
                        ADD = "1";
                    }
                    

                    // MessageBox.Show(ID);
                    // CustomerID = ID;
                }
                else
                {
                   
                    CustomerID = Convert.ToString(dtGetCustomerByName.Rows[0]["ID"]);
                    int T1 = CustomerID.IndexOf("-");
                    if (T1 != -1)
                    {
                        CustomerID = Convert.ToString(GetCustomerByName2(OrdName,CUSTTYPE).Rows[0]["ID"]);
                    }


                    DataTable dtAddress = GetcomCustAddress(CustomerID, LinkMan);

                    DataTable dtAddrID = GetcomCustAddressID(CustomerID);

                    try
                    {
                        AddrID = (Convert.ToInt32(dtAddrID.Rows[0]["AddrID"]) + 1).ToString("000");
                    }
                    catch
                    {
                        AddrID = "001";
                    }

                    
                    //AddrID = Convert.ToString(dtAddrID.Rows[0]["AddrID"]);


                    if (dtAddress.Rows.Count == 0)
                    {
                        if (rbDb02.Checked)
                        {
                            AddcomCustAddress(xFlag, CustomerID, AddrID, ZipCode, CustAddress, LinkMan, LinkTelephone);
                            ADD = "1";
                        }
                        
                    }


                    //AddrID 是否帶入 -> 沒有影響

                }


                //寫入訂單



                //判斷是否有運費
                Int32 ShipFee = 0;

                try
                {
                    ShipFee = Convert.ToInt32(row.Cells["ShipFee"].Value);
                }
                catch
                {

                }



                //int BillDate = Convert.ToInt32(DateTime.Now.ToString("yyyyMMdd"));

                int BillDate = Convert.ToInt32(row.Cells["CreateDate"].Value);


                //單號
                string BillNO = GetOrderKey(BillDate.ToString());
                // string CustomerID = "W00002";

                // 地址
                string AddressID = AddrID;
                // string ZipCode = "";



                //業務人員必須輸入
                //string SalesMan = "SI11"; // 陳那慈
                string SalesMan = "C0007";


                if (!string.IsNullOrEmpty(SALES))
                {
                    SalesMan = SALES;
                }
                string DepartID = "C1"; //生物科技

                string CurrID = "NTD";
                int ExchRate = 1;

                //20140731
      

                //銷售金額
                int SumBTaxAmt = Amount;
                //
                //int TaxType = 0;

                //免稅
                int TaxType = 1;
                //稅
                int SumTax = 0;
                //數量
                int SumQty = Qty;


                //帳月
                int AccMonth = Convert.ToInt32(BillDate.ToString().Substring(0, 6));

                //總計
                int SumAmtATax = SumBTaxAmt + SumTax;


                //本幣
                int LocalTotal = SumBTaxAmt;
                int LocalTax = SumTax;

                //訂單
                int Flag = 2;

                //誰來處理
                string Maker = "SandyLo";
                string MakerID = "SI25";


                OrderNo = BillNO;
                gOrderNo = OrderNo;


                //結案註記 //已結
                // int BillStatus = 1;
                int BillStatus = 0;

                //客戶訂單編號
                string CustBillNo = Convert.ToString(row.Cells["ID"].Value);


                string Remark = string.Empty;

                //付款方式
                string GatherStyle = string.Empty;
                string GatherOther = string.Empty;

                string TransMark = string.Empty;

                TransMark = Convert.ToString(row.Cells["TransMark"].Value);

                if (TransMark == "貨到付款")
                {
                    GatherStyle = "0";
                }
                //20140904 月結30day -> 月結30days 
                else if (TransMark == "月結30days")
                {
                    GatherStyle = "2";
                }
               // else if (TransMark == "現金" || TransMark == "電匯" || TransMark == "員工付現")
                else if (TransMark == "現金" || TransMark == "員工付現")
                {
                    GatherStyle = "3";
                    GatherOther = "現金";
                }

                //20140912
                else if (TransMark == "信用卡付款")
                {
                        GatherStyle = "3";
                        GatherOther = "信用卡";

                }
    
        
                else 
                {
                    GatherStyle = "3";
                    GatherOther = TransMark;
                }

                //TransMark
                //貨到付款
                //FOC
                //月結30days
                //現金
                //電匯
                //員工付現
                //GatherStyle=0 貨到 1次月 2月結 3其他
                //CASE A.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END as 付款







                //UserDef1 = Convert.ToString(row.Cells["DelDate"].Value);
                //UserDef2 = Convert.ToString(row.Cells["Serv"].Value); 
                //20140912
                UserDef1 = DateToStr(StrToDate(DelDate).AddDays(-1));

                //帳款歸屬

                string DueTo = CustomerID;
                string ENGCLASS = "";
                System.Data.DataTable CUST1 = GetCustomerCLASS(DueTo);
                if (CUST1.Rows.Count > 0)
                {

                    ENGCLASS = CUST1.Rows[0][0].ToString();
                }
                if (OrdEmail == "")
                {

                    if (ENGCLASS == "C" && !String.IsNullOrEmpty(UNIT))
                    {
                        OrdEmail = "acmegb-fin@acmegb.com";
                    }

                    if (ENGCLASS == "B" || ENGCLASS == "C_Web" || ENGCLASS == "有機通路")
                    {
                        OrdEmail = "acmegb-fin@acmegb.com";
                    }
                }

                if (String.IsNullOrEmpty(OrderPin))
                {
                    OrderPin = CUSTNO;
                }

                string INVADDR = "";
                string INVSEND = "";
                string INVNUIT = "";
                if (InvNo == "紙本發票")
                {
                    INVADDR = DelAddr;
                    INVSEND = "紙本寄送";
                }
                if (InvNo == "公司發票")
                {
                    INVADDR = DelAddr;
                    INVNUIT = "#" + UNIT + "#" + OrdCom2;
                }
                //公司發票
                //#22468373# 聿豐實業股份有限公司
                Remark += "1.備註:" + RivaNote + "\r\n";
                Remark += "2.快遞單號:" + "\r\n";
                Remark += "3.外部訂單單號:" + OrderPin + "\r\n";
                Remark += "4.外部訂單總金額:" + Amount.ToString() + "\r\n";
                Remark += "5.付款人:" + OrdName + "\r\n";
                Remark += "6.紙本:" + INVSEND+ "\r\n";
                Remark += "7.統編:" + INVNUIT + "\r\n";
         
                Remark += "8.發票地址:" + INVADDR + "\r\n";
                if (ADD == "1")
                {
                    Remark += "9.訂購人Email:" + OrdEmail + "\r\n";
                    Remark += "10.是否附DM:是";
                }
                else
                {
                    Remark += "9.訂購人Email:" + OrdEmail;
                }
                //Remark += string.Format("5.PO:{0}", ID) + "\r\n";
                //Remark += "6.付款人:"  + "\r\n";
                //if (!string.IsNullOrEmpty(OrderPin))
                //{
                //Remark += "7.網訂號碼:" + OrderPin;
                //}


                if(PROJECT =="FISH2016")
                {
                    TaxType = 0;
                }
                    

                AddOrdBillMain(BillDate, CustomerID, AddressID, ZipCode, CustAddress, SalesMan, CurrID, ExchRate, SumBTaxAmt, TaxType, SumTax, SumQty, AccMonth, SumAmtATax, LocalTotal, LocalTax, Flag, BillNO, Maker, MakerID, DepartID,
                    LinkMan, LinkTelephone, CustBillNo, BillStatus,
                     UserDef1, UserDef2, Remark, GatherStyle, GatherOther, DueTo, "",PROJECT);
                




                //讀取明細檔


                DataTable dtD = GetPotatoDetail(ID);

                //MessageBox.Show(BillNo);

                int RowNO = 0;
                int SerNO = 0;

                string PreInDate = string.Empty;
                string ProdID = string.Empty;
                string ProdName = string.Empty;
                string ItemRemark = string.Empty;
                
                int Quantity = 0;

                //int Price = 0;
                Double Price = 0;
                Double QuantityF = 0;
                Double PriceA = 0;
                Double dAmount = 0;

                double TaxRate = 0;
                int TaxAmt = 0;

                int Discount = 0;
               // int IsGift = 0;
                Double sPrice = 0;
                int sQuantity = 0;
                Double dAmountT = 0;
                double TAX = 0;
                for (int dj = 0; dj <= dtD.Rows.Count - 1; dj++)
                {
                     //IsGift = 0;
                    //單筆範例
                    ProdID = Convert.ToString(dtD.Rows[dj]["ItemCode"]);
                  
                    ProdName = Convert.ToString(dtD.Rows[dj]["ItemName"]);


                    //個別的數量及金額
                    Quantity = Convert.ToInt16(dtD.Rows[dj]["Qty"]);
                    QuantityF = Convert.ToDouble(dtD.Rows[dj]["Qty"]);
                    Price = Convert.ToDouble(dtD.Rows[dj]["Price"]);
                    ItemRemark = Convert.ToString(dtD.Rows[dj]["ItemRemark"]);
                    string GS = ProdID.Substring(0, 1);
                    if (GS == "P")
                    {
                        Price =  GetPRICE(Price);
    
                    }

                    TaxRate = 0;
                    if (checkBox1.Checked)
                    {
                        if (GS == "P")
                        {
                            TaxRate = 0.05;
                        }
                    }

                    TaxAmt = Convert.ToInt32((QuantityF * Price) * TaxRate);
                    dAmount = (QuantityF * Price);
                    dAmountT += dAmount;
                    
                

                  

                    Discount = 1;

                    int T1 = ProdID.ToUpper().IndexOf("-G");
                    if (T1 != -1)
                    {
                        ProdID = ProdID.ToUpper().Replace("-G", "");
                        Price = 0;
                        System.Data.DataTable G1 = GetOITM(ProdID);

                        if (G1.Rows.Count > 0)
                        {
                            ProdName = G1.Rows[0][0].ToString();
                            PriceA = Convert.ToDouble(G1.Rows[0][2]);
                        }

                    }

                    int T2 = ProdID.ToUpper().IndexOf("-S");
                    if (T2 != -1)
                    {
                        ProdID = ProdID.ToUpper().Replace("-S", "");

                        ItemRemark = "短效品 " + ItemRemark;

                        System.Data.DataTable G1 = GetOITM(ProdID);

                        if (G1.Rows.Count > 0)
                        {
                            ProdName = G1.Rows[0][0].ToString();
                        }
                    }

                    int T3 = ProdID.ToUpper().IndexOf("-P");
                    if (T3 != -1)
                    {
                        ProdID = ProdID.ToUpper().Replace("-P", "");

                        System.Data.DataTable G1 = GetOITM(ProdID);

                        if (G1.Rows.Count > 0)
                        {
                            ProdName = G1.Rows[0][0].ToString();
                        }
                    }

                    //int Flag="";
                    //string BillNO="";

                    string Price2 = (Price * Quantity).ToString();
                    double FF = Convert.ToDouble(Price);
                    sPrice = Price;
                    sQuantity = Quantity;

                    PreInDate = DelDate;
            //        double DdAmount = 0;
              //      double DdAmount2 = 0;
                    //20140829 - 手動輸入模式

                    //取貨日期
                    //UserDef1 = Convert.ToString(dtFriends.Rows[0]["SDate"]);

                    if (string.IsNullOrEmpty(PreInDate))
                    {
                        PreInDate = UserDef1;
                    }

                    //if (GS == "G")
                    //{
                    //    //LLEYTON
                    
                    //    System.Data.DataTable L2 = GetBOMD2(ProdID);
                    //    double TAMT = Convert.ToDouble(L2.Rows[0][0]);
                    //    System.Data.DataTable L1 = GetBOMD(ProdID, TAMT, FF);

                    //    if (L2.Rows.Count > 0 && L1.Rows.Count > 0)
                    //    {
    
                    //        for (int ds = 0; ds <= L1.Rows.Count - 1; ds++)
                    //        {

               
                    //            string QTY = L1.Rows[ds]["QTY"].ToString();
                    //            double CUTP = Convert.ToDouble(L1.Rows[ds]["QTY2"]);

                    //            string CODE = L1.Rows[ds]["CODE"].ToString();
                    //            string CODENAME = "";
                    //            System.Data.DataTable H1 = GetOITM(CODE);
                    //            if (H1.Rows.Count > 0)
                    //            {
                    //                CODENAME = H1.Rows[0][0].ToString();
                    //                PriceA = Convert.ToDouble(H1.Rows[0][2]);
                    //            }

                 
                    //            if (String.IsNullOrEmpty(QTY))
                    //            {
                    //                QTY = "0";
                    //            }
                    //            int Q4 = Convert.ToInt16(QTY) * Quantity;

                    //            string G1 = CODE.Substring(0, 1);
                    //            Price = PriceA * CUTP;
                    //            if (G1 == "P")
                    //            {
                    //                Price = GetPRICE(Price);
                    //            }
                    //            if (checkBox1.Checked)
                    //            {
                    //                if (GS == "P")
                    //                {
                    //                    TaxRate = 0.05;
                    //                }
                    //            }

                    //            TaxAmt = Convert.ToInt32((Q4 * Price) * TaxRate);
                    //            dAmount = (Price * Q4);
                       

                    //            RowNO = RowNO + 1;
                    //            SerNO = SerNO + 1;


                    //            AddOrdBillSub(BillDate, SerNO, CODE, CODENAME, Q4, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate, ProdID, Price2, "");
                                
                    //        }
                    //    }
                    //}
                    //else
                    //{
                        RowNO = RowNO + 1;
                        SerNO = SerNO + 1;

                        AddOrdBillSubF(BillDate, SerNO, ProdID, ProdName, QuantityF, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate, ItemRemark, "", "");
                        
                  //  }
                }
                System.Data.DataTable TT1 = GetBOMTOTAL(BillNO);
                if (TT1.Rows.Count > 0)
                {
                    for (int dj = 0; dj <= TT1.Rows.Count - 1; dj++)
                    {

                        string PRODID = TT1.Rows[dj]["PRODID"].ToString();
                        System.Data.DataTable G1 = GetOITM(PRODID);
                        if (G1.Rows.Count > 0)
                        {
                            string PRODNAME = G1.Rows[0][1].ToString();

                            System.Data.DataTable TT2 = GetBOMTOTAL2(BillNO, PRODID);
                            if (TT2.Rows.Count > 0)
                            {
                                int D1 = Convert.ToInt16(TT2.Rows[0][0]);
                                if (D1 != 0)
                                {
                                    System.Data.DataTable PP1 = GetBOMTOTAL3(BillNO, PRODID);
                                    if (PP1.Rows.Count > 0)
                                    {
                                        int ROWNO = Convert.ToInt16(PP1.Rows[0][0]);
                                        int QTY = Convert.ToInt16(PP1.Rows[0][1]);
                                        int AMT = Convert.ToInt16(PP1.Rows[0][2]) - D1;
                                        Double tmpPrice = Convert.ToDouble(AMT) / QTY;

                                        UPOrdBillSub(BillNO, ROWNO, tmpPrice, AMT);
                                        UPOrdBillSub2(BillNO, PRODID, PRODNAME);
                                    }

                                }

                            }
                        }

                    }

                }
                if (ShipFee > 0)
                {
                    RowNO = RowNO + 1;
                    SerNO = SerNO + 1;

                    ProdID = "FREIGHT01";
                    ProdName = "Freight 運費";
                    Quantity = 1;
                    // Price = 150;
                    //Price = 142.8571;
                    Price = 171;

                    dAmount = Convert.ToInt32(Quantity * Price);

                    TaxRate = 0;
             

                    // TaxAmt = Convert.ToInt32(Amount * TaxRate);
                    TaxAmt = 0;
                    if (checkBox1.Checked)
                    {
                        TaxRate = 0.05;
                        TaxAmt = 9;
                    }
                    Discount = 1;


                    sPrice = Price;
                    sQuantity = Quantity;
                    dAmountT += dAmount;
                    AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate, "", "", "");
                    
                }

                System.Data.DataTable P1 = GETFOC(Amount, CUSTTYPE);
                if (P1.Rows.Count > 0)
                {
                    for (int S = 0; S <= P1.Rows.Count - 1; S++)
                    {
                        RowNO = RowNO + 1;
                        SerNO = SerNO + 1;
                        string ITEMCODE = P1.Rows[S]["ITEMCODE"].ToString();
                        string QTY = P1.Rows[S]["QTY"].ToString();
                        ProdID = ITEMCODE;
                        string CODENAME = "";
                        System.Data.DataTable H1 = GetOITM(ITEMCODE);
                        if (H1.Rows.Count > 0)
                        {
                            CODENAME = H1.Rows[0][0].ToString();
                        }
                        ProdName = CODENAME;
                        Quantity = Convert.ToInt16(QTY);

                        Price = 0;

                        dAmount = 0;

                        TaxRate = 0;
                        TaxAmt = 0;

                        Discount = 1;


                        AddOrdBillSub(BillDate, SerNO, ProdID, ProdName, Quantity, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate, "", "", "");
                    }
                }
                AddOrdBillSubU(BillNO, Convert.ToInt32(dAmountT));

                if (checkBox1.Checked)
                {
                    System.Data.DataTable TA1 = GetBillSUBTAX();

         
                        for (int F = 0; F <= TA1.Rows.Count - 1; F++)
                        {
                            int TTaxType = 0;

                            string RATE = "";
                            if (F == 0)
                            {
                                RATE = "0";
                                TTaxType = 0;
                            }
                            if (F == 1)
                            {
                                RATE = "0.05";
                                TTaxType = 1;
                            }
                            string TBillNO = GetOrderKey(BillDate.ToString());
                            AddOrdBillMain(BillDate, CustomerID, AddressID, ZipCode, CustAddress, SalesMan, CurrID, ExchRate, SumBTaxAmt, TTaxType, SumTax, SumQty, AccMonth, SumAmtATax, LocalTotal, LocalTax, Flag, TBillNO, Maker, MakerID, DepartID,
            LinkMan, LinkTelephone, CustBillNo, BillStatus,
             UserDef1, UserDef2, Remark, GatherStyle, GatherOther, DueTo, "1","");
                            System.Data.DataTable TA2 = GetBillSUB(RATE);
                            if (TA2.Rows.Count > 0)
                            {
                                for (int F2 = 0; F2 <= TA2.Rows.Count - 1; F2++)
                                {
                                    DataRow dd = TA2.Rows[F2];
                                    int TBillDate = Convert.ToInt32(dd["BillDate"]);
                                    int TSerNO = F2 + 1;
                                    string TProdID = dd["ProdID"].ToString();
                                    string TProdName = dd["ProdName"].ToString();
                                    int TQuantity = Convert.ToInt16(dd["Quantity"]);
                                    double TPrice = Convert.ToDouble(dd["Price"]);
                                    double TdAmount = Convert.ToDouble(dd["Amount"]);
                                    double TTaxRate = Convert.ToDouble(dd["TaxRate"]);
                                    int TTaxAmt = Convert.ToInt32(dd["TaxAmt"]);
                                    int TDiscount = Convert.ToInt16(dd["Discount"]);
                                    int TFlag = Convert.ToInt16(dd["Flag"]);

                                    int TRowNO = F2 + 1;
                                    string TPreInDate = dd["PreInDate"].ToString();

                                    AddOrdBillSub(TBillDate, TSerNO, TProdID, TProdName, TQuantity, TPrice, TdAmount, TTaxRate, TTaxAmt, TDiscount, TFlag, TBillNO, TRowNO, TPreInDate, "", "", "1");
                                }
                            }
                        }
                    
                }
                //正式區
                if (rbDb02.Checked)
                {
                  //回寫 成立

                    string memo = string.Format("單號-{0} {1}", BillNO, DateTime.Now.ToString("yyyyMMddHHmmss"));
     
                        UpdatePotato(ID, memo);
                    

                }

            } //for


            MessageBox.Show("己轉入筆數:"+dataGridView1.SelectedRows.Count.ToString());

            //重新整理
            DataTable dt = GetPotato();
            dataGridView1.DataSource = dt;

            
        }
        private double GetPRICE(double OPRICE)
        {

            double PRICE = Math.Round((OPRICE / 1.05), 4, MidpointRounding.AwayFromZero);

            return PRICE;
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

            SqlConnection connection = globals.Connection;

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
        private DataTable GetCustomerByName(string FullName,string CUSTTYPE)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ID FROM comCustomer ");
            sb.Append("Where FullName like '%" + FullName + "%'   ");

            sb.Append("and  flag=1 ");
            if (CUSTTYPE == "員購")
            {
                sb.Append("AND SUBSTRING(ID,1,1) <>0 AND ID <> '90143-170'  ");
            }
            sb.Append(" order by id desc");

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
        private DataTable GetCustomerByName2(string FullName, string CUSTTYPE)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ID FROM comCustomer ");
            sb.Append("Where FullName like '%" + FullName + "%'   ");

            sb.Append("and  flag=1  ");
            if (CUSTTYPE == "員購")
            {
                sb.Append("AND SUBSTRING(ID,1,1) <>0 AND ID <> '90143-170'  ");
            }
            sb.Append("  order by CAST(ltrim(substring(ID,CHARINDEX('-', ID)+1,10)) AS INT)      ");

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
        private DataTable GetCustomerCLASS(string ID)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("      select T1.EngName from comCustomer T0 LEFT JOIN  comCustClass T1 ON (T0.ClassID =T1.ClassID AND T1.Flag =1) WHERE T0.ID =@ID AND T0.Flag =1");



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
        private DataTable GetFriends(string ID)
        {

            SqlConnection connection = globals.Connection;

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

            sb.Append("              SELECT IsNull(Max(ID),'00000') FROM comCustomer  ");
            sb.Append("              where Flag=1 and SUBSTRING(id,1,6)='90143-' and LEN(id)=9 ");


            

 


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

            string s = "90143-";
            string n = ID.Substring(6, ID.Length - 6);

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
        public void AddcomCustomer(string ID, string FullName, string ClassId, string PersonID, string email, string Telephone1)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert into comCustomer(Flag,ID ,FundsAttribution,TransNewID,CurrencyID ,FullName,InvoiceHead,ShortName,ClassId,PersonID,email,Telephone1) " +
            "values (@Flag,@ID,@FundsAttribution,@TransNewID,@CurrencyID,@FullName,@InvoiceHead,@ShortName,@ClassId,@PersonID,@email,@Telephone1)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@FundsAttribution", ID));
            command.Parameters.Add(new SqlParameter("@TransNewID", ID));
            command.Parameters.Add(new SqlParameter("@CurrencyID", "NTD"));
            command.Parameters.Add(new SqlParameter("@FullName", FullName));
            command.Parameters.Add(new SqlParameter("@InvoiceHead", FullName));
            command.Parameters.Add(new SqlParameter("@email", email));
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
            command.Parameters.Add(new SqlParameter("@PersonID", PersonID));
            command.Parameters.Add(new SqlParameter("@Telephone1", Telephone1));
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
        public void AddcomCustomerD(string ID, string FullName, string ClassId, string PersonID, string Moderm)
        {
            SqlConnection connection = new SqlConnection(strCn2);
            string sql = "Insert into comCustomer(Flag,ID ,FundsAttribution,TransNewID,CurrencyID ,FullName,InvoiceHead,ShortName,ClassId,PersonID,Moderm) " +
            "values (@Flag,@ID,@FundsAttribution,@TransNewID,@CurrencyID,@FullName,@InvoiceHead,@ShortName,@ClassId,@PersonID,@Moderm)";
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
            command.Parameters.Add(new SqlParameter("@PersonID", PersonID));
            command.Parameters.Add(new SqlParameter("@Moderm", Moderm));
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

        public void AddcomCustomerD2(string ID, string FullName, string ClassId, string PersonID,string STRN)
        {
            SqlConnection connection = new SqlConnection(STRN);
            string sql = "Insert into comCustomer(Flag,ID ,FundsAttribution,TransNewID,CurrencyID ,FullName,InvoiceHead,ShortName,ClassId,PersonID,Moderm) " +
            "values (@Flag,@ID,@FundsAttribution,@TransNewID,@CurrencyID,@FullName,@InvoiceHead,@ShortName,@ClassId,@PersonID,@Moderm)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@FundsAttribution", ID));
            command.Parameters.Add(new SqlParameter("@TransNewID", ID));
            command.Parameters.Add(new SqlParameter("@CurrencyID", "NTD"));
            command.Parameters.Add(new SqlParameter("@FullName", FullName));
            command.Parameters.Add(new SqlParameter("@InvoiceHead", ID));
            command.Parameters.Add(new SqlParameter("@ShortName", FullName));
   

            //ClassId
            command.Parameters.Add(new SqlParameter("@ClassId", ClassId));
            command.Parameters.Add(new SqlParameter("@PersonID", PersonID));
            command.Parameters.Add(new SqlParameter("@Moderm", ID));
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
        public void AddcomCustDesc(string ID, string AddField1)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);
            string sql = "Insert into comCustDesc(Flag,ID,AddField1 ) " +
            "values (@Flag,@ID,@AddField1)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@AddField1", AddField1));
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

        public void AddcomCustDescD(string ID, string AddField1)
        {
            SqlConnection connection = new SqlConnection(strCn2);
            string sql = "Insert into comCustDesc(Flag,ID,AddField1) " +
            "values (@Flag,@ID,@AddField1)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@AddField1", AddField1));
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
        public void AddcomCustDescD2(string ID, string AddField1, string RateOfDiscount, string STRN)
        {
            SqlConnection connection = new SqlConnection(STRN);
            string sql = "Insert into comCustDesc(Flag,ID,AddField1,RateOfDiscount) " +
            "values (@Flag,@ID,@AddField1,@RateOfDiscount)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@AddField1", AddField1));
            command.Parameters.Add(new SqlParameter("@RateOfDiscount", RateOfDiscount));
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

        public void AddOrdBillMain(int BillDate, string CustomerID, string AddressID, string ZipCode, string CustAddress, string SalesMan, string CurrID, int ExchRate, int SumBTaxAmt, int TaxType, int SumTax, int SumQty, int AccMonth, int SumAmtATax, int LocalTotal, int LocalTax, int Flag, string BillNO, string Maker, string MakerID,
            string DepartID, string LinkMan, string LinkTelephone, string CustBillNo, int BillStatus,
            string UserDef1, string UserDef2, string Remark, string GatherStyle, string GatherOther, string DueTo, string con, string ProjectID)
        {

            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }

            if (con == "1")
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "Insert Into OrdBillMain (BillDate,CustomerID,AddressID,ZipCode,CustAddress,SalesMan,CurrID,ExchRate,SumBTaxAmt,TaxType,SumTax,SumQty,AccMonth,SumAmtATax,LocalTotal,LocalTax,Flag,BillNO,Maker,MakerID,DepartID,LinkMan,LinkTelephone,BillStatus,CustBillNo,UserDef1, UserDef2 ,Remark,GatherStyle,GatherOther,DueTo,FormalCust,ProjectID) " +
            "values (@BillDate,@CustomerID,@AddressID,@ZipCode,@CustAddress,@SalesMan,@CurrID,@ExchRate,@SumBTaxAmt,@TaxType,@SumTax,@SumQty,@AccMonth,@SumAmtATax,@LocalTotal,@LocalTax,@Flag,@BillNO,@Maker,@MakerID,@DepartID,@LinkMan,@LinkTelephone,@BillStatus,@CustBillNo,@UserDef1, @UserDef2 ,@Remark,@GatherStyle,@GatherOther,@DueTo,@FormalCust,@ProjectID)";
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
            command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));


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
            int Flag, string BillNO, int RowNO, string PreInDate, string ItemRemark, string Detail,string con)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            if (con == "1")
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "Insert Into OrdBillSub (BillDate,SerNO,ProdID,ProdName,Quantity,Price,Amount,TaxRate,TaxAmt,Discount,Flag,BillNO,RowNO,sQuantity,sPrice,QtyRemain,PreInDate,ItemRemark,Detail) " +
            "values (@BillDate,@SerNO,@ProdID,@ProdName,@Quantity,@Price,@Amount,@TaxRate,@TaxAmt,@Discount,@Flag,@BillNO,@RowNO,@sQuantity,@sPrice,@QtyRemain,@PreInDate,@ItemRemark,@Detail)";
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
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@Detail", Detail));
            




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


        public void AddOrdBillSubD(int BillDate, int SerNO, string ProdID, string ProdName, int Quantity, Double Price, Double Amount, double TaxRate, int TaxAmt, int Discount,
            int Flag, string BillNO, int RowNO, string PreInDate, string ItemRemark, string Detail, string con)
        {
            SqlConnection connection = new SqlConnection(strCn3);
            string sql = "Insert Into OrdBillSub (BillDate,SerNO,ProdID,ProdName,Quantity,Price,Amount,TaxRate,TaxAmt,Discount,Flag,BillNO,RowNO,sQuantity,sPrice,QtyRemain,PreInDate,ItemRemark,Detail) " +
            "values (@BillDate,@SerNO,@ProdID,@ProdName,@Quantity,@Price,@Amount,@TaxRate,@TaxAmt,@Discount,@Flag,@BillNO,@RowNO,@sQuantity,@sPrice,@QtyRemain,@PreInDate,@ItemRemark,@Detail)";
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
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@Detail", Detail));





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
        public void AddOrdBillSubF(int BillDate, int SerNO, string ProdID, string ProdName, Double Quantity, Double Price, Double Amount, double TaxRate, int TaxAmt, int Discount,
    int Flag, string BillNO, int RowNO, string PreInDate, string ItemRemark, string Detail, string con)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            if (con == "1")
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "Insert Into OrdBillSub (BillDate,SerNO,ProdID,ProdName,Quantity,Price,Amount,TaxRate,TaxAmt,Discount,Flag,BillNO,RowNO,sQuantity,sPrice,QtyRemain,PreInDate,ItemRemark,Detail) " +
            "values (@BillDate,@SerNO,@ProdID,@ProdName,@Quantity,@Price,@Amount,@TaxRate,@TaxAmt,@Discount,@Flag,@BillNO,@RowNO,@sQuantity,@sPrice,@QtyRemain,@PreInDate,@ItemRemark,@Detail)";
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
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@Detail", Detail));





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
        public void AddOrdBillSubU(string BillNO, int SumAmtATax)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "UPDATE ordBillMain SET SumAmtATax=@SumAmtATax,LocalTotal=@SumAmtATax,SumBTaxAmt=@SumAmtATax  where BillNO =@BillNO ";
           
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@SumAmtATax", SumAmtATax));




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
        public void AddOrdBillSubU2(string BillNO, int SumAmtATax)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            //select SumAmtATax,LocalTotal,SumBTaxAmt,SumTax,LocalTax  from ordBillMain where BillNO ='J105010479'
            string sql = "UPDATE ordBillMain SET SumAmtATax=@SumAmtATax,LocalTotal=@SumAmtATax,SumBTaxAmt=@SumAmtATax  where BillNO =@BillNO ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@SumAmtATax", SumAmtATax));




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
        public void trunOrdBill()
        {
            SqlConnection connection = globals.Connection;

            string sql = "truncate table  ordBillMain truncate table  ordBillsub ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;






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
        public void UPOrdBillSub(string BillNO, int ROWNO, double PRICE, int Amount)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "UPDATE OrdBillSub SET PRICE=@PRICE,SPRICE=@PRICE,Amount=@Amount where BillNO =@BillNO AND ROWNO =@ROWNO ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@ROWNO", ROWNO));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));



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
        public void UPOrdBillSub2(string BillNO, string PRODID, string ItemRemark)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }
            string sql = "UPDATE OrdBillSub SET ItemRemark=@ItemRemark,detail='' where BillNO =@BillNO AND ItemRemark =@PRODID ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));




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
            string sql = "Insert into comCustTrade(Flag,ID,InvoiceType,TaxKind,AccReceivable,InvoiceStyle,AccBillRecv,AccAdvRecv ) " +
            "values (@Flag,@ID,@InvoiceType,@TaxKind,@AccReceivable,@InvoiceStyle,@AccBillRecv,@AccAdvRecv)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            //20140902
            command.Parameters.Add(new SqlParameter("@InvoiceType", "35"));
            command.Parameters.Add(new SqlParameter("@TaxKind", "1"));
            command.Parameters.Add(new SqlParameter("@AccReceivable", "1144992"));
            command.Parameters.Add(new SqlParameter("@InvoiceStyle", "1"));
            command.Parameters.Add(new SqlParameter("@AccBillRecv", "1141000"));
            command.Parameters.Add(new SqlParameter("@AccAdvRecv", "2283000"));

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
        public void AddcomCustTradeD(string ID)
        {
            SqlConnection connection = new SqlConnection(strCn2);
            string sql = "Insert into comCustTrade(Flag,ID,InvoiceType,TaxKind,AccReceivable,InvoiceStyle,AccBillRecv,AccAdvRecv ) " +
            "values (@Flag,@ID,@InvoiceType,@TaxKind,@AccReceivable,@InvoiceStyle,@AccBillRecv,@AccAdvRecv)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            //20140902
            command.Parameters.Add(new SqlParameter("@InvoiceType", "35"));
            command.Parameters.Add(new SqlParameter("@TaxKind", "1"));
            command.Parameters.Add(new SqlParameter("@AccReceivable", "1144992"));
            command.Parameters.Add(new SqlParameter("@InvoiceStyle", "1"));
            command.Parameters.Add(new SqlParameter("@AccBillRecv", "1141000"));
            command.Parameters.Add(new SqlParameter("@AccAdvRecv", "2283000"));

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
        public void AddcomCustTradeD2(string ID,string STRN)
        {
            SqlConnection connection = new SqlConnection(STRN);
            string sql = "Insert into comCustTrade(Flag,ID,InvoiceType,TaxKind,AccReceivable,InvoiceStyle,AccBillRecv,AccAdvRecv ) " +
            "values (@Flag,@ID,@InvoiceType,@TaxKind,@AccReceivable,@InvoiceStyle,@AccBillRecv,@AccAdvRecv)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@Flag", "1"));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            //20140902
            command.Parameters.Add(new SqlParameter("@InvoiceType", "35"));
            command.Parameters.Add(new SqlParameter("@TaxKind", "1"));
            command.Parameters.Add(new SqlParameter("@AccReceivable", "1144992"));
            command.Parameters.Add(new SqlParameter("@InvoiceStyle", "1"));
            command.Parameters.Add(new SqlParameter("@AccBillRecv", "1141000"));
            command.Parameters.Add(new SqlParameter("@AccAdvRecv", "2283000"));

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

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM gb_potato2 where ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
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
        private DataTable GETFOC(int AMT,string CUSTTYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,QTY  FROM GB_FOC T0");
            sb.Append(" LEFT JOIN GB_FOC2 T1 ON (T0.ID=T1.ID)");
            sb.Append(" LEFT JOIN GB_FOC3 T2 ON (T0.ID=T2.ID)");
            sb.Append(" WHERE T0.STATUS='進行中' AND T1.STATUS='TRUE' ");
            sb.Append(" AND @DATE BETWEEN STARTDATE AND ENDDATE");
            sb.Append(" AND @AMT>=AMT AND CUSTTYPE=@CUSTTYPE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE", DateTime.Now.ToString("yyyyMMdd")));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@CUSTTYPE", CUSTTYPE));
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
        private DataTable GetBOMD(string ProdID, double TAMT, double TPRICE2)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT CombSubID CODE,CAST(Amount AS INT) QTY,@TPRICE2/@TQTY QTY2  FROM CHIComp02.DBO.comProdCombine  WHERE ProdID =@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@TQTY", TAMT));
            command.Parameters.Add(new SqlParameter("@TPRICE2", TPRICE2));
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
        private DataTable GetBOMD2(string ProdID)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append(" select SUM(T1.Amount*T0.SalesPriceA) AMT  from CHIComp02.DBO.comProduct T0 LEFT JOIN CHIComp02.DBO.comProdCombine T1 ON (T0.ProdID=T1.CombSubID) where T1.ProdID =@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

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
        private DataTable GetBillMAIN()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM ordBillMAIN ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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
        private DataTable GetBillSUB(string taxrate)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM ordBillSUB where taxrate=@taxrate ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@taxrate", taxrate));

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
        private DataTable GetBillSUBTAX()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT TAXRATE FROM ordBillSUB ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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
        private DataTable GetBOMTOTAL(string BillNO)
        {
            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }


            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT DISTINCT ItemRemark PRODID,DETAIL AMT  FROM OrdBillSub WHERE BillNO =@BillNO AND ISNULL(ItemRemark,'') <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

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
        private DataTable GetBOMTOTAL2(string BillNO, string PRODID)
        {

            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }

            StringBuilder sb = new StringBuilder();


            sb.Append("SELECT CAST(SUM(CASE WHEN SUBSTRING(PRODID,1,1)='P' THEN CAST(ROUND(AMOUNT*1.05,0) AS INT) ELSE CAST(Amount AS INT)   END)-MAX(Detail) AS int) AMT  FROM OrdBillSub WHERE BillNO =@BillNO AND ItemRemark =@PRODID AND ISNULL(ItemRemark,'') <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));
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
        private DataTable GetBOMTOTAL3(string BillNO, string PRODID)
        {

            SqlConnection connection = null;
            if (checkBox1.Checked)
            {
                connection = globals.Connection;
            }
            else
            {
                connection = new SqlConnection(ConnectiongString);
            }

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT TOP 1 ROWNO,Quantity,AMOUNT  FROM OrdBillSub WHERE BillNO =@BillNO AND ItemRemark =@PRODID AND ISNULL(ItemRemark,'') <> '' AND Price <> 0 AND SUBSTRING(PRODID,1,1)='M' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));
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
        private DataTable GetOITM(string ProdID)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("       select InvoProdName,ProdName,SalesPriceA From comProduct A  Where A.ProdID = @ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

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
            SqlConnection connection = globals.Connection;
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

        private void rbDb03_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDb02.Checked)
            {
                ConnectiongString = ConnectiongString.Replace("CHIComp92", "CHIComp02");
            }

            if (rbDb03.Checked)
            {
                ConnectiongString = ConnectiongString.Replace("CHIComp02", "CHIComp92");
            }
        }

        private void fmRg2Chi_Load(object sender, EventArgs e)
        {

            if (globals.GroupID.ToString().Trim() == "EEP")
            {
                rbDb03.Checked = true;
            }
            DataTable dt = GetPotato();
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    FileName = openFileDialog1.FileName;


            //    WriteExcelGBPICK3(FileName);

            //}
            for (int w = 2001; w <= 3000; w++)
            {
                string CustomerID = w.ToString("000000");

                AddcomCustomerD2(CustomerID, "東門臨時會員", "034", "SI30", strCn2);
                AddcomCustDescD2(CustomerID, "", "1", strCn2);
                AddcomCustTradeD2(CustomerID, strCn2);

                AddcomCustomerD2(CustomerID, "東門臨時會員", "034", "SI30", strCn3);
                AddcomCustDescD2(CustomerID, "", "1", strCn3);
                AddcomCustTradeD2(CustomerID, strCn3);
            }


           // AddOrdBillSubD("20190116", SerNO, CODE, CODENAME, Q4, Price, dAmount, TaxRate, TaxAmt, Discount, Flag, BillNO, RowNO, PreInDate, ProdID, Price2, "");
        }
        private void WriteExcelGBPICK3(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;

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



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string SERNO;
                string CARDCODE;
                string CARDNAME;
                string QTY;
                string PRICE;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERNO = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CARDCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    QTY = range.Text.ToString().Trim();

                    int QQ = Convert.ToInt32(Convert.ToDouble(QTY));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    PRICE = range.Text.ToString().Trim();
                    Int32 G1 = 20190116;
                    //if (ID != "")
                    //{
                    double TAX = 0;
                    double AMOUNT = Convert.ToDouble(QTY) * Convert.ToDouble(PRICE);
                    AddOrdBillSubD(G1, Convert.ToInt16(SERNO), CARDCODE, CARDNAME, QQ, Convert.ToDouble(PRICE), AMOUNT, TAX, 0, 1, 4, "J108011603", Convert.ToInt16(SERNO), "20190116", "", "", "");
                  //  }
                }




            }
            finally
            {



                //try
                //{
                //    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //}
                //catch
                //{
                //}
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


               // System.Diagnostics.Process.Start(NewFileName);


            }



        }

    } //p
} //n 
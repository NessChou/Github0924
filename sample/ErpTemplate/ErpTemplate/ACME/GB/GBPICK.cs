using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ACME
{
    public partial class GBPICK : ACME.fmBase1
    {
        private string FileName;
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn03 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBPICK()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            gB_PICKTableAdapter.Connection = MyConnection;
            gB_PICK2TableAdapter.Connection = MyConnection;
        }
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            uPDATEUSERTextBox.ReadOnly = true;
            cREATEUSERTextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            cHECKEDCheckBox.Enabled = false;
            cHECKEDDATETextBox.ReadOnly = true;
            button7.Enabled = true;
        }
   
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            uPDATEUSERTextBox.ReadOnly = false;
            cREATEUSERTextBox.ReadOnly = false;
            dOCDATETextBox.ReadOnly = false;
            comboBox3.SelectedIndex = -1;
        }
        public override void AfterDelete()
        {
                   DialogResult result;
            result = MessageBox.Show("您確認是否要刪除", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                DELPICK(shippingCodeTextBox.Text);
            

            }
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                pOTATO.GB_PICK.RejectChanges();
                pOTATO.GB_PICK2.RejectChanges();
            }
            catch
            {
            }
            return true;

        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override void EndEdit()
        {
            WW();

        }

        public override void AfterEdit()
        {
            shippingCodeTextBox.ReadOnly = true;

            uPDATEUSERTextBox.Text = fmLogin.LoginID.ToString();

            for (int i = 0; i <= 10; i++)
            {
                if (i != 4)
                {
                    gB_PICK2DataGridView.Columns[i].ReadOnly = true;
                }
            }
        }
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = gB_PICKBindingSource;
            MyTableName = "GB_PICK";
            MyIDFieldName = "ShippingCode";

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "GP" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            string username = fmLogin.LoginID.ToString();
            dOCDATETextBox.Text = GetMenu.Day();
            cREATEUSERTextBox.Text = username;
            this.gB_PICKBindingSource.EndEdit();
            kyes = null;

            gBTYPETextBox.Text = "零售";
        }
        public override void FillData()
        {
            try
            {

                gB_PICKTableAdapter.Fill(pOTATO.GB_PICK, MyID);
                gB_PICK2TableAdapter.Fill(pOTATO.GB_PICK2, MyID);
                comboBox1.Text = "取貨日期";

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

                gB_PICK2BindingSource.MoveFirst();

                for (int i = 0; i <= gB_PICK2BindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)gB_PICK2BindingSource.Current;

                    row3["LINE"] = i;

                    gB_PICK2BindingSource.EndEdit();

                    gB_PICK2BindingSource.MoveNext();

                }




                gB_PICKTableAdapter.Connection.Open();

                gB_PICKBindingSource.EndEdit();
                gB_PICK2BindingSource.EndEdit();


                tx = gB_PICKTableAdapter.Connection.BeginTransaction();

                SqlDataAdapter Adapter = util.GetAdapter(gB_PICKTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(gB_PICK2TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                gB_PICKTableAdapter.Update(pOTATO.GB_PICK);
                pOTATO.GB_PICK.AcceptChanges();

                gB_PICK2TableAdapter.Update(pOTATO.GB_PICK2);
                pOTATO.GB_PICK2.AcceptChanges();


                tx.Commit();

                this.MyID = this.shippingCodeTextBox.Text;

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
                this.gB_PICKTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("請輸入資料");
                return;
            }

            System.Data.DataTable dt1 = null;
            if (comboBox2.Text == "門市")
            {
                dt1 = DD2(textBox1.Text, textBox2.Text);
            }
            else
            {
                dt1 = DD1(textBox1.Text, textBox2.Text, "0");
            }
            if (dt1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            string BILLNO = "";
            System.Data.DataTable dt2 = pOTATO.GB_PICK2;
            DataRow drw2 = null;
            string DUP = "";
            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                string GTYPE = gBTYPETextBox.Text;
                drw2["GBGROUP"] = drw["產品別"];
                drw2["GTYPE"] = GTYPE;
                drw2["BILLNO"] = drw["BILLNO"];
                drw2["ROWNO"] = drw["ROWNO"];
                drw2["BILLDATE"] = drw["BILLDATE"];
                drw2["CARDCODE"] = drw["CARDCODE"];
                drw2["CARDNAME"] = drw["CARDNAME"];
                drw2["CHI"] = drw["CLASS"];
                drw2["BILLNO2"] = drw["CUSTBILLNO"];
                
                string CUSTID = drw["CUSTBILLNO"].ToString();
                System.Data.DataTable F11 = GETRIVA(CUSTID);
                if (F11.Rows.Count > 0)
                {
                    drw2["CUSTMEMO"] = F11.Rows[0][0].ToString();

                }
                else
                {
                    string REMARK = drw["REMARK"].ToString();
                    int T1 = REMARK.IndexOf("1.備註:");
                    int T2 = REMARK.IndexOf("2.");

                    if (T1 != -1 && T2 != -1)
                    {
                        string H1 = REMARK.Substring(T1 + 5, T2 - T1 - 5);


                        drw2["CUSTMEMO"] = H1;
                    }
              
                }
                string BILLNOS = drw["BILLNO"].ToString() + " " + drw["ROWNO"];

          
                string GROUP = "";
                string ITEMCODE = drw["ITEMCODE"].ToString().Substring(0, 3);
                if (ITEMCODE == "MCK")
                {
                    GROUP = "雞";
                }
                if (ITEMCODE == "MPK")
                {
                    GROUP = "豬";
                }
                if (ITEMCODE == "MSR")
                {
                    GROUP = "蝦";
                }
                if (ITEMCODE.Substring(0,1) == "P")
                {
                    GROUP = "加工品";
                }

                drw2["ITEMCODE"] = drw["ITEMCODE"].ToString();
                drw2["ITEMNAME"] = GROUP + "_" + drw["ITEMNAME"];
                drw2["QTY2"] = drw["QTY2"];
                drw2["UNIT"] = drw["UNIT"];
                decimal QTY2 = Convert.ToDecimal(drw["QTY2"]);
                drw2["QTY"] = drw["QTY"];
                //1.備註:2.快遞單號:
                drw2["PRICE"] = drw["PRICE"];
                if (BILLNO != drw["BILLNO"].ToString())
                {
                    string REMARK = drw["REMARK"].ToString();
                    int T1 = REMARK.IndexOf("4.外部訂單總金額:");
                    int T2 = REMARK.IndexOf("5.");
                    int L1 = 0;
                    System.Data.DataTable K1 = DD23T(drw["BILLNO"].ToString());
                    if (K1.Rows.Count > 0)
                    {
                        L1 = 9;
                    }
                           string TRADE = drw["交易方式"].ToString();
                 
                  
                    if (TRADE == "貨到付款")
                    {
                        int F1 = Convert.ToInt32(drw["AMT"]);
                        if (F1 == 0)
                        {
                            L1 = 0;
                        }
                        if (T1 != -1 && T2 != -1)
                        {
                            string H1 = REMARK.Substring(T1 + 10, T2 - T1 - 10);
                          
                            H1 = H1.Replace("\r\n", "");
                       
                            drw2["AMT"] = Convert.ToInt32(H1);
                        }
                        else
                        {
                            drw2["AMT"] = Convert.ToInt32(drw["AMT"]) + L1;
                        }
            
                    }
                    else
                    {
                        drw2["AMT"] = 0;
                    }
                }
                else
                {
                    drw2["AMT"] = 0;
                }
                drw2["TRADE"] = drw["交易方式"];
                drw2["CustAddress"] = drw["CustAddress"];
                drw2["LinkMan"] = drw["LinkMan"];
                drw2["LinkTelephone"] = drw["LinkTelephone"];
                drw2["FAXNO"] = drw["FaxNo"];
                drw2["ORDMAN"] = drw["ORDMAN"];
                drw2["BARCODEID"] = drw["BARCODEID"];
                drw2["UserDef1"] = drw["UserDef1"];
                drw2["PreInDate"] = drw["PreInDate"];
                drw2["UserDef2"] = drw["UserDef2"];
                drw2["PACK1"] = 1;
                drw2["AMT2"] = 0;
                drw2["AMT3"] = 0;
                drw2["MEMO"] = drw["ITEMREMARK"];
                drw2["AMT4"] = drw["AMT4"];
                if (DUP == BILLNOS)
                {
                    drw2["AMT4"] = 0;
                }
                drw2["CUSTTYPE"] = drw["客戶類別"];
                string ITEMREMARK=drw["ITEMREMARK"].ToString();
                //System.Data.DataTable DTP = GETPRODNAME(ITEMREMARK);
                //if (DTP.Rows.Count > 0)
                //{
                //    drw2["MEMO"] = DTP.Rows[0][0].ToString();
                //}
                string ENDDATE = "";
                string CD = drw2["MEMO"].ToString();
                drw2["PACKAGE"] = drw["ADDRID"];

                if (drw["PRICE"].ToString() == "0")
                {
                    ENDDATE = "D6";
                }
                else if (drw2["MEMO"].ToString().Trim() == "短效品")
                {
                    ENDDATE = "D7";
                }
                else if (drw2["MEMO"].ToString().Trim() == "促銷品")
                {
                    ENDDATE = "D8";
                }
                else
                {
                    ENDDATE = "D5";
                }
                System.Data.DataTable DD = DDSMANF(drw["ITEMCODE"].ToString(), ENDDATE);
                if (DD.Rows.Count > 0)
                {
                    drw2["DEADDATE"] = DD.Rows[0][0].ToString();
                    string REMARK = drw["REMARK"].ToString();
                    int GG1 = REMARK.IndexOf("捐贈");
                    if (GG1 != -1)
                    {
                        DateTime T1 = Convert.ToDateTime(DD.Rows[0][1]);
                        DateTime T2 = DateTime.Now.AddMonths(3);

                        if (T1 > T2)
                        {
                            drw2["DEADDATE"] = "";
                        
                        }
                    }

                }
                DUP = BILLNOS;
      
                dt2.Rows.Add(drw2);
                if (BILLNO != drw["BILLNO"].ToString())
                {
                    string REMARK = drw["REMARK"].ToString();
                    int GG1 = REMARK.IndexOf("是否附DM:是");
                    if (GG1 != -1)
                    {
                        drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["GTYPE"] = gBTYPETextBox.Text;
                        drw2["BILLNO"] = drw["BILLNO"];
                        drw2["ROWNO"] = drw["ROWNO"];

                        drw2["BILLDATE"] = drw["BILLDATE"];
                        drw2["CARDCODE"] = drw["CARDCODE"];
                        drw2["CARDNAME"] = drw["CARDNAME"];
                        drw2["CHI"] = drw["CLASS"];

                        drw2["ITEMCODE"] = "";
                        drw2["ITEMNAME"] = "DM";
                        drw2["QTY"] = 1;
                        drw2["QTY2"] = 1;
                        drw2["PRICE"] = 0;
                        drw2["AMT"] = 0;
                        drw2["CustAddress"] = drw["CustAddress"];
                        drw2["LinkMan"] = drw["LinkMan"];
                        drw2["LinkTelephone"] = drw["LinkTelephone"];
                        drw2["ORDMAN"] = drw["ORDMAN"];
                        drw2["UserDef1"] = drw["UserDef1"];
                        drw2["PreInDate"] = drw["PreInDate"];
                        drw2["UserDef2"] = drw["UserDef2"];
                        drw2["PACK1"] = 1;
                        drw2["UNIT"] = "";
                        drw2["TRADE"] = drw["交易方式"];
                        drw2["AMT2"] = 0;
                        drw2["AMT3"] = 0;
                        dt2.Rows.Add(drw2);
                    }
                }

                 BILLNO = drw["BILLNO"].ToString();

            }

         


            for (int j = 0; j <= gB_PICK2DataGridView.Rows.Count - 1; j++)
            {
                gB_PICK2DataGridView.Rows[j].Cells[0].Value = j.ToString();
            }


            gB_PICKBindingSource.EndEdit();
            gB_PICK2BindingSource.EndEdit();
        }

        private void GBPICK_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
            comboBox1.Text = "取貨日期";
            button6.Enabled = true;
            comboBox2.Text = "聿豐";
            WW();
        }
        public System.Data.DataTable DD1(string CreateDate, string CreateDate2, string KEY)
        {
         
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                                            Select  G.ProdID,a.billdate,H.CLASSID , A.BillNO BILLNO,A.CUSTBILLNO,G.PRODNAME,  ");
            sb.Append("                                                G.ProdID ITEMCODE,G.RowNO ROWNO,CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END 交易方式,    ");
            sb.Append("                                                                       h.InvoProdName  ITEMNAME,    ");
            sb.Append("                                                                       G.Quantity  QTY,    ");
            sb.Append("                                                                    G.Quantity  QTY2,     ");
            sb.Append("                                                                           A.CustAddress ,    ");
            sb.Append("                                                                           A.LinkMan ,    ");
            sb.Append("                                                                           A.LinkTelephone ,    ");
            sb.Append("                                                                           A.UserDef1 , T11.FaxNo,   ");
            sb.Append("                                                                           Convert(varchar(8),G.PreInDate) PreInDate,H.BARCODEID BARCODEID,    ");
            sb.Append("                                                                                    CASE  SUBSTRING(A.UserDef2,0,CHARINDEX('(', A.UserDef2)) WHEN '' THEN  ");
            sb.Append("                                                                                    CASE A.UserDef2 WHEN '中午前' THEN '中午前'  WHEN '下午' THEN '12-17時' WHEN '晚上' THEN '17-20時'  ");
            sb.Append("                                                                                     ELSE A.UserDef2 END  WHEN '中午前' THEN '中午前' WHEN '下午' THEN '12-17時' WHEN '晚上' THEN '17-20時' End UserDef2 ,G.PRICE,    ");
            sb.Append("                                                                                      convert(int, case when A.GatherStyle=0 then  A.SumAmtATax else 0 end) as AMT,A.CustomerID CARDCODE,B.FullName CARDNAME,A.BillDate BILLDATE,     ");
            sb.Append("                                                      CASE CHARINDEX('-', B.ShortName) WHEN 0 THEN B.ShortName ELSE  SUBSTRING(B.ShortName,0,CHARINDEX('-', B.ShortName)) END+' TEL :'+CASE WHEN ISNULL(Telephone1,'') = '' THEN LinkTelephone ELSE Telephone1 END ORDMAN,     ");
            sb.Append("                                                     CAST(A.AddressID  AS INT)  ADDRID,     ");
            sb.Append("                                                       Case when H.ClassID ='AWC200' then '朝貢雞'     ");
            sb.Append("                                                        when H.ClassID ='ARC200' then '朝貢雞'     ");
            sb.Append("                                                        when H.ClassID ='AWP200' then '朝貢豬'     ");
            sb.Append("                                                        when H.ClassID ='ARP200' then '朝貢豬'     ");
            sb.Append("                                                        when H.ClassID ='AWS200' then '大力蝦'     ");
            sb.Append("                                                        when H.ClassID ='ARS200' then '大力蝦'     ");
            sb.Append("                                                        when H.ClassID ='AWS210' then '白金蝦'     ");
            sb.Append("                                                        when H.ClassID ='ARS210' then '白金蝦'     ");
            sb.Append("                                                        when H.ClassID ='ARG100' then '禮盒'     ");
            sb.Append("                                when SUBSTRING(G.ProdID,1,1)='P' THEN '加工品'      ");
            sb.Append("                                                       else '空白' end 產品別,A.SumAmtATax AMT2,G.ITEMREMARK,A.UserDef2,A.UserDef2,A.Remark REMARK,G.Amount  AMT4 ,L.ClassName 客戶類別,B.ClassID CLASS,H.UNIT    ");
            sb.Append("                                                                           From OrdBillMain A     ");
            sb.Append("                                                                           Inner Join OrdBillSub G On G.Flag=A.Flag And G.BillNO=A.BillNO     ");
            sb.Append("                                                                           Left Join comCustomer B On B.Flag=A.Flag-1 And B.ID=A.CustomerID     ");
            sb.Append("                                                                           Left Join comPerson E On E.PersonID=A.SalesMan     ");
            sb.Append("                                                                           Left Join comProduct H On H.ProdID=G.ProdID     ");
            sb.Append("                                                                                      LEFT JOIN comCustAddress T11 ON (A.AddressID=T11.AddrID AND A.CustomerID=T11.ID )   ");
            sb.Append("               Left Join comCustClass L On L.ClassID =b.ClassID and L.Flag =1 ");
            sb.Append("                    Where A.Flag=2 AND G.ProdID<> 'FREIGHT01' ");
            if (textBox5.Text == "")
            {
                if (comboBox1.Text == "訂購日期")
                {
                    sb.Append("             And  A.BILLDATE Between @CreateDate and @CreateDate2 ");
                }
                if (comboBox1.Text == "取貨日期")
                {
                    sb.Append("             And   A.UserDef1 Between @CreateDate and @CreateDate2 ");
                }

                if (KEY == "1" && KEY == "2")
                {
                    if (gBTYPETextBox.Text == "批發")
                    {
                        sb.Append("    and B.ClassID = ('011')   ");
                    }
                    else if (gBTYPETextBox.Text == "零售")
                    {
                        sb.Append("      and   H.CLASSID IN ('ARP200','ARC200','ARS200','ARS210','BPK010','BPK020','BPK030','BPK040','ARG100','BFH010','BCK010')  and    B.ClassID <> ('014') ");
                    }

                    else if (gBTYPETextBox.Text == "通路")
                    {
                        sb.Append("      and   B.ClassID ='014' ");
                    }
                }

                if (KEY == "1")
                {
                    sb.Append("   and A.CustomerID= 'TW90144-94' ");
                }
                else if (KEY == "2")
                {
                    sb.Append("   and A.CustomerID= 'TW90146-89' ");
                }
                else
                {
                    sb.Append("   and A.CustomerID not in  ('TW90144-94','TW90146-89') ");
                }
            }
            else
            {
                sb.Append("   and A.BillNO between  @BillNO1 and @BillNO2 ");
            }


            sb.Append("             Order By A.BillDate,A.BillNO, G.ProdID,G.ROWNO");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateDate2", CreateDate2));
            command.Parameters.Add(new SqlParameter("@CUST", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@BillNO1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillNO2", textBox6.Text));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable DD2(string CreateDate, string CreateDate2)
        {

            SqlConnection MyConnection = new SqlConnection(strCn03);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select  G.ProdID,a.billdate,H.CLASSID , A.FundBillNo  BILLNO,'' CUSTBILLNO,G.PRODNAME,  ");
            sb.Append(" CASE ISNULL(T0.CombSubID,'')  WHEN '' THEN G.ProdID  ELSE T0.CombSubID END ITEMCODE,G.RowNO ROWNO,");
            sb.Append(" '' 交易方式,    ");
            sb.Append(" CASE ISNULL(T0.CombSubID,'') WHEN '' THEN H.InvoProdName ELSE ISNULL(T1.InvoProdName,'') END   ITEMNAME,    ");
            sb.Append(" CASE ISNULL(T0.CombSubID,'') WHEN '' THEN G.Quantity ELSE T0.Amount*G.Quantity   END  QTY,    ");
            sb.Append(" CASE ISNULL(T0.CombSubID,'') WHEN '' THEN G.Quantity ELSE T0.Amount*G.Quantity   END  QTY2,     ");
            sb.Append(" A.CustAddress ,    ");
            sb.Append("             B.FullName    LinkMan,     ");
            sb.Append("               A.CustID  LinkTelephone ,  ");
            sb.Append(" '' UserDef1, T11.FaxNo,   ");
            sb.Append(" '' PreInDate,H.BARCODEID BARCODEID,    ");
            sb.Append(" '' UserDef2 ,G.PRICE,    ");
            sb.Append(" 0 AMT,A.CustID  CARDCODE,B.FullName CARDNAME,A.BillDate BILLDATE,     ");
            sb.Append(" '' ORDMAN,     ");
            sb.Append(" CAST(A.AddrID   AS INT)  ADDRID,     ");
            sb.Append(" Case when H.ClassID ='AWC200' then '朝貢雞'     ");
            sb.Append(" when H.ClassID ='ARC200' then '朝貢雞'     ");
            sb.Append(" when H.ClassID ='AWP200' then '朝貢豬'     ");
            sb.Append(" when H.ClassID ='ARP200' then '朝貢豬'     ");
            sb.Append(" when H.ClassID ='AWS200' then '大力蝦'     ");
            sb.Append(" when H.ClassID ='ARS200' then '大力蝦'     ");
            sb.Append(" when H.ClassID ='AWS210' then '白金蝦'     ");
            sb.Append(" when H.ClassID ='ARS210' then '白金蝦'     ");
            sb.Append(" when H.ClassID ='ARG100' then '禮盒'     ");
            sb.Append(" when SUBSTRING(G.ProdID,1,1)='P' THEN '加工品'      ");
            sb.Append(" else '空白' end 產品別,0 AMT2,G.ITEMREMARK,A.UDef1 ,A.UDef2,A.Remark REMARK,G.Amount  AMT4 ,L.ClassName 客戶類別,B.ClassID CLASS,H.UNIT    ");
            sb.Append(" From ComProdRec G");
            sb.Append(" Inner Join comBillAccounts  A ON G.BillNO=A.FundBillNo AND CASE G.Flag WHEN 701 THEN 698 ELSE G.Flag END=A.Flag     ");
            sb.Append("               Left Join comCustomer B On B.Flag=1 And B.ID=A.CustID           ");
            sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan     ");
            sb.Append(" Left Join comProduct H On H.ProdID=G.ProdID     ");
            sb.Append(" Left Join comProdCombine T0 On G.ProdID=T0.ProdID     ");
            sb.Append(" Left Join comProduct T1 On T0.CombSubID=T1.ProdID     ");
            sb.Append(" LEFT JOIN comCustAddress T11 ON (A.AddrID=T11.AddrID AND A.CustID=T11.ID )   ");
            sb.Append(" Left Join comCustClass L On L.ClassID =b.ClassID and L.Flag =1 ");
            sb.Append(" Where A.Flag=500 AND G.ProdID<> 'FREIGHT01'  ");
            if (textBox5.Text == "")
            {
                sb.Append("             And  A.BILLDATE Between @CreateDate and @CreateDate2 ");

            }
            else
            {
                sb.Append("   and A.FundBillNo between  @BillNO1 and @BillNO2 ");
            }


            sb.Append("             Order By A.BillDate,A.FundBillNo, G.ProdID,G.ROWNO");
            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateDate2", CreateDate2));
            command.Parameters.Add(new SqlParameter("@CUST", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@BillNO1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillNO2", textBox6.Text));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable DD23T(string BILLNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BILLNO FROM OrdBillSub WHERE BILLNO=@BILLNO AND SUBSTRING(ProdID,1,3) = 'FRE' ");
          


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
          
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable GETPRODNAME(string ProdID)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT InvoProdName   FROM comProduct WHERE ProdID=@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }

        public System.Data.DataTable DDSMANF(string ITEMCODE, string TYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT Convert(varchar(8),DEADLINE,112)   DEADLINE,DEADLINE DEADLINE2 FROM GB_DEADLINE WHERE ITEMCODE=@ITEMCODE  ");

            if (TYPE == "D1")
            {
                sb.Append("    AND D1='TRUE'  ");
            }

            if (TYPE == "D2")
            {
                sb.Append("    AND D2='TRUE'  ");
            }

            if (TYPE == "D3")
            {
                sb.Append("    AND D3='TRUE'  ");
            }

            if (TYPE == "D4")
            {
                sb.Append("    AND D4='TRUE'  ");
            }
            if (TYPE == "D5")
            {
                sb.Append("    AND D5='TRUE'  ");
            }
            if (TYPE == "D6")
            {
                sb.Append("    AND D6='TRUE'  ");
            }
            if (TYPE == "D7")
            {
                sb.Append("    AND D7='TRUE'  ");
            }
            if (TYPE == "D8")
            {
                sb.Append("    AND D8='TRUE'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable DD23(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY2) QTY FROM GB_PICK T0");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE ");
            sb.Append(" GROUP BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }

        public System.Data.DataTable GETRIVA(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                 SELECT RIVANOTE FROM GB_POTATO WHERE CAST(ID AS VARCHAR)=@ID AND ISNULL(RIVANOTE,'') <>'' ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        
                textBox1.Text = GetMenu.Day();
                textBox2.Text = GetMenu.Day();
            
        }


        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHIPCARD();

            if (LookupValues != null)
            {
                textBox4.Text = Convert.ToString(LookupValues[0]);
                textBox3.Text = Convert.ToString(LookupValues[1]);
            }
        }



        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            gBTYPETextBox.Text = comboBox3.Text;
        }

        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.Getfee("GBTYPE");
            comboBox3.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelProduct4(FileName);

            }
        }
        private void WriteExcelProduct4(string ExcelFile)
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
                string ITEMCODE;
                string KG;
                string PACK;
                string CAL;

                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    KG = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    PACK = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    CAL = range.Text.ToString().Trim();

                    CAL = Environment.NewLine + "for 出貨材積計算：##" + CAL + "##";

                    if (ITEMCODE != "")
                    {

                        if (ITEMCODE != "料號")
                        {
                            if (!String.IsNullOrEmpty(CAL))
                            {
                                AddTEMP21(ITEMCODE, CAL, Convert.ToInt32(PACK), Convert.ToDecimal(KG));
                            }
              

                        }
                    }
                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
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


                System.Diagnostics.Process.Start(NewFileName);


            }



        }
        public void AddTEMP21(string ProdID, string ProdDesc, int PackAmt1,decimal  CtmWeight)
        {
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand("UPDATE comProduct SET ProdDesc=ProdDesc+@ProdDesc,PackAmt1=@PackAmt1,CtmWeight=@CtmWeight,CtmUnit='公克',PackUnit1='箱'   WHERE ProdID =@ProdID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdDesc", ProdDesc));
            command.Parameters.Add(new SqlParameter("@PackAmt1", PackAmt1));
            command.Parameters.Add(new SqlParameter("@CtmWeight", CtmWeight));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        public void DELPICK(string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("DELETE  gB_PICK WHERE SHIPPINGCODE=@SHIPPINGCODE DELETE  gB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE", connection);
            command.CommandType = CommandType.Text;
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
        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog3.FileName;


                WriteExcelGB(FileName);

            }
        }
        private void WriteExcelGB(string ExcelFile)
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

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string BILLNO;
                string EZNO;
      
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    BILLNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    EZNO = range.Text.ToString().Trim().Replace("'", "");



                    if (BILLNO != "")
                    {

                        if (BILLNO != "訂單編號")
                        {
                            if (BILLNO.Length > 9)
                            {
                                BILLNO = BILLNO.Substring(0, 10);
                                AddTEMPG1(EZNO, BILLNO);
                            }
                        }
                    }
                }




            }
            finally
            {

     
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
        public void AddTEMPG1(string STORE, string BILLNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE GB_PICK2 SET STORE=@STORE WHERE BILLNO=@BILLNO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@STORE", STORE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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


        private void button6_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gB_PICK2DataGridView);
        }



    

    }
}

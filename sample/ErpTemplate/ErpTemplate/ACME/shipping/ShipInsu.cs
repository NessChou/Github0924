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

namespace ACME
{
    public partial class ShipInsu : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public ShipInsu()
        {
            InitializeComponent();
        }

        private System.Data.DataTable OrderData;
        private void ShipInsu_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(shipyearComboBox, "shipyear", shipBuyBindingSource, "shipyear");
            UtilSimple.SetLookupBinding(shipmonthComboBox, "shipmonth", shipBuyBindingSource, "shipmonth");
            UtilSimple.SetLookupBinding(shipdateComboBox, "shipdate", shipBuyBindingSource, "shipdate");
            // TODO: 這行程式碼會將資料載入 'ship.ShipBuy' 資料表。您可以視需要進行移動或移除。
            this.shipBuyTableAdapter.Fill(this.ship.ShipBuy);
            string aa = "//acmesrv01//SAP_Share//shipping//pic//";
            // TODO: 這行程式碼會將資料載入 'ship.temp' 資料表。您可以視需要進行移動或移除。
            if (fmLogin.LoginID.ToString() == "lleytonchen" || fmLogin.LoginID.ToString() == "joychen")
            {
                groupBox1.Visible = true;
            }
            textBox3.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox1.Text = GetMenu.DFirst();
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();
            System.Data.DataTable dt1 = SumOwhr();
            DataRow drw = dt1.Rows[0];
            if (drw["temp1"].ToString() == "1")
            {

                button2.Image = Image.FromFile(aa+"Yes.gif");
            }
            else
            {
                button2.Image = Image.FromFile(aa+"cancel.gif");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {

            Microsoft.VisualBasic.Devices.Computer computer = new Computer();
            if (computer.Name.ToString().ToUpper() == "ACMEW08R2RDP")
            {
         
                    OrderData = ExecuteQuerySZ();

                
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\close.xls";

                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                //string FileName = string.Empty;
                //string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                //FileName = lsAppDir + "\\Excel\\close.xls";

                if (radioButton3.Checked)
                {
                    //取得 Excel 資料
                    OrderData = ExecuteQueryclose();
                }
                else
                {
                    OrderData = ExecuteQuery();

                }
                string GG = @"\SHIPPING" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

                dataGridView2.DataSource = OrderData;
                //  ExcelReport.GridViewToCSV(dataGridView2,Environment.CurrentDirectory + @"\SHIPPING.csv");
                ExcelReport.GridViewToCSV(dataGridView2, Environment.CurrentDirectory + "\\Excel\\temp\\" + GG);
            }
            ////Excel的樣版檔
            //string ExcelTemplate = FileName;

            ////輸出檔
            //string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
            //      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            ////產生 Excel Report
            //ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
        }
        private System.Data.DataTable ExecuteQuery()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select a.cardname 客戶名稱,a.shippingcode 工單號碼,a.closeday 結關日,a.ForecastDay 開航日,a.arriveday  抵達日,a.shipment 裝船港,a.unloadCargo 卸貨港,");
            sb.Append(" b.itemcode 品名,sum(b.Quantity) 數量,a.notifyMemo 異常通知 ");
            sb.Append(" from shipping_main a left join shipping_item b on(a.shippingcode=b.shippingcode)");
            sb.Append("where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" group by a.shippingcode,a.cardname,a.closeday ,a.ForecastDay ,a.arriveday ,a.shipment ,a.unloadCargo ,");
            sb.Append(" b.itemcode  ,a.notifyMemo  order by a.closeday");
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

        private System.Data.DataTable ExecuteQuerySZ()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select a.cardname 客戶名稱,a.shippingcode 工單號碼,a.closeday 結關日,a.ForecastDay 開航日,a.arriveday  抵達日,a.shipment 裝船港,a.unloadCargo 卸貨港,");
            sb.Append(" b.itemcode 品名,sum(b.Quantity) 數量,a.notifyMemo 異常通知,a.boatCompany 船公司 ");
            sb.Append(" from shipping_main a left join shipping_item b on(a.shippingcode=b.shippingcode)");
            sb.Append("where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" group by a.shippingcode,a.cardname,a.closeday ,a.ForecastDay ,a.arriveday ,a.shipment ,a.unloadCargo ,");
            sb.Append(" b.itemcode  ,a.notifyMemo,a.boatCompany  order by a.closeday");
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
        private System.Data.DataTable ExecuteQueryclose()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("            select a.cardname 客戶名稱,a.shippingcode 工單號碼,a.closeday 結關日,a.ForecastDay 開航日,a.arriveday  抵達日,a.shipment 裝船港,a.unloadCargo 卸貨港,");
            sb.Append("            b.itemcode 品名,sum(b.Quantity) 數量,a.notifyMemo 異常通知 ");
            sb.Append("            from shipping_main a left join shipping_item b on(a.shippingcode=b.shippingcode)");
            sb.Append("            where a.buCardcode='checked'");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  a.buCardname between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("            group by a.shippingcode,a.cardname,a.closeday ,a.ForecastDay ,a.arriveday ,a.shipment ,a.unloadCargo ,");
            sb.Append("            b.itemcode  ,a.notifyMemo  order by a.closeday");
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

        private void button15_Click(object sender, EventArgs e)
        {

            try
            {

                System.Data.DataTable dt = GetOrderData6();
                DataRow dr = null;
                System.Data.DataTable dtCost = MakeTableINSU();
                string 貿易類別 = "";
                string INV保險費 = "";
                string INV保險費1 = "";
                string 運送方式 = "";
                string SEQNO = "";
                string 保險費 = "";
                string X1 = "";
                int L1 = 0;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    dr = dtCost.NewRow();
                    string job = dt.Rows[i]["工單號碼"].ToString();
                    string invo = dt.Rows[i]["invo"].ToString();
  
                    if (X1 == job)
                    {
                        L1 += 1;

                    }
                    else
                    {
                        L1 = 1;
                    }
               
                    貿易類別 = dt.Rows[i]["貿易類別"].ToString();
                    INV保險費 = dt.Rows[i]["INV保險費"].ToString();
                    INV保險費1 = dt.Rows[i]["INV保險費1"].ToString();
                    運送方式 = dt.Rows[i]["運送方式"].ToString();

                    SEQNO = dt.Rows[i]["SEQNO"].ToString();
                    dr["結關日"] = dt.Rows[i]["結關日"].ToString();
                    dr["貿易類別"] = dt.Rows[i]["貿易類別"].ToString();
                    dr["工單號碼"] = job;
                    dr["invo"] = invo;
                    dr["品名"] = dt.Rows[i]["品名"].ToString();
                    dr["數量"] = dt.Rows[i]["數量"].ToString();
            
                    dr["貿易條件"] = dt.Rows[i]["貿易條件"].ToString();
                    dr["客戶"] = dt.Rows[i]["客戶"].ToString();
                    dr["價格"] = dt.Rows[i]["價格"].ToString();
                    string INSUCHECK = dt.Rows[i]["INSUCHECK"].ToString().Trim();
                    string INSSHIPWAY = dt.Rows[i]["INSSHIPWAY"].ToString().Trim();
                   

                    保險費 = dt.Rows[i]["保險費"].ToString();
                    X1 = job;
                    if (INSUCHECK == "Checked")
                    {
                        if (!String.IsNullOrEmpty(INSSHIPWAY))
                        {

                            運送方式 = INSSHIPWAY;
                            INV保險費 = INV保險費1;
                        }
                    }
                    dr["運送方式"] = 運送方式;
                    if (!String.IsNullOrEmpty(INV保險費))
                    {
                        decimal  G1 = Convert.ToDecimal(INV保險費);
                        if (SEQNO == "0" && L1.ToString()=="1")
                        {
  
                            dr["保險費"] = INV保險費;

                        }
                        else
                        {
                            dr["保險費"] = "-";
                        }

                        if (運送方式 == "SEA" || 運送方式 == "AIR")
                        {
                            if (SEQNO == "0" && L1.ToString() == "1")
                            {
                                if (G1 < 5)
                                {
                         
                                        dr["保險費"] = "5";
                                    
                                    
                                }
                            }
                            else
                            {
                                dr["保險費"] = "-";
                            }
                        }
                        else
                        {

                            if (貿易類別 == "三角" && 運送方式.ToUpper() == "TRUCK")
                            {

                                if (SEQNO == "0")
                                {
                                    if (G1 < 2)
                                    {
                                
                                            dr["保險費"] = "2";
                                        
                                    }
                                }
                                else
                                {
                                    dr["保險費"] = "-";
                                }
                            }
                            else
                            {
                                dr["保險費"] = 保險費;
                            }

                        }

                    }

                    dtCost.Rows.Add(dr);
                }
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\SHIP\\INSU.xlsx";
               // FileName = lsAppDir + "\\Excel\\INSU.xls";
                OrderData = dtCost;
                GetExcelinsuWORK(FileName);
         

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable GetOrderData6()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" select * from (       ");
            sb.Append("     select a.shippingcode 工單號碼,a.tradeCondition 貿易條件 ,a.boardcountno 貿易類別 ,b.invoiceno+b.invoiceno_seq invo,b.INDescription 品名");
            sb.Append("                    ,b.InQty 數量,a.receiveday 運送方式,a.cardname 客戶,a.closeday 結關日,b.amount 價格,SEQNO,a.INSUCHECK,a.INSSHIPWAY,");
            sb.Append("                    保險費=case a.receiveday ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.amount as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.amount as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end,");
            sb.Append("                    INV保險費=case a.receiveday ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end,");
            sb.Append("                    INV保險費1=case a.INSSHIPWAY ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end");
            sb.Append("                    from shipping_main a ");
            sb.Append("                    INNER join invoiceM C on(A.shippingcode=C.shippingcode)");
            sb.Append("                     INNER join invoiced b on(C.shippingcode=b.shippingcode AND C.INVOICENO=B.INVOICENO AND C.INVOICENO_SEQ=B.INVOICENO_SEQ)");
            sb.Append("                    INNER JOIN (SELECT SUM(AMOUNT) INV金額,SHIPPINGCODE FROM invoiced WHERE INDescription NOT LIKE  '%FREIGHT%'   GROUP BY SHIPPINGCODE) D on (D.shippingcode=b.shippingcode )");
            sb.Append("                     WHERE A.CFS='checked' AND b.InQty <>0 AND b.INDescription  NOT LIKE  '%FREIGHT%'    and isnull(add1,'') not like '%正航%'   AND A.QUANTITY<>'取消' ");
            sb.Append(" and  SUBSTRING(a.shippingcode,1,2) <>'SI'");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

            sb.Append(" union all");
            sb.Append("            select a.shippingcode 工單號碼,a.tradeCondition 貿易條件 ,a.boardcountno 貿易類別 ,b.invoiceno+b.invoiceno_seq invo,b.INDescription 品名");
            sb.Append("                    ,b.InQty 數量,a.receiveday 運送方式,a.cardname 客戶,a.closeday 結關日,b.CHOamount 價格,SEQNO,a.INSUCHECK,a.INSSHIPWAY,");
            sb.Append("                    保險費=case a.receiveday ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.CHOamount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.CHOamount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.CHOamount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.CHOamount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.CHOamount as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.CHOamount as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end,");
            sb.Append("                    INV保險費=case a.receiveday ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end,");
            sb.Append("                    INV保險費1=case a.INSSHIPWAY ");
            sb.Append("                    when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end");
            sb.Append("                    when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end");
            sb.Append("                    from shipping_main a ");
            sb.Append("                    INNER join invoiceM C on(A.shippingcode=C.shippingcode)");
            sb.Append("                     INNER join invoiced b on(C.shippingcode=b.shippingcode AND C.INVOICENO=B.INVOICENO AND C.INVOICENO_SEQ=B.INVOICENO_SEQ)");
            sb.Append("                    INNER JOIN (SELECT SUM(CHOAMOUNT) INV金額,SHIPPINGCODE FROM invoiced WHERE INDescription NOT LIKE  '%FREIGHT%'   GROUP BY SHIPPINGCODE) D on (D.shippingcode=b.shippingcode )");
            sb.Append("                     WHERE A.CFS='checked' AND b.InQty <>0 AND b.INDescription  NOT LIKE  '%FREIGHT%'    and isnull(add1,'')  like '%正航%'    AND A.QUANTITY<>'取消' ");
            sb.Append(" and  SUBSTRING(a.shippingcode,1,2) <>'SI'");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

            sb.Append(" union all");
          sb.Append(" select a.shippingcode 工單號碼,a.tradeCondition 貿易條件 ,a.boardcountno 貿易類別 ,'' invo,U_ITEMNAME COLLATE  Chinese_Taiwan_Stroke_CI_AS +' '+REPLACE(ISNULL(U_MODEL,''),'NON','')   品名 ");
sb.Append(" ,b.Quantity  數量,a.receiveday 運送方式,a.cardname 客戶,a.closeday 結關日,b.ItemAmount  價格,SEQNO,a.INSUCHECK,a.INSSHIPWAY, ");
sb.Append(" 保險費=case a.receiveday  ");
sb.Append(" when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.ItemAmount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.ItemAmount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.ItemAmount as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.ItemAmount as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(b.ItemAmount as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(b.ItemAmount as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end, ");
sb.Append(" INV保險費=case a.receiveday  ");
sb.Append(" when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end, ");
sb.Append(" INV保險費1=case a.INSSHIPWAY  ");
sb.Append(" when 'SEA'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'AIR'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*1.1*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*1.1*cast(a.shiptodate as decimal(6,3))/100 end else 0 end ");
sb.Append(" when 'TRUCK'  then case a.cfs when 'checked' then case a.shiptodate when '其他費率' then cast(D.INV金額 as decimal(12,3))*cast(a.ModifyDate as decimal(12,3))/100 else cast(D.INV金額 as numeric)*cast(a.shiptodate as decimal(6,3))/100 end else 0 end end ");
sb.Append(" from shipping_main a  ");
sb.Append(" INNER join Shipping_Item  b on(A.shippingcode=b.shippingcode) ");
sb.Append(" INNER JOIN (SELECT SUM(ItemAmount) INV金額,SHIPPINGCODE FROM Shipping_Item  WHERE Dscription NOT LIKE  '%FREIGHT%'  GROUP BY SHIPPINGCODE) D on (D.shippingcode=b.shippingcode ) ");
sb.Append("  LEFT JOIN  ACMESQL02.DBO.OITM T2 ON (b.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
sb.Append(" WHERE A.CFS='checked' AND b.Quantity  <>0 AND b.Dscription  NOT LIKE  '%FREIGHT%'   and isnull(add1,'') not like '%正航%'   AND A.QUANTITY<>'取消'  ");
sb.Append(" and  SUBSTRING(a.shippingcode,1,2) ='SI'");

            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" ) as A ORDER BY 工單號碼,CAST(SEQNO AS INT)");
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


        private System.Data.DataTable GetOrderData7()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select a.boardCountNo 貿易條件 ,count(*) 筆數 from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("  group by a.boardCountNo order by a.boardCountNo ");
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
        private System.Data.DataTable GetOrderData19()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("            SELECT  t1.docentry DOC,T1.LINENUM LINE,SUBSTRING(CLOSEDAY,0,7) AA,T3.SHIPPINGCODE JOBNO,CASE T3.[CardCODE] WHEN '0511-00' THEN  T3.[ADD6] WHEN '0257-00' THEN T3.[ADD6] ELSE T3.[CardName] END 客戶, T1.ITEMCODE 料號, T1.[Dscription] 說明,");
            sb.Append("                        sum(t4.QUANTITY) 數量 ,");
            sb.Append("                       t1.linetotal 金額,T3.boardcountno 貿易條件,T3.receiveDay 運送方式,T2.[SLPNAME] 業務");
            sb.Append(" ,MAX(CASE ITMSGRPCOD ");
            sb.Append("                 WHEN 1032 THEN 'TFT事業部' ");
            sb.Append("                 WHEN 1033 THEN 'LED事業部' ");
            sb.Append("                  WHEN 233 THEN 'LED事業部' ");
            sb.Append("                 WHEN 1034 THEN 'PCBA事業部' ");
            sb.Append("                WHEN 102 THEN '太陽能事業部'  END  )  部門 ,T2.[SlpName] 業務 FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append("                           LEFT JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("                        left join acmesqlsp.dbo.shipping_main T3 on (t0.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                          left join acmesqlsp.dbo.shipping_item T4 on (T3.shippingcode=T4.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS )");
            sb.Append("                        LEFT JOIN acmesql02.dbo.OSLP T2 ON T0.SLPCODE = T2.SLPCODE     ");
            sb.Append(" LEFT JOIN         acmesql02.dbo.OITM T5 ON T4.ITEMCODE = T5.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS");
            sb.Append("                            where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T3.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  T3.SHIPPINGCODE between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("                       group by t1.linetotal,t1.docentry,T3.SHIPPINGCODE ,T3.[CardCODE],T3.[ADD6] ,T1.LINENUM, T3.[CardName], T1.ITEMCODE , T1.[Dscription],T3.boardcountno,T3.receiveDay,T2.[SLPNAME],T3.CLOSEDAY            ORDER BY T3.SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
        private System.Data.DataTable GetOrderData20(string aa,string bb,string cc)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT  t1.docentry DOC,T1.LINENUM LINE,SUBSTRING(CLOSEDAY,0,7) AA,T3.SHIPPINGCODE JOBNO,CASE T3.[CardCODE] WHEN '0511-00' THEN  T3.[ADD6] WHEN '0257-00' THEN T3.[ADD6] ELSE T3.[CardName] END 客戶, T1.ITEMCODE 料號, T1.[Dscription] 說明,T3.[CardCODE] 客戶編號,");
            sb.Append("               sum(t4.QUANTITY) 數量 ,");
            sb.Append("              t1.linetotal 金額,T3.boardcountno 貿易條件,T3.receiveDay 運送方式,T2.[SLPNAME] 業務 FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append("                  LEFT JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("               left join acmesqlsp.dbo.shipping_main T3 on (t0.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                 left join acmesqlsp.dbo.shipping_item T4 on (T3.shippingcode=T4.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS )");
            sb.Append("               LEFT JOIN acmesql02.dbo.OSLP T2 ON T0.SLPCODE = T2.SLPCODE            ");
            sb.Append("               where T3.SHIPPINGCODE = @aa AND T1.LINENUM=@bb AND t1.docentry=@cc ");
            sb.Append("              group by t1.linetotal,t1.docentry,T3.SHIPPINGCODE ,T3.[CardCODE],T3.[ADD6] ,T1.LINENUM, T3.[CardName], T1.ITEMCODE , T1.[Dscription],T3.boardcountno,T3.receiveDay,T2.[SLPNAME],T3.CLOSEDAY");
            
                        SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@cc", cc));
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
        private System.Data.DataTable GetOrderData17()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select a.boardCountNo 貿易條件 ,count(*) 筆數 from shipping_main a  where ");
            sb.Append(" closeday between @startday and @endday group by a.boardCountNo ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startday", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@endday ", textBox3.Text));

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
        private System.Data.DataTable GetOrderData9()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select a.boardCountNo 貿易條件 ,count(*) 筆數 from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("   and substring(cardcode,1,1)='S' group by a.boardCountNo order by a.boardCountNo ");
           
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
        private System.Data.DataTable GetOrderData8()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select a.shippingcode 工單號碼,a.boardCountNo 貿易類別,case when cardcode in ('0257-00','0511-00') then add6  else a.cardname end 客戶  from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("  order by a.boardCountNo  ");
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
        private System.Data.DataTable GetOrderData8S()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select a.shippingcode 工單號碼,a.boardCountNo 貿易類別,A.ADD7 所有人  from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("  order by a.boardCountNo  ");
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
        private System.Data.DataTable GetOrderData10()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select a.shippingcode 工單號碼,a.boardCountNo 貿易類別,A.ADD7 所有人  from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("   and substring(cardcode,1,1)='S' order by a.boardCountNo  ");
             
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

        private System.Data.DataTable GetOrderData10S()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select a.shippingcode 工單號碼,a.boardCountNo 貿易類別,A.ADD7 所有人  from shipping_main a  where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  closeday  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  a.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("   and substring(cardcode,1,1)='S' order by a.boardCountNo  ");

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
        private void GetExcelinsu(string ExcelFile)
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
                            DetailRow1 = 24;
                            DetailRow2 = 24;
                            DetailRow3 = 24;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
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
                            SetRow(aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }

                for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData.Rows.Count - 1)
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
                        SetRow(aRow, sTemp, ref FieldValue);

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

        private void GetExcelinsuWORK(string ExcelFile)
        {
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false ;
            excelApp.DisplayAlerts = false;
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
                            DetailRow1 = 24;
                            DetailRow2 = 24;
                            DetailRow3 = 24;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
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
                            SetRow(aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }

                for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData.Rows.Count - 1)
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
                        SetRow(aRow, sTemp, ref FieldValue);

                        range.Value2 = FieldValue;


                    }

                    DetailRow++;
                }
                for (int aRow = 1; aRow <= OrderData.Rows.Count + 1; aRow++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 9]);
                    range.Select();
                    string a = (string)range.Text;
                    a = a.Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                    range.Select();
                    string b = (string)range.Text;
                    b = b.Trim();

                    int n;

                    if (int.TryParse(a, out n) && int.TryParse(b, out n))
                    {

                        if (Convert.ToInt16(b) > Convert.ToInt16(a))
                        {
                            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }
                }

                for (int aRow = 1; aRow <= OrderData.Rows.Count + 1; aRow++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 9]);
                    range.Select();
                    string a = (string)range.Text;
                    a = a.Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                    range.Select();
                

                    int n;

                    if (a=="")
                    {

                
                            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        
                    }
                }
            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(ExcelFile).Replace("\\SHIP","") + "\\temp\\" +
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

        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (radioButton1.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade1.xls";

                    OrderData = GetOrderData7();

                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
                else if (radioButton2.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade2.xls";

                    OrderData = GetOrderData8();
                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
                else if (radioButton4.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade3.xls";

                    OrderData = GetOrderData8S();
                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    

        private void button2_Click(object sender, EventArgs e)
        {
            string aa = "//acmesrv01//SAP_Share//shipping//pic//";
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            System.Data.DataTable dt1 = SumOwhr();
            DataRow drw = dt1.Rows[0];
            if (drw["temp1"].ToString() == "1")
            {
                UpdateMasterSQL1("2");
                button2.Image = Image.FromFile(aa + "cancel.gif");
                MessageBox.Show("排程日期檢查已關閉");
            }
            else
            {
                UpdateMasterSQL1("1");
               
                button2.Image = Image.FromFile(aa + "Yes.gif");
                MessageBox.Show("排程日期檢查已開啟");
            }
        }

        private void UpdateMasterSQL1(string temp1)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  update temp set temp1=@temp1 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@temp1", temp1));
  

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
        public static System.Data.DataTable SumOwhr()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from temp";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "temp");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["temp"];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ViewBatchPayment();

            GridViewToExcel(dataGridView1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ViewBatchPayment1();

            GridViewToExcel(dataGridView1);
        }
        private void ViewBatchPayment()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T3.SHIPPINGCODE JOBNO,CASE T3.[CardCODE] WHEN '0511-00' THEN  T3.[ADD6] WHEN '0257-00' THEN T3.[ADD6] ELSE T3.[CardName] END 客戶, T1.ITEMCODE 料號, T1.[Dscription] 說明,");
            sb.Append("  sum(t4.QUANTITY) 數量 ,");
            sb.Append(" t1.linetotal 金額,T3.boardcountno 貿易條件,T2.[SLPNAME] 業務 FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append("              LEFT JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              left join acmesqlsp.dbo.shipping_main T3 on (t0.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("    left join acmesqlsp.dbo.shipping_item T4 on (T3.shippingcode=T4.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("  LEFT JOIN acmesql02.dbo.OSLP T2 ON T0.SLPCODE = T2.SLPCODE            ");
            sb.Append("  where T0.[U_Shipping_no] <> '' and T3.CLOSEDAY between @startdate and @enddate");
            sb.Append(" group by t1.linetotal,t1.docentry,T3.SHIPPINGCODE ,T3.[CardCODE],T3.[ADD6] ,T1.LINENUM, T3.[CardName], T1.ITEMCODE , T1.[Dscription],T3.boardcountno,T2.[SLPNAME]");
      
            sb.Append(" order by T3.SHIPPINGCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

            command.Parameters.Add(new SqlParameter("@startdate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@enddate", textBox3.Text));
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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }
        private void ETSAI()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  MAX(CASE WHEN REPLACE(REPLACE(ISNULL(MEMO2,''),CHAR(10),''),CHAR(13),'') <> '' THEN '有'   ELSE '無'  END) 異常,COUNT(CASE WHEN REPLACE(REPLACE(ISNULL(MEMO2,''),CHAR(10),''),CHAR(13),'') <> '' THEN '有' ELSE '無' END) ' ' ");
            sb.Append(" FROM SHIPPING_MAIN WHERE 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  SHIPPINGCODE between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" GROUP  BY CASE WHEN REPLACE(REPLACE(ISNULL(MEMO2,''),CHAR(10),''),CHAR(13),'') <> '' THEN '有' ELSE '無' END");

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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView2.DataSource = bindingSource1;

        }
        private void ETSAI2()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE SUBSTRING(CARDCODE,1,1) WHEN 'S' THEN 'S' ELSE 'C' END '供應商/客戶'");
            sb.Append(" ,SUBSTRING(tradeCondition,1,1)  貿易條件,SHIPPINGCODE 單號,receiveDay 運送方式");
            sb.Append(" ,boardCountNo 貿易形式 FROM SHIPPING_MAIN  WHERE REPLACE(REPLACE(ISNULL(MEMO2,''),CHAR(10),''),CHAR(13),'') <> ''");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  SHIPPINGCODE between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView2.DataSource = bindingSource1;

        }
        private void ViewBatchPayment1()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T3.SHIPPINGCODE JOBNO, T1.ITEMCODE 料號, T1.[Dscription] 說明,SUBSTRING(Convert(varchar(8),t0.docdate,112),5,2) 月份,T3.FORECASTDAY ETD, SUM(T1.[LineTotal]) 金額 FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append(" LEFT JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left join acmesqlsp.dbo.shipping_main T3 on (t0.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where T0.[U_Shipping_no] <> ''  ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T3.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  T3.SHIPPINGCODE between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" GROUP BY T3.SHIPPINGCODE ,T0.DOCDATE , T1.ITEMCODE , T1.[Dscription] ,T3.FORECASTDAY, t3.closeday");

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


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("說明", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("業務", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTableINSU()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("結關日", typeof(string));
            dt.Columns.Add("貿易類別", typeof(string));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("invo", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("運送方式", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("價格", typeof(string));
            dt.Columns.Add("保險費", typeof(string));
            return dt;
        }
        private void button5_Click(object sender, EventArgs e)
        {

            try
            {


                string 單號;
                string 料號;
                string SAP;
                System.Data.DataTable dt = GetOrderData19();
                System.Data.DataTable dtCost = MakeTableCombine();
                System.Data.DataTable dtDoc = null;
                string DuplicateKey = "";
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    單號 = dt.Rows[i]["JOBNO"].ToString();
                    料號 = dt.Rows[i]["LINE"].ToString();
                    SAP = dt.Rows[i]["DOC"].ToString();
                    dtDoc = GetOrderData20(單號, 料號, SAP);

                    dr = dtCost.NewRow();
                    dr["JOBNO"] = Convert.ToString(dtDoc.Rows[0]["JOBNO"]);
                    dr["客戶"] = Convert.ToString(dtDoc.Rows[0]["客戶"]);
                    dr["說明"] = Convert.ToString(dtDoc.Rows[0]["說明"]);
                    dr["金額"] = Convert.ToString(dtDoc.Rows[0]["金額"]);
                    dr["貿易條件"] = Convert.ToString(dtDoc.Rows[0]["貿易條件"]);
                    dr["業務"] = Convert.ToString(dtDoc.Rows[0]["業務"]);
                    

                    if (單號 != DuplicateKey)
                    {

                        dr["數量"] = Convert.ToString(dtDoc.Rows[0]["數量"]);

                    }
                    DuplicateKey = 單號;

                    dtCost.Rows.Add(dr);

                }
                bindingSource1.DataSource = dtCost;
                dataGridView2.DataSource = bindingSource1.DataSource;
                GridViewToExcel(dataGridView2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GridViewToExcel(DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;

            try
            {

               

                for (int i = 0; i < dgv.Columns.Count; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;


                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                        }

                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);

                        }

                    }

                }

                wapp.Visible = true;


            }

            catch (Exception ex1)
            {

                MessageBox.Show(ex1.Message);

            }

            wapp.UserControl = true;

        }
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton1.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade1.xls";

                    OrderData = GetOrderData9();
                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
                else if (radioButton2.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade2.xls";

                    OrderData = GetOrderData10();
                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
                else if (radioButton4.Checked)
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\trade3.xls";

                    OrderData = GetOrderData10S();
                    dataGridView2.DataSource = OrderData;
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            UpdateSQL(buyTextBox.Text, sellTextBox.Text, shipyearComboBox.Text, shipmonthComboBox.Text, shipdateComboBox.Text);
            MessageBox.Show("更新成功");
        }
        private void UpdateSQL(string shippingcode, string notifymemo, string shipyear, string shipmonth, string shipdate)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update shipbuy set Buy=@Buy,Sell=@Sell,shipyear=@shipyear,shipmonth=@shipmonth,shipdate=@shipdate");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@Buy", shippingcode));
            command.Parameters.Add(new SqlParameter("@Sell", notifymemo));
            command.Parameters.Add(new SqlParameter("@shipyear", shipyear));
            command.Parameters.Add(new SqlParameter("@shipmonth", shipmonth));
            command.Parameters.Add(new SqlParameter("@shipdate", shipdate));
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

        private void button8_Click(object sender, EventArgs e)
        {

            try
            {

                System.Data.DataTable dt = Getwork();
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = null;
                    System.Data.DataTable dtCost = MakeTableWORK();

                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {

                        dr = dtCost.NewRow();

                        dr["供應商/客戶"] = dt.Rows[i]["供應商/客戶"].ToString();
                        dr["JOBNO"] = dt.Rows[i]["JOBNO"].ToString();
                        dr["貿易條件"] = dt.Rows[i]["貿易條件"].ToString();
                        dr["貿易形式"] = dt.Rows[i]["貿易形式"].ToString();
                        dr["運送方式"] = dt.Rows[i]["運送方式"].ToString();
                        string data1 = dt.Rows[i]["起始日期"].ToString().Replace("/", "");
                        string data2 = dt.Rows[i]["結關日"].ToString().Replace("/", "");
                        string data3 = dt.Rows[i]["結案日期"].ToString().Replace("/", "");
                        string data4 = DateTime.Now.ToString("yyyyMMdd");
                        dr["所有人"] = dt.Rows[i]["所有人"].ToString();
                        dr["起始日期"] = dt.Rows[i]["起始日期"].ToString();
                        dr["結關日"] = dt.Rows[i]["結關日"].ToString();
                        string 結案周期 = "";
                        string 運送周期 = "";
                        string 結案日期 = dt.Rows[i]["結案日期"].ToString();
                        if (結案日期 == "//")
                        {
                            結案日期 = "";
                        }

                        dr["結案日期"] = 結案日期;

                        int f1 = data1.Length;
                        int f2 = data2.Length;
                        int f3 = data3.Length;


                        if (f2 == 8 && f3 == 8)
                        {
                            System.Data.DataTable t1 = Getwork2(data2, data3);
                            if (t1.Rows.Count > 0)
                            {
                                運送周期 = t1.Rows[0][0].ToString();
                            }
                        }
                        if (結案日期 == "")
                        {
                            System.Data.DataTable t1 = Getwork2(data1, data4);
                            if (t1.Rows.Count > 0)
                            {
                                結案周期 = t1.Rows[0][0].ToString();
                            }
                        }
                        else
                        {
                            if (f1 == 8 && f3 == 8)
                            {
                                System.Data.DataTable t1 = Getwork2(data1, data3);
                                if (t1.Rows.Count > 0)
                                {
                                    結案周期 = t1.Rows[0][0].ToString();
                                }
                            }

                        }

                        dr["結案/取消"] = dt.Rows[i]["結案/取消"].ToString();
                        dr["結案周期"] = 結案周期;
                        dr["運送周期"] = 運送周期;
                        dtCost.Rows.Add(dr);
                    }
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\SHIP\\Work.xlsx";
                  //  FileName = lsAppDir + "\\Excel\\Work.xls";

                

                    if (checkBox1.Checked)
                    {
                        dataGridView2.DataSource = dtCost;
                        ExcelReport.GridViewToCSV(dataGridView2, Environment.CurrentDirectory + @"\SHIPPING" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
                    }
                    else
                    {
                        OrderData = dtCost;
                        GetExcelinsuWORK(FileName);
                    }

                }
                else
                {
                    MessageBox.Show("沒有資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        
        }



        private System.Data.DataTable Getwork()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                       SELECT CASE SUBSTRING(T0.CARDCODE,1,1) WHEN 'S' THEN 'S' ELSE 'C' END '供應商/客戶',T0.SHIPPINGCODE JOBNO,SUBSTRING(tradeCondition,1,1) 貿易條件,boardCountNo 貿易形式,receiveDay 運送方式,add7 所有人,");
            sb.Append("                       SUBSTRING(T0.SHIPPINGCODE,3,4)+'/'+SUBSTRING(T0.SHIPPINGCODE,7,2)+'/'+SUBSTRING(T0.SHIPPINGCODE,9,2)  起始日期,isnull(SUBSTRING(CLOSEDAY,1,4)+'/'+SUBSTRING(CLOSEDAY,5,2)+'/'+SUBSTRING(CLOSEDAY,7,2),'')  結關日,");
            sb.Append("                               isnull(SUBSTRING(buCardname,1,4)+'/'+SUBSTRING(buCardname,5,2)+'/'+SUBSTRING(buCardname,7,2),'')  結案日期,");
            sb.Append(" CASE WHEN buCardcode='Checked' THEN 1 WHEN quantity= '已結' THEN 1 WHEN quantity='取消' THEN 3 ELSE 2 END '結案/取消'  ");
            sb.Append("                    FROM SHIPPING_MAIN T0  WHERE 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  T0.SHIPPINGCODE between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" ORDER BY T0.SHIPPINGCODE");
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
        private System.Data.DataTable Getwork2(string AA,string BB)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select count(*) workday From   acmesqlsp.dbo.Y_2004 ");
            sb.Append(" where Convert(varchar(10),date_time,112) ");
            sb.Append(" between '" + AA + "' AND '" + BB+ "' and  isrestday=0");
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
        private void button9_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetAT();
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTable2();
            System.Data.DataTable dt2 = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string 工單號碼=dt.Rows[i]["工單號碼"].ToString();
                string TYPE = dt.Rows[i]["TYPE"].ToString();
                dr = dtCost.NewRow();
                dr["貿易形式"] = "進口";
                dr["運送方式"] = "SEA";
                dr["報單號碼"] = 工單號碼;
                dr["工單號碼"] = dt.Rows[i]["報單號碼"].ToString();
                StringBuilder sb = new StringBuilder();
                if (TYPE == "採購訂單")
                {
                    dt2 = GetAT1(工單號碼);

                    for (int j = 0; j <= dt2.Rows.Count - 1; j++)
                    {
                        DataRow dv = dt2.Rows[j];
                        string GH = dv["原廠"].ToString();
                        if (!String.IsNullOrEmpty(GH))
                        {
                            sb.Append(GH + "/");

                        }
                    }
                }
                else if (TYPE == "調撥單")
                {
                    dt2 = GetAT2(工單號碼);

                    for (int j = 0; j <= dt2.Rows.Count - 1; j++)
                    {
                        DataRow dv = dt2.Rows[j];
                        string GH = dv["原廠"].ToString();
                        if (!String.IsNullOrEmpty(GH))
                        {
                            sb.Append(GH + "/");

                        }
                    }
                }
                if (!String.IsNullOrEmpty(sb.ToString()))
                {
                    sb.Remove(sb.Length - 1, 1);
                }
                dr["原廠INVOICE NO"] = sb.ToString();
                dtCost.Rows.Add(dr);
            }
            dataGridView2.DataSource = dtCost;
            GridViewToExcel(dataGridView2);

        }

        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("貿易形式", typeof(string));
            dt.Columns.Add("運送方式", typeof(string));
            dt.Columns.Add("報單號碼", typeof(string));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("原廠INVOICE NO", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTableWORK()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("供應商/客戶", typeof(string));
            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("貿易形式", typeof(string));
            dt.Columns.Add("運送方式", typeof(string));
            dt.Columns.Add("所有人", typeof(string));
            dt.Columns.Add("起始日期", typeof(string));
            dt.Columns.Add("結關日", typeof(string));
            dt.Columns.Add("結案日期", typeof(string));
            dt.Columns.Add("結案周期", typeof(string));
            dt.Columns.Add("運送周期", typeof(string));
            dt.Columns.Add("結案/取消", typeof(string));
            return dt;
        }

        private System.Data.DataTable GetAT()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT DISTINCT T1.ITEMREMARK TYPE,add9 報單號碼,T0.SHIPPINGCODE 工單號碼");
            sb.Append("  from shipping_main T0");
            sb.Append("  LEFT JOIN shipping_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" where substring(add9,1,2)='AT'");
            sb.Append(" AND T0.arriveDay  BETWEEN '" + textBox5.Text.ToString() + "' AND '" + textBox6.Text.ToString() + "' ");
            sb.Append("  AND receiveDay='SEA' AND boardCountNo='進口'");
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
        private System.Data.DataTable GetAT1(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT T4.U_ACME_INV '原廠' FROM SHIPPING_ITEM T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.POR1 T2 ON (T0.DOCENTRY=T2.DOCENTRY AND T0.LINENUM=T2.LINENUM)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.PDN1 T3 ON (t3.baseentry=T2.docentry and  t3.baseline=t2.linenum and t3.basetype='22'  )");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN T4 ON (t3.DOCENTRY=T4.DOCENTRY   )");
            sb.Append("  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(T4.U_ACME_INV,'') <> ''");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        private System.Data.DataTable GetAT2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT T2.U_ACME_INV '原廠' FROM SHIPPING_ITEM T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OWTR T2 ON (T0.DOCENTRY=T2.DOCENTRY)");
            sb.Append("  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(T2.U_ACME_INV,'') <> ''");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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

        private void button10_Click(object sender, EventArgs e)
        {
            string H1 = textBox7.Text.ToString().ToUpper();
            int I1 = H1.IndexOf("V.");
            int I2 = H1.IndexOf(".");
            string MODEL = "";
            string VER = "";
            if (I2 == -1)
            {
                MessageBox.Show("請輸入正確格式");
                return;
            }

            if (I1 == -1)
            {
                MODEL = H1.Substring(0, I2).Trim();
                VER = H1.Substring(I2 + 1, 1).Trim();
            }
            else
            {
                MODEL = H1.Substring(0, I1).Trim();
                VER = H1.Substring(I2 + 1, 1).Trim();
            }

            System.Data.DataTable T1 = GetF(MODEL, VER);
            if (T1.Rows.Count > 0)
            {
                dataGridView2.DataSource = T1;
                GridViewToExcel(dataGridView2);
            }
            else
            {
                MessageBox.Show("沒有資料");
            }

        }

        private System.Data.DataTable GetF(string MODEL,string VER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("          SELECT DISTINCT add9 報單號碼,ARRIVEDAY  預計抵達日 FROM Shipping_Main T0 ");
            sb.Append(" LEFT JOIN Shipping_Item T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                WHERE SUBSTRING(T0.CARDCODE,1,1) in ('S','U') AND BOARDCOUNTNO='進口' ");
            sb.Append("                AND ISNULL(ADD9,'') <> '' AND T2.U_TMODEL like '%" + MODEL + "%' AND T2.U_VERSION=@VER ORDER BY ARRIVEDAY DESC  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
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


        private System.Data.DataTable MakeTableCombineFee()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("子公司", typeof(string));
            dt.Columns.Add("說明", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            dt.Columns.Add("貿易條件", typeof(string));
            dt.Columns.Add("運送方式", typeof(string));
            dt.Columns.Add("業務", typeof(string));

            //
            dt.Columns.Add("部門", typeof(string));
            dt.Columns.Add("種類", typeof(string));




            return dt;
        }

        private void button11_Click(object sender, EventArgs e)
        {

            //以明細第一筆作為分攤部門

            try
            {

                this.Cursor = Cursors.WaitCursor;

                //原客戶費用
                //費用
                //電子帳單

                //OPOR採購單

                string 單號;
                string 料號;
                string 單別;


                string SAP;
                System.Data.DataTable dt = GetData_Fee();
                if (dt.Rows.Count == 0)
                {

                    MessageBox.Show("沒有資料");
                    return;

                }
                System.Data.DataTable dtCost = MakeTableCombineFee();
                System.Data.DataTable dtDoc = null;
                string DuplicateKey = "";
                DataRow dr = null;

                string DocEntry = "";
                string ItemRemark = "";
                System.Data.DataTable dtSap = null;

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    單號 = dt.Rows[i]["JOBNO"].ToString();

                    dtDoc = GetData_ItemQty(單號);

                    try
                    {
                        DocEntry = Convert.ToString(dtDoc.Rows[0]["DocEntry"]);
                        ItemRemark = Convert.ToString(dtDoc.Rows[0]["ItemRemark"]);
                    }
                    catch
                    {

                    }

                    dr = dtCost.NewRow();
                    dr["JOBNO"] = Convert.ToString(dt.Rows[i]["JOBNO"]);
                    dr["客戶編號"] = Convert.ToString(dt.Rows[i]["客戶編號"]);
                    dr["客戶名稱"] = Convert.ToString(dt.Rows[i]["客戶"]);
                    dr["說明"] = Convert.ToString(dt.Rows[i]["說明"]);
                    dr["金額"] = Convert.ToString(dt.Rows[i]["金額"]);
                    dr["貿易條件"] = Convert.ToString(dt.Rows[i]["貿易條件"]);
                    dr["運送方式"] = Convert.ToString(dt.Rows[i]["運送方式"]);
                    dr["子公司"] = Convert.ToString(dt.Rows[i]["子公司"]);
                    dr["種類"] = "電子帳單";


                    if (ItemRemark == "")
                    {
                        dr["業務"] = "JOY";
                        dr["部門"] = "TFT事業部";
                    }
                    else if (ItemRemark.ToUpper() == "CHOICE")
                    {


                        dtSap = GetData_Doc2(DocEntry);
                        dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                        dr["部門"] = "TFT事業部";
                    }
                    else if (ItemRemark.ToUpper() == "INFINITE")
                    {


                        dtSap = GetData_Doc3(DocEntry);
                        dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                        dr["部門"] = "TFT事業部";
                    }
                    else
                    {
                        try
                        {
                            dtSap = GetData_Doc(ItemRemark, DocEntry);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message +ItemRemark+ DocEntry);
                        }
                        dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                        dr["部門"] = Convert.ToString(dtSap.Rows[0]["部門"]);


                    }




                    if (單號 != DuplicateKey)
                    {

                        dr["數量"] = Convert.ToString(dtDoc.Rows[0]["數量"]);

                    }
                    DuplicateKey = 單號;

                    dtCost.Rows.Add(dr);



                }





                dataGridView2.DataSource = dtCost;
                ExcelReport.GridViewToCSV(dataGridView2, Environment.CurrentDirectory + @"\SHIPPING" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }


        /// <summary>
        /// 電子帳單費用
        /// </summary>
        /// <returns></returns>
        private System.Data.DataTable GetData_Fee()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT T0.ShippingCode JOBNO,");
            sb.Append(" CASE T3.[CardCODE] WHEN '0511-00' THEN  T3.[ADD6] WHEN '0257-00' THEN T3.[ADD6] ELSE T3.[CardName] END 客戶,T3.[CardCODE] 客戶編號,");
            sb.Append(" T0.ITEM 說明, T0.Amount 金額,T3.boardcountno 貿易條件,T3.receiveDay 運送方式,SUBCOMPANY 子公司 ");
         
            sb.Append(" FROM Shipping_Fee T0");
            sb.Append(" LEFT JOIN acmesqlsp.dbo.shipping_main T3 on (T0.shippingcode=T3.shippingcode )");

            sb.Append(" where 1=1 ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T3.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox2.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and T3.shippingcode between '" + textBox2.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" ORDER BY T3.shippingcode");

            
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
        private System.Data.DataTable GetData_Fee2(string U_Shipping_no)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("                                             SELECT DISTINCT T0.U_Shipping_no  FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append("                                                  LEFT JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("                                                           INNER join acmesqlsp.dbo.shipping_main T3 on (t0.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS or t1.U_Shipping_no=T3.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                    WHERE ISNULL(T0.U_Shipping_no,'') <> '' AND    T3.CLOSEDAY between @startday and @endday AND T0.U_Shipping_no=@U_Shipping_no");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startday", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@endday", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@U_Shipping_no", U_Shipping_no));
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

        private System.Data.DataTable GetData_ItemQty(string shippingcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("select sum(QUANTITY) 數量 ,Max(DocEntry) DocEntry,Max(ItemRemark) ItemRemark from shipping_item  where shippingcode='" + shippingcode + "'  ");


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

        /// <summary>
        /// 銷售訂單->交貨->才有部門別
        /// 
        /// </summary>
        /// <returns></returns>
        private System.Data.DataTable GetData_Doc(string DocType,string DocNum)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            if (DocType == "AR貸項")
            {
                sb.Append(" SELECT CASE T1.GROUPCODE ");
                sb.Append(" WHEN 103 THEN 'TFT事業部' ");
                sb.Append(" WHEN 104 THEN 'LED事業部' ");
                sb.Append(" WHEN 105 THEN '太陽能事業部'  END  部門,T2.[SlpName] 業務");
                sb.Append("  FROM ORIN T0");
                sb.Append(" LEFT JOIN OCRD T1 ON (T0.CARDCODE=T1.CARDCODE)");
                sb.Append(" LEFT JOIN dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");
                sb.Append(" WHERE DocNum=@DocNum ");
            }

            if (DocType == "銷售訂單")
            {
                sb.Append(" SELECT CASE T1.GROUPCODE ");
                sb.Append(" WHEN 103 THEN 'TFT事業部' ");
                sb.Append(" WHEN 104 THEN 'LED事業部' ");
                sb.Append(" WHEN 105 THEN '太陽能事業部'  END  部門,T2.[SlpName] 業務");
                sb.Append("  FROM ORDR T0");
                sb.Append(" LEFT JOIN OCRD T1 ON (T0.CARDCODE=T1.CARDCODE)");
                sb.Append(" LEFT JOIN dbo.OSLP T2 ON T0.SlpCode = T2.SlpCode");
                sb.Append(" WHERE DocNum=@DocNum ");
            }
            
            if (DocType == "採購訂單")
            {
                sb.Append(" SELECT DISTINCT CASE ITMSGRPCOD ");
                sb.Append(" WHEN 1032 THEN 'TFT事業部' ");
                sb.Append(" WHEN 1034 THEN 'TFT事業部' ");
                sb.Append(" WHEN 1033 THEN 'LED事業部' ");
                sb.Append(" WHEN 233 THEN 'LED事業部' ");
                sb.Append(" WHEN 102 THEN '太陽能事業部'  END  部門,T3.[SlpName] 業務");
                sb.Append("   FROM OPOR T0");
                sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
                sb.Append(" LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
                sb.Append(" LEFT JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode");
                sb.Append(" WHERE DocNum=@DocNum ");
            }
            if (DocType == "調撥單")
            {
                sb.Append(" SELECT DISTINCT CASE ITMSGRPCOD ");
                sb.Append(" WHEN 1032 THEN 'TFT事業部' ");
                sb.Append(" WHEN 1033 THEN 'LED事業部' ");
                sb.Append(" WHEN 1034 THEN 'TFT事業部' ");
                sb.Append(" WHEN 233 THEN 'LED事業部' ");
                sb.Append(" WHEN 102 THEN '太陽能事業部'  END  部門,T3.[SlpName] 業務");
                sb.Append("   FROM OWTR T0");
                sb.Append(" LEFT JOIN WTR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
                sb.Append(" LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
                sb.Append(" LEFT JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode");
                sb.Append(" WHERE DocNum=@DocNum ");
            }
            if (DocType == "發貨單")
            {
                sb.Append(" SELECT DISTINCT CASE ITMSGRPCOD ");
                sb.Append(" WHEN 1032 THEN 'TFT事業部' ");
                sb.Append(" WHEN 1033 THEN 'LED事業部' ");
                sb.Append(" WHEN 1034 THEN 'TFT事業部' ");
                sb.Append(" WHEN 233 THEN 'LED事業部' ");
                sb.Append(" WHEN 102 THEN '太陽能事業部'  END  部門,T3.[SlpName] 業務");
                sb.Append("   FROM OIGE T0");
                sb.Append(" LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
                sb.Append(" LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
                sb.Append(" LEFT JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode");
                sb.Append(" WHERE DocNum=@DocNum ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocNum", SqlDbType.Int));
            command.Parameters["@DocNum"].Value = DocNum;

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
        private System.Data.DataTable GetData_Doc3(string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn22);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PersonName 業務 FROM OrdBillMain A Left Join comPerson T5 ON (A.SalesMan=T5.PersonID) where BillNO=@BillNO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


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
        private System.Data.DataTable GetData_Doc4(string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn20);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PersonName 業務 FROM OrdBillMain A Left Join comPerson T5 ON (A.SalesMan=T5.PersonID) where BillNO=@BillNO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


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
        private System.Data.DataTable GetData_Doc2(string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PersonName 業務 FROM OrdBillMain A Left Join comPerson T5 ON (A.SalesMan=T5.PersonID) where BillNO=@BillNO ");

      
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


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
        //sb.Append(" LEFT JOIN acmesqlsp.dbo.shipping_item T4 on (T3.shippingcode=T4.shippingcode COLLATE Chinese_Taiwan_Stroke_CI_AS )");

        private void button12_Click(object sender, EventArgs e)
        {
            //try
            //{
            DELETEFILE();
                this.Cursor = Cursors.WaitCursor;


                //原客戶費用
                //費用
                //電子帳單
                string 單號;
                string 料號;
                string SAP;

                //OPOR採購單
                System.Data.DataTable dt = GetOrderData19();



                System.Data.DataTable dtCost = MakeTableCombineFee();
                System.Data.DataTable dtDoc = null;
                string DuplicateKey = "";

                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    單號 = dt.Rows[i]["JOBNO"].ToString();
                    料號 = dt.Rows[i]["LINE"].ToString();
                    SAP = dt.Rows[i]["DOC"].ToString();
                    dtDoc = GetOrderData20(單號, 料號, SAP);

 

                    dr = dtCost.NewRow();
                    dr["JOBNO"] = Convert.ToString(dtDoc.Rows[0]["JOBNO"]);
                    dr["客戶編號"] = Convert.ToString(dtDoc.Rows[0]["客戶編號"]);
                    dr["客戶名稱"] = Convert.ToString(dtDoc.Rows[0]["客戶"]);
                    dr["說明"] = Convert.ToString(dtDoc.Rows[0]["說明"]);
                    dr["金額"] = Convert.ToString(dtDoc.Rows[0]["金額"]);
                    dr["貿易條件"] = Convert.ToString(dtDoc.Rows[0]["貿易條件"]);
                    dr["運送方式"] = Convert.ToString(dtDoc.Rows[0]["運送方式"]);
                    dr["業務"] = Convert.ToString(dtDoc.Rows[0]["業務"]);
                    dr["部門"] = Convert.ToString(dt.Rows[i]["部門"]);
                    
                    dr["種類"] = "費用";

                    if (單號 != DuplicateKey)
                    {

                        dr["數量"] = Convert.ToString(dtDoc.Rows[0]["數量"]);

                    }


        
                    DuplicateKey = 單號;

                    dtCost.Rows.Add(dr);

                }


                //加入電子帳單

                
              dt = GetData_Fee();
                
                DuplicateKey = "";
                

                string DocEntry = "";
                string ItemRemark = "";
                System.Data.DataTable dtSap = null;

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    單號 = dt.Rows[i]["JOBNO"].ToString();

                    dtDoc = GetData_ItemQty(單號);

                    try
                    {
                        DocEntry = Convert.ToString(dtDoc.Rows[0]["DocEntry"]);
                        ItemRemark = Convert.ToString(dtDoc.Rows[0]["ItemRemark"]);
                    }
                    catch
                    {

                    }
             
                    dr = dtCost.NewRow();
                    dr["JOBNO"] = Convert.ToString(dt.Rows[i]["JOBNO"]);
                    dr["客戶編號"] = Convert.ToString(dt.Rows[i]["客戶編號"]);
                    dr["客戶名稱"] = Convert.ToString(dt.Rows[i]["客戶"]);
                    dr["說明"] = Convert.ToString(dt.Rows[i]["說明"]);
                    dr["金額"] = Convert.ToString(dt.Rows[i]["金額"]);
                    dr["貿易條件"] = Convert.ToString(dt.Rows[i]["貿易條件"]);
                    dr["運送方式"] = Convert.ToString(dt.Rows[0]["運送方式"]);

                    dr["種類"] = "電子帳單";

                    if (ItemRemark == "" )
                    {
                        dr["業務"] = "JOY";
                        dr["部門"] = "TFT事業部";
                    }
                    else if (ItemRemark.ToUpper() == "CHOICE")
                    {


                        dtSap = GetData_Doc2(DocEntry);
                        if (dtSap.Rows.Count > 0)
                        {
                            dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                            dr["部門"] = "TFT事業部";
                        }
                    }
                    else if (ItemRemark.ToUpper() == "INFINITE")
                    {
                        dtSap = GetData_Doc3(DocEntry);
                        if (dtSap.Rows.Count > 0)
                        {
                            dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                            dr["部門"] = "TFT事業部";
                        }
                    }
                    else if (ItemRemark.ToUpper() == "TOP GARDEN")
                    {
                        dtSap = GetData_Doc4(DocEntry);
                        if (dtSap.Rows.Count > 0)
                        {
                            dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                            dr["部門"] = "TFT事業部";
                        }
                    }
                    else
                    {

                        dtSap = GetData_Doc(ItemRemark, DocEntry);
                        if (dtSap.Rows.Count > 0)
                        {
                            dr["業務"] = Convert.ToString(dtSap.Rows[0]["業務"]);
                            dr["部門"] = Convert.ToString(dtSap.Rows[0]["部門"]);
                        }


                    }


                    if (單號 != DuplicateKey)
                    {

                        dr["數量"] = Convert.ToString(dtDoc.Rows[0]["數量"]);

                    }

                    System.Data.DataTable dtSap2 = GetData_Fee2(單號);

                    if (dtSap2.Rows.Count > 0)
                    {
                        dr["數量"] = "";
                    }
                    DuplicateKey = 單號;

                    dtCost.Rows.Add(dr);

                   

                }


                dataGridView2.DataSource = dtCost;


                ExcelReport.GridViewToCSV(dataGridView2, Environment.CurrentDirectory + @"\SHIPPING" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
      
           
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir;
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {
                    int  a = file.IndexOf("csv");

                    if (a != -1)
                    {
                        File.Delete(file);
                    }

                }
            }
            catch { }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton5.Checked)
                {
                    ETSAI();
                    GridViewToExcel(dataGridView2);
                }
                else if (radioButton6.Checked)
                {
                    ETSAI2();
                    GridViewToExcel(dataGridView2);
                }
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            System.Data.DataTable F1 = null;
            if (comboBox16.Text == "A. TW出口運輸費用 (友福)")
            {
                F1 = GETFEEA();

            }
            if (comboBox16.Text == "B. TW出口報關費用 (AIR)")
            {
                F1 = GETFEEB();

            }
            if (comboBox16.Text == "B. TW出口報關費用 (SEA)")
            {
                F1 = GETFEEB2();

            }
            if (comboBox16.Text == "C. TW出口 local 費用 (AIR)")
            {
                F1 = GETFEEC1();
            }
            if (comboBox16.Text == "C. TW出口 local 費用 (SEA)")
            {
                F1 = GETFEEC2();
            }
            if (comboBox16.Text == "D. DHL出口運費")
            {
                F1 = GETFEED();
            }
            dataGridView2.DataSource = F1;
            GridViewToExcel(dataGridView2);
//            A. TW出口運輸費用 (友福)
//B. TW出口報關費用 (AIR)
//B. TW出口報關費用 (SEA)
//C. TW出口 local 費用 (AIR)
//C. TW出口 local 費用 (SEA)
//D. DHL出口運費
        }
        private System.Data.DataTable GETFEEA()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,T1.AT 噸數,T1.aA 起點,T1.aD 迄點");
            sb.Append(" ,aGA 加點費,aSHA 加卸費,aE 議價,aBIN 併單,aAMT 總金額");
            sb.Append("  FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode)");
            sb.Append(" WHERE ISNULL(T1.aAMT,'') <> ''");

            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }
         

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
        private System.Data.DataTable GETFEEB()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,bA 報關行,bT 推貿費,bTM 原因,bAE 議價");
            sb.Append(" ,bAAMT 總金額 ");
            sb.Append(" FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode) ");
            sb.Append(" WHERE ISNULL(T1.bAAMT,'') <> '' ");

            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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

        private System.Data.DataTable GETFEEB2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,bS 報關行,bSG 櫃,cT 推貿費,cTM 原因,bSE 議價");
            sb.Append(" ,cSG '機械使用費(港區費用)',bSAMT 總金額 ");
            sb.Append(" FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode) ");
            sb.Append(" WHERE ISNULL(T1.bSAMT,'') <> '' ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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

        private System.Data.DataTable GETFEEC1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,cAK 計費重量,cAT 提單文件費");
            sb.Append(" ,cAZ 倉租費,cAS '艙單申報(ENS/AMS/AFR)',cAAMT 總金額 ");
            sb.Append(" FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode) ");
            sb.Append(" WHERE ISNULL(T1.cAAMT,'') <> '' ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
        private System.Data.DataTable GETFEEC2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,cSC 計費材積,cSV 'VGM(其他)'");
            sb.Append(" ,cST 提單文件費,cSD 提單電放費,cSB 併櫃費,cSH 操作手續費,cSS '艙單申報(ENS/AMS/AFR)',cSAMT 總金額 ");
            sb.Append(" FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode) ");
            sb.Append(" WHERE ISNULL(T1.cSAMT,'') <> '' ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
        private System.Data.DataTable GETFEED()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CloseDay,T0.ShippingCode,receivePlace 收貨地,goalPlace 目的地,dHL1 計費重量,dHL2 總金額 ");
            sb.Append(" FROM SHIPPING_MAIN T0 LEFT JOIN Ship_Fee  T1 ON (T0.ShippingCode =T1.ShippingCode) ");
            sb.Append(" WHERE ISNULL(T1.dHL2,'') <> '' ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and  T0.CLOSEDAY  between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
    }
}
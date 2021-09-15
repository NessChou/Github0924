using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{

    public partial class GBCHOICE : Form
    {

        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn3 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBCHOICE()
        {
            InitializeComponent();
        }



        private void button6_Click(object sender, EventArgs e)
        {

            if (comboBox8.Text !="")
            {
                if(comboBox8.Text=="銷貨總表")
                {
                    System.Data.DataTable L1 = GetCHO3ANDNOTCLOSE();
                    dataGridView7.DataSource = L1;
                    if (comboBox6.Text == "聿豐")
                    {
                        L1.DefaultView.RowFilter = " 公司='聿豐' ";
                    }
                    else if (comboBox6.Text == "忠孝")
                    {
                        L1.DefaultView.RowFilter = " 公司='忠孝' ";
                    }
                    for (int i = 27; i <= 32; i++)
                    {
                        DataGridViewColumn col = dataGridView7.Columns[i];


                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        col.DefaultCellStyle.Format = "#,##0.0000";


                    }
                }

                if (comboBox8.Text == "費用單")
                {
                    System.Data.DataTable L1 = GetCHO6();
                    dataGridView6.DataSource = L1;
                    if (comboBox6.Text == "聿豐")
                    {
                        L1.DefaultView.RowFilter = " 公司='聿豐' ";
                    }
                    else if (comboBox6.Text == "忠孝")
                    {
                        L1.DefaultView.RowFilter = " 公司='忠孝' ";
                    }
                    for (int i = 7; i <= 7; i++)
                    {
                        DataGridViewColumn col = dataGridView6.Columns[i];
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        col.DefaultCellStyle.Format = "#,##0.00";
                    }
                }

                if (comboBox8.Text == "採購已結")
                {
                    System.Data.DataTable L1 = GetCHO4();
                    dataGridView2.DataSource = L1;
                    if (comboBox6.Text == "聿豐")
                    {
                        L1.DefaultView.RowFilter = " 公司='聿豐' ";
                    }
                    else if (comboBox6.Text == "忠孝")
                    {
                        L1.DefaultView.RowFilter = " 公司='忠孝' ";
                    }
                    for (int i = 20; i <= 25; i++)
                    {
                        DataGridViewColumn col2 = dataGridView2.Columns[i];

                        col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        col2.DefaultCellStyle.Format = "#,##0.00";
                    }
                }
                if (comboBox8.Text == "進退/進折")
                {
                    System.Data.DataTable L1 = GetCH11("1");
                    dataGridView11.DataSource = L1;
                    if (comboBox6.Text == "聿豐")
                    {
                        L1.DefaultView.RowFilter = " 公司='聿豐' ";
                    }
                    else if (comboBox6.Text == "忠孝")
                    {
                        L1.DefaultView.RowFilter = " 公司='忠孝' ";
                    }
                    for (int i = 20; i <= 25; i++)
                    {
                        DataGridViewColumn col2 = dataGridView11.Columns[i];

                        col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        col2.DefaultCellStyle.Format = "#,##0.00";
                    }
                }
                //進退/進折
            }
            else
            {

                System.Data.DataTable dtCost = MakeTableCombine();
                DataRow dr = null;
                System.Data.DataTable DT1 = GetCHO3();
                for (int i = 0; i <= DT1.Rows.Count - 1; i++)
                {
                    DataRow dd = DT1.Rows[i];
                    dr = dtCost.NewRow();
                    dr["年"] = dd["年"].ToString();
                    dr["月"] = dd["月"].ToString();
                    string COMPANY = dd["公司"].ToString();
                    dr["公司"] = COMPANY;
                    dr["客戶類別"] = dd["客戶類別"].ToString();
                    dr["類別名稱"] = dd["類別名稱"].ToString();
                    dr["來源"] = dd["來源"].ToString();
                    dr["門市"] = dd["門市"].ToString();
                    dr["B"] = dd["B"].ToString();
                    dr["品項"] = dd["品項"].ToString();
                    dr["類別"] = dd["類別"].ToString();
                    dr["零售"] = dd["零售"].ToString();
                    dr["客戶代碼"] = dd["客戶代碼"].ToString();
                    dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                    dr["取貨日期"] = dd["取貨日期"].ToString();
                    dr["訂購憑單日期"] = dd["訂購憑單日期"].ToString();
                    dr["訂購憑單週數"] = dd["訂購憑單週數"].ToString();
                    dr["訂購憑單號碼"] = dd["訂購憑單號碼"].ToString();
                    dr["銷貨單據日期"] = dd["銷貨單據日期"].ToString();
                    dr["銷貨單據週數"] = dd["銷貨單據週數"].ToString();
                    string BILLNO = dd["銷貨單據號碼"].ToString();
                    string INVOICENO = dd["INVOICENO"].ToString();
                    dr["銷貨單據號碼"] = BILLNO;

                    if (!String.IsNullOrEmpty(INVOICENO))
                    {
                        System.Data.DataTable DTINV = null;
                        if (COMPANY == "聿豐")
                        {
                            DTINV = GetINVOICE(INVOICENO);
                        }
                        if (COMPANY == "忠孝")
                        {
                            BILLNO = BILLNO.Replace("T", "");
                            DTINV = GetINVOICE2(INVOICENO);
                        }
                        if (DTINV.Rows.Count > 0)
                        {
                            StringBuilder sb2 = new StringBuilder();
                            StringBuilder sb3 = new StringBuilder();
                            for (int s = 0; s <= DTINV.Rows.Count - 1; s++)
                            {
                                string 發票日期 = DTINV.Rows[s]["發票日期"].ToString();
                                string 發票號碼 = DTINV.Rows[s]["發票號碼"].ToString();

                                sb2.Append(發票日期 + "/");
                                sb3.Append(發票號碼 + "/");

                            }
                            sb2.Remove(sb2.Length - 1, 1);
                            sb3.Remove(sb3.Length - 1, 1);
                            dr["發票日期"] = sb2.ToString();
                            dr["發票號碼"] = sb3.ToString();
                        }
                    }


                    string PRODID = dd["產品編號"].ToString();
                    dr["產品編號"] = PRODID;
                    dr["品名規格"] = dd["品名規格"].ToString();
                    dr["發票品名"] = dd["發票品名"].ToString();
                    dr["倉別"] = dd["倉別"].ToString();
                    dr["數量"] = dd["數量"].ToString();
                    dr["單位"] = dd["單位"].ToString();
                    dr["單價"] = dd["單價"].ToString();
                    dr["金額"] = dd["金額"].ToString();
                    dr["成本"] = dd["成本"].ToString();
                    dr["毛利"] = dd["毛利"].ToString();
                    dr["稅"] = dd["稅"].ToString();
                    dr["金額含稅"] = dd["金額含稅"].ToString();
                    dr["分錄備註"] = dd["分錄備註"].ToString();
                    string G = dd["快遞單號"].ToString().Trim();
                    dr["快遞單號"] = G;

                    if (!String.IsNullOrEmpty(G))
                    {
                        System.Data.DataTable A1 = GTODOWN(G);
                        if (A1.Rows.Count > 0)
                        {
                            dr["下載"] = "下載";
                        }
                    }
                    dr["是否為贈品"] = dd["是否為贈品"].ToString();
                    dr["細項描述"] = dd["細項描述"].ToString();
                    dr["筆數"] = dd["筆數"].ToString();
                    dr["收款方式"] = dd["收款方式"].ToString();
                    dr["天"] = dd["天"].ToString();
                    dr["部門"] = dd["部門"].ToString();
                    dtCost.Rows.Add(dr);
                }

               //銷貨已結
                System.Data.DataTable L1 = dtCost;
                //採購已結
                System.Data.DataTable L2 = GetCHO4();
                //銷售未結
                System.Data.DataTable L3 = GetCHO3NOTCLOSE();
                //採購未結
                System.Data.DataTable L4 = GetCHO4NOTCLOSE();
                //調整
                System.Data.DataTable L5 = GetCHO5();
                //費用單
                System.Data.DataTable L6 = GetCHO6();
                //銷貨總表
                System.Data.DataTable L7 = GetCHO3ANDNOTCLOSE();
                //銷退/銷折
                System.Data.DataTable L9 = GetCHO9();
                //進退/進折
                System.Data.DataTable L11 = GetCH11("1");
                //承銷品銷貨總表
                System.Data.DataTable L12 = GetCH11("2");
                //承銷品銷貨總表
                System.Data.DataTable L10 = GetCHO3ANDNOTCLOSE2F();
                if (comboBox6.Text == "聿豐")
                {
                    L1.DefaultView.RowFilter = " 公司='聿豐' ";
                    L2.DefaultView.RowFilter = " 公司='聿豐' ";
                    L3.DefaultView.RowFilter = " 公司='聿豐' ";
                    L4.DefaultView.RowFilter = " 公司='聿豐' ";
                    L5.DefaultView.RowFilter = " 公司='聿豐' ";
                    L6.DefaultView.RowFilter = " 公司='聿豐' ";
                    L7.DefaultView.RowFilter = " 公司='聿豐' ";
                    L9.DefaultView.RowFilter = " 公司='聿豐' ";
                    L10.DefaultView.RowFilter = " 公司='聿豐' ";
                    L11.DefaultView.RowFilter = " 公司='聿豐' ";
                }
                else if (comboBox6.Text == "忠孝")
                {
                    L1.DefaultView.RowFilter = " 公司='忠孝' ";
                    L2.DefaultView.RowFilter = " 公司='忠孝' ";
                    L3.DefaultView.RowFilter = " 公司='忠孝' ";
                    L4.DefaultView.RowFilter = " 公司='忠孝' ";
                    L5.DefaultView.RowFilter = " 公司='忠孝' ";
                    L6.DefaultView.RowFilter = " 公司='忠孝' ";
                    L7.DefaultView.RowFilter = " 公司='忠孝' ";
                    L9.DefaultView.RowFilter = " 公司='忠孝' ";
                    L10.DefaultView.RowFilter = " 公司='忠孝' ";
                    L11.DefaultView.RowFilter = " 公司='忠孝' ";
                }

                dataGridView1.DataSource = L1;

                //採購已結
                dataGridView2.DataSource = L2;
                //銷售未結
                dataGridView3.DataSource = L3;
                //採購未結
                dataGridView4.DataSource = L4;
                //調整
                dataGridView5.DataSource = L5;
                //費用單
                dataGridView6.DataSource = L6;
                //銷貨總表
                dataGridView7.DataSource = L7;
                //庫存總表
                if (comboBox6.Text == "忠孝")
                {
                    dataGridView8.DataSource = GetCHO82();
                }
                else
                {
                    dataGridView8.DataSource = GetCHO8();
                }
                //銷退/銷折
                dataGridView9.DataSource = L9;

                //進退/進折
                dataGridView11.DataSource = L11;

                //承銷品進貨總表
                dataGridView12.DataSource = L12;

                //承銷品銷貨總表
                dataGridView10.DataSource = L10;
                for (int i = 26; i <= 33; i++)
                {
                    DataGridViewColumn col = dataGridView1.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.0000";


                }

                for (int i = 23; i <= 28; i++)
                {
                    DataGridViewColumn col = dataGridView9.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.00";


                }
                for (int i = 27; i <= 32; i++)
                {
                    DataGridViewColumn col = dataGridView7.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.0000";


                }
                for (int i = 20; i <= 25; i++)
                {


                    DataGridViewColumn col2 = dataGridView2.Columns[i];


                    col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col2.DefaultCellStyle.Format = "#,##0.00";
                }
                for (int i = 20; i <= 25; i++)
                {


                    DataGridViewColumn col2 = dataGridView11.Columns[i];


                    col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col2.DefaultCellStyle.Format = "#,##0.00";
                }

                for (int i = 21; i <= 27; i++)
                {


                    DataGridViewColumn col2 = dataGridView12.Columns[i];


                    col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col2.DefaultCellStyle.Format = "#,##0.00";
                }


                for (int i = 22; i <= 26; i++)
                {
                    DataGridViewColumn col = dataGridView3.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.0000";


                }

                for (int i = 13; i <= 17; i++)
                {
                    DataGridViewColumn col2 = dataGridView4.Columns[i];
                    col2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col2.DefaultCellStyle.Format = "#,##0.00";
                }

                for (int i = 12; i <= 15; i++)
                {
                    DataGridViewColumn col = dataGridView5.Columns[i];
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0.00";
                }

                for (int i = 7; i <= 7; i++)
                {
                    DataGridViewColumn col = dataGridView6.Columns[i];
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0.00";
                }
                if (comboBox6.Text == "忠孝")
                {
                    for (int i = 5; i <= 7; i++)
                    {
                        DataGridViewColumn col = dataGridView8.Columns[i];
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        col.DefaultCellStyle.Format = "#,##0.00";
                    }
                }
                else
                {
                    for (int i = 6; i <= 8; i++)
                    {
                        DataGridViewColumn col = dataGridView8.Columns[i];
                        col.DefaultCellStyle.ForeColor = Color.Blue;
                    }
                    for (int i = 5; i <= 33; i++)
                    {
                        DataGridViewColumn col = dataGridView8.Columns[i];
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        col.DefaultCellStyle.Format = "#,##0.00";
                    }
                }

      

                for (int i = 26; i <= 30; i++)
                {
                    DataGridViewColumn col = dataGridView10.Columns[i];
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0.00";
                }
                for (int i = 31; i <= 32; i++)
                {
                    DataGridViewColumn col = dataGridView10.Columns[i];
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0";
                }
            }
                if (globals.GroupID.ToString().Trim() == "GBSALES")
                {
                    tabControl1.TabPages.Remove(tabControl1.TabPages["採購單已結報表"]);
                    tabControl1.TabPages.Remove(tabControl1.TabPages["採購單未結報表"]);
                    tabControl1.TabPages.Remove(tabControl1.TabPages["調整憑單報表"]);
                    tabControl1.TabPages.Remove(tabControl1.TabPages["庫存總表"]);
                    tabControl1.TabPages.Remove(tabControl1.TabPages["銷折銷退"]);
                }
            
        }

        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,          ");
            sb.Append(" '聿豐' 公司,L.ClassID 客戶類別,L.ClassName 類別名稱,      ");
            sb.Append(" case when T.CustID = 'tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,         ");
            sb.Append(" case when T.CustID = 'tw90146-16'   THEN  O2.LinkMan      ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN      ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan      ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')      ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN      ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan      ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')     ");
            sb.Append(" END End 門市,         ");
            sb.Append("  L.ENGNAME B,CASE WHEN SUBSTRING(A.ProdID ,1,2)='A1' THEN '外購品' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'FRE' THEN '運費' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MCK' THEN '雞' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MPK' THEN '豬' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSR' THEN '蝦' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSF' THEN '魚' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PCK' THEN '加工品-雞' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PFH' THEN '加工品-魚' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PPK' THEN '加工品-豬'  ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'           ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'           ");
            sb.Append(" END 零售,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.UserDef1 取貨日期,    ");
            sb.Append(" O2.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 訂購憑單週數,A.FromNO 訂購憑單號碼,    ");
            sb.Append(" A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(A.BillDate as varchar) as datetime)) 銷貨單據週數,    ");
            sb.Append(" A.BillNO 銷貨單據號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,CAST(ROUND(A.MLAmount,0) AS INT) 金額,CAST(ROUND(A.TaxAmt,0) AS INT) 稅            ");
            sb.Append(" ,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) 金額含稅,CostForAcc 成本 ,A.MLAmount-A.CostForAcc 毛利,O.ItemRemark 分錄備註,T.UDef2 快遞單號,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述              ");
            sb.Append(" ,  CASE RANK() OVER( PARTITION BY O.BILLNO   ORDER BY O.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數 ,  ");
            sb.Append(" CASE F2.GatherStyle  WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN  F2.GatherOther END 收款方式,F2.GatherDelay  天,INVOICENO,T.DeptID 部門   ");
            sb.Append(" From CHICOMP02.DBO.ComProdRec A                        ");
            sb.Append(" Left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag            ");
            sb.Append(" Left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1            ");
            sb.Append(" Left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID            ");
            sb.Append(" Left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2             ");
            sb.Append(" Left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2                                                                                                                                      ");
            sb.Append(" Left join CHICOMP02.DBO.comPerson P ON (T.Salesman=P.PersonID)            ");
            sb.Append(" Left join CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )            ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID            ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID            ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1             ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1        ");
            sb.Append(" Left Join CHICOMP02.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO       ");
            sb.Append(" Left join CHICOMP02.DBO.stkBillMain F2 On  F.BillNO=F2.BillNO and F.Flag =F2.Flag            ");
            sb.Append(" Where A.Flag=500 AND ISNULL(O2.BillStatus,0) <> 2   ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }


            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {

                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {

                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DeptID BETWEEN @C1 AND @C2 ");
            }
            sb.Append(" UNION ALL  ");
            sb.Append(" Select year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,           ");
            sb.Append(" '忠孝' 公司,L.ClassID 客戶類別,L.ClassName 類別名稱,M.ADDFIELD1 來源,'' 門市,          ");
            sb.Append(" L.ENGNAME B, CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'   WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'         ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'            ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'             ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'            ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'            ");
            sb.Append(" END 零售,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,''取貨日期,     ");
            sb.Append(" '' 訂購憑單日期,'' 訂購憑單週數,A.FromNO 訂購憑單號碼,     ");
            sb.Append(" A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(A.BillDate as varchar) as datetime)) 銷貨單據週數,     ");
            sb.Append(" A.BillNO 銷貨單據號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,CAST(ROUND(A.MLAmount,0) AS INT) 金額,CAST(ROUND(A.TaxAmt,0) AS INT) 稅             ");
            sb.Append(" ,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) 金額含稅,CostForAcc 成本 ,A.MLAmount-A.CostForAcc 毛利,A.ItemRemark 分錄備註,T.UDef2 快遞單號,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述               ");
            sb.Append(" ,  CASE RANK() OVER( PARTITION BY A.BILLNO   ORDER BY A.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數 ,   ");
            sb.Append(" U3.FullName  收款方式,U2.DistDays 天,INVOICENO,T.DeptID 部門    ");
            sb.Append(" From CHICOMP03.DBO.ComProdRec A                         ");
            sb.Append(" Left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag             ");
            sb.Append(" Left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID                                                                                                                                                ");
            sb.Append(" Left join CHICOMP03.DBO.comPerson P ON (T.Salesman=P.PersonID)             ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID             ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID             ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1              ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1         ");
            sb.Append(" Left Join CHICOMP03.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO        ");
            sb.Append(" Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1        ");
            sb.Append(" Left join CHICOMP03.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1       ");
            sb.Append(" Where A.Flag=500 AND K.ClassName <> '承銷品' ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }


            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {

                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {

                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DeptID BETWEEN @C1 AND @C2 ");
            }
               sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC2lassName", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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

        public System.Data.DataTable GetCHO9()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select CASE A.Flag WHEN 600 THEN '銷退' WHEN 701 THEN '銷折'  END 類別,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,       ");
            sb.Append(" '聿豐' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,   ");
            sb.Append(" L.ENGNAME B, CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'         ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'        ");
            sb.Append(" END 零售,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,A.FromNO 銷貨憑單號碼,A.BillDate  銷折單據日期,A.BillNO 銷折單據號碼,         ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,");
            sb.Append(" (CASE WHEN A.Flag=600 THEN  A.Quantity ELSE 0 END)*-1   數量,J.Unit 單位,A.Price 單價,");
            sb.Append(" CASE A.Flag WHEN 600 THEN A.Amount*-1  ELSE  F.Dist*-1 END 金額,");
            sb.Append(" CASE A.Flag WHEN 600 THEN A.TaxAmt*-1 ELSE  F.DistTaxAmt*-1  END 稅,");
            sb.Append(" CASE A.Flag WHEN 600 THEN (A.Amount+A.TaxAmt)*-1 ELSE (F.Dist+F.DistTaxAmt)*-1  END 金額含稅");
            sb.Append(" ,A.CostForAcc*-1 成本 ,");
            sb.Append(" CASE A.Flag WHEN 600 THEN  (A.Amount-A.CostForAcc)*-1 ELSE (F.Dist-A.CostForAcc)*-1  END  毛利,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式                                            ");
            sb.Append(" From CHICOMP02.DBO.ComProdRec A           ");
            sb.Append(" Left Join CHICOMP02.DBO.comWareHouse D On D.WareHouseID=A.WareID         ");
            sb.Append(" left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag         ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID                                                                                                                                     ");
            sb.Append(" left join CHICOMP02.DBO.comPerson P ON (T.Salesman=P.PersonID)             ");
            sb.Append(" Left Join CHICOMP02.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =4 AND I.IsCancel <> 1 AND T.Flag= I.SrcBillFlag           ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID         ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1          ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append(" Left Join CHICOMP02.DBO.stkDistSub F On A.BillNO =F.DISTNO and A.Flag =F.Flag AND A.RowNO =F.RowNO    ");
            sb.Append(" Left join CHICOMP02.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1      ");
            sb.Append(" Where A.Flag IN  (600,701) ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {

                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            sb.Append("  UNION ALL ");
            sb.Append(" Select CASE A.Flag WHEN 600 THEN '銷退' WHEN 701 THEN '銷折'  END 類別,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,       ");
            sb.Append(" '忠孝' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,   ");
            sb.Append(" L.ENGNAME B, CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'         ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'        ");
            sb.Append(" END 零售,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,A.FromNO 銷貨憑單號碼,A.BillDate  銷折單據日期,A.BillNO 銷折單據號碼,         ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,");
            sb.Append(" (CASE WHEN A.Flag=600 THEN  A.Quantity ELSE 0 END)*-1   數量,J.Unit 單位,A.Price 單價,");
            sb.Append(" CASE A.Flag WHEN 600 THEN A.Amount*-1  ELSE  F.Dist*-1 END 金額,");
            sb.Append(" CASE A.Flag WHEN 600 THEN A.TaxAmt*-1 ELSE  F.DistTaxAmt*-1  END 稅,");
            sb.Append(" CASE A.Flag WHEN 600 THEN (A.Amount+A.TaxAmt)*-1 ELSE (F.Dist+F.DistTaxAmt)*-1  END 金額含稅");
            sb.Append(" ,A.CostForAcc*-1 成本 ,");
            sb.Append(" CASE A.Flag WHEN 600 THEN  (A.Amount-A.CostForAcc)*-1 ELSE (F.Dist-A.CostForAcc)*-1  END  毛利,U3.FullName 收款方式                                                 ");
            sb.Append(" From CHICOMP03.DBO.ComProdRec A           ");
            sb.Append(" Left Join CHICOMP03.DBO.comWareHouse D On D.WareHouseID=A.WareID         ");
            sb.Append(" left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag         ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID                                                                                                                                     ");
            sb.Append(" left join CHICOMP03.DBO.comPerson P ON (T.Salesman=P.PersonID)             ");
            sb.Append(" Left Join CHICOMP03.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =4 AND I.IsCancel <> 1 AND T.Flag= I.SrcBillFlag           ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID         ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1          ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append(" Left Join CHICOMP03.DBO.stkDistSub F On A.BillNO =F.DISTNO and A.Flag =F.Flag AND A.RowNO =F.RowNO    ");
            sb.Append(" Left join CHICOMP03.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1       ");
            sb.Append(" Where A.Flag IN  (600,701) ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {

                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC2lassName", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
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

        public System.Data.DataTable GetINVOICE(string InvoiceNO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT InvoiceDate 發票日期,InvoiceNO 發票號碼  FROM CHICOMP02.DBO.comInvoice WHERE InvoiceNO =@InvoiceNO AND Flag =2 AND IsCancel <> 1      ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNO", InvoiceNO));
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
        public System.Data.DataTable GetINVOICE2(string InvoiceNO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT InvoiceDate 發票日期,InvoiceNO 發票號碼  FROM CHICOMP03.DBO.comInvoice WHERE InvoiceNO =@InvoiceNO AND Flag =2 AND IsCancel <> 1      ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNO", InvoiceNO));
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
        public System.Data.DataTable GetCHO3ANDNOTCLOSE()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select CASE WHEN O2.BillStatus=2 THEN '無效' ELSE  '已結' END 單況,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,      ");
            sb.Append(" '聿豐' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,   ");
            sb.Append(" case when T.CustID ='tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,     ");
            sb.Append(" case when T.CustID = 'tw90146-16'   THEN  O2.LinkMan    ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan    ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan    ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append(" END End 門市, L.ENGNAME 'B/C',       ");
            sb.Append(" CASE WHEN SUBSTRING(A.ProdID ,1,2)='A1' THEN '外購品' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'FRE' THEN '運費' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MCK' THEN '雞' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MPK' THEN '豬' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSR' THEN '蝦' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSF' THEN '魚' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PCK' THEN '加工品-雞' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PFH' THEN '加工品-魚' ");
            sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PPK' THEN '加工品-豬'  ");
            sb.Append(" END  品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'       ");
            sb.Append(" END '零售/批發',''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.UserDef1 取貨日期,O2.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 訂購憑單週數,A.FromNO 訂購憑單號碼  ");
            sb.Append(" ,A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 銷貨單據週數,A.BillNO 銷貨單據號碼,        ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,CAST(ROUND(A.MLAmount,0) AS INT) '金額(未稅)',A.TaxAmt 稅        ");
            sb.Append(" ,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) '金額(含稅)',O.ItemRemark 分錄備註,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述,P.PersonName 業務,O2.ProjectID 專案         ");
            sb.Append(" , CASE RANK() OVER( PARTITION BY O.BILLNO   ORDER BY O.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數            ");
            sb.Append(" ,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天         From CHICOMP02.DBO.ComProdRec A     ");
            sb.Append(" Left Join CHICOMP02.DBO.comWareHouse D On D.WareHouseID=A.WareID         ");
            sb.Append(" left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag         ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID         ");
            sb.Append(" left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2          ");
            sb.Append(" left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2          ");
            sb.Append(" left join CHICOMP02.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID          ");
            sb.Append(" left join CHICOMP02.DBO.COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)         ");
            sb.Append(" left join CHICOMP02.DBO.comPerson P ON (S.Salesman=P.PersonID)         ");
            sb.Append(" LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )         ");
            sb.Append(" Left Join CHICOMP02.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1 AND T.InvoiceNo =I.InvoiceNO               ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID         ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1          ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append(" Left Join CHICOMP02.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO    ");
            sb.Append(" Left join CHICOMP02.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1  ");
            sb.Append(" Where A.Flag=500   ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {
                    if (comboBox3.Text == "官網")
                    {
                        sb.Append("  AND (K.ClassName  = @ClassName OR LEN(O2.CustBillNo) = 17 ) ");
                    }
                    else
                    {
                        sb.Append("  AND L.ClassName  = @CClassName ");
                    }
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }


            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
            sb.Append("                                        UNION ALL");
            sb.Append(" select CASE WHEN T0.BillStatus=2 THEN '無效' ELSE  '未結' END 單況,year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,      ");
            sb.Append(" '聿豐' 公司,T0.DepartID  部門,L.ClassID 客戶類別,L.ClassName 類別名稱,     case when T0.CustomerID = 'tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,    ");
            sb.Append(" case when T0.CustomerID = 'tw90146-16'   THEN  T0.LinkMan    ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan    ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan    ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append(" END End 門市, L.ENGNAME 'B/C',       ");
            sb.Append(" CASE WHEN SUBSTRING(T1.ProdID ,1,2)='A1' THEN '外購品' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'FRE' THEN '運費' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MCK' THEN '雞' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MPK' THEN '豬' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MSR' THEN '蝦' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MSF' THEN '魚' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PCK' THEN '加工品-雞' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PFH' THEN '加工品-魚' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PPK' THEN '加工品-豬'  ");
            sb.Append(" END 品項,K.ClassName  類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'        ");
            sb.Append(" END '零售/批發',        ");
            sb.Append(" ''''+T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.UserDef1 取貨日期,T0.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(T0.BillDate as varchar) as datetime)) 訂購憑單週數,T0.BillNO 訂購憑單號碼,''  銷貨單據日期,''  銷貨單據週數,'' 銷貨單據號碼,'','', T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名,''       ");
            sb.Append(" ,T1.Quantity 數量,J.Unit 單位,CAST(T1.Price AS decimal(18,4)) 單價,CAST(ROUND(T1.Amount,0) AS INT) '金額(未稅)',T1.TaxAmt 稅         ");
            sb.Append(" ,CAST(ROUND(T1.Amount+T1.TaxAmt,0) AS INT) '金額(含稅)',T1.ItemRemark 分錄備註,CASE T1.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,T1.Detail 細項描述,P.PersonName 業務,T0.ProjectID 專案,  CASE RANK() OVER( PARTITION BY T1.BILLNO   ORDER BY T1.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數               ");
            sb.Append(" ,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天        from CHICOMP02.DBO.ordBillMain T0         ");
            sb.Append(" left join CHICOMP02.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID          ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On L.ClassID =U.ClassID and L.Flag =1             ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1      ");
            sb.Append(" LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (T0.AddressID=AD.AddrID AND T0.CustomerID=AD.ID )   left join CHICOMP02.DBO.comPerson P ON (T0.Salesman=P.PersonID)      ");
            sb.Append(" Left join CHICOMP02.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1  ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus IN (0,2)  and QtyRemain  > 0    ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {
                    if (comboBox3.Text == "官網")
                    {
                        sb.Append("  AND (K.ClassName  = @ClassName OR LEN(T0.CustBillNo) = 17 ) ");
                    }
                    else
                    {
                        sb.Append("  AND L.ClassName  = @CClassName ");
                    }
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }
            sb.Append(" UNION ALL ");
            sb.Append(" Select CASE WHEN O2.BillStatus=2 THEN '無效' ELSE  '已結' END 單況,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,     ");
            sb.Append(" '忠孝' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,  ");
            sb.Append(" case when T.CustID ='tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,    ");
            sb.Append(" case when T.CustID = 'tw90146-16'   THEN  O2.LinkMan   ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN   ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan   ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN   ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan   ");
            sb.Append(" ELSE REPLACE(O2.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')  ");
            sb.Append(" END End 門市, L.ENGNAME 'B/C' ,     ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'    ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'          ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'      ");
            sb.Append(" END '零售/批發',''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.UserDef1 取貨日期,O2.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 訂購憑單週數,A.FromNO 訂購憑單號碼 ");
            sb.Append(" ,A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 銷貨單據週數,A.BillNO 銷貨單據號碼,       ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,CAST(ROUND(A.MLAmount,0) AS INT) '金額(未稅)',A.TaxAmt 稅       ");
            sb.Append(" ,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) '金額(含稅)',O.ItemRemark 分錄備註,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述,P.PersonName 業務,O2.ProjectID 專案        ");
            sb.Append(" , CASE RANK() OVER( PARTITION BY O.BILLNO   ORDER BY O.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數           ");
            sb.Append(" ,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天         From CHICOMP03.DBO.ComProdRec A    ");
            sb.Append(" Left Join CHICOMP03.DBO.comWareHouse D On D.WareHouseID=A.WareID        ");
            sb.Append(" left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag        ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1        ");
            sb.Append(" left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID        ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2         ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2         ");
            sb.Append(" left join CHICOMP03.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID         ");
            sb.Append(" left join CHICOMP03.DBO.COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)        ");
            sb.Append(" left join CHICOMP03.DBO.comPerson P ON (S.Salesman=P.PersonID)        ");
            sb.Append(" LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )        ");
            sb.Append(" Left Join CHICOMP03.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1  AND T.InvoiceNo =I.InvoiceNO       ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID        ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID        ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1         ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1    ");
            sb.Append(" Left Join CHICOMP03.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO   ");
            sb.Append(" Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1 ");
            sb.Append(" Where A.Flag=500 AND K.ClassName <> '承銷品'   ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
            sb.Append("                                        UNION ALL");
            sb.Append(" select CASE WHEN T0.BillStatus=2 THEN '無效' ELSE  '未結' END 單況,year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,     ");
            sb.Append(" '忠孝' 公司,T0.DepartID  部門,L.ClassID 客戶類別,L.ClassName 類別名稱,     case when T0.CustomerID = 'tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,   ");
            sb.Append(" case when T0.CustomerID = 'tw90146-16'   THEN  T0.LinkMan   ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN   ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan   ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN   ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan   ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')  ");
            sb.Append(" END End 門市,      ");
            sb.Append(" L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'     ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.ProdID,1,1)='P' THEN '加工品'          ");
            sb.Append(" END 品項,K.ClassName  類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'       ");
            sb.Append(" END '零售/批發',       ");
            sb.Append(" ''''+T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.UserDef1 取貨日期,T0.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(T0.BillDate as varchar) as datetime)) 訂購憑單週數,T0.BillNO 訂購憑單號碼,''  銷貨單據日期,''  銷貨單據週數,'' 銷貨單據號碼,'','', T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名,''      ");
            sb.Append(" ,T1.Quantity 數量,J.Unit 單位,CAST(ROUND(T1.Price,0) AS INT) 單價,CAST(ROUND(T1.Amount,0) AS INT) '金額(未稅)',T1.TaxAmt 稅        ");
            sb.Append(" ,CAST(ROUND(T1.Amount+T1.TaxAmt,0) AS INT) '金額(含稅)',T1.ItemRemark 分錄備註,CASE T1.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,T1.Detail 細項描述,P.PersonName 業務,T0.ProjectID 專案,  CASE RANK() OVER( PARTITION BY T1.BILLNO   ORDER BY T1.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數              ");
            sb.Append(" ,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天        from CHICOMP03.DBO.ordBillMain T0        ");
            sb.Append(" left join CHICOMP03.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)          ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1         ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On T1.ProdID =J.ProdID       ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On L.ClassID =U.ClassID and L.Flag =1            ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append(" LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (T0.AddressID=AD.AddrID AND T0.CustomerID=AD.ID )   left join CHICOMP03.DBO.comPerson P ON (T0.Salesman=P.PersonID)     ");
            sb.Append(" Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1 ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus IN (0,2)  and QtyRemain  > 0 AND K.ClassName <> '承銷品'    ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }

            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }

            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }
            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandTimeout = 0;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC2lassName", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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
        public System.Data.DataTable GetCHO3ANDNOTCLOSE2F()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("       Select  CASE A.Flag WHEN 600 THEN '銷退' WHEN 701 THEN '銷折' ELSE '銷貨' END 類別,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,       ");
            sb.Append("                           '聿豐' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,    ");
            sb.Append("                           case when T.CustID ='tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,      ");
            sb.Append("                           case when T.CustID = 'tw90146-16'   THEN  O2.LinkMan     ");
            sb.Append("                           ELSE CASE WHEN U.ShortName='棉花田' THEN     ");
            sb.Append("                           REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan     ");
            sb.Append("                           ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')     ");
            sb.Append("                           WHEN U.ShortName='安永鮮物' THEN     ");
            sb.Append("                           REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan     ");
            sb.Append("                           ELSE REPLACE(O2.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append("                           END End 門市,        ");
            sb.Append("                           L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append("                           WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'        ");
            sb.Append("                           WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'            ");
            sb.Append("                           END 品項,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'        ");
            sb.Append("                           WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'        ");
            sb.Append("                           END '零售/批發',''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.UserDef1 取貨日期,O2.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 訂購憑單週數,A.FromNO 訂購憑單號碼   ");
            sb.Append("                           ,A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 銷貨單據週數,A.BillNO 銷貨單據號碼,         ");
            sb.Append("                           I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,");
            sb.Append("						   CASE A.Flag WHEN 500 THEN A.Quantity ELSE A.Quantity*-1 END  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,");
            sb.Append("						   CASE A.Flag WHEN 500 THEN CAST(ROUND(A.MLAmount,0) AS INT) ELSE CAST(ROUND(A.MLAmount,0) AS INT)*-1 END  '金額(未稅)'");
            sb.Append("						   , CASE A.Flag WHEN 500 THEN CAST(ROUND(A.TaxAmt ,0) AS INT) ELSE   CAST(ROUND(A.TaxAmt ,0) AS INT)*-1 END 稅         ");
            sb.Append("                           ,CASE A.Flag WHEN 500 THEN CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) ELSE CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT)*-1 END '金額(含稅)',O.ItemRemark 分錄備註,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述,P.PersonName 業務          ");
            sb.Append("                           , CASE RANK() OVER( PARTITION BY O.BILLNO   ORDER BY O.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數             ");
            sb.Append("                           ,U3.FullName  收款方式,U2.DistDays 天    From CHICOMP02.DBO.ComProdRec A      ");
            sb.Append("                           Left Join CHICOMP02.DBO.comWareHouse D On D.WareHouseID=A.WareID          ");
            sb.Append("                           left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag          ");
            sb.Append("                           left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1          ");
            sb.Append("                           left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID          ");
            sb.Append("                           left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2           ");
            sb.Append("                           left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2           ");
            sb.Append("                           left join CHICOMP02.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID           ");
            sb.Append("                           left join CHICOMP02.DBO.COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)          ");
            sb.Append("                           left join CHICOMP02.DBO.comPerson P ON (S.Salesman=P.PersonID)          ");
            sb.Append("                           LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )          ");
            sb.Append("                           Left Join CHICOMP02.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1          ");
            sb.Append("                           Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID          ");
            sb.Append("                           Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID          ");
            sb.Append("                           Left Join CHICOMP02.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1           ");
            sb.Append("                           Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1      ");
            sb.Append("                           Left Join CHICOMP02.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO     ");
            sb.Append("                           Left join CHICOMP02.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1   ");
            sb.Append("                           Left join CHICOMP02.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1     ");
            sb.Append("                           Where A.Flag IN (500,600) AND K.ClassName = '承銷品'    ");


            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }

            sb.Append(" UNION ALL ");
            sb.Append("                Select CASE A.Flag WHEN 600 THEN '銷退' WHEN 701 THEN '銷折' ELSE '銷貨' END 類別,year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,      ");
            sb.Append("              '忠孝' 公司,T.DeptID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,   ");
            sb.Append("              case when T.CustID ='tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,     ");
            sb.Append("              case when T.CustID = 'tw90146-16'   THEN  O2.LinkMan    ");
            sb.Append("              ELSE CASE WHEN U.ShortName='棉花田' THEN    ");
            sb.Append("              REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan    ");
            sb.Append("              ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append("              WHEN U.ShortName='安永鮮物' THEN    ");
            sb.Append("              REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan    ");
            sb.Append("              ELSE REPLACE(O2.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append("              END End 門市,       ");
            sb.Append("              L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'     ");
            sb.Append("              WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'       ");
            sb.Append("              WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'           ");
            sb.Append("              END 品項,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'       ");
            sb.Append("              WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'       ");
            sb.Append("              END '零售/批發',''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.UserDef1 取貨日期,O2.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 訂購憑單週數,A.FromNO 訂購憑單號碼  ");
            sb.Append("              ,A.BillDate  銷貨單據日期,DATEPART(wk, cast(cast(O2.BillDate as varchar) as datetime)) 銷貨單據週數,A.BillNO 銷貨單據號碼,        ");
            sb.Append("              I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,CAST(A.Price AS decimal(18,4)) 單價,CAST(ROUND(A.MLAmount,0) AS INT) '金額(未稅)',A.TaxAmt 稅        ");
            sb.Append("              ,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) '金額(含稅)',O.ItemRemark 分錄備註,CASE F.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,F.Detail 細項描述,P.PersonName 業務       ");
            sb.Append("              , CASE RANK() OVER( PARTITION BY O.BILLNO   ORDER BY O.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數            ");
            sb.Append("              ,U3.FullName  收款方式,U2.DistDays 天    From CHICOMP03.DBO.ComProdRec A     ");
            sb.Append("              Left Join CHICOMP03.DBO.comWareHouse D On D.WareHouseID=A.WareID         ");
            sb.Append("              left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag         ");
            sb.Append("              left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append("              left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID         ");
            sb.Append("              left join CHICOMP03.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2          ");
            sb.Append("              left join CHICOMP03.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2          ");
            sb.Append("              left join CHICOMP03.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID          ");
            sb.Append("              left join CHICOMP03.DBO.COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)         ");
            sb.Append("              left join CHICOMP03.DBO.comPerson P ON (S.Salesman=P.PersonID)         ");
            sb.Append("              LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )         ");
            sb.Append("              Left Join CHICOMP03.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1         ");
            sb.Append("              Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID         ");
            sb.Append("              Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append("              Left Join CHICOMP03.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1          ");
            sb.Append("              Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append("              Left Join CHICOMP03.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO    ");
            sb.Append("              Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1  ");
            sb.Append("              Left join CHICOMP03.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1    ");
            sb.Append("              Where A.Flag IN (500,600,701) AND K.ClassName = '承銷品'    ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
 
            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandTimeout = 0;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC2lassName", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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
        public System.Data.DataTable GetCHO3ANDNOTCLOSE2(string YEAR, string MONTH, string COMPANY, string TYPE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                                                               SELECT CAST(SUM(數量) AS INT) FROM (            Select '已結' ' ',year(cast(cast(I.InvoiceDate as varchar) as datetime)) 年,month(cast(cast(I.InvoiceDate as varchar) as datetime)) 月,   ");
            sb.Append("                                                                                L.ClassID 客戶類別,L.ClassName 類別名稱, ");
            sb.Append("                     case when T.CustID = 'tw90146-16' THEN  O2.LinkMan ELSE  CASE WHEN U.ShortName='棉花田' THEN REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', O2.LinkMan),0) <> 0 THEN O2.LinkMan ELSE REPLACE(O2.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','') END End 門市, ");
            sb.Append("                                                                                L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'   ");
            sb.Append("                                                         WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'    ");
            sb.Append("                                                         WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'       ");
            sb.Append("                                                         END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'    ");
            sb.Append("                                                         WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'    ");
            sb.Append("                                                         END '零售/批發',''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,O2.BillDate 訂購憑單日期,A.FromNO 訂購憑單號碼,A.BillDate  銷貨單據日期,A.BillNO 銷貨單據號碼,     ");
            sb.Append("                                                                                               I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount '金額(未稅)',A.TaxAmt 稅     ");
            sb.Append("                                                                                               ,A.MLAmount+A.TaxAmt '金額(含稅)'     ");
            sb.Append("                                                                                             From ComProdRec A       ");
            sb.Append("                                                                                              Left Join comWareHouse D On D.WareHouseID=A.WareID      ");
            sb.Append("                                                                                             left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag      ");
            sb.Append("                                                                                                  left join comCustomer U On  U.ID=T.CustID AND U.Flag =1      ");
            sb.Append("                                                                                                         left join comWareHouse W On  A.WareID=W.WareHouseID      ");
            sb.Append("                                                                                                     left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2       ");
            sb.Append("                                                                                                          left join OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2       ");
            sb.Append("                                                                                                     left join comCurrencySys C On  O2.CurrID=C.CurrencyID       ");
            sb.Append("                                                                                                        left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)      ");
            sb.Append("                                                                                                                   left join comPerson P ON (S.Salesman=P.PersonID)      ");
            sb.Append("                                                                                       LEFT JOIN comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )      ");
            sb.Append("                                                                                        Left Join comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =2 AND I.IsCancel <> 1      ");
            sb.Append("                                                                                                 Left Join comProduct J On A.ProdID =J.ProdID      ");
            sb.Append("                                                                                                         Left Join comProductClass K On J.ClassID =K.ClassID      ");
            sb.Append("                                                                                           Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1       ");
            sb.Append("                                                                                                  Where A.Flag=500   AND U.ShortName ='棉花田' AND A.MLAmount <> 0 ");

            sb.Append("                                                                                                  AND year(cast(cast(O.PreInDate as varchar) as datetime))=@YEAR AND month(cast(cast(O.PreInDate as varchar) as datetime))=@MONTH");


                sb.Append("                                                                                           AND     SUBSTRING(K.ClassID,3,1)=@TYPE ");
            
            sb.Append("                                                                                                  UNION ALL");
            sb.Append("                        select  '未結' ' ',year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,   ");
            sb.Append("                                                                                 L.ClassID 客戶類別,L.ClassName 類別名稱, ");
            sb.Append("                   case when T0.CustomerID = 'tw90146-16' THEN  T0.LinkMan ELSE      CASE WHEN U.ShortName='棉花田' THEN REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan ELSE REPLACE(T0.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','') END End 門市, ");
            sb.Append("                                                                                 L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'    ");
            sb.Append("                                                                       WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'     ");
            sb.Append("                                                                       WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.ProdID,1,1)='P' THEN '加工品'       ");
            sb.Append("                                                                       END 品項,K.ClassName  類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'     ");
            sb.Append("                                                                       WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'     ");
            sb.Append("                                                                       END '零售/批發',     ");
            sb.Append("                                                                                       ''''+T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.BillDate 訂購憑單日期,T0.BillNO 訂購憑單號碼,''  銷貨單據日期,'' 銷貨單據號碼,'','', T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名,''    ");
            sb.Append("                                                                                     ,T1.Quantity 數量,J.Unit 單位,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅      ");
            sb.Append("                                                                                                             ,T1.Amount+T1.TaxAmt '金額(含稅)'  from ordBillMain T0      ");
            sb.Append("                                                                                     left join ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)        ");
            sb.Append("                                                                                      left join comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1       ");
            sb.Append("                                                                                                           Left Join comProduct J On T1.ProdID =J.ProdID     ");
            sb.Append("                                                                                             Left Join comProductClass K On J.ClassID =K.ClassID       ");
            sb.Append("                                                                                                                                               Left Join comCustClass L On L.ClassID =U.ClassID and L.Flag =1          ");
            sb.Append("                                                                                     WHERE  T0.Flag =2 AND T0.BillStatus =0  and QtyRemain  > 0      AND U.ShortName ='棉花田' AND T1.Amount <> 0 ");

            sb.Append("                                                                                                         AND   year(cast(cast(T1.PreInDate as varchar) as datetime))=@YEAR AND  month(cast(cast(T1.PreInDate as varchar) as datetime))=@MONTH ");

                sb.Append("                                                                                            AND    SUBSTRING(K.ClassID,3,1)=@TYPE ");
                sb.Append("                                                                                                    ) AS A");
                sb.Append("                                                                                                         WHERE 門市=@COMPANY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
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
        public System.Data.DataTable GetCHO4()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select  year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,'聿豐' 公司,T.DeptID 部門,  ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品'  WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'    ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'      ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'   ");
            sb.Append(" END '零售/批發',T.UDef1 批號,   ");
            sb.Append(" T.CustID 廠商代碼,U.ShortName 廠商簡稱,O2.BillDate 採購憑單日期,A.FromNO 採購憑單號碼,A.BillDate  進貨單據日期,A.BillNO 進貨單據號碼,    ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,''''+T.VoucherNO 傳票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount '金額(未稅)',A.TaxAmt 稅    ");
            sb.Append(" ,A.MLAmount+A.TaxAmt '金額(含稅)',A.Detail 細項描述,A.ItemRemark 分錄備註,T.Remark  備註 ");
            sb.Append(" From CHICOMP02.DBO.ComProdRec A      ");
            sb.Append(" Left Join CHICOMP02.DBO.comWareHouse D On D.WareHouseID=A.WareID     ");
            sb.Append(" left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag     ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =2     ");
            sb.Append(" left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID     ");
            sb.Append(" left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4     ");
            sb.Append(" left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4      ");
            sb.Append(" left join CHICOMP02.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID      ");
            sb.Append(" left join CHICOMP02.DBO.comPerson P ON (T.Salesman=P.PersonID)     ");
            sb.Append(" LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )     ");
            sb.Append(" Left Join CHICOMP02.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =1 AND I.IsCancel <> 1     ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID     ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID     ");
            sb.Append(" Where A.Flag=100     AND K.ClassID <> 'CI1010'    ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
            sb.Append(" UNION ALL ");
            sb.Append(" Select  year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,'忠孝' 公司,T.DeptID 部門,  ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'    ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'      ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'   ");
            sb.Append(" END '零售/批發',T.UDef1 批號,   ");
            sb.Append(" T.CustID 廠商代碼,U.ShortName 廠商簡稱,O2.BillDate 採購憑單日期,A.FromNO 採購憑單號碼,A.BillDate  進貨單據日期,A.BillNO 進貨單據號碼,    ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,''''+T.VoucherNO 傳票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount '金額(未稅)',A.TaxAmt 稅    ");
            sb.Append(" ,A.MLAmount+A.TaxAmt '金額(含稅)',O.Detail 細項描述,O.ItemRemark 分錄備註,T.Remark  備註 ");
            sb.Append(" From CHICOMP03.DBO.ComProdRec A      ");
            sb.Append(" Left Join CHICOMP03.DBO.comWareHouse D On D.WareHouseID=A.WareID     ");
            sb.Append(" left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag     ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =2     ");
            sb.Append(" left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID     ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4     ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4      ");
            sb.Append(" left join CHICOMP03.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID      ");
            sb.Append(" left join CHICOMP03.DBO.comPerson P ON (T.Salesman=P.PersonID)     ");
            sb.Append(" LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )     ");
            sb.Append(" Left Join CHICOMP03.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =1 AND I.IsCancel <> 1     ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID     ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID     ");
            sb.Append(" Where A.Flag=100      AND K.ClassID <> 'CI1010'  ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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
        public System.Data.DataTable GetCH11(string DOCTYPE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                                Select CASE  A.Flag WHEN 700 THEN '進折' WHEN 200 THEN '進退'  WHEN 100 THEN '進貨'  END 類別 ,  year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,'聿豐' 公司,T.DeptID 部門,     ");
            sb.Append("                                        CASE WHEN K.ClassID='ARP100' THEN '外購品'  WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'       ");
            sb.Append("                                        WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'      ");
            sb.Append("                                        WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'         ");
            sb.Append("                                        END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'      ");
            sb.Append("                                        WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'      ");
            sb.Append("                                        END '零售/批發',T.UDef1 批號,      ");
            sb.Append("                                        T.CustID 廠商代碼,U.ShortName 廠商簡稱,O2.BillDate 採購憑單日期,A.FromNO 採購憑單號碼,A.BillDate  進貨單據日期,A.BillNO 進貨單據號碼,       ");
            sb.Append("                                        I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,''''+T.VoucherNO 傳票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,");
            sb.Append("										CASE A.Flag WHEN 700 THEN 0 ELSE A.Quantity END 數量,J.Unit 單位,");
            sb.Append("											CASE A.Flag WHEN 700 THEN 0 ELSE A.Price END 單價,     ");
            sb.Append("                                        CASE A.Flag WHEN 700 THEN F.Dist*-1  ELSE A.Amount  END '金額(未稅)',   ");
            sb.Append("                                        CASE A.Flag WHEN 700 THEN  F.DistTaxAmt*-1 ELSE A.TaxAmt    END 稅,   ");
            sb.Append("                                        CASE A.Flag WHEN 700 THEN (F.Dist+F.DistTaxAmt)*-1 ELSE (A.Amount+A.TaxAmt)  END '金額(含稅)',  ");
            sb.Append("                                        A.Detail 細項描述,A.ItemRemark 分錄備註,T.Remark  備註    ");
            sb.Append("                                        From CHICOMP02.DBO.ComProdRec A         ");
            sb.Append("                                        Left Join CHICOMP02.DBO.comWareHouse D On D.WareHouseID=A.WareID        ");
            sb.Append("                                        left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=CASE T.Flag WHEN 298 THEN 700 ELSE  T.Flag END       ");
            sb.Append("                                        left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =2        ");
            sb.Append("                                        left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID        ");
            sb.Append("                                        left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4        ");
            sb.Append("                                        left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4         ");
            sb.Append("                                        left join CHICOMP02.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID         ");
            sb.Append("                                        left join CHICOMP02.DBO.comPerson P ON (T.Salesman=P.PersonID)        ");
            sb.Append("                                        LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )        ");
            sb.Append("                                        Left Join CHICOMP02.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag =1 AND I.IsCancel <> 1        ");
            sb.Append("                                        Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID        ");
            sb.Append("                                        Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID        ");
            sb.Append("             						    Left Join CHICOMP02.DBO.stkDistSub F On A.BillNO =F.DISTNO and A.Flag =F.Flag AND A.RowNO =F.RowNO ");




            if (DOCTYPE == "1")
            {
                sb.Append("              Where A.Flag IN (200,700)     ");
            }
            else

            {
                sb.Append("              Where A.Flag IN (100,200,700)     ");
            }

            //CI1010
            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }

            if (DOCTYPE == "1")
            {

                sb.Append("    AND K.ClassID <> 'CI1010' ");
            }
            else
            {

                sb.Append("    AND K.ClassID = 'CI1010'");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }

            sb.Append(" UNION ALL ");
            sb.Append("              Select CASE  A.Flag WHEN 700 THEN '進折' WHEN 200 THEN '進退'  WHEN 100 THEN '進貨'  END 類別 ,  year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,'忠孝' 公司,T.DeptID 部門,   ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'    ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'      ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'   ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'   ");
            sb.Append(" END '零售/批發',T.UDef1 批號,   ");
            sb.Append(" T.CustID 廠商代碼,U.ShortName 廠商簡稱,O2.BillDate 採購憑單日期,A.FromNO 採購憑單號碼,A.BillDate  進貨單據日期,A.BillNO 進貨單據號碼,    ");
            sb.Append(" I.InvoiceDate 發票日期,I.InvoiceNO 發票號碼,''''+T.VoucherNO 傳票號碼,A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount '金額(未稅)',A.TaxAmt 稅    ");
            sb.Append(" ,A.MLAmount+A.TaxAmt '金額(含稅)',O.Detail 細項描述,O.ItemRemark 分錄備註,T.Remark  備註 ");
            sb.Append(" From CHICOMP03.DBO.ComProdRec A      ");
            sb.Append(" Left Join CHICOMP03.DBO.comWareHouse D On D.WareHouseID=A.WareID     ");
            sb.Append(" left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag     ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =2     ");
            sb.Append(" left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID     ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4     ");
            sb.Append(" left join CHICOMP03.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4      ");
            sb.Append(" left join CHICOMP03.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID      ");
            sb.Append(" left join CHICOMP03.DBO.comPerson P ON (T.Salesman=P.PersonID)     ");
            sb.Append(" LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )     ");
            sb.Append(" Left Join CHICOMP03.DBO.comInvoice I On  A.BillNO=I.SrcBillNO AND I.Flag IN (1,3)  AND I.IsCancel <> 1     ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID     ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID     ");
            if (DOCTYPE == "1")
            {
                sb.Append("              Where A.Flag IN (200,700)     ");
            }
            else
            {
                sb.Append("              Where A.Flag IN (100,200,700)     ");
            }

            if (checkBox1.Checked)
            {
                sb.Append(" and  T.[CustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                if (comboBox7.Text == "出貨日期")
                {
                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
                }
                else
                {
                    sb.Append("  AND O2.BillDate  between @BillDate1 and @BillDate2 ");
                }
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }

            if (DOCTYPE == "1")
            {

                sb.Append("    AND K.ClassID <> 'CI1010' ");
            }
            else
            {

                sb.Append("    AND K.ClassID = 'CI1010'");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }

            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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

        public System.Data.DataTable GetCHO5()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                       Select '聿豐' 公司,T1.DepartID 部門, ");
            sb.Append("               CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'  ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞' ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'     ");
            sb.Append("               END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛' ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發' ");
            sb.Append("               END '零售/批發',T2.ClassName 調整類別,T1.UDef1 批號, A.BillDate  調整憑單日期,A.BillNO  調整憑單號碼,  ");
            sb.Append("                                                    A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount 金額  ");
            sb.Append("                                 ,A.ItemRemark 分錄備註,T1.Remark 備註  From CHICOMP02.DBO.ComProdRec A    ");
            sb.Append("                                                   LEFT JOIN CHICOMP02.DBO.StkAdjustMain T1 ON (A.BillNO =T1.AdjustNO)");
            sb.Append("                                                           LEFT JOIN CHICOMP02.DBO.stkAdjustClass T2 ON (T1.AdjustType =T2.ClassID)");
            sb.Append("                                                               left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID   ");
            sb.Append("                                                       Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID   ");
            sb.Append("                                                                  Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID   ");
            sb.Append("                                                        Where A.Flag=300   ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }
            if (comboBox5.Text != "")
            {

                sb.Append(" AND T2.ClassName=@CC ");
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text+ "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T1.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T1.DepartID  BETWEEN @C1 AND @C2 ");
            }

            sb.Append(" UNION ALL");
            sb.Append("                                       Select '忠孝' 公司,T1.DepartID 部門, ");
            sb.Append("               CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'  ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞' ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(A.ProdID,1,1)='P' THEN '加工品'     ");
            sb.Append("               END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛' ");
            sb.Append("               WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發' ");
            sb.Append("               END '零售/批發',T2.ClassName 調整類別,T1.UDef1 批號, A.BillDate  調整憑單日期,A.BillNO  調整憑單號碼,  ");
            sb.Append("                                                    A.ProdID 產品編號,A.ProdName 品名規格,J.InvoProdName 發票品名, W.WareHouseName 倉別,A.Quantity  數量,J.Unit 單位,A.Price 單價,A.MLAmount 金額  ");
            sb.Append("                                 ,A.ItemRemark 分錄備註,T1.Remark 備註  From CHICOMP03.DBO.ComProdRec A    ");
            sb.Append("                                                   LEFT JOIN CHICOMP03.DBO.StkAdjustMain T1 ON (A.BillNO =T1.AdjustNO)");
            sb.Append("                                                           LEFT JOIN CHICOMP03.DBO.stkAdjustClass T2 ON (T1.AdjustType =T2.ClassID)");
            sb.Append("                                                               left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID   ");
            sb.Append("                                                       Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID   ");
            sb.Append("                                                                  Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID   ");
            sb.Append("                                                        Where A.Flag=300   ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }
            if (comboBox5.Text != "")
            {

                sb.Append(" AND T2.ClassName=@CC ");
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND A.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox1.Text != "")
            {

                sb.Append("  AND T1.UDef1  LIKE '%" + textBox1.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T1.DepartID  BETWEEN @C1 AND @C2 ");
            }
            sb.Append("  ORDER BY 公司,A.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CC", comboBox5.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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

        public System.Data.DataTable GetCHO6()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select '聿豐' 公司,year(cast(cast(T0.MakeDate as varchar) as datetime)) 年,month(cast(cast(T0.MakeDate as varchar) as datetime)) 月,T0.MakeDate 傳票日期,''''+A.VoucherNo 傳票號碼,A.SubjectID 科目,C.SubjectName 科目名稱,A.Amount 金額,A.Summary 備註");
            sb.Append(" From CHICOMP02.DBO.AccVoucherSub A ");
            sb.Append(" Left Join CHICOMP02.DBO.comDepartment B On B.DepartID=A.DepartID ");
            sb.Append(" Left Join CHICOMP02.DBO.ComSubject C On C.SubjectID=A.SubjectID ");
            sb.Append(" Left Join CHICOMP02.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
            sb.Append(" WHERE  A.DepartID IN ('C1','F1','F2','C2') ");
            sb.Append(" and (A.SubjectID in (1201002,6202002,6205000,6208005,6210020,6210060,6223000,6226001,6226002,6226003,6227000,6236000,6237000,6238000 ");
            sb.Append(" ,6251000,6271001,7105000,7304000,6242004) OR A.Summary LIKE '%DM%' )");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.MakeDate  between @BillDate1 and @BillDate2 ");
            }

            if (!String.IsNullOrEmpty(K))
            {
                sb.Append(" and  A.SubjectID in ( " + K + ") ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" Select '忠孝' 公司,year(cast(cast(T0.MakeDate as varchar) as datetime)) 年,month(cast(cast(T0.MakeDate as varchar) as datetime)) 月,T0.MakeDate 傳票日期,''''+A.VoucherNo 傳票號碼,A.SubjectID 科目,C.SubjectName 科目名稱,A.Amount 金額,A.Summary 備註");
            sb.Append(" From CHICOMP03.DBO.AccVoucherSub A ");
            sb.Append(" Left Join CHICOMP03.DBO.comDepartment B On B.DepartID=A.DepartID ");
            sb.Append(" Left Join CHICOMP03.DBO.ComSubject C On C.SubjectID=A.SubjectID ");
            sb.Append(" Left Join CHICOMP03.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
            sb.Append(" WHERE  A.DepartID IN ('C1','F1','F2','C2') ");
            sb.Append(" and (A.SubjectID in (1201002,6202002,6205000,6208005,6210020,6210060,6223000,6226001,6226002,6226003,6227000,6236000,6237000,6238000 ");
            sb.Append(" ,6251000,6271001,7105000,7304000,6242004) OR A.Summary LIKE '%DM%' )");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.MakeDate  between @BillDate1 and @BillDate2 ");
            }

            if (!String.IsNullOrEmpty(K))
            {
                sb.Append(" and  A.SubjectID in ( " + K + ") ");
            }
            sb.Append("  ORDER BY 公司,T0.MakeDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GetCHO8()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'         ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T0.ProdID,1,1)='P' THEN '加工品'           ");
            sb.Append(" END 品項,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'         ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'         ");
            sb.Append(" END '零售/批發',         ");
            sb.Append(" T0.ProdID 料號,T2.InvoProdName 發票品名,T2.Unit 單位,       ");
            sb.Append(" ISNULL(CAST(T1.QTY AS decimal(10,2)),0)-ISNULL((select CAST(SUM(QtyRemain) AS  decimal(10,2)) from OrdBillSub TT left join ordBillMain TS ON (TT.Flag =TS.Flag AND TT.BillNO=TS.BillNO)  WHERE T0.PRODID=TT.PRODID AND TS.Flag =2 AND TS.BillStatus =0  GROUP BY TT.PRODID),0) 可用量,       ");
            sb.Append(" (select CAST(SUM(QtyRemain) AS  decimal(10,2)) from OrdBillSub TT left join ordBillMain TS ON (TT.Flag =TS.Flag AND TT.BillNO=TS.BillNO)  WHERE T0.PRODID=TT.PRODID AND TS.Flag =2 AND TS.BillStatus =0  GROUP BY TT.PRODID) 訂單未交量,      ");
            sb.Append(" ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A16' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A17' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A18' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A19' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N01' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A14' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A21' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A06' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'OT002' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A12' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N02' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'R10' AND T0.PRODID=W20.PRODID ),0)  庫存總量,        ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID ) 逢泰數量,      ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID )*T0.CAvgCost  逢泰金額,   ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N01' AND T0.PRODID=W20.PRODID ) 內湖數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'N01' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   內湖金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A14' AND T0.PRODID=W20.PRODID ) 麟洛數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A14' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   麟洛金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A16' AND T0.PRODID=W20.PRODID ) EzWaven數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A16' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   EzWaven金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A17' AND T0.PRODID=W20.PRODID ) 台灣好食材數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A17' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   台灣好食材金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A18' AND T0.PRODID=W20.PRODID ) MOMO數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A18' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   MOMO金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A19' AND T0.PRODID=W20.PRODID ) 博客來數量,        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A19' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   博客來金額,  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A21' AND T0.PRODID=W20.PRODID ) '逢泰-客倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A21' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '逢泰-客倉金額',  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A06' AND T0.PRODID=W20.PRODID ) '聿豐-在途倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A06' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '聿豐-在途倉金額',  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'OT002' AND T0.PRODID=W20.PRODID ) '聿-捐贈倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'OT002' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '聿-捐贈倉金額',  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A12' AND T0.PRODID=W20.PRODID ) '逢泰-不良品倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'A12' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '逢泰-不良品倉金額',  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N02' AND T0.PRODID=W20.PRODID ) '內湖-不良品倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'N02' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '內湖-不良品倉金額',  ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'R10' AND T0.PRODID=W20.PRODID ) '食品加工倉數量',        ");
            sb.Append(" (select SUM(Quantity)from comWareAmount W20  WHERE  WareID = 'R10' AND T0.PRODID=W20.PRODID )*T0.CAvgCost   '食品加工倉金額'  ");
            sb.Append(" FROM comProduct T0       ");
            sb.Append(" LEFT JOIN (select SUM(Quantity) QTY,PRODID from comWareAmount TS  GROUP BY TS.PRODID) T1 ON (T0.PRODID=T1.PRODID)      ");
            sb.Append(" LEFT JOIN comProduct T2 ON (T0.ProdID=T2.ProdID)      ");
            sb.Append(" LEFT JOIN comProductClass K On T2.ClassID =K.ClassID     ");
            sb.Append(" WHERE  T0.ProdName NOT LIKE '%FEE%'           AND                ");
            sb.Append(" ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A08' AND T0.PRODID=W20.PRODID),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A16' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A17' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A18' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A19' AND T0.PRODID=W20.PRODID ),0)     ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N01' AND T0.PRODID=W20.PRODID ),0)    ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A14' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A21' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A06' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'OT002' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'A12' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'N02' AND T0.PRODID=W20.PRODID ),0)   ");
            sb.Append(" +ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'R10' AND T0.PRODID=W20.PRODID ),0)  ");
            sb.Append(" > 0    ");

            if (checkBox2.Checked)
            {
                sb.Append(" and   T0.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T0.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T2.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
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



        public System.Data.DataTable GetCHO82()
        {

            SqlConnection MyConnection = new SqlConnection(strCn3);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'  WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'          ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T0.ProdID,1,1)='P' THEN '加工品'            ");
            sb.Append(" END 品項,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'          ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'          ");
            sb.Append(" END '零售/批發',          ");
            sb.Append(" T0.ProdID 料號,T2.InvoProdName 發票品名,T2.Unit 單位,        ");
            sb.Append(" ISNULL(CAST(T1.QTY AS decimal(10,2)),0)-ISNULL((select CAST(SUM(QtyRemain) AS  decimal(10,2)) from OrdBillSub TT left join ordBillMain TS ON (TT.Flag =TS.Flag AND TT.BillNO=TS.BillNO)  WHERE T0.PRODID=TT.PRODID AND TS.Flag =2 AND TS.BillStatus =0  GROUP BY TT.PRODID),0) 可用量,             ");
            sb.Append(" ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'B01' AND T0.PRODID=W20.PRODID),0)   庫存總量,         ");
            sb.Append(" (select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'B01' AND T0.PRODID=W20.PRODID ) 忠孝數量");
            sb.Append(" FROM comProduct T0        ");
            sb.Append(" LEFT JOIN (select SUM(Quantity) QTY,PRODID from comWareAmount TS  GROUP BY TS.PRODID) T1 ON (T0.PRODID=T1.PRODID)       ");
            sb.Append(" LEFT JOIN comProduct T2 ON (T0.ProdID=T2.ProdID)       ");
            sb.Append(" LEFT JOIN comProductClass K On T2.ClassID =K.ClassID      ");
            sb.Append(" WHERE  T0.ProdName NOT LIKE '%FEE%'           AND                 ");
            sb.Append(" ISNULL((select CAST(SUM(Quantity) AS decimal(10,2))  from comWareAmount W20  WHERE  WareID = 'B01' AND T0.PRODID=W20.PRODID),0)    ");
            sb.Append(" > 0   ");
            if (checkBox2.Checked)
            {
                sb.Append(" and   T0.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T0.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T2.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
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
        public System.Data.DataTable GetCHO3NOTCLOSE()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,      ");
            sb.Append(" '聿豐' 公司,T0.DepartID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,   ");
            sb.Append(" case when T0.CustomerID = 'tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,     ");
            sb.Append(" case when T0.CustomerID  = 'tw90146-16'   THEN  T0.LinkMan    ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN     ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan     ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')     ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN     ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan     ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append(" END End 門市,L.ENGNAME 'B/C',      ");
            sb.Append(" CASE WHEN SUBSTRING(T1.ProdID ,1,2)='A1' THEN '外購品' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'FRE' THEN '運費' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MCK' THEN '雞' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MPK' THEN '豬' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MSR' THEN '蝦' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'MSF' THEN '魚' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PCK' THEN '加工品-雞' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PFH' THEN '加工品-魚' ");
            sb.Append(" WHEN SUBSTRING(T1.ProdID ,1,3) = 'PPK' THEN '加工品-豬'  ");
            sb.Append(" END 品項,K.ClassName  類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'        ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'        ");
            sb.Append(" END '零售/批發',        ");
            sb.Append(" ''''+T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.UserDef1 取貨日期,  ");
            sb.Append(" T0.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(T0.BillDate as varchar) as datetime)) 訂購憑單週數  ");
            sb.Append(" ,T0.BillNO 訂購憑單號碼,''''+T0.CustBillNo 客戶訂單,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名         ");
            sb.Append(" ,T1.Quantity 數量,J.Unit 單位,CAST(T1.Price AS decimal(18,4)) 單價,CAST(ROUND(T1.Amount,0) AS INT) '金額(未稅)',CAST(ROUND(T1.TaxAmt,0) AS INT) 稅         ");
            sb.Append(" ,CAST(ROUND(T1.Amount+T1.TaxAmt,0) AS INT) '金額(含稅)',T1.ItemRemark 分錄備註,CASE T1.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,T1.Detail 細項描述,CASE RANK() OVER( PARTITION BY T1.BILLNO   ORDER BY T1.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數  ");
            sb.Append(" ,CASE T0.GatherStyle  WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN T0.GatherOther END 收款方式,T0.GatherDelay  天     from CHICOMP02.DBO.ordBillMain T0         ");
            sb.Append(" left join CHICOMP02.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID          ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On L.ClassID =U.ClassID and L.Flag =1            ");
            sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1      ");
            sb.Append(" LEFT JOIN CHICOMP02.DBO.comCustAddress AD ON (T0.AddressID=AD.AddrID AND T0.CustomerID=AD.ID )    ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0 ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }

            sb.Append(" UNION ALL ");
            sb.Append(" select  year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,     ");
            sb.Append(" '忠孝' 公司,T0.DepartID 部門,L.ClassID 客戶類別,L.ClassName 類別名稱,  ");
            sb.Append(" case when T0.CustomerID = 'tw90146-16' then AD.LinkManProf  ELSE  M.AddField1 END 來源,    ");
            sb.Append(" case when T0.CustomerID  = 'tw90146-16'   THEN  T0.LinkMan   ");
            sb.Append(" ELSE CASE WHEN U.ShortName='棉花田' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan    ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'棉花田',''),'-',''),'門市',''),'門巿','')    ");
            sb.Append(" WHEN U.ShortName='安永鮮物' THEN    ");
            sb.Append(" REPLACE(REPLACE(REPLACE(REPLACE(CASE WHEN ISNULL(CHARINDEX('新店', T0.LinkMan),0) <> 0 THEN T0.LinkMan    ");
            sb.Append(" ELSE REPLACE(T0.LinkMan,'店','') END,'安永鮮物',''),'-',''),'門市',''),'門巿','')   ");
            sb.Append(" END End 門市,     ");
            sb.Append(" L.ENGNAME 'B/C',CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.ProdID,1,1)='P' THEN '加工品'         ");
            sb.Append(" END 品項,K.ClassName  類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '零售'   WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'       ");
            sb.Append(" END '零售/批發',       ");
            sb.Append("  ''''+T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.UserDef1 取貨日期, ");
            sb.Append(" T0.BillDate 訂購憑單日期,DATEPART(wk, cast(cast(T0.BillDate as varchar) as datetime)) 訂購憑單週數 ");
            sb.Append(" ,T0.BillNO 訂購憑單號碼,''''+T0.CustBillNo 客戶訂單,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名        ");
            sb.Append(" ,T1.Quantity 數量,J.Unit 單位,CAST(T1.Price AS decimal(18,4)) 單價,CAST(ROUND(T1.Amount,0) AS INT) '金額(未稅)',CAST(ROUND(T1.TaxAmt,0) AS INT) 稅        ");
            sb.Append(" ,CAST(ROUND(T1.Amount+T1.TaxAmt,0) AS INT) '金額(含稅)',T1.ItemRemark 分錄備註,CASE T1.IsGift WHEN 1 THEN '是' ELSE '否' END 是否為贈品,T1.Detail 細項描述,CASE RANK() OVER( PARTITION BY T1.BILLNO   ORDER BY T1.SerNO ) WHEN 1 THEN 1 ELSE 0 END 筆數 ");
            sb.Append(" ,CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天     from CHICOMP03.DBO.ordBillMain T0        ");
            sb.Append(" left join CHICOMP03.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)          ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1         ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On T1.ProdID =J.ProdID       ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID         ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On L.ClassID =U.ClassID and L.Flag =1           ");
            sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1     ");
            sb.Append(" LEFT JOIN CHICOMP03.DBO.comCustAddress AD ON (T0.AddressID=AD.AddrID AND T0.CustomerID=AD.ID )   ");
            sb.Append(" Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1 ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0");
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {

                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }

            if (checkBox5.Checked)
            {
                sb.Append(" and    L.ClassName in ( " + MM + ") ");
            }
            else
            {
                if (comboBox3.Text != "")
                {

                    sb.Append("  AND L.ClassName  = @CClassName ");
                }
            }

            if (checkBox6.Checked)
            {
                sb.Append(" and    L.EngName in ( " + MM2 + ") ");
            }
            else
            {
                if (comboBox4.Text != "")
                {

                    sb.Append("  AND L.EngName  = @CC2lassName ");
                }
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }

            sb.Append("  ORDER BY 公司,T0.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC2lassName", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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
        public System.Data.DataTable GetCHO4NOTCLOSE()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,'聿豐' 公司,T0.DepartID  部門,  ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.ProdID,1,1)='P' THEN '加工品'     ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'  ");
            sb.Append(" END '零售/批發',T0.CustomerID  廠商代碼,U.ShortName 廠商簡稱,T0.BillDate 採購憑單日期,T0.BillNO 採購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名   ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.MLAmount '金額(未稅)',T1.TaxAmt 稅   ");
            sb.Append(" ,T1.MLAmount+T1.TaxAmt '金額(含稅)',T1.Detail 細項描述,T1.ItemRemark 分錄備註,T0.REMARK 備註 from CHICOMP02.DBO.ordBillMain T0   ");
            sb.Append(" left join CHICOMP02.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)     ");
            sb.Append(" left join CHICOMP02.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =2    ");
            sb.Append(" Left Join CHICOMP02.DBO.comProduct J On T1.ProdID =J.ProdID  ");
            sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID    ");
            sb.Append(" WHERE  T0.Flag =4 AND T0.BillStatus =0    ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {
                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }

            sb.Append(" UNION ALL");
            sb.Append(" select year(cast(cast(T0.BillDate as varchar) as datetime)) 年,month(cast(cast(T0.BillDate as varchar) as datetime)) 月,'忠孝' 公司,T0.DepartID  部門,  ");
            sb.Append(" CASE WHEN K.ClassID='ARP100' THEN '外購品' WHEN K.ClassID='ACME M' THEN '豬' WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費' WHEN K.ClassID IN ('AWS220','ARS220') THEN '烏魚'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.ProdID,1,1)='P' THEN '加工品'     ");
            sb.Append(" END 品項,K.ClassName 類別,CASE WHEN K.ClassID='ACME M' THEN '毛' WHEN K.ClassID='ACMECM' THEN '毛'  ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,2,1)='R' THEN '零售' WHEN SUBSTRING(K.ClassID,2,1)='W' THEN '批發'  ");
            sb.Append(" END '零售/批發',T0.CustomerID  廠商代碼,U.ShortName 廠商簡稱,T0.BillDate 採購憑單日期,T0.BillNO 採購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名   ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.MLAmount '金額(未稅)',T1.TaxAmt 稅   ");
            sb.Append(" ,T1.MLAmount+T1.TaxAmt '金額(含稅)' ,T1.Detail 細項描述,T1.ItemRemark 分錄備註,T0.REMARK 備註 from CHICOMP03.DBO.ordBillMain T0   ");
            sb.Append(" left join CHICOMP03.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)     ");
            sb.Append(" left join CHICOMP03.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =2    ");
            sb.Append(" Left Join CHICOMP03.DBO.comProduct J On T1.ProdID =J.ProdID  ");
            sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID    ");
            sb.Append(" WHERE  T0.Flag =4 AND T0.BillStatus =0    ");

            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CustomerID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T0.[CustomerID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND T0.BillDate  between @BillDate1 and @BillDate2 ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   T1.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T1.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + M + ") ");
            }
            else
            {
                if (comboBox2.Text != "")
                {
                    sb.Append("  AND K.ClassName  = @ClassName ");
                }
            }
            if (textBox2.Text != "")
            {
                sb.Append("  AND T1.ProdName  LIKE '%" + textBox2.Text + "%' ");
            }
            if (textBox3.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T0.DepartID BETWEEN @C1 AND @C2 ");
            }

   
            sb.Append("  ORDER BY 公司,T0.BillDate ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@ClassName", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
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
        private void SACHOICE_Load(object sender, EventArgs e)
        {
            comboBox7.Text = "出貨日期";
            UtilSimple.SetLookupBinding(comboBox3, BU(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, BC(), "DataValue", "DataValue");

            System.Data.DataTable dt3 = GetBU();

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }

            System.Data.DataTable dt4 = GetBUT();

            comboBox5.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox5.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }

            comboBox1.Text = "客戶";
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();

            if (globals.GroupID.ToString().Trim() != "EEP")
            {

                button7.Visible = false;
                button9.Visible = false;
    
            }

        }

        public  System.Data.DataTable BU()
        {
            SqlConnection con = new SqlConnection(strCn);

            string sql = "SELECT '' DataValue UNION ALL  SELECT CLASSNAME  DataValue from comCustClass where Flag =1 AND ClassID >008  ";

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
        public System.Data.DataTable BC()
        {
            SqlConnection con = new SqlConnection(strCn);

            string sql = "SELECT '' DataValue UNION ALL  SELECT DISTINCT EngName   from comCustClass where Flag =1 AND ClassID >008 AND ISNULL(EngName,'') <> ''   ";

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
        public string c;
        private void button8_Click(object sender, EventArgs e)
        {

            APS1CHOICE frm1 = new APS1CHOICE();
            frm1.CARDTYPE = comboBox1.Text;
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox1.Checked = true;
                c = frm1.q;

            }
        }
        public System.Data.DataTable GetBU()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT '' UNION ALL SELECT ClassName FROM comProductClass WHERE SUBSTRING(ClassID,1,1)='A' AND ClassID NOT IN ('ACMEFR','AMM100','AMO100','AMV100')  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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


        public System.Data.DataTable GetBUT()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT '' UNION ALL SELECT CLASSNAME  FROM stkAdjustClass WHERE CLASSID > '16'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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


        private void button3_Click(object sender, EventArgs e)
        {

    
            if (tabControl1.SelectedTab  == 採購單已結報表)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedTab == 採購單未結報表)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedTab == 銷貨已結報表)
            {
              //  ExcelReport.GridViewToExcel(dataGridView1);
                string GG = @"\銷貨已結報表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                ExcelReport.GridViewToCSV(dataGridView1, Environment.CurrentDirectory + "\\Excel\\temp\\" + GG);
            }
            else if (tabControl1.SelectedTab == 銷貨未結報表)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedTab == 調整憑單報表)
            {
                ExcelReport.GridViewToExcel(dataGridView5);
            }
            else if (tabControl1.SelectedTab == 費用單)
            {
                ExcelReport.GridViewToExcel(dataGridView6);
            }
            else if (tabControl1.SelectedTab == 銷貨總表)
            {
             //   ExcelReport.GridViewToExcel(dataGridView7);

                string GG = @"\銷貨總表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                ExcelReport.GridViewToCSV(dataGridView7, Environment.CurrentDirectory + "\\Excel\\temp\\" + GG);
            }
            else if (tabControl1.SelectedTab == 庫存總表)
            {
                ExcelReport.GridViewToExcel(dataGridView8);
            }
            else if (tabControl1.SelectedTab == 銷折銷退)
            {
                ExcelReport.GridViewToExcel(dataGridView9);
            }
            else if (tabControl1.SelectedTab == 進折進退)
            {
                ExcelReport.GridViewToExcel(dataGridView11);
            }
            else if (tabControl1.SelectedTab == 承銷品進貨總表進貨及進退折)
            {
                ExcelReport.GridViewToExcel(dataGridView12);
            }
            else if (tabControl1.SelectedTab == 承銷品銷貨總表)
            {
                ExcelReport.GridViewToExcel(dataGridView10);
            }

            //承銷品進貨總表進貨及進退折
        }



        private void textBox7_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            if (comboBox1.Text == "客戶")
            {
                LookupValues = GetMenu.GetCHIPCARD();
            }
            if (comboBox1.Text == "廠商")
            {
                LookupValues = GetMenu.GetCHIPCARD2();
            }

            if (LookupValues != null)
            {
                textBox7.Text = Convert.ToString(LookupValues[0]);
                //   textBox9.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox8_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            if (comboBox1.Text == "客戶")
            {
                LookupValues = GetMenu.GetCHIPCARD();
            }
            if (comboBox1.Text == "廠商")
            {
                LookupValues = GetMenu.GetCHIPCARD2();
            }
            if (LookupValues != null)
            {
                textBox8.Text = Convert.ToString(LookupValues[0]);
                //   textBox9.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox10_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetOitmGB();

            if (LookupValues != null)
            {
                textBox10.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox9_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetOitmGB();

            if (LookupValues != null)
            {
                textBox9.Text = Convert.ToString(LookupValues[0]);

            }
        }
        public string d;
        private void button4_Click(object sender, EventArgs e)
        {
            APS2CHOICE frm1 = new APS2CHOICE();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                CalcTotals1(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                CalcTotals1(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                CalcTotals1(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                CalcTotals1(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                CalcTotals1(dataGridView5);
            }
        }

        private void CalcTotals1(DataGridView dgv)
        {

            Int32 iTotal = 0;

            int i = dgv.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dgv.SelectedRows[iRecs].Cells["數量"].Value);
            }

            textBox11.Text = iTotal.ToString("#,##0");
        }
        public string K;
        private void button1_Click(object sender, EventArgs e)
        {
            APS3CHOICE frm1 = new APS3CHOICE();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

                K = frm1.q;
                dataGridView6.DataSource = GetCHO6();
            }
        }
        public string M;
        private void button2_Click(object sender, EventArgs e)
        {

            APS4CHOICE frm1 = new APS4CHOICE();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox3.Checked = true;
                M = frm1.M;

            }
        }

        private void GetExcelProduct(string ExcelFile, string NewFileName)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true ;

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

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
  
                string MANS;
                string YEARS;
                string MARK;
                string TYPE = "";
                int F1 = 0;
                for (int A1 = 5; A1 <= iColCnt; A1++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, A1]);
                    range.Select();
                    MARK = range.Text.ToString().Trim();
                    if (MARK == "總計")
                    {
                        F1 = A1;
                    }
                }
              
                for (int A2 = 5; A2 <= F1; A2++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, A2]);
                    range.Select();
                    YEARS = range.Text.ToString().Trim();

                    int K1 = 0;
                    int KS = 0;
                 
                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                        range.Select();
                        string AA = range.Text.ToString().Trim();

                        if (AA == "豬")
                        {
                            TYPE = "P";
                        }

                        if (AA == "雞")
                        {
                            TYPE = "C";
                        }
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                            range.Select();
                            MANS = range.Text.ToString().Trim();
                            System.Data.DataTable T1 = null;
                            string K2 = "";
                               if (A2 != F1)
                               {
                                   int F = YEARS.Length;
                                   string TYEAR = YEARS.Substring(0, 4);
                                   string TMON = YEARS.Substring(5, F - 5);
                                   T1 = GetCHO3ANDNOTCLOSE2(TYEAR, TMON, MANS, TYPE);
                                   K2 = T1.Rows[0][0].ToString();
                               }
                               else
                               {
                                   int KH = 0;
                                   for (int A2S = 5; A2S <= F1-1; A2S++)
                                   {
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2S]);
                                       range.Select();
                                       string G = range.Text.ToString().Trim().Replace(",", "");
                                       if (String.IsNullOrEmpty(G))
                                       {
                                           G = "0";
                                       }
                                       KH += Convert.ToInt16(G);
                                   }

                                   K2 = KH.ToString();
                               }

                            if (!String.IsNullOrEmpty(K2))
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2]);
                                range.Value2 = K2;
                                if (TYPE == "P")
                                {
                                    KS += Convert.ToInt16(K2);
                                }
                                if (TYPE == "C")
                                {
                                    K1 += Convert.ToInt16(K2);
                                }
                              
                            }
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2]);
                            if (AA == "雞 合計")
                            {
                                if (A2 == F1)
                                {
                                    range.Value2 = (K1 / 2).ToString();
                                }
                                else
                                {
                                    range.Value2 = K1.ToString();
                                }
                            }
                            if (AA == "豬 合計")
                            {
                              
                                if (A2 == F1)
                                {
                                    range.Value2 = (KS / 2).ToString();
                                }
                                else
                                {
                                    range.Value2 = KS.ToString();
                                }
                            }
                            if (AA == "總計")
                            {
                                if (A2 == F1)
                                {
                                    range.Value2 = ((KS + K1) / 3).ToString();
                                }
                                else
                                {
                                    range.Value2 = (KS + K1).ToString();
                                }
 
                            }


                   
                    }



                }

            }
            finally
            {

              
      

                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                System.Diagnostics.Process.Start(NewFileName);
            }

        }
        private void GetCHOICE(string ExcelFile, string NewFileName)
        {

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

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;

                string MANS;
                string YEARS;
                string MARK;
                int F1 = 0;
                for (int A1 = 1; A1 <= iRowCnt; A1++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[A1, 1]);
                    range.Select();
                    MARK = range.Text.ToString().Trim();

                    UPDATEPICKCHECK(MARK);
       
                }

   

            }
            finally
            {




                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                System.Diagnostics.Process.Start(NewFileName);
            }

        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {

                  string   FileNameS = openFileDialog1.FileName;

                  string AA = Path.GetDirectoryName(FileNameS) + "\\" +

               DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileNameS);
               GetExcelProduct(FileNameS, AA);

        
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {

                    string FileNameS = openFileDialog1.FileName;

                    string AA = Path.GetDirectoryName(FileNameS) + "\\" +

                 DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileNameS);
                    GetCHOICE(FileNameS, AA);


                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }
        public  void UPDATEPICKCHECK(string ID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE comCustomer SET PersonID='C0007' where ID=@ID    ");
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





        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                AddADOWNLOADt();
                string d = @"\\acmesrv01\Public\ARMAS\acmebg-ezcat\";

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {


                    FileInfo info = new FileInfo(file);
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");

                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");

                        string M2 = NAME.Substring(0, J1);
                       
                        try
                        {
                            string server2 = d + NAME;

                            AddADOWNLOAD2(M2, server2);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }


                    }
                }

                MessageBox.Show("上傳完成");
           

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void AddADOWNLOADt()
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("TRUNCATE TABLE GB_EZCAT ", connection);
            command.CommandType = CommandType.Text;



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
        public void AddADOWNLOAD2(string EZCAT, string EZPATH)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand(" Insert into GB_EZCAT(EZCAT,EZPATH) values(@EZCAT,@EZPATH)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@EZCAT", EZCAT));
            command.Parameters.Add(new SqlParameter("@EZPATH", EZPATH));

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
       

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("年", typeof(string));
            dt.Columns.Add("月", typeof(string));
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("客戶類別", typeof(string));
            dt.Columns.Add("類別名稱", typeof(string));
            dt.Columns.Add("來源", typeof(string));
            dt.Columns.Add("門市", typeof(string));
            dt.Columns.Add("B", typeof(string));
            dt.Columns.Add("品項", typeof(string));
            dt.Columns.Add("類別", typeof(string));
            dt.Columns.Add("零售", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("訂購憑單日期", typeof(string));
            dt.Columns.Add("訂購憑單週數", typeof(string));
            dt.Columns.Add("訂購憑單號碼", typeof(string));
            dt.Columns.Add("銷貨單據日期", typeof(string));
            dt.Columns.Add("銷貨單據週數", typeof(string));
            dt.Columns.Add("銷貨單據號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("發票品名", typeof(string));
            dt.Columns.Add("倉別", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單位", typeof(string));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("金額", typeof(decimal));
            dt.Columns.Add("稅", typeof(decimal));
            dt.Columns.Add("金額含稅", typeof(decimal));
            dt.Columns.Add("成本", typeof(decimal));
            dt.Columns.Add("毛利", typeof(decimal));
            dt.Columns.Add("分錄備註", typeof(string));
            dt.Columns.Add("快遞單號", typeof(string));
            dt.Columns.Add("下載", typeof(string));
            dt.Columns.Add("是否為贈品", typeof(string));
            dt.Columns.Add("細項描述", typeof(string));
            dt.Columns.Add("筆數", typeof(string));
            dt.Columns.Add("收款方式", typeof(string));
            dt.Columns.Add("天", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            return dt;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "下載")
                {
                    string sd = dataGridView1.CurrentRow.Cells["快遞單號"].Value.ToString();

                        System.Data.DataTable dt1 = GTODOWN(sd);

                 
                            for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                            {
                                DataRow drw = dt1.Rows[j];



                                System.Diagnostics.Process.Start(drw[0].ToString());

                            }

                            DataGridViewLinkCell cell =

                                (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                            cell.LinkVisited = true;
                     
                    
          
                }
     
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GTODOWN(string EZCAT)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT EZPATH FROM GB_EZCAT WHERE EZCAT LIKE '%" + EZCAT + "%' ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EZCAT", EZCAT));
  

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        public string MM;
        private void button10_Click(object sender, EventArgs e)
        {
            APS5CHOICE frm1 = new APS5CHOICE();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox5.Checked = true;
                MM = frm1.MM;

            }
        }
        public string MM2;
        private void button12_Click(object sender, EventArgs e)
        {
            APS6CHOICE frm1 = new APS6CHOICE();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox6.Checked = true;
                MM2 = frm1.MM2;

            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView5.DataSource = GetCHO5();
        }

 

    }
}

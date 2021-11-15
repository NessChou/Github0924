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
using System.Collections;

namespace ACME
{
    
    public partial class SACHOICE : Form
    {
     
        public SACHOICE()
        {
            InitializeComponent();
        }

    

        private void button6_Click(object sender, EventArgs e)
        {
            string strCn = "";

            if (textBox5.Text == "" && textBox6.Text == "" && textBox17.Text == "" && textBox18.Text == "" && textBox7.Text == "" && !checkBox2.Checked )
            {
                MessageBox.Show("請輸入條件");
                return;
            }


            if (radioButton1.Checked)
            {
                 strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                 strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
    
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            System.Data.DataTable t1 = GetCHO3(strCn);
            System.Data.DataTable t1T = GetCHO3T(strCn);
            if (t1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }


            dataGridView1.DataSource = t1;
            dataGridView7.DataSource = t1T;

            for (int i = 14; i <= 18; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }

        }
        public System.Data.DataTable GetCHO1(string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select   T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱  ,T0.ProjectID 專案代碼,AD2.ProjectName 專案名稱,T0.BillDate 訂購憑單日期 ");
            sb.Append(" , T1.PREINDATE 預交貨日,''''+T0.BillNO 訂購憑單號碼,T0.CustBillNo 客戶訂單 ,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名         ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅         ");
            sb.Append(" ,T1.Amount+T1.TaxAmt '金額(含稅)' ,P.PersonName 業務 from ordBillMain T0         ");
            sb.Append(" left join ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" LEFT JOIN comProject  AD2 ON (T0.ProjectID=AD2.ProjectID )   ");
            sb.Append(" left join comPerson P ON (T0.Salesman=P.PersonID)   ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0   ");
            sb.Append(" and year(cast(cast(T0.BillDate as varchar) as datetime))>2015  ");


            if (textBox12.Text != "")
            {
                sb.Append("          AND  T0.CustomerID=@CC  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CC", textBox12.Text));
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

        public System.Data.DataTable GetCHO2(string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select   T0.CustomerID  供應商代碼,U.ShortName 供應商簡稱  ,T0.ProjectID 專案代碼,AD2.ProjectName 專案名稱,");
            sb.Append(" T0.BillDate 採購憑單日期, T1.PREINDATE 預進貨日,''''+T0.BillNO 採購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名        ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅        ");
            sb.Append(" ,T1.Amount+T1.TaxAmt '金額(含稅)'  from ordBillMain T0        ");
            sb.Append(" left join ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)          ");
            sb.Append(" left join comCustomer U On  U.ID=T0.CustomerID AND U.Flag =2         ");
            sb.Append(" Left Join comProduct J On T1.ProdID =J.ProdID       ");
            sb.Append(" LEFT JOIN comProject  AD2 ON (T0.ProjectID=AD2.ProjectID )  ");
            sb.Append(" WHERE  T0.Flag =4 AND T0.BillStatus = 0  and QtyRemain  > 0  and year(cast(cast(T0.BillDate as varchar) as datetime))>2015 ");

            if (textBox13.Text != "")
            {
                sb.Append("AND  T0.CustomerID=@CC  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CC", textBox13.Text));
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
        public System.Data.DataTable GetCHOF()
        {

            SqlConnection MyConnection = new SqlConnection("Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock");
            StringBuilder sb = new StringBuilder();
            sb.Append(" select * from (                   ");
            sb.Append(" select 'TOP' OBU, T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.BillDate 訂購憑單日期,''''+T0.BillNO 訂購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名         ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅         ");
            sb.Append(" ,T1.Amount+T1.TaxAmt '金額(含稅)', ");
            sb.Append(" cast((select sum(Quantity) QTY from CHIComp20.DBO.StkYearMonthQty where ProdID =T1.ProdID AND yearmonth<substring(CONVERT(VARCHAR(10) , DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime)) , 112 ),1,6) ) as int) 立單45天後庫存數量             ");
            sb.Append(" from CHIComp20.DBO.ordBillMain T0         ");
            sb.Append(" left join CHIComp20.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join CHIComp20.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join CHIComp20.DBO.comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0  and year(cast(cast(T0.BillDate as varchar) as datetime))>2015  AND  DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime))  < GETDATE() ");
            sb.Append(" UNION ALL  select 'CHOICE' OBU, T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.BillDate 訂購憑單日期,''''+T0.BillNO 訂購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名         ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅         ");
            sb.Append(" ,T1.Amount+T1.TaxAmt '金額(含稅)', ");
            sb.Append(" cast((select sum(Quantity) QTY from CHIComp21.DBO.StkYearMonthQty where ProdID =T1.ProdID AND yearmonth<substring(CONVERT(VARCHAR(10) , DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime)) , 112 ),1,6) ) as int) 立單45天後庫存數量              ");
            sb.Append(" from CHIComp21.DBO.ordBillMain T0         ");
            sb.Append(" left join CHIComp21.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join CHIComp21.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join CHIComp21.DBO.comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0  and year(cast(cast(T0.BillDate as varchar) as datetime))>2015  AND  DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime))  < GETDATE() ");
            sb.Append(" UNION ALL  select 'INFINITE' OBU, T0.CustomerID  客戶代碼,U.ShortName 客戶簡稱,T0.BillDate 訂購憑單日期,''''+T0.BillNO 訂購憑單號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名         ");
            sb.Append(" ,T1.Quantity 數量,T1.Price 單價,T1.Amount '金額(未稅)',T1.TaxAmt 稅         ");
            sb.Append(" ,T1.Amount+T1.TaxAmt '金額(含稅)', ");
            sb.Append(" cast((select sum(Quantity) QTY from CHIComp22.DBO.StkYearMonthQty where ProdID =T1.ProdID AND yearmonth<substring(CONVERT(VARCHAR(10) , DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime)) , 112 ),1,6) ) as int) 立單45天後庫存數量            ");
            sb.Append(" from CHIComp22.DBO.ordBillMain T0         ");
            sb.Append(" left join CHIComp22.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)           ");
            sb.Append(" left join CHIComp22.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1          ");
            sb.Append(" Left Join CHIComp22.DBO.comProduct J On T1.ProdID =J.ProdID        ");
            sb.Append(" WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0  and year(cast(cast(T0.BillDate as varchar) as datetime))>2015  AND  DATEADD(D,45,cast(substring(T0.BillNO,1,8) as datetime))  < GETDATE() ");
            sb.Append(" ) as a where 立單45天後庫存數量 > 數量 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox7.Text));
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
        public System.Data.DataTable GetCHO3(string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append("                 Select A.BillDate 日期,O.PreInDate 預交貨日 ");
            sb.Append("   ,O2.ProjectID 專案代碼,AD2.ProjectName 專案名稱");
            sb.Append("   ,''''+A.FromNO 訂購單號,''''+A.BillNO 銷貨單號,O2.CustBillNo 客戶訂單, U.FULLNAME 客戶名稱, ");
            sb.Append("                                   A.ProdID 產品編號,A.ProdName 品名規格,C.CurrencyName+''+ cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價, ");
            sb.Append("                W.WareHouseName 倉別,P.PersonName 業務,AD.MEMO 送貨地址,CAST(A.Quantity AS INT) 數量 ,O.Amount '美金銷售總額(未稅)',O.Amount+O.TaxAmt '美金銷售總額(含稅)',A.MLAmount '台幣銷售總額(未稅)',A.MLAmount+A.TaxAmt '台幣銷售總額(含稅)',S.Remark 備註 ");
            sb.Append("                       From ComProdRec A  ");
            sb.Append("                        Left Join comWareHouse D On D.WareHouseID=A.WareID ");
            sb.Append("                       left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag ");
            sb.Append("                            left join comCustomer U On  U.ID=T.CustID AND U.Flag =1 ");
            sb.Append("                                   left join comWareHouse W On  A.WareID=W.WareHouseID ");
            sb.Append("                               left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2  ");
            sb.Append("                                    left join OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2  ");
            sb.Append("                               left join comCurrencySys C On  O2.CurrID=C.CurrencyID  ");
            sb.Append("                                  left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500) ");
            sb.Append("                                             left join comPerson P ON (S.Salesman=P.PersonID) ");
            sb.Append("                 LEFT JOIN comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID ) ");
            sb.Append("                 LEFT JOIN comProject  AD2 ON (O2.ProjectID=AD2.ProjectID ) ");
            sb.Append("                            Where A.Flag=500   ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {
                sb.Append(" and   A.BillDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
            }

            if (textBox7.Text != "")
            {
                sb.Append("          AND T.CustID=@CC  ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox17.Text != "" && textBox18.Text != "")
                {
                    sb.Append(" and   A.ProdID between '" + textBox18.Text.ToString() + "' and '" + textBox17.Text.ToString() + "' ");
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CC", textBox7.Text));
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

        public System.Data.DataTable GetCHO3T(string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append("                              Select T.CustID 客戶編號,");
            sb.Append("          U.FULLNAME 客戶名稱,    SUM(CAST(A.Quantity AS INT)) 數量 ,SUM(O.Amount) '美金銷售總額(未稅)',SUM(O.Amount+O.TaxAmt) '美金銷售總額(含稅)',SUM(A.MLAmount) '台幣銷售總額(未稅)',SUM(A.MLAmount+A.TaxAmt) '台幣銷售總額(含稅)'");
            sb.Append("                                    From ComProdRec A   ");
            sb.Append("                                     Left Join comWareHouse D On D.WareHouseID=A.WareID  ");
            sb.Append("                                    left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag  ");
            sb.Append("                                         left join comCustomer U On  U.ID=T.CustID AND U.Flag =1  ");
            sb.Append("                                                left join comWareHouse W On  A.WareID=W.WareHouseID  ");
            sb.Append("                                            left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2   ");
            sb.Append("                                                 left join OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2   ");
            sb.Append("                                            left join comCurrencySys C On  O2.CurrID=C.CurrencyID   ");
            sb.Append("                                               left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =500)  ");
            sb.Append("                                                          left join comPerson P ON (S.Salesman=P.PersonID)  ");
            sb.Append("                              LEFT JOIN comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )  ");
            sb.Append("                              LEFT JOIN comProject  AD2 ON (O2.ProjectID=AD2.ProjectID )  ");
            sb.Append("                                         Where A.Flag=500   ");


            if (textBox5.Text != "" && textBox6.Text != "")
            {
                sb.Append(" and   A.BillDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
            }

            if (textBox7.Text != "")
            {
                sb.Append("          AND T.CustID=@CC  ");
            }

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox17.Text != "" && textBox18.Text != "")
                {
                    sb.Append(" and   A.ProdID between '" + textBox18.Text.ToString() + "' and '" + textBox17.Text.ToString() + "' ");
                }
            }
            sb.Append("    GROUP BY     T.CustID,      U.FULLNAME ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CC", textBox7.Text));
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
        public System.Data.DataTable GetCHO4(string strCn1, string COMPANY)
        {
            string username = fmLogin.LoginID.ToString().ToUpper();
            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            if (COMPANY == "AD")
            {
                sb.Append("            Select A.BillDate 日期,O.PreInDate 預交貨日  ,O2.ProjectID 專案代碼,AD2.ProjectName 專案名稱,''''+A.FromNO 採購單號,''''+A.BillNO 進貨單號,  T.UDef1 原廠Invoice,U.FULLNAME 供應商名稱, ");
                sb.Append("                       A.ProdID 產品編號,A.ProdName 品名規格, ");
                if (username == "KIKILEE")
                {
                    sb.Append("  C.CurrencyName+''+ cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價, ");
                }
                sb.Append("                W.WareHouseName 倉別,P.PersonName 業務,CAST(A.Quantity AS INT) 數量 ,  ");
                if (username == "KIKILEE")
                {
                    sb.Append("              O.Amount+O.TaxAmt 美金採購總額,A.MLAmount+A.TaxAmt 台幣採購總額, ");
                }
                sb.Append("       S.Remark 備註  ");
                sb.Append("                       From ComProdRec A  ");
                sb.Append("                        Left Join comWareHouse D On D.WareHouseID=A.WareID ");
                sb.Append("                       left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag ");
                sb.Append("                            left join comCustomer U On  U.ID=T.CustID AND U.Flag =2 ");
                sb.Append("                                   left join comWareHouse W On  A.WareID=W.WareHouseID ");
                sb.Append("                               left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4 ");
                sb.Append("                                    left join OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4  ");
                sb.Append("                               left join comCurrencySys C On  O2.CurrID=C.CurrencyID ");
                sb.Append("                                  left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =100) ");
                sb.Append("                                             left join comPerson P ON (S.Salesman=P.PersonID) ");
                sb.Append("                 LEFT JOIN comProject  AD2 ON (O2.ProjectID=AD2.ProjectID ) ");
                sb.Append("                            Where A.Flag=100");

                if (textBox1.Text != "" && textBox3.Text != "")
                {
                    sb.Append(" and   A.BillDate between '" + textBox1.Text.ToString() + "' and '" + textBox3.Text.ToString() + "' ");
                }

                if (textBox4.Text != "")
                {
                    sb.Append("          AND T.CustID=@CC  ");
                }
                if (checkBox1.Checked)
                {
                    sb.Append(" and   A.ProdID in ( " + d + ") ");
                }
                else
                {
                    if (textBox20.Text != "" && textBox19.Text != "")
                    {
                        sb.Append(" and   A.ProdID between '" + textBox20.Text.ToString() + "' and '" + textBox19.Text.ToString() + "' ");
                    }
                }
            }
            else
            {
                sb.Append("            Select A.BillDate 日期,O.PreInDate 預交貨日,''''+A.FromNO 採購單號,A.BillNO 進貨單號, U.FULLNAME 供應商名稱, ");
                sb.Append("                       A.ProdID 產品編號,A.ProdName 品名規格, ");
                sb.Append("  C.CurrencyName+''+ cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價, ");
                sb.Append("                W.WareHouseName 倉別,P.PersonName 業務,CAST(A.Quantity AS INT) 數量 ,  ");
                sb.Append("              O.Amount+O.TaxAmt 美金採購總額,A.MLAmount+A.TaxAmt 台幣採購總額, ");
                sb.Append("        S.Remark 備註,T.UDef1 原廠Invoice  ");
                sb.Append("                       From ComProdRec A  ");
                sb.Append("                        Left Join comWareHouse D On D.WareHouseID=A.WareID ");
                sb.Append("                       left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag ");
                sb.Append("                            left join comCustomer U On  U.ID=T.CustID AND U.Flag =2 ");
                sb.Append("                                   left join comWareHouse W On  A.WareID=W.WareHouseID ");
                sb.Append("                               left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4 ");
                sb.Append("                                    left join OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4  ");
                sb.Append("                               left join comCurrencySys C On  O2.CurrID=C.CurrencyID ");
                sb.Append("                                  left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =100) ");
                sb.Append("                                             left join comPerson P ON (S.Salesman=P.PersonID) ");
                sb.Append("                            Where A.Flag=100");

                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and   A.BillDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
                }
                if (textBox4.Text != "")
                {
                    sb.Append("          AND T.CustID=@CC  ");
                }

                if (checkBox1.Checked)
                {
                    sb.Append(" and   A.ProdID in ( " + d + ") ");
                }
                else
                {
                    if (textBox20.Text != "" && textBox19.Text != "")
                    {
                        sb.Append(" and   A.ProdID between '" + textBox20.Text.ToString() + "' and '" + textBox19.Text.ToString() + "' ");
                    }
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox4.Text));
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

        public System.Data.DataTable GetBRROW(string strCn1)
        {

            SqlConnection MyConnection = new SqlConnection(strCn1);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.CustomerID  客戶代碼,T0.CustomerName  客戶名稱,T0.BorrowDate  借出日期,''''+T0.BorrowNO 借出號碼,T1.ProdID 產品編號,T1.ProdName 品名規格,J.InvoProdName 發票品名        ");
            sb.Append(" ,T1.Quantity 數量,QtyRemain 未還數量,T2.StyleName 借出類別, T0.Remark 備註  from StkBorrowMain T0        ");
            sb.Append(" left join StkBorrowSub T1 ON (T0.Flag =T1.Flag AND T0.BorrowNO=T1.BorrowNO)           ");
            sb.Append(" Left Join comProduct J On T1.ProdID =J.ProdID       ");
            sb.Append(" LEFT JOIN stkBorrowStyle T2 ON (T0.BorrowStyle=T2.StyleID) ");
            sb.Append(" WHERE year(cast(cast(T0.BorrowDate as varchar) as datetime))>2015");
            sb.Append(" AND T0.Flag IN  (10) AND QtyRemain > 0 ");
            sb.Append(" AND T0.BorrowDate BETWEEN @AA AND @BB   ");
            if (textBox15.Text != "")
            {
                sb.Append(" AND T0.CustomerID=@CC  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox14.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox15.Text));
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
        string strCn = "";
        private void SACHOICE_Load(object sender, EventArgs e)
        {
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();

            textBox1.Text = GetMenu.DFirst();
            textBox3.Text = GetMenu.DLast();

            textBox9.Text = GetMenu.DFirst();
            textBox14.Text = GetMenu.DLast();

            txbYear.Text = DateTime.Now.ToString("yyyy");

        }

        private void button8_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            if (radioButton1.Checked)
            {
                LookupValues = GetMenu.GetCHICUST();
            }

            if (radioButton2.Checked)
            {
                LookupValues = GetMenu.GetCHICUST12();
            }
            if (radioButton3.Checked)
            {
                LookupValues = GetMenu.GetCHICUST13();
            }
            if (radioButton4.Checked)
            {
                LookupValues = GetMenu.GetCHICUST14();
            }

                if (LookupValues != null)
                {
                    textBox7.Text = Convert.ToString(LookupValues[0]);
                    textBox8.Text = Convert.ToString(LookupValues[1]);
                }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strCn = "";
            string COMPANY = "";
            if (textBox5.Text == "" && textBox6.Text == "" && textBox20.Text == "" && textBox19.Text == "" && textBox4.Text == ""   && !checkBox1.Checked)
            {
                MessageBox.Show("請輸入條件");
                return;
            }

  
            if (radioButton1.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                COMPANY = "AD";
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            string username = fmLogin.LoginID.ToString().ToUpper();
            System.Data.DataTable t1 = GetCHO4(strCn, COMPANY);
            if (t1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }

            dataGridView2.DataSource = t1;
            if ( COMPANY == "AD")
            {

            }
            else
            {
              

                for (int i = 9; i <= 11; i++)
                {
                    DataGridViewColumn col = dataGridView2.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;


            if (radioButton1.Checked)
            {
                LookupValues = GetMenu.GetCHICUST2();
            }

            if (radioButton2.Checked)
            {
                LookupValues = GetMenu.GetCHICUST222();
            }
            if (radioButton3.Checked)
            {
                LookupValues = GetMenu.GetCHICUST223();
            }
            if (radioButton4.Checked)
            {
                LookupValues = GetMenu.GetCHICUST224();
            }
            if (LookupValues != null)
            {
                textBox4.Text = Convert.ToString(LookupValues[0]);
                textBox2.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string strCn = "";
            if (radioButton1.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            System.Data.DataTable t1 = GetCHO1(strCn);
            if (t1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }


            dataGridView3.DataSource = t1;

            for (int i = 12; i <= 16; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            if (radioButton1.Checked)
            {
                LookupValues = GetMenu.GetCHICUST();
            }

            if (radioButton2.Checked)
            {
                LookupValues = GetMenu.GetCHICUST12();
            }
            if (radioButton3.Checked)
            {
                LookupValues = GetMenu.GetCHICUST13();
            }
            if (radioButton4.Checked)
            {
                LookupValues = GetMenu.GetCHICUST14();
            }

            if (LookupValues != null)
            {
                textBox12.Text = Convert.ToString(LookupValues[0]);
                textBox10.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string strCn = "";
            if (radioButton1.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            System.Data.DataTable t1 = GetCHO2(strCn);
            if (t1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }


            dataGridView4.DataSource = t1;

            for (int i = 10; i <= 14; i++)
            {
                DataGridViewColumn col = dataGridView4.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView3);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView4);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;


            if (radioButton1.Checked)
            {
                LookupValues = GetMenu.GetCHICUST2();
            }

            if (radioButton2.Checked)
            {
                LookupValues = GetMenu.GetCHICUST222();
            }
            if (radioButton3.Checked)
            {
                LookupValues = GetMenu.GetCHICUST223();
            }
            if (radioButton4.Checked)
            {
                LookupValues = GetMenu.GetCHICUST224();
            }
            if (LookupValues != null)
            {
                textBox13.Text = Convert.ToString(LookupValues[0]);
                textBox16.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = GetCHOF();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView5);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string strCn = "";
            if (radioButton1.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            System.Data.DataTable t1 = GetBRROW(strCn);
            if (t1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }


            dataGridView6.DataSource = t1;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            if (radioButton1.Checked)
            {
                LookupValues = GetMenu.GetCHICUST();
            }

            if (radioButton2.Checked)
            {
                LookupValues = GetMenu.GetCHICUST12();
            }
            if (radioButton3.Checked)
            {
                LookupValues = GetMenu.GetCHICUST13();
            }
            if (radioButton4.Checked)
            {
                LookupValues = GetMenu.GetCHICUST14();
            }

            if (LookupValues != null)
            {
                textBox15.Text = Convert.ToString(LookupValues[0]);
                textBox11.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView6);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string AccGroup = tabControl1.SelectedIndex.ToString();

            if (AccGroup == "5")

            {
                radioButton4.Checked = true;
            }
        }
        public string d;
        private void button18_Click(object sender, EventArgs e)
        {
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox1.Checked = true;
                d = frm1.q;

            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            string strCn = "";
            string COMPANY = "";


            if (textBox5.Text == "" && textBox6.Text == "" && textBox17.Text == "" && textBox18.Text == "" && textBox7.Text == "" && !checkBox2.Checked)
            {
                MessageBox.Show("請輸入條件");
                return;
            }


            if (radioButton1.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (radioButton2.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton3.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (radioButton4.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                COMPANY = "AD";
            }
            if (radioButton5.Checked)
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string DOCYEAR = txbYear.Text;
            string FileName = lsAppDir + "\\Excel\\wh\\單據筆數.xlsx";
            string OutPutFile = "";
            System.Data.DataTable dt = new System.Data.DataTable();

            if (cmbType.Text == "單據數量")
            {
                OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + cmbType.Text + ".xlsx";
                dt = GGY1(strCn, DOCYEAR, OutPutFile);
            }
            else
            {
                OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + cmbType.Text + ".xlsx";
                dt = GGY2(strCn, DOCYEAR, OutPutFile);
            }
            


            MakeExcel(FileName, OutPutFile, dt);


            System.Diagnostics.Process.Start(OutPutFile);
        }
        private System.Data.DataTable GGY1(string strCn ,string DOCYEAR,string OutPutFile)
        {
            System.Data.DataTable dt;
            SqlConnection MyConnection = new SqlConnection(strCn); 
            StringBuilder sb = new StringBuilder();
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'01%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'01%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'02%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'02%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'03%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'03%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'04%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'04%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'05%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'05%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'06%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'06%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'07%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'07%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'08%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'08%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'09%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'09%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'10%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'10%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'11%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'11%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=500 ) and   A.BillDate LIKE  @DOCYEAR" + "+'12%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select distinct(count(a.BillNO))  From ComProdRec A  Where ( A.Flag=100 ) and   A.BillDate LIKE  @DOCYEAR" + "+'12%'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCYEAR", DOCYEAR));


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
        private System.Data.DataTable GGY2(string strCn, string DOCYEAR, string OutPutFile)
        {
            System.Data.DataTable dt;
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'01%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'01%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'02%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'02%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'03%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'03%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'04%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'04%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'05%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'05%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'06%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'06%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'07%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'07%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'08%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'08%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'09%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'09%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'10%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'10%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'11%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'11%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=500 ) and  A.BillDate LIKE  @DOCYEAR" + "+'12%'");
            sb.Append(" UNION ALL ");
            sb.Append("  Select SUM(MLAmount) From ComProdRec A  Where ( A.Flag=100 ) and  A.BillDate LIKE  @DOCYEAR" + "+'12%'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCYEAR", DOCYEAR));


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
        private void MakeExcel(string ExcelFile,string OutPutFile,System.Data.DataTable dt) 
        {
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;

            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);
            Microsoft.Office.Interop.Excel.Range range = null;
            Microsoft.Office.Interop.Excel.Range range2 = null;



            //Open the worksheet file


            try
            {
                object SelectCell = "A1";
                range = excelSheet.get_Range("A1", "M10");
                
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                range.Select();
                range.Value2 = txbYear.Text + "年單據筆數";

                iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                ht = new Hashtable(iRowCnt);

                range2 = null;

                SelectCell = "A1";
                range2 = excelSheet.get_Range(SelectCell, SelectCell);
               
                
                for (int i = 0; i < 12; i++)
                {
                    string monthcode = "";
                    //聯倉CardCode U0193
                    if (i.ToString().Length == 1)
                    {
                        monthcode = "0" + (i + 1).ToString();
                    }
                    else
                    {
                        monthcode = (i + 1).ToString();
                    }

                    //聯倉i月理貨
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 2]);
                    range.Select();
                    range.Value2 = dt.Rows[i * 2][0].ToString() == "" ? "0" : dt.Rows[i * 2][0].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 2]);
                    range.Select();
                    range.Value2 = dt.Rows[i * 2 + 1][0].ToString() == "" ? "0" : dt.Rows[i * 2 + 1][0].ToString();

                }

            }
            catch (Exception ex)
            {

            }
            finally
            {

                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);
            }

        }
    }
}

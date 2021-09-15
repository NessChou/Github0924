using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.SqlClient;
using System.IO;

namespace ACME
{
    public partial class GB_GETMONEY : Form
    {
        public string c;
        System.Data.DataTable DTC = null;
        string CONN = "";
        string strCn2 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn3 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_GETMONEY()
        {
            InitializeComponent();
        }

        private void GB_GETMONEY_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "聿豐";
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();
            EXEC();

            for (int i = 6; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];

                if (i == 6 || i == 10 || i == 12 || i == 14)
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0";
                 
                }
            }

            dataGridView2.DataSource = GetOrderData4();
            for (int i = 2; i <= dataGridView2.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];

             
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        col.DefaultCellStyle.Format = "#,##0.0000";
                    

                
            }
            dataGridView3.DataSource = GetOrderData5();
            for (int i = 6; i <= dataGridView3.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];

                if (i == 6 || i == 9 || i == 11 || i == 13)
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    if (i == 6)
                    {
                        col.DefaultCellStyle.Format = "#,##0";
                    }
                    else
                    {
                        col.DefaultCellStyle.Format = "#,##0.0000";
                    }

                }
            }
            dataGridView4.DataSource = GetOrderData6();
        }
        private System.Data.DataTable GetOrderData3()
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            if (rbAr1.Checked)
            {
                sb.Append(" Select       ''''+A.CustID 客戶編號,       ");
                sb.Append(" B.ShortName  客戶名稱,       ");
                sb.Append(" A.DueTo  帳款歸屬編號,       ");
                sb.Append(" F.ShortName As 帳款歸屬,    ");
                sb.Append(" CONVERT(VARCHAR(10) , cast(CAST(A.BillDate AS VARCHAR) as datetime),111)    日期, ");
                sb.Append(" CASE WHEN A.Flag =698 THEN CONVERT(VARCHAR(10) ,CHICOMP02.dbo.fun_CreditDate(CASE C.RecvWay WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN  '月結'   WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),C.DistDays ), 111 )  ELSE  ");
                sb.Append(" CONVERT(VARCHAR(10) ,CHICOMP02.dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN CASE WHEN M.GatherDelay =45 THEN  '月結45' ELSE '月結' END  WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay), 111 ) END  逾期日期       ");
                sb.Append(" ,				 CASE WHEN A.Flag =698 THEN datediff(day,       ");
                sb.Append(" CHICOMP02.dbo.fun_CreditDate(CASE C.RecvWay WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN  '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),C.DistDays )        ");
                sb.Append(" , GETDATE()) ELSE datediff(day,       ");
                sb.Append(" CHICOMP02.dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN  '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)        ");
                sb.Append(" , GETDATE())  END 逾期天數       ");
                sb.Append(" ,     A.FundBillNO 單號,       ");
                sb.Append(" A.InvoiceNO 發票號碼,         ");
                sb.Append(" case A.Flag when 500 then a.Total+A.Tax  when 595 then a.Total+A.Tax  when 600 then  -(A.Total+A.Tax) when 698 then  -(A.Total+A.Tax) end 應收金額,        ");
                sb.Append(" ''''+ A.VoucherNo AS 傳票編號,        ");
                sb.Append(" A.CashPay+A.VisaPay+A.OtherPay+ A.OffSet as 沖款金額,       ");
                if (comboBox1.Text == "聿豐")
                {
                    sb.Append(" CASE WHEN A.Flag =698 THEN CASE  C.RecvWay WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(C.DistDays AS VARCHAR)+'天' WHEN 3 THEN  B.GatherOther END ELSE  ");
                    sb.Append(" CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(M.GatherDelay AS VARCHAR)+'天' WHEN 3 THEN M.GatherOther END END  付款方式,       ");
                }

                if (comboBox1.Text == "東門店")
                {
                    sb.Append(" CASE C.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END  付款方式,  ");
                }
                sb.Append(" CAST((case A.Flag when 500 then a.Total+A.Tax  when 595 then a.Total+A.Tax  when 600 then  -(A.Total+A.Tax) when 698 then  -(A.Total+A.Tax) end-A.CashPay-A.VisaPay-A.OtherPay- A.OffSet) AS DECIMAL(12,4)) as 結餘,       ");
                sb.Append(" Case When  substring(B.ShortName,CHARINDEX('-',B.ShortName)+1,1) IN ('進','聿','博','能') then  substring(B.ShortName,0,CHARINDEX('-',B.ShortName)) end as 員工,E.PersonName  業務,A.DEPTID 部門       ");
                sb.Append(" From comBillAccounts A        ");
                sb.Append(" Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO        ");
                sb.Append(" Left Join      comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag      ");
                sb.Append(" Left Join      comCustomer F On  A.DueTo = F.ID AND A.CustFlag = F.Flag    ");
                sb.Append(" Left Join      comCustTrade  C On A.CustID=C.ID And A.CustFlag=C.Flag    ");
                sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan         ");
                sb.Append(" Where     A.Flag <> 298 And A.HasCheck = 1 And A.CustFlag =1 AND A.CurrID ='NTD'   ");
                sb.Append(" And (A.Status IN (1, 3))    ");
                sb.Append(" and ((Case When A.Flag IN(297,697,200,600,210,201,698) Then -(A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)  Else (A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)    ");
                sb.Append(" End - (A.Offset+A.NoCheckOffSet + A.Discount + A.NoCheckDisCount))<>0)  And A.YearCompressType <> 1   ");
                sb.Append(" AND  A.BILLDATE BETWEEN @CreateDate AND @CreateDate1 ");
                if (checkBox4.Checked)
                {
                    sb.Append(" and  A.[CustID] in ( " + c + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  A.[CustID] between @CustID1 and @CustID2 ");
                    }
                }
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  A.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox1.Text != "" )
                {
                    sb.Append(" and  A.DueTo  =@DUETO ");
                }
                sb.Append("      Order By A.CustID , A.BillDate, A.FundBillNO   ");

            }



            if (rbAr2.Checked)
            {
                sb.Append(" Select   ");
                sb.Append(" ''''+A.CustID 客戶編號,  ");
                sb.Append(" B.ShortName  客戶名稱,  ");
                sb.Append(" A.DueTo  帳款歸屬編號,  ");
                sb.Append(" C.ShortName As 帳款歸屬,  ");
                sb.Append(" CONVERT(VARCHAR(10) , cast(CAST(A.BillDate AS VARCHAR) as datetime),111)    日期,    ");
                sb.Append(" A.FundBillNO 單號,  ");
                sb.Append(" A.InvoiceNO 發票號碼,    ");
                sb.Append(" A.Total+A.Tax 應收金額,   ");
                sb.Append(" ''''+A.VoucherNo AS 傳票編號,   ");
                sb.Append(" A.CashPay+A.VisaPay+A.OtherPay as 沖款金額,");
                if (comboBox1.Text == "聿豐")
                {
                    sb.Append(" CASE WHEN A.Flag =698 THEN CASE F.RecvWay  WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(F.DistDays AS VARCHAR)+'天' WHEN 3 THEN B.GatherOther END ELSE  CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(M.GatherDelay AS VARCHAR)+'天' WHEN 3 THEN M.GatherOther END END 付款方式,  ");
                }
                if (comboBox1.Text == "東門店")
                {
                    sb.Append(" CASE F.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END  付款方式,  ");
                }
                sb.Append(" (A.Total+A.Tax-A.CashPay-A.VisaPay-A.OtherPay) as 結餘,  ");
                sb.Append(" Case When  substring(B.ShortName,CHARINDEX('-',B.ShortName)+1,1) IN ('進','聿','博','能') then  substring(B.ShortName,0,CHARINDEX('-',B.ShortName)) end as 員工,E.PersonName  業務,A.DEPTID 部門   ");
                sb.Append(" FROM comBillAccounts A         ");
                sb.Append(" Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO  ");
                sb.Append(" Left Join (ComFundSub D         ");
                sb.Append(" Inner JOIN comFundMain H ON D.Flag = H.Flag AND D.FundBillNo = H.FundBillID And H.YearCompressType <> 1)         ");
                sb.Append(" ON A.Flag = D.OriginFlag AND A.FundBillNo = D.OriginBillNo        ");
                sb.Append(" Left Join ComCustomer B ON A.CustID = B.ID AND A.CustFlag = B.Flag         ");
                sb.Append(" Left JOIN ComCustomer C ON A.DueTo = C.ID AND A.CustFlag = C.Flag    ");
                sb.Append(" Left Join      comCustTrade  F On A.CustID=F.ID And A.CustFlag=F.Flag        ");
                sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan    ");
                sb.Append(" Left Join impAAMain V ON A.Flag = V.Flag And A.FundBillNO = V.AANO WHERE A.CustFlag = 1   ");
                sb.Append(" AND A.Status IN (1, 3)   ");
                sb.Append(" And A.YearCompressType <> 1 AND D.AccFlag = 1  And A.HasCheck=1         ");
                sb.Append(" AND ((A.Offset + A.Discount)<>0    ");
                sb.Append(" Or (Exists (Select J.OriginBillNO From ComFundSub J   ");
                sb.Append(" Where A.Flag = J.OriginFlag  And A.FundBillNO = J.OriginBillNO   ");
                sb.Append(" And J.OriginFlag <> 0 And Left(J.OriginBillNO,1) <> '*') ))   ");
                sb.Append(" AND  A.BILLDATE BETWEEN @CreateDate AND @CreateDate1  ");
                if (checkBox4.Checked)
                {
                    sb.Append(" and  A.[CustID] in ( " + c + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  A.[CustID] between @CustID1 and @CustID2 ");
                    }
                }
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  A.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox1.Text != "")
                {
                    sb.Append(" and  A.DueTo  =@DUETO ");
                }
                sb.Append(" Order By A.DueTo , A.BillDate, A.FundBillNO");

            }





            if (rbAr3.Checked)
            {



                //全部

                sb.Append(" Select  ");
                sb.Append(" ''''+A.CustID 客戶編號, ");
                sb.Append(" B.ShortName  客戶名稱, ");
                sb.Append(" A.DueTo  帳款歸屬編號, ");
                sb.Append(" C.ShortName As 帳款歸屬, ");
                sb.Append(" CONVERT(VARCHAR(10) , cast(CAST(A.BillDate AS VARCHAR) as datetime),111)    日期,   ");
                sb.Append(" A.FundBillNO 單號, ");
                sb.Append(" A.InvoiceNO 發票號碼,   ");
                sb.Append(" A.Total+A.Tax 應收金額,  ");
                sb.Append(" ''''+A.VoucherNo AS 傳票編號,  ");
                sb.Append(" A.CashPay+A.VisaPay+A.OtherPay as 沖款金額, ");
                if (comboBox1.Text == "聿豐")
                {
                    sb.Append(" CASE WHEN A.Flag =698 THEN CASE  F.RecvWay WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(F.DistDays AS VARCHAR)+'天' WHEN 3 THEN  B.GatherOther END ELSE  ");
                    sb.Append(" CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(M.GatherDelay AS VARCHAR)+'天' WHEN 3 THEN M.GatherOther END END  付款方式,    ");
                }
                if (comboBox1.Text == "東門店")
                {
                    sb.Append(" CASE F.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END  付款方式,  ");
                }
                sb.Append(" (A.Total+A.Tax-A.CashPay-A.VisaPay-A.OtherPay) as 結餘, ");
                sb.Append(" Case When  substring(B.ShortName,CHARINDEX('-',B.ShortName)+1,1) IN ('進','聿','博','能') then  substring(B.ShortName,0,CHARINDEX('-',B.ShortName)) end as 員工,E.PersonName  業務,A.DEPTID 部門  ");
                sb.Append(" From comBillAccounts A  ");
                sb.Append(" Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO  ");
                sb.Append(" Left Join      comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag  ");
                sb.Append(" Left Join      comCustomer C On A.DueTo= C.ID And A.CustFlag=C.Flag      Left Join impAAMain V On A.Flag = V.Flag And A.Fundbillno = V.AANO  ");
                sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan   ");
                sb.Append(" Left Join      comCustTrade  F On A.CustID=F.ID And A.CustFlag=F.Flag    ");
                sb.Append(" Where     A.Flag <> 298 And A.HasCheck = 1 And A.CustFlag =1 AND A.CurrID ='NTD' ");
                sb.Append(" And (A.Status IN (1, 3))  ");
                sb.Append(" and ((Case When A.Flag IN(297,697,200,600,210,201,698) Then -(A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)  Else (A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)  ");
                sb.Append(" End - (A.Offset+A.NoCheckOffSet + A.Discount + A.NoCheckDisCount))<>0)  And A.YearCompressType <> 1 ");
                sb.Append(" AND  A.BILLDATE BETWEEN @CreateDate AND @CreateDate1 ");
                if (checkBox4.Checked)
                {
                    sb.Append(" and  A.[CustID] in ( " + c + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  A.[CustID] between @CustID1 and @CustID2 ");
                    }
                }
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  A.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox1.Text != "")
                {
                    sb.Append(" and  A.DueTo  =@DUETO ");
                }
                sb.Append(" Union All ");
                sb.Append(" Select  ");
                sb.Append(" ''''+A.CustID 客戶編號, ");
                sb.Append(" B.ShortName  客戶名稱, ");
                sb.Append(" A.DueTo  帳款歸屬編號, ");
                sb.Append(" C.ShortName As 帳款歸屬, ");
                sb.Append(" CONVERT(VARCHAR(10) , cast(CAST(A.BillDate AS VARCHAR) as datetime),111)    日期,   ");
                sb.Append(" A.FundBillNO 單號, ");
                sb.Append(" L.InvoiceNO 發票號碼,   ");
                sb.Append(" A.Total+A.Tax 應收金額,  ");
                sb.Append(" ''''+A.VoucherNo AS 傳票編號,  ");
                sb.Append(" A.CashPay+A.VisaPay+A.OtherPay as 沖款金額, ");
                if (comboBox1.Text == "聿豐")
                {
                    sb.Append(" CASE WHEN A.Flag =698 THEN CASE  F.RecvWay WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(F.DistDays AS VARCHAR)+'天' WHEN 3 THEN  B.GatherOther END ELSE  ");
                    sb.Append(" CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+CAST(M.GatherDelay AS VARCHAR)+'天' WHEN 3 THEN M.GatherOther END END  付款方式,       ");
                }
                if (comboBox1.Text == "東門店")
                {
                    sb.Append(" CASE F.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END  付款方式,  ");
                }
                sb.Append(" (A.Total+A.Tax-A.CashPay-A.VisaPay-A.OtherPay) as 結餘, ");
                sb.Append(" Case When  substring(B.ShortName,CHARINDEX('-',B.ShortName)+1,1) IN ('進','聿','博','能') then  substring(B.ShortName,0,CHARINDEX('-',B.ShortName)) end as 員工,E.PersonName  業務,A.DEPTID 部門  ");
                sb.Append(" FROM comBillAccounts A        ");
                sb.Append(" Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO ");
                sb.Append(" Left Join (ComFundSub D        ");
                sb.Append(" Inner JOIN comFundMain H ON D.Flag = H.Flag AND D.FundBillNo = H.FundBillID And H.YearCompressType <> 1)        ");
                sb.Append(" ON A.Flag = D.OriginFlag AND A.FundBillNo = D.OriginBillNo       ");
                sb.Append(" Left JOIN ComInvoice L on A.InvoFlag = L.Flag And A.InvoBillNo = L.InvoBillNo        ");
                sb.Append(" Left Join ComCustomer B ON A.CustID = B.ID AND A.CustFlag = B.Flag        ");
                sb.Append(" Left JOIN ComCustomer C ON A.DueTo = C.ID AND A.CustFlag = C.Flag       ");
                sb.Append(" Left Join      comCustTrade  F On A.CustID=F.ID And A.CustFlag=F.Flag     ");
                sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan   ");
                sb.Append(" Left Join impAAMain V ON A.Flag = V.Flag And A.FundBillNO = V.AANO WHERE A.CustFlag = 1  ");
                sb.Append(" AND A.Status IN (1, 3)  ");
                sb.Append(" And A.YearCompressType <> 1 AND D.AccFlag = 1  And A.HasCheck=1        ");
                sb.Append(" AND ((A.Offset + A.Discount)<>0   ");
                sb.Append(" Or (Exists (Select J.OriginBillNO From ComFundSub J  ");
                sb.Append(" Where A.Flag = J.OriginFlag  And A.FundBillNO = J.OriginBillNO  ");
                sb.Append(" And J.OriginFlag <> 0 And Left(J.OriginBillNO,1) <> '*') ))  ");
                sb.Append(" AND  A.BILLDATE BETWEEN @CreateDate AND @CreateDate1 ");
                if (checkBox4.Checked)
                {
                    sb.Append(" and  A.[CustID] in ( " + c + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  A.[CustID] between @CustID1 and @CustID2 ");
                    }
                }
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  A.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox1.Text != "")
                {
                    sb.Append(" and  A.DueTo  =@DUETO ");
                }
            //    sb.Append(" Order By A.DueTo , A.BillDate, A.FundBillNO");



            }



           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@C2", textBox12.Text));
            command.Parameters.Add(new SqlParameter("@DUETO", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData4()
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
          sb.Append("   SELECT 客戶編號,客戶名稱,SUM(金額) 金額,SUM(A) '<0',SUM(B) '0~30',SUM(C) '31~60',SUM(D) '61~90',SUM(E) '>90' FROM (");
sb.Append("                  Select ");
sb.Append("                  ''''+A.CustID 客戶編號,");
sb.Append("                  B.ShortName  客戶名稱,");
sb.Append("                  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) 金額 ");
sb.Append("                      ,CASE WHEN datediff(day,");
sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
sb.Append(" , GETDATE()) < 0 THEN  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) END 'A'");
sb.Append("      ,CASE WHEN datediff(day,");
sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
sb.Append(" , GETDATE()) BETWEEN  0 AND 30 THEN  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) END 'B'");
sb.Append("      ,CASE WHEN datediff(day,");
sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
sb.Append(" , GETDATE()) BETWEEN  31 AND 60 THEN  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) END 'C'");
sb.Append("  ,CASE WHEN datediff(day,");
sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
sb.Append(" , GETDATE()) BETWEEN  61 AND 90 THEN  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) END 'D',");
sb.Append(" CASE WHEN datediff(day,");
sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
sb.Append(" , GETDATE()) > 90 THEN  A.Total+A.Tax-(A.CashPay+A.VisaPay+A.OtherPay) END 'E'");
sb.Append("                  From comBillAccounts A ");
sb.Append("                  Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO ");
sb.Append("                  Left Join ComInvoice L on A.InvoFlag = L.Flag And A.InvoBillNo = L.InvoBillNo ");
sb.Append("                  Left Join      comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag ");
sb.Append("    Left Join comPerson E On E.PersonID=A.SalesMan  ");
sb.Append("                  Left Join      comCustomer C On A.DueTo= C.ID And A.CustFlag=C.Flag      Left Join impAAMain V On A.Flag = V.Flag And A.Fundbillno = V.AANO ");
sb.Append("                  Where A.CustFlag=1 And A.YearCompressType <> 1 And  A.Status In (1, 3) And (A.Offset + A.Discount)=0  ");
sb.Append("                  And Not Exists (Select H.OriginBillNO From ComFundSub H ");
sb.Append("                  Where A.Flag = H.OriginFlag  ");
sb.Append("                  And A.FundBillNO = H.OriginBillNO And H.OriginFlag <> 0 ");
sb.Append("                  And Left(H.OriginBillNO,1) <> '*') And A.HasCheck=1  ");
sb.Append("                  And A.BillDate >=20140101 and  A.Total > 0");
sb.Append("           ) AS A GROUP BY 客戶編號,客戶名稱");







            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData5()
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("                 Select ");
            sb.Append("                 ''''+A.CustID 客戶編號,");
            sb.Append("                 B.ShortName  客戶名稱,");
            sb.Append("                 A.DueTo  帳款歸屬編號,");
            sb.Append("                 C.ShortName As 帳款歸屬,");
            sb.Append("                A.BillDate 日期, ");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay) 逾期日期");
            sb.Append("               , datediff(day,");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
            sb.Append(" , GETDATE()) 逾期天數");
            sb.Append("            ,     A.FundBillNO 單號,");
            sb.Append("                 A.InvoBillNo 發票號碼,  ");
            sb.Append("                 A.Total+A.Tax 應收金額, ");
            sb.Append("                 ''''+A.VoucherNo AS 傳票編號, ");
            sb.Append("                 A.CashPay+A.VisaPay+A.OtherPay as 沖款金額,");
            sb.Append("                 CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END as 付款方式,");
            sb.Append("                  (A.Total+A.Tax-A.CashPay-A.VisaPay-A.OtherPay) as 結餘,");
            sb.Append("                                     Case When  substring(B.ShortName,CHARINDEX('-',B.ShortName)+1,1) IN ('進','聿','博') then  substring(B.ShortName,0,CHARINDEX('-',B.ShortName)) end as 員工,E.PersonName  業務");
            sb.Append("                 From comBillAccounts A ");
            sb.Append("                 Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO ");
            sb.Append("                 Left Join ComInvoice L on A.InvoFlag = L.Flag And A.InvoBillNo = L.InvoBillNo ");
            sb.Append("                 Left Join      comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag ");
            sb.Append("    Left Join comPerson E On E.PersonID=A.SalesMan  ");
            sb.Append("                 Left Join      comCustomer C On A.DueTo= C.ID And A.CustFlag=C.Flag      Left Join impAAMain V On A.Flag = V.Flag And A.Fundbillno = V.AANO ");
            sb.Append("                 Where A.CustFlag=1 And A.YearCompressType <> 1 And  A.Status In (1, 3) And (A.Offset + A.Discount)=0  ");
            sb.Append("                 And Not Exists (Select H.OriginBillNO From ComFundSub H ");
            sb.Append("                 Where A.Flag = H.OriginFlag  ");
            sb.Append("                 And A.FundBillNO = H.OriginBillNO And H.OriginFlag <> 0 ");
            sb.Append("                 And Left(H.OriginBillNO,1) <> '*') And A.HasCheck=1  ");
            sb.Append("                 And A.BillDate >=20140101 and  A.Total > 0");
            sb.Append(" AND datediff(day,");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
            sb.Append(" , GETDATE()) > 0");
            sb.Append("                 Order By datediff(day,");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
            sb.Append(" , GETDATE()) DESC,A.CustID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData6()
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT SUM(結餘) 金額,逾期天數 FROM (     Select ");
            sb.Append("                ''''+A.CustID 客戶編號,");
            sb.Append("                 B.ShortName  客戶名稱,");
            sb.Append("                 A.DueTo  帳款歸屬編號,");
            sb.Append("                 C.ShortName As 帳款歸屬,");
            sb.Append("                A.BillDate 日期, ");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay) 逾期日期");
            sb.Append("               , datediff(day,");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),M.GatherDelay)");
            sb.Append(" , GETDATE()) 逾期天數");
            sb.Append("            ,     A.FundBillNO 單號,");
            sb.Append("                 L.InvoiceNO 發票號碼,  ");
            sb.Append("                 A.Total+A.Tax 應收金額, ");
            sb.Append("                 ''''+A.VoucherNo AS 傳票編號, ");
            sb.Append("                 A.CashPay+A.VisaPay+A.OtherPay as 沖款金額,");
            sb.Append("                 CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END as 付款方式,");
            sb.Append("                  (A.Total+A.Tax-A.CashPay-A.VisaPay-A.OtherPay) as 結餘,");
            sb.Append("                 Case When  CHARINDEX('-',B.ShortName) > 0 then  substring(B.ShortName, CHARINDEX('-',B.ShortName)+1,1) end as 員工");
            sb.Append("                 From comBillAccounts A ");
            sb.Append("                 Left Join stkBillMain M on M.Flag = A.Flag And M.BillNO = A.FundBillNO ");
            sb.Append("                 Left Join ComInvoice L on A.InvoBillNo = L.InvoBillNo ");
            sb.Append("                 Left Join      comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag ");
            sb.Append("                 Left Join      comCustomer C On A.DueTo= C.ID And A.CustFlag=C.Flag      Left Join impAAMain V On A.Flag = V.Flag And A.Fundbillno = V.AANO ");
            sb.Append("                 Where A.CustFlag=1 And A.YearCompressType <> 1 And  A.Status In (1, 3) And (A.Offset + A.Discount)=0  ");
            sb.Append("                 And Not Exists (Select H.OriginBillNO From ComFundSub H ");
            sb.Append("                 Where A.Flag = H.OriginFlag  ");
            sb.Append("                 And A.FundBillNO = H.OriginBillNO And H.OriginFlag <> 0 ");
            sb.Append("                 And Left(H.OriginBillNO,1) <> '*') And A.HasCheck=1  ");
            sb.Append("                 And A.BillDate >=20140101 and  A.Total > 0");
            sb.Append(" AND datediff(day,");
            sb.Append("               dbo.fun_CreditDate(CASE M.GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN M.GatherOther END,cast(CAST(A.BillDate AS VARCHAR) as datetime),GatherDelay)");
            sb.Append(" , GETDATE()) > 0 ) AS A GROUP BY 逾期天數 ORDER BY 逾期天數 DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetA1(string BillNO)
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT FromNO    FROM ComProdRec WHERE BillNO=@BillNO AND FLAG=500 ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetA2(string BillNO)
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("  select CustBillNo from ordBillMain where BillNO =@BillNO  AND Flag =2 ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPAY(string OriginBillNo)
        {
            if (comboBox1.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox1.Text == "東門店")
            {
                CONN = strCn3;
            }
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("  select FundBillID ID,VoucherNo VOUCHER  from comFundSub T0");
            sb.Append("  LEFT JOIN comFundMain T1 ON (T0.FundBillNo =T1.FundBillID AND T0.Flag =T1.Flag)");
            sb.Append(" where T0.OriginBillNo =@OriginBillNo AND T0.Flag =82");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OriginBillNo", OriginBillNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetA3(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ORDERPIN FROM GB_POTATO WHERE CAST(ID AS VARCHAR)=@ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void rbAr1_Click(object sender, EventArgs e)
        {
        //    EXEC();
        }

        private void rbAr2_Click(object sender, EventArgs e)
        {
         //   EXEC();
        }

        private void rbAr3_Click(object sender, EventArgs e)
        {
    
        }

        private void button4_Click(object sender, EventArgs e)
        {
               if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
               else if (tabControl1.SelectedIndex == 2)
               {
                   ExcelReport.GridViewToExcel(dataGridView3);
               }
               else if (tabControl1.SelectedIndex == 3)
               {
                   ExcelReport.GridViewToExcel(dataGridView4);
               }

        }

        public void EXEC()
        {
            DTC = MakeTableCombine();
            System.Data.DataTable DT = GetOrderData3();
            DataRow dr = null;
            for (int i = 0; i <= DT.Rows.Count - 1; i++)
            {
                dr = DTC.NewRow();
                DataRow dd = DT.Rows[i];
                dr["客戶編號"] = dd["客戶編號"].ToString();
                dr["客戶名稱"] = dd["客戶名稱"].ToString();
                dr["帳款歸屬編號"] = dd["帳款歸屬編號"].ToString();
                dr["帳款歸屬"] = dd["帳款歸屬"].ToString();
                dr["日期"] = dd["日期"].ToString();
                if (rbAr1.Checked)
                {
                    dr["逾期日期"] = dd["逾期日期"].ToString();
                    dr["逾期天數"] = dd["逾期天數"].ToString();
                }
                string DOC = dd["單號"].ToString();
                System.Data.DataTable T1 = GetA1(DOC);

                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();
                if (T1.Rows.Count > 0)
                {
                    for (int J = 0; J <= T1.Rows.Count - 1; J++)
                    {
                        DataRow SS = T1.Rows[J];
                        string SO = SS[0].ToString();
                        sb.Append(SO + "/-");
                        System.Data.DataTable T2 = GetA2(SO);
                        if (T2.Rows.Count > 0)
                        {
                            string PSO = T2.Rows[0][0].ToString();
                            if (!String.IsNullOrEmpty(PSO))
                            {
                                if (PSO.Length > 4)
                                {
                                    sb2.Append("'"+PSO + "/");
                                }
                                //else
                                //{
                                //    System.Data.DataTable T3 = GetA3(PSO);
                                //    if (T3.Rows.Count > 0)
                                //    {
                                //        string PINSO = T3.Rows[0][0].ToString();
                                //        if (!String.IsNullOrEmpty(PINSO))
                                //        {
                                //            sb2.Append(PINSO + "/");
                                //        }

                                //    }
                                //}

                            }
                        }
                    }

                    sb.Remove(sb.Length - 2, 2);
                    if (sb2.Length != 0)
                    {
                        sb2.Remove(sb2.Length - 1, 1);
                    }
                    dr["訂單單號"] = sb.ToString();
                    dr["外部訂單編號"] = sb2.ToString();
                    
                    
                
                }
                dr["銷售單號"] = dd["單號"].ToString();
                dr["發票號碼"] = dd["發票號碼"].ToString();
                dr["應收金額"] = Convert.ToInt32(dd["應收金額"]);
                dr["傳票編號"] = dd["傳票編號"].ToString();
                dr["沖款金額"] = dd["沖款金額"].ToString();
                dr["付款方式"] = dd["付款方式"].ToString();
                dr["結餘"] = Convert.ToInt32(dd["結餘"]);
                dr["員工"] = dd["員工"].ToString();
                dr["業務"] = dd["業務"].ToString();
                dr["部門"] = dd["部門"].ToString();
                System.Data.DataTable TT1 = GetPAY(DOC);
                if (TT1.Rows.Count > 0)
                {
                    dr["付款單號"] = TT1.Rows[0]["ID"].ToString();
                    dr["付款傳票"] = TT1.Rows[0]["VOUCHER"].ToString();
                }
                DTC.Rows.Add(dr);
            }
            dataGridView1.DataSource = DTC;
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (rbAr1.Checked)
            {
                if (e.RowIndex >= dataGridView1.Rows.Count)
                    return;
                DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
                try
                {
                    if (!String.IsNullOrEmpty(dgr.Cells["逾期天數"].Value.ToString()))
                    {

                        if (Convert.ToInt32(dgr.Cells["逾期天數"].Value.ToString()) >= 0)
                        {

                            dgr.DefaultCellStyle.BackColor = Color.Pink;
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private System.Data.DataTable MakeTableCombine()
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("帳款歸屬編號", typeof(string));
            dt.Columns.Add("帳款歸屬", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            if (rbAr1.Checked)
            {
                dt.Columns.Add("逾期日期", typeof(string));
                dt.Columns.Add("逾期天數", typeof(string));
            }
            dt.Columns.Add("銷售單號", typeof(string));
            dt.Columns.Add("訂單單號", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("應收金額", typeof(int));
            dt.Columns.Add("傳票編號", typeof(string));
            dt.Columns.Add("沖款金額", typeof(string));
            dt.Columns.Add("付款方式", typeof(string));
            dt.Columns.Add("結餘", typeof(int));
            dt.Columns.Add("員工", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("外部訂單編號", typeof(string));
            dt.Columns.Add("付款單號", typeof(string));
            dt.Columns.Add("付款傳票", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelUtils.DataGridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelUtils.DataGridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelUtils.DataGridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelUtils.DataGridViewToExcel(dataGridView4);
            }

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EXEC();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            APS1CHOICE frm1 = new APS1CHOICE();
            frm1.CARDTYPE = "客戶";
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox4.Checked = true;
                c = frm1.q;

            }
        }


       
    }
}

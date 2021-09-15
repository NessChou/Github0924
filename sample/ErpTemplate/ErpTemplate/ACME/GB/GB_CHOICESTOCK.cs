using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class GB_CHOICESTOCK : Form
    {
                public string c;
        string strCn ="";
        System.Data.DataTable G1 = null;
        System.Data.DataTable G2 = null;
        System.Data.DataTable dtCost = null;
        public GB_CHOICESTOCK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
                 G1 = GetCHO11();
                 B2();
             //    G2 = GetCHO21();
            dataGridView1.DataSource = G1;
            dataGridView2.DataSource = G2;

            DataRow row;
            Int32[] Total = new Int32[G1.Columns.Count - 1];
            for (int i = 0; i <= G1.Rows.Count - 1; i++)
            {
                for (int j = 9; j <= 11; j++)
                {
                    try
                    {
                        Total[j - 1] += Convert.ToInt32(G1.Rows[i][j]);
                    }
                    catch
                    {
                        Total[j - 1] += 0;
                    }
                }
            }
            row = G1.NewRow();

            row[8] = "合計";
            for (int j = 9; j <= 11; j++)
            {
                row[j] = Total[j - 1];
            }
            G1.Rows.Add(row);
            for (int i = 9; i <= 11; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.DefaultCellStyle.Format = "#,##0";
            }


            DataRow row2;
            Int32[] Total2 = new Int32[G2.Columns.Count - 1];
            for (int i = 0; i <= G2.Rows.Count - 1; i++)
            {
                for (int j = 21; j <= 27; j++)
                {
                    try
                    {
                        Total2[j - 1] += Convert.ToInt32(G2.Rows[i][j]);
                    }
                    catch
                    {
                        Total2[j - 1] += 0;
                    }
                }
            }
            row2 = G2.NewRow();

            row2[2] = "合計";
            for (int j = 21; j <= 27; j++)
            {
                row2[j] = Total2[j - 1];
            }
            G2.Rows.Add(row2);
            for (int i = 21; i <= 27; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.DefaultCellStyle.Format = "#,##0";
            }


            System.Data.DataTable dt = GetCHO31();
             dtCost = MakeTableCombine();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string 會計科目 = dt.Rows[i]["會計科目"].ToString();
                string 類別名稱 = dt.Rows[i]["類別名稱"].ToString();

                System.Data.DataTable dt2 = GetCHO32(會計科目);
                for (int i2 = 0; i2 <= dt2.Rows.Count - 1; i2++)
                {
                    dr = dtCost.NewRow();

                    if (i2 == 0)
                    {
                        dr["會計科目"] = 會計科目;
                        dr["類別名稱"] = 類別名稱;
                    }
                    else
                    {
                        dr["會計科目"] = "";
                        dr["類別名稱"] = "";
                    }
                    dr["發票品名"] = dt2.Rows[i2]["發票品名"].ToString();
                    dr["數量"] = dt2.Rows[i2]["數量"].ToString();
                    dr["金額"] = dt2.Rows[i2]["金額"].ToString();

                    dtCost.Rows.Add(dr);
                }

                dr = dtCost.NewRow();
                dr["會計科目"] = 會計科目 + 類別名稱 + "合計";
                dr["類別名稱"] = "";
                dr["發票品名"] = "";
                dr["數量"] = dt.Rows[i]["數量"].ToString();
                dr["金額"] = dt.Rows[i]["金額"].ToString();
                dtCost.Rows.Add(dr);
            }

            System.Data.DataTable dt3 = GetCHO33();
            dr = dtCost.NewRow();
            dr["會計科目"] = "總計";
            dr["類別名稱"] = "";
            dr["發票品名"] = "";
            dr["數量"] = dt3.Rows[0]["數量"].ToString();
            dr["金額"] = dt3.Rows[0]["金額"].ToString();
            dtCost.Rows.Add(dr);
            dataGridView3.DataSource = dtCost;

            for (int i = 3; i <= 4; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void B2()
        {
 
            System.Data.DataTable dtCost = MakeTableCombineB();
            DataRow dr = null;
            System.Data.DataTable DT1 = GetCHOB1();
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                DataRow dd = DT1.Rows[i];
                dr = dtCost.NewRow();
                string COMPANY = dd["公司別"].ToString();
                string 部門 = dd["部門"].ToString();
                dr["公司別"] = COMPANY;
                dr["部門"] = 部門;
                dr["單據類別"] = dd["單據類別"].ToString();
                dr["季別"] = dd["季別"].ToString();
                dr["群組"] = dd["群組"].ToString();
                dr["年"] = dd["年"].ToString();
                dr["月"] = dd["月"].ToString();
                dr["類別名稱"] = dd["類別名稱"].ToString();   
                dr["B"] = dd["B"].ToString();
                dr["贈品"] = dd["贈品"].ToString();
                //string PIN = dd["品項"].ToString();
                string CLASSNAME = dd["類別"].ToString();
                dr["類別"] = CLASSNAME;
                //if (CLASSNAME == "加工品-調味肉品")
                //{
                //    PIN = "加工品";
                //}
                dr["品項"] =  dd["品項"].ToString();
                dr["客戶代碼"] = dd["客戶代碼"].ToString();
                dr["客戶簡稱"] = dd["客戶簡稱"].ToString();
                dr["訂購憑單號碼"] = dd["訂購憑單號碼"].ToString();
                dr["銷貨單據日期"] = dd["銷貨單據日期"].ToString();
                string BILLNO = dd["銷貨單據號碼"].ToString();
                dr["銷貨單據號碼"] = BILLNO;

                System.Data.DataTable DTINV = null;
                if (COMPANY == "聿豐")
                {
                    DTINV = GetINVOICE(BILLNO);
                }
                if (COMPANY == "東門")
                {
                    BILLNO = BILLNO.Replace("T", "");
                    DTINV = GetINVOICE2(BILLNO);
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
                string PRODID = dd["產品編號"].ToString();
                dr["產品編號"] = PRODID;
                dr["品名規格"] = dd["品名規格"].ToString();
                dr["倉別"] = dd["倉別"].ToString();
                dr["數量"] = dd["數量"].ToString();
                dr["單位"] = dd["單位"].ToString();
                dr["單價"] = dd["單價"].ToString();
                dr["金額未稅"] = dd["金額未稅"].ToString();
                dr["稅"] = dd["稅"].ToString();
                dr["金額含稅"] = dd["金額含稅"].ToString();
                dr["成本"] = dd["成本"].ToString();
                dr["毛利"] = dd["毛利"].ToString();
                if (部門 == "C2")
                {
                    dr["收款方式"] = dd["帳款歸屬"].ToString();
                }
                else
                {
                    dr["收款方式"] = dd["收款方式"].ToString();
                }
                dr["銷貨收入科目"] = dd["銷貨收入科目"].ToString();
                dr["傳票號碼"] = dd["傳票號碼"].ToString();
                dtCost.Rows.Add(dr);
            }
            G2 = dtCost;
        }
        public System.Data.DataTable GetINVOICE(string SrcBillNO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT InvoiceDate 發票日期,InvoiceNO 發票號碼  FROM CHICOMP02.DBO.comInvoice WHERE SrcBillNO =@SrcBillNO AND Flag IN (2,4) AND IsCancel <> 1      ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));
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
        public System.Data.DataTable GetINVOICE2(string SrcBillNO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT InvoiceDate 發票日期,InvoiceNO 發票號碼  FROM CHICOMP03.DBO.comInvoice WHERE  REPLACE(SrcBillNO,'T','') =@SrcBillNO AND Flag =2 AND IsCancel <> 1      ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));
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
        public System.Data.DataTable GetCHO11()
        {
            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  CASE A.Flag WHEN 500 THEN '銷貨' WHEN 600 THEN '銷退' WHEN 698 THEN '銷折' END 單據別,convert(varchar,CAST(CAST(A.BillDate AS VARCHAR) AS DATETIME), 111) 單據日期,A.FundBillNo 單據號碼,B.ShortName 客戶簡稱 ");
           sb.Append(" , H.InvoiceType 發票聯式,H.InvoiceNO  發票號碼,Q.FullName 付款方式,H.TaxRegNO 統一編號, M.PersonName 業務人員,");
           sb.Append(" CASE A.Flag WHEN 500 THEN A.Total WHEN 600 THEN A.Total*-1 WHEN 698 THEN A.Total*-1 END 未稅金額,");
           sb.Append(" CASE A.Flag WHEN 500 THEN A.TAX WHEN 600 THEN A.TAX*-1 WHEN 698 THEN A.TAX*-1 END 稅額,");
           sb.Append(" CASE A.Flag WHEN 500 THEN A.Total+A.Tax  WHEN 600 THEN (A.Total+A.Tax )*-1 WHEN 698 THEN (A.Total+A.Tax )*-1 END 含稅金額");
           sb.Append(" FROM comBillAccounts A  ");
           sb.Append(" LEFT  JOIN comCustomer B ON B.Flag = A.CustFlag AND A.CustID = B.ID  ");
           sb.Append(" LEFT  JOIN comInvoice H ON A.InvoFlag = H.Flag AND A.InvoBillNo = H.InvoBillNO ");
           sb.Append(" LEFT  JOIN comPerson M ON A.Salesman = M.PersonID  ");
           sb.Append(" Left join comCustomer Q ON A.DueTo = Q.ID And Q.Flag = A.CustFlag  ");
            sb.Append(" WHERE A.YearCompressType=0 and  A.BillDate Between @BillDate1 And @BillDate2 AND ");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append(" A.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" A.Flag=600");
            }
            else if (comboBox1.Text == "銷折")
            {
                sb.Append(" A.Flag=698");
            }
            else
            {
                sb.Append(" A.Flag IN  (500,600,698)");
            }

            if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND A.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND A.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND A.CustID NOT IN ('C00004','C00005') ");
            }
            else if (comboBox3.Text != "")
            {
                sb.Append("  AND A.CustID   = '" + comboBox3.Text + "' ");
            }

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

            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
            }

     
            sb.Append(" Order by A.BillDate, A.FundBillNO");
        

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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
        public System.Data.DataTable GetCHO12()
        {

            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  Q.FullName 付款方式,SUM(A.Total) 未稅金額,SUM(A.TAX) 稅額,SUM(A.Total+A.Tax) 含稅金額 ");
            sb.Append(" FROM comBillAccounts A  ");
            sb.Append(" LEFT  JOIN comCustomer B ON B.Flag = A.CustFlag AND A.CustID = B.ID  ");
            sb.Append(" LEFT  JOIN comInvoice H ON A.InvoFlag = H.Flag AND A.InvoBillNo = H.InvoBillNO ");
            sb.Append(" LEFT  JOIN comPerson M ON A.Salesman = M.PersonID  ");
            sb.Append(" Left join comCustomer Q ON A.DueTo = Q.ID And Q.Flag = A.CustFlag  ");
            sb.Append(" WHERE A.YearCompressType=0 and  A.BillDate Between @BillDate1 And @BillDate2 AND ");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append(" A.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" A.Flag=600");
            }
            else
            {
                sb.Append(" A.Flag IN  (500,600)");
            }
            if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND A.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND A.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND A.CustID NOT IN ('C00004','C00005') ");
            }
            else if (comboBox3.Text != "")
            {
                sb.Append("  AND A.CustID   = '" + comboBox3.Text + "' ");
            }


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

            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
            }

            sb.Append(" GROUP BY Q.FullName");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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

        public System.Data.DataTable GetCHO21()
        {
            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select J.ProdName 品名規格,W.WareHouseName 倉別,J.Unit 單位,A.Quantity  數量,CAST(ROUND(A.Price,0) AS INT) 單價");
            sb.Append(" ,CAST(ROUND(A.MLAmount,0) AS INT) 金額未稅,CAST(ROUND(A.TaxAmt,0) AS INT) 稅,CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT) 金額含稅,CostForAcc 成本 ,A.MLAmount-A.CostForAcc 毛利,              ");
            sb.Append(" U3.FullName  收款方式, K.AccSale+U4.SubjectName  銷貨收入科目  ");
            sb.Append(" From DBO.ComProdRec A                          ");
            sb.Append(" Left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag              ");
            sb.Append(" Left join DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1              ");
            sb.Append(" Left join DBO.comWareHouse W On  A.WareID=W.WareHouseID                                                                                                                                                 ");
            sb.Append(" Left join DBO.comPerson P ON (T.Salesman=P.PersonID)              ");
            sb.Append(" Left Join DBO.comProduct J On A.ProdID =J.ProdID              ");
            sb.Append(" Left Join DBO.comProductClass K On J.ClassID =K.ClassID                         ");
            sb.Append(" Left join DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1        ");
            sb.Append(" Left join DBO.ComSubject U4 On  K.AccSale =U4.SubjectID      ");
            sb.Append("   LEFT  JOIN comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO  ");
            sb.Append(" WHERE A.YearCompressType=0 and  A.BillDate Between @BillDate1 And @BillDate2 AND  ");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append(" A.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" A.Flag=600");
            }
            else
            {
                sb.Append(" A.Flag IN  (500,600)");
            }

            if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND T.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND T.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
            }
            else if (comboBox3.Text != "")
            {
                sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
            }
            if (checkBox4.Checked)
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

            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }

            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
            }
            sb.Append(" ORDER BY A.ProdID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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
        private void button5_Click(object sender, EventArgs e)
        {
         
                System.Data.DataTable G12 = GetCHO12();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string D1 = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
                string D2 = textBox2.Text.Substring(0, 4) + "/" + textBox2.Text.Substring(4, 2) + "/" + textBox2.Text.Substring(6, 2);
                string D3 = "日期區間: " + D1 + " ～ " + D2;
                string D4 = "";
                string D5 = textBox1.Text.Substring(4, 2);
                string D6 = comboBox1.Text;
                if (comboBox2.Text == "聿豐")
                {
                    D4 = "聿豐實業股份有限公司";
                }
                else if (comboBox2.Text == "東門")
                {
                    D4 = "聿豐實業股份有限公司東門店";
                }


                //Excel的樣版檔


                //輸出檔
             

                //產生 Excel Report
                if (tabControl1.SelectedIndex == 0)
                {
                    FileName = lsAppDir + "\\Excel\\GW\\銷售報表1.xlsx";

                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                   DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                    ExcelReport.ExcelGBCHOICE1(G1, G12, D3, D4, FileName, OutPutFile);
                }
                if (tabControl1.SelectedIndex == 1)
                {
                    ExcelReport.GridViewToExcel(dataGridView2);
                }
                if (tabControl1.SelectedIndex == 2)
                {
                    FileName = lsAppDir + "\\Excel\\GW\\銷售報表2.xlsx";
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                   DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                    ExcelReport.ExcelGBCHOICE2(FileName, OutPutFile, dtCost, D1, D5, D6);
                }
       
        }

        private void GB_CHOICESTOCK_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
            comboBox2.Text = "聿豐";


        }


        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("會計科目", typeof(string));
            dt.Columns.Add("類別名稱", typeof(string));
            dt.Columns.Add("發票品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("金額", typeof(decimal));
            return dt;
        }

        public System.Data.DataTable GetCHO31()
        {
            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select  K.AccSale 會計科目,U4.SubjectName 類別名稱, ISNULL(CAST(SUM(A.Quantity) AS INT),0)  數量 ,SUM(CAST(ROUND(A.MLAmount,0) AS INT)) 金額 From DBO.ComProdRec A  ");
            sb.Append(" Left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag            ");
            sb.Append("   LEFT  JOIN DBO.comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO  ");
            sb.Append(" Left Join DBO.comProduct J On A.ProdID =J.ProdID               ");
            sb.Append(" Left Join DBO.comProductClass K On J.ClassID =K.ClassID                               ");
            sb.Append(" Left join DBO.ComSubject U4 On  K.AccSale =U4.SubjectID       ");
            sb.Append(" WHERE A.YearCompressType=0  and  A.BillDate Between @BillDate1 And @BillDate2  AND ");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append(" A.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" A.Flag=600");
            }
            else
            {
                sb.Append(" A.Flag IN  (500,600)");
            }

            if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND T.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND T.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
            }
            else if (comboBox3.Text != "")
            {
                sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
            }

            if (checkBox4.Checked)
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
            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  = '" + textBox4.Text + "' ");
            }

            sb.Append(" GROUP BY K.AccSale,U4.SubjectName");
            sb.Append(" ORDER BY  K.AccSale ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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
        public System.Data.DataTable GetCHO32(string AccSale)
        {
            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select J.InvoProdName 發票品名,SUM(A.Quantity)  數量 ,SUM(CAST(ROUND(A.MLAmount,0) AS INT)) 金額");
            sb.Append(" From DBO.ComProdRec A");
            sb.Append(" Left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag            ");
            sb.Append(" LEFT  JOIN DBO.comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO  ");
            sb.Append(" Left Join DBO.comProduct J On A.ProdID =J.ProdID               ");
            sb.Append(" Left Join DBO.comProductClass K On J.ClassID =K.ClassID                               ");
            sb.Append(" WHERE A.YearCompressType=0  and  A.BillDate Between @BillDate1 And @BillDate2    ");
            sb.Append(" AND  K.AccSale=@AccSale  AND");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append(" A.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" A.Flag=600");
            }
            else
            {
                sb.Append(" A.Flag IN  (500,600)");
            }


             if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND T.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND T.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
            }
             else if (comboBox3.Text !="")
             {
                 sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
             }
             if (checkBox4.Checked)
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
             if (textBox10.Text != "" && textBox12.Text != "")
             {
                 sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
             }
            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  = '" + textBox4.Text + "' ");
            }
            sb.Append(" GROUP BY  A.ProdID,J.InvoProdName");
            sb.Append(" ORDER BY  A.ProdID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AccSale", AccSale));
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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

        public System.Data.DataTable GetCHO33()
        {
            if (comboBox2.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            else if (comboBox2.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select   ISNULL(CAST(SUM(A.Quantity) AS INT),0)  數量 ,ISNULL(CAST(ROUND(SUM(A.MLAMOUNT),0) AS INT),0) 金額 From DBO.ComProdRec A");
            sb.Append(" Left join DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag            ");
            sb.Append("   LEFT  JOIN DBO.comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO  ");
            sb.Append(" WHERE A.YearCompressType=0 and  A.BillDate Between @BillDate1 And @BillDate2 AND ");
            if (comboBox1.Text == "銷貨")
            {
                sb.Append("  T.Flag=500");
            }
            else if (comboBox1.Text == "銷退")
            {
                sb.Append(" T.Flag=600");
            }
            else if (comboBox1.Text == "銷折")
            {
                sb.Append(" T.Flag=698");
            }
            else
            {
                sb.Append(" T.Flag IN  (500,600,698)");
            }

            if (comboBox3.Text == "禾豐黃舉昇")
            {
                sb.Append(" AND T.CustID = 'C00004' ");
            }
            else if (comboBox3.Text == "聿豐東門店")
            {
                sb.Append(" AND T.CustID = 'C00005' ");
            }
            else if (comboBox3.Text == "其他")
            {
                sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
            }
            else if (comboBox3.Text != "")
            {
                sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
            }

            if (checkBox4.Checked)
            {
                sb.Append(" and  ATCustID] in ( " + c + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T.[CustID] between @CustID1 and @CustID2 ");
                }
            }
            if (textBox10.Text != "" && textBox12.Text != "")
            {
                sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
            }

            if (textBox4.Text != "")
            {
                sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
            }









            //sb.Append(" Order by A.BillDate, A.FundBillNO");
        
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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

        public System.Data.DataTable GetCHOB1()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            if (comboBox2.Text == "聿豐")
            {
                sb.Append(" Select  A.Flag ,'聿豐' 公司別,T.DeptID 部門,CASE A.Flag WHEN 500 THEN '銷貨' WHEN 600 THEN '銷退' WHEN 701 THEN '銷折' END 單據類別,'Q'+CAST(datepart(qq,(cast(cast(A.BillDate as varchar) as datetime))) AS VARCHAR)   季別,    ");
                sb.Append(" CASE WHEN K.ClassName LIKE '%材料%' THEN 'Trading' ELSE 'Brand' END 群組,   ");
                sb.Append(" year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,             ");
                sb.Append(" L.ClassName 類別名稱,L.ENGNAME B,              ");
                sb.Append(" CASE WHEN SUBSTRING(A.ProdID ,1,2)='A1' THEN '外購品'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'FRE' THEN '運費'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MCK' THEN '雞'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MPK' THEN '豬'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSR' THEN '蝦'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSF' THEN '魚'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PCK' THEN '加工品-雞'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PFH' THEN '加工品-魚'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PPK' THEN '加工品-豬'   ");
                sb.Append(" END 品項, A.TAXRATE,");
                sb.Append(" K.ClassName 類別,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,A.FromNO 訂購憑單號碼,       ");
                sb.Append(" A.BillDate  銷貨單據日期,A.BillNO 銷貨單據號碼,A.ProdID 產品編號,J.InvoProdName 品名規格, W.WareHouseName 倉別,CASE WHEN A.Flag=500 THEN A.Quantity WHEN A.Flag=701 THEN 0 WHEN A.Flag = 600 THEN A.Quantity*-1 END   數量,J.Unit 單位,CAST(ROUND(A.Price,0) AS INT) 單價,");
                sb.Append(" CASE WHEN A.Flag=500 THEN CAST(ROUND(A.MLAmount,0) AS INT)  WHEN A.Flag = 600 THEN CAST(ROUND(A.MLAmount,0) AS INT)*-1  WHEN A.Flag =701 THEN MLDIST*-1 END  金額未稅,");
                sb.Append(" CASE WHEN A.Flag=500 THEN CAST(ROUND(A.TaxAmt,0) AS INT) WHEN A.Flag  = 600THEN CAST(ROUND(A.TaxAmt,0) AS INT)*-1 WHEN A.Flag =701 THEN CAST(ROUND(MLDIST*(CASE WHEN A.TaxAmt >0 THEN 0.05 ELSE 0 END),0)*-1 AS INT) END   稅               ");
                sb.Append(" ,CASE WHEN A.Flag=500 THEN CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT)  WHEN A.Flag =600 THEN CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT)*-1 WHEN A.Flag =701 THEN (MLDIST+  CAST(ROUND(MLDIST*(CASE WHEN A.TaxAmt >0 THEN 0.05 ELSE 0 END),0)*-1 AS INT) )*-1 END  金額含稅,");
                sb.Append(" CASE WHEN A.Flag=500 THEN CostForAcc WHEN A.Flag IN (600,701) THEN CostForAcc*-1 END 成本  ,");
                sb.Append(" CASE WHEN A.Flag=500 THEN (A.MLAmount-A.CostForAcc) WHEN A.Flag =600 THEN (A.MLAmount-A.CostForAcc)*-1 WHEN A.Flag=701 THEN (A.MLDIST-A.CostForAcc)*-1  END  毛利,                   ");
                sb.Append(" CASE WHEN T.CustID ='90143-12' THEN 'Trading / '+ (     CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END)+CAST(F2.GatherDelay AS VARCHAR)  ELSE    ");
                sb.Append(" CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結'+ISNULL(CAST(F2.GatherDelay AS VARCHAR),'')  WHEN 3 THEN '其他' END END 收款方式   ");
                sb.Append(" , K.AccSale+U4.SubjectName  銷貨收入科目,''''+T.VoucherNO 傳票號碼,U3.FULLNAME 帳款歸屬,CASE  F.IsGift  WHEN '1' THEN '是' ELSE '否' END 贈品       ");
                sb.Append(" From CHICOMP02.DBO.ComProdRec A                           ");
                sb.Append(" Left join CHICOMP02.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag               ");
                sb.Append(" Left join CHICOMP02.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1               ");
                sb.Append(" Left join CHICOMP02.DBO.comWareHouse W On  A.WareID=W.WareHouseID               ");
                sb.Append(" Left join CHICOMP02.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2                ");
                sb.Append(" Left join CHICOMP02.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2                                                                                                                                         ");
                sb.Append(" Left join CHICOMP02.DBO.comPerson P ON (T.Salesman=P.PersonID)               ");
                sb.Append(" Left join CHICOMP02.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )               ");
                sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID               ");
                sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID               ");
                sb.Append(" Left Join CHICOMP02.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1                ");
                sb.Append(" Left Join CHICOMP02.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1           ");
                sb.Append(" Left Join CHICOMP02.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO          ");
                sb.Append(" Left Join CHICOMP02.DBO.stkBillMAIN F2 On A.BillNO =F2.BillNO and A.Flag =F2.Flag      ");
                sb.Append(" Left join CHICOMP02.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1          ");
                sb.Append(" Left join CHICOMP03.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1         ");
                sb.Append(" Left join CHICOMP02.DBO.ComSubject U4 On  K.AccSale =U4.SubjectID           ");
                sb.Append(" LEFT  JOIN CHICOMP02.DBO.comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO  ");
                sb.Append(" Where ISNULL(O2.BillStatus,0) <> 2 ");


                if (textBox1.Text != "" && textBox2.Text != "")
                {

                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 AND ");
                }
                if (comboBox1.Text == "銷貨")
                {
                    sb.Append(" A.Flag=500");
                }
                else if (comboBox1.Text == "銷退")
                {
                    sb.Append(" A.Flag =600");
                }
                else if (comboBox1.Text == "銷折")
                {
                    sb.Append(" A.Flag =701");
                }
                else
                {
                    sb.Append(" A.Flag IN  (500,600,701)");
                }

                if (comboBox3.Text == "禾豐黃舉昇")
                {
                    sb.Append(" AND T.CustID = 'C00004' ");
                }
                else if (comboBox3.Text == "聿豐東門店")
                {
                    sb.Append(" AND T.CustID = 'C00005' ");
                }
                else if (comboBox3.Text == "其他")
                {
                    sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
                }
                else if (comboBox3.Text != "")
                {
                    sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
                }
                if (checkBox4.Checked)
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
                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox4.Text != "")
                {
                    sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
                }
            }

           // sb.Append(" UNION ALL  ");
            if (comboBox2.Text == "東門")
            {
                sb.Append(" Select '東門' 公司別,T.DeptID 部門,CASE A.Flag WHEN 500 THEN '銷貨' WHEN 600 THEN '銷退' WHEN 701 THEN '銷折' END 單據類別,'Q'+CAST(datepart(qq,(cast(cast(A.BillDate as varchar) as datetime))) AS VARCHAR)   季別,  ");
                sb.Append(" CASE WHEN K.ClassName LIKE '%材料%' THEN 'Trading' ELSE 'Brand' END 群組, ");
                sb.Append(" year(cast(cast(A.BillDate as varchar) as datetime)) 年,month(cast(cast(A.BillDate as varchar) as datetime)) 月,           ");
                sb.Append(" L.ClassName 類別名稱,L.ENGNAME B,            ");
                sb.Append(" CASE WHEN SUBSTRING(A.ProdID ,1,2)='A1' THEN '外購品'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) IN ('AME','AMM','AMO','AMV') THEN 'Trading'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'FRE' THEN '運費'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MCK' THEN '雞'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MPK' THEN '豬'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSR' THEN '蝦'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'MSF' THEN '魚'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PCK' THEN '加工品-雞'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PFH' THEN '加工品-魚'  ");
                sb.Append(" WHEN SUBSTRING(A.ProdID ,1,3) = 'PPK' THEN '加工品-豬'  WHEN K.ClassName= '商業折扣' THEN '商業折扣'   ");
                sb.Append(" END 品項, K.ClassName 類別,''''+T.CustID 客戶代碼,U.ShortName 客戶簡稱,A.FromNO 訂購憑單號碼,     ");
                sb.Append(" A.BillDate  銷貨單據日期,A.BillNO 銷貨單據號碼,A.ProdID 產品編號,J.InvoProdName 品名規格, W.WareHouseName 倉別,CASE A.Flag WHEN 500 THEN A.Quantity WHEN 701 THEN 0 WHEN  600 THEN A.Quantity*-1 END   數量,J.Unit 單位,CAST(ROUND(A.Price,0) AS INT) 單價,CASE A.Flag WHEN 500 THEN CAST(ROUND(A.MLAmount,0) AS INT)  WHEN 600 THEN CAST(ROUND(A.MLAmount,0) AS INT)*-1 WHEN 701 THEN CAST(ROUND(A.MLAmount,0) AS INT)*-1 END  金額未稅,CASE WHEN A.Flag =500 THEN CAST(ROUND(A.TaxAmt,0) AS INT) WHEN A.Flag IN (600,701) THEN CAST(ROUND(A.TaxAmt,0) AS INT)*-1 END   稅             ");
                sb.Append(" ,CASE WHEN A.Flag=500 THEN CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT)  WHEN A.Flag IN (600,701) THEN CAST(ROUND(A.MLAmount+A.TaxAmt,0) AS INT)*-1 END  金額含稅,CASE WHEN A.Flag=500 THEN CostForAcc WHEN A.Flag IN (600,701) THEN CostForAcc*-1 END 成本  ,CASE WHEN A.Flag=500 THEN (A.MLAmount-A.CostForAcc) WHEN A.Flag IN (600,701) THEN (A.MLAmount-A.CostForAcc)*-1 END  毛利,         ");
                sb.Append(" U3.FULLNAME 收款方式, K.AccSale+U4.SubjectName  銷貨收入科目,''''+T.VoucherNO 傳票號碼,U3.FULLNAME 帳款歸屬,CASE  F.IsGift  WHEN '1' THEN 'V' END 贈品    ");
                sb.Append(" From CHICOMP03.DBO.ComProdRec A                         ");
                sb.Append(" Left join CHICOMP03.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag             ");
                sb.Append(" Left join CHICOMP03.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
                sb.Append(" Left join CHICOMP03.DBO.comWareHouse W On  A.WareID=W.WareHouseID             ");
                sb.Append(" Left join CHICOMP03.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2              ");
                sb.Append(" Left join CHICOMP03.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2                                                                                                                                       ");
                sb.Append(" Left join CHICOMP03.DBO.comPerson P ON (T.Salesman=P.PersonID)             ");
                sb.Append(" Left join CHICOMP03.DBO.comCustAddress AD ON (O2.AddressID=AD.AddrID AND O2.CustomerID=AD.ID )             ");
                sb.Append(" Left Join CHICOMP03.DBO.comProduct J On A.ProdID =J.ProdID             ");
                sb.Append(" Left Join CHICOMP03.DBO.comProductClass K On J.ClassID =K.ClassID             ");
                sb.Append(" Left Join CHICOMP03.DBO.comCustClass L On U.ClassID =L.ClassID and L.Flag =1              ");
                sb.Append(" Left Join CHICOMP03.DBO.comCustDesc M On U.ID =M.ID and M.Flag =1         ");
                sb.Append(" Left Join CHICOMP03.DBO.stkBillSub F On A.BillNO =F.BillNO and A.Flag =F.Flag AND A.RowNO =F.RowNO        ");
                sb.Append(" Left join CHICOMP03.DBO.comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1        ");
                sb.Append(" Left join CHICOMP03.DBO.comCustomer U3 On  U3.ID=T.DueTo  AND U3.Flag =1         ");
                sb.Append(" Left join CHICOMP03.DBO.ComSubject U4 On  K.AccSale =U4.SubjectID         ");
                sb.Append(" LEFT  JOIN CHICOMP03.DBO.comInvoice H ON T.InvoFlag = H.Flag AND T.InvoBillNo = H.InvoBillNO   ");
                sb.Append(" Where 1=1  ");


                if (textBox1.Text != "" && textBox2.Text != "")
                {

                    sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 AND ");
                }
                if (comboBox1.Text == "銷貨")
                {
                    sb.Append(" A.Flag=500");
                }
                else if (comboBox1.Text == "銷退")
                {
                    sb.Append(" A.Flag =600");
                }
                else if (comboBox1.Text == "銷折")
                {
                    sb.Append(" A.Flag =701");
                }
                else
                {
                    sb.Append(" A.Flag IN  (500,600,701)");
                }
                if (comboBox3.Text == "禾豐黃舉昇")
                {
                    sb.Append(" AND T.CustID = 'C00004' ");
                }
                else if (comboBox3.Text == "聿豐東門店")
                {
                    sb.Append(" AND T.CustID = 'C00005' ");
                }
                else if (comboBox3.Text == "其他")
                {
                    sb.Append(" AND T.CustID NOT IN ('C00004','C00005') ");
                }
                else if (comboBox3.Text != "")
                {
                    sb.Append("  AND T.CustID   = '" + comboBox3.Text + "' ");
                }
                if (checkBox4.Checked)
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

                if (textBox10.Text != "" && textBox12.Text != "")
                {
                    sb.Append(" and  T.DEPTID BETWEEN @C1 AND @C2 ");
                }
                if (textBox4.Text != "")
                {
                    sb.Append("  AND H.TaxRegNO  LIKE '%" + textBox4.Text + "%' ");
                }
    //            sb.Append("  ORDER BY 公司別,A.BillDate ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CustID1", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@CustID2", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@C1", textBox10.Text));
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

        private System.Data.DataTable MakeTableCombineB()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司別", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            dt.Columns.Add("單據類別", typeof(string));
            dt.Columns.Add("季別", typeof(string));
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("年", typeof(string));
            dt.Columns.Add("月", typeof(string));
            dt.Columns.Add("類別名稱", typeof(string));
            dt.Columns.Add("B", typeof(string));
            dt.Columns.Add("品項", typeof(string));
            dt.Columns.Add("類別", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("訂購憑單號碼", typeof(string));
            dt.Columns.Add("銷貨單據日期", typeof(string));
            dt.Columns.Add("銷貨單據號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("倉別", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("單位", typeof(string));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("金額未稅", typeof(decimal));
            dt.Columns.Add("稅", typeof(decimal));
            dt.Columns.Add("金額含稅", typeof(decimal));
            dt.Columns.Add("成本", typeof(decimal));
            dt.Columns.Add("毛利", typeof(decimal));
            dt.Columns.Add("收款方式", typeof(string));
            dt.Columns.Add("銷貨收入科目", typeof(string));
            dt.Columns.Add("傳票號碼", typeof(string));
            dt.Columns.Add("贈品", typeof(string));
            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G12 = GetCHO12();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string D1 = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
            string D2 = textBox2.Text.Substring(0, 4) + "/" + textBox2.Text.Substring(4, 2) + "/" + textBox2.Text.Substring(6, 2);
            string D3 = "日期區間: " + D1 + " ～ " + D2;
            string D4 = "";
            string D5 = textBox1.Text.Substring(4, 2);
            string D6 = comboBox1.Text;
            if (comboBox2.Text == "聿豐")
            {
                D4 = "聿豐實業股份有限公司";
            }
            else if (comboBox2.Text == "東門")
            {
                D4 = "聿豐實業股份有限公司東門店";
            }


            //Excel的樣版檔


            //輸出檔


            //產生 Excel Report
            FileName = lsAppDir + "\\Excel\\GW\\銷售報表.xlsx";

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
            ExcelReport.ExcelGBCHOICE(G1, G12, D3, D4, FileName, OutPutFile, G2, dtCost, D1, D5, D6);
       
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView2.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
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

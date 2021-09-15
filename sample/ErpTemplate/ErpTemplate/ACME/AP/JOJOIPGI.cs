using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace ACME
{
    public partial class JOJOIPGI : Form
    {
        string strCn = "";
        public JOJOIPGI()
        {
            InitializeComponent();
        }
        public string cs;
        private void button1_Click_1(object sender, EventArgs e)
        {
            

            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
              
                cs = frm1.q;
          
                string ed = "";

                if (!String.IsNullOrEmpty(cs))
                {
                   System.Data.DataTable dt2 = Getbb(cs);
                    dataGridView1.DataSource = dt2;



                    System.Data.DataTable dt3 = Getcc(cs, ed);
                    dataGridView2.DataSource = dt3;

                    decimal[] Total = new decimal[dt3.Columns.Count - 1];

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        string S = dt3.Rows[i][3].ToString();

                        Total[2] += Convert.ToDecimal(dt3.Rows[i][4]);

                    }


                    DataRow row;

                    row = dt3.NewRow();

                    row[3] = "合計";

                    row[4] = Total[2];


                    dt3.Rows.Add(row);

                    System.Data.DataTable dt4 = Getdd(cs);
                    dataGridView3.DataSource = dt4;


                    decimal[] Total2 = new decimal[dt4.Columns.Count - 1];

                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {


                        Total2[2] += Convert.ToDecimal(dt4.Rows[i][7]);

                    }


                    DataRow row2;

                    row2 = dt4.NewRow();

                    row2[6] = "合計";

                    row2[7] = Total2[2];


                    dt4.Rows.Add(row2);


                }
            }
        }

 
        public System.Data.DataTable Getbb(string cs)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.PRODID 項目料號,CAST(T1.QTY AS INT) 庫存,T0.CAvgCost 平均成本 FROM comProduct T0");
            sb.Append("  LEFT JOIN (select SUM(Quantity) QTY,PRODID from DBO.comWareAmount TS   GROUP BY TS.PRODID) T1 ON (T0.PRODID=T1.PRODID)");
            sb.Append(" where  T0.PRODID in ( " + cs + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
         

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
        public System.Data.DataTable Getcc(string cs,string es)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("                         Select '1' SEQ,A.PRODID 項目料號,A.BillDate 過帳日期, U.FULLNAME 客戶,CAST(A.Quantity AS INT) 數量,C.CurrencyName 幣別,");
            sb.Append("						 cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價");
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
            sb.Append("                                         Where A.Flag=500");

            if (cs != "")
            {
                sb.Append(" AND A.PRODID  in ( " + cs + ")  ");
            }
            if (es != "")
            {

                sb.Append(" AND  T.CustID in ( " + es + ") and CAST(A.Quantity AS INT) <> 0  ");
            }
    
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable Getdd(string cs)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("                         Select '1' SEQ,A.PRODID 項目料號,O2.BillDate 採購單日期,A.BillDate 收貨採購單日期,A.BILLNO 收貨採購單單號,U.FULLNAME 廠商,");
            sb.Append("						 cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價,CAST(A.Quantity AS INT) 數量,A.ItemRemark 備註 ");
            sb.Append("                                    From ComProdRec A   ");
            sb.Append("                                                            Left Join comWareHouse D On D.WareHouseID=A.WareID  ");
            sb.Append("                                        left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag  ");
            sb.Append("                                             left join comCustomer U On  U.ID=T.CustID AND U.Flag =2  ");
            sb.Append("                                                    left join comWareHouse W On  A.WareID=W.WareHouseID  ");
            sb.Append("                                                left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4  ");
            sb.Append("                                                     left join OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4   ");
            sb.Append("                                                left join comCurrencySys C On  O2.CurrID=C.CurrencyID  ");
            sb.Append("                                                   left join COMBILLACCOUNTS S ON (A.BillNO =S.FundBillNo AND S.Flag =100)  ");
            sb.Append("                                                              left join comPerson P ON (S.Salesman=P.PersonID)  ");
            sb.Append("                                  LEFT JOIN comProject  AD2 ON (O2.ProjectID=AD2.ProjectID )  ");
            sb.Append("                                             Where A.Flag=100 ");

            sb.Append(" AND  A.PRODID   in ( " + cs + ")    ");
      
                sb.Append(" ORDER BY A.PRODID  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable Getee(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct Convert(varchar(10),t0.docdate,112) 日期,t1.itemcode 原始料號,t2.itemcode 調整料號,cast(t1.quantity as int) 數量 from oige t0");
            sb.Append(" inner join ige1 t1 on (t0.docentry=t1.docentry)");
            sb.Append(" inner join (select t0.docdate,t1.itemcode,t1.quantity,t1.u_base_doc from oign t0 left join ign1 t1 on (t0.docentry=t1.docentry) where t0.u_acme_kind1 like '%料號調整%' ) t2 on (t0.docdate=t2.docdate and t1.quantity=t2.quantity and t1.docentry=t2.u_base_doc)");
            sb.Append(" where t0.u_acme_kind1 like '%料號調整%'  and   T1.[ItemCode] in ( " + cs + ") and t1.itemcode <> t2.itemcode ");
            sb.Append(" union all");
            sb.Append(" select '','數量加總','',sum(cast(t1.quantity as int)) 數量 from oige t0");
            sb.Append(" inner join ige1 t1 on (t0.docentry=t1.docentry)");
            sb.Append(" inner join (select t0.docdate,t1.itemcode,t1.quantity,t1.u_base_doc from oign t0 left join ign1 t1 on (t0.docentry=t1.docentry) where t0.u_acme_kind1 like '%料號調整%' ) t2 on (t0.docdate=t2.docdate and t1.quantity=t2.quantity and t1.docentry=t2.u_base_doc )");
            sb.Append(" where t0.u_acme_kind1 like '%料號調整%'  and   T1.[ItemCode] in ( " + cs + ") and t1.itemcode <> t2.itemcode ");
            sb.Append("    order by t1.itemcode,Convert(varchar(10),t0.docdate,112)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable Getee1()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                   Select A.BILLDATE 過帳日期,O.BILLNO 採購單號,A.BILLNO 收貨採購單單號,U.FULLNAME 廠商,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本");
            sb.Append("				   ,(Substring(T11.[ItemCode],15,1)) 地區,CAST(A.Quantity AS INT) 數量,cast(cast(O.PRICE as numeric(16,2)) as varchar) 單價,C.CurrencyName 幣別,A.ITEMREMARK 備註");
            sb.Append("                                        From otherDB.CHIComp22.DBO.ComProdRec A                  ");
            sb.Append("                                        left join otherDB.CHIComp22.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag  ");
            sb.Append("                                             left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =2 ");
            sb.Append("                                                    left join otherDB.CHIComp22.DBO.comWareHouse W On  A.WareID=W.WareHouseID  ");
            sb.Append("                                                left join otherDB.CHIComp22.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =4  ");
            sb.Append("                                                     left join otherDB.CHIComp22.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =4   ");
            sb.Append("                                                left join otherDB.CHIComp22.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID  ");
            sb.Append("															  LEFT JOIN ACMESQL02.DBO.OITM T11 ON (A.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                                             Where  A.FLAG=100 ");

            if (textBox1.Text != "" && textBox2.Text != "")
            {

                sb.Append(" and  A.BILLDATE between @DocDate1 and @DocDate2 ");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" and T.CustID ='" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" order by (A.BILLNO) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        public System.Data.DataTable GetALL()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("			      SELECT U_BU 項目群組,CAST(sum(T0.CTotalCost) AS INT) 存貨金額   FROM otherDB.CHIComp22.DBO.comProduct T0 ");
            sb.Append("			   LEFT JOIN ACMESQL02.DBO.OITM T11 ON (T0.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("              INNER  JOIN  ACMESQL02.DBO.[OITB] T2  ON  T11.itmsgrpcod = T2.itmsgrpcod  ");
            sb.Append("              where T0.CTotalCost>0 and t0.PRODID not in (select itemcode COLLATE  Chinese_Taiwan_Stroke_CI_AS from ACMESQL02.DBO.oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z')) ");
            sb.Append("              And substring(t0.PRODID,1,2) <> 'ZR' ");
            sb.Append("              And substring(t0.PRODID,1,2) <> 'ZA' ");
            sb.Append("              And substring(t0.PRODID,1,2) <> 'ZB' ");
            sb.Append("              group by U_BU  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

           
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

        public System.Data.DataTable Ged1()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("			                select   T0.BILLDATE  製單日期,U.FULLNAME 客戶名稱  ,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本 ");
            sb.Append("             				   ,(Substring(T11.[ItemCode],15,1)) 地區,CAST(T1.Quantity AS INT) 未結數量,cast(cast(T1.PRICE as numeric(16,2)) as varchar) 單價");
            sb.Append("							   ,C.CurrencyName 幣別");
            sb.Append("              , T1.PREINDATE 訂單交期,''''+T0.BILLNO 單號 from otherDB.CHIComp22.DBO.ordBillMain T0          ");
            sb.Append("              left join otherDB.CHIComp22.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)            ");
            sb.Append("              left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1           ");
            sb.Append("              Left Join otherDB.CHIComp22.DBO.comProduct J On T1.ProdID =J.ProdID          ");
            sb.Append("              left join otherDB.CHIComp22.DBO.comPerson P ON (T0.Salesman=P.PersonID)   ");
            sb.Append("			                   left join otherDB.CHIComp22.DBO.comCurrencySys C On  T0.CurrID=C.CurrencyID   ");
            sb.Append("			    LEFT JOIN ACMESQL02.DBO.OITM T11 ON (T1.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("              WHERE  T0.Flag =2 AND T0.BillStatus = 0  and QtyRemain  > 0    ");
            sb.Append("              and year(cast(cast(T0.BillDate as varchar) as datetime))>2015   ");

            if (textBox10.Text != "")
            {
                sb.Append(" and  T0.CustomerID ='" + textBox10.Text.ToString() + "' ");
            }
            if (textBox12.Text != "" && textBox11.Text != "")
            {

                sb.Append(" and  T0.BILLDATE between @DocDate1 and @DocDate2 ");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and Substring (T11.[ItemCode],2,8) ='" + comboBox1.SelectedValue.ToString() + "'  ");
            }
            sb.Append(" order by (t0.BILLDATE) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox12.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox11.Text));
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
        public System.Data.DataTable Ged2()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("			                select   T0.BILLDATE  製單日期,U.FULLNAME 廠商名稱   ,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本 ");
            sb.Append("             				   ,(Substring(T11.[ItemCode],15,1)) 地區,CAST(T1.Quantity AS INT) 未結數量,cast(cast(T1.PRICE as numeric(16,2)) as varchar) 單價");
            sb.Append("							   ,C.CurrencyName 幣別");
            sb.Append("              , T1.PREINDATE 訂單交期,''''+T0.BILLNO 單號 from otherDB.CHIComp22.DBO.ordBillMain T0          ");
            sb.Append("              left join otherDB.CHIComp22.DBO.ordBillSUB T1 ON (T0.Flag =T1.Flag AND T0.BillNO=T1.BillNO)            ");
            sb.Append("              left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T0.CustomerID AND U.Flag =2           ");
            sb.Append("              Left Join otherDB.CHIComp22.DBO.comProduct J On T1.ProdID =J.ProdID          ");
            sb.Append("              left join otherDB.CHIComp22.DBO.comPerson P ON (T0.Salesman=P.PersonID)   ");
            sb.Append("			                   left join otherDB.CHIComp22.DBO.comCurrencySys C On  T0.CurrID=C.CurrencyID   ");
            sb.Append("			    LEFT JOIN ACMESQL02.DBO.OITM T11 ON (T1.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("              WHERE  T0.Flag =4 AND T0.BillStatus = 0  and QtyRemain  > 0    ");
            sb.Append("              and year(cast(cast(T0.BillDate as varchar) as datetime))>2015   ");

            if (textBox14.Text != "")
            {
                sb.Append(" and  T0.CustomerID ='" + textBox14.Text.ToString() + "' ");
            }
            if (textBox16.Text != "" && textBox15.Text != "")
            {

                sb.Append(" and  T0.BILLDATE between @DocDate1 and @DocDate2 ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and Substring (T11.[ItemCode],2,8) ='" + comboBox2.SelectedValue.ToString() + "'  ");
            }
            sb.Append(" order by (t0.BILLDATE) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox16.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox15.Text));
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

        private void button2_Click_1(object sender, EventArgs e)
        {
            APS1 frm1 = new APS1();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                cs = frm1.q;
                string ed = "";
                if (!String.IsNullOrEmpty(cs))
                {
                    

                    System.Data.DataTable dt3 = Getcc(ed, cs);
                    dataGridView2.DataSource = dt3;

                    decimal[] Total = new decimal[dt3.Columns.Count - 1];

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {


                        Total[2] += Convert.ToDecimal(dt3.Rows[i][4]);

                    }
                    

                    DataRow row;

                    row = dt3.NewRow();

                    row[3] = "合計";
           
                        row[4] = Total[2];

                    
                    dt3.Rows.Add(row);
                 
                }
            }
            tabControl1.SelectedIndex = 3;
        }

      

       

        private void button5_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToCSV2(dataGridView1, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToCSV2(dataGridView3, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            //else if (tabControl1.SelectedIndex == 2)
            //{
            //    ExcelReport.GridViewToCSV2(dataGridView4, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            //}
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToCSV2(dataGridView2, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToCSV2(dataGridView5, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToCSV2(dataGridView6, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                ExcelReport.GridViewToCSV2(dataGridView7, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                ExcelReport.GridViewToCSV(dataGridView8, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 7)
            {
                ExcelReport.GridViewToCSV2(dataGridView9, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            }
            else if (tabControl1.SelectedIndex == 8)
            {
                ExcelReport.GridViewToCSV2(dataGridView10, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
           
            }

     
        }

      
        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetCHICUSTM();

            if (LookupValues != null)
            {
              textBox4.Text = Convert.ToString(LookupValues[0]);
              textBox3.Text = Convert.ToString(LookupValues[1]);

            }
        
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            dataGridView5.DataSource = Getee1();
        }

        private void JOJO_Load(object sender, EventArgs e)
        {

            strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();
            textBox12.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox11.Text = GetMenu.Day();
            textBox16.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox15.Text = GetMenu.Day();
            textBox17.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox18.Text = GetMenu.Day();
        
            UtilSimple.SetLookupBinding(comboBox1, GetOslp(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetOslp2(), "DataValue", "DataValue");
            label20.Text = "";
        }
        System.Data.DataTable GetOslp()
        {

            SqlConnection con = globals.shipConnection;
            string sql = " SELECT DISTINCT Substring (T1.[ItemCode],2,8) DataText,Substring (T1.[ItemCode],2,8) DataValue FROM RDR1 T1 iNNER JOIN OWHS T4 ON T4.whsCode = T1.whscode WHERE T1.[LINESTATUS] ='O' UNION ALL SELECT '0', 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM'  order by DataText ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ousr");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ousr"];
        }
        System.Data.DataTable GetOslp2()
        {

            SqlConnection con = globals.shipConnection;
            string sql = " SELECT DISTINCT Substring (T1.[ItemCode],2,8) DataText,Substring (T1.[ItemCode],2,8) DataValue FROM por1 T1  INNER  JOIN [dbo].[OITM] T11  ON  T1.[ItemCode] = T11.ItemCode WHERE T1.[LINESTATUS] ='O'  AND ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  UNION ALL SELECT '0', 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM'  order by DataText ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ousr");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ousr"];
        }
        private void button6_Click(object sender, EventArgs e)
        {

            System.Data.DataTable t1 = Gen_201004();
            dataGridView6.DataSource = t1;
        }



        private void button8_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetCHICUSTM2();

            if (LookupValues != null)
            {
                textBox7.Text = Convert.ToString(LookupValues[0]);
                textBox8.Text = Convert.ToString(LookupValues[1]);

            }
        }

  

   
        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView7.DataSource = Ged1();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = Ged2();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetCHICUSTM2();

            if (LookupValues != null)
            {
                textBox10.Text = Convert.ToString(LookupValues[0]);
                textBox9.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetCHICUSTM();

            if (LookupValues != null)
            {
                textBox14.Text = Convert.ToString(LookupValues[0]);
                textBox13.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            dataGridView9.DataSource = GetALL();
        }



        private System.Data.DataTable Gen_201004()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;



            StringBuilder sb = new StringBuilder();

            //彙總
            sb.Append("                              Select A.BILLDATE 過帳日期,U.FULLNAME 客戶名稱,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本 ");
            sb.Append("             				   ,(Substring(T11.[ItemCode],15,1)) 地區,CAST(A.Quantity AS INT) 數量,cast(cast(O.PRICE as numeric(16,2)) as varchar) 訂單單價");
            sb.Append("							   ,cast(cast(A.PRICE as numeric(16,2)) as varchar) 台幣單價	   ,cast(cast(A.AMOUNT as numeric(16,2)) as varchar) 台幣金額,");
            sb.Append("							   	   cast(cast(ISNULL(A.AMOUNT,0)-ISNULL(A.CostForAcc,0) as numeric(16,2)) as varchar) 台幣毛利,P.PersonName 業務, A.BillNO 單據");
            sb.Append("                                                     From otherDB.CHIComp22.DBO.ComProdRec A                   ");
            sb.Append("                                                     left join otherDB.CHIComp22.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag   ");
            sb.Append("                                                          left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1   ");
            sb.Append("                                                                 left join otherDB.CHIComp22.DBO.comWareHouse W On  A.WareID=W.WareHouseID   ");
            sb.Append("                                                             left join otherDB.CHIComp22.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =2   ");
            sb.Append("                                                                  left join otherDB.CHIComp22.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =2    ");
            sb.Append("                                                             left join otherDB.CHIComp22.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID   ");
            sb.Append("															      left join  otherDB.CHIComp22.DBO.comPerson P ON (T.Salesman=P.PersonID) ");
            sb.Append("             															  LEFT JOIN ACMESQL02.DBO.OITM T11 ON (A.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("                                                          Where  A.FLAG=500  ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and   A.BILLDATE between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T.CustID ='" + textBox7.Text.ToString() + "' ");
            }
            sb.Append(" UNION ALL");
            sb.Append("														            Select A.BILLDATE 過帳日期,U.FULLNAME 客戶名稱,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本 ");
            sb.Append("             				   ,(Substring(T11.[ItemCode],15,1)) 地區,CAST(A.Quantity AS INT) 數量,cast(cast(O.PRICE as numeric(16,2)) as varchar) 訂單單價");
            sb.Append("							   ,cast(cast(A.PRICE as numeric(16,2)) as varchar) 台幣單價	   ,cast(cast(A.AMOUNT as numeric(16,2))*-1 as varchar) 台幣金額,");
            sb.Append("							   	   cast(cast(ISNULL(A.AMOUNT,0)-ISNULL(A.CostForAcc,0) as numeric(16,2))*-1 as varchar) 台幣毛利,P.PersonName 業務, '貸項'+A.BillNO 單據");
            sb.Append("                                                     From otherDB.CHIComp22.DBO.ComProdRec A                   ");
            sb.Append("                                                     left join otherDB.CHIComp22.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag   ");
            sb.Append("                                                          left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1   ");
            sb.Append("                                                                 left join otherDB.CHIComp22.DBO.comWareHouse W On  A.WareID=W.WareHouseID   ");
            sb.Append("                                                             left join otherDB.CHIComp22.DBO.OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO AND O.Flag =2   ");
            sb.Append("                                                                  left join otherDB.CHIComp22.DBO.OrdBillMain O2 On  O.BillNO=O2.BillNO AND O2.Flag =2    ");
            sb.Append("                                                             left join otherDB.CHIComp22.DBO.comCurrencySys C On  O2.CurrID=C.CurrencyID   ");
            sb.Append("															      left join  otherDB.CHIComp22.DBO.comPerson P ON (T.Salesman=P.PersonID) ");
            sb.Append("             															  LEFT JOIN ACMESQL02.DBO.OITM T11 ON (A.PRODID=T11.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("                                                          Where  A.FLAG IN (600,701)");


            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and   A.BILLDATE between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T.CustID='" + textBox7.Text.ToString() + "' ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox5.Text));

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox6.Text));





            SqlDataAdapter da = new SqlDataAdapter(command);



            DataSet ds = new DataSet();

            try
            {

                connection.Open();

                da.Fill(ds, "OINV");

            }

            finally
            {

                connection.Close();

            }

            return ds.Tables[0];



        }

        private System.Data.DataTable Gen_201005()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            string FS = textBox19.Text;

            StringBuilder sb = new StringBuilder();

            //彙總
            sb.Append("");
            sb.Append("																	                                              Select A.PRODID 項目料號,U.FULLNAME 客戶名稱,A.BILLDATE 過帳日期,CAST(A.Quantity AS INT) 數量  ");
            sb.Append("																												  ,cast(cast(A.PRICE as numeric(16,2)) as varchar) 單價,cast(cast(A.AMOUNT as numeric(16,2)) as varchar) 銷售金額");
            sb.Append("																												  ,cast(cast(A.CostForAcc as numeric(16,2)) as varchar) 成本, cast(cast(ISNULL(A.AMOUNT,0)-ISNULL(A.CostForAcc,0) as numeric(16,2)) as varchar) 毛利");
            sb.Append("																												  , CAST(((cast(ISNULL(A.AMOUNT,0)-ISNULL(A.CostForAcc,0) as numeric(16,2)))/CAST(A.Quantity AS INT)) AS INT) 每片毛利");
            sb.Append("																												  ,'銷貨單'+A.BillNO 單據");
            sb.Append("                                                                  From otherDB.CHIComp22.DBO.ComProdRec A                    ");
            sb.Append("                                                                  left join otherDB.CHIComp22.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag    ");
            sb.Append("                                                                       left join otherDB.CHIComp22.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1    ");
            sb.Append("                                                                       Where  A.FLAG=500 ");



            if (textBox17.Text != "" && textBox18.Text != "")
            {

                sb.Append(" and A.BILLDATE  between @DocDate1 and @DocDate2 ");
            }

            sb.Append(" and A.PRODID   in ( " + FS + ")");

            sb.Append("  ORDER BY A.PRODID ,A.BILLDATE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox17.Text));

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox18.Text));





            SqlDataAdapter da = new SqlDataAdapter(command);



            DataSet ds = new DataSet();

            try
            {

                connection.Open();

                da.Fill(ds, "OINV");

            }

            finally
            {

                connection.Close();

            }

            return ds.Tables[0];



        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox19.Text == "")
            {
                MessageBox.Show("請輸入項目號碼");
                return;
            }

            System.Data.DataTable t1 = Gen_201005();
            if (t1.Rows.Count > 0)
            {
                dataGridView10.DataSource = t1;

                string g = t1.Compute("AVG(每片毛利)", null).ToString();


                decimal sh = Convert.ToDecimal(g);

                label20.Text = "平均每片毛利 " + sh.ToString("#,##0");

            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
  
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

                cs = frm1.q;

            

                if (!String.IsNullOrEmpty(cs))
                {
                    textBox19.Text = cs;
                }
            }
        }


        private System.Data.DataTable MakeTabe()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("生產日期", typeof(string));
            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("銷售訂單", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("銷售單價", typeof(string));
            dt.Columns.Add("毛利", typeof(string));
            dt.Columns.Add("生產訂單", typeof(string));
            dt.Columns.Add("廠商", typeof(string));
            dt.Columns.Add("項目號碼", typeof(string));
            dt.Columns.Add("項目說明", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(int));
            return dt;
        }
        private System.Data.DataTable MakeTabeCHECKPAID()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("銷售單號", typeof(string));
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("型號", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("美金總額", typeof(string));
            dt.Columns.Add("台幣總額", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("到帳日期", typeof(string));
            dt.Columns.Add("入帳日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("付款方法", typeof(string));
            return dt;
        }


        private System.Data.DataTable GetTABLE2(string DocNum)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT W1.ItemCode 子件編號, T1.[ItemName] 產品名稱,");
            sb.Append("  單價 = CAST((SELECT abs(Convert(int,Sum(T7.[TransValue])))   FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum AND  T7.[AppObjLine] = W1.LineNum AND T7.[ItemCode] = W1.ItemCode AND  T7.[AppObjType] = 'C'  )/W1.[PlannedQty]   AS INT)");
            sb.Append("  FROM OWOR T0 ");
            sb.Append("  INNER JOIN WOR1 W1 ON W1.DocEntry=T0.DocNum");
            sb.Append("  Left JOIN OITM T1 ON T1.ItemCode= W1.ItemCode");
            sb.Append("  Where T0.DocNum=@DocNum");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));

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

        private System.Data.DataTable GetTABLE3(string BASEREF)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T0.DOCDATE,112)  FROM OIGN T0 ");
            sb.Append(" LEFT JOIN IGN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T1.BASEREF=@BASEREF");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BASEREF", BASEREF));

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
        private System.Data.DataTable GetTABLE3CHECKPAID(string DOCENTRY, string LINENUM)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select PAYCHECK 付款方法,TTDATE 入帳日期 from satt1 t0 left join satt t1 on (t0.ttcode=t1.ttcode) ");
            sb.Append("                                 left join satt2 t2 on (t2.ttcode=t0.ttcode AND T2.ID=T0.SEQNO) ");
            sb.Append("           WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        DateTime T1;


        private void dataGridView9_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView9.SelectedRows.Count > 0)
            {

                string da = dataGridView9.SelectedRows[0].Cells["項目群組"].Value.ToString();

                JOJO2IPGI a = new JOJO2IPGI();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

   
    }
}
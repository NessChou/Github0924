using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace ACME
{
    
    public partial class APCHOICE : Form
    {
        private string DRS = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql05";
        string ss1 = "";
        string ss2 = "";
        string ss3 = "";
        string ss4 = "";
        string ss5 = "";
        string ss6 = "";
        string ss7 = "";
        string ss8 = "";
        string ss9 = "";
        string ss10 = "";
        string ss101 = "";
        string ss11 = "";
        string ss12 = "";
        string BASEDOC = "";
        string BASELINE = "";
        string A1 = "";
        string A2 = "";
        public APCHOICE()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "進金生")
            {
                DELATC1();
                A1 = textBox5.Text;
                A2 = textBox6.Text;
                string YEAR = A1.Substring(0, 4);
                int M1 = Convert.ToInt16(A1.Substring(4, 2));
                int M2 = Convert.ToInt16(A2.Substring(4, 2));
                Category(8, "", "Account_Temp61");

                System.Data.DataTable k2 = GetODLNN3();
                if (k2.Rows.Count > 0)
                {

                    System.Data.DataTable FS = MakeTable();
                    DataRow dr = null;
                    for (int i = 0; i <= k2.Rows.Count - 1; i++)
                    {
                        dr = FS.NewRow();

                        string ITEMCODE = k2.Rows[i]["產品編號"].ToString().Trim();
                        string DOCENTRY = k2.Rows[i]["DOCENTRY"].ToString().Trim();
                        string DOCDATE = k2.Rows[i]["過帳日期"].ToString().Trim();
                        dr["產品編號"] = ITEMCODE;
                        dr["過帳日期"] = DOCDATE;
                        dr["年"] = Convert.ToInt16(DOCDATE.Substring(0, 4));
                        dr["月"] = Convert.ToInt16(DOCDATE.Substring(4, 2));
                        dr["客戶名稱"] = k2.Rows[i]["客戶名稱"].ToString().Trim();
                        dr["BU"] = k2.Rows[i]["BU"].ToString().Trim();
                        dr["群組"] = k2.Rows[i]["群組"].ToString().Trim();

                        dr["群組次分類"] = k2.Rows[i]["群組次分類"].ToString().Trim();
                        dr["MODEL"] = k2.Rows[i]["MODEL"].ToString().Trim();
                        dr["VER"] = k2.Rows[i]["VER"].ToString().Trim();
                        dr["品名敘述"] = k2.Rows[i]["品名敘述"].ToString().Trim();
                        dr["數量"] = k2.Rows[i]["數量"].ToString().Trim();
                        dr["金額"] = k2.Rows[i]["金額"].ToString().Trim();
                        dr["成本"] = k2.Rows[i]["成本"].ToString().Trim();
                        dr["毛利"] = k2.Rows[i]["毛利"].ToString().Trim();
                        dr["毛利率"] = k2.Rows[i]["毛利率"].ToString().Trim();
                        dr["美金單價"] = k2.Rows[i]["美金單價"].ToString().Trim();
                        dr["業務"] = k2.Rows[i]["業務"].ToString().Trim();
                        string CARDNAME = k2.Rows[i]["供應商"].ToString().Trim();

                        if (String.IsNullOrEmpty(CARDNAME))
                        {
                            System.Data.DataTable FF1 = GetODLNN4(ITEMCODE);
                            if (FF1.Rows.Count > 0)
                            {
                                CARDNAME = FF1.Rows[0][0].ToString();


                            }
                        }
                        if (String.IsNullOrEmpty(CARDNAME))
                        {

                            int F1 = ITEMCODE.IndexOf("ACME");
                            if (F1 != -1)
                            {
                                System.Data.DataTable FF2 = GetODLNN5(DOCENTRY,ITEMCODE);
                                if (FF2.Rows.Count > 0)
                                {
                                    int VIS = Convert.ToInt16(FF2.Rows[0][0]);

                                    VIS++;

                                    System.Data.DataTable FF3 = GetODLNN6(DOCENTRY, VIS.ToString());
                                    if (FF3.Rows.Count > 0)
                                    {
                                        string ITEMCODE2 = FF3.Rows[0][0].ToString();
                                        System.Data.DataTable FF4 = GetODLNN7(ITEMCODE2);
                                        if (FF4.Rows.Count > 0)
                                        {
                                            CARDNAME = FF4.Rows[0][0].ToString();
                                        }
                                        else
                                        {
                                            VIS++;
                                            System.Data.DataTable FF5 = GetODLNN6(DOCENTRY, VIS.ToString());
                                            if (FF5.Rows.Count > 0)
                                            {
                                                CARDNAME = FF5.Rows[0][0].ToString();
                                            }
                                        }

                                    }

                                }



                            
                            }


                        }
                        dr["供應商"] = CARDNAME;
                        FS.Rows.Add(dr);
                    }

          



                    for (int i = 0; i <= FS.Rows.Count - 1; i++)
                    {

                        DataRow dd1 = FS.Rows[i];
                        string CARDNAME = dd1["供應商"].ToString();
                        string 產品編號 = dd1["產品編號"].ToString();
                        string 客戶名稱 = dd1["客戶名稱"].ToString();
                        string 過帳日期 = dd1["過帳日期"].ToString();
                        int 年 = Convert.ToInt16(dd1["年"]);
                        int 月 = Convert.ToInt16(dd1["月"]);
                        int 數量 = Convert.ToInt32(dd1["數量"]);
                        int 金額 = Convert.ToInt32(dd1["金額"]);
                        int 成本 = Convert.ToInt32(dd1["成本"]);
                        int 毛利 = Convert.ToInt32(dd1["毛利"]);
                        string BU = dd1["BU"].ToString();
                        string 群組 = dd1["群組"].ToString();
                        string 群組次分類 = dd1["群組次分類"].ToString();
                        string MODEL = dd1["MODEL"].ToString();
                        string VER = dd1["VER"].ToString();

                        string 品名敘述 = dd1["品名敘述"].ToString();
                        string 毛利率 = dd1["毛利率"].ToString();
                        string 美金單價 = dd1["美金單價"].ToString();
                        string 業務 = dd1["業務"].ToString();
                        string 供應商 = dd1["供應商"].ToString();


                        AddATC1(供應商, 產品編號, 過帳日期, 年, 月, BU, 群組, 群組次分類, MODEL, VER, 品名敘述, 數量, 金額, 成本, 毛利, 毛利率, 美金單價, 業務, 客戶名稱);
                    }

                    System.Data.DataTable K3 = GetODLNN4F();
                    if (K3.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = GETJO();
                        dataGridView2.DataSource = K3;
                        dataGridView3.DataSource = GETJO2();
                        for (int i = 1; i <= dataGridView2.Columns.Count - 1; i++)
                        {
                            DataGridViewColumn col = dataGridView2.Columns[i];


                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                            col.DefaultCellStyle.Format = "#,##0";
                        }
                    }
                }
            }
            if (comboBox1.Text == "達睿生")
            {
                TRUNG();

                System.Data.DataTable k1 = GetODLNN();
                if (k1.Rows.Count > 0)
                {
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {
                        decimal PRICE = 0;
                        DataRow dd1 = k1.Rows[i];
                        string 項目料號 = dd1["項目料號"].ToString();
                        string 客戶編號 = dd1["客戶編號"].ToString();
                        string 客戶名稱 = dd1["客戶名稱"].ToString();
                        string 過帳日期 = dd1["過帳日期"].ToString();
                        string 訂購單號 = dd1["訂購單號"].ToString();
                        string BU = dd1["BU"].ToString();
                        if (comboBox1.Text == "進金生")
                        {
                            PRICE = Convert.ToDecimal(dd1["美金單價"]);
                        }
                        string MODEL = dd1["MODEL"].ToString();
                        int 數量 = Convert.ToInt32(dd1["數量"]);
                        int 銷售金額 = Convert.ToInt32(dd1["銷售金額"]);
                        int 成本 = Convert.ToInt32(dd1["成本"]);
                        string 業務 = dd1["業務"].ToString();
                        string AR = dd1["AR"].ToString();
                        string CHI = dd1["CHI"].ToString();
                        int LINENUM = Convert.ToInt32(dd1["LINENUM"]);
                        int BUAMT = 0;

             

                        AddG(項目料號, 客戶編號, 客戶名稱, 數量, 銷售金額, 成本, 業務, 過帳日期, BUAMT, AR, CHI, LINENUM, BU, PRICE);


                    }
                }
                System.Data.DataTable k2 = GetODLNN2();
                if (k2.Rows.Count > 0)
                {

                    dataGridView1.DataSource = k2;
                }
            }
            if (comboBox1.Text == "CHOICE")
            {
                string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                System.Data.DataTable k3 = GETCHOICE(strCn);
                if (k3.Rows.Count > 0)
                {

                    dataGridView1.DataSource = k3;
                }
            }
            if (comboBox1.Text == "宇豐")
            {
                string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                System.Data.DataTable k3 = GETCHOICE(strCn);
                if (k3.Rows.Count > 0)
                {

                    dataGridView1.DataSource = k3;
                }
            }
            if (comboBox1.Text == "IPGI")
            {
                string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                System.Data.DataTable k3 = GETCHOICE(strCn);
                if (k3.Rows.Count > 0)
                {

                    dataGridView1.DataSource = k3;
                }
            }
            //IPGI 
        }
        public System.Data.DataTable GetSAP(string BILLNO, string ProdID, int Quantity)
        {

            SqlConnection MyConnection = new SqlConnection(DRS);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CAST(T1.LINETOTAL/T9.RATE AS INT) AMT FROM oinv T0 ");
            sb.Append(" INNER JOIN inv1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
            sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
            sb.Append(" LEFT JOIN ORTT T9 ON (Convert(varchar(8),T0.DOCDATE,112)=Convert(varchar(8),T9.RATEDATE,112) AND T9.CURRENCY='NTD')");
            sb.Append(" WHERE T8.U_SAPDOC=@BILLNO AND T1.ITEMCODE=@ProdID AND CAST(T1.Quantity AS INT)=@Quantity ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));

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
        public  System.Data.DataTable GetODLNN()
        {

            SqlConnection connection = null;

            if (comboBox1.Text == "進金生")
            {
                connection = globals.shipConnection;
            }
            if (comboBox1.Text == "達睿生")
            {
                connection = new SqlConnection(DRS);
            }
            StringBuilder sb = new StringBuilder();

            sb.Append("                                                 SELECT  T2.ITEMCODE 項目料號,T0.CARDCODE 客戶編號,case when T0.cardname like  '%TOP GARDEN INT%' then 'TOP GARDEN' when T0.cardname like  '%CHOICE CHANNEL%' then 'CHOICE' when T0.cardname like  '%Infinite Power Group%' then 'INFINITE' when T0.cardname like  '%宇豐光電股份有限公司%' then '宇豐' when T0.cardname like  '%達睿生%' then 'DRS' else t0.cardname end+CASE ISNULL(T0.U_BENEFICIARY,'') WHEN '' THEN '' ELSE '-'+T0.U_BENEFICIARY END 客戶名稱,Convert(varchar(8),T0.[DocDate],112)  過帳日期,   ");
            sb.Append("                                                             (CAST(T2.Quantity AS INT)) 數量,T2.[Price] 單價,CASE T5.CURRENCY WHEN 'USD' THEN ISNULL(T5.[Price],0) ELSE 0 END 美金單價,    ");
            sb.Append("                                                      CAST(T2.LineTotal AS INT) 銷售金額,CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END 成本,  ");
            sb.Append("                                                      CAST((T2.LineTotal) - (CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END) AS INT) 毛利, T3.SLPNAME 業務,  ");
            sb.Append("                                                     CAST(T0.DOCENTRY AS VARCHAR) AR, ISNULL(t8.U_CHI_NO,'') CHI,T2.LINENUM,T5.DOCENTRY 訂購單號,T11.U_TMODEL MODEL,SUBSTRING(ITMSGRPNAM,4,20) BU FROM OINV T0   ");
            sb.Append("                                                      INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry   ");
            sb.Append("                               left join dln1 t4 on (t2.baseentry=T4.docentry and  t2.baseline=t4.linenum  and t2.basetype='15')  ");
            sb.Append("                              left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15') left join ordr t8 on (t8.docentry=T5.docentry  )  ");
            sb.Append("                                         INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode   ");
            sb.Append("                                   left JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE   ");
            sb.Append("                               left JOIN OITB T12 ON T12.ITMSGRPCOD = T11.ITMSGRPCOD   ");
            sb.Append("                                                      WHERE T0.[DocType] ='I'   ");
            if (comboBox1.Text == "進金生")
            {
                sb.Append("                                and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.U_IN_BSTYC <> '1'   AND T11.itmsgrpcod IN ('1032','1034')      ");
            }
            sb.Append("               AND  Convert(varchar(8),T0.[DocDate],112) between @DATE and  @DATE1 ");
            sb.Append("           UNION ALL ");
            sb.Append("                                                               SELECT  T2.ITEMCODE 項目料號,T0.CARDCODE 客戶編號,case when T0.cardname like  '%TOP GARDEN INT%' then 'TOP GARDEN' when T0.cardname like  '%CHOICE CHANNEL%' then 'CHOICE' when T0.cardname like  '%Infinite Power Group%' then 'INFINITE' when T0.cardname like  '%宇豐光電股份有限公司%' then '宇豐' when T0.cardname like  '%達睿生%' then 'DRS' else t0.cardname end+CASE ISNULL(T0.U_BENEFICIARY,'') WHEN '' THEN '' ELSE '-'+T0.U_BENEFICIARY END 客戶名稱,Convert(varchar(8),T0.[DocDate],112)  過帳日期,    ");
            sb.Append("                                                                           (CAST(T2.Quantity AS INT)) 數量,T2.[Price] 單價,T2.U_ACME_INV 美金單價,     ");
            sb.Append("                                                                    CAST(T2.LineTotal AS INT)*-1 銷售金額,CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END*-1 成本,   ");
            sb.Append("                                                                    CAST((T2.LineTotal) - (CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END) AS INT)*-1 毛利, T3.SLPNAME 業務,   ");
            sb.Append("                                                                   CAST(T0.DOCENTRY AS VARCHAR) AR, '' CHI,T2.LINENUM,T0.DOCENTRY  訂購單號,T11.U_TMODEL MODEL,SUBSTRING(ITMSGRPNAM,4,20) BU FROM ORIN T0    ");
            sb.Append("                                                                    INNER JOIN RIN1 T2 ON T0.DocEntry = T2.DocEntry    ");
            sb.Append("                                                       INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode    ");
            sb.Append("                                                 left JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE    ");
            sb.Append("                                             left JOIN OITB T12 ON T12.ITMSGRPCOD = T11.ITMSGRPCOD    ");
            sb.Append("                                                                    WHERE T0.[DocType] ='I'    ");
            if (comboBox1.Text == "進金生")
            {
                sb.Append("                                              and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.U_IN_BSTYC <> '1'   AND T11.itmsgrpcod IN ('1032','1034')       ");
            }
            sb.Append("               AND  Convert(varchar(8),T0.[DocDate],112) between @DATE and  @DATE1 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DATE", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DATE1", textBox6.Text));

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
        public System.Data.DataTable GETCHOICE(string strCn)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select A.ProdID 產品編號,A.BillDate  過帳日期 ,U.FullName 客戶名稱,A.ProdID 產品編號,A.ProdName 產品名稱,");
            sb.Append(" A.Quantity 數量,A.MLAmount 金額,(A.CostForAcc) 成本 ,(A.MLAmount-A.CostForAcc) 毛利, ");
            sb.Append(" CAST(CAST(( CASE (A.MLAmount) WHEN 0 THEN 0 ELSE (A.MLAmount-A.CostForAcc)/(A.MLAmount) END)*100 AS decimal(10,2)) AS VARCHAR)+'%' 毛利率,");
            sb.Append(" C.CurrencyName+''+ cast(cast(O.PRICE as numeric(16,2)) as varchar) 訂單單價,P.PersonName 業務");
            sb.Append(" ,(SELECT TOP 1  U.FullName      FROM ComProdRec S      ");
            sb.Append(" left join comBillAccounts T ON S.BillNO=T.FundBillNo AND S.Flag=T.Flag        ");
            sb.Append(" left join comCustomer U On  U.ID=T.CustID  AND U.Flag =2  ");
            sb.Append(" WHERE S.ProdID=A.ProdID AND S.Flag IN (100,200)");
            sb.Append(" ORDER BY S.BillDate  DESC) 供應商");
            sb.Append(" From ComProdRec A           ");
            sb.Append(" left join comBillAccounts T ON A.BillNO=T.FundBillNo AND A.Flag=T.Flag         ");
            sb.Append(" left join comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" left join OrdBillSub O On  A.FromNO=O.BillNO AND A.FromRow=O.RowNO  AND O.Flag =2  ");
            sb.Append("      left join OrdBillMain O2 On  O.BillNO=O2.BillNO  AND O2.Flag =2 ");
            sb.Append(" left join comPerson P ON (T.Salesman =P.PersonID)    ");
            sb.Append(" left join comCurrencySys C On  O2.CurrID=C.CurrencyID ");
            sb.Append(" Where A.Flag IN (500,600)  ");
            sb.Append(" AND CAST(A.BillDate AS VARCHAR) BETWEEN @DATE and  @DATE1 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DATE", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DATE1", textBox6.Text));

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
        public System.Data.DataTable GetODLNN2()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                                                                   SELECT T0.ITEMCODE,T0.DOCDATE 過帳日期,T0.CARDNAME 客戶名稱,BU,T1.U_GROUP 群組, ");
            sb.Append("                                                                        T3.ITEM1  群組次分類,T1.U_TMODEL MODEL,   ");
            sb.Append("                                                                  CASE T1.ITEMNAME WHEN 'one shot order' THEN T2.DSCRIPTION ELSE T1.ITEMNAME END 品名敘述,    ");
            sb.Append("                                                                     T0.QUANTITY 數量,T0.AMT 金額,T0.COST 成本,T0.AMT-T0.COST 毛利,    ");
            sb.Append("                                                                     CAST(CAST(CASE WHEN T0.AMT = 0 THEN '0' WHEN T0.AMT < 0 THEN ((CAST(T0.AMT-T0.COST AS DECIMAL))/T0.AMT)*100*-1  ELSE    ");
            sb.Append("                                           ((CAST(T0.AMT-T0.COST AS DECIMAL))/T0.AMT)*100 END AS DECIMAL(10,2)) AS VARCHAR)+'%' 毛利率,    ");
            sb.Append("                                                                     T0.PRICE 美金單價,T0.SALES 業務,    ");
            sb.Append("                                                                     (SELECT TOP 1 CARDNAME  FROM ACMESQL02.DBO.POR1 A    ");
            sb.Append("                                                                     LEFT JOIN ACMESQL02.DBO.OPOR B ON (A.DOCENTRY=B.DOCENTRY)    ");
            sb.Append("                                                                      WHERE ITEMCODE=T0.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS    ");
            sb.Append("                                                                     ORDER BY T0.DOCDATE DESC) 供應商,CASE WHEN T2.PRICE=0 THEN T4.U_ACME_USER END 'FOC 的出貨單'     ");
            sb.Append("                                                                      FROM SALES_BUREPORT T0    ");
                sb.Append("                                                                     LEFT JOIN ACMESQL05.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
                sb.Append("                                                                     LEFT JOIN ACMESQL05.DBO.INV1 T2 ON (T0.AR=T2.DOCENTRY AND T0.LINENUM=T2.LINENUM)    ");
            sb.Append("                                           LEFT JOIN ACMESQLSP.DBO.WH_ITM1 T3 ON (T1.U_GROUP=T3.VALUE1)   ");
            sb.Append("      LEFT JOIN ACMESQL02.DBO.OINV T4 ON (T2.DOCENTRY=T4.DOCENTRY)   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DATE", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DATE1", textBox6.Text));

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

        public System.Data.DataTable GETJO()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMCODE 產品編號,DYEAR 年,DMONTH 月,DOCDATE 過帳日期,CARDNAME2 客戶名稱,BU,DG 群組,DG2 群組次分類,MODEL,VER");
            sb.Append("  ,ITEMNAME  品名敘述,GQTY 數量,GTOTAL 金額,GVALUE 成本,GM 毛利,GM2  毛利率,USD 美金單價,SALES 業務,CARDNAME 供應商 FROM AP_JO WHERE  USERS=@USERS  ");
            if (comboBox2.Text != "")
            {
                sb.Append("  AND CARDNAME=@CARDNAME ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME",comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GETJO2()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 20 CARDNAME2 客戶名稱,SUM(GQTY) 數量 FROM AP_JO WHERE  USERS=@USERS  GROUP BY CARDNAME2 ORDER BY SUM(GQTY) DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetODLNN3()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT T0.ITEMCODE 產品編號,Convert(varchar(8),T0.DDATE,112) 過帳日期,T0.CARDNAME 客戶名稱,SUBSTRING(ITMSGRPNAM,4,5) BU,T1.U_GROUP 群組, ");
            sb.Append("               CASE WHEN ISNULL(T1.U_JGROUP,'') = '' THEN T3.ITEM1 ELSE T1.U_JGROUP END  COLLATE  Chinese_Taiwan_Stroke_CI_AS 群組次分類,T1.U_TMODEL MODEL,T1.U_VERSION VER  ");
            sb.Append("               ,CASE ISNULL(T1.ITEMNAME,'') WHEN 'one shot order' THEN T0.ITEMNAME WHEN '' THEN T0.ITEMNAME  ELSE T1.ITEMNAME END COLLATE  Chinese_Taiwan_Stroke_CI_AS 品名敘述, ");
            sb.Append("                 T0.GQTY 數量,T0.GTOTAL 金額,T0.GVALUE 成本,T0.GTOTAL-T0.GVALUE 毛利, ");
            sb.Append("                CAST(CAST(CASE WHEN T0.GTOTAL = 0 THEN '0' WHEN T0.GTOTAL < 0 THEN ((CAST(T0.GTOTAL-T0.GVALUE AS DECIMAL))/T0.GTOTAL)*100*-1  ELSE      ");
            sb.Append("               ((CAST(T0.GTOTAL-T0.GVALUE AS DECIMAL))/T0.GTOTAL)*100 END AS DECIMAL(10,2)) AS VARCHAR)+'%' 毛利率 ");
            sb.Append("               ,T5.PRICE 美金單價,T0.SALES 業務,(SELECT TOP 1 CARDNAME  FROM ACMESQL02.DBO.PCH1 A      ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OPCH B ON (A.DOCENTRY=B.DOCENTRY)      ");
            sb.Append("               WHERE ITEMCODE=T0.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS      ");
            sb.Append("               ORDER BY 	 B.DOCDATE DESC) 供應商 ");
            sb.Append("               ,CASE WHEN T0.GTOTAL=0 THEN T6.U_ACME_USER END 'FOC 的出貨單',T0.DOCENTRY  FROM Account_Temp61 T0 ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OITB T2 ON (T1.itmsgrpcod = T2.itmsgrpcod) ");
            sb.Append("               LEFT JOIN ACMESQLSP.DBO.WH_ITM1 T3 ON (T1.U_GROUP=T3.VALUE1)     ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.dln1 t4 on (T0.BASEDOC=T4.docentry and  T0.BASELINE=t4.linenum  ) ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.odln t9 on (t4.docentry=T9.docentry ) ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OINV T6 ON (T0.DOCENTRY=T6.DOCENTRY) WHERE CARDGROUP=103     ");
   

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
   
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

        public System.Data.DataTable GetODLNN4F()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDNAME 廠商,  SUM(T0.GQTY) 數量,SUM(T0.GTOTAL) 金額,SUM(T0.GVALUE) 成本,SUM(T0.GTOTAL-T0.GVALUE) 毛利,");
            sb.Append(" CAST(CAST(CASE WHEN SUM(T0.GTOTAL) = 0 THEN '0' WHEN SUM(T0.GTOTAL) < 0 THEN ((CAST(SUM(T0.GTOTAL-T0.GVALUE)  AS DECIMAL))/SUM(T0.GTOTAL))*100*-1  ELSE       ");
            sb.Append(" ((CAST(SUM(T0.GTOTAL-T0.GVALUE) AS DECIMAL))/SUM(T0.GTOTAL) )*100 END AS DECIMAL(10,2)) AS VARCHAR)+'%' 毛利率   FROM AP_JO T0");
            sb.Append(" WHERE CARDNAME <> '' AND  USERS=@USERS  ");
    
            sb.Append("  GROUP BY CARDNAME HAVING SUM(T0.GTOTAL) <>0");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

        public System.Data.DataTable GetCARDNAME()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT CARDNAME  FROM AP_JO   WHERE USERS=@USERS AND ISNULL(CARDNAME,'') <> '' ORDER BY CARDNAME");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public System.Data.DataTable GetODLNN4(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 U_CARDNAME FROM OWOR WHERE ITEMCODE=@ITEMCODE	ORDER BY DueDate DESC");
          
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetODLNN5(string DOCENTRY, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT VisOrder  FROM INV1 WHERE DOCENTRY=@DOCENTRY  AND ITEMCODE=@ITEMCODE  ORDER BY VisOrder ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public System.Data.DataTable GetODLNN6(string DOCENTRY, string VisOrder)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE  FROM INV1 WHERE DOCENTRY=@DOCENTRY  AND VisOrder=@VisOrder  AND TreeType ='I' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@VisOrder", VisOrder));
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

        public System.Data.DataTable GetODLNN7(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 CARDNAME  FROM ACMESQL02.DBO.PCH1 A       ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPCH B ON (A.DOCENTRY=B.DOCENTRY)       ");
            sb.Append(" WHERE ITEMCODE=@ITEMCODE  ");
            sb.Append(" ORDER BY 	 B.DOCDATE DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public void AddG(string ITEMCODE, string CARDCODE, string CARDNAME, int QUANTITY, int AMT, int COST, string SALES, string DOCDATE, int BUAMT, string AR, string CHI, int LINENUM, string BU, decimal PRICE)
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into SALES_BUREPORT(ITEMCODE,CARDCODE,CARDNAME,QUANTITY,AMT,COST,SALES,DOCDATE,BUAMT,AR,CHI,LINENUM,BU,PRICE) values(@ITEMCODE,@CARDCODE,@CARDNAME,@QUANTITY,@AMT,@COST,@SALES,@DOCDATE,@BUAMT,@AR,@CHI,@LINENUM,@BU,@PRICE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@QUANTITY", QUANTITY));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@COST", COST));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@BUAMT", BUAMT));
            command.Parameters.Add(new SqlParameter("@AR", AR));
            command.Parameters.Add(new SqlParameter("@CHI", CHI));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
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
        public void TRUNG()
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand(" TRUNCATE TABLE SALES_BUREPORT ", connection);
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
        


        private void APCHOICE_Load(object sender, EventArgs e)
        {
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();

            comboBox1.Text = "進金生";

        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex  == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }

            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }
        private void Category(int year, string ff, string TABLE)
        {
        
                AddAUOGD1(TABLE);
            
            System.Data.DataTable dt = null;


            dt = GetMenu.GetSAPRevenue(A1, A2);

            
            System.Data.DataTable dtCost = MakeTableCombine2();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string 單據;
            string 科目代號;

            Int32 單號;
            DateTime 日期;

            Int32 基礎單號;
            Int32 基礎列;

            //20080904
            //宣告 DuplicateKey 來檢查
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                單據 = Convert.ToString(dt.Rows[i]["單別"]);
                單號 = Convert.ToInt32(dt.Rows[i]["DocNum"]);
                日期 = Convert.ToDateTime(dt.Rows[i]["日期"]);
                科目代號 = Convert.ToString(dt.Rows[i]["科目代號"]);

                //if (單號 == 29159)
                //{
                //    MessageBox.Show("A");
                //}

                dtDoc = GetSAPDoc(單據, 單號, 科目代號, A1, A2);


                基礎單號 = -1;
                基礎列 = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();

                    dr["收入單據"] = 單據;
                    dr["收入單號"] = 單號;


                    dr["日期"] = 日期;
                    dr["科目代號"] = 科目代號;
                    dr["客戶編號"] = "'" + Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["客戶名稱"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["產品編號"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["產品名稱"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);



                    if (year == 7)
                    {
                        dr["客戶群組"] = Convert.ToString(dt.Rows[i]["部門"]);
                    }
                    else
                    {
                        dr["客戶群組"] = Convert.ToString(dtDoc.Rows[j]["GROUPCODE"]);
                    }
                    string D = dtDoc.Rows[j]["LineTotal"].ToString();
                    dr["數量"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["單價"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["金額"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]);
                    dr["單號總成本"] = 0;
                    dr["項目成本"] = 0;


                    //業務員
                    dr["業務員編號"] = Convert.ToString(dt.Rows[i]["業務員編號"]);
                    dr["姓名"] = Convert.ToString(dt.Rows[i]["姓名"]);



                    if (單據 == "AR" || 單據 == "貸項" || 單據 == "AR預")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            基礎單號 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["基礎單號"] = 基礎單號;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            基礎列 = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["基礎列"] = 基礎列;
                        }

                    }


                    //總收入寫在最後一筆
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (單據 == "AR" || 單據 == "AR-服務" || 單據 == "AR預")
                        {
                            dr["單號總收入"] = Convert.ToInt32(dt.Rows[i]["總成本"]);

                        }
                        else
                        {

                            dr["單號總收入"] = Convert.ToInt32(dt.Rows[i]["總成本"]) * (-1);
                        }
                    }

                    if (單據 == "貸項" || 單據 == "貸項-服務" || 單據 == "銷退" || 單據 == "JE")
                    {
                        dr["金額"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  數量改成 負數
                        dr["數量"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }

                    if (單據 == "AR" || 單據 == "AR預")
                    {
                        //0303
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (基礎單號.ToString() == "3169" && 單號.ToString() == "3429")
                            {
                                dr["項目成本"] = 0;
                                dr["單號總成本"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (基礎單號.ToString() == "3167" && 單號.ToString() == "3404")
                            {
                                dr["項目成本"] = 0;
                                dr["單號總成本"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }

                            dtDocLine = GetSAPDocByLine("交貨", 基礎單號, 基礎列);

                            dr["成本單據"] = "交貨";
                            dr["成本單號"] = 基礎單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                               * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["單號總成本"] = 0;
                                    if (單號 != DuplicateKey)
                                    {

                                        dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }
                                    DuplicateKey = 單號;
                                }
                                //20091204一對多
                                if (基礎單號.ToString() == "5394" && 單號.ToString() == "5673")
                                {

                                    dr["單號總成本"] = 2111964;
                                    dr["單號總收入"] = 0;

                                }
                                //2010331多對一
                                if (單號.ToString() == "6975")
                                {
                                    dr["單號總成本"] = 0;

                                }
                                //2010409訂單轉AR
                                if (單號.ToString() == "7022")
                                {
                                    dr["單號總成本"] = "5476";
                                    dr["項目成本"] = "5476";

                                }

                                //20150506 AR跟AR預共存
                                if (基礎單號.ToString() == "26223" && 基礎列.ToString() == "0")
                                {
                                    dr["單號總成本"] = "1005608";


                                }

                                if (基礎單號.ToString() == "26441" && Convert.ToInt16(dtDoc.Rows[j]["Quantity"]) == 4)
                                {
                                    dr["單號總成本"] = "261458";


                                }
                                System.Data.DataTable GT = TF(基礎單號.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["序號"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (單號.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["單號總成本"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }

                            //成本必須來自至於分錄

                        }
                        //沒有基礎單號
                        else
                        {
                            //成本資料為自已
                            dr["成本單據"] = 單據;
                            dr["成本單號"] = 單號;


                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            if (dtDocLine != null)
                            {

                                if (dtDocLine.Rows.Count == 1)
                                {
                                    dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                                   * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                        {
                                            dr["單號總成本"] = 0;
                                        }
                                        else
                                        {
                                            //反回去找銷貨成本
                                            System.Data.DataTable dtSalesCost = GetSalesCost(單號.ToString());
                                            try
                                            {
                                                dr["單號總成本"] = Convert.ToInt32(dtSalesCost.Rows[0]["總成本"]);
                                            }
                                            catch
                                            {
                                                dr["單號總成本"] = 0;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    //Rows.Count =0 成本為零
                                    dr["項目成本"] = 0;
                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }

                        }

                    }

                    // 3 月案例沒有來源單號

                    //20081007 增加銷退..成本為負

                    if (單據 == "貸項" || 單據 == "貸項-服務" || 單據 == "銷退")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////要判斷來源單種類
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));

                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            //成本資料為自已

                            dr["成本單據"] = 單據;

                            dr["成本單號"] = 單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (單號 != DuplicateKey)
                                        {

                                            dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                        }
                                        DuplicateKey = 單號;

                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;

                                    //20081231
                                    if (單號 != DuplicateKey)
                                    {

                                        dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }
                                    DuplicateKey = 單號;
                                }
                            }


                        }
                        else
                        {


                            dtDocLine = GetSAPDocByLine(單據, 單號);

                            //成本資料為自已

                            dr["成本單據"] = 單據;

                            dr["成本單號"] = 單號;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["項目成本"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["總成本"]))
                                    {
                                        dr["單號總成本"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (單號 != DuplicateKey)
                                        {

                                            dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                        }
                                        DuplicateKey = 單號;

                                        // dr["單號總成本"] = Convert.ToInt32(dtDocLine.Rows[0]["總成本"]);
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 成本為零
                                dr["項目成本"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["單號總成本"] = 0;
                                }
                            }
                        }
                    }

                    dtCost.Rows.Add(dr);


                }
            }


            dataGridView8.DataSource = dtCost;
            for (int i = 0; i <= dataGridView8.Rows.Count - 1; i++)
            {


                ss1 = dataGridView8.Rows[i].Cells["客戶編號"].Value.ToString();
                ss2 = dataGridView8.Rows[i].Cells["客戶名稱"].Value.ToString();
                ss3 = dataGridView8.Rows[i].Cells["姓名"].Value.ToString();
                ss4 = dataGridView8.Rows[i].Cells["數量"].Value.ToString();
                ss5 = dataGridView8.Rows[i].Cells["單號總收入"].Value.ToString();
                ss6 = dataGridView8.Rows[i].Cells["單號總成本"].Value.ToString();
                ss7 = dataGridView8.Rows[i].Cells["收入單號"].Value.ToString();
                //if (ss7 == "24252")
                //{
                //    MessageBox.Show("A");
                //}
                ss8 = dataGridView8.Rows[i].Cells["科目代號"].Value.ToString();
                ss9 = dataGridView8.Rows[i].Cells["客戶群組"].Value.ToString();
                ss10 = dataGridView8.Rows[i].Cells["項目成本"].Value.ToString();
                ss101 = dataGridView8.Rows[i].Cells["金額"].Value.ToString();
                ss11 = dataGridView8.Rows[i].Cells["產品編號"].Value.ToString();
                ss12 = dataGridView8.Rows[i].Cells["產品名稱"].Value.ToString();
                BASEDOC = dataGridView8.Rows[i].Cells["基礎單號"].Value.ToString();
                BASELINE = dataGridView8.Rows[i].Cells["基礎列"].Value.ToString();
                //20150324 太陽能
                if (ss4 == "1")
                {
                    if (ss7 == "24018")
                    {
                        ss5 = "31600";
                    }
                    if (ss7 == "25636")
                    {
                        ss5 = "30000";
                    }
                }

                //鈺緯
                if (ss7 == "26976")
                {
                    if (ss4 == "206")
                    {
                        ss6 = "575134";
                    }

                    if (ss4 == "66")
                    {
                        ss6 = "143783";
                    }

                    if (ss4 == "110")
                    {
                        ss6 = "239639";
                    }
                }
                if (ss7 == "28510")
                {
                    if (ss4 == "108")
                    {
                        ss6 = "221455";
                    }
                }

                //20151207
                if (ss7 == "30027")
                {
                    if (ss4 == "1")
                    {
                        ss101 = "5790";
                    }
                }
                //2050901
                if (ss7 == "27312")
                {
                    if (ss4 == "264")
                    {
                        ss6 = "571912";
                        ss101 = "608889";

                    }
                }
                if (ss7 == "27487")
                {
                    if (ss4 == "100")
                    {
                        ss101 = "265338";
                    }
                    if (ss4 == "900")
                    {
                        ss101 = "2388046";
                        ss6 = "2450000";
                    }
                    if (ss4 == "200")
                    {
                        ss101 = "530677";
                    }
                }
                if (ss7 == "30013")
                {
                    if (ss4 == "160")
                    {
                        ss6 = "436545";
                    }

                    if (ss4 == "16")
                    {
                        ss101 = "38166";
                    }
                }


                //20160513

                if (ss7 == "29133")
                {
                    if (ss4 == "384")
                    {
                        ss6 = "787394";
                    }
                }

                if (ss7 == "30993")
                {
                    if (ss4 == "325")
                    {
                        ss6 = "1220106";
                    }
                }
                if (ss7 == "31526")
                {
                    if (ss4 == "240")
                    {
                        ss6 = "654818";
                    }
                }
                if (ss7 == "31702")
                {
                    if (ss4 == "3212")
                    {
                        ss6 = "9366132";

                        ss101 = "9484876";
                    }
                }
                if (ss7 == "29133")
                {
                    if (ss4 == "960")
                    {
                        ss6 = "2070704";
                    }
                }
                //臨時戶
                if (ss7 == "29010")
                {
                    if (ss101 == "79625")
                    {
                        ss6 = "69647";
                    }
                }


                //豐藝
                if (ss7 == "32778")
                {

                    ss6 = "0";

                }

                //20150625 太陽能差1元
                if (ss7 == "22848")
                {
                    ss101 = "3648959";
                }
                if (ss7 == "27445")
                {
                    ss101 = "1107541";
                }


                if (ss7 == "31193")
                {
                    if (ss101 == "13272")
                    {
                        ss6 = "7256";
                    }
                }
                if (ss7 == "31194")
                {
                    if (ss101 == "26543")
                    {
                        ss6 = "14513";
                    }
                    if (ss101 == "4976")
                    {
                        ss101 = "4977";
                    }
                }
                if (ss7 == "31195")
                {
                    if (ss101 == "39815")
                    {
                        ss6 = "21769";
                    }
                }
                if (ss7 == "33318")
                {
                    ss101 = "19489";
                    ss5 = "19489";
                }
                if (ss7 == "1532" || ss7 == "1533")
                {
                    if (ss2 == "YSHENG  HOLDINGS  LIMITED")
                    {
                        ss101 = "0";
                        ss5 = "0";
                    }
                }

                //20180104
                if (ss7 == "40574")
                {
                    if (ss4 == "12")
                    {
                        ss6 = "41747";
                        ss10 = "41747";
                    }
                }
                //27445     
                DateTime dd = Convert.ToDateTime(dataGridView8.Rows[i].Cells["日期"].Value);

                if (String.IsNullOrEmpty(ss6))
                {
                    ss6 = "0";
                }
                if (String.IsNullOrEmpty(ss5))
                {
                    ss5 = "0";
                }
      
                    AddAUOGD61(ss11, ss12, ss1, ss2, ss3, ss4, ss101, ss10, dd, ss7, ss8, ss9, ss5, ss6, BASEDOC, BASELINE);
             

            }
        }
        public void AddAUOGD1(string TABLE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table " + TABLE + " ", connection);
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


        public void AddAUOGD61(string ITEMCODE, string ITEMNAME, string CARDCODE, string CARDNAME, string SALES, string GQty, string GTotal, string GValue, DateTime 日期, string DOCENTRY, string ACCOUNT, string CARDGROUP, string GSUMTotal, string GSUMValue, string BASEDOC, string BASELINE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Account_Temp61(ITEMCODE,ITEMNAME,CARDCODE,CARDNAME,SALES,GQty,GTotal,GValue,DDATE,DOCENTRY,ACCOUNT,CARDGROUP,GSUMTotal,GSUMValue,BASEDOC,BASELINE) values(@ITEMCODE,@ITEMNAME,@CARDCODE,@CARDNAME,@SALES,@GQty,@GTotal,@GValue,@DDATE,@DOCENTRY,@ACCOUNT,@CARDGROUP,@GSUMTotal,@GSUMValue,@BASEDOC,@BASELINE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@GQty", GQty));
            command.Parameters.Add(new SqlParameter("@GTotal", GTotal));
            command.Parameters.Add(new SqlParameter("@GValue", GValue));
            command.Parameters.Add(new SqlParameter("@DDATE", 日期));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", CARDGROUP));
            command.Parameters.Add(new SqlParameter("@GSUMTotal", GSUMTotal));
            command.Parameters.Add(new SqlParameter("@GSUMValue", GSUMValue));
            command.Parameters.Add(new SqlParameter("@BASEDOC", BASEDOC));
            command.Parameters.Add(new SqlParameter("@BASELINE", BASELINE));
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
        private System.Data.DataTable GetSalesCost(string BaseRef)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (T1.[Debit] - T1.[Credit])  總成本");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" WHERE T0.TransType=13 and  T0.BaseRef=@BaseRef and T1.[Account] like '5110%' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@BaseRef", BaseRef));

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
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {


            }
            else if (DocKind == "銷退")
            {


            }
            else if (DocKind == "貸項")
            {

                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");


            }
            else if (DocKind == "AR" || DocKind == "AR-服務" | DocKind == "AR預")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");



            }
            else if (DocKind == "JE")
            {

            }
            else if (DocKind == "貸項-服務")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) 總成本 ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
                connection.Close();
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum, Int32 LineNum)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {

                sb.Append(" SELECT LINENUM,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append(" T0.[LineTotal], T0.[StockPrice],T2.總成本 ");
                sb.Append(" FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append(" INNER JOIN (SELECT SUM([Debit]-[Credit]) 總成本,TransId FROM JDT1 WHERE [Account]='51100101' GROUP BY TransId) T2 ");
                sb.Append(" ON(T1.TransId=T2.TransId)");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");
                sb.Append("and   T0.LineNum =@LineNum   ");



            }
            else if (DocKind == "銷退")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "貸項")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");

                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT * FROM (SELECT T0.ACCTCODE,SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='I'   ");
                sb.Append("UNION ALL   ");


            }
            else if (DocKind == "AR預")
            {



                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0   and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2  ");
            }
            else if (DocKind == "JE")
            {

                sb.Append("SELECT  T0.Account as CardCode, T0.LineMemo as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price], ");
                sb.Append("  T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice] FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID  ");
                sb.Append("WHERE T1.TransID =@DocEntry   ");

            }
            else if (DocKind == "貸項-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));

            command.Parameters.Add(new SqlParameter("@A1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox6.Text));
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
        private System.Data.DataTable GetSAPDoc(string DocKind, Int32 DocNum, string AcctCode, string A1, string A2)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "交貨")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");


            }
            else if (DocKind == "銷退")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "貸項")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='I'   ");



            }
            else if (DocKind == "AR預")
            {

                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 ");
                sb.Append("INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append("INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)");
                sb.Append("LEFT JOIN ODLN T4 ON (T3.DOCENTRY=T4.DOCENTRY)");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   T1.UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0 and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2     ");


            }
            else if (DocKind == "JE")
            {

                sb.Append("                  SELECT  T0.U_REMARK1 as CardCode, T2.CARDNAME as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price],  ");
                sb.Append("                    T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice],'103' GROUPCODE FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID   ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.U_REMARK1 = T2.CARDCODE");
                sb.Append(" WHERE T1.TransID =@DocEntry   AND T0.REF2='XX'  ");

            }
            else if (DocKind == "JE2")
            {

                sb.Append("                  SELECT  T0.U_REMARK1 as CardCode, T2.CARDNAME as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price],  ");
                sb.Append("                   0 as  [LineTotal],T0.[Debit] - T0.[Credit]  as [StockPrice],'103' GROUPCODE FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID   ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.U_REMARK1 = T2.CARDCODE");
                sb.Append(" WHERE T1.TransID+T0.Line_ID =@DocEntry   AND T0.REF2='XX'  ");

            }
            else if (DocKind == "貸項-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //加入基礎單號 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR-服務")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //加入基礎單號 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            if (DocKind != "JE" && DocKind != "JE2")
            {
                //20081009 增加 科目代號
                sb.Append("AND  T0.AcctCode =@AcctCode   ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            //20081009 增加 科目代號
            command.Parameters.Add(new SqlParameter("@AcctCode", AcctCode));
            command.Parameters.Add(new SqlParameter("@A1", A1));
            command.Parameters.Add(new SqlParameter("@A2", A2));
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


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenue(string DocDate1, string DocDate2)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            //AR服務
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )   總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' ");
            sb.Append(" and (((T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%' )  and isnull(t2.u_acme_arap,'') <> 'xx' ) OR (T1.[Account]='22610103' AND (U_LOCATION)='XX' ))");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            sb.Append(" SELECT '貸項' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112)  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            //貸項服務
            sb.Append(" union all");
            sb.Append(" SELECT '貸項-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  )  ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            sb.Append(" SELECT 'AR' as 單別,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append(" MAX(T7.AcctCode)  科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,0 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" 0  總成本,");
            sb.Append(" 0  - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append(" inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("  AND T2.[DocTotal] = 0 ");
            sb.Append(" GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR預' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account 科目代號,");
            sb.Append("              T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T6.[DOCDATE])  日期,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("              INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("              INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            //20150721AR預切41100101
            //           sb.Append("              WHERE T2.[DocType] ='I' and ((T1.[Account] = '22610103') OR (T2.DOCENTRY in (10198,24001,24555,26572))) ");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101') ");
            sb.Append("              and T6.DOCDATE BETWEEN Convert(varchar(8),@DocDate1,112) and  Convert(varchar(8),@DocDate2,112) ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            sb.Append(" union all");
            //20120419 AR貸項沒有收入
            sb.Append("             SELECT '貸項' as 單別,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T6.AcctCode)  科目代號,");
            sb.Append("             T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,0 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append("             SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append("             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("             INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("             INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("             INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append("             WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("             GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            //20150916 AR預開貸項服務
            sb.Append("               SELECT '貸項-服務' as 單別,T2.[DocNum],T0.[TransId], ");
            sb.Append("               T1.Account 科目代號, ");
            sb.Append("               T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組, ");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T6.[DOCDATE])  日期,");
            sb.Append("               SUM(T1.[Debit] - T1.[Credit])  總成本, ");
            sb.Append("               (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利 ");
            sb.Append("               FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               INNER JOIN ORIN T2 ON T0.TransId = T2.TransId  ");
            sb.Append("               INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("               INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE  ");
            sb.Append("               INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("        INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0 ");
            sb.Append("               LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13 ");
            sb.Append("               GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.U_ACME_ARAP=T6.BASEENTRY) ");
            sb.Append("               WHERE T2.[DocType] ='S' AND T1.ACCOUNT='22610103' AND U_LOCATION='XX'");
            sb.Append("              and T6.DOCDATE BETWEEN Convert(varchar(8),@DocDate1,112) and  Convert(varchar(8),@DocDate2,112) ");
            sb.Append("     GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            sb.Append(" union all");
            //20151006  折讓貸項
            sb.Append("                        SELECT 'JE' as 單別,T0.TransId,T0.[TransId],  ");
            sb.Append("                             T1.Account 科目代號,  ");
            sb.Append("                           T1.REF1 業務員編號, T3.SlpName  姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append("                             SUM(T1.debit)*(-1) 總金額,MAX(T0.REFDATE) 日期,  ");
            sb.Append("                             0  總成本,  ");
            sb.Append("                                     SUM(T1.debit)*(-1) 總毛利  ");
            sb.Append("                             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("                             INNER JOIN OSLP T3 ON T1.REF1 = T3.SlpCode  ");
            sb.Append(" INNER JOIN OCRD T2 ON T1.U_REMARK1 = T2.CARDCODE");
            sb.Append("               INNER JOIN OCRG T5 ON T2.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("                             WHERE T1.ACCOUNT='41900101' and isnull(T1.REF2,'')  ='xx'");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("                             GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


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
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("年", typeof(int));
            dt.Columns.Add("月", typeof(int));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("群組", typeof(string));
            dt.Columns.Add("群組次分類", typeof(string));
            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("品名敘述", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            dt.Columns.Add("成本", typeof(string));
            dt.Columns.Add("毛利", typeof(string));
            dt.Columns.Add("毛利率", typeof(string)); 
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("供應商", typeof(string));


            return dt;
        }
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("收入單據", typeof(string));
            dt.Columns.Add("收入單號", typeof(Int32));

            dt.Columns.Add("成本單據", typeof(string));
            dt.Columns.Add("成本單號", typeof(Int32));

            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("客戶群組", typeof(string));

            //20081008
            //業務員   
            dt.Columns.Add("業務員編號", typeof(string));
            dt.Columns.Add("姓名", typeof(string));


            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("數量", typeof(Int32));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("金額", typeof(Int32));

            dt.Columns.Add("項目成本", typeof(Int32));   //有成本時寫入此欄位
            dt.Columns.Add("單號總成本", typeof(Int32)); //有成本時寫入此欄位

            dt.Columns.Add("單號總收入", typeof(Int32));
            dt.Columns.Add("基礎單號", typeof(Int32));
            dt.Columns.Add("基礎列", typeof(Int32));

            dt.Columns.Add("日期", typeof(DateTime));
            dt.Columns.Add("科目代號", typeof(string));




            return dt;
        }
        System.Data.DataTable TF(string TRGETENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT RANK() OVER (ORDER BY DOCENTRY DESC) AS 序號,DOCENTRY AR,TRGETENTRY 交貨 FROM  INV1 WHERE  TRGETENTRY IN (");
            sb.Append(" select docentry  from dln1 where BASEtype='13'");
            sb.Append(" GROUP BY DOCENTRY HAVING COUNT (DISTINCT BASEENTRY) >1) AND    DOCENTRY NOT IN (SELECT DOCENTRY FROM OINV WHERE DOCTOTAL=0) AND TRGETENTRY=@TRGETENTRY  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRGETENTRY", TRGETENTRY));
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
        public void AddATC1(string CARDNAME, string ITEMCODE, string DOCDATE, int DYEAR, int DMONTH, string BU, string DG, string DG2, string MODEL, string VER, string ITEMNAME, int GQty, int GTotal, int GValue, int GM, string GM2, string USD, string SALES, string CARDNAME2)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into AP_JO(CARDNAME,ITEMCODE,DOCDATE,DYEAR,DMONTH,BU,DG,DG2,MODEL,VER,ITEMNAME,GQty,GTotal,GValue,GM,GM2,USD,SALES,CARDNAME2,USERS) values(@CARDNAME,@ITEMCODE,@DOCDATE,@DYEAR,@DMONTH,@BU,@DG,@DG2,@MODEL,@VER,@ITEMNAME,@GQty,@GTotal,@GValue,@GM,@GM2,@USD,@SALES,@CARDNAME2,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@DYEAR", DYEAR));
            command.Parameters.Add(new SqlParameter("@DMONTH", DMONTH));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@DG", DG));
            command.Parameters.Add(new SqlParameter("@DG2", DG2));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@GQty", GQty));
            command.Parameters.Add(new SqlParameter("@GTotal", GTotal));
            command.Parameters.Add(new SqlParameter("@GValue", GValue));
            command.Parameters.Add(new SqlParameter("@GM", GM));
            command.Parameters.Add(new SqlParameter("@GM2", GM2));
            command.Parameters.Add(new SqlParameter("@USD", USD));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@CARDNAME2", CARDNAME2));

            //GQty,GTotal,GValue,GM,GM2,USD,SALES,CARDNAME2,USERS

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
        public void DELATC1()
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("DELETE AP_JO WHERE USERS=@USERS", connection);
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

        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetCARDNAME();

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void dataGridView3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {

                string da = dataGridView3.SelectedRows[0].Cells["客戶名稱"].Value.ToString();

                JOJO4 a = new JOJO4();
                a.PublicString = da;

                a.ShowDialog();
            }
        }
    }
}

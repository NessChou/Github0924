using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using CarlosAg.ExcelXmlWriter;
using System.IO;
namespace ACME
{
    public partial class TRAIL : Form
    {
        int FixedCol = 2;
        string SAPConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
        public TRAIL()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string SAPConnStr2 = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql05";
            string SAPCH16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH21 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH02 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH09 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH03 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH06 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP06;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            string SAPCH23 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP23;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            //進金生
            System.Data.DataTable T1 = VOIDACME(SAPConnStr);

            ArrayList al = new ArrayList();
            string ff = "";
            if (listBox1.SelectedItems.Count > 0)
            {
                for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                {
                    al.Add(listBox1.SelectedItems[i].ToString());

                    string fd = listBox1.SelectedItems[i].ToString();
                    ff = listBox1.SelectedItems[0].ToString();
                }
                StringBuilder sb = new StringBuilder();



                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);

                if (ff != "All")
                {
                    T1.DefaultView.RowFilter = " Dept in (" + sb + ") ";
                }
            }
            進金生.DataSource = T1;
            FixedCol = 6;
            for (int i = FixedCol; i <= 進金生.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 進金生.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //達睿生
            System.Data.DataTable T2 = VOID(SAPConnStr2);
            達睿生.DataSource = T2;
            for (int i = FixedCol; i <= 達睿生.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 達睿生.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //宇豐
            System.Data.DataTable T3 = VOID2(SAPCH16);
            宇豐.DataSource = T3;
            for (int i = FixedCol; i <= 宇豐.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 宇豐.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //創田
            System.Data.DataTable T4 = VOID2(SAPCH21);
            創田.DataSource = T4;
            for (int i = FixedCol; i <= 創田.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 創田.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }


            //鼎園
            System.Data.DataTable T5 = VOID2(SAPCH20);
            鼎園.DataSource = T5;
            for (int i = FixedCol; i <= 鼎園.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 鼎園.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //ARMAS
            System.Data.DataTable T6 = VOID2(SAPCH02);
            聿豐.DataSource = T6;
            for (int i = FixedCol; i <= 聿豐.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 聿豐.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //博豐
            System.Data.DataTable T7 = VOID2(SAPCH09);
            博豐.DataSource = T7;
            for (int i = FixedCol; i <= 博豐.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 博豐.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //Infinite
            System.Data.DataTable T8 = VOID2(SAPCH22);
            Infinite.DataSource = T8;
            for (int i = FixedCol; i <= Infinite.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = Infinite.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //禾中
            System.Data.DataTable T9 = VOID2(SAPCH23);
            禾中.DataSource = T9;
            for (int i = FixedCol; i <= 禾中.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 禾中.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //東門
            System.Data.DataTable T10 = VOID2(SAPCH03);
            東門.DataSource = T10;
            for (int i = FixedCol; i <= 東門.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 東門.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

            //禾豐
            System.Data.DataTable T11 = VOID2(SAPCH06);
            禾豐.DataSource = T11;
            for (int i = FixedCol; i <= 禾豐.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 禾豐.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }
        }
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            //dt.Columns.Add("GruopAc", typeof(string));
            //dt.Columns.Add("GruopName", typeof(string));
            dt.Columns.Add("OpenBalance", typeof(Int64));

            int StartMon = Convert.ToInt32(cbMon1.Text);
            int EndMon = Convert.ToInt32(cbMon2.Text);

            for (int i = StartMon; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }



            dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }

        private System.Data.DataTable MakeTable2ACME()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("Dept", typeof(string));
            dt.Columns.Add("DeptName", typeof(string));
            dt.Columns.Add("GroupAccount", typeof(string));
            dt.Columns.Add("GroupName", typeof(string));
            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("OpenBalance", typeof(Int64));
            
            int StartMon = Convert.ToInt32(cbMon1.Text);
            int EndMon = Convert.ToInt32(cbMon2.Text);

            for (int i = StartMon; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }



            dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }
        private System.Data.DataTable GetOpenBalance(string Year,string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account], SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account]");
            sb.Append(" ORDER BY T0.[Account]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@P1", Year + "0101"));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Account"];
            dt.PrimaryKey = colPk;

            return dt;


        }
        private System.Data.DataTable GetOpenBalanceACME(string Year, string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account], ISNULL(T0.[ProfitCode],'')  ProfitCode, SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account], ISNULL(T0.[ProfitCode],'') ");
            sb.Append(" ORDER BY T0.[Account]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@P1", Year + "0101"));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            dt.PrimaryKey = colPk;

            return dt;


        }
        private System.Data.DataTable GetOpenBalanceCHI(string Year, string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("   Select distinct A.SubjectID Account,IsNull((Select Sum(IsNull(CurAmount, 0)*(2*DebitCredit-1))  From AccSurplus Where SubjectID = A.SubjectID ");
            sb.Append("   ),0) + IsNull((Select Sum(IsNull(DebitAmount,0)-IsNull(CreditAmount,0)) ");
            sb.Append("   From AccYearFile Where  Flag in (1, 3)  and SubjectID= A.SubjectID and HistYearMonth < @P1 ),0)");
            sb.Append("    as Balance  From ComSubject A WHERE  A.IsUseSubject = 1");
            sb.Append("          Group By A.SubjectID");

     

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@P1", Year));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Account"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetAccount(string CONN)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[AcctCode], T0.[AcctName] ,T0.[Postable],T0.[fathernum],T3.[AcctName] GROUPNAME");
            sb.Append(" FROM  [OACT] T0");
            sb.Append(" LEFT join  [OACT] T3  ON  T3.[AcctCode] = T0.fathernum ");
            sb.Append(" ORDER BY T0.[AcctCode]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0]; ;


        }

        private System.Data.DataTable GetAccountACME(string RefDate1, string RefDate2, string CONN)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT distinct T2.[AcctCode] , T2.[AcctName] ,T2.[Postable],ISNULL(ProfitCode,'') ProfitCode,PRCNAME ProfitName ");
            sb.Append(" ,T2.fathernum FATHER,T4.[AcctName] ACCGROUP");
            sb.Append("                                  FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId    ");
            sb.Append("        LEFT JOIN [OACT] T2 ON (T0.[Account]=T2.[AcctCode])   LEFT JOIN [OPRC] T3 ON (T0.ProfitCode=T3.PRCCODE )  ");
            sb.Append(" LEFT join  [OACT] T4  ON  T2.fathernum = T4.[AcctCode]     ");
            sb.Append(" WHERE Convert(varchar(8),T0.[RefDate],112)   BETWEEN  @RefDate1  AND   @RefDate2  ");
            sb.Append("                                  AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'  ORDER BY  T2.[AcctCode],ISNULL(ProfitCode,'')  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0]; ;


        }
        private System.Data.DataTable GetAccountCHI(string CONN)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" Select  A.SubjectID AcctCode,A.SubjectName AcctName From ComSubject A   WHERE  A.IsUseSubject = 1  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0]; ;


        }
        private System.Data.DataTable GetBalance(string RefDate1, string RefDate2, string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account], SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account]");
            sb.Append(" ORDER BY T0.[Account]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Account"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetBalanceACME(string RefDate1, string RefDate2, string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'')  ProfitCode, SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append("         GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],'') ");
            sb.Append(" ORDER BY T0.[Account]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            dt.PrimaryKey = colPk;

            return dt;


        }
        private System.Data.DataTable GetBalanceCHI(string RefDate1, string RefDate2, string CONN)
        {

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append(" Select distinct A.SubjectID Account,Sum(IsNull(D.DebitAmount, 0))- Sum(IsNull(D.CreditAmount, 0)) as Balance");
            sb.Append("  From ComSubject A ");
            sb.Append("  left Join AccYearFile D on (D.SubjectID = A.SubjectID and  D.Flag in (1, 3) )  And D.HistYearMonth Between @RefDate1 And @RefDate2 And D.DeptProjID Between 'A' And 'S'");
            sb.Append("   WHERE  A.IsUseSubject = 1");
            sb.Append("   Group By A.SubjectID, A.SubjectName, A.SubjectEngName, A.ParentSubID, A.SubjectType, A.SubLevel, A.IsUseSubject");
            sb.Append("   ORDER BY  A.SubjectID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }

            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Account"];
            dt.PrimaryKey = colPk;

            return dt;


        }
        private void TRAIL_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 12; i++)
            {
                cbMon1.Items.Add(i.ToString("00"));
                cbMon2.Items.Add(i.ToString("00"));
            }

            int currentMon = DateTime.Now.Month - 1;
            cbMon1.SelectedIndex = 0;
            cbMon2.SelectedIndex = currentMon;
            UtilSimple.SetLookupBinding(cbYear, GetMenu.Year(), "DataValue", "DataValue");

            System.Data.DataTable  dtDept2 = GetDept2();
            DataRow dr;
            for (int i = 0; i <= dtDept2.Rows.Count - 1; i++)
            {
                dr = dtDept2.Rows[i];

                string PrcCode = Convert.ToString(dr["PrcCode"]);

                listBox1.Items.Add(PrcCode);
            }
        }
        private System.Data.DataTable GetDept2()
        {

            SqlConnection connection = new SqlConnection(SAPConnStr);

            StringBuilder sb = new StringBuilder();
            sb.Append(" select param_desc COLLATE Chinese_Taiwan_Stroke_CI_AS PrcCode from acmesqlsp.dbo.rma_params where id=3");
            sb.Append("   union all SELECT T0.[PrcCode] FROM  [OPRC] T0  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRC");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables["OPRC"];



        }
        private System.Data.DataTable VOID(string CONN)
        {
            long QTY = 0;

            DataTable dt = null;
            dt = MakeTable2();

            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);
 

            DataTable dtAccount = GetAccount(CONN);

            DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00") + "01";
                string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

                ArrayDt[j] = GetBalance(Date1, Date2, CONN);
            }

            ArrayDt[0] = GetOpenBalance(sYear, CONN);


            string AcctCode = "";
            string AcctName = "";
            string Postable = "";

            DataRow dr;
            DataRow row;

            for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
            {
                dr = dtAccount.Rows[i];

                AcctCode = Convert.ToString(dr["AcctCode"]);
                AcctName = Convert.ToString(dr["AcctName"]);
                Postable = Convert.ToString(dr["Postable"]);

                if (Postable == "N")
                {
                    continue;
                }

                row = dt.NewRow();
                row["AccountCode"] = AcctCode;
                row["AccountName"] = AcctName;

                DataRow drFind;
                Int64 Total = 0;

                drFind = ArrayDt[0].Rows.Find(AcctCode);
                if (drFind != null)
                {

                    row["OpenBalance"] = drFind["Balance"];
                    Total += Convert.ToInt64(drFind["Balance"]);

                }
                else
                {
                    row["OpenBalance"] = 0;
                }

                for (int j = iMon1; j <= iMon2; j++)
                {


                    drFind = ArrayDt[j].Rows.Find(AcctCode);
                    if (drFind != null)
                    {
                        if (String.IsNullOrEmpty(drFind["Balance"].ToString()))
                        {
                            drFind["Balance"] = 0;
                        }
                        row[j.ToString("00")] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }
                    else
                    {
                        row[j.ToString("00")] = 0;
                    }

                }
                row["Total"] = Total;

                dt.Rows.Add(row);

                if (!String.IsNullOrEmpty(row["01"].ToString()))
                {
                    QTY += Convert.ToInt64(row["01"]);
                }
            }
            return dt;

        }

        private System.Data.DataTable VOIDACME(string CONN)
        {
            long QTY = 0;

            DataTable dt = null;
            dt = MakeTable2ACME();

            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);
            string Date11 = sYear + iMon1.ToString("00") + "01";
            string Date21 = sYear + iMon2.ToString("00") + "31";
            if (cbYear.Text == "2014" && iMon1 == 1)
            {
                Date11 = "20140102";
            }
            DataTable dtAccount = GetAccountACME(Date11, Date21, CONN);

            DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00") + "01";
                string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

                if (cbYear.Text == "2014" && j == 1)
                {
                    Date1 = "20140102";
                }

                ArrayDt[j] = GetBalanceACME(Date1, Date2, CONN);
            }

            ArrayDt[0] = GetOpenBalanceACME(sYear, CONN);


            string AcctCode = "";
            string AcctName = "";
            string Postable = "";
            string ProfitCode = "";
            string ProfitName = "";
            string FATHER = "";
            string ACCGROUP = "";
            DataRow dr;
            DataRow row;

            for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
            {
                dr = dtAccount.Rows[i];

                AcctCode = Convert.ToString(dr["AcctCode"]);
                AcctName = Convert.ToString(dr["AcctName"]);
                Postable = Convert.ToString(dr["Postable"]);
                ProfitCode = Convert.ToString(dr["ProfitCode"]);
                ProfitName = Convert.ToString(dr["ProfitName"]);
                FATHER = Convert.ToString(dr["FATHER"]);
                ACCGROUP = Convert.ToString(dr["ACCGROUP"]);
                if (Postable == "N")
                {
                    continue;
                }

                row = dt.NewRow();
                row["AccountCode"] = AcctCode;
                row["AccountName"] = AcctName;
                row["Dept"] = ProfitCode;
                row["DeptName"] = ProfitName;
                row["GroupAccount"] = FATHER;
                row["GroupName"] = ACCGROUP;
                DataRow drFind;
                Int64 Total = 0;
                Object[] Key = new object[] { AcctCode, ProfitCode };
                drFind = ArrayDt[0].Rows.Find(Key);
                if (drFind != null)
                {

                    row["OpenBalance"] = drFind["Balance"];
                    Total += Convert.ToInt64(drFind["Balance"]);

                }
                else
                {
                    row["OpenBalance"] = 0;
                }

                for (int j = iMon1; j <= iMon2; j++)
                {


                    drFind = ArrayDt[j].Rows.Find(Key);
                    if (drFind != null)
                    {
                        if (String.IsNullOrEmpty(drFind["Balance"].ToString()))
                        {
                            drFind["Balance"] = 0;
                        }
                        row[j.ToString("00")] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }
                    else
                    {
                        row[j.ToString("00")] = 0;
                    }

                }
                row["Total"] = Total;

                dt.Rows.Add(row);

                if (!String.IsNullOrEmpty(row["01"].ToString()))
                {
                    QTY += Convert.ToInt64(row["01"]);
                }
            }
            return dt;

        }
        private System.Data.DataTable VOID2(string CONN)
        {
            long QTY = 0;

            DataTable dt = null;
            dt = MakeTable2();

            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);
           
            DataTable dtAccount = GetAccountCHI(CONN);

            DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00");
                string Date2 = sYear + j.ToString("00");

                ArrayDt[j] = GetBalanceCHI(Date1, Date2, CONN);
            }
            ArrayDt[0] = GetOpenBalanceCHI(cbYear.Text + cbMon1.Text, CONN);


            string AcctCode = "";
            string AcctName = "";


            DataRow dr;
            DataRow row;

            for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
            {
                dr = dtAccount.Rows[i];

                AcctCode = Convert.ToString(dr["AcctCode"]);
                AcctName = Convert.ToString(dr["AcctName"]);
            

                row = dt.NewRow();
                row["AccountCode"] = AcctCode;
                row["AccountName"] = AcctName;

                DataRow drFind;
                Int64 Total = 0;

                drFind = ArrayDt[0].Rows.Find(AcctCode);
                if (drFind != null)
                {
                    row["OpenBalance"] = drFind["Balance"];
                    Total += Convert.ToInt64(drFind["Balance"]);

                }

                for (int j = iMon1; j <= iMon2; j++)
                {


                    drFind = ArrayDt[j].Rows.Find(AcctCode);
                    if (drFind != null)
                    {
                        row[j.ToString("00")] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }

                }
                row["Total"] = Total;

                dt.Rows.Add(row);

                if (!String.IsNullOrEmpty(row["01"].ToString()))
                {
                    QTY += Convert.ToInt64(row["01"]);
                }
            }
            return dt;





        }

   

        private void button3_Click(object sender, EventArgs e)
        {
            if (進金生.Rows.Count == 0)
            {
                MessageBox.Show("請先點選查詢");
                return;

            }

            CarlosAg.ExcelXmlWriter.Workbook book = new CarlosAg.ExcelXmlWriter.Workbook();
            WorksheetStyle headerStyle = book.Styles.Add("headerStyleID");
            headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            headerStyle.Alignment.WrapText = true;
            headerStyle.Interior.Color = "#284775";
            headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;
            headerStyle.Font.Color = "white";
            headerStyle.Font.Bold = true;

            WorksheetStyle defaultStyle = book.Styles.Add("workbookStyleID");
            defaultStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            defaultStyle.Alignment.WrapText = true;
            defaultStyle.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");

            WorksheetStyle defaultStyle2 = book.Styles.Add("workbookStyleID2");
            defaultStyle2.Alignment.Horizontal = StyleHorizontalAlignment.Right;
            defaultStyle2.Alignment.WrapText = true;
            defaultStyle2.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle2.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle2.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle2.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");
            foreach (Control x in this.Controls)
            {
                if (x is TabControl)
                {
                    if (x.HasChildren)
                    {
                        foreach (Control CHILD in x.Controls)
                        {
                            foreach (Control RCHILD in CHILD.Controls)
                            {
                                DataGridView aTextBox = (DataGridView)RCHILD;
                                WH(book, aTextBox, aTextBox.Name.ToString());

                            }
                        }
                    }
                }
            }
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
    DateTime.Now.ToString("yyyyMMddHHmmss") + "試算表.xls";
            book.Save(OutPutFile);
            System.Diagnostics.Process.Start(OutPutFile);
     

        }
        private void WH(CarlosAg.ExcelXmlWriter.Workbook book, DataGridView DGV,string DD)
        {



            Worksheet sheet = book.Worksheets.Add(DD);
            WorksheetRow headerRow = sheet.Table.Rows.Add();
            for (int i = 0; i < DGV.Columns.Count; i++)
            {
                headerRow.Cells.Add(DGV.Columns[i].HeaderText, DataType.String, "headerStyleID");
            }

            for (int i = 0; i < DGV.Rows.Count - 1; i++)
            {

                DataGridViewRow row = DGV.Rows[i];
                WorksheetRow rowS = sheet.Table.Rows.Add();

                for (int j = 0; j < row.Cells.Count ; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                 

                    if (DD == "進金生")
                    {
                        if (j == 0 || j == 1 || j == 2 || j == 3 || j == 4 || j == 5)
                        {
                            rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                        }
                        else
                        {
                            rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                        }
                    }
                    else
                    {
                        if (j == 0 || j == 1)
                        {
                            rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                        }
                        else
                        {
                            rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                        }
                    }
                    rowS.AutoFitHeight = true;
                    rowS.Table.DefaultColumnWidth = 100;
                   
                }

            }
        }

    }
}
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
    public partial class fmAcmeTrialReport : Form
    {


        public  System.Data .DataTable dt = null;
        public System.Data.DataTable dt3 = null;
        public System.Data.DataTable dt4 = null;
        public System.Data.DataTable dt5 = null;
        public System.Data.DataTable dt6 = null;
        public System.Data.DataTable dt7 = null;
        public System.Data.DataTable dtS = null;
        private DataTable dtOcrd;
        DataRow dr22 = null;
        DataRow dr23 = null;
        private DataTable dtProject;

        private DataTable dtDept, dtDept2, dtDept3;

        int FixedCol = 6;

        public fmAcmeTrialReport()
        {
            InitializeComponent();
        }

        private void fmAcmeTrialReport_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 12; i++)
            {
        
                cbMon2.Items.Add(i.ToString("00"));
            }
            cbMon1.Items.Add(1.ToString("00"));
            int currentMon = DateTime.Now.Month - 1;
            if (currentMon == 0)
            {
                currentMon = 1;
            }
            cbMon1.SelectedIndex = 0;
            cbMon2.SelectedIndex = cbMon2.Items.IndexOf(currentMon.ToString("00"));

            UtilSimple.SetLookupBinding(cbYear, GetMenu.Year(), "DataValue", "DataValue");


            dtOcrd = GetOCRD();
            dtProject = GetProject();
            dtDept = GetDept();
            dtDept2 = GetDept2();
            dtDept3 = GetDept3();

            if (globals.GroupID.ToString().Trim() == "ACCS")
            {
                listBox2.Visible = false;
                listBox1.Visible = false;
                label1.Visible = false;
                checkBox1.Visible = false;
            }
            DataRow dr;
            for (int i = 0; i <= dtDept2.Rows.Count - 1; i++)
            {
                dr = dtDept2.Rows[i];

                string PrcCode = Convert.ToString(dr["PrcCode"]);

                listBox1.Items.Add(PrcCode);
            }


            DataRow dr2;
            for (int i = 0; i <= dtDept3.Rows.Count - 1; i++)
            {
                dr2 = dtDept3.Rows[i];

                string BU = Convert.ToString(dr2["BU"]);

                listBox2.Items.Add(BU);
            }
     
        }

        private void button1_Click(object sender, EventArgs e)
        {
          

            long QTY = 0;

            DataTable dt = null;
            DataTable dtS = null;
            dt = MakeTable2();

            int iYear =Convert.ToInt32(cbYear.Text);
            string sYear =cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);

            DataTable dtAccount = GetAccount();
        
            DataTable[] ArrayDt = new DataTable[iMon2-iMon1+2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00") + "01";
                string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

                ArrayDt[j] = GetBalance(Date1, Date2);
            }
            ArrayDt[0] = GetOpenBalance(sYear);


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
               row["GruopAc"] = Convert.ToString(dr["fathernum"]);
               row["GruopName"] = Convert.ToString(dr["GROUPNAME"]);
                DataRow drFind;
                Int64 Total =0;

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
               for (int j = iMon1; j <= iMon2; j++)
               {
                   if (String.IsNullOrEmpty(row[j.ToString("00")].ToString()))
                   {
                       row[j.ToString("00")] = "0";
                   }
               }
               dt.Rows.Add(row);

               if (!String.IsNullOrEmpty(row["01"].ToString()))
               {
                   QTY += Convert.ToInt64(row["01"]);
               }
            }
       
            月.DataSource = dt;

            for (int i = 4; i <= 月.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 月.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";
            
            }


            dtS = MakeTable2S();
            for (int l = 0; l <= dt.Rows.Count - 1; l++)
            {
                DataRow dz = dt.Rows[l];
                dr22 = dtS.NewRow();
                string AccountCode = dz["AccountCode"].ToString();
        
                dr22["AccountCode"] = AccountCode;
                dr22["AccountName"] = dz["AccountName"].ToString();
                dr22["GruopAc"] = dz["GruopAc"].ToString();
                dr22["GruopName"] = dz["GruopName"].ToString();

                dr22["OpenBalance"] = dz["OpenBalance"];

                string Q1 = "";
                string Q2 = "";
                string Q3 = "";
                string Q4 = "";
                int EndMon = Convert.ToInt32(cbMon2.Text);
                if (EndMon == 1)
                {
                    Q1 = dt.Compute("SUM([01])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 2)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])", "AccountCode='" + AccountCode + "'    ").ToString();
                }
                else if (EndMon == 3)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 4)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 5)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 6)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 7)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 8)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])+SUM([08])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 9)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 10)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q4 = dt.Compute("SUM([10])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 11)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q4 = dt.Compute("SUM([10])+SUM([11])", "AccountCode='" + AccountCode + "'   ").ToString();
                }
                else if (EndMon == 12)
                {
                    Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "'   ").ToString();
                    Q4 = dt.Compute("SUM([10])+SUM([11])+SUM([12])", "AccountCode='" + AccountCode + "'   ").ToString();
                }

                dr22["Q1"] = Q1;
                dr22["Q2"] = Q2;
                dr22["Q3"] = Q3;
                dr22["Q4"] = Q4;
                dr22["Total"] = dz["Total"];
                dtS.Rows.Add(dr22);
            }
            季.DataSource = dtS;
            for (int i = 4; i <= 季.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = 季.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }

        }


        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("GruopName", typeof(string));
            dt.Columns.Add("DeptCode", typeof(string));
            dt.Columns.Add("Dept", typeof(string));
            dt.Columns.Add("參考3", typeof(string));
            dt.Columns.Add("ProjectCode", typeof(string));
            dt.Columns.Add("Project", typeof(string));
            dt.Columns.Add("CustomerCode", typeof(string));
            dt.Columns.Add("Customer", typeof(string));

            dt.Columns.Add("OpenBalance", typeof(Int64));

            int StartMon = Convert.ToInt32(cbMon1.Text);
            int EndMon   = Convert.ToInt32(cbMon2.Text);

            for (int i = StartMon; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }



            dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }
        private System.Data.DataTable MakeTableS()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("GruopName", typeof(string));
            dt.Columns.Add("DeptCode", typeof(string));
            dt.Columns.Add("Dept", typeof(string));
            dt.Columns.Add("參考3", typeof(string));
            dt.Columns.Add("ProjectCode", typeof(string));
            dt.Columns.Add("Project", typeof(string));
            dt.Columns.Add("CustomerCode", typeof(string));
            dt.Columns.Add("Customer", typeof(string));
            dt.Columns.Add("OpenBalance", typeof(Int64));
            dt.Columns.Add("Q1", typeof(string));
            dt.Columns.Add("Q2", typeof(string));
            dt.Columns.Add("Q3", typeof(string));
            dt.Columns.Add("Q4", typeof(string));
            dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }
        private System.Data.DataTable MakeTableJ1()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("GruopName", typeof(string));
            dt.Columns.Add("DeptCode", typeof(string));
            dt.Columns.Add("Dept", typeof(string));
            dt.Columns.Add("參考3", typeof(string));
            dt.Columns.Add("ProjectCode", typeof(string));
            dt.Columns.Add("Project", typeof(string));
            dt.Columns.Add("CustomerCode", typeof(string));
            dt.Columns.Add("Customer", typeof(string));
            dt.Columns.Add("OpenBalance", typeof(Int64));
            dt.Columns.Add("Total", typeof(Int64));

            return dt;
        }

        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("GruopName", typeof(string));
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

        private System.Data.DataTable MakeTable2S()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("GruopName", typeof(string));
            dt.Columns.Add("OpenBalance", typeof(Int64));
            dt.Columns.Add("Q1", typeof(string));
            dt.Columns.Add("Q2", typeof(string));
            dt.Columns.Add("Q3", typeof(string));
            dt.Columns.Add("Q4", typeof(string));
            dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }

        private System.Data.DataTable GetOpenBalance(string Year)
        {

            SqlConnection connection = globals.shipConnection;

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

            command.Parameters.Add(new SqlParameter("@P1", Year+"0101"));

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

        



        private System.Data.DataTable GetBalance(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account], SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account]");
            sb.Append(" ORDER BY T0.[Account]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@RefDate1", SqlDbType.VarChar));
            command.Parameters["@RefDate1"].Value = RefDate1;
            command.Parameters.Add(new SqlParameter("@RefDate2", SqlDbType.VarChar));
            command.Parameters["@RefDate2"].Value = RefDate2;

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


        private System.Data.DataTable GetBalance_Sub(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project, SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" ,ISNULL(T0.REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),ISNULL(T0.REF3LINE,'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),ISNULL(T0.REF3LINE,'')");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RefDate1", SqlDbType.VarChar));
            command.Parameters["@RefDate1"].Value = RefDate1;
            command.Parameters.Add(new SqlParameter("@RefDate2", SqlDbType.VarChar));
            command.Parameters["@RefDate2"].Value = RefDate2;

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
            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }


        /// <summary>
        /// 取出 T0.[Account],T1.[ProfitCode],T1.[Project] 列表
        /// </summary>
        /// <param name="RefDate1"></param>
        /// <param name="RefDate2"></param>
        /// <returns></returns>
        private System.Data.DataTable GetAccount_Sub(string RefDate1, string RefDate2)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T0.[Account] as AcctCode,T2.[AcctName],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project ");
            sb.Append(" ,ISNULL(REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" Inner join  [OACT] T2  ON  T2.[AcctCode] = T0.Account   ");

            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),ISNULL(T0.REF3LINE,'')");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@RefDate1", SqlDbType.VarChar));
            command.Parameters["@RefDate1"].Value = RefDate1;

            command.Parameters.Add(new SqlParameter("@RefDate2", SqlDbType.VarChar));
            command.Parameters["@RefDate2"].Value = RefDate2;

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
            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        /// <summary>
        /// 依 T0.[Account],T1.[ProfitCode],T1.[Project] 取得餘額
        /// </summary>
        /// <param name="Year"></param>
        /// <returns></returns>
        private System.Data.DataTable GetOpenBalance_Sub(string Year)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project, SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" ,ISNULL(T0.REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),ISNULL(T0.REF3LINE,'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),ISNULL(T0.REF3LINE,'')");


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
            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }




        private System.Data.DataTable GetAccount()
        {
            SqlConnection connection = globals.shipConnection;

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
            return ds.Tables[0];;


        }


        private System.Data.DataTable GetOCRD()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[CardCode], T0.[CardName] ");
            sb.Append(" FROM  [OCRD] T0");           


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            //command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }
            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["CardCode"];
            dt.PrimaryKey = colPk;

            return dt;



        }

        private System.Data.DataTable GetProject()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[PrjCode], T0.[PrjName] ");
            sb.Append(" FROM  [OPRJ] T0");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

          

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["PrjCode"];
            dt.PrimaryKey = colPk;

            return dt;



        }


        private System.Data.DataTable GetDept()
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[PrcCode], T0.[PrcName] ");
            sb.Append(" FROM  [OPRC] T0");

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
            DataTable dt = ds.Tables[0];
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["PrcCode"];
            dt.PrimaryKey = colPk;

            return dt;

        }

        private System.Data.DataTable GetDept2()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select param_desc COLLATE Chinese_Taiwan_Stroke_CI_AS PrcCode from acmesqlsp.dbo.rma_params where id=3");
            sb.Append("   union all SELECT T0.[PrcCode] FROM  [OPRC] T0  WHERE SUBSTRING(T0.[PrcCode],1,1) LIKE '%[0-9]%' ");



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


        private System.Data.DataTable GetDept3()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  BU  FROM Account_Bu2 order by seq");
          


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

        private System.Data.DataTable GetBU(string BU)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DEPT FROM ACCOUNT_BU WHERE BU=@BU");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BU", BU));
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
        private System.Data.DataTable GetAccount_Sub_Ocrd(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T0.[Account] as AcctCode,T2.[AcctName],T3.[AcctName] ACCGROUP,T2.fathernum,ISNULL(T0.[ProfitCode],'') as ProfitCode,T0.ShortName,ISNULL(T0.[Project],'') as Project");
            sb.Append(" ,ISNULL(T0.REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" Inner join  [OACT] T2  ON  T2.[AcctCode] = T0.Account   ");
            sb.Append(" LEFT join  [OACT] T3  ON  T3.[AcctCode] = T2.fathernum   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName,ISNULL(T0.REF3LINE,'')");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
            command.CommandType = CommandType.Text;



            command.Parameters.Add(new SqlParameter("@RefDate1", SqlDbType.VarChar));
            command.Parameters["@RefDate1"].Value = RefDate1;
            command.Parameters.Add(new SqlParameter("@RefDate2", SqlDbType.VarChar));
            command.Parameters["@RefDate2"].Value = RefDate2;

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
            DataColumn[] colPk = new DataColumn[5];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["ShortName"];
            colPk[4] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetOpenBalance_Sub_Ocrd(string Year)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project,T0.ShortName, SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" ,ISNULL(T0.REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName,ISNULL(T0.REF3LINE,'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName,ISNULL(T0.REF3LINE,'')");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandTimeout = 0;
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
            DataColumn[] colPk = new DataColumn[5];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["ShortName"];
            colPk[4] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetBalance_Sub_Ocrd(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project,T0.ShortName, SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" ,ISNULL(T0.REF3LINE,'') REF FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName,ISNULL(T0.REF3LINE,'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName ,ISNULL(T0.REF3LINE,'')");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            
            command.Parameters.Add(new SqlParameter("@RefDate1", SqlDbType.VarChar));
            command.Parameters["@RefDate1"].Value = RefDate1;
            command.Parameters.Add(new SqlParameter("@RefDate2", SqlDbType.VarChar));
            command.Parameters["@RefDate2"].Value = RefDate2;

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
            DataColumn[] colPk = new DataColumn[5];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["ShortName"];
            colPk[4] = dt.Columns["REF"];
            dt.PrimaryKey = colPk;

            return dt;


        }



        /// <summary>
        /// 測試
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTest_Click(object sender, EventArgs e)
        {
            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);
            string RefDate1=cbYear.Text+cbMon1.Text+"01";
            string RefDate2=cbYear.Text+cbMon2.Text+DateTime.DaysInMonth(iYear,iMon2);
            DataTable dtAccount = GetAccount_Sub_Ocrd(RefDate1, RefDate2);

            月.DataSource = dtAccount;
        }

        private void btnAcme_Click(object sender, EventArgs e)
        {
      
            long  QTY = 0;

            月.DataSource = null;
            季.DataSource = null;
            TFTALL.DataSource = null;
            TFTBU.DataSource = null;
            TFTPurchasing.DataSource = null;
            TFTLogistic.DataSource = null;
            TFTEngService.DataSource = null;

            dtS = null;
            if (checkBox1.Checked)
            {
                button1_Click(sender,e);
            }

            else
            {

                dt = MakeTable();

    

                int iYear = Convert.ToInt32(cbYear.Text);
                string sYear = cbYear.Text;
                int iMon1 = Convert.ToInt32(cbMon1.Text);
                int iMon2 = Convert.ToInt32(cbMon2.Text);



                string RefDate1 = cbYear.Text + cbMon1.Text + "01";
                if (RefDate1 == "20140101")
                {
                    RefDate1 = "20140102";
                }
                string RefDate2 = cbYear.Text + cbMon2.Text + DateTime.DaysInMonth(iYear, iMon2);
                DataTable dtAccount = GetAccount_Sub_Ocrd(RefDate1, RefDate2);

                int K1 = 0;
                DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
                for (int j = iMon1; j <= iMon2; j++)
                {
                    string Date1 = sYear + j.ToString("00") + "01";
                    string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);
                    K1++;
                    if (cbYear.Text == "2014" && j==1)
                    {
                        Date1 = "20140102";
                    }

                    ArrayDt[j] = GetBalance_Sub_Ocrd(Date1, Date2);
                }
                ArrayDt[0] = GetOpenBalance_Sub_Ocrd(sYear);


                string AcctCode = "";
                string AcctName = "";


                string ProfitCode;
                string Project;

                string ShortName;
                string 參考3 = "";
                DataRow dr;
                DataRow row;


                DataRow drFind;
                for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
                {

          
                    dr = dtAccount.Rows[i];

                    AcctCode = Convert.ToString(dr["AcctCode"]);
                    AcctName = Convert.ToString(dr["AcctName"]);


                    ProfitCode = Convert.ToString(dr["ProfitCode"]);
                    Project = Convert.ToString(dr["Project"]);
                    參考3 = Convert.ToString(dr["REF"]);          
                    ShortName = Convert.ToString(dr["ShortName"]);

                    Object[] Key = new object[] { AcctCode, ProfitCode, Project, ShortName, 參考3 };

                  

                    row = dt.NewRow();
                    row["AccountCode"] = AcctCode;
                    row["AccountName"] = AcctName;

                    row["DeptCode"] = ProfitCode;

                    row["ProjectCode"] = Project;

                    row["參考3"] = 參考3;
                    row["GruopAc"] = Convert.ToString(dr["fathernum"]);
                    row["GruopName"] = Convert.ToString(dr["ACCGROUP"]);

                 
                    if (!string.IsNullOrEmpty(ProfitCode))
                    {
                        drFind = dtDept.Rows.Find(ProfitCode);
                        if (drFind != null)
                        {

                            row["Dept"] = drFind["PrcName"];
                        }
                    }


                    if (!string.IsNullOrEmpty(Project))
                    {
                        drFind = dtProject.Rows.Find(Project);
                        if (drFind != null)
                        {
                            row["Project"] = drFind["PrjName"];
                        }
        
                    }


                    if (ShortName == AcctCode)
                    {

                    }
                    else
                    {
                        drFind = dtOcrd.Rows.Find(ShortName);
                        if (drFind != null)
                        {
                            row["CustomerCode"] = ShortName;
                            row["Customer"] = drFind["CardName"];
                        }

                    }




                    Int64 Total = 0;


             

                        drFind = ArrayDt[0].Rows.Find(Key);
                        if (drFind != null)
                        {
                            row["OpenBalance"] = drFind["Balance"];
                            if (!String.IsNullOrEmpty(drFind["Balance"].ToString()))
                            {
                                Total += Convert.ToInt64(drFind["Balance"]);
                            }

                        }
                 

                    for (int j = iMon1; j <= iMon2; j++)
                    {


                        drFind = ArrayDt[j].Rows.Find(Key);
                        if (drFind != null)
                        {
                            string GG = drFind["Balance"].ToString();
                            row[j.ToString("00")] = drFind["Balance"];
                            if (!String.IsNullOrEmpty(GG))
                            {
                                Total += Convert.ToInt64(drFind["Balance"]);
                            }

                        }


                    }



                    row["Total"] = Total;
                    for (int j = iMon1; j <= iMon2; j++)
                    {
                        if (String.IsNullOrEmpty(row[j.ToString("00")].ToString()))
                        {
                            row[j.ToString("00")] = "0";
                        }
                    }
                 
                    dt.Rows.Add(row);

                    if (!String.IsNullOrEmpty(row["01"].ToString()))
                    {
                        QTY += Convert.ToInt64(row["01"]);
                    }


                }

         
                ArrayList al = new ArrayList();
                ArrayList al2 = new ArrayList();
                string ff = "";
                StringBuilder sb = new StringBuilder();
                if (listBox1.SelectedItems.Count > 0)
                {
                    for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                    {
                        al.Add(listBox1.SelectedItems[i].ToString());

                        string fd = listBox1.SelectedItems[i].ToString();
                        ff = listBox1.SelectedItems[0].ToString();
                    }

                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);


                }
                StringBuilder sb2 = new StringBuilder();
                StringBuilder sb3 = new StringBuilder();
                StringBuilder sb4 = new StringBuilder();


                if (listBox2.SelectedItems.Count > 0)
                {
       
                    for (int i = 0; i <= listBox2.SelectedItems.Count - 1; i++)
                    {
                        al2.Add(listBox2.SelectedItems[i].ToString());

                    }

                    foreach (string v in al2)
                    {
                        sb2.Append("'" + v + "',");
                    }
                    sb2.Remove(sb2.Length - 1, 1);
                }
                string J1 = sb2.ToString();
                if (J1.IndexOf("TFT-ALL") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT-ALL");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                else
                {

                    if (J1.IndexOf("TFT-BU") != -1)
                    {
                        System.Data.DataTable k1 = GetBU("TFT-BU");
                        for (int i = 0; i <= k1.Rows.Count - 1; i++)
                        {

                            DataRow dd = k1.Rows[i];

                            string F = dd["DEPT"].ToString();
                            sb3.Append("'" + F + "',");

                        }

                    }
                    if (J1.IndexOf("TFT-Purchasing") != -1)
                    {
                        System.Data.DataTable k1 = GetBU("TFT-Purchasing");
                        for (int i = 0; i <= k1.Rows.Count - 1; i++)
                        {

                            DataRow dd = k1.Rows[i];

                            string F = dd["DEPT"].ToString();
                            sb3.Append("'" + F + "',");

                        }

                    }
                    if (J1.IndexOf("TFT-Logistic") != -1)
                    {
                        System.Data.DataTable k1 = GetBU("TFT-Logistic");
                        for (int i = 0; i <= k1.Rows.Count - 1; i++)
                        {

                            DataRow dd = k1.Rows[i];

                            string F = dd["DEPT"].ToString();
                            sb3.Append("'" + F + "',");

                        }

                    }
                    if (J1.IndexOf("TFT-Eng.Service") != -1)
                    {
                        System.Data.DataTable k1 = GetBU("TFT-Eng.Service");
                        for (int i = 0; i <= k1.Rows.Count - 1; i++)
                        {

                            DataRow dd = k1.Rows[i];

                            string F = dd["DEPT"].ToString();
                            sb3.Append("'" + F + "',");

                        }

                    }
                
                
                }
                if (J1.IndexOf("ESCO") != -1)
                {
                    System.Data.DataTable k1 = GetBU("ESCO");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                 if (J1.IndexOf("SOLAR") != -1)
                 {
                     System.Data.DataTable k1 = GetBU("SOLAR");
                     for (int i = 0; i <= k1.Rows.Count - 1; i++)
                     {

                         DataRow dd = k1.Rows[i];

                         string F = dd["DEPT"].ToString();
                         sb3.Append("'" + F + "',");
        
                     }

                 }
                 if (J1.IndexOf("NON-BU") != -1)
                {
                    System.Data.DataTable k1 = GetBU("NON-BU");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }

                 if (globals.GroupID.ToString().Trim() == "ACCS" )
                 {
                     System.Data.DataTable k1 = GetBU("SOLAR");
                     for (int i = 0; i <= k1.Rows.Count - 1; i++)
                     {

                         DataRow dd = k1.Rows[i];

                         string F = dd["DEPT"].ToString();
                         sb4.Append("'" + F + "',");
                     }
                     sb4.Remove(sb4.Length - 1, 1);
                     dt.DefaultView.RowFilter = " DeptCode in (" + sb4 + ") ";
                 }
                 else
                 {
                     if (listBox2.SelectedItems.Count > 0)
                     {

                         sb3.Remove(sb3.Length - 1, 1);
                         dt.DefaultView.RowFilter = " DeptCode in (" + sb3 + ") ";
                     }
                     else
                     {
                         if (ff != "All" && ff != "")
                         {
                             dt.DefaultView.RowFilter = " DeptCode in (" + sb + ") ";
                         }
                     }
                 }


                 月.DataSource = dt.DefaultView;
  


     

                for (int i = FixedCol; i <= 月.Columns.Count - 1; i++)
                {
                    DataGridViewColumn c = 月.Columns[i];
                    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    c.DefaultCellStyle.Format = "#,##0";

                }

                if (dt.Rows.Count > 0)
                {
                    dtS = MakeTableS();
                    for (int l = 0; l <= dt.Rows.Count - 1; l++)
                    {
                        DataRow dz = dt.Rows[l];
                        dr22 = dtS.NewRow();
                        string AccountCode = dz["AccountCode"].ToString();
                        string GruopAc = dz["GruopAc"].ToString();
                        string DeptCode = dz["DeptCode"].ToString();
                        dr22["AccountCode"] = AccountCode;
                        dr22["AccountName"] = dz["AccountName"].ToString();
                        dr22["GruopAc"] = dz["GruopAc"].ToString();
                        dr22["GruopName"] = dz["GruopName"].ToString();
                        dr22["DeptCode"] = dz["DeptCode"].ToString();
                        dr22["Dept"] = dz["Dept"].ToString();
                        dr22["參考3"] = dz["參考3"].ToString();

                        dr22["ProjectCode"] = dz["ProjectCode"].ToString();
                        dr22["Project"] = dz["Project"].ToString();
                        dr22["CustomerCode"] = dz["CustomerCode"].ToString();
                        dr22["Customer"] = dz["Customer"].ToString();
                        dr22["OpenBalance"] = dz["OpenBalance"];

                        string Q1 = "";
                        string Q2 = "";
                        string Q3 = "";
                        string Q4 = "";
                        int EndMon = Convert.ToInt32(cbMon2.Text);
                        if (EndMon == 1)
                        {
                            Q1 = dt.Compute("SUM([01])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 2)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 3)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 4)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 5)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 6)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 7)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 8)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])+SUM([08])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 9)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 10)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q4 = dt.Compute("SUM([10])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 11)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q4 = dt.Compute("SUM([10])+SUM([11])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }
                        else if (EndMon == 12)
                        {
                            Q1 = dt.Compute("SUM([01])+SUM([02])+SUM([03])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q2 = dt.Compute("SUM([04])+SUM([05])+SUM([06])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q3 = dt.Compute("SUM([07])+SUM([08])+SUM([09])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                            Q4 = dt.Compute("SUM([10])+SUM([11])+SUM([12])", "AccountCode='" + AccountCode + "' and GruopAc='" + GruopAc + "' and DeptCode='" + DeptCode + "'  ").ToString();
                        }

                        dr22["Q1"] = Q1;
                        dr22["Q2"] = Q2;
                        dr22["Q3"] = Q3;
                        dr22["Q4"] = Q4;
                        dr22["Total"] = dz["Total"];
                        dtS.Rows.Add(dr22);
                    }

                    if (globals.GroupID.ToString().Trim() == "ACCS" )
                    {

                        dtS.DefaultView.RowFilter = " DeptCode in (" + sb4 + ") ";
                    }
                    else
                    {
                        if (listBox2.SelectedItems.Count > 0)
                        {

      
                            dtS.DefaultView.RowFilter = " DeptCode in (" + sb3 + ") ";
                        }
                        else
                        {
                            if (ff != "All" && ff != "")
                            {
                                dtS.DefaultView.RowFilter = " DeptCode in (" + sb + ") ";
                            }
                        }
                    }
                    季.DataSource = dtS;
       
                    for (int i = FixedCol; i <= 季.Columns.Count - 1; i++)
                    {
                        DataGridViewColumn c = 季.Columns[i];
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        c.DefaultCellStyle.Format = "#,##0";

                    }


                    StringBuilder sbm3 = new StringBuilder();
                    System.Data.DataTable km1 = GetBU("TFT-ALL");
                    for (int i = 0; i <= km1.Rows.Count - 1; i++)
                    {

                        DataRow dd = km1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm3.Append("'" + F + "',");
                    }
                    sbm3.Remove(sbm3.Length - 1, 1);
                    GDABLE(sbm3,TFTALL);

                    StringBuilder sbm4 = new StringBuilder();
                    System.Data.DataTable km4 = GetBU("TFT-BU");
                    for (int i = 0; i <= km4.Rows.Count - 1; i++)
                    {

                        DataRow dd = km4.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm4.Append("'" + F + "',");
                    }
                    sbm4.Remove(sbm4.Length - 1, 1);
                    GDABLE(sbm4, TFTBU);

                    StringBuilder sbm5 = new StringBuilder();
                    System.Data.DataTable km5 = GetBU("TFT-Purchasing");
                    for (int i = 0; i <= km5.Rows.Count - 1; i++)
                    {

                        DataRow dd = km5.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm5.Append("'" + F + "',");
                    }
                    sbm5.Remove(sbm5.Length - 1, 1);
                    GDABLE(sbm5, TFTPurchasing);

                    StringBuilder sbm6 = new StringBuilder();
                    System.Data.DataTable km6 = GetBU("TFT-Logistic");
                    for (int i = 0; i <= km6.Rows.Count - 1; i++)
                    {

                        DataRow dd = km6.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm6.Append("'" + F + "',");
                    }
                    sbm6.Remove(sbm6.Length - 1, 1);
                    GDABLE(sbm6, TFTLogistic);


                    StringBuilder sbm7 = new StringBuilder();
                    System.Data.DataTable km7 = GetBU("TFT-Eng.Service");
                    for (int i = 0; i <= km7.Rows.Count - 1; i++)
                    {

                        DataRow dd = km7.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm7.Append("'" + F + "',");
                    }
                    sbm7.Remove(sbm7.Length - 1, 1);
                    GDABLE(sbm7, TFTEngService);

                    StringBuilder sbm8 = new StringBuilder();
                    System.Data.DataTable km8 = GetBU("ESCO");
                    for (int i = 0; i <= km8.Rows.Count - 1; i++)
                    {

                        DataRow dd = km8.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm8.Append("'" + F + "',");
                    }
                    sbm8.Remove(sbm8.Length - 1, 1);
                    GDABLE(sbm8, ESCO);

                    StringBuilder sbm9 = new StringBuilder();
                    System.Data.DataTable km9 = GetBU("NON-BU");
                    for (int i = 0; i <= km9.Rows.Count - 1; i++)
                    {

                        DataRow dd = km9.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sbm9.Append("'" + F + "',");
                    }
                    sbm9.Remove(sbm9.Length - 1, 1);
                    GDABLE(sbm9, NONBU);

                }

   
            }

  
        }
        private void GDABLE(StringBuilder  sbm3,DataGridView GD)
        {
            if (dt.Rows.Count > 0)
            {
                dt3 = MakeTable();
                for (int l = 0; l <= dt.Rows.Count - 1; l++)
                {
                    DataRow dz = dt.Rows[l];
                    dr23 = dt3.NewRow();
                    string AccountCode = dz["AccountCode"].ToString();
                    string GruopAc = dz["GruopAc"].ToString();
                    string DeptCode = dz["DeptCode"].ToString();
                    dr23["AccountCode"] = AccountCode;
                    dr23["AccountName"] = dz["AccountName"].ToString();
                    dr23["GruopAc"] = dz["GruopAc"].ToString();
                    dr23["GruopName"] = dz["GruopName"].ToString();
                    dr23["DeptCode"] = dz["DeptCode"].ToString();
                    dr23["Dept"] = dz["Dept"].ToString();
                    dr23["參考3"] = dz["參考3"].ToString();

                    dr23["ProjectCode"] = dz["ProjectCode"].ToString();
                    dr23["Project"] = dz["Project"].ToString();
                    dr23["CustomerCode"] = dz["CustomerCode"].ToString();
                    dr23["Customer"] = dz["Customer"].ToString();
                    dr23["OpenBalance"] = dz["OpenBalance"];

                    int StartMon = Convert.ToInt32(cbMon1.Text);
                    int EndMon = Convert.ToInt32(cbMon2.Text);

                    for (int i = StartMon; i <= EndMon; i++)
                    {
                        dr23[i.ToString("00")] = dz[i.ToString("00")];
                    }
                    dr23["Total"] = dz["Total"];
                    dt3.Rows.Add(dr23);
                }

     
    
     
                dt3.DefaultView.RowFilter = " DeptCode in (" + sbm3 + ") ";

                GD.DataSource = dt3.DefaultView;

                for (int i = FixedCol; i <= GD.Columns.Count - 1; i++)
                {
                    DataGridViewColumn c = GD.Columns[i];
                    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    c.DefaultCellStyle.Format = "#,##0";

                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
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
            //if (tabControl1.SelectedIndex == 0)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView1);
            //}
            //else if (tabControl1.SelectedIndex == 1)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView2);
            //}
            //else if (tabControl1.SelectedIndex == 2)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView3);
            //}
            //else if (tabControl1.SelectedIndex == 3)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView4);
            //}
            //else if (tabControl1.SelectedIndex == 4)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView5);
            //}
            //else if (tabControl1.SelectedIndex == 5)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView6);
            //}
            //else if (tabControl1.SelectedIndex == 6)
            //{
            //    ExcelReport.GridViewToExcel(dataGridView7);
            //}
        }
        private void WH(CarlosAg.ExcelXmlWriter.Workbook book, DataGridView DGV, string DD)
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

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];



                    if (j == 0 || j == 1 || j == 2 || j == 3 || j == 4 || j == 5 || j == 6 || j == 7 || j == 8 || j == 9 || j == 10)
                    {
                        rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                    }
                    else
                    {
                        rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                    }
                    
                    rowS.AutoFitHeight = true;
                    rowS.Table.DefaultColumnWidth = 100;

                }

            }
        }
        private void listBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            StringBuilder sb2 = new StringBuilder();
   
            ArrayList al2 = new ArrayList();
            if (listBox2.SelectedItems.Count > 0)
            {

                for (int i = 0; i <= listBox2.SelectedItems.Count - 1; i++)
                {
                    al2.Add(listBox2.SelectedItems[i].ToString());

                }

                foreach (string v in al2)
                {
                    sb2.Append("'" + v + "',");
                }
                sb2.Remove(sb2.Length - 1, 1);
            }
            string J1 = sb2.ToString();
            if (J1.IndexOf("TFT-ALL") != -1)
            {
                System.Data.DataTable k1 = GetBU("TFT-ALL");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox1.Items.Add(PrcCode);
                }


            }
            else
            {
                if (J1.IndexOf("TFT-BU") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT-BU");
                    DataRow dr;
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {
                        dr = k1.Rows[i];

                        string PrcCode = Convert.ToString(dr["DEPT"]);

                        listBox1.Items.Add(PrcCode);
                    }

                }
                if (J1.IndexOf("TFT-Purchasing") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT-Purchasing");
                    DataRow dr;
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {
                        dr = k1.Rows[i];

                        string PrcCode = Convert.ToString(dr["DEPT"]);

                        listBox1.Items.Add(PrcCode);
                    }

                }
                if (J1.IndexOf("TFT-Logistic") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT-Logistic");
                    DataRow dr;
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {
                        dr = k1.Rows[i];

                        string PrcCode = Convert.ToString(dr["DEPT"]);

                        listBox1.Items.Add(PrcCode);
                    }

                }
                if (J1.IndexOf("TFT-Eng.Service") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT-Eng.Service");
                    DataRow dr;
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {
                        dr = k1.Rows[i];

                        string PrcCode = Convert.ToString(dr["DEPT"]);

                        listBox1.Items.Add(PrcCode);
                    }

                }
            }
            if (J1.IndexOf("ESCO") != -1)
            {
                System.Data.DataTable k1 = GetBU("ESCO");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox1.Items.Add(PrcCode);
                }

            }
            if (J1.IndexOf("SOLAR") != -1)
            {
                System.Data.DataTable k1 = GetBU("SOLAR");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox1.Items.Add(PrcCode);
                }

            }
            if (J1.IndexOf("NON-BU") != -1)
            {
                System.Data.DataTable k1 = GetBU("NON-BU");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox1.Items.Add(PrcCode);
                }

   

            }
        }

  

    }//
}//




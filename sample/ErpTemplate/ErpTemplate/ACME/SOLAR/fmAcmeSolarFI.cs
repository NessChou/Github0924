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
    public partial class fmAcmeSolarFI : Form
    {


        string ProfitCode;

        private DataTable dtOcrd;

        private DataTable dtProject;

        private DataTable dtDept;

        int FixedCol = 6;

        public fmAcmeSolarFI()
        {
            InitializeComponent();
        }

        private void fmAcmeTrialReport_Load(object sender, EventArgs e)
        {

            dataGridView1.AutoGenerateColumns = false;
            dataGridView2.AutoGenerateColumns = false;
            dataGridView3.AutoGenerateColumns = false;

            txtStartDate.Text = DateTime.Now.ToString("yyyy")+"0101";
            txtEndDate.Text = DateTime.Now.ToString("yyyyMMdd");
            txtDate.Text = txtEndDate.Text;

            ProfitCode = GetProfitCode(); 



        }

        private void button1_Click(object sender, EventArgs e)
        {
        //    ProfitCode = GetProfitCode();

        //    DataTable dt = MakeTable();

        //    DataTable dtOpen = GetOpenBalance(cbYear.Text);

        //  // dataGridView1.DataSource = dtOpen;

        //    int iYear =Convert.ToInt32(cbYear.Text);
        //    string sYear =cbYear.Text;
        //    int iMon1 = Convert.ToInt32(cbMon1.Text);
        //    int iMon2 = Convert.ToInt32(cbMon2.Text);

    
        //    string StartDate


        //    DataTable dtAccount = GetAccount();
        


        //    DataTable[] ArrayDt = new DataTable[iMon2-iMon1+2];
        //    for (int j = iMon1; j <= iMon2; j++)
        //    {
        //        string Date1 = sYear + j.ToString("00") + "01";
        //        string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

        //        ArrayDt[j] = GetBalance(Date1, Date2);
        //    }

        //    //ArrayDt[0] = GetOpenBalance(sYear);
        //    ArrayDt[0] = dtOpen;


        //    string AcctCode = "";
        //    string AcctName = "";
        //    string Postable = "";

        //    DataRow dr;
        //    DataRow row;

        //    for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
        //    {
        //       dr = dtAccount.Rows[i];

        //       AcctCode = Convert.ToString(dr["AcctCode"]);
        //       AcctName = Convert.ToString(dr["AcctName"]);
        //       Postable = Convert.ToString(dr["Postable"]);

        //       if (Postable == "N")
        //       {
        //           continue;
        //       }

        //       row = dt.NewRow();
        //       row["AccountCode"] = AcctCode;
        //       row["AccountName"] = AcctName;


        //        DataRow drFind;
        //        Int64 Total =0;

        //        drFind = ArrayDt[0].Rows.Find(AcctCode);
        //        if (drFind != null)
        //        {
        //            row["OpenBalance"] = drFind["Balance"];
        //            Total += Convert.ToInt64(drFind["Balance"]);

        //        }

        //       for (int j = iMon1; j <= iMon2; j++)
        //       {


        //           drFind = ArrayDt[j].Rows.Find(AcctCode);
        //           if (drFind != null)
        //           {
        //               row[j.ToString("00")] = drFind["Balance"];
        //               Total += Convert.ToInt64(drFind["Balance"]);

        //           }

        //       }



        //       row["Total"] = Total;

        //       dt.Rows.Add(row);


        //    }

        //    dataGridView1.DataSource = dt;

        //    for (int i = FixedCol; i <= dataGridView1.Columns.Count - 1; i++)
        //    {
        //        DataGridViewColumn c = dataGridView1.Columns[i];
        //        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //        c.DefaultCellStyle.Format = "#,##0";
            
        //    }

        //    dataGridView1.Columns[2].Visible = false;
        //    dataGridView1.Columns[3].Visible = false;
        //    dataGridView1.Columns[4].Visible = false;
        //    dataGridView1.Columns[5].Visible = false;


        }


        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("AccountCode", typeof(string));
            dt.Columns.Add("AccountName", typeof(string));
            dt.Columns.Add("GruopAc", typeof(string));
            dt.Columns.Add("Dept", typeof(string));
            dt.Columns.Add("Project", typeof(string));
            dt.Columns.Add("Customer", typeof(string));

            dt.Columns.Add("OpenBalance", typeof(Int64));

            int StartMon = Convert.ToInt32(cbMon1.Text);
            int EndMon   = Convert.ToInt32(cbMon2.Text);

            for (int i = StartMon; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }



            dt.Columns.Add("Total", typeof(Int64));

            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["SlpCode"];
            //dt.PrimaryKey = colPk;


            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }





        private System.Data.DataTable GetOpenBalance(string Year)
        {
    
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account], SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");

            sb.Append(string.Format(" AND  T0.ProfitCode in ({0}) ",ProfitCode));
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
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" AND  T0.ProfitCode in ({0}) ", ProfitCode));
            }
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


        private System.Data.DataTable GetBalance_Sub(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project, SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],'')");


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
            DataColumn[] colPk = new DataColumn[3];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
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
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" Inner join  [OACT] T2  ON  T2.[AcctCode] = T0.Account   ");

            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],'')");


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
            DataColumn[] colPk = new DataColumn[3];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
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
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],'')");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],'')");


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
            DataColumn[] colPk = new DataColumn[3];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            dt.PrimaryKey = colPk;

            return dt;


        }




        private System.Data.DataTable GetAccount()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT T0.[AcctCode], T0.[AcctName] ,T0.[Postable]");
            //sb.Append(" FROM  [OACT] T0");
            //sb.Append(" ORDER BY T0.[AcctCode]");


            sb.Append(" SELECT distinct T1.[AcctCode], T1.[AcctName] ,T1.[Postable]");
            sb.Append(" FROM  [dbo].[JDT1] T0");
            sb.Append(" Inner join  [OACT] T1  ON  T1.AcctCode = T0.Account");
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" WHERE T0.ProfitCode in ({0}) ", ProfitCode));
            }
            sb.Append(" and T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" ORDER BY T1.[AcctCode]");

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

            ////

            //command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

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

            ////

            //command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

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


        /// <summary>
        /// 取出 T0.[Account],T1.[ProfitCode],T1.[Project] 列表
        /// </summary>
        /// <param name="RefDate1"></param>
        /// <param name="RefDate2"></param>
        /// <returns></returns>
        private System.Data.DataTable GetAccount_Sub_Ocrd(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T0.[Account] as AcctCode,T2.[AcctName],ISNULL(T0.[ProfitCode],'') as ProfitCode,T0.ShortName,ISNULL(T0.[Project],'') as Project ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" Inner join  [OACT] T2  ON  T2.[AcctCode] = T0.Account   ");

            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName");


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
            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["ShortName"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetOpenBalance_Sub_Ocrd(string Year)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project,T0.ShortName, SUM(T0.[SYSDeb]) Debit,SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) as Balance");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] < @P1 ");
            sb.Append(" AND  T0.[TransType] <> '-3'  ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName");


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
            colPk[3] = dt.Columns["ShortName"];
            dt.PrimaryKey = colPk;

            return dt;


        }

        private System.Data.DataTable GetBalance_Sub_Ocrd(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[Account],ISNULL(T0.[ProfitCode],'') as ProfitCode,ISNULL(T0.[Project],'') as Project,T0.ShortName, SUM(T0.[SYSDeb]) Debit, SUM(T0.[SYSCred]) Credit,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred]) Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId   ");
            sb.Append(" WHERE T0.[RefDate] >= @RefDate1  AND  T0.[RefDate] <= @RefDate2  ");
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   ");
            sb.Append(" GROUP BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName");
            sb.Append(" ORDER BY T0.[Account],ISNULL(T0.[ProfitCode],''),ISNULL(T0.[Project],''),T0.ShortName");


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
            DataColumn[] colPk = new DataColumn[4];
            colPk[0] = dt.Columns["Account"];
            colPk[1] = dt.Columns["ProfitCode"];
            colPk[2] = dt.Columns["Project"];
            colPk[3] = dt.Columns["ShortName"];
            dt.PrimaryKey = colPk;

            return dt;


        }


        private void button2_Click(object sender, EventArgs e)
        {
            

            DataTable dt = MakeTable();

            DataTable dtOpen = GetOpenBalance(cbYear.Text);

            dataGridView1.DataSource = dtOpen;

            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);



            string RefDate1 = cbYear.Text + cbMon1.Text + "01";
            string RefDate2 = cbYear.Text + cbMon2.Text + DateTime.DaysInMonth(iYear, iMon2);
            DataTable dtAccount = GetAccount_Sub(RefDate1, RefDate2);
            

            DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00") + "01";
                string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

                ArrayDt[j] = GetBalance_Sub(Date1, Date2);
            }
            ArrayDt[0] = GetOpenBalance_Sub(sYear);


            string AcctCode = "";
            string AcctName = "";
           

            string ProfitCode;
            string Project;

            DataRow dr;
            DataRow row;

            

            for (int i = 0; i <= dtAccount.Rows.Count - 1; i++)
            {
                dr = dtAccount.Rows[i];

                AcctCode = Convert.ToString(dr["AcctCode"]);
                AcctName = Convert.ToString(dr["AcctName"]);


                ProfitCode = Convert.ToString(dr["ProfitCode"]);
                Project = Convert.ToString(dr["Project"]);

                Object[] Key = new object[]{AcctCode,ProfitCode,Project};

                //Postable = Convert.ToString(dr["Postable"]);

                //if (Postable == "N")
                //{
                //    continue;
                //}

                row = dt.NewRow();
                row["AccountCode"] = AcctCode;
                row["AccountName"] = AcctName;


                row["Dept"] = ProfitCode;
                row["Project"] = Project;


                DataRow drFind;
                Int64 Total = 0;

                drFind = ArrayDt[0].Rows.Find(Key);
                if (drFind != null)
                {
                    row["OpenBalance"] = drFind["Balance"];
                    Total += Convert.ToInt64(drFind["Balance"]);

                }

                for (int j = iMon1; j <= iMon2; j++)
                {


                    drFind = ArrayDt[j].Rows.Find(Key);
                    if (drFind != null)
                    {
                        row[j.ToString("00")] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }

                }



                row["Total"] = Total;

                dt.Rows.Add(row);


            }

            dataGridView1.DataSource = dt;

            for (int i = FixedCol; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = dataGridView1.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }
            dataGridView1.Columns[5].Visible = false;
            
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

            dataGridView1.DataSource = dtAccount;
        }

        private void btnAcme_Click(object sender, EventArgs e)
        {

 


            DataTable dt = MakeTable();

            DataTable dtOpen = GetOpenBalance(cbYear.Text);

            dataGridView1.DataSource = dtOpen;

            int iYear = Convert.ToInt32(cbYear.Text);
            string sYear = cbYear.Text;
            int iMon1 = Convert.ToInt32(cbMon1.Text);
            int iMon2 = Convert.ToInt32(cbMon2.Text);



            string RefDate1 = cbYear.Text + cbMon1.Text + "01";
            string RefDate2 = cbYear.Text + cbMon2.Text + DateTime.DaysInMonth(iYear, iMon2);
            DataTable dtAccount = GetAccount_Sub_Ocrd(RefDate1, RefDate2);


            DataTable[] ArrayDt = new DataTable[iMon2 - iMon1 + 2];
            for (int j = iMon1; j <= iMon2; j++)
            {
                string Date1 = sYear + j.ToString("00") + "01";
                string Date2 = sYear + j.ToString("00") + DateTime.DaysInMonth(iYear, j);

                ArrayDt[j] = GetBalance_Sub_Ocrd(Date1, Date2);
            }
            ArrayDt[0] = GetOpenBalance_Sub_Ocrd(sYear);


            string AcctCode = "";
            string AcctName = "";


            string ProfitCode;
            string Project;

            string ShortName;

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



                ShortName = Convert.ToString(dr["ShortName"]);

                Object[] Key = new object[] { AcctCode, ProfitCode, Project, ShortName };

                //Postable = Convert.ToString(dr["Postable"]);

                //if (Postable == "N")
                //{
                //    continue;
                //}

                row = dt.NewRow();
                row["AccountCode"] = AcctCode;
                row["AccountName"] = AcctName;



                if (! string.IsNullOrEmpty(ProfitCode))
                {
                    drFind = dtDept.Rows.Find(ProfitCode);
                    if (drFind != null)
                    {
                        
                        row["Dept"] = drFind["PrcName"];
                    }
                    //row["Dept"] = ProfitCode;
                
                }


                if (!string.IsNullOrEmpty(Project))
                {
                    drFind = dtProject.Rows.Find(Project);
                    if (drFind != null)
                    {

                        row["Project"] = drFind["PrjName"];
                    }
                    //row["Dept"] = ProfitCode;

                }


                //row["Dept"] = ProfitCode;
                //row["Project"] = Project;


                
                if (ShortName == AcctCode)
                {

                }
                else
                {
                    drFind = dtOcrd.Rows.Find(ShortName);
                    if (drFind != null)
                    {
                        //row["Customer"] = ShortName;
                        row["Customer"] = drFind["CardName"];
                    }
                    
                }



                
                Int64 Total = 0;


                if (AcctCode.Substring(0, 1) == "1" || AcctCode.Substring(0, 1) == "2" || AcctCode.Substring(0, 1) == "3")
                {

                    drFind = ArrayDt[0].Rows.Find(Key);
                    if (drFind != null)
                    {
                        row["OpenBalance"] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }
                }

                for (int j = iMon1; j <= iMon2; j++)
                {


                    drFind = ArrayDt[j].Rows.Find(Key);
                    if (drFind != null)
                    {
                        row[j.ToString("00")] = drFind["Balance"];
                        Total += Convert.ToInt64(drFind["Balance"]);

                    }

                }



                row["Total"] = Total;

                dt.Rows.Add(row);


            }

            dataGridView1.DataSource = dt;

            for (int i = FixedCol; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = dataGridView1.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }
            //dataGridView1.Columns[5].Visible = false;
        }



        private string GetProfitCode()
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT dept COLLATE Chinese_PRC_CI_AS FROM acmesqlsp..ACCOUNT_BU where bu='solar' ");



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
                da.Fill(ds, "OPRC");
            }
            finally
            {
                connection.Close();
                this.Cursor = Cursors.Default;
            }
            DataTable dt = ds.Tables[0];
            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["PrcCode"];
            //dt.PrimaryKey = colPk;

            string s = "";

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                s = s + "'" + Convert.ToString(dt.Rows[i][0]) + "',";
            }

            s = s.Substring(0, s.Length - 1);

            return s;



        }


        public DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        public string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtStartDate.Text) ||
                string.IsNullOrEmpty(txtEndDate.Text) ||
                string.IsNullOrEmpty(txtDate.Text))
            {

                MessageBox.Show("日期必須輸入");
                return;
            }

            try
            {
                StrToDate(txtStartDate.Text);
                StrToDate(txtEndDate.Text);
                StrToDate(txtEndDate.Text);

            }

            catch

            {
                MessageBox.Show("請輸入正確日期");
                return;
            }


            string RefDate1 = txtStartDate.Text;
            string RefDate2 = txtEndDate.Text;

            DataTable dt = GetBalanceNew( RefDate1,  RefDate2);
            dataGridView1.DataSource = dt;


            RefDate2 = DateToStr(StrToDate(RefDate1).AddDays(-1));
            RefDate1 = "20070101";

            DataTable dtOpen = GetBalanceNew_Fixed(RefDate2);

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dtOpen.Columns["科目代號"];
            dtOpen.PrimaryKey = colPk;


            DataRow dr;
            DataRow drFind;
            string AccCode;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dt.Rows[i];

                AccCode = Convert.ToString(dr["科目代號"]);
                dr.BeginEdit();

                drFind = dtOpen.Rows.Find(AccCode);

                if (drFind != null)
                {

                    dr["期初餘額"] = Convert.ToInt32(drFind["餘額"]);
                }
                dr.EndEdit();
            
            }


            DataView dv = dt.DefaultView;
            

            dv.RowFilter = "科目代號 > '3'";
            DataTable dtT = dv.ToTable();

            dv.RowFilter = "科目代號 > ''";

            dataGridView3.DataSource = dtT;



            RefDate1 ="20070101";
            RefDate2 = txtDate.Text;

            DataTable dtBS = GetBalanceNewBS_Fixed(RefDate2);
     
            dataGridView2.DataSource = dtBS;

            DataView dv1 = dtBS.DefaultView;
            dv1.RowFilter = "科目代號 < '4'";

        }



        private System.Data.DataTable GetBalanceNew(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T1.AcctCode as 科目代號,T1.[AcctName] 科目名稱,convert(int,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred])) as 餘額,0 as 期初餘額");
            sb.Append(" FROM  [OACT] T1 ");
            sb.Append(" Inner join  [JDT1] T0  ON  T1.AcctCode = T0.Account");
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" Where T0.RefDate >=@RefDate1 and  T0.RefDate <=@RefDate2 and T0.ProfitCode in ({0}) ", ProfitCode));
            }
            else
            {
                sb.Append(" Where T0.RefDate >=@RefDate1 and  T0.RefDate <=@RefDate2 ");
            }
            sb.Append(" group by T1.AcctCode,T1.[AcctName]");
            sb.Append(" order by T1.AcctCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            command.CommandTimeout = 300;

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
        

            return dt;


        }


        private System.Data.DataTable GetBalanceNew_Fixed(string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T1.AcctCode as 科目代號,T1.[AcctName] 科目名稱,convert(int,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred])) as 餘額,0 as 期初餘額");
            sb.Append(" FROM  [OACT] T1 ");
            sb.Append(" Inner join  [JDT1] T0  ON  T1.AcctCode = T0.Account");
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" Where T0.RefDate <='{0}' and T0.ProfitCode in ({1}) ", RefDate2, ProfitCode));
            }
            else
            {
                sb.Append(string.Format(" Where T0.RefDate <='{0}' ", RefDate2));
            }
            sb.Append(" group by T1.AcctCode,T1.[AcctName]");
            sb.Append(" order by T1.AcctCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            command.CommandTimeout = 300;

            //

            //command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));
            // command.Parameters.Add(new SqlParameter("@ProfitCode", ProfitCode));

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


            return dt;


        }


        private System.Data.DataTable GetBalanceNewBS(string RefDate1, string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T1.AcctCode as 科目代號,T1.[AcctName] 科目名稱,convert(int,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred])) as 餘額");
            sb.Append(" FROM  [OACT] T1 ");
            sb.Append(" Inner join  [JDT1] T0  ON  T1.AcctCode = T0.Account");
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" Where T0.RefDate >=@RefDate1 and  T0.RefDate <=@RefDate2 and T0.ProfitCode in ({0}) ", ProfitCode));
            }
            else
            {
                sb.Append("Where T0.RefDate >=@RefDate1 and  T0.RefDate <=@RefDate2  ");
            }
            sb.Append(" group by T1.AcctCode,T1.[AcctName]");
            sb.Append(" order by T1.AcctCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));
            // command.Parameters.Add(new SqlParameter("@ProfitCode", ProfitCode));

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


            return dt;


        }


        private System.Data.DataTable GetBalanceNewBS_Fixed(string RefDate2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T1.AcctCode as 科目代號,T1.[AcctName] 科目名稱,convert(int,SUM(T0.[SYSDeb])-SUM(T0.[SYSCred])) as 餘額");
            sb.Append(" FROM  [OACT] T1 ");
            sb.Append(" Inner join  [JDT1] T0  ON  T1.AcctCode = T0.Account");
            if (globals.DBNAME == "進金生")
            {
                sb.Append(string.Format(" Where T0.RefDate <='{0}' and T0.ProfitCode in ({1}) ", RefDate2, ProfitCode));
            }
            else
            {
                sb.Append(string.Format(" Where T0.RefDate <='{0}'  ", RefDate2));
            }
            sb.Append(" group by T1.AcctCode,T1.[AcctName]");
            sb.Append(" order by T1.AcctCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            //command.Parameters.Add(new SqlParameter("@RefDate1", RefDate1));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));
            //// command.Parameters.Add(new SqlParameter("@ProfitCode", ProfitCode));

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


            return dt;


        }

        //傳入參數
        //dataGridView
        //輸出文字檔 ,附檔名為 csv
        //使用範例  GridViewToCSV(dataGridView1, Environment.CurrentDirectory + @"\dataGridview.csv");
        private void GridViewToCSV(DataGridView dgv, string FileName)
        {

            StringBuilder sbCSV = new StringBuilder();
            int intColCount = dgv.Columns.Count;
            //int intColCount = dgv.Cells.Count;


            //表頭
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                sbCSV.Append(dgv.Columns[i].HeaderText);

                if ((i + 1) != intColCount)
                {
                    sbCSV.Append(",");
                    //tab
                    // sbCSV.Append("\t");
                }

            }
            sbCSV.Append("\n");

            foreach (DataGridViewRow dr in dgv.Rows)
            {

                //資料內容
                for (int x = 0; x < intColCount; x++)
                {

                    if (dr.Cells[x].Value != null)
                    {

                        sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));
                    }
                    else
                    {
                        sbCSV.Append("");
                    }


                    if ((x + 1) != intColCount)
                    {
                        sbCSV.Append(",");
                        // sbCSV.Append("\t");
                    }
                }
                sbCSV.Append("\n");
            }
            using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.Default))
            {
                sw.Write(sbCSV.ToString());
            }

            System.Diagnostics.Process.Start(FileName);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView1,"試算表.csv");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView2, "資產負債表.csv");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView3, "損益表.csv");
        }

        private void txtEndDate_TextChanged(object sender, EventArgs e)
        {
            txtDate.Text = txtEndDate.Text;
        }

    }//
}




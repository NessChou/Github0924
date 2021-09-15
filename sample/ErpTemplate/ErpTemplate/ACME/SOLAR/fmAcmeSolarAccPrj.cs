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
    public partial class fmAcmeSolarAccPrj : Form
    {

 

        string ProfitCode;

        public fmAcmeSolarAccPrj()
        {
            InitializeComponent();

            dataGridView1.AutoGenerateColumns = false;
            dataGridView2.AutoGenerateColumns = false;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
          //  MessageBox.Show( tabControl1.SelectedIndex.ToString());

            //string AccGroup = tabControl1.SelectedIndex.ToString();

            //DataTable dt  = GetAcc(AccGroup);
            //dataGridView1.DataSource = dt;

            //if (dt.Rows.Count > 0)
            //{

            //    // dataGridView1.Rows[0].Selected=true;
            //    dataGridView1.Focus();
            //    dataGridView1_SelectionChanged(dataGridView1, e);
            //}
        }


        private string GetProfitCode()
        {
            this.Cursor = Cursors.AppStarting;
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
           
            sb.Append(" SELECT dept COLLATE Chinese_PRC_CI_AS FROM acmesqlsp..ACCOUNT_BU where bu='solar' ");

           

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
                this.Cursor = Cursors.Default;
            }
            DataTable dt = ds.Tables[0];
    

            string s="";

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                s = s +"'"+ Convert.ToString(dt.Rows[i][0])+"',";
            }

            s = s.Substring(0, s.Length - 1);

            return s;



        }

        private System.Data.DataTable GetAcc(string AccGroup)
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
        
            sb.Append(" SELECT distinct T1.AcctCode as 科目代號,T1.[AcctName] 科目名稱");
            sb.Append(" FROM  [dbo].[JDT1] T0");
            sb.Append(" Inner join  [OACT] T1  ON  T1.AcctCode = T0.Account WHERE 1=1 ");
            if (globals.DBNAME != "進金生能源服務")
            {
                sb.Append(string.Format(" AND T0.ProfitCode in ({0}) ", ProfitCode));
            }
            if (AccGroup == "0")
            {
                sb.Append(" and substring(T1.AcctCode,1,1)='1'");
            }
            else if (AccGroup == "1")
            {
                sb.Append(" and substring(T1.AcctCode,1,1)='2'");
            } 
            else if (AccGroup == "2")
            {
                sb.Append(" and substring(T1.AcctCode,1,1)='3'");
            } 
            else if (AccGroup == "3")
            {
                sb.Append(" and substring(T1.AcctCode,1,1)='4'");
            } 
            else if (AccGroup == "4")
            {
                sb.Append(" and substring(T1.AcctCode,1,1) in ('5','6','7','8','9')");
            } 


          //  sb.Append(" group by T1.AcctCode,T1.AcctName");
            sb.Append(" order by T1.AcctCode");


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
                this.Cursor = Cursors.Default;
            }
            DataTable dt = ds.Tables[0];


            return dt;

            

        }


        private System.Data.DataTable GetSolarPrj()
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT T0.[PrjCode], T0.[PrjName],Convert(varchar(10), T0.[U_PRJDATE],111) as U_PRJDATE ,  T0.[U_BU] FROM OPRJ T0 WHERE T0.[U_BU] ='SOLAR'");
          

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
                this.Cursor = Cursors.Default;
            }
            DataTable dt = ds.Tables[0];
            //DataColumn[] colPk = new DataColumn[1];
            //colPk[0] = dt.Columns["PrcCode"];
            //dt.PrimaryKey = colPk;

            return dt;



        }

        //'11030','11131','12131','12132','12331','12631','12632','12633','12731','12831'
        private void fmAcmeSolarAcc_Load(object sender, EventArgs e)
        {

            DataTable dt = GetSolarPrj();
            dataGridView1.DataSource = dt;



        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
           // return;

            //避免觸發
            DataGridView dgv = (DataGridView)sender;
            if (!dgv.Focused) return;


            //
            string PrjCode ="";



            PrjCode = dgv.CurrentRow.Cells[0].Value.ToString();
            //MessageBox.Show(Account);

            DataTable dt = GetAccDetail_Prj(PrjCode);

            dataGridView2.DataSource = dt;


            DataTable dtSum = GetAccSummary_Prj(PrjCode);

            dataGridView3.DataSource = dtSum;

             DataTable dt4= GetTotal_Prj(PrjCode);

            dataGridView4.DataSource = dt4;


        
            

            //DataRow dr;
            //Int32 Balance = 0;
            //for (int i = 0; i <= dt.Rows.Count - 1; i++)
            //{
            //    dr = dt.Rows[i];
            //    dr.BeginEdit();

            //    Balance = Balance + Convert.ToInt32(dr["SYSDeb"]) - Convert.ToInt32(dr["SYSCred"]);
            //    dr["Balance"] = Balance;
            //    dr.EndEdit();
            //}

            //label3.Text = Balance.ToString("#,##0");

            //label4.Text = Account + "-" + dgv.CurrentRow.Cells[1].Value.ToString();

        }

        private System.Data.DataTable GetAccDetail(string Account)
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.* ,0 as Balance");
            sb.Append(" FROM  [dbo].[JDT1] T0 WHERE 1=1  ");
            if (globals.DBNAME != "進金生能源服務")
            {
                sb.Append(string.Format(" AND T0.ProfitCode in ({0}) ", ProfitCode));
            }
            sb.Append(" and  T0.[Account]=@Account");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            command.Parameters.Add(new SqlParameter("@Account", Account));
            

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

            return dt;

            

        }

        private System.Data.DataTable GetAccDetail_Prj(string PrjCode)
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.* ,T1.AcctName ");
            sb.Append(" FROM  [dbo].[JDT1] T0  ");
            sb.Append(" Inner Join OACT T1  on T1.AcctCode = T0.Account  WHERE 1=1 ");
            if (globals.DBNAME != "進金生能源服務")
            {
                sb.Append(string.Format(" AND T0.Project=@PrjCode ", PrjCode));
            }
            sb.Append(" order by T0.[TransId], T0.[Line_ID]");
            
           




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));


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

            return dt;



        }

        private System.Data.DataTable GetAccSummary_Prj(string PrjCode)
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.Account,T1.AcctName,Sum(T0.SYSDeb) as SYSDeb,Sum(T0.SYSCred) as SYSCred,Sum(T0.SYSDeb-T0.SYSCred) as Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  ");
            sb.Append(" Inner Join OACT T1  on T1.AcctCode = T0.Account WHERE 1=1  ");
            if (globals.DBNAME != "進金生能源服務")
            {
                sb.Append(string.Format(" AND T0.Project=@PrjCode ", PrjCode));
            }
            sb.Append(" Group by T0.Account,T1.AcctName");
            sb.Append(" order by T0.Account");






            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));


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

            return dt;



        }


        private System.Data.DataTable GetTotal_Prj(string PrjCode)
        {
            this.Cursor = Cursors.AppStarting;
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Case when substring(T0.Account,1,1)='4' then '收入' when substring(T0.Account,1,1)='5' then '成本' when substring(T0.Account,1,1)='6' then '費用'  else '' end as AcctName,Sum( Case when substring(T0.Account,1,1)='4' then (T0.SYSDeb-T0.SYSCred) * (-1) else T0.SYSDeb-T0.SYSCred end ) as Balance ");
            sb.Append(" FROM  [dbo].[JDT1] T0  ");
            sb.Append(" Inner Join OACT T1  on T1.AcctCode = T0.Account WHERE 1=1 ");

            if (globals.DBNAME != "進金生能源服務")
            {
                sb.Append(string.Format(" AND T0.Project=@PrjCode ", PrjCode));
            }

            sb.Append(" and substring(T0.Account,1,1) >='4' ");

            sb.Append(" Group by substring(T0.Account,1,1)");
          






            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));


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

            return dt;



        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

            if (e.RowIndex == -1) return;

            DataGridView dgv = (DataGridView)sender;
            DataGridViewRow dgr = dgv.Rows[e.RowIndex];
            DataRowView row = (DataRowView)dgv.Rows[e.RowIndex].DataBoundItem;

           
            if (e.ColumnIndex ==6)
            {
                string s = Convert.ToString(e.Value);

                if (s == "30")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "JE";
                }
                else if (s == "15")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Delivery";
                }
                else if (s == "16")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Returns";
                }
                else if (s == "13")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/R Invoice";
                }
                else if (s == "14")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/R Credit Memo";
                }
                else if (s == "132")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Correction Invoice";
                }
                else if (s == "20")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Receipt";
                }
                else if (s == "21")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Returns";
                }
                else if (s == "18")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/P Invoice";
                }
                else if (s == "19")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/P Credit Memo";
                }
                else if (s == "-2")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Opening Balance";
                }
                else if (s == "58")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Stock Update";
                }
                else if (s == "59")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Receipt";
                }
                else if (s == "60")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Issue";
                }
                else if (s == "67")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Inventory Transfers";
                }
                else if (s == "67")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Inventory Transfers";
                }
                else if (s == "68")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Work Instructions";
                }
                else if (s == "-1")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "All Transactions";
                }
            }

            //   Case
            //WHEN T0.[TransType]=15  THEN 'Delivery' 
            //WHEN T0.[TransType]=16  THEN 'Returns'
            //WHEN T0.[TransType]=13  THEN 'A/R Invoice'
            //WHEN T0.[TransType]=14  THEN 'A/R Credit Memo'
            //WHEN T0.[TransType]=132 THEN 'Correction Invoice'
            //WHEN T0.[TransType]=20  THEN 'Goods Receipt'
            //WHEN T0.[TransType]=21  THEN 'Goods Returns'
            //WHEN T0.[TransType]=18  THEN 'A/P Invoice'
            //WHEN T0.[TransType]=19  THEN 'A/P Credit Memo'
            //WHEN T0.[TransType]=-2  THEN 'Opening Balance'
            //WHEN T0.[TransType]=58  THEN 'Stock Update'
            //WHEN T0.[TransType]=59  THEN 'Goods Receipt'
            //WHEN T0.[TransType]=60  THEN 'Goods Issue'
            //WHEN T0.[TransType]=67  THEN 'Inventory Transfers'
            //WHEN T0.[TransType]=68  THEN 'Work Instructions'
            //WHEN T0.[TransType]=-1  THEN 'All Transactions'
            //ELSE 'Other'
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            string TransID = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            //  MessageBox.Show(TransID);

            fmSolarAcc2 f = new fmSolarAcc2(TransID);
            f.ShowDialog();
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

        private void button1_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView2,  DateTime.Now.ToString("HHmmss")+".csv");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView3, DateTime.Now.ToString("HHmmss") + ".csv");
        }


    }
}


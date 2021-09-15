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
    public partial class FmRivaRep02 : Form
    {

        public static string ConnectiongString02 = "server=10.10.1.40;pwd=riv@green168;uid=rivagreen;database=CHIComp02";
        public static string EEPConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=Acmesqlsp";

        public FmRivaRep02()
        {
            InitializeComponent();
        }

        private void button95_Click(object sender, EventArgs e)
        {
            string StartDate = txtStartDate.Text;
            string EndDate = txtEndDate.Text;


            //  DataTable dt = GetBill();

            //訂單單號 -> 分批時 ->可能找到多筆
            //零售 批發 規則

            DataTable dt = GetBill_DN(StartDate, EndDate);
            dataGridView1.DataSource = dt;

            DataRow dr;
            string Key = "";
            string BillNo;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = dt.Rows[i];

                BillNo = Convert.ToString(dr["訂單單號"]);

                if (BillNo != Key)
                {

                    Key = BillNo;


                    dr.BeginEdit();


                    DataTable dtBill = GetBill(BillNo);

                    dr["理貨箱"] = "0";

                    if (dtBill.Rows.Count == 1)
                    {
                        dr["理貨箱"] = Convert.ToString(dtBill.Rows[0]["箱"]);
                    }
                    else if (dtBill.Rows.Count > 1)
                    {
                        dr["理貨箱"] = "資料多筆-" + Convert.ToString(dtBill.Rows[0]["箱"]);
                    }

                    dr.EndEdit();
                }
                else
                {

                }


            }
        }



        private DataTable GetBill_DN(string StartDate, string EndDate)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString02);

            StringBuilder sb = new StringBuilder();

           
            sb.Append("  Select  SubString(C.ClassName,5,2)  類別,");
            sb.Append("              A.BillNO 訂單單號,");
            sb.Append("              A.BillDate 訂單日期,");
            sb.Append("              A.CustomerID 客戶編號,");
            sb.Append("              B.ShortName 簡稱,");
            sb.Append("              Convert(varchar(8),G.PreInDate) 預交日期,");
            sb.Append("              S.UDef2 快遞單號,");
            sb.Append("              O.[ProdID] 產品編號,");
            sb.Append("       O.[ProdName] 產品名稱,");
            sb.Append("           O.[Quantity] 數量,");
            sb.Append("           S.FundBillNo 銷貨單號, ''as 理貨箱");
            sb.Append("              From OrdBillMain A ");
            sb.Append("              Inner Join OrdBillSub G On G.Flag=A.Flag And G.BillNO=A.BillNO ");

            sb.Append("              Left Join comCustomer B On B.Flag=A.Flag-1 And B.ID=A.CustomerID ");
            // sb.Append("              Left Join comCustClass  C on B.ClassID =C.ClassID and C.Flag=1");


            sb.Append("              left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500");
            sb.Append("              left join comBillAccounts S ON (O.BillNO =S.FundBillNo AND S.Flag =500)");
            sb.Append("              Inner Join comProduct R on R.ProdId=O.Prodid");
            sb.Append("              Left  Join  comProductClass  C on R.ClassID =C.ClassID ");

            sb.Append("              Where A.Flag=2 ");
            // sb.Append("              And A.BillDate >= '20140601'  And A.BillDate <='20140630'");
            sb.Append("              And A.BillDate >= @StartDate  And A.BillDate <=@EndDate ");
            //FREIGHT01
            sb.Append("              And O.ProdID <> 'FREIGHT01' order by A.BillNo");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            command.Parameters.Add(new SqlParameter("@EndDate", EndDate));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private DataTable GetBill(string BILLNO)
        {

            SqlConnection connection = new SqlConnection(EEPConnectiongString);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT gtype 客戶別,BILLNO 訂單號碼, ");
            sb.Append("case when gtype='零售' then max(PACK1) else sum(PACK1) end as 箱  ");
            sb.Append("from GB_PICK2  ");
            sb.Append("WHERE  BILLNO=@BILLNO ");
            sb.Append("group by gtype,BILLNO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private void FmRivaRep02_Load(object sender, EventArgs e)
        {
            txtStartDate.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "01"; ;
            txtEndDate.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + DateTime.DaysInMonth(DateTime.Now.AddMonths(-1).Year, DateTime.Now.AddMonths(-1).Month).ToString();
        }

        private void button86_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView1, "對帳箱數.csv");
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
    }
}
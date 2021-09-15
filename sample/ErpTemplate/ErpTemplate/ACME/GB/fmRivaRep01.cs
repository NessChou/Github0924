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
    public partial class fmRivaRep01 : Form
    {

        public static string ConnectiongString02 = "server=10.10.1.40;pwd=riv@green168;uid=rivagreen;database=CHIComp02";

        public fmRivaRep01()
        {
            InitializeComponent();
        }

        private System.Data.DataTable MakeTableEngName()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("排行", typeof(string));
            dt.Columns.Add("部位", typeof(string));
            dt.Columns.Add("庫存量", typeof(Int32));
            dt.Columns.Add("庫存佔比", typeof(string));
            dt.Columns.Add("庫存金額", typeof(Int32));
            dt.Columns.Add("銷貨成本", typeof(Int32));
            dt.Columns.Add("銷貨數量", typeof(Int32));

            dt.Columns.Add("週轉率", typeof(string));


            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["部位"];
            dt.PrimaryKey = colPk;


            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }


        private DataTable GetChiStockHis_EngName_Weight_All(string StartDate, string EndDate, string Category)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString02);

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT B.EngName, ");
            sb.Append("SUM(CASE WHEN A.Flag BETWEEN 100 AND 199 THEN A.CostForAcc ");
            sb.Append("WHEN A.Flag BETWEEN 200 AND 299 THEN -A.CostForAcc ");
            sb.Append("WHEN A.Flag BETWEEN 300 AND 399  AND A.Flag NOT IN (318, 319, 320)  THEN A.CostForAcc ");
            sb.Append("WHEN A.Flag BETWEEN 500 AND 599 THEN -A.CostForAcc ");
            sb.Append("WHEN A.Flag BETWEEN 600 AND 699 THEN A.CostForAcc ");
            sb.Append("WHEN A.Flag = 700 THEN -A.CostForAcc ELSE 0 END) AS 庫存金額, ");

            //數量換算
            sb.Append("SUM(Case When B.CtmWeight<>0 then CASE WHEN A.Flag BETWEEN 100 AND 199 THEN A.Quantity * B.CtmWeight ");
            sb.Append("WHEN A.Flag BETWEEN 200 AND 299 THEN -A.Quantity * B.CtmWeight ");
            sb.Append("WHEN A.Flag BETWEEN 300 AND 399  AND A.Flag NOT IN (318, 319, 320)  THEN A.Quantity * B.CtmWeight ");
            sb.Append("WHEN A.Flag BETWEEN 500 AND 599 THEN -(A.Quantity - A.QuanComb) * B.CtmWeight ");
            sb.Append("WHEN A.Flag BETWEEN 600 AND 699 THEN A.Quantity * B.CtmWeight ELSE 0 END else ");
            sb.Append("CASE WHEN A.Flag BETWEEN 100 AND 199 THEN A.Quantity ");
            sb.Append("WHEN A.Flag BETWEEN 200 AND 299 THEN -A.Quantity ");
            sb.Append("WHEN A.Flag BETWEEN 300 AND 399  AND A.Flag NOT IN (318, 319, 320)  THEN A.Quantity ");
            sb.Append("WHEN A.Flag BETWEEN 500 AND 599 THEN -(A.Quantity - A.QuanComb) ");
            sb.Append("WHEN A.Flag BETWEEN 600 AND 699 THEN A.Quantity ELSE 0 END end) AS 庫存數量 ");



            sb.Append("FROM ComProdRec A ");
            sb.Append("Inner Join ComProduct B on A.ProdId=B.ProdID ");

            if (Category == "豬")
            {
                sb.Append("Where B.EngName in (select distinct EngName from ComProduct where ClassId in ('AWP200','ARP200') and EngName<>'' ) ");
            }
            else if (Category == "雞")
            {
                //R.ClassId in ('AWC200','ARC200','AWP200','ARP200')
                sb.Append("Where B.EngName in (select distinct EngName from ComProduct where ClassId in ('AWC200','ARC200')) ");
            }
            sb.Append("AND A.Needupdate = 1 ");
            sb.Append("AND A.HasCheck = 1 ");
            sb.Append("AND A.BillDate >=@StartDate ");
            sb.Append("AND A.BillDate <=@EndDate ");
            sb.Append("GROUP BY B.EngName");


            //AWC200// 朝貢雞-批發
            //ARC200// 朝貢雞-零售
            //AWP200// 朝貢豬-批發
            //ARP200// 朝貢豬-零售

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            // command.Parameters.Add(new SqlParameter("@EngName", EngName));

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


        /// <summary>
        /// 20170722 部位群組
        /// </summary>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
        /// <param name="Flag"></param>
        /// <returns></returns>
        private DataTable GetChiStockIn_EngName(string StartDate, string EndDate, string Flag, string Category)
        {

            SqlConnection connection = new SqlConnection(ConnectiongString02);

            StringBuilder sb = new StringBuilder();
            sb.Append("select ");
            sb.Append("B.EngName, ");


            //if (ckFoc.Checked)
            //{
            //    sb.Append("Case WHEN T4.IsGift=1 then Sum(T1.Quantity)  else 0 end as 贈品數量,");
            //    sb.Append("Case WHEN T4.IsGift=1 then 'Y'  else 'N' end as 贈品,");

            //}

            if (Flag == "銷")
            {
                sb.Append("Case WHEN T1.Flag/100 in (1, 5) THEN Sum(T1.CostForAcc)  ELSE -Sum(T1.CostForAcc) END AS 成本, ");
            }


            //數量換算
            //sb.Append("SUM(Case When B.CtmWeight<>0 then CASE WHEN A.Flag BETWEEN 100 AND 199 THEN A.Quantity * B.CtmWeight ");
            //sb.Append("WHEN A.Flag BETWEEN 200 AND 299 THEN -A.Quantity * B.CtmWeight ");
            //sb.Append("WHEN A.Flag BETWEEN 300 AND 399  AND A.Flag NOT IN (318, 319, 320)  THEN A.Quantity * B.CtmWeight ");
            //sb.Append("WHEN A.Flag BETWEEN 500 AND 599 THEN -(A.Quantity - A.QuanComb) * B.CtmWeight ");
            //sb.Append("WHEN A.Flag BETWEEN 600 AND 699 THEN A.Quantity * B.CtmWeight ELSE 0 END else ");
            //sb.Append("CASE WHEN A.Flag BETWEEN 100 AND 199 THEN A.Quantity ");
            //sb.Append("WHEN A.Flag BETWEEN 200 AND 299 THEN -A.Quantity ");
            //sb.Append("WHEN A.Flag BETWEEN 300 AND 399  AND A.Flag NOT IN (318, 319, 320)  THEN A.Quantity ");
            //sb.Append("WHEN A.Flag BETWEEN 500 AND 599 THEN -(A.Quantity - A.QuanComb) ");
            //sb.Append("WHEN A.Flag BETWEEN 600 AND 699 THEN A.Quantity ELSE 0 END end) AS 庫存數量 ");



            // sb.Append("Case WHEN T1.Flag/100 in (1, 5) THEN Sum(T1.Quantity) ELSE -Sum(T1.Quantity) END AS 數量,");



            sb.Append("Case When B.CtmWeight<>0 then Case WHEN T1.Flag/100 in (1, 5) THEN Sum(T1.Quantity * B.CtmWeight)  ELSE -Sum(T1.Quantity* B.CtmWeight) END ");
            sb.Append("ELSE Case WHEN T1.Flag/100 in (1, 5) THEN Sum(T1.Quantity) ELSE -Sum(T1.Quantity) END END  AS 數量,");





            sb.Append("Case WHEN T1.Flag/100 in (1, 5) THEN Sum(T1.Amount) ELSE -Sum(T1.Amount) END AS 金額 ");
            sb.Append(" from comProdRec T1 ");



            sb.Append("inner Join StkBillSub T4 on T4.Flag = T1.Flag and  T4.BillNo = T1.BillNo and T4.RowNo=T1.RowNo ");
            sb.Append("Inner Join ComProduct B on T1.ProdId=B.ProdID ");


            //500:銷,600銷退,100進,200進退,300調撥
            if (Flag == "進")
            {
                sb.Append("where T1.Flag in (100,200) ");
            }
            else
            {
                sb.Append("where T1.Flag in (500,600) ");
            }



            //  sb.Append("and (T1.ProdID like 'MP%' or T1.ProdID like 'MC%' or T1.ProdID like 'GMX%'  or T1.ProdID like 'GPX%' ) ");


            sb.Append("and T1.BillDate >= @StartDate ");
            sb.Append("and T1.BillDate <= @EndDate ");


            if (Category == "豬")
            {
                sb.Append("and B.EngName in (select distinct EngName from ComProduct where ClassId in ('AWP200','ARP200') and EngName<>'' ) ");
            }
            else if (Category == "雞")
            {
                //R.ClassId in ('AWC200','ARC200','AWP200','ARP200')
                sb.Append("and B.EngName in (select distinct EngName from ComProduct where ClassId in ('AWC200','ARC200')) ");
            }

            //依贈品不計算
            // sb.Append("group by B.EngName,T1.Flag,T4.IsGift ");
            sb.Append("group by B.EngName, T1.Flag,B.CtmWeight ");


            //if (ckFoc.Checked)
            //{
            //    sb.Append("group by T1.ProdID,T1.ProdName,T1.Flag,T4.IsGift ");
            //}
            //else
            //{
            //    sb.Append("group by T1.ProdID,T1.ProdName,T1.Flag ");
            //}


            //sb.Append("order by ProdID,ProdName,BillDate,Flag ");
            // sb.Append("order by T1.ProdID,T1.ProdName ");


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

        private void button93_Click(object sender, EventArgs e)
        {
            string CalType = comboBox1.SelectedItem.ToString();

            //MessageBox.Show(CalType);

            // return;

            DataTable dtData = MakeTableEngName();
            dataGridView1.DataSource = dtData;



            string StartDate = "20140101";
            string EndDate = txtEndDate.Text;


            DataTable dt = GetChiStockHis_EngName_Weight_All(StartDate, EndDate, CalType);


            StartDate = txtStartDate.Text;
            DataTable dtCost = GetChiStockIn_EngName(StartDate, EndDate, "銷", CalType);

            // dt.Columns.Add("排行", typeof(string));
            //dt.Columns.Add("部位", typeof(string));
            //dt.Columns.Add("庫存量", typeof(Int32));
            //dt.Columns.Add("庫存佔比", typeof(Int32));
            //dt.Columns.Add("週轉率", typeof(double));


            Double TotalKg = 0;
            DataRow dr;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtData.NewRow();

                dr["部位"] = Convert.ToString(dt.Rows[i]["EngName"]);

                if (Convert.ToString(dr["部位"]) == "朝貢雞")
                {
                    dr["部位"] = "朝貢雞(分解)";
                }


                dr["庫存量"] = Convert.ToInt32(dt.Rows[i]["庫存數量"]);
                dr["庫存金額"] = Convert.ToInt32(dt.Rows[i]["庫存金額"]);
                dr["銷貨成本"] = 0;
                dr["銷貨數量"] = 0;


                TotalKg = TotalKg + Convert.ToInt32(dt.Rows[i]["庫存數量"]);
                dtData.Rows.Add(dr);

            }



            for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
            {
                dr = dtData.Rows.Find(Convert.ToString(dtCost.Rows[i]["EngName"]));


                if (dr != null)
                {

                    dr.BeginEdit();
                    dr["銷貨成本"] = Convert.ToInt32(dr["銷貨成本"]) + Convert.ToInt32(dtCost.Rows[i]["成本"]);

                    dr["銷貨數量"] = Convert.ToInt32(dr["銷貨數量"]) + Convert.ToInt32(dtCost.Rows[i]["數量"]); ;

                    dr.EndEdit();
                }


            }


            for (int i = 0; i <= dtData.Rows.Count - 1; i++)
            {
                dr = dtData.Rows[i];

                dr.BeginEdit();

                double ratio = 0;

                ratio = Convert.ToInt32(dtData.Rows[i]["庫存量"]) / TotalKg * 100;

                dr["庫存佔比"] = ratio.ToString("#,##0.00");


                if (Convert.ToInt32(dtData.Rows[i]["銷貨成本"]) > 0)
                {

                    decimal turnover = 0;




                    //週轉率
                    try
                    {



                        // turnover  = Convert.ToInt32(dtData.Rows[i]["銷貨成本"]) / Convert.ToDecimal(dtData.Rows[i]["庫存金額"]) * 100;

                        turnover = Convert.ToInt32(dtData.Rows[i]["銷貨數量"]) / Convert.ToDecimal(dtData.Rows[i]["庫存量"]) * 100;
                        dr["週轉率"] = turnover.ToString("#,##0.00");

                    }
                    catch
                    {
                        // dr["週轉率"] = 0;
                    }
                }

                dr.EndEdit();

            }

            DataView dv = dtData.DefaultView;


            dv.Sort = "庫存量 desc";


            DataTable dtT = dv.ToTable();


            for (int i = 0; i <= dtT.Rows.Count - 1; i++)
            {
                dr = dtT.Rows[i];

                dr.BeginEdit();


                dr["排行"] = (i + 1).ToString();


                dr.EndEdit();

            }


            dataGridView1.DataSource = dtT;
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

        private void button86_Click(object sender, EventArgs e)
        {
            GridViewToCSV(dataGridView1, "產品資料.csv");
        }

        private void fmRivaRep01_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;

            txtStartDate.Text = DateTime.Now.AddDays(-30).ToString("yyyyMMdd");
            txtEndDate.Text = DateTime.Now.AddDays(0).ToString("yyyyMMdd");

 
        }
    }
}
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
    public partial class ReportForm : Form
    {
        private decimal sd;
        public ReportForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text + textBox3.Text == "" && INVOICE1.Text + INVOICE2.Text == "")
            {
                MessageBox.Show("請輸入查詢條件");
                return;
            }

            if ( INVOICE1.Text + INVOICE2.Text != "")
            {
                int num1;
                if (int.TryParse(INVOICE1.Text, out num1) == false || int.TryParse(INVOICE1.Text, out num1) == false) 
                {
                    MessageBox.Show("SO#請輸入數字");
                    return;
                }
            }
            System.Data.DataTable dt;
            System.Data.DataTable dt1 = null;
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTableCombine();
      
                dt = GetOrderDataAP();
         

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string JOBNO = dt.Rows[i]["JOBNO"].ToString();

                dt1 = GetOrderDataAP1(JOBNO);
                dr = dtCost.NewRow();
                dr["JOBNO"] = JOBNO;
                dr["客戶名稱"] = dt.Rows[i]["客戶名稱"].ToString();
                dr["收貨地"] = dt.Rows[i]["收貨地"].ToString();
                dr["目的地"] = dt.Rows[i]["目的地"].ToString();
                dr["幣別"] = dt.Rows[i]["幣別"].ToString();
                dr["目的地"] = dt.Rows[i]["目的地"].ToString();
                dr["幣別"] = dt.Rows[i]["幣別"].ToString();
                dr["貿易形式"] = dt.Rows[i]["貿易形式"].ToString();
                dr["報單號碼"] = dt.Rows[i]["報單號碼"].ToString();
                dr["LINK"] = dt.Rows[i]["LINK"].ToString();
                dr["LINKL"] = dt.Rows[i]["LINKL"].ToString();
                
                sd = 0;
                string DOCENTRY = dt1.Rows[0]["單據號碼"].ToString();
                System.Data.DataTable DT2 = GetDOCRATE(DOCENTRY);
                if (DT2.Rows.Count > 0)
                {
                    dr["匯率"] = DT2.Rows[0][0].ToString();
                
                }
                System.Data.DataTable DT3 = GetINVOICENO(JOBNO);
                if (DT3.Rows.Count > 0)
                {
                    dr["INVOICENO"] = DT3.Rows[0][0].ToString();
                    dr["金額"] = DT3.Rows[0][1].ToString();
                }
                for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                {
                 
                  
                    if (dt1.Rows.Count == 1)
                    {

                        dr["單據總類"] = "S0#" + dt1.Rows[j]["單據號碼"].ToString();

                  
                    }
                    else
                    {

                        if (j == dt1.Rows.Count - 1)
                        {

                            dr["單據總類"] += "S0#" + dt1.Rows[j]["單據號碼"].ToString();

                        }
                        else
                        {


                            dr["單據總類"] += "S0#" + dt1.Rows[j]["單據號碼"].ToString() + "/";
               
                        }
                    }
                }
              

                dtCost.Rows.Add(dr);
            }
            bindingSource1.DataSource = dtCost;
            dataGridView1.DataSource = bindingSource1.DataSource;
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox3.Text = GetMenu.DLast();
        }
        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT  t0.shippingcode JOBNO,cardname 客戶名稱,CASE MAX(ISNULL(T3.[PATH],'')) WHEN '' THEN '' ELSE  '附件下載' END LINK,MAX(T3.[PATH]) LINK2 ");
            sb.Append(" ,CASE MAX(ISNULL(T4.[PATH],'')) WHEN '' THEN '' ELSE  '附件下載' END LINKL,MAX(T4.[PATH]) LINKL2,T0.receivePlace 收貨地,T0.goalPlace 目的地");
            sb.Append(" ,T5.CURRENCY 幣別,T0.boardCountNo 貿易形式,T0.add9 報單號碼");
            sb.Append(" from shipping_main t0 ");
            sb.Append(" left join shipping_item t5 on (t0.shippingcode=t5.shippingcode)  ");
            sb.Append(" LEFT JOIN (SELECT SHIPPINGCODE,[PATH] FROM download WHERE [PATH] LIKE '%I&P%' OR REPLACE(REPLACE([PATH],'SHIPPING',''),'ZIP','')   LIKE '%IP%') T3 ON (T0.SHIPPINGCODE=T3.SHIPPINGCODE) ");
            sb.Append(" LEFT JOIN (SELECT [PATH],T0.SHIPPINGCODE FROM SHIPPING_MAIN T0 ");
            sb.Append(" LEFT JOIN download2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE  ) WHERE  ");
            sb.Append(" SUBSTRING([FILENAME],0,CHARINDEX('.', [FILENAME]))=T0.ADD9 COLLATE  Chinese_Taiwan_Stroke_CI_AS ");
            sb.Append(" ) T4 ON (T0.SHIPPINGCODE=T4.SHIPPINGCODE)");
            sb.Append(" where  ITEMREMARK='銷售訂單'   ");
            if (INVOICE1.Text != "" && INVOICE2.Text != "")
            {
                sb.Append(" and  T5.DOCENTRY between @INVOICENO1 and @INVOICENO2 ");
            }
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append("  AND SUBSTRING(t0.shippingcode,3,8) BETWEEN @aa and @bb ");
            }
            sb.Append(" group by t0.shippingcode ,cardname ,add9,T0.goalPlace,T5.CURRENCY,T0.receivePlace,T0.boardCountNo,T0.add9");
            sb.Append("  HAVING MAX(ISNULL(T3.[PATH],''))+MAX(ISNULL(T4.[PATH],'')) <> ''  "); 
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox3.Text));
            if (INVOICE1.Text != "" && INVOICE2.Text != "")
            {
                command.Parameters.Add(new SqlParameter("@INVOICENO1", Convert.ToInt32(INVOICE1.Text)));
                command.Parameters.Add(new SqlParameter("@INVOICENO2", Convert.ToInt32(INVOICE2.Text)));
            }

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

        private System.Data.DataTable GetOrderDataAPL(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                   select  t0.shippingcode JOBNO,cardname 客戶名稱,boardCountNo 貿易形式,add9 報單號碼,add7 所有人,sum(cast(t2.amount as int)) 金額,boatName 港名,CASE MAX(ISNULL(T3.[PATH],'')) WHEN '' THEN '' ELSE  '附件下載' END LINK,INVOICENO,MAX(T3.[PATH]) LINK2");
            sb.Append(" ,CASE MAX(ISNULL(T4.[PATH],'')) WHEN '' THEN '' ELSE  '附件下載' END LINKL,MAX(T4.[PATH]) LINKL2");
            sb.Append("                        from shipping_main t0");
            sb.Append("                       left join invoiced t2 on (t0.shippingcode=t2.shippingcode)");
            sb.Append(" LEFT JOIN (SELECT SHIPPINGCODE,[PATH] FROM download WHERE [PATH] LIKE '%I&P%' OR REPLACE(REPLACE([PATH],'SHIPPING',''),'ZIP','')  LIKE '%IP%') T3 ON (T0.SHIPPINGCODE=T3.SHIPPINGCODE)");
            sb.Append(" LEFT JOIN (SELECT [PATH],T0.SHIPPINGCODE FROM SHIPPING_MAIN T0");
            sb.Append(" LEFT JOIN download2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE  ) WHERE ");
            sb.Append(" SUBSTRING([FILENAME],0,CHARINDEX('.', [FILENAME]))=T0.ADD9 COLLATE  Chinese_Taiwan_Stroke_CI_AS");
            sb.Append(" ) T4 ON (T0.SHIPPINGCODE=T4.SHIPPINGCODE)");
            sb.Append("           where T0.SHIPPINGCODE=@SHIPPINGCODE ");
          
            sb.Append(" group by t0.shippingcode ,cardname ,boardCountNo ,add9 ,add7,boatName,INVOICENO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
        
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
        private System.Data.DataTable GetOrderDataAP1(string aa)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select  distinct t1.itemremark 單據總類,t1.DOCENTRY 單據號碼");
            sb.Append("               from shipping_main t0");
            sb.Append("              left join shipping_item t1 on (t0.shippingcode=t1.shippingcode)");
            sb.Append(" where  t0.shippingcode=@aa ");
            if (INVOICE1.Text != "" && INVOICE2.Text != "")
            {

                sb.Append(" and  T1.DOCENTRY between @INVOICENO1 and @INVOICENO2 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@INVOICENO1", INVOICE1.Text));
            command.Parameters.Add(new SqlParameter("@INVOICENO2", INVOICE2.Text));
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
        private System.Data.DataTable GetINVOICENO(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(InvoiceNo) INVOICENO,CAST(SUM(AMOUNT) AS INT) 金額 FROM InvoiceD WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        private System.Data.DataTable GetDOCRATE(string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT DocRate FROM ORDR WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("單據總類", typeof(string));
            dt.Columns.Add("收貨地", typeof(string));
            dt.Columns.Add("目的地", typeof(string));
            dt.Columns.Add("INVOICENO", typeof(string));
            dt.Columns.Add("幣別", typeof(string));
            dt.Columns.Add("匯率", typeof(decimal));
            dt.Columns.Add("金額", typeof(int));
            dt.Columns.Add("貿易形式", typeof(string));
            dt.Columns.Add("報單號碼", typeof(string));
            dt.Columns.Add("LINK", typeof(string));
            dt.Columns.Add("LINKL", typeof(string));

            return dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {
                    string sd = dataGridView1.CurrentRow.Cells["JOBNO"].Value.ToString();
                    System.Data.DataTable dt1 = GetOrderDataAPL(sd);
                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                        {
                            DataRow drw = dt1.Rows[i];
                            System.Diagnostics.Process.Start(drw["LINK2"].ToString());
                        }
                            DataGridViewLinkCell cell =

                                (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];
                        
                    }
                }
                if (dgv.Columns[e.ColumnIndex].Name == "LINKL")
                {
                    string sd = dataGridView1.CurrentRow.Cells["JOBNO"].Value.ToString();
                    System.Data.DataTable dt1 = GetOrderDataAPL(sd);

                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                        {
                            DataRow drw = dt1.Rows[i];

                            System.Diagnostics.Process.Start(drw["LINKL2"].ToString());

                        }
                            DataGridViewLinkCell cell =

                                (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];
                        
                    }
                }
        }
 
    }
}
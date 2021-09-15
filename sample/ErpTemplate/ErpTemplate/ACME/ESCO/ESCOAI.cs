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
    public partial class ESCOAI : Form
    {
        public ESCOAI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string DD = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + "01";
            string DD2 = Convert.ToDateTime(DD).AddMonths(1).ToString("yyyy") + "/" + Convert.ToDateTime(DD).AddMonths(1).ToString("MM");
            System.Data.DataTable dtCost = MakeTable();
            System.Data.DataTable dt = GetProject2();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                string 代碼 = dt.Rows[i]["PRJCODE"].ToString();

                dr["代碼"] = 代碼;
                dr["名稱"] =  dt.Rows[i]["PRJNAME"].ToString();
             
                System.Data.DataTable dt2 = GetProject(代碼);
                if (dt2.Rows.Count > 0)
                {
               
                    dr["金額"] = Convert.ToInt32(dt2.Rows[0]["金額"]);
                }
                System.Data.DataTable dt3 = GetProject3(代碼);
                if (dt3.Rows.Count > 0)
                {
                    dr["AI輸入日期"] = DD2 +"/"+ dt3.Rows[0][0].ToString();
                    dr["AWS編號"] = dt3.Rows[0]["AWS編號"].ToString();
                    dr["訂單號碼"] = dt3.Rows[0]["訂單號碼"].ToString();
                 
                }
                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("代碼", typeof(string));
            dt.Columns.Add("名稱", typeof(string));
            dt.Columns.Add("AWS編號", typeof(string));
            dt.Columns.Add("訂單號碼", typeof(string));
            dt.Columns.Add("金額", typeof(int));
            dt.Columns.Add("AI輸入日期", typeof(string));
            return dt;
        }
        public DataTable GetProject2()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT distinct O.PrjCode PRJCODE,O.PrjName PRJNAME,U_MEMO MEMO,U_MEMO2 MEMO2  FROM OPRJ O    WHERE  Substring(O.PrjCode,1,1)='4' AND ISNULL(U_MEMO,'') NOT　IN ('專案碼刪除','撤案') ORDER BY PrjCode   ");

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
            return ds.Tables[0];
        }
        public DataTable GetProject(string PROJECT)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,T1.PROJECT 專案,T0.DOCENTRY 訂單號碼,CAST(PriceAfVAT AS INT) 金額,U_CUSTDOCENTRY 輸入日期,T0.U_ACME_Warranty AWS編號    FROM ORDR T0");
            sb.Append(" LEFT JOIN RDR1  T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("  WHERE SUBSTRING(U_CUSTDOCENTRY,1,6)=@U_CUSTDOCENTRY AND T1.PROJECT=@PROJECT ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_CUSTDOCENTRY", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
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
            return ds.Tables[0];
        }
        public DataTable GetProject3(string PROJECT)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE WHEN LEN(T0.U_LOCATION)=1 THEN '0'+T0.U_LOCATION ELSE T0.U_LOCATION  END BDATE,T0.U_ACME_Warranty AWS編號,T0.DOCENTRY 訂單號碼     FROM ORDR T0 ");
            sb.Append(" LEFT JOIN RDR1  T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" WHERE ISNULL(T0.U_LOCATION,'') <> '' AND T1.PROJECT=@PROJECT ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
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
            return ds.Tables[0];
        }

        private void ESCOAI_Load(object sender, EventArgs e)
        {
            textBox1.Text= DateTime.Now.AddMonths(-1).ToString("yyyy") + DateTime.Now.AddMonths(-1).ToString("MM");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

   
    }
}

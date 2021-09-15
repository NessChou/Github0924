using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class SUPSum : Form
    {
        public SUPSum()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTable(int EndMon)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            for (int i = 1; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString(), typeof(string));
            }



            return dt;
        }
        private System.Data.DataTable MakeTable2(int YEAR)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add(((YEAR - 2) + "數量").ToString(), typeof(string));
            dt.Columns.Add(((YEAR - 2) + "未稅金額").ToString(), typeof(string));
            dt.Columns.Add(((YEAR - 1) + "數量").ToString(), typeof(string));
            dt.Columns.Add(((YEAR - 1) + "未稅金額").ToString(), typeof(string));
            dt.Columns.Add(((YEAR) + "數量").ToString(), typeof(string));
            dt.Columns.Add(((YEAR) + "未稅金額").ToString(), typeof(string));
            


            return dt;
        }
             //SELECT CARDCODE 客戶編號,CARDNAME 客戶名稱,CAST(SUM(金額) AS INT) 金額 FROM ( SELECT CARDCODE,CARDNAME,SUM(U_IN_BSAMN) 金額 FROM OINV T0  
             // WHERE  T0.[U_IN_BSAPP] = '2020/12/15'
             // GROUP BY CARDCODE,CARDNAME 
             // UNION ALL
             //               SELECT CARDCODE,CARDNAME,SUM(T0.U_BSAMN2)*-1
             //  FROM [@CADMEN_CMD1] T0 
             //  left join [@CADMEN_CMD]  T1 on T0.DOCENTRY=T1.DOCENTRY 
             //  left join [ORIN]  T2 on T1.U_BSREN=T2.DOCENTRY 
             // WHERE  T2.[U_RI_BSAPP]  = '2020/12/15'
             //     GROUP BY CARDCODE,CARDNAME ) AS A
             //     GROUP BY CARDCODE,CARDNAME
        private void PROD()
        {

            string YEAR = comboBox1.Text;
            int EMONTH = 12;
            if (comboBox1.Text == DateTime.Now.Year.ToString())
            {
                EMONTH = DateTime.Now.Month;
            }
            System.Data.DataTable dt = MakeTable(EMONTH);
            System.Data.DataTable dt2 = MakeTable(EMONTH);


            System.Data.DataTable dtSIZE = GETAMT("1", YEAR, 0, "");
            DataRow dr;
            DataRow dr2;

            for (int l = 0; l <= dtSIZE.Rows.Count - 1; l++)
            {
                DataRow dz = dtSIZE.Rows[l];


                dr = dt.NewRow();
                dr2 = dt2.NewRow();
            
                string 廠商編號 = dz["廠商編號"].ToString();
                string 廠商名稱 = dz["廠商名稱"].ToString();
                string 付款條件 = dz["付款條件"].ToString();
                string 數量 = dz["數量"].ToString();
                string 未稅金額 = dz["未稅金額"].ToString();
                dr["廠商編號"] = 廠商編號;
                dr["廠商名稱"] = 廠商名稱;
                dr["付款條件"] = 付款條件;

                dr2["廠商編號"] = 廠商編號;
                dr2["廠商名稱"] = 廠商名稱;
                dr2["付款條件"] = 付款條件;


                for (int M = 1; M <= EMONTH; M++)
                {
                    System.Data.DataTable dh = null;
                    dh = GETAMT("2", YEAR, M, 廠商編號);
                    string DHV = "0";
                    string DHV2 = "0";
                    if (dh.Rows.Count > 0)
                    {
                        DHV = dh.Rows[0][3].ToString();
                        DHV2 = dh.Rows[0][4].ToString();
                    }

                    dr[M.ToString()] = Convert.ToDecimal(DHV).ToString("#,##0");
                    dr2[M.ToString()] = Convert.ToDecimal(DHV2).ToString("#,##0");
                }
                dt.Rows.Add(dr);
                dt2.Rows.Add(dr2);

            }
            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = dt2;
       
            //for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            //{
            //    DataGridViewColumn c = dataGridView1.Columns[i];
            //    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    c.DefaultCellStyle.Format = "#,##0";

            //}
        }

        private void PROD2()
        {
            dataGridView3.DataSource = null;
            string YEAR = comboBox1.Text;
            int YEAR2 = Convert.ToInt16(comboBox1.Text);

            System.Data.DataTable dt3 = MakeTable2(YEAR2);

            System.Data.DataTable dtSIZE = GETAMT2(YEAR2 - 2, YEAR2);

            DataRow dr3;
            for (int l = 0; l <= dtSIZE.Rows.Count - 1; l++)
            {
                DataRow dz = dtSIZE.Rows[l];
             
  
   
                    dr3 = dt3.NewRow();
                    string 廠商編號 = dz["廠商編號"].ToString();
                    string 廠商名稱 = dz["廠商名稱"].ToString();
                    string 付款條件 = dz["付款條件"].ToString();

                    dr3["廠商編號"] = 廠商編號;
                    dr3["廠商名稱"] = 廠商名稱;
                    dr3["付款條件"] = 付款條件;
                    System.Data.DataTable dtSIZE1 = GETAMT("3", (YEAR2).ToString(), 0, 廠商編號);
                    System.Data.DataTable dtSIZE2 = GETAMT("3", (YEAR2 - 2).ToString(), 0, 廠商編號);
                    System.Data.DataTable dtSIZE3 = GETAMT("3", (YEAR2 - 1).ToString(), 0, 廠商編號);
                    if (dtSIZE2.Rows.Count > 0)
                    {
                        dr3[((YEAR2 - 2) + "數量")] = Convert.ToDecimal(dtSIZE2.Rows[0]["數量"]).ToString("#,##0");
                        dr3[((YEAR2 - 2) + "未稅金額")] = Convert.ToDecimal(dtSIZE2.Rows[0]["未稅金額"]).ToString("#,##0");
                    }
                    if (dtSIZE3.Rows.Count > 0)
                    {
                        dr3[((YEAR2 - 1) + "數量")] = Convert.ToDecimal(dtSIZE3.Rows[0]["數量"]).ToString("#,##0");
                        dr3[((YEAR2 - 1) + "未稅金額")] = Convert.ToDecimal(dtSIZE3.Rows[0]["未稅金額"]).ToString("#,##0");
                    }
                    if (dtSIZE1.Rows.Count > 0)
                    {
                        dr3[((YEAR2) + "數量")] = Convert.ToDecimal(dtSIZE1.Rows[0]["數量"]).ToString("#,##0");
                        dr3[((YEAR2) + "未稅金額")] = Convert.ToDecimal(dtSIZE1.Rows[0]["未稅金額"]).ToString("#,##0");
                    }
           
            
        
                    dt3.Rows.Add(dr3);
            }
  
            dataGridView3.DataSource = dt3;
            //for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            //{
            //    DataGridViewColumn c = dataGridView1.Columns[i];
            //    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    c.DefaultCellStyle.Format = "#,##0";

            //}
        }


        System.Data.DataTable GETAMT(string DOCTYPE, string YEAR, int MONTH, string CARDCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            // LIKE  '%" + NAME + "%' 
            sb.Append("	SELECT 廠商編號,廠商名稱,MAX(付款條件) 付款條件,SUM(數量) 數量,SUM(未稅金額) 未稅金額 FROM (");
            sb.Append("	SELECT T0.CARDCODE 廠商編號,T2.CARDNAME 廠商名稱,MAX(T4.PymntGroup) 付款條件 , SUM(CAST(T1.QUANTITY AS INT)) 數量,");
            sb.Append("	SUM(CAST(T1.LINETOTAL AS float)) 未稅金額  FROM OPCH  T0");
            sb.Append("	INNER JOIN PCH1 T1 ON T0.DOCENTRY = T1.DOCENTRY ");
            sb.Append("	INNER JOIN OCRD T2 ON T0.CARDCODE = T2.CARDCODE ");
            sb.Append("	left join ACMESQL02.DBO.octg t4 on(t2.groupnum=t4.groupnum)");
            sb.Append("	WHERE  SUBSTRING(T0.CARDCODE,1,1)='S' AND YEAR(T0.DOCDATE)=@YEAR ");
            if (DOCTYPE == "2")
            {
                sb.Append(" AND MONTH(T0.DOCDATE)=@MONTH AND T0.CARDCODE=@CARDCODE ");
            }
            if (DOCTYPE == "3")
            {
                sb.Append(" AND T0.CARDCODE=@CARDCODE ");
            }
            if (textBox1.Text != "")
            {
                sb.Append(" AND T2.CARDNAME  LIKE '%" + textBox1.Text + "%' ");
            }
            sb.Append("	GROUP BY T0.CARDCODE ,T2.CARDNAME");
            sb.Append("	UNION ALL");
            sb.Append("	SELECT T0.CARDCODE 廠商編號,T2.CARDNAME 廠商名稱,MAX(T4.PymntGroup) 付款條件 , SUM(CAST(T1.QUANTITY AS INT))*-1 數量,");
            sb.Append("	SUM(CAST(T1.LINETOTAL AS float))*-1 未稅金額  FROM ORPC  T0");
            sb.Append("	INNER JOIN RPC1 T1 ON T0.DOCENTRY = T1.DOCENTRY ");
            sb.Append("	INNER JOIN OCRD T2 ON T0.CARDCODE = T2.CARDCODE ");
            sb.Append("	left join ACMESQL02.DBO.octg t4 on(t2.groupnum=t4.groupnum)");
            sb.Append("	WHERE  SUBSTRING(T0.CARDCODE,1,1)='S' AND YEAR(T0.DOCDATE)=@YEAR ");

            if (DOCTYPE == "2")
            {
                sb.Append(" AND MONTH(T0.DOCDATE)=@MONTH AND T0.CARDCODE=@CARDCODE ");

            }
            if (DOCTYPE == "3")
            {
                sb.Append(" AND T0.CARDCODE=@CARDCODE ");
            }
            if (textBox1.Text != "")
            {
                sb.Append(" AND T2.CARDNAME LIKE  '%" + textBox1.Text + "%' ");
            }
            sb.Append("	GROUP BY T0.CARDCODE ,T2.CARDNAME) AS A GROUP BY 廠商編號,廠商名稱");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH ", MONTH));
            command.Parameters.Add(new SqlParameter("@CARDCODE ", CARDCODE));
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

        System.Data.DataTable GETAMT2(int YEAR, int YEAR2)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            // LIKE  '%" + NAME + "%' 
            sb.Append("	SELECT DISTINCT 廠商編號,廠商名稱,付款條件 FROM (");
            sb.Append("	SELECT DISTINCT T0.CARDCODE 廠商編號,T2.CARDNAME 廠商名稱,(T4.PymntGroup) 付款條件   FROM OPCH  T0 ");
            sb.Append("	INNER JOIN OCRD T2 ON T0.CARDCODE = T2.CARDCODE ");
            sb.Append("	left join ACMESQL02.DBO.octg t4 on(t2.groupnum=t4.groupnum)");
            sb.Append("	WHERE  SUBSTRING(T0.CARDCODE,1,1)='S' AND YEAR(T0.DOCDATE) BETWEEN @YEAR AND @YEAR2 ");
            if (textBox1.Text != "")
            {
                sb.Append(" AND T0.CARDNAME  LIKE '%" + textBox1.Text + "%' ");
            }

            sb.Append("	UNION ALL");
            sb.Append("	SELECT DISTINCT T0.CARDCODE 廠商編號,T2.CARDNAME 廠商名稱,(T4.PymntGroup) 付款條件  FROM ORPC  T0 ");
            sb.Append("	INNER JOIN OCRD T2 ON T0.CARDCODE = T2.CARDCODE ");
            sb.Append("	left join ACMESQL02.DBO.octg t4 on(t2.groupnum=t4.groupnum)");
            sb.Append("	WHERE  SUBSTRING(T0.CARDCODE,1,1)='S' AND YEAR(T0.DOCDATE) BETWEEN @YEAR AND @YEAR2 ");


            if (textBox1.Text != "")
            {
                sb.Append(" AND T0.CARDNAME LIKE  '%" + textBox1.Text + "%' ");
            }

            sb.Append("	) AS A");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@YEAR2", YEAR2));

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

        private void ProductSum_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year2017(), "DataValue", "DataValue");
            PROD();
            PROD2();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            PROD();
            PROD2();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
       
        }
    }
}

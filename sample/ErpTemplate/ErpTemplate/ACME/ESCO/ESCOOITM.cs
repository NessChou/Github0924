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
    public partial class ESCOOITM : Form
    {
        public ESCOOITM()
        {
            InitializeComponent();
        }

        public DataTable GetExpense(string WHSCODE)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE 設備編碼,T1.DSCRIPTION 設備名稱,T1.PROJECT 專案代碼,");
            sb.Append(" SUBSTRING(T2.PRJNAME,1,4) 客戶地區,T0.DOCENTRY 收採號碼,Convert(varchar(10),  T0.DOCDATE,111) 收採日期");
            sb.Append(" ,T1.QUANTITY 數量,T1.LINETOTAL 總成本,T3.USERTEXT  ");
            sb.Append(" FROM OPDN T0");
            sb.Append(" LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OPRJ T2 ON (T1.PROJECT=T2.PRJCODE)");
            sb.Append(" LEFT JOIN OITM T3 ON (T1.ITEMCODE=T3.ITEMCODE)");
            sb.Append(" WHERE  T3.itmsgrpcod='101'   AND T1.WHSCODE=@WHSCODE AND T0.DOCENTRY NOT IN (19496,19459,19460,19498,19499)  order by T1.ITEMCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("項目", typeof(string));
            dt.Columns.Add("設備編碼", typeof(string));
            dt.Columns.Add("設備名稱", typeof(string));
            dt.Columns.Add("專案代碼", typeof(string));
            dt.Columns.Add("客戶地區", typeof(string));
            dt.Columns.Add("收採號碼", typeof(string));
            dt.Columns.Add("收採日期", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("總成本", typeof(decimal));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("安裝地址", typeof(string));
            dt.Columns.Add("聯絡人", typeof(string));
            dt.Columns.Add("電話", typeof(string));
            return dt;
        }
        private void ESCOOITM_Load(object sender, EventArgs e)
        {
            for (int s = 0; s <= 1; s++)
            {
                string WHS = "";
                if (s == 0)
                {
                    WHS = "FA001";
                }
                if (s == 1)
                {
                    WHS = "Z0015";
                }
                DataTable dt = GetExpense(WHS);
                DataRow dr = null;
                System.Data.DataTable dtCost = MakeTable();
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();

                    dr["項目"] = (i + 1).ToString();
                    dr["設備編碼"] = dt.Rows[i]["設備編碼"].ToString();
                    dr["設備名稱"] = dt.Rows[i]["設備名稱"].ToString();
                    dr["專案代碼"] = dt.Rows[i]["專案代碼"].ToString();


                    dr["客戶地區"] = dt.Rows[i]["客戶地區"].ToString();
                    dr["收採號碼"] = dt.Rows[i]["收採號碼"].ToString();
                    dr["收採日期"] = dt.Rows[i]["收採日期"].ToString();
                    dr["數量"] = dt.Rows[i]["數量"].ToString();

                    dr["總成本"] = dt.Rows[i]["總成本"].ToString();
                    string USERTEXT = dt.Rows[i]["USERTEXT"].ToString();
                    int F1 = USERTEXT.Length;
                    int G1 = USERTEXT.IndexOf("客戶名稱");
                    if (G1 != -1)
                    {
                        string aa = USERTEXT.Substring(G1 + 4, F1 - G1 - 4);
                        int gt = aa.IndexOf("\r");
                        int gt2 = aa.IndexOf("。");
                        string FF1 = "";
                        if (gt != -1)
                        {
                            FF1 = aa.Substring(1, gt - 1);
                        }
                        else if (gt2 != -1)
                        {
                            FF1 = aa.Substring(1, gt2 - 1);
                        }
                        else
                        {
                            FF1 = aa;
                        }
                        dr["客戶名稱"] = FF1.Replace("。", "");
                    }

                    int G2 = USERTEXT.IndexOf("安裝地址");
                    if (G2 != -1)
                    {
                        string aa = USERTEXT.Substring(G2 + 4, F1 - G2 - 4);
                        int gt = aa.IndexOf("\r");
                        int gt2 = aa.IndexOf("。");
                        string FF1 = "";
                        if (gt != -1)
                        {
                            FF1 = aa.Substring(1, gt - 1);
                        }
                        else if (gt2 != -1)
                        {
                            FF1 = aa.Substring(1, gt2 - 1);
                        }
                        else
                        {
                            FF1 = aa;
                        }
                        dr["安裝地址"] = FF1.Replace("。", "");
                    }
                    else
                    {
                        int G2S = USERTEXT.IndexOf("地址");
                        if (G2S != -1)
                        {
                            string aa = USERTEXT.Substring(G2S + 2, F1 - G2S - 2);
                            int gt = aa.IndexOf("\r");
                            int gt2 = aa.IndexOf("。");
                            string FF1 = "";
                            if (gt != -1)
                            {
                                FF1 = aa.Substring(1, gt - 1);
                            }
                            else if (gt2 != -1)
                            {
                                FF1 = aa.Substring(1, gt2 - 1);
                            }
                            else
                            {
                                FF1 = aa;
                            }
                            dr["安裝地址"] = FF1.Replace("。", "");
                        }
                    }

                    int G3 = USERTEXT.IndexOf("聯絡人");
                    if (G3 != -1)
                    {
                        string aa = USERTEXT.Substring(G3 + 3, F1 - G3 - 3);
                        int gt = aa.IndexOf("\r");
                        int gt2 = aa.IndexOf("。");
                        string FF1 = "";
                        if (gt != -1)
                        {
                            FF1 = aa.Substring(1, gt - 1);
                        }
                        else if (gt2 != -1)
                        {
                            FF1 = aa.Substring(1, gt2 - 1);
                        }
                        else
                        {
                            FF1 = aa;
                        }
                        dr["聯絡人"] = FF1.Replace("。", "");
                    }

                    int G4 = USERTEXT.IndexOf("電話");

                    if (G4 != -1)
                    {
                        string aa = USERTEXT.Substring(G4 + 2, F1 - G4 - 2);
                        int gt = aa.IndexOf("\r");
                        int gt2 = aa.IndexOf("。");
                        string FF1 = "";
                        if (gt != -1)
                        {
                            FF1 = aa.Substring(1, gt - 1);
                        }
                        else if (gt2 != -1)
                        {
                            FF1 = aa.Substring(1, gt2 - 1);
                        }
                        else
                        {
                            FF1 = aa;
                        }
                        dr["電話"] = FF1.Replace("。", "");
                    }
                    dtCost.Rows.Add(dr);
                }



                if (s == 0)
                {
                    dataGridView1.DataSource = dtCost;
                }
                if (s == 1)
                {
                    dataGridView2.DataSource = dtCost;
                }
              
            }
            for (int i = 7; i <= 8; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }
    }
}

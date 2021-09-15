using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace ACME
{
    public partial class POTATOCHECK : Form
    {
        public POTATOCHECK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("請選擇訂購人公司");
                return;
            }

            System.Data.DataTable DT1 = DT();

            if (DT1.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\請款對帳單.xls";

                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private System.Data.DataTable DT()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                                SELECT ORDCOM 客戶名稱,T0.DOCID 訂單單號,T0.ORDNO 運輸單號,T1.QTY 箱,");
            sb.Append("                                SPERSON 聯絡人,SADDRESS 地址,STEL 電話,ISNULL(T3.UNIT,'') 統一編號,");
            sb.Append("                                SUBSTRING(SDATE,1,4)+'/'+SUBSTRING(SDATE,5,2)+'/'+SUBSTRING(SDATE,7,2) ");
            sb.Append("                                +' 單號:'+CAST(DOCID AS VARCHAR) 到貨日期,T2.ITEMNAME,T1.PRICE 單價,T1.AMOUNT 小計  FROM GB_FRIEND T0");
            sb.Append("                                LEFT JOIN GB_POTATO2 T1 ON (T0.DOCID=T1.ID)");
            sb.Append("                                LEFT JOIN GB_OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
            sb.Append("                                LEFT JOIN GB_POTATO T3 ON (T0.DOCID=T3.ID)");
            sb.Append("                                 WHERE  T2.BIG='True' AND T0.DelRemark between @CreateDate and @CreateDate2 ");
            if (textBox3.Text != "")
            {
                sb.Append("                    AND ORDCOM=@COMPANY ");
            }
            sb.Append(" ORDER BY DOCID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", textBox3.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT2()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                     SELECT T0.DOCID,T0.ORDNO 運輸單號,T1.QTY 箱,'訂單號碼: '+CAST(T0.DOCID AS VARCHAR) 訂單號碼,");
            sb.Append("                     '收貨人: '+SPERSON 收貨人,'收貨地址: '+SADDRESS 收貨地址,'聯繫電話: '+STEL 聯繫電話,");
            sb.Append("                     SUBSTRING(SDATE,1,4)+'/'+SUBSTRING(SDATE,5,2)+'/'+SUBSTRING(SDATE,7,2) ");
            sb.Append("                     +'  '+ISNULL(STIME,'') 到貨日期, '訂單日期: '+SUBSTRING(CREATEDATE,1,4)+'/'+SUBSTRING(CREATEDATE,5,2)+'/'+SUBSTRING(CREATEDATE,7,2) 訂單日期,");
            sb.Append(" T2.ITEMNAME 品名,RANK() OVER (ORDER BY T2.ITEMNAME DESC) AS [NO] FROM GB_FRIEND T0");
            sb.Append("                               LEFT JOIN GB_POTATO2 T1 ON (T0.DOCID=T1.ID)");
            sb.Append("                   LEFT JOIN GB_OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
            sb.Append("                LEFT JOIN GB_POTATO T3 ON (T0.DOCID=T3.ID)");
            sb.Append("  WHERE T0.DOCID=@DOCID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCID", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void POTATOCHECK_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetGBOCRD();

            if (LookupValues != null)
            {
                textBox3.Text = Convert.ToString(LookupValues[2]);
          


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable DT1 = DT();
            dataGridView1.DataSource = DT1;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                MessageBox.Show("請輸入訂單號碼");
                return;
            }

            System.Data.DataTable DT1 = DT2();

            if (DT1.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\出貨簽收單.xls";

                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

    }
}

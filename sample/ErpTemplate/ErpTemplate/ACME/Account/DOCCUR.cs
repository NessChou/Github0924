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
    public partial class DOCCUR : Form
    {
        public DOCCUR()
        {
            InitializeComponent();
        }


        private System.Data.DataTable DTOPOR(string ITEMCODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE,T1.QUANTITY 數量,CAST(T2.LINETOTAL/T1.TOTALFRGN AS DECIMAL(10,4)) 匯率 FROM OPOR T0");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN PDN1 T2 ON (T1.docentry=T2.baseentry AND T1.linenum=T2.baseline)");
            sb.Append("  WHERE T1.ITEMCODE=@ITEMCODE  AND T1.TOTALFRGN <> 0 AND t2.basetype='22'  ");
            sb.Append(" ORDER BY T0.DOCDATE DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button1_Click(object sender, EventArgs e)
        {

        

            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt = OINV();
           
            DataRow dr = null;

           
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["交貨單date"] = dt.Rows[i]["交貨單date"].ToString();
                dr["交貨單#"] = dt.Rows[i]["交貨單#"].ToString();
                dr["銷售訂單#"] = dt.Rows[i]["銷售訂單#"].ToString();
                dr["AR發票"] = dt.Rows[i]["AR發票"].ToString();
                dr["客戶群組"] = dt.Rows[i]["客戶群組"].ToString();
                dr["客戶碼"] = dt.Rows[i]["客戶碼"].ToString();
                dr["客戶名稱"] = dt.Rows[i]["客戶名稱"].ToString();
                dr["幣別"] = dt.Rows[i]["幣別"].ToString();
                decimal 未稅原始金額 = Convert.ToDecimal(dt.Rows[i]["未稅原始金額"].ToString());
                dr["未稅原始金額"] = 未稅原始金額.ToString("#,##0");
                decimal 匯率 = Convert.ToDecimal(dt.Rows[i]["匯率"].ToString());
                dr["匯率"] = 匯率.ToString("#,##0.0000");
                decimal 帳載金額 = Convert.ToDecimal(dt.Rows[i]["帳載金額"].ToString());
                dr["帳載金額"] = 帳載金額.ToString("#,##0");
                decimal 銷貨成本 = Convert.ToDecimal(dt.Rows[i]["銷貨成本"].ToString());
                dr["銷貨成本"] = 銷貨成本.ToString("#,##0.0000");
                decimal 銷售訂單匯率 = Convert.ToDecimal(dt.Rows[i]["訂單匯率"].ToString());
                dr["銷售訂單匯率"] = 銷售訂單匯率.ToString("#,##0.0000");
                string ITEMCODE = dt.Rows[i]["品名"].ToString();
                dr["品名"] = ITEMCODE;
                decimal QTY = Convert.ToDecimal(dt.Rows[i]["數量"].ToString());
                dr["數量"] = QTY.ToString("#,##0");
                DataTable dt1 = DTOPOR(ITEMCODE);
                decimal QUANTITY = 0;
                decimal FINAL = 0;
                decimal DD = 0;
                for (int j = 0; j <= dt1.Rows.Count - 1; i++)
                {
                    decimal aa = Convert.ToDecimal(dt1.Rows[j]["匯率"].ToString());
                    decimal bb = Convert.ToDecimal(dt1.Rows[j]["數量"].ToString());
              

                    QUANTITY += bb;
                    decimal F1 = QTY - QUANTITY;
                    if (F1 < 0)
                    {
                        DD = QTY - (QUANTITY - bb);
                    }
                    if (F1 >= 0)
                    {
                        FINAL += aa * bb;
                    }
                    else
                    {
                        FINAL += aa * (DD);
                        break;
                    }
                }

                decimal 採購匯率 = FINAL / QTY;
                dr["採購匯率"] = (採購匯率).ToString("#,##0.0000");
                decimal 匯兌損益 = (匯率 - 採購匯率) * 未稅原始金額;
                dr["匯兌損益"] = (匯兌損益).ToString("#,##0.0000");

                decimal 匯兌損益2 = (匯率 - 銷售訂單匯率) * 未稅原始金額;
                dr["匯兌損益2"] = (匯兌損益2).ToString("#,##0.0000");
                dtCost.Rows.Add(dr);


            }
            dataGridView1.DataSource = dtCost;

            for (int i = 8; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

        }
        private System.Data.DataTable OINV()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  Convert(varchar(8),T7.docdate,112) 交貨單date,T7.DOCENTRY 交貨單#,T8.DOCENTRY 銷售訂單#,T0.DOCENTRY AR發票");
            sb.Append(" ,SUBSTRING(T11.GROUPNAME,4,5) 客戶群組,T0.CARDCODE 客戶碼,T0.CARDNAME 客戶名稱,T9.DOCCUR 幣別,T8.TOTALFRGN 未稅原始金額,T1.LINETOTAL 帳載金額");
            sb.Append(" ,CAST(T1.LINETOTAL/T8.TOTALFRGN AS DECIMAL(10,4)) 匯率,T1.GrossBuyPr 銷貨成本,T1.QUANTITY 數量,T1.PRICE 單價,T1.ITEMCODE 品名,ISNULL(T8.RATE,0) 訂單匯率 FROM OINV T0  ");
            sb.Append(" LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append(" LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append(" LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append(" LEFT JOIN OCRD T10 ON (T0.CARDCODE=T10.CARDCODE)");
            sb.Append(" LEFT JOIN OCRG T11 ON (T10.GROUPCODE=T11.GROUPCODE)");
            sb.Append(" where t1.basetype='15' AND T8.GtotalFC <> 0 AND T9.DOCCUR='USD' ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append("   and Convert(varchar(8),T7.docdate,112) between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'"); 
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append("   and T7.DOCENTRY between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'");
            }
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                sb.Append("   and T9.DOCENTRY between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "'");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
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
            dt.Columns.Add("交貨單date", typeof(string));
            dt.Columns.Add("交貨單#", typeof(string));
            dt.Columns.Add("銷售訂單#", typeof(string));
            dt.Columns.Add("AR發票", typeof(string));
            dt.Columns.Add("客戶群組", typeof(string));
            dt.Columns.Add("客戶碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("幣別", typeof(string));
            dt.Columns.Add("未稅原始金額", typeof(string));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("帳載金額", typeof(string));
            dt.Columns.Add("銷貨成本", typeof(string));
            dt.Columns.Add("採購匯率", typeof(string));
            dt.Columns.Add("匯兌損益", typeof(string));
            dt.Columns.Add("銷售訂單匯率", typeof(string));
            dt.Columns.Add("匯兌損益2", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            return dt;
        }

        private void DOCCUR_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}
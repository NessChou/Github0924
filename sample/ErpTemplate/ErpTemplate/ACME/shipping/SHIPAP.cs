using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class SHIPAP : Form 
    {
        public string cardcode;
        public string CLOSE;
        public string DOCTYPE;
        public string a;
        public SHIPAP()
        {
            InitializeComponent();
        }
        private void S1()
        {
            string NAME = globals.DBNAME;
            if (NAME == "進金生" || NAME == "測試98")
            {
                if (DOCTYPE == "採購報價")
                {

                    ViewBatchPayment2Q(cardcode);
                }
                else
                {

                    ViewBatchPayment2(cardcode);
                }
            }
            else
            {
                ViewBatchPaymentDRS(cardcode);
            }
        }
        private void AP_Load(object sender, EventArgs e)
        {

            S1();                  
        }
        private void ViewBatchPayment2(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select T0.Docentry 單號,T0.LINENUM 欄號,CASE ISNULL(T0.itemcode,'') WHEN '' THEN T0.DSCRIPTION ELSE T0.itemcode END 品名,T0.QUANTITY 數量,T2.QUANTITY 已沖數量,CAST(T0.QUANTITY-ISNULL(T2.QUANTITY,0) AS INT) 未沖數量 from POR1 T0 ");
            sb.Append("                 left join OPOR T1 on (T0.docentry=T1.docentry) INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T0.ItemCode  ");
            sb.Append("               left join  (SELECT DOCENTRY,LINENUM,SUM(QUANTITY) QUANTITY,ITEMCODE FROM ACMESQLSP.DBO.SHIPPING_ITEM WHERE ITEMREMARK='採購訂單' ");
            sb.Append("               GROUP BY DOCENTRY,LINENUM,ITEMCODE) T2 ON (CAST(T0.DOCENTRY AS VARCHAR)=CAST(T2.DOCENTRY AS VARCHAR)  AND T0.LINENUM=T2.LINENUM AND T0.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("  where T1.cardcode=@cardcode ");
            if (!checkBox1.Checked)
            {
                sb.Append("  AND T0.OPENCREQTY > 0 AND  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND CAST(T0.QUANTITY-ISNULL(T2.QUANTITY,0) AS INT) <> 0   ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("  AND T0.Docentry = @Docentry ");
            }

            sb.Append(" order by T0.Docentry ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment2Q(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select T0.Docentry 單號,T0.LINENUM 欄號,CASE ISNULL(T0.itemcode,'') WHEN '' THEN T0.DSCRIPTION ELSE T0.itemcode END 品名,T0.QUANTITY 數量,T2.QUANTITY 已沖數量,CAST(T0.QUANTITY-ISNULL(T2.QUANTITY,0) AS INT) 未沖數量 from PQT1 T0 ");
            sb.Append("                 left join OPQT T1 on (T0.docentry=T1.docentry) INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T0.ItemCode  ");
            sb.Append("               left join  (SELECT DOCENTRY,LINENUM,SUM(QUANTITY) QUANTITY,ITEMCODE FROM ACMESQLSP.DBO.SHIPPING_ITEM WHERE ITEMREMARK='採購報價' ");
            sb.Append("               GROUP BY DOCENTRY,LINENUM,ITEMCODE) T2 ON (CAST(T0.DOCENTRY AS VARCHAR)=CAST(T2.DOCENTRY AS VARCHAR)  AND T0.LINENUM=T2.LINENUM AND T0.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("  where T1.cardcode=@cardcode ");
            if (!checkBox1.Checked)
            {
                sb.Append("  AND T0.OPENCREQTY > 0 AND  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND CAST(T0.QUANTITY-ISNULL(T2.QUANTITY,0) AS INT) <> 0   ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("  AND T0.Docentry = @Docentry ");
            }

            sb.Append(" order by T0.Docentry ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPaymentDRS(string cardcode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select T0.Docentry 單號,T0.LINENUM 欄號,CASE ISNULL(T0.itemcode,'') WHEN '' THEN T0.DSCRIPTION ELSE T0.itemcode END 品名,T0.QUANTITY 數量,");
            sb.Append("              0  已沖數量,T0.QUANTITY 未沖數量 from POR1 T0 ");
            sb.Append("                 left join OPOR T1 on (T0.docentry=T1.docentry) INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T0.ItemCode  ");
            sb.Append("  where T1.cardcode=@cardcode ");
            if (!checkBox1.Checked)
            {
                sb.Append("   AND  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND T0.QUANTITY <> 0       ");
            }
            if (CLOSE != "Checked")
            {
                sb.Append("  AND T0.OPENCREQTY > 0 ");
            }
            else
            {
                sb.Append("  AND T0.OPENCREQTY = 0 ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("  AND T0.Docentry = @Docentry ");
            }
            sb.Append(" order by T0.Docentry ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    StringBuilder sb = new StringBuilder();
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];

                        sb.Append("'" + row.Cells["單號"].Value.ToString() + " " + row.Cells["欄號"].Value.ToString() + "',");
                       
                    }


                    sb.Remove(sb.Length - 1, 1);
                    string q = sb.ToString();
                    a = q;
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            ViewBatchPayment2(cardcode);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            ViewBatchPayment2(cardcode);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            S1();  
        }







   
    }
}
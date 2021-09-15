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
    public partial class SHIPINFINITE : Form
    {
        private decimal sd;
        public SHIPINFINITE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            System.Data.DataTable T1 = GetOrderDataAP();
            if (T1.Rows.Count > 0)
            {
                dataGridView1.DataSource = T1;
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
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
            sb.Append(" SELECT T0.SHIPPINGCODE 工單號碼,itemremark 單據類型,ITEMCODE 料號,DSCRIPTION 品名,T1.ItemPrice 單價,");
            sb.Append(" receivePlace 收貨地,GOALPLACE 目的地 FROM SHIPPING_MAIN T0");
            sb.Append(" LEFT JOIN SHIPPING_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE (T0.CARDCODE ='1030-00'  OR ISNULL(ADD6,'') LIKE '%DRS%'  OR ISNULL(ADD6,'') LIKE '%達睿生%')  ");
            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append("  AND SUBSTRING(t0.shippingcode,3,8) BETWEEN @aa and @bb ");
            }
            if (textBox2.Text != "")
            {
                sb.Append(" AND CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
                sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
                sb.Append(" Substring (T1.[ItemCode],2,8) END=@CC");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox2.Text));

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

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

 
 
    }
}
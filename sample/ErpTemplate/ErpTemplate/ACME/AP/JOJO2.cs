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
    public partial class JOJO2 : Form
    {
        public string PublicString;
        public JOJO2()
        {
            InitializeComponent();
        }

        private void JOJO2_Load(object sender, EventArgs e)
        {
            System.Data.DataTable T1 = GetTABLE(PublicString);
            dataGridView1.DataSource = T1;

            for (int i = 3; i <= 4; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }
        }
        private System.Data.DataTable GetTABLE(string U_BU)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT U_BU 項目群組,ITEMCODE 產品編號,ITEMNAME 產品名稱,T0.ONHAND 數量,CAST(T0.[StockValue] AS INT) 存貨金額   FROM OITM T0  ");
            sb.Append("               INNER  JOIN [dbo].[OITB] T2  ON  T0.itmsgrpcod = T2.itmsgrpcod  ");
            sb.Append("               where T0.[OnHand]>0 and t0.itemcode not in (select itemcode from oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z')) ");
            sb.Append("               And substring(t0.itemcode,1,2) <> 'ZR' ");
            sb.Append("               And substring(t0.itemcode,1,2) <> 'ZA' ");
            sb.Append("               And substring(t0.itemcode,1,2) <> 'ZB' ");
            sb.Append(" AND U_BU=@U_BU ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_BU", U_BU));


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

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToCSV2(dataGridView1, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
        }
    }
}

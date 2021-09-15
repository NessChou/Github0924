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
    public partial class JOJO4 : Form
    {
        public string PublicString;
        public JOJO4()
        {
            InitializeComponent();
        }

        private void JOJO4_Load(object sender, EventArgs e)
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
        private System.Data.DataTable GetTABLE(string CARDNAME2)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("			                  SELECT ITEMCODE 產品編號,DG 群組");
            sb.Append("               ,ITEMNAME  品名敘述,GQTY 數量,GTOTAL 金額,CARDNAME 供應商 FROM AP_JO  ");
            sb.Append("			   WHERE CARDNAME2=@CARDNAME2 AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARDNAME2", CARDNAME2));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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

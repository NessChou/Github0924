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
    public partial class HROVER : Form
    {
        public HROVER()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetSGROUP();
        }

        private System.Data.DataTable GetSGROUP()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT ODATE 日期,USERS 員工,YESNO 是否加班,REASON 不加班留下原因 FROM HR_OVERTIMED WHERE ODATE=@ODATE ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ODATE", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
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

        private void HROVER_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyyMMdd");
        }
    }
}

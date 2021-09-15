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
    public partial class ACCOUNT : Form
    {

        string strCn = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn21 = "";
        public ACCOUNT()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "進金生")
            {
                dataGridView1.DataSource = GetORDR();
            }
            else
            {

                if (comboBox1.Text == "CHOICE")
                {
                    strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
                if (comboBox1.Text == "博豐")
                {
                    strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
            
                dataGridView1.DataSource = GetORDR2();

            
            }
        }

        private System.Data.DataTable GetORDR()
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" select case postable when 'N' THEN '財報科目' WHEN 'Y' THEN '一般科目' END 分類,");
            sb.Append(" ACCTCODE 科目代碼,ACCTNAME 科目名稱,LEVELS 層次 from oact");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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
        private System.Data.DataTable GetORDR2()
        {

            SqlConnection connection = new SqlConnection(strCn21);

            StringBuilder sb = new StringBuilder();
            sb.Append(" select case IsUseSubject when 0 THEN '財報科目' WHEN 1 THEN '一般科目' END 分類 ,SubjectID 科目代碼,SubjectName 科目名稱,SubLevel 層次 from ComSubject");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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

        private void ACCOUNT_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "進金生";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}
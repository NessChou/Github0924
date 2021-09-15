using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{

    public partial class GB_CUSTINV : Form
    {

        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_CUSTINV()
        {
            InitializeComponent();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("請輸入資料");
                return;
            }
            System.Data.DataTable T1 = GetINVOICE(textBox1.Text);

            if (T1.Rows.Count == 0)
            {
                DialogResult result;
                result = MessageBox.Show("正航沒有此發票，是否要直接新增統編?", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    AddAUOGD4(textBox1.Text, textBox3.Text, DateTime.Now.ToString("yyyyMM"));

                    System.Data.DataTable K1 = GetCUSTINV();
                    dataGridView1.DataSource = K1;

                }
                else
                {
                    return;
                }
            }

            UPDATEINVOICE(textBox3.Text, textBox1.Text);
            MessageBox.Show("更新成功");
        }
        public System.Data.DataTable GetINVOICE(string SrcBillNO)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM comInvoice WHERE SrcBillNO  =@SrcBillNO ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetCUSTINV()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT YEARMON 新增年月,BILLNO 銷貨單號,CUSTNO 客戶統編  FROM GB_CUSTINV ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public void AddAUOGD4(string BILLNO, string CUSTNO, string YEARMON)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_CUSTINV(BILLNO,CUSTNO,YEARMON) values(@BILLNO,@CUSTNO,@YEARMON)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@CUSTNO", CUSTNO));
            command.Parameters.Add(new SqlParameter("@YEARMON", YEARMON));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
   
        public void UPDATEINVOICE(string TaxRegNO, string SrcBillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE comInvoice SET TaxRegNO =@TaxRegNO WHERE SrcBillNO  =@SrcBillNO   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@TaxRegNO", TaxRegNO));
            command.Parameters.Add(new SqlParameter("@SrcBillNO", SrcBillNO));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        private void GB_CUSTINV_Load(object sender, EventArgs e)
        {

            System.Data.DataTable K1 = GetCUSTINV();
            dataGridView1.DataSource = K1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

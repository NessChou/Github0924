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
            if (comboBox1.Text == "�i����")
            {
                dataGridView1.DataSource = GetORDR();
            }
            else
            {

                if (comboBox1.Text == "CHOICE")
                {
                    strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
                if (comboBox1.Text == "����")
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
            sb.Append(" select case postable when 'N' THEN '�]�����' WHEN 'Y' THEN '�@����' END ����,");
            sb.Append(" ACCTCODE ��إN�X,ACCTNAME ��ئW��,LEVELS �h�� from oact");
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
            sb.Append(" select case IsUseSubject when 0 THEN '�]�����' WHEN 1 THEN '�@����' END ���� ,SubjectID ��إN�X,SubjectName ��ئW��,SubLevel �h�� from ComSubject");

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
            comboBox1.Text = "�i����";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}
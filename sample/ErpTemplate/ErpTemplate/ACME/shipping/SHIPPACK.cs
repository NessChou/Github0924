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
    public partial class SHIPPACK : Form
    {
        public SHIPPACK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = dataGridView1.Rows[i];
                sb.Append("'" + row.Cells["來源Invoice"].Value.ToString() + "',");

            }
            if (sb.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);
            }
            else
            {
                MessageBox.Show("請輸入來源Invoice");
                return;
            }
     
            System.Data.DataTable T1 = GETDT1(sb.ToString());
            if (T1.Rows.Count > 0)
            {
                dataGridView2.DataSource = T1;
            }
            else
            {
                MessageBox.Show("請輸入來源Invoice");
            }
        }
        private System.Data.DataTable GETDT1(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.*,T1.Memo   from PackingListD T0 LEFT JOIN PackingListM T1 ON (T0.ShippingCode =T1.ShippingCode) WHERE T0. plno IN (" + SHIPPINGCODE + ")");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

     
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
        private System.Data.DataTable GETDT2(string plno)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT SHIPPINGCODE  from PackingListM WHERE plno=@plno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@plno", plno));

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
            System.Data.DataTable T2 = GETDT2(textBox1.Text.Trim());
            if (T2.Rows.Count == 0 || textBox1.Text == "")
            {
                MessageBox.Show("請輸入目的Invoice");
                return;
            }
            string ShippingCode = T2.Rows[0][0].ToString();
            if (dataGridView2.Rows.Count  == 0)
            {
                MessageBox.Show("請先預覽");
                return;
            }

            string Doctentry;
            string SeqNo;
            string PackageNo;
            string CNo;
            string DescGoods;
            string Quantity;

            string Net;
            string Gross;
            string MeasurmentCM;
            string TREETYPE;
            string VISORDER;
            string SOID;
            string PACKMARK;
            string SeqNo2;
            string MEMO;
            for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
            {


                Doctentry = dataGridView2.Rows[i].Cells["Doctentry"].Value.ToString();
                SeqNo = i.ToString();
                PackageNo = dataGridView2.Rows[i].Cells["PackageNo"].Value.ToString();
                CNo = dataGridView2.Rows[i].Cells["CNo"].Value.ToString();
                DescGoods = dataGridView2.Rows[i].Cells["DescGoods"].Value.ToString();
                Quantity = dataGridView2.Rows[i].Cells["Quantity"].Value.ToString();
                Net = dataGridView2.Rows[i].Cells["Net"].Value.ToString();
                Gross = dataGridView2.Rows[i].Cells["Gross"].Value.ToString();
                MeasurmentCM = dataGridView2.Rows[i].Cells["MeasurmentCM"].Value.ToString();
                TREETYPE = dataGridView2.Rows[i].Cells["TREETYPE"].Value.ToString();
                VISORDER = dataGridView2.Rows[i].Cells["VISORDER"].Value.ToString();
                SOID = dataGridView2.Rows[i].Cells["SOID"].Value.ToString();
                PACKMARK = dataGridView2.Rows[i].Cells["PACKMARK"].Value.ToString();
                MEMO = dataGridView2.Rows[i].Cells["MEMO"].Value.ToString();
                SeqNo2 = i.ToString();
                InsertPacking(ShippingCode, textBox1.Text, Doctentry, SeqNo, PackageNo, CNo, DescGoods, Quantity, Net, Gross, MeasurmentCM, TREETYPE, VISORDER, SOID, PACKMARK, SeqNo2);
                UPPacking(MEMO, textBox1.Text);
            }
            MessageBox.Show("匯入成功");
        }


        private void InsertPacking(string ShippingCode, string PLNo, string Doctentry, string SeqNo, string PackageNo, string CNo, string DescGoods, string Quantity, string Net, string Gross, string MeasurmentCM, string TREETYPE, string VISORDER, string SOID, string PACKMARK, string SeqNo2)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO packinglistd (ShippingCode,PLNo,Doctentry,SeqNo,PackageNo,CNo,DescGoods,Quantity,Net,Gross,MeasurmentCM,TREETYPE,VISORDER,SOID,PACKMARK,SeqNo2) VALUES(@ShippingCode,@PLNo,@Doctentry,@SeqNo,@PackageNo,@CNo,@DescGoods,@Quantity,@Net,@Gross,@MeasurmentCM,@TREETYPE,@VISORDER,@SOID,@PACKMARK,@SeqNo2)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@Doctentry", Doctentry));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@PackageNo", PackageNo));
            command.Parameters.Add(new SqlParameter("@CNo", CNo));
            command.Parameters.Add(new SqlParameter("@DescGoods", DescGoods));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Net", Net));
            command.Parameters.Add(new SqlParameter("@Gross", Gross));
            command.Parameters.Add(new SqlParameter("@MeasurmentCM", MeasurmentCM));
            command.Parameters.Add(new SqlParameter("@TREETYPE", TREETYPE));
            command.Parameters.Add(new SqlParameter("@VISORDER", VISORDER));
            command.Parameters.Add(new SqlParameter("@SOID", SOID));
            command.Parameters.Add(new SqlParameter("@PACKMARK", PACKMARK));
            command.Parameters.Add(new SqlParameter("@SeqNo2", SeqNo2));
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

        private void UPPacking(string MEMO, string PLNO)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE packinglistM SET MEMO=@MEMO WHERE PLNO=@PLNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@PLNO", PLNO));

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
    }
}

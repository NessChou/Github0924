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
    public partial class StockFreeze : Form
    {
        public StockFreeze()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt = GetordOBillSub();
                bindingSource1.DataSource = dt;

   
                dataGridView1.DataSource = bindingSource1.DataSource;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable GetordOBillSub()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.[ItemCode] 項目號碼, T0.[ItemName] 項目說明,T1.[OnHand], Convert(varchar(10),T0.[lastpurdat],111 ) 最近交易日,");
            sb.Append(" datediff(d,lastpurdat,cast('" + textBox3.Text.ToString() + "' as datetime)) 呆滯天數,T0.AvgPrice 成本  FROM OITM T0  INNER JOIN OITW T1 ");
            sb.Append(" ON T0.ItemCode = T1.ItemCode WHERE T1.[OnHand] > 0 and Substring(T0.ItemCode,1,1)<>'Z' and datediff(d,lastpurdat,cast('" + textBox3.Text.ToString() + "' as datetime))>='" + textBox4.Text.ToString() + "' ");
            if (checkBox2.Checked)
            {
                sb.Append(" and  T0.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[ItemCode] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
            }
       
                sb.Append(" and (datediff(d,lastpurdat,cast('" + textBox3.Text.ToString() + "' as datetime)) >=0)");
         
            sb.Append(" order by T0.[ItemCode]");
          
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public string d;
        private void button4_Click(object sender, EventArgs e)
        {
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
       
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.Get0itm();

            if (LookupValues != null)
            {
                textBox1.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox2_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.Get0itm();

            if (LookupValues != null)
            {
                textBox2.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void StockFreeze_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("yyyyMMdd");
        }

    }
}
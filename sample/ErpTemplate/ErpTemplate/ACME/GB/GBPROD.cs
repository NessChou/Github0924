using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace ACME
{
    public partial class GBPROD : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBPROD()
        {
            InitializeComponent();
        }
        public string d;
        private void button4_Click(object sender, EventArgs e)
        {

            APS2CHOICE frm1 = new APS2CHOICE();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetCHO3();
            dataGridView2.DataSource = GetCHO3U();
        }
        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT ProdID 料號,T1.ClassName  產品類別,BarCodeID 條碼編號,ProdName 品名規格,InvoProdName 發票品名,T0.EngName 英文品名,Unit 計量單位,SuggestPrice 建議售價, ");
            sb.Append("                CAST(PackAmt1 AS VARCHAR)+PackUnit1 包裝1,CAST(PackAmt2 AS VARCHAR)+PackUnit2 包裝2,CASE  WHEN ISNULL(ConverUnit,'') <> '' THEN");
            sb.Append("                 ConverUnit+'='+CAST(CONVERRATE AS VARCHAR)+Unit ELSE '' END 換算單位,ProdDesc 產品說明   ");
            sb.Append("                  FROM comProduct  T0 LEFT JOIN comProductClass T1 ON (T0.ClassID=T1.ClassID) ");
            sb.Append("                  WHERE (SUBSTRING(T0.ClassID,1,1)='A'  OR  T0.ClassID IN ('BCK010','BFH010','BPK050'))  AND T0.ClassID NOT IN ('ACMEFR','AMM100','AMO100','AMV100','ASC100')    ");
            sb.Append("                 AND T0.ProdName NOT LIKE '%FEE%'    ");
   
            if (checkBox2.Checked)
            {
                sb.Append(" and   T0.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T0.ProdID  between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }

            if (comboBox1.Text != "")
            {

                sb.Append("  AND T1.ClassName  = '" + comboBox1.Text + "' ");
            }

            if (textBox2.Text != "")
            {

                sb.Append("  AND T0.ProdName  LIKE '%" + textBox2.Text.ToString() + "%' ");
            }
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
        public System.Data.DataTable GetCHO3U()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ProdID 項目號碼,ProdName 品名規格,InvoProdName 發票品名,T1.ClassName 產品說明,UDef1 上架日期,UDef2 下架日期,SuggestPrice 建議售價 FROM comProduct T0  ");
            sb.Append(" LEFT JOIN comProductClass T1 ON (T0.ClassID=T1.ClassID)");
            sb.Append("   WHERE (SUBSTRING(T0.ClassID,1,1)='A' OR  T0.ClassID IN ('BCK010','BFH010','BPK050'))  AND T0.ClassID NOT IN ('ACMEFR','AMM100','AMO100','AMV100','ASC100')     ");
            sb.Append(" AND ProdName NOT LIKE '%FEE%' ");

            if (checkBox2.Checked)
            {
                sb.Append(" and   T0.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND T0.ProdID  between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }

            if (comboBox1.Text != "")
            {

                sb.Append("  AND T1.ClassName  = '" + comboBox1.Text + "' ");
            }

            if (textBox2.Text != "")
            {

                sb.Append("  AND T0.ProdName  LIKE '%" + textBox2.Text.ToString() + "%' ");
            }
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
        public  System.Data.DataTable GetBU()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT '' UNION ALL SELECT ClassName FROM comProductClass WHERE (SUBSTRING(ClassID,1,1)='A' OR  ClassID IN ('BCK010','BFH010','BPK050')) AND ClassID NOT IN ('ACMEFR','AMM100','AMO100','AMV100')  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void GBPROD_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt3 = GetBU();

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }
    }
}

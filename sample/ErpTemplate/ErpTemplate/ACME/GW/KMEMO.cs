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
    public partial class KMEMO : Form
    {
        public KMEMO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GETAMT2();
            dataGridView2.DataSource = GETAMT3();
        }

        private void KMEMO_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year2017(), "DataValue", "DataValue");
        }

        System.Data.DataTable GETAMT2()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT Convert(varchar(11),DUEDATE,111)   日期,T1.AcctName 類別, CAST(DEBIT-CREDIT AS INT) 金額,LINEMEMO 備註 FROM JDT1  T0 ");
            sb.Append(" LEFT JOIN OACT T1 ON (T0.Account =T1.AcctCode)");
            sb.Append(" WHERE YEAR(T0.TaxDate )=@YEAR AND ProfitCode IN ('11802','11803') AND  ACCOUNT BETWEEN 62110101 AND 79999999");
            if (textBox1.Text != "")
            {
                sb.Append(" and LINEMEMO LIKE '%" + textBox1.Text + "%'  ");
            }
            sb.Append(" ORDER BY DUEDATE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));


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
        System.Data.DataTable GETAMT3()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT T1.AcctName 類別, CAST(SUM(DEBIT-CREDIT) AS INT) 金額  FROM JDT1  T0  ");
            sb.Append("              LEFT JOIN OACT T1 ON (T0.Account =T1.AcctCode) ");
            sb.Append("              WHERE YEAR(T0.TaxDate )=@YEAR AND ProfitCode IN ('11802','11803') AND  ACCOUNT BETWEEN 62110101 AND 79999999 ");


            if (textBox1.Text != "")
            {
                sb.Append(" and LINEMEMO LIKE '%" + textBox1.Text + "%'  ");
            }
            sb.Append("			  GROUP BY T1.AcctName ORDER BY T1.AcctName ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));


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

        private void button2_Click(object sender, EventArgs e)
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

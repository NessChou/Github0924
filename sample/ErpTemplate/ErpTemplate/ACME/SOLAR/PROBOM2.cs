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
    public partial class PROBOM2 : Form
    {
        public PROBOM2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                System.Data.DataTable dt = GetTABLE2();
                dataGridView1.DataSource = dt;
            }
            else
            {
                System.Data.DataTable dt = GetTABLE();
                dataGridView1.DataSource = dt;
            }
        }

        private System.Data.DataTable GetTABLE()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT T0.SHIPPINGCODE JOBNO,PROJECTCODE 專案代碼,PROJECTNAME 專案名稱,DOCDATE 建立日期");
            sb.Append("  FROM sOLAR_PROBOM  T0 LEFT JOIN sOLAR_PROBOM2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("   LEFT JOIN ACMESQL02.DBO.OWOR T2  ON (T1.OWORDOC=T2.DOCENTRY)");
            sb.Append(" WHERE ISNULL(PROJECTCODE,'') <> ''");
            sb.Append(" AND T0.DOCDATE BETWEEN @DOCDATE AND @DOCDATE1  ");
            if (comboBox1.Text == "已結")
            {
                sb.Append(" AND T2.STATUS IN ('L')  ");
            }
            if (comboBox1.Text == "未結")
            {
                sb.Append(" AND T2.STATUS IN ('R')  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetTABLE2()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT  DISTINCT T0.SHIPPINGCODE JOBNO,PROJECTCODE 專案代碼,PROJECTNAME 專案名稱,DOCDATE 建立日期");
            sb.Append("        FROM sOLAR_PROBOM4 T0 ");
            sb.Append("   LEFT JOIN ACMESQL02.DBO.OWOR T2  ON (T0.OWORDOC=T2.DOCENTRY)");
            sb.Append("       WHERE ISNULL(PROJECTCODE,'') <> ''");
            sb.Append(" AND T0.DOCDATE BETWEEN @DOCDATE AND @DOCDATE1 AND VER=@VER  ");
            if (comboBox1.Text == "已結")
            {
                sb.Append(" AND T2.STATUS IN ('L')  ");
            }
            if (comboBox1.Text == "未結")
            {
                sb.Append(" AND T2.STATUS IN ('R')  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@VER", textBox3.Text.Trim()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private void SOLAPAY2_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
    
        }



        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                if (textBox3.Text == "")
                {

                    string da = dataGridView1.SelectedRows[0].Cells["JOBNO"].Value.ToString();
                    PROBOM a = new PROBOM();
                    a.PublicString2 = da;
                    a.ShowDialog();
                }
                else
                {
                    string da = dataGridView1.SelectedRows[0].Cells["JOBNO"].Value.ToString();
                    PROVER a = new PROVER();
                    a.PublicString2 = da;
                    a.PublicString3 = textBox3.Text;
                    a.ShowDialog();
                }
            }
        }



    }
}

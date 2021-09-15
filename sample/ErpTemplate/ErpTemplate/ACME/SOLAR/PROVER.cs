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
    public partial class PROVER : Form
    {
        public string N2;
        public string PublicString2;
        public string PublicString3;
        public PROVER()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetBOMVER(string shippingcode, string VER)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  PROJECTCODE 專案代碼,PROJECTNAME 專案名稱,DOCTYPE 類型,FATHER 母件編號,");
            sb.Append(" ITEMCODE 子件編號,ITEMNAME 產品名稱,DOCENTRY  採購單號,QTY  數量,PRICE 單價,OPCOST 採購成本,");
            sb.Append(" PCOST 已付採購成本,PRECOST 未付採購成本,COST 預估成本");
            sb.Append("  FROM SOLAR_PROBOM4");
            sb.Append(" WHERE shippingcode=@shippingcode AND VER=@VER ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@VER", VER));

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
        private System.Data.DataTable GetVER(string shippingcode)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT VER DataValue FROM SOLAR_PROBOM4 WHERE shippingcode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

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
        private System.Data.DataTable GetVER2(string shippingcode, string VER)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT VER DataValue FROM SOLAR_PROBOM4 WHERE shippingcode=@shippingcode AND @VER ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@VER", VER));
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
        private void PROVER_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt = null;
            if (!String.IsNullOrEmpty(PublicString2))
            {
                comboBox1.DataSource = GetVER2(PublicString2, PublicString3);
                comboBox1.DisplayMember = "DataValue";
                comboBox1.ValueMember = "DataValue";

                dt = GetBOMVER(PublicString2, PublicString3);
                dataGridView1.DataSource = dt;

            }
            else
            {

                comboBox1.DataSource = GetVER(N2);
                comboBox1.DisplayMember = "DataValue";
                comboBox1.ValueMember = "DataValue";
                
                dt = GetBOMVER(N2, comboBox1.Text.Trim());
                dataGridView1.DataSource = dt;
            }

            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 7; j <= 12; j++)
                {
                    Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);

                }
            }

            DataRow row;

            row = dt.NewRow();

            row[3] = "合計";
            for (int j = 7; j <= 12; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Data.DataTable dt = null;
            dt = GetBOMVER(N2, comboBox1.Text.Trim());
            dataGridView1.DataSource = dt;

            //加入一筆合計
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 7; j <= 12; j++)
                {
                    Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);

                }
            }

            DataRow row;

            row = dt.NewRow();

            row[3] = "合計";
            for (int j = 7; j <= 12; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }
    }
}

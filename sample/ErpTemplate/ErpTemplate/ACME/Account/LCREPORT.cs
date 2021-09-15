using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class LCREPORT : Form
    {
        public LCREPORT()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTableMONTH()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("TYPE", typeof(string));
            dt.Columns.Add("BANK", typeof(string));
            dt.Columns.Add("1月", typeof(decimal));
            dt.Columns.Add("2月", typeof(decimal));
            dt.Columns.Add("3月", typeof(decimal));
            dt.Columns.Add("4月", typeof(decimal));
            dt.Columns.Add("5月", typeof(decimal));
            dt.Columns.Add("6月", typeof(decimal));
            dt.Columns.Add("7月", typeof(decimal));
            dt.Columns.Add("8月", typeof(decimal));
            dt.Columns.Add("9月", typeof(decimal));
            dt.Columns.Add("10月", typeof(decimal));
            dt.Columns.Add("11月", typeof(decimal));
            dt.Columns.Add("12月", typeof(decimal));

            return dt;

        }

        System.Data.DataTable GetTemp(string YEAR, int MONTH, string lCTYPE, string BANK, string BANK2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(Amount),0) AMT FROM Account_LC T0 LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE SUBSTRING(LCDATE,1,4)=@YEAR AND CAST(SUBSTRING(LCDATE,5,2) AS INT) =@MONTH AND lCTYPE=@lCTYPE");

            if (BANK2 == "BANK")
            {
                sb.Append("  AND BANK=@BANK");
            }
            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@lCTYPE", lCTYPE));
            command.Parameters.Add(new SqlParameter("@BANK", BANK));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
        System.Data.DataTable GetTempC(string YEAR, int MONTH, string lCTYPE, string BANK, string BANK2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(Amount),0) AMT FROM Account_LC T0 LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE SUBSTRING(LCDATE,1,4)=@YEAR AND CAST(SUBSTRING(LCDATE,5,2) AS INT) =@MONTH AND lCTYPE=@lCTYPE");

            if (BANK2 == "BANK")
            {
                sb.Append("  AND CARDNAME=@BANK");
            }
            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@lCTYPE", lCTYPE));
            command.Parameters.Add(new SqlParameter("@BANK", BANK));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
        System.Data.DataTable GetTemp3(string YEAR, int MONTH)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(Amount),0) AMT FROM Account_LC T0 LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE SUBSTRING(LCDATE,1,4)=@YEAR AND CAST(SUBSTRING(LCDATE,5,2) AS INT) =@MONTH AND  ISNULL(LCTYPE,'') <> '' ");

            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
        System.Data.DataTable GetTemp2(string YEAR, string lCTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(Amount),0) AMT FROM Account_LC T0 LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE SUBSTRING(LCDATE,1,4)=@YEAR  AND lCTYPE=@lCTYPE");
            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@lCTYPE", lCTYPE));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
        private void button1_Click(object sender, EventArgs e)
        {

            if (radioButton1.Checked)
            {
                 
                EXEC("1");
                dataGridView1.Columns[1].HeaderText = "銀行";
            }
            if (radioButton2.Checked)
            {
                EXEC("2");
                dataGridView1.Columns[1].HeaderText = "客戶"; 
            }

            dataGridView1.Columns[0].HeaderText = ""; 
            for (int i = 2; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = dataGridView1.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0.00";

            }
        }
        private void EXEC(string F)
        {
            System.Data.DataTable dtemp5 = GetTemp61();
            System.Data.DataTable dtCostDD = MakeTableMONTH();
            DataRow drtemp5 = null;
            string YEAR = comboBox1.Text;

            drtemp5 = dtCostDD.NewRow();
            drtemp5["TYPE"] = "";
            drtemp5["BANK"] = "加總";
            for (int y = 1; y <= 12; y++)
            {

                System.Data.DataTable dh = null;
                dh = GetTemp3(YEAR, y);
                drtemp5[y + "月"] = dh.Rows[0]["AMT"].ToString();

            }
            dtCostDD.Rows.Add(drtemp5);

            string TYPE2 = "";
            for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
            {

                string TYPE = dtemp5.Rows[i]["PARAM_NO"].ToString();

                System.Data.DataTable dtemp52 = null;
                if (F == "2")
                {
                    dtemp52 = GetTemp62C(TYPE);
                }
                else
                {
                    dtemp52 = GetTemp62(TYPE);
                }
                for (int i2 = 0; i2 <= dtemp52.Rows.Count - 1; i2++)
                {
                    drtemp5 = dtCostDD.NewRow();
                    string BANK = dtemp52.Rows[i2]["BANK"].ToString();
                    if (TYPE2 != TYPE)
                    {
                        drtemp5["TYPE"] = TYPE;
                    }
                    drtemp5["BANK"] = BANK;
                    TYPE2 = TYPE;
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = null;
                        if (F == "2")
                        {
                            dh = GetTempC(YEAR, y, TYPE, BANK, "BANK");
                        }
                        else
                        {
                            dh = GetTemp(YEAR, y, TYPE, BANK, "BANK");
                        }
             
                        drtemp5[y + "月"] = dh.Rows[0]["AMT"].ToString();

                    }
                    dtCostDD.Rows.Add(drtemp5);
                }
                System.Data.DataTable G1 = GetTemp2(comboBox1.Text, TYPE);
                if (G1.Rows[0][0].ToString() != "0.000")
                {
                    drtemp5 = dtCostDD.NewRow();
                    drtemp5["TYPE"] = "";
                    drtemp5["BANK"] = "小計";

                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = null;
                        dh = GetTemp(YEAR, y, TYPE, "", "");
                        drtemp5[y + "月"] = dh.Rows[0]["AMT"].ToString();

                    }
                    dtCostDD.Rows.Add(drtemp5);
                }

            }
            dataGridView1.DataSource = dtCostDD;
        }

        private void LCREPORT_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
            radioButton1.Checked = true;
            comboBox2.Text = "ALL";
        }
        System.Data.DataTable GetTemp61()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                       SELECT PARAM_NO FROM PARAMS WHERE PARAM_KIND='ACCOUNTLC'  ");
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
        System.Data.DataTable GetTemp62(string LCTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT BANK  FROM Account_LC T0 ");
            sb.Append(" LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE ISNULL(BANK,'') <> '' AND LCTYPE=@LCTYPE AND SUBSTRING(LCDATE,1,4)=@YEAR ");
            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }
            sb.Append(" GROUP BY BANK HAVING ISNULL(SUM(Amount),0) <> 0");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCTYPE", LCTYPE));
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
        System.Data.DataTable GetTemp62C(string LCTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT CARDNAME BANK FROM Account_LC T0 ");
            sb.Append(" LEFT JOIN Account_LC2 T1 ON (T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE ISNULL(CARDNAME,'') <> '' AND LCTYPE=@LCTYPE AND SUBSTRING(LCDATE,1,4)=@YEAR ");
            if (comboBox2.Text != "ALL")
            {
                sb.Append(" AND T0.COMPANY=@COMPANY ");
            }
            sb.Append(" GROUP BY CARDNAME HAVING ISNULL(SUM(Amount),0) <> 0");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCTYPE", LCTYPE));
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", comboBox2.Text));
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
            ExcelReport.GridViewToExcel(dataGridView1); 
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView1.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["BANK"].Value.ToString() == "小計")
                {
                    dgr.DefaultCellStyle.BackColor = Color.Yellow;
                    dgr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else if (dgr.Cells["BANK"].Value.ToString() == "加總")
                {

                    dgr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dgr.DefaultCellStyle.BackColor = Color.YellowGreen;
                }
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }
    }

}

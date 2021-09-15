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
    public partial class SHIPRIGHT : Form
    {
        public SHIPRIGHT()
        {
            InitializeComponent();
        }

        private void SHIPRIGHT_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt3 = GETSHIPRIGHT();

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }

            dataGridView1.DataSource = Getshipitem07();
        }
        public System.Data.DataTable Getshipitem07()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CELLNAME 單據種類,T1.USERNAME 使用者,T2.USERNAME 部門,EDIT 可修改資料  FROM USERSSHIP T0");
            sb.Append(" LEFT JOIN dbo.[RIGHT] T1 ON (T0.USERID=T1.category)");
            sb.Append(" LEFT JOIN dbo.[USERS] T2 ON (T2.USERID=T1.category)");
            sb.Append(" LEFT JOIN dbo.USERMENUS T3 ON (T3.USERID=T1.category AND T3.MENUID='fmShip') ");
            sb.Append(" WHERE ISNULL(T1.USERNAME,'') <> '' AND T2.USERNAME <> '管理者'");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND T2.USERNAME=@USERNAME  ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CELLNAME,T1.USERNAME,T2.USERNAME,EDIT 可修改  FROM USERSSHIP T0");
            sb.Append(" LEFT JOIN dbo.[RIGHT] T1 ON (T0.USERID=substring(T1.category,0,CHARINDEX(',',T1.category)))");
            sb.Append(" LEFT JOIN dbo.[USERS] T2 ON (T2.USERID=substring(T1.category,0,CHARINDEX(',',T1.category)))");
            sb.Append(" LEFT JOIN dbo.USERMENUS T3 ON (T3.USERID=substring(T1.category,0,CHARINDEX(',',T1.category)) AND T3.MENUID='fmShip') ");
            sb.Append(" WHERE ISNULL(T1.USERNAME,'') <> '' AND T2.USERNAME <> '管理者'");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND T2.USERNAME=@USERNAME  ");
            }
            sb.Append(" ORDER BY T1.USERNAME");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", comboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GETSHIPRIGHT()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT T2.USERNAME FROM USERSSHIP T0");
            sb.Append(" LEFT JOIN dbo.[RIGHT] T1 ON (T0.USERID=substring(T1.category,CHARINDEX(',',T1.category)+1,10))");
            sb.Append(" LEFT JOIN dbo.[USERS] T2 ON (T2.USERID=substring(T1.category,CHARINDEX(',',T1.category)+1,10))");
            sb.Append(" WHERE T2.USERNAME <> '管理者'");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getshipitem07();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

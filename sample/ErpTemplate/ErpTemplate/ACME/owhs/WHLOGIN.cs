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
    public partial class WHLOGIN : Form
    {
        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public WHLOGIN()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetOrderDataAPL(string User_Code)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select case when A.User_Host like '%sz%' then '深圳' else   H.[Name] end  使用者, SUBSTRING(convert(varchar,A.Login_Time),12,8) 登入時間  ");
            sb.Append(" from   acmesqlsp..Acme_Audit_log A ");
            sb.Append(" Left Join acmesqlsp..HostName H on H.HostName=A.User_Host COLLATE  Chinese_Taiwan_Stroke_CI_AS ");
            sb.Append(" Where Convert(varchar(8),A.Login_Time,112) = Convert(varchar(8),GetDate(),112) ");
            sb.Append(" and User_Code =@User_Code ORDER BY A.Login_Time DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@User_Code", User_Code));
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

        private System.Data.DataTable GetOrderDataAPL2(string User_Code)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select case when A.User_Host like '%sz%' then '深圳' else   H.[Name] end  使用者, SUBSTRING(convert(varchar,A.Login_Time),12,8) 登入時間  ,");
            sb.Append(" CASE WHEN C.SAP <> '' THEN '使用中' END　SAP");
            sb.Append(" from   acmesqlsp..Acme_Audit_log A  ");
            sb.Append(" Left Join acmesqlsp..HostName H on H.HostName=A.User_Host COLLATE  Chinese_Taiwan_Stroke_CI_AS  ");
            sb.Append(" LEFT JOIN (SELECT MAX(HOSTNAME) HOSTNAME,MAX(program_name) SAP FROM sys.sysprocesses　WHERE  program_name ='SAP Business One' GROUP BY HOSTNAME) C ON (A.USER_HOST=C.hostname  COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" Where Convert(varchar(8),A.Login_Time,112) = Convert(varchar(8),GetDate(),112)  ");
            sb.Append(" AND User_Code=@User_Code ");
            if (textBox1.Text == "A02")
            {
                sb.Append("  AND  H.[Name] IN('cloudiawu', 'maggieweng', 'rebeccalin')");
            }
            //AND  H.[Name] IN('cloudiawu', 'maggieweng', 'rebeccalin')
            sb.Append(" ORDER BY A.Login_Time DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@User_Code", User_Code));
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
        private void button1_Click(object sender, EventArgs e)
        {
            AA();
        }

        private void ESCOLOGIN_Load(object sender, EventArgs e)
        {
            string GROUP = globals.GroupID.ToString().Trim();
            if (GROUP != "EEP" && GROUP != "ShipBuy")
            {
                textBox1.Text = "A01";
            }
            else
            {
                textBox1.Text = "A02";
            }

            AA();

        }
        private void AA()
        {
            string GROUP = globals.GroupID.ToString().Trim();
            if (GROUP != "EEP" && GROUP != "ShipBuy")
            {
                dataGridView1.DataSource = GetOrderDataAPL(textBox1.Text);
            }
            else
            {
                dataGridView1.DataSource = GetOrderDataAPL2(textBox1.Text);
            }
        }
    }
}

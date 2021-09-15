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
    public partial class ESCOLOGIN : Form
    {
        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public ESCOLOGIN()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetOrderDataAPL()
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct A.User_Code 帳號,case when A.User_Host like '%sz%' then '深圳' else   H.[Name] end  使用者, SUBSTRING(convert(varchar,A.Login_Time),12,8) 登入時間  ");
            sb.Append(" from   acmesqlsp..Acme_Audit_log A ");
            sb.Append(" Left Join acmesqlsp..HostName H on H.HostName=A.User_Host COLLATE  Chinese_Taiwan_Stroke_CI_AS ");
            sb.Append(" Where Convert(varchar(8),A.Login_Time,112) = Convert(varchar(8),GetDate(),112) ");
            sb.Append(" and User_Code IN ('A03','S01') ORDER BY A.User_Code,SUBSTRING(convert(varchar,A.Login_Time),12,8)");
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

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOrderDataAPL();
        }

        private void ESCOLOGIN_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOrderDataAPL();
        }
    }
}

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
    public partial class SASAPPER : Form
    {
        string ID = "";
        string USERNAME = "";
        public SASAPPER()
        {
            InitializeComponent();
        }
        private void UpdateMasterSQL()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Drop Table ACMESQLSP.DBO.CPRF ");
            sb.Append("               Select * Into ACMESQLSP.DBO.CPRF From  ACMESQL96.DBO.CPRF ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
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
        private void UpdateMasterSQL22(string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @FormID nvarchar(20)");
            sb.Append(" Declare @From nvarchar(20)");
            sb.Append(" Declare @To nvarchar(20)");
            sb.Append(" set @FormID = " + comboBox6.SelectedValue.ToString() + "");
            sb.Append(" set @From = " + ID + "");
            sb.Append(" set @To = " + ID + "");
            sb.Append(" if exists(Select 1 from ACMESQL98.DBO.OUSR where UserID=@To)");
            sb.Append(" begin");
            sb.Append(" Delete From  ACMESQL02.DBO.CPRF Where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@To");
            sb.Append(" Insert Into  ACMESQLSP.DBO.CPRF2 Select * From ACMESQLSP.DBO.CPRF where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@From");
            sb.Append(" update ACMESQLSP.DBO.CPRF2 set UserSign=@to");
            sb.Append(" Insert Into  ACMESQL02.DBO.CPRF Select * From ACMESQLSP.DBO.CPRF2");
            sb.Append(" truncate table ACMESQLSP.DBO.CPRF2 ");
            sb.Append(" end");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
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
        private System.Data.DataTable GetSCE()
        {
            SqlConnection connection = globals.shipConnection;
            //  SqlConnection connection2 =globals_SAP.
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct t0.hostname 電腦,t1.[name] 登入者,login_time 登入時間,last_batch 最後動作時間 from master..sysprocesses t0");
            sb.Append(" left join acmesqlsp.dbo.hostname t1 on(t0.hostname=t1.hostname COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where program_name='SAP Business One' and dbid=13 ");
            sb.Append(" order by t1.[name] ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetUSERCODE(string CODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select userid   from ousr where  user_code=@CODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CODE", CODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void SASAPPER_Load(object sender, EventArgs e)
        {

            UtilSimple.SetLookupBinding(comboBox6, Getpdn1(), "DataText", "DataValue");
            System.Data.DataTable T1 = GetUSER(fmLogin.LoginID.ToString().Trim());
            if (T1.Rows.Count > 0)
            {
                 USERNAME = T1.Rows[0]["USERNAME"].ToString().Trim();
                 ID = T1.Rows[0]["ID"].ToString().Trim();
                 label1.Text = "SAP帳號 : " + ID;
                 label3.Text = "使用者 : " + USERNAME;
            }
        }
        public static System.Data.DataTable Getpdn1()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY as datavalue, FormName as datatext FROM FORMID ");
            sb.Append(" WHERE DOCENTRY IN (149,139,60126)");
            sb.Append(" ORDER BY DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "FORMID");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["FORMID"];
        }
        private System.Data.DataTable GetUSER(string USERNAME)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("select USER_CODE ID,U_NAME USERNAME from   ACMESQL02.DBO.ousr  WHERE  u_name LIKE '%" + USERNAME + "%'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ID == "")
            {
                MessageBox.Show("SAP帳號尚未設定連結，請通知Lleyton設定");
                return;
            }
                      DialogResult result;
                      result = MessageBox.Show("請確定SAP使用者 '" + USERNAME + "' 是否已登出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                System.Data.DataTable t1 = GetUSERCODE(ID);
                if (t1.Rows.Count > 0)
                {
                    ID = t1.Rows[0][0].ToString();
                    UpdateMasterSQL22(ID);
                    MessageBox.Show("重整完成");
                }
                else
                {
                    MessageBox.Show("重整失敗，請聯絡MIS");
                }
                
            }
        }

    }
}

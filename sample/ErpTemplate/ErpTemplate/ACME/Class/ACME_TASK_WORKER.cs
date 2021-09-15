 using System;
 using System.Data;
 using System.Configuration;
 using System.Web;
 using System.Web.Security;
 using System.Web.UI;
 using System.Web.UI.WebControls;
 using System.Web.UI.WebControls.WebParts;
 using System.Web.UI.HtmlControls;
 using System.Data.SqlClient;
 
 /// <summary>
 /// Summary description for ACME_TASK_WORKER
 /// 作者:
 /// </summary>
// ACME_TASK_WORKER 資料結構
namespace ACME
{
    public class ACME_TASK_WORKER
    {
        private int _ID;
        private string _UserCode;
        private string _UserName;
        private string _Enabled;

        public int ID { get { return _ID; } set { _ID = value; } }
        public string UserCode { get { return _UserCode; } set { _UserCode = value; } }
        public string UserName { get { return _UserName; } set { _UserName = value; } }
        public string Enabled { get { return _Enabled; } set { _Enabled = value; } }

        public ACME_TASK_WORKER(int ID, string UserCode, string UserName, string Enabled)
        {
            _ID = ID;
            _UserCode = UserCode;
            _UserName = UserName;
            _Enabled = Enabled;
        }
        public ACME_TASK_WORKER()
        {
        }
        // ACME_TASK_WORKER Insert
        public static void AddACME_TASK_WORKER(ACME_TASK_WORKER row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into ACME_TASK_WORKER(UserCode,UserName,Enabled) values(@UserCode,@UserName,@Enabled)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserName", SqlDbType.VarChar, 50, "UserName"));
            command.Parameters["@UserName"].Value = row.UserName;
            if (String.IsNullOrEmpty(row.UserName))
            {
                command.Parameters["@UserName"].IsNullable = true;
                command.Parameters["@UserName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Enabled", SqlDbType.VarChar, 1, "Enabled"));
            command.Parameters["@Enabled"].Value = row.Enabled;
            if (String.IsNullOrEmpty(row.Enabled))
            {
                command.Parameters["@Enabled"].IsNullable = true;
                command.Parameters["@Enabled"].Value = "";
            }
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        // ACME_TASK_WORKER Update
        public static void UpdateACME_TASK_WORKER(ACME_TASK_WORKER row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE ACME_TASK_WORKER SET UserCode = @UserCode,UserName = @UserName,Enabled = @Enabled WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserName", SqlDbType.VarChar, 50, "UserName"));
            command.Parameters["@UserName"].Value = row.UserName;
            if (String.IsNullOrEmpty(row.UserName))
            {
                command.Parameters["@UserName"].IsNullable = true;
                command.Parameters["@UserName"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Enabled", SqlDbType.VarChar, 1, "Enabled"));
            command.Parameters["@Enabled"].Value = row.Enabled;
            if (String.IsNullOrEmpty(row.Enabled))
            {
                command.Parameters["@Enabled"].IsNullable = true;
                command.Parameters["@Enabled"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@wID", row.ID));
           
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        // ACME_TASK_WORKER Delete
        public static void DeleteACME_TASK_WORKER(ACME_TASK_WORKER row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE ACME_TASK_WORKER WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        // ACME_TASK_WORKER Select
        public static DataTable GetACME_TASK_WORKER(ACME_TASK_WORKER row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,UserCode,UserName,Enabled FROM ACME_TASK_WORKER WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_WORKER");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_WORKER"];
        }

        // ACME_TASK_WORKER Select
        public static DataTable GetACME_TASK_WORKER(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,UserCode,UserName,Enabled FROM ACME_TASK_WORKER WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_WORKER");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_WORKER"];
        }
        // Condition 版本
        public static DataTable GetACME_TASK_WORKER_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,UserCode,UserName,Enabled FROM ACME_TASK_WORKER WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_WORKER");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_WORKER"];
        }

    }
}
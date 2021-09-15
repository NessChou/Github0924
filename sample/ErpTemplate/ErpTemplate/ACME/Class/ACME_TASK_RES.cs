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
 /// Summary description for ACME_TASK_RES
 /// 作者:
 /// </summary>
// ACME_TASK_RES 資料結構
namespace ACME
{
    public class ACME_TASK_RES
    {
        private int _ID;
        private string _PrjCode;
        private string _UserCode;
        private string _Task;
        private string _StartDate;
        private string _EndDate;
        private Decimal _WorkDate;

        public int ID { get { return _ID; } set { _ID = value; } }
        public string PrjCode { get { return _PrjCode; } set { _PrjCode = value; } }
        public string UserCode { get { return _UserCode; } set { _UserCode = value; } }
        public string Task { get { return _Task; } set { _Task = value; } }
        public string StartDate { get { return _StartDate; } set { _StartDate = value; } }
        public string EndDate { get { return _EndDate; } set { _EndDate = value; } }
        public Decimal WorkDate { get { return _WorkDate; } set { _WorkDate = value; } }

        public ACME_TASK_RES(int ID, string PrjCode, string UserCode, string Task, string StartDate, string EndDate, decimal WorkDate)
        {
            _ID = ID;
            _PrjCode = PrjCode;
            _UserCode = UserCode;
            _Task = Task;
            _StartDate = StartDate;
            _EndDate = EndDate;
            _WorkDate = WorkDate;
        }
        public ACME_TASK_RES()
        {
        }
        // ACME_TASK_RES Insert
        public static void AddACME_TASK_RES(ACME_TASK_RES row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into ACME_TASK_RES(PrjCode,UserCode,Task,StartDate,EndDate,WorkDate) values(@PrjCode,@UserCode,@Task,@StartDate,@EndDate,@WorkDate)", connection);
            command.CommandType = CommandType.Text;
       
            command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
            command.Parameters["@PrjCode"].Value = row.PrjCode;
            if (String.IsNullOrEmpty(row.PrjCode))
            {
                command.Parameters["@PrjCode"].IsNullable = true;
                command.Parameters["@PrjCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar, 50, "Task"));
            command.Parameters["@Task"].Value = row.Task;
            if (String.IsNullOrEmpty(row.Task))
            {
                command.Parameters["@Task"].IsNullable = true;
                command.Parameters["@Task"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.VarChar, 8, "StartDate"));
            command.Parameters["@StartDate"].Value = row.StartDate;
            if (String.IsNullOrEmpty(row.StartDate))
            {
                command.Parameters["@StartDate"].IsNullable = true;
                command.Parameters["@StartDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.VarChar, 8, "EndDate"));
            command.Parameters["@EndDate"].Value = row.EndDate;
            if (String.IsNullOrEmpty(row.EndDate))
            {
                command.Parameters["@EndDate"].IsNullable = true;
                command.Parameters["@EndDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WorkDate", row.WorkDate));
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

        // ACME_TASK_RES Update
        public static void UpdateACME_TASK_RES(ACME_TASK_RES row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE ACME_TASK_RES SET PrjCode = @PrjCode,UserCode = @UserCode,Task = @Task,StartDate = @StartDate,EndDate = @EndDate,WorkDate = @WorkDate WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
            command.Parameters["@PrjCode"].Value = row.PrjCode;
            if (String.IsNullOrEmpty(row.PrjCode))
            {
                command.Parameters["@PrjCode"].IsNullable = true;
                command.Parameters["@PrjCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UserCode", SqlDbType.VarChar, 20, "UserCode"));
            command.Parameters["@UserCode"].Value = row.UserCode;
            if (String.IsNullOrEmpty(row.UserCode))
            {
                command.Parameters["@UserCode"].IsNullable = true;
                command.Parameters["@UserCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar, 50, "Task"));
            command.Parameters["@Task"].Value = row.Task;
            if (String.IsNullOrEmpty(row.Task))
            {
                command.Parameters["@Task"].IsNullable = true;
                command.Parameters["@Task"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.VarChar, 8, "StartDate"));
            command.Parameters["@StartDate"].Value = row.StartDate;
            if (String.IsNullOrEmpty(row.StartDate))
            {
                command.Parameters["@StartDate"].IsNullable = true;
                command.Parameters["@StartDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.VarChar, 8, "EndDate"));
            command.Parameters["@EndDate"].Value = row.EndDate;
            if (String.IsNullOrEmpty(row.EndDate))
            {
                command.Parameters["@EndDate"].IsNullable = true;
                command.Parameters["@EndDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WorkDate", row.WorkDate));

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

        // ACME_TASK_RES Delete
        public static void DeleteACME_TASK_RES(ACME_TASK_RES row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE ACME_TASK_RES WHERE ID=@ID";
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

        // ACME_TASK_RES Select
        public static DataTable GetACME_TASK_RES(ACME_TASK_RES row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,UserCode,Task,StartDate,EndDate,WorkDate FROM ACME_TASK_RES WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_RES");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_RES"];
        }

        // ACME_TASK_RES Select
        public static DataTable GetACME_TASK_RES(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,UserCode,Task,StartDate,EndDate,WorkDate FROM ACME_TASK_RES WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_RES");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_RES"];
        }
        // Condition 版本
        public static DataTable GetACME_TASK_RES_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,UserCode,Task,StartDate,EndDate,WorkDate FROM ACME_TASK_RES WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_RES");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_RES"];
        }

    }
}
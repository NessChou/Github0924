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
 /// Summary description for ACME_CREDIT_UNLOCK
 /// 作者:
 /// </summary>
// ACME_CREDIT_UNLOCK 資料結構
namespace ACME
{
    public class ACME_CREDIT_UNLOCK
    {
        private int _DocEntry;
        private string _CardCode;
        private string _CardName;
        private string _StartDate;
        private string _EndDate;
        private string _Reason;
        private string _Applicate;
        private string _Handler;
        private string _CreateDate;
        private string _CreateTime;
        private string _FlowFlag;

        public int DocEntry { get { return _DocEntry; } set { _DocEntry = value; } }
        public string CardCode { get { return _CardCode; } set { _CardCode = value; } }
        public string CardName { get { return _CardName; } set { _CardName = value; } }
        public string StartDate { get { return _StartDate; } set { _StartDate = value; } }
        public string EndDate { get { return _EndDate; } set { _EndDate = value; } }
        public string Reason { get { return _Reason; } set { _Reason = value; } }
        public string Applicate { get { return _Applicate; } set { _Applicate = value; } }
        public string Handler { get { return _Handler; } set { _Handler = value; } }
        public string CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
        public string CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
        public string FlowFlag { get { return _FlowFlag; } set { _FlowFlag = value; } }

        public ACME_CREDIT_UNLOCK(int DocEntry, string CardCode,string CardName, string StartDate, string EndDate, string Reason, string Applicate, string Handler, string CreateDate, string CreateTime, string FlowFlag)
        {
            _DocEntry = DocEntry;

            _CardCode = CardCode;
            _CardName = CardName;
            _StartDate = StartDate;
            _EndDate = EndDate;
            _Reason = Reason;
            _Applicate = Applicate;
            _Handler = Handler;
            _CreateDate = CreateDate;
            _CreateTime = CreateTime;
            _FlowFlag = FlowFlag;
        }
        public ACME_CREDIT_UNLOCK()
        {
        }
        // ACME_CREDIT_UNLOCK Insert
        public static void AddACME_CREDIT_UNLOCK(ACME_CREDIT_UNLOCK row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into ACME_CREDIT_UNLOCK(CardCode,CardName,StartDate,EndDate,Reason,Applicate,Handler,CreateDate,CreateTime,FlowFlag) values(@CardCode,@CardName,@StartDate,@EndDate,@Reason,@Applicate,@Handler,@CreateDate,@CreateTime,@FlowFlag)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CardCode", SqlDbType.VarChar, 20, "CardCode"));
            command.Parameters["@CardCode"].Value = row.CardCode;
            if (String.IsNullOrEmpty(row.CardCode))
            {
                command.Parameters["@CardCode"].IsNullable = true;
                command.Parameters["@CardCode"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@CardName", SqlDbType.VarChar, 20, "CardName"));
            command.Parameters["@CardName"].Value = row.CardCode;
            if (String.IsNullOrEmpty(row.CardCode))
            {
                command.Parameters["@CardName"].IsNullable = true;
                command.Parameters["@CardName"].Value = "";
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
            command.Parameters.Add(new SqlParameter("@Reason", SqlDbType.VarChar, 100, "Reason"));
            command.Parameters["@Reason"].Value = row.Reason;
            if (String.IsNullOrEmpty(row.Reason))
            {
                command.Parameters["@Reason"].IsNullable = true;
                command.Parameters["@Reason"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Applicate", SqlDbType.VarChar, 20, "Applicate"));
            command.Parameters["@Applicate"].Value = row.Applicate;
            if (String.IsNullOrEmpty(row.Applicate))
            {
                command.Parameters["@Applicate"].IsNullable = true;
                command.Parameters["@Applicate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Handler", SqlDbType.VarChar, 20, "Handler"));
            command.Parameters["@Handler"].Value = row.Handler;
            if (String.IsNullOrEmpty(row.Handler))
            {
                command.Parameters["@Handler"].IsNullable = true;
                command.Parameters["@Handler"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
            command.Parameters["@CreateDate"].Value = row.CreateDate;
            if (String.IsNullOrEmpty(row.CreateDate))
            {
                command.Parameters["@CreateDate"].IsNullable = true;
                command.Parameters["@CreateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 4, "CreateTime"));
            command.Parameters["@CreateTime"].Value = row.CreateTime;
            if (String.IsNullOrEmpty(row.CreateTime))
            {
                command.Parameters["@CreateTime"].IsNullable = true;
                command.Parameters["@CreateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@FlowFlag", SqlDbType.VarChar, 1, "FlowFlag"));
            command.Parameters["@FlowFlag"].Value = row.FlowFlag;
            if (String.IsNullOrEmpty(row.FlowFlag))
            {
                command.Parameters["@FlowFlag"].IsNullable = true;
                command.Parameters["@FlowFlag"].Value = "";
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

        // ACME_CREDIT_UNLOCK Update
        public static void UpdateACME_CREDIT_UNLOCK(ACME_CREDIT_UNLOCK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE ACME_CREDIT_UNLOCK SET CardCode = @CardCode,CardName=@CardName,StartDate = @StartDate,EndDate = @EndDate,Reason = @Reason,Applicate = @Applicate,Handler = @Handler,CreateDate = @CreateDate,CreateTime = @CreateTime,FlowFlag = @FlowFlag WHERE DocEntry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", row.DocEntry));

            command.Parameters.Add(new SqlParameter("@CardCode", SqlDbType.VarChar, 20, "CardCode"));
            command.Parameters["@CardCode"].Value = row.CardCode;
            if (String.IsNullOrEmpty(row.CardCode))
            {
                command.Parameters["@CardCode"].IsNullable = true;
                command.Parameters["@CardCode"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@CardName", SqlDbType.NVarChar, 200, "CardName"));
            command.Parameters["@CardName"].Value = row.CardName;
            if (String.IsNullOrEmpty(row.CardName))
            {
                command.Parameters["@CardName"].IsNullable = true;
                command.Parameters["@CardName"].Value = "";
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
            command.Parameters.Add(new SqlParameter("@Reason", SqlDbType.VarChar, 100, "Reason"));
            command.Parameters["@Reason"].Value = row.Reason;
            if (String.IsNullOrEmpty(row.Reason))
            {
                command.Parameters["@Reason"].IsNullable = true;
                command.Parameters["@Reason"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Applicate", SqlDbType.VarChar, 20, "Applicate"));
            command.Parameters["@Applicate"].Value = row.Applicate;
            if (String.IsNullOrEmpty(row.Applicate))
            {
                command.Parameters["@Applicate"].IsNullable = true;
                command.Parameters["@Applicate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Handler", SqlDbType.VarChar, 20, "Handler"));
            command.Parameters["@Handler"].Value = row.Handler;
            if (String.IsNullOrEmpty(row.Handler))
            {
                command.Parameters["@Handler"].IsNullable = true;
                command.Parameters["@Handler"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
            command.Parameters["@CreateDate"].Value = row.CreateDate;
            if (String.IsNullOrEmpty(row.CreateDate))
            {
                command.Parameters["@CreateDate"].IsNullable = true;
                command.Parameters["@CreateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 4, "CreateTime"));
            command.Parameters["@CreateTime"].Value = row.CreateTime;
            if (String.IsNullOrEmpty(row.CreateTime))
            {
                command.Parameters["@CreateTime"].IsNullable = true;
                command.Parameters["@CreateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@FlowFlag", SqlDbType.VarChar, 1, "FlowFlag"));
            command.Parameters["@FlowFlag"].Value = row.FlowFlag;
            if (String.IsNullOrEmpty(row.FlowFlag))
            {
                command.Parameters["@FlowFlag"].IsNullable = true;
                command.Parameters["@FlowFlag"].Value = "";
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

        // ACME_CREDIT_UNLOCK Delete
        public static void DeleteACME_CREDIT_UNLOCK(ACME_CREDIT_UNLOCK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE ACME_CREDIT_UNLOCK WHERE DocEntry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", row.DocEntry));
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

        // ACME_CREDIT_UNLOCK Select
        public static DataTable GetACME_CREDIT_UNLOCK(ACME_CREDIT_UNLOCK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT DocEntry,CardCode,CardName,StartDate,EndDate,Reason,Applicate,Handler,CreateDate,CreateTime,FlowFlag FROM ACME_CREDIT_UNLOCK WHERE 1= 1  AND DocEntry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", row.DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_CREDIT_UNLOCK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_CREDIT_UNLOCK"];
        }

        // ACME_CREDIT_UNLOCK Select
        public static DataTable GetACME_CREDIT_UNLOCK(int DocEntry)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT DocEntry,CardCode,CardName,StartDate,EndDate,Reason,Applicate,Handler,CreateDate,CreateTime,FlowFlag FROM ACME_CREDIT_UNLOCK WHERE 1= 1  AND DocEntry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_CREDIT_UNLOCK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_CREDIT_UNLOCK"];
        }
        // Condition 版本
        public static DataTable GetACME_CREDIT_UNLOCK_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT DocEntry,CardCode,CardName,StartDate,EndDate,Reason,Applicate,Handler,CreateDate,CreateTime,FlowFlag FROM ACME_CREDIT_UNLOCK WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_CREDIT_UNLOCK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_CREDIT_UNLOCK"];
        }

    }

}
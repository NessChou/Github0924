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
 /// Summary description for ACME_MIS_TASK
 /// 作者:
 /// </summary>
// ACME_MIS_TASK 資料結構
namespace ACME
{
    public class ACME_MIS_TASK
    {
        private int _ID;
        private string _Kind;
        private string _Task;
        private string _StartDate;
        private string _EndDate;
        private string _AcDate;
        private string _Owner;
        private string _CreateDate;
        private string _CreateTime;
        private string _CreateUser;
        private string _UpdateDate;
        private string _UpdateTime;
        private string _UpdateUser;
        private string _BU;
        private string _UNIT;
        private string _EDIT;
        private string _ITGROUP;

        public int ID { get { return _ID; } set { _ID = value; } }
        public string Kind { get { return _Kind; } set { _Kind = value; } }
        public string Task { get { return _Task; } set { _Task = value; } }
        public string StartDate { get { return _StartDate; } set { _StartDate = value; } }
        public string EndDate { get { return _EndDate; } set { _EndDate = value; } }
        public string AcDate { get { return _AcDate; } set { _AcDate = value; } }
        public string Owner { get { return _Owner; } set { _Owner = value; } }
        public string CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
        public string CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
        public string CreateUser { get { return _CreateUser; } set { _CreateUser = value; } }
        public string UpdateDate { get { return _UpdateDate; } set { _UpdateDate = value; } }
        public string UpdateTime { get { return _UpdateTime; } set { _UpdateTime = value; } }
        public string UpdateUser { get { return _UpdateUser; } set { _UpdateUser = value; } }
        public string BU { get { return _BU; } set { _BU = value; } }
        public string UNIT { get { return _UNIT; } set { _UNIT = value; } }
        public string EDIT { get { return _EDIT; } set { _EDIT = value; } }
        public string ITGROUP { get { return _ITGROUP; } set { _ITGROUP = value; } }
        public ACME_MIS_TASK(int ID, string Kind, string Task, string StartDate, string EndDate, string AcDate, string Owner, string
      CreateDate, string CreateTime, string CreateUser, string UpdateDate, string UpdateTime, string UpdateUser, string BU, string UNIT, string EDIT, string ITGROUP)
        {
            _ID = ID;
            _Kind = Kind;
            _Task = Task;
            _StartDate = StartDate;
            _EndDate = EndDate;
            _AcDate = AcDate;
            _Owner = Owner;
            _CreateDate = CreateDate;
            _CreateTime = CreateTime;
            _CreateUser = CreateUser;
            _UpdateDate = UpdateDate;
            _UpdateTime = UpdateTime;
            _UpdateUser = UpdateUser;
            _BU = BU;
            _UNIT = UNIT;
            _EDIT = EDIT;
            _ITGROUP = ITGROUP;
        }
        public ACME_MIS_TASK()
        {
        }
        // ACME_MIS_TASK Insert
        public static void AddACME_MIS_TASK(ACME_MIS_TASK row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into ACME_MIS_TASK(Kind,Task,StartDate,EndDate,AcDate,Owner,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser,BU,UNIT,EDIT,ITGROUP) values(@Kind,@Task,@StartDate,@EndDate,@AcDate,@Owner,@CreateDate,@CreateTime,@CreateUser,@UpdateDate,@UpdateTime,@UpdateUser,@BU,@UNIT,@EDIT,@ITGROUP)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Kind", SqlDbType.VarChar, 50, "Kind"));
            command.Parameters["@Kind"].Value = row.Kind;
            if (String.IsNullOrEmpty(row.Kind))
            {
                command.Parameters["@Kind"].IsNullable = true;
                command.Parameters["@Kind"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar, 200, "Task"));
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
            command.Parameters.Add(new SqlParameter("@AcDate", SqlDbType.VarChar, 8, "AcDate"));
            command.Parameters["@AcDate"].Value = row.AcDate;
            if (String.IsNullOrEmpty(row.AcDate))
            {
                command.Parameters["@AcDate"].IsNullable = true;
                command.Parameters["@AcDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Owner", SqlDbType.VarChar, 20, "Owner"));
            command.Parameters["@Owner"].Value = row.Owner;
            if (String.IsNullOrEmpty(row.Owner))
            {
                command.Parameters["@Owner"].IsNullable = true;
                command.Parameters["@Owner"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
            command.Parameters["@CreateDate"].Value = row.CreateDate;
            if (String.IsNullOrEmpty(row.CreateDate))
            {
                command.Parameters["@CreateDate"].IsNullable = true;
                command.Parameters["@CreateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 6, "CreateTime"));
            command.Parameters["@CreateTime"].Value = row.CreateTime;
            if (String.IsNullOrEmpty(row.CreateTime))
            {
                command.Parameters["@CreateTime"].IsNullable = true;
                command.Parameters["@CreateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateUser", SqlDbType.VarChar, 20, "CreateUser"));
            command.Parameters["@CreateUser"].Value = row.CreateUser;
            if (String.IsNullOrEmpty(row.CreateUser))
            {
                command.Parameters["@CreateUser"].IsNullable = true;
                command.Parameters["@CreateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateDate", SqlDbType.VarChar, 8, "UpdateDate"));
            command.Parameters["@UpdateDate"].Value = row.UpdateDate;
            if (String.IsNullOrEmpty(row.UpdateDate))
            {
                command.Parameters["@UpdateDate"].IsNullable = true;
                command.Parameters["@UpdateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateTime", SqlDbType.VarChar, 6, "UpdateTime"));
            command.Parameters["@UpdateTime"].Value = row.UpdateTime;
            if (String.IsNullOrEmpty(row.UpdateTime))
            {
                command.Parameters["@UpdateTime"].IsNullable = true;
                command.Parameters["@UpdateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
            command.Parameters["@UpdateUser"].Value = row.UpdateUser;
            if (String.IsNullOrEmpty(row.UpdateUser))
            {
                command.Parameters["@UpdateUser"].IsNullable = true;
                command.Parameters["@UpdateUser"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@BU", SqlDbType.VarChar, 20, "BU"));
            command.Parameters["@BU"].Value = row.BU;
            if (String.IsNullOrEmpty(row.BU))
            {
                command.Parameters["@BU"].IsNullable = true;
                command.Parameters["@BU"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UNIT", SqlDbType.VarChar, 50, "UNIT"));
            command.Parameters["@UNIT"].Value = row.UNIT;
            if (String.IsNullOrEmpty(row.UNIT))
            {
                command.Parameters["@UNIT"].IsNullable = true;
                command.Parameters["@UNIT"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@EDIT", SqlDbType.NChar, 10, "EDIT"));
            command.Parameters["@EDIT"].Value = row.EDIT;
            if (String.IsNullOrEmpty(row.EDIT))
            {
                command.Parameters["@EDIT"].IsNullable = true;
                command.Parameters["@EDIT"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@ITGROUP", SqlDbType.NVarChar , 50, "ITGROUP"));
            command.Parameters["@ITGROUP"].Value = row.ITGROUP;
            if (String.IsNullOrEmpty(row.EDIT))
            {
                command.Parameters["@ITGROUP"].IsNullable = true;
                command.Parameters["@ITGROUP"].Value = "";
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

        // ACME_MIS_TASK Update
        public static void UpdateACME_MIS_TASK(ACME_MIS_TASK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE ACME_MIS_TASK SET Kind = @Kind,Task = @Task,StartDate = @StartDate,EndDate = @EndDate,AcDate = @AcDate,Owner = @Owner,CreateDate = @CreateDate,CreateTime = @CreateTime,CreateUser = @CreateUser,UpdateDate = @UpdateDate,UpdateTime = @UpdateTime,UpdateUser = @UpdateUser,BU = @BU,UNIT = @UNIT,EDIT=@EDIT,ITGROUP=@ITGROUP WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@Kind", SqlDbType.VarChar, 50, "Kind"));
            command.Parameters["@Kind"].Value = row.Kind;
            if (String.IsNullOrEmpty(row.Kind))
            {
                command.Parameters["@Kind"].IsNullable = true;
                command.Parameters["@Kind"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar, 200, "Task"));
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
            command.Parameters.Add(new SqlParameter("@AcDate", SqlDbType.VarChar, 8, "AcDate"));
            command.Parameters["@AcDate"].Value = row.AcDate;
            if (String.IsNullOrEmpty(row.AcDate))
            {
                command.Parameters["@AcDate"].IsNullable = true;
                command.Parameters["@AcDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@Owner", SqlDbType.VarChar, 20, "Owner"));
            command.Parameters["@Owner"].Value = row.Owner;
            if (String.IsNullOrEmpty(row.Owner))
            {
                command.Parameters["@Owner"].IsNullable = true;
                command.Parameters["@Owner"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar, 8, "CreateDate"));
            command.Parameters["@CreateDate"].Value = row.CreateDate;
            if (String.IsNullOrEmpty(row.CreateDate))
            {
                command.Parameters["@CreateDate"].IsNullable = true;
                command.Parameters["@CreateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateTime", SqlDbType.VarChar, 6, "CreateTime"));
            command.Parameters["@CreateTime"].Value = row.CreateTime;
            if (String.IsNullOrEmpty(row.CreateTime))
            {
                command.Parameters["@CreateTime"].IsNullable = true;
                command.Parameters["@CreateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@CreateUser", SqlDbType.VarChar, 20, "CreateUser"));
            command.Parameters["@CreateUser"].Value = row.CreateUser;
            if (String.IsNullOrEmpty(row.CreateUser))
            {
                command.Parameters["@CreateUser"].IsNullable = true;
                command.Parameters["@CreateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateDate", SqlDbType.VarChar, 8, "UpdateDate"));
            command.Parameters["@UpdateDate"].Value = row.UpdateDate;
            if (String.IsNullOrEmpty(row.UpdateDate))
            {
                command.Parameters["@UpdateDate"].IsNullable = true;
                command.Parameters["@UpdateDate"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateTime", SqlDbType.VarChar, 6, "UpdateTime"));
            command.Parameters["@UpdateTime"].Value = row.UpdateTime;
            if (String.IsNullOrEmpty(row.UpdateTime))
            {
                command.Parameters["@UpdateTime"].IsNullable = true;
                command.Parameters["@UpdateTime"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UpdateUser", SqlDbType.VarChar, 20, "UpdateUser"));
            command.Parameters["@UpdateUser"].Value = row.UpdateUser;
            if (String.IsNullOrEmpty(row.UpdateUser))
            {
                command.Parameters["@UpdateUser"].IsNullable = true;
                command.Parameters["@UpdateUser"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@BU", SqlDbType.VarChar, 20, "BU"));
            command.Parameters["@BU"].Value = row.BU;
            if (String.IsNullOrEmpty(row.BU))
            {
                command.Parameters["@BU"].IsNullable = true;
                command.Parameters["@BU"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@UNIT", SqlDbType.VarChar, 50, "UNIT"));
            command.Parameters["@UNIT"].Value = row.UNIT;
            if (String.IsNullOrEmpty(row.UNIT))
            {
                command.Parameters["@UNIT"].IsNullable = true;
                command.Parameters["@UNIT"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@EDIT", SqlDbType.NChar, 10, "EDIT"));
            command.Parameters["@EDIT"].Value = row.EDIT;
            if (String.IsNullOrEmpty(row.EDIT))
            {
                command.Parameters["@EDIT"].IsNullable = true;
                command.Parameters["@EDIT"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@ITGROUP", SqlDbType.NVarChar, 50, "ITGROUP"));
            command.Parameters["@ITGROUP"].Value = row.ITGROUP;
            if (String.IsNullOrEmpty(row.EDIT))
            {
                command.Parameters["@ITGROUP"].IsNullable = true;
                command.Parameters["@ITGROUP"].Value = "";
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

        // ACME_MIS_TASK Delete
        public static void DeleteACME_MIS_TASK(ACME_MIS_TASK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE ACME_MIS_TASK WHERE ID=@ID";
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

        // ACME_MIS_TASK Select
        public static DataTable GetACME_MIS_TASK(ACME_MIS_TASK row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,Kind,Task,StartDate,EndDate,AcDate,Owner,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser,BU,UNIT,EDIT,ITGROUP FROM ACME_MIS_TASK WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }

        // ACME_MIS_TASK Select
        public static DataTable GetACME_MIS_TASK(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,Kind,Task,StartDate,EndDate,AcDate,Owner,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser,BU,UNIT,EDIT,ITGROUP FROM ACME_MIS_TASK WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }
        // Condition 版本
        public static DataTable GetACME_MIS_TASK_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,Kind,Task,StartDate,EndDate,AcDate,Owner,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser,BU,UNIT,EDIT FROM ACME_MIS_TASK WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }
    }
}


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
 /// Summary description for ACME_TASK_File
 /// 作者:
 /// </summary>
// ACME_TASK_File 資料結構
namespace ACME
{
    public class ACME_TASK_File
    {
        private int _ID;
        private string _PrjCode;
        private string _ObjType;
        private string _DocNo;
        private string _FileName;
        private string _CreateDate;
        private string _CreateTime;
        private string _CreateUser;
        private string _UpdateDate;
        private string _UpdateTime;
        private string _UpdateUser;

        public int ID { get { return _ID; } set { _ID = value; } }
        public string PrjCode { get { return _PrjCode; } set { _PrjCode = value; } }
        public string ObjType { get { return _ObjType; } set { _ObjType = value; } }
        public string DocNo { get { return _DocNo; } set { _DocNo = value; } }
        public string FileName { get { return _FileName; } set { _FileName = value; } }
        public string CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
        public string CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
        public string CreateUser { get { return _CreateUser; } set { _CreateUser = value; } }
        public string UpdateDate { get { return _UpdateDate; } set { _UpdateDate = value; } }
        public string UpdateTime { get { return _UpdateTime; } set { _UpdateTime = value; } }
        public string UpdateUser { get { return _UpdateUser; } set { _UpdateUser = value; } }

        public ACME_TASK_File(int ID, string PrjCode, string ObjType, string DocNo, string FileName, string CreateDate, string
      CreateTime, string CreateUser, string UpdateDate, string UpdateTime, string UpdateUser)
        {
            _ID = ID;
            _PrjCode = PrjCode;
            _ObjType = ObjType;
            _DocNo = DocNo;
            _FileName = FileName;
            _CreateDate = CreateDate;
            _CreateTime = CreateTime;
            _CreateUser = CreateUser;
            _UpdateDate = UpdateDate;
            _UpdateTime = UpdateTime;
            _UpdateUser = UpdateUser;
        }
        public ACME_TASK_File()
        {
        }
        // ACME_TASK_File Insert
        public static void AddACME_TASK_File(ACME_TASK_File row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into ACME_TASK_File(PrjCode,ObjType,DocNo,FileName,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser) values(@PrjCode,@ObjType,@DocNo,@FileName,@CreateDate,@CreateTime,@CreateUser,@UpdateDate,@UpdateTime,@UpdateUser)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
            command.Parameters["@PrjCode"].Value = row.PrjCode;
            if (String.IsNullOrEmpty(row.PrjCode))
            {
                command.Parameters["@PrjCode"].IsNullable = true;
                command.Parameters["@PrjCode"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@ObjType", SqlDbType.VarChar, 20, "ObjType"));
            command.Parameters["@ObjType"].Value = row.ObjType;
            if (String.IsNullOrEmpty(row.ObjType))
            {
                command.Parameters["@ObjType"].IsNullable = true;
                command.Parameters["@ObjType"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar, 20, "DocNo"));
            command.Parameters["@DocNo"].Value = row.DocNo;
            if (String.IsNullOrEmpty(row.DocNo))
            {
                command.Parameters["@DocNo"].IsNullable = true;
                command.Parameters["@DocNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar, 250, "FileName"));
            command.Parameters["@FileName"].Value = row.FileName;
            if (String.IsNullOrEmpty(row.FileName))
            {
                command.Parameters["@FileName"].IsNullable = true;
                command.Parameters["@FileName"].Value = "";
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

        // ACME_TASK_File Update
        public static void UpdateACME_TASK_File(ACME_TASK_File row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE ACME_TASK_File SET PrjCode = @PrjCode,ObjType = @ObjType,DocNo = @DocNo,FileName = @FileName,CreateDate = @CreateDate,CreateTime = @CreateTime,CreateUser = @CreateUser,UpdateDate = @UpdateDate,UpdateTime = @UpdateTime,UpdateUser = @UpdateUser WHERE ID=@ID";
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
            command.Parameters.Add(new SqlParameter("@ObjType", SqlDbType.VarChar, 20, "ObjType"));
            command.Parameters["@ObjType"].Value = row.ObjType;
            if (String.IsNullOrEmpty(row.ObjType))
            {
                command.Parameters["@ObjType"].IsNullable = true;
                command.Parameters["@ObjType"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar, 20, "DocNo"));
            command.Parameters["@DocNo"].Value = row.DocNo;
            if (String.IsNullOrEmpty(row.DocNo))
            {
                command.Parameters["@DocNo"].IsNullable = true;
                command.Parameters["@DocNo"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar, 250, "FileName"));
            command.Parameters["@FileName"].Value = row.FileName;
            if (String.IsNullOrEmpty(row.FileName))
            {
                command.Parameters["@FileName"].IsNullable = true;
                command.Parameters["@FileName"].Value = "";
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

        // ACME_TASK_File Delete
        public static void DeleteACME_TASK_File(ACME_TASK_File row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE ACME_TASK_File WHERE ID=@ID";
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

        // ACME_TASK_File Select
        public static DataTable GetACME_TASK_File(ACME_TASK_File row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,ObjType,DocNo,FileName,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_File WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }

        // ACME_TASK_File Select
        public static DataTable GetACME_TASK_File(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,ObjType,DocNo,FileName,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_File WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        // Condition 版本
        public static DataTable GetACME_TASK_File_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,PrjCode,ObjType,DocNo,FileName,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_File WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }

    }
}
 using System;
 using System.Data;
 using System.Configuration;

 using System.Data.SqlClient;
 
 /// <summary>
 /// Summary description for ACME_TASK_Q
 /// 作者:
 /// </summary>
// ACME_TASK_Q 資料結構
public class ACME_TASK_Q
{
    public static string AcmesqlSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
    private int _ID;
    private string _PrjCode;
    private string _Title;
    private string _Subject;
    private string _Solution;
    private string _StartDate;
    private string _EndDate;
    private string _AcDate;
    private string _CreateDate;
    private string _CreateTime;
    private string _CreateUser;
    private string _UpdateDate;
    private string _UpdateTime;
    private string _UpdateUser;

    public int ID { get { return _ID; } set { _ID = value; } }
    public string PrjCode { get { return _PrjCode; } set { _PrjCode = value; } }
    public string Title { get { return _Title; } set { _Title = value; } }
    public string Subject { get { return _Subject; } set { _Subject = value; } }
    public string Solution { get { return _Solution; } set { _Solution = value; } }
    public string StartDate { get { return _StartDate; } set { _StartDate = value; } }
    public string EndDate { get { return _EndDate; } set { _EndDate = value; } }
    public string AcDate { get { return _AcDate; } set { _AcDate = value; } }
    public string CreateDate { get { return _CreateDate; } set { _CreateDate = value; } }
    public string CreateTime { get { return _CreateTime; } set { _CreateTime = value; } }
    public string CreateUser { get { return _CreateUser; } set { _CreateUser = value; } }
    public string UpdateDate { get { return _UpdateDate; } set { _UpdateDate = value; } }
    public string UpdateTime { get { return _UpdateTime; } set { _UpdateTime = value; } }
    public string UpdateUser { get { return _UpdateUser; } set { _UpdateUser = value; } }

    public ACME_TASK_Q(int ID, string Title, string Subject, string Solution, string StartDate, string EndDate, string AcDate, string
  CreateDate, string CreateTime, string CreateUser, string UpdateDate, string UpdateTime, string UpdateUser)
    {
        _ID = ID;
        _Title = Title;
        _Subject = Subject;
        _Solution = Solution;
        _StartDate = StartDate;
        _EndDate = EndDate;
        _AcDate = AcDate;
        _CreateDate = CreateDate;
        _CreateTime = CreateTime;
        _CreateUser = CreateUser;
        _UpdateDate = UpdateDate;
        _UpdateTime = UpdateTime;
        _UpdateUser = UpdateUser;
    }
    public ACME_TASK_Q()
    {
    }
    // ACME_TASK_Q Insert
    public static void AddACME_TASK_Q(ACME_TASK_Q row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        SqlCommand command = new SqlCommand("Insert into ACME_TASK_Q(PrjCode,Title,Subject,Solution,StartDate,EndDate,AcDate,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser) values(@PrjCode,@Title,@Subject,@Solution,@StartDate,@EndDate,@AcDate,@CreateDate,@CreateTime,@CreateUser,@UpdateDate,@UpdateTime,@UpdateUser)", connection);
        command.CommandType = CommandType.Text;


        command.Parameters.Add(new SqlParameter("@PrjCode", SqlDbType.VarChar, 50, "PrjCode"));
        command.Parameters["@PrjCode"].Value = row.PrjCode;
        if (String.IsNullOrEmpty(row.PrjCode))
        {
            command.Parameters["@PrjCode"].IsNullable = true;
            command.Parameters["@PrjCode"].Value = "";
        }

        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 50, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Subject", SqlDbType.VarChar, 200, "Subject"));
        command.Parameters["@Subject"].Value = row.Subject;
        if (String.IsNullOrEmpty(row.Subject))
        {
            command.Parameters["@Subject"].IsNullable = true;
            command.Parameters["@Subject"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Solution", SqlDbType.VarChar, 200, "Solution"));
        command.Parameters["@Solution"].Value = row.Solution;
        if (String.IsNullOrEmpty(row.Solution))
        {
            command.Parameters["@Solution"].IsNullable = true;
            command.Parameters["@Solution"].Value = "";
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

    // ACME_TASK_Q Update
    public static void UpdateACME_TASK_Q(ACME_TASK_Q row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "UPDATE ACME_TASK_Q SET PrjCode = @PrjCode,Title = @Title,Subject = @Subject,Solution = @Solution,StartDate = @StartDate,EndDate = @EndDate,AcDate = @AcDate,CreateDate = @CreateDate,CreateTime = @CreateTime,CreateUser = @CreateUser,UpdateDate = @UpdateDate,UpdateTime = @UpdateTime,UpdateUser = @UpdateUser WHERE ID=@ID";
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



        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 50, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Subject", SqlDbType.VarChar, 200, "Subject"));
        command.Parameters["@Subject"].Value = row.Subject;
        if (String.IsNullOrEmpty(row.Subject))
        {
            command.Parameters["@Subject"].IsNullable = true;
            command.Parameters["@Subject"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Solution", SqlDbType.VarChar, 200, "Solution"));
        command.Parameters["@Solution"].Value = row.Solution;
        if (String.IsNullOrEmpty(row.Solution))
        {
            command.Parameters["@Solution"].IsNullable = true;
            command.Parameters["@Solution"].Value = "";
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

    // ACME_TASK_Q Delete
    public static void DeleteACME_TASK_Q(ACME_TASK_Q row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "DELETE ACME_TASK_Q WHERE ID=@ID";
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

    // ACME_TASK_Q Select
    public static DataTable GetACME_TASK_Q(ACME_TASK_Q row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,PrjCode,Title,Subject,Solution,StartDate,EndDate,AcDate,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Q WHERE 1= 1  AND ID=@ID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ID", row.ID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Q");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Q"];
    }

    // ACME_TASK_Q Select
    public static DataTable GetACME_TASK_Q(int ID)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,PrjCode,Title,Subject,Solution,StartDate,EndDate,AcDate,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Q WHERE 1= 1  AND ID=@ID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ID", ID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Q");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Q"];
    }
    // Condition 版本
    public static DataTable GetACME_TASK_Q_Condition(string Condition)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,PrjCode,Title,Subject,Solution,StartDate,EndDate,AcDate,CreateDate,CreateTime,CreateUser,UpdateDate,UpdateTime,UpdateUser FROM ACME_TASK_Q WHERE 1= 1 ";
        sql += Condition;
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_Q");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_Q"];
    }

}

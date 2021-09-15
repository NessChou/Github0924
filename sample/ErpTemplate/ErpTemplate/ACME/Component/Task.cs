using System;
using System.Data;
using System.Configuration;
using System.Web;

using System.Data.SqlClient;

/// <summary>
/// Summary description for Task
/// 作者:
/// </summary>
// Task 資料結構
public class Task
{

    public static string AcmesqlSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";

    private int _TaskID;
    private int _ParentTaskID;
    private string _Title;
    private int _Duration;
    private string _DurationType;
    private DateTime _StartDate;
    private DateTime _EndDate;
    private int _PercentComplete;
    private DateTime _DueDate;
    private string _Status;
    private int _CategoryID;
    private string _Priority;
    private string _ColorCategory;
    private string _Flag;

    public int TaskID { get { return _TaskID; } set { _TaskID = value; } }
    public int ParentTaskID { get { return _ParentTaskID; } set { _ParentTaskID = value; } }
    public string Title { get { return _Title; } set { _Title = value; } }
    public int Duration { get { return _Duration; } set { _Duration = value; } }
    public string DurationType { get { return _DurationType; } set { _DurationType = value; } }
    public DateTime StartDate { get { return _StartDate; } set { _StartDate = value; } }
    public DateTime EndDate { get { return _EndDate; } set { _EndDate = value; } }
    public int PercentComplete { get { return _PercentComplete; } set { _PercentComplete = value; } }
    public DateTime DueDate { get { return _DueDate; } set { _DueDate = value; } }
    public string Status { get { return _Status; } set { _Status = value; } }
    public int CategoryID { get { return _CategoryID; } set { _CategoryID = value; } }
    public string Priority { get { return _Priority; } set { _Priority = value; } }
    public string ColorCategory { get { return _ColorCategory; } set { _ColorCategory = value; } }
    public string Flag { get { return _Flag; } set { _Flag = value; } }

    public Task(int TaskID, int ParentTaskID, string Title, int Duration, string DurationType, DateTime StartDate, DateTime EndDate, int
  PercentComplete, DateTime DueDate, string Status, int CategoryID, string Priority, string ColorCategory, string Flag)
    {
        _TaskID = TaskID;
        _ParentTaskID = ParentTaskID;
        _Title = Title;
        _Duration = Duration;
        _DurationType = DurationType;
        _StartDate = StartDate;
        _EndDate = EndDate;
        _PercentComplete = PercentComplete;
        _DueDate = DueDate;
        _Status = Status;
        _CategoryID = CategoryID;
        _Priority = Priority;
        _ColorCategory = ColorCategory;
        _Flag = Flag;
    }
    public Task()
    {
    }
    // Task Insert
    public static void AddTask(Task row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        SqlCommand command = new SqlCommand("Insert into Task(TaskID,ParentTaskID,Title,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag) values(@TaskID,@ParentTaskID,@Title,@Duration,@DurationType,@StartDate,@EndDate,@PercentComplete,@DueDate,@Status,@CategoryID,@Priority,@ColorCategory,@Flag)", connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@ParentTaskID", row.ParentTaskID));
        command.Parameters.Add(new SqlParameter("@Title", row.Title));
        command.Parameters.Add(new SqlParameter("@Duration", row.Duration));
        command.Parameters.Add(new SqlParameter("@DurationType", SqlDbType.VarChar, 50, "DurationType"));
        command.Parameters["@DurationType"].Value = row.DurationType;
        if (String.IsNullOrEmpty(row.DurationType))
        {
            command.Parameters["@DurationType"].IsNullable = true;
            command.Parameters["@DurationType"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate", row.StartDate));
        command.Parameters.Add(new SqlParameter("@EndDate", row.EndDate));
        command.Parameters.Add(new SqlParameter("@PercentComplete", row.PercentComplete));
        command.Parameters.Add(new SqlParameter("@DueDate", row.DueDate));
        command.Parameters.Add(new SqlParameter("@Status", SqlDbType.VarChar, 50, "Status"));
        command.Parameters["@Status"].Value = row.Status;
        if (String.IsNullOrEmpty(row.Status))
        {
            command.Parameters["@Status"].IsNullable = true;
            command.Parameters["@Status"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CategoryID", row.CategoryID));
        command.Parameters.Add(new SqlParameter("@Priority", SqlDbType.VarChar, 50, "Priority"));
        command.Parameters["@Priority"].Value = row.Priority;
        if (String.IsNullOrEmpty(row.Priority))
        {
            command.Parameters["@Priority"].IsNullable = true;
            command.Parameters["@Priority"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ColorCategory", SqlDbType.VarChar, 50, "ColorCategory"));
        command.Parameters["@ColorCategory"].Value = row.ColorCategory;
        if (String.IsNullOrEmpty(row.ColorCategory))
        {
            command.Parameters["@ColorCategory"].IsNullable = true;
            command.Parameters["@ColorCategory"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Flag", SqlDbType.VarChar, 50, "Flag"));
        command.Parameters["@Flag"].Value = row.Flag;
        if (String.IsNullOrEmpty(row.Flag))
        {
            command.Parameters["@Flag"].IsNullable = true;
            command.Parameters["@Flag"].Value = "";
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

    // Task Update
    public static void UpdateTask(Task row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "UPDATE Task SET ParentTaskID = @ParentTaskID,Title = @Title,Duration = @Duration,DurationType = @DurationType,StartDate = @StartDate,EndDate = @EndDate,PercentComplete = @PercentComplete,DueDate = @DueDate,Status = @Status,CategoryID = @CategoryID,Priority = @Priority,ColorCategory = @ColorCategory,Flag = @Flag WHERE TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@ParentTaskID", row.ParentTaskID));
        command.Parameters.Add(new SqlParameter("@Title", row.Title));
        command.Parameters.Add(new SqlParameter("@Duration", row.Duration));
        command.Parameters.Add(new SqlParameter("@DurationType", SqlDbType.VarChar, 50, "DurationType"));
        command.Parameters["@DurationType"].Value = row.DurationType;
        if (String.IsNullOrEmpty(row.DurationType))
        {
            command.Parameters["@DurationType"].IsNullable = true;
            command.Parameters["@DurationType"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@StartDate", row.StartDate));
        command.Parameters.Add(new SqlParameter("@EndDate", row.EndDate));
        command.Parameters.Add(new SqlParameter("@PercentComplete", row.PercentComplete));
        command.Parameters.Add(new SqlParameter("@DueDate", row.DueDate));
        command.Parameters.Add(new SqlParameter("@Status", SqlDbType.VarChar, 50, "Status"));
        command.Parameters["@Status"].Value = row.Status;
        if (String.IsNullOrEmpty(row.Status))
        {
            command.Parameters["@Status"].IsNullable = true;
            command.Parameters["@Status"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@CategoryID", row.CategoryID));
        command.Parameters.Add(new SqlParameter("@Priority", SqlDbType.VarChar, 50, "Priority"));
        command.Parameters["@Priority"].Value = row.Priority;
        if (String.IsNullOrEmpty(row.Priority))
        {
            command.Parameters["@Priority"].IsNullable = true;
            command.Parameters["@Priority"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@ColorCategory", SqlDbType.VarChar, 50, "ColorCategory"));
        command.Parameters["@ColorCategory"].Value = row.ColorCategory;
        if (String.IsNullOrEmpty(row.ColorCategory))
        {
            command.Parameters["@ColorCategory"].IsNullable = true;
            command.Parameters["@ColorCategory"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Flag", SqlDbType.VarChar, 50, "Flag"));
        command.Parameters["@Flag"].Value = row.Flag;
        if (String.IsNullOrEmpty(row.Flag))
        {
            command.Parameters["@Flag"].IsNullable = true;
            command.Parameters["@Flag"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@wTaskID", row.TaskID));
        
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

    // Task Delete
    public static void DeleteTask(Task row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "DELETE Task WHERE TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
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

    // Task Select
    public static DataTable GetTask(Task row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT TaskID,ParentTaskID,Title,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM Task WHERE 1= 1  AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "Task");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["Task"];
    }

    // Task Select
    public static DataTable GetTask(int TaskID)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT TaskID,ParentTaskID,Title,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM Task WHERE 1= 1  AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TaskID", TaskID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "Task");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["Task"];
    }
    // Condition 版本
    public static DataTable GetTask_Condition(string Condition)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT TaskID,ParentTaskID,Title,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM Task WHERE 1= 1 ";
        sql += Condition;
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "Task");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["Task"];
    }

}

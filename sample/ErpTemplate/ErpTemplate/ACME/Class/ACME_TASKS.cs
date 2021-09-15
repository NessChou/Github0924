 using System;
 using System.Data;
 using System.Configuration;

 using System.Data.SqlClient;
 
 /// <summary>
 /// Summary description for ACME_TASKS
 /// 作者:
 /// </summary>
// ACME_TASKS 資料結構
public class ACME_TASKS
{
    public static string AcmesqlSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
    private string _ProjectID;
    private int _TaskID;
    private int _ParentTaskID;
    private string _Title;
    private string _Description;
    private int _Duration;
    private string _DurationType;
    private string _StartDate;
    private string _EndDate;
    private int _PercentComplete;
    private string _DueDate;
    private string _Status;
    private int _CategoryID;
    private string _Priority;
    private string _ColorCategory;
    private string _Flag;

    public string ProjectID { get { return _ProjectID; } set { _ProjectID = value; } }
    public int TaskID { get { return _TaskID; } set { _TaskID = value; } }
    public int ParentTaskID { get { return _ParentTaskID; } set { _ParentTaskID = value; } }
    public string Title { get { return _Title; } set { _Title = value; } }
    public string Description { get { return _Description; } set { _Description = value; } }
    public int Duration { get { return _Duration; } set { _Duration = value; } }
    public string DurationType { get { return _DurationType; } set { _DurationType = value; } }
    public string StartDate { get { return _StartDate; } set { _StartDate = value; } }
    public string EndDate { get { return _EndDate; } set { _EndDate = value; } }
    public int PercentComplete { get { return _PercentComplete; } set { _PercentComplete = value; } }
    public string DueDate { get { return _DueDate; } set { _DueDate = value; } }
    public string Status { get { return _Status; } set { _Status = value; } }
    public int CategoryID { get { return _CategoryID; } set { _CategoryID = value; } }
    public string Priority { get { return _Priority; } set { _Priority = value; } }
    public string ColorCategory { get { return _ColorCategory; } set { _ColorCategory = value; } }
    public string Flag { get { return _Flag; } set { _Flag = value; } }

    public ACME_TASKS(string ProjectID, int TaskID, int ParentTaskID, string Title, string Description, int Duration, string
  DurationType, string StartDate, string EndDate, int PercentComplete, string DueDate, string Status, int CategoryID, string Priority, string
  ColorCategory, string Flag)
    {
        _ProjectID = ProjectID;
        _TaskID = TaskID;
        _ParentTaskID = ParentTaskID;
        _Title = Title;
        _Description = Description;
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
    public ACME_TASKS()
    {
    }
    // ACME_TASKS Insert
    public static void AddACME_TASKS(ACME_TASKS row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        SqlCommand command = new SqlCommand("Insert into ACME_TASKS(ProjectID,TaskID,ParentTaskID,Title,Description,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag) values(@ProjectID,@TaskID,@ParentTaskID,@Title,@Description,@Duration,@DurationType,@StartDate,@EndDate,@PercentComplete,@DueDate,@Status,@CategoryID,@Priority,@ColorCategory,@Flag)", connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ProjectID", SqlDbType.VarChar, 20, "ProjectID"));
        command.Parameters["@ProjectID"].Value = row.ProjectID;
        if (String.IsNullOrEmpty(row.ProjectID))
        {
            command.Parameters["@ProjectID"].IsNullable = true;
            command.Parameters["@ProjectID"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@ParentTaskID", row.ParentTaskID));
        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 250, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Description", SqlDbType.VarChar, 250, "Description"));
        command.Parameters["@Description"].Value = row.Description;
        if (String.IsNullOrEmpty(row.Description))
        {
            command.Parameters["@Description"].IsNullable = true;
            command.Parameters["@Description"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Duration", row.Duration));
        command.Parameters.Add(new SqlParameter("@DurationType", SqlDbType.VarChar, 50, "DurationType"));
        command.Parameters["@DurationType"].Value = row.DurationType;
        if (String.IsNullOrEmpty(row.DurationType))
        {
            command.Parameters["@DurationType"].IsNullable = true;
            command.Parameters["@DurationType"].Value = "";
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
        command.Parameters.Add(new SqlParameter("@PercentComplete", row.PercentComplete));
        command.Parameters.Add(new SqlParameter("@DueDate", SqlDbType.VarChar, 8, "DueDate"));
        command.Parameters["@DueDate"].Value = row.DueDate;
        if (String.IsNullOrEmpty(row.DueDate))
        {
            command.Parameters["@DueDate"].IsNullable = true;
            command.Parameters["@DueDate"].Value = "";
        }
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

    // ACME_TASKS Update
    public static void UpdateACME_TASKS(ACME_TASKS row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "UPDATE ACME_TASKS SET ParentTaskID = @ParentTaskID,Title = @Title,Description = @Description,Duration = @Duration,DurationType = @DurationType,StartDate = @StartDate,EndDate = @EndDate,PercentComplete = @PercentComplete,DueDate = @DueDate,Status = @Status,CategoryID = @CategoryID,Priority = @Priority,ColorCategory = @ColorCategory,Flag = @Flag WHERE ProjectID=@ProjectID AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ProjectID", SqlDbType.VarChar, 20, "ProjectID"));
        command.Parameters["@ProjectID"].Value = row.ProjectID;
        if (String.IsNullOrEmpty(row.ProjectID))
        {
            command.Parameters["@ProjectID"].IsNullable = true;
            command.Parameters["@ProjectID"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@ParentTaskID", row.ParentTaskID));
        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 250, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Description", SqlDbType.VarChar, 250, "Description"));
        command.Parameters["@Description"].Value = row.Description;
        if (String.IsNullOrEmpty(row.Description))
        {
            command.Parameters["@Description"].IsNullable = true;
            command.Parameters["@Description"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@Duration", row.Duration));
        command.Parameters.Add(new SqlParameter("@DurationType", SqlDbType.VarChar, 50, "DurationType"));
        command.Parameters["@DurationType"].Value = row.DurationType;
        if (String.IsNullOrEmpty(row.DurationType))
        {
            command.Parameters["@DurationType"].IsNullable = true;
            command.Parameters["@DurationType"].Value = "";
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
        command.Parameters.Add(new SqlParameter("@PercentComplete", row.PercentComplete));
        command.Parameters.Add(new SqlParameter("@DueDate", SqlDbType.VarChar, 8, "DueDate"));
        command.Parameters["@DueDate"].Value = row.DueDate;
        if (String.IsNullOrEmpty(row.DueDate))
        {
            command.Parameters["@DueDate"].IsNullable = true;
            command.Parameters["@DueDate"].Value = "";
        }
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

    // ACME_TASKS Delete
    public static void DeleteACME_TASKS(ACME_TASKS row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "DELETE ACME_TASKS WHERE ProjectID=@ProjectID AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ProjectID", row.ProjectID));
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

    // ACME_TASKS Select
    public static DataTable GetACME_TASKS(ACME_TASKS row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ProjectID,TaskID,ParentTaskID,Title,Description,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM ACME_TASKS WHERE 1= 1  AND ProjectID=@ProjectID AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ProjectID", row.ProjectID));
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASKS");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASKS"];
    }

    // ACME_TASKS Select
    public static DataTable GetACME_TASKS(string ProjectID, int TaskID)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ProjectID,TaskID,ParentTaskID,Title,Description,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM ACME_TASKS WHERE 1= 1  AND ProjectID=@ProjectID AND TaskID=@TaskID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));
        command.Parameters.Add(new SqlParameter("@TaskID", TaskID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASKS");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASKS"];
    }
    // Condition 版本
    public static DataTable GetACME_TASKS_Condition(string Condition)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ProjectID,TaskID,ParentTaskID,Title,Description,Duration,DurationType,StartDate,EndDate,PercentComplete,DueDate,Status,CategoryID,Priority,ColorCategory,Flag FROM ACME_TASKS WHERE 1= 1 ";
        sql += Condition;
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASKS");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASKS"];
    }

}

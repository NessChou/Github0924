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
 /// Summary description for ACME_TASK_TL
 /// 作者:
 /// </summary>
// ACME_TASK_TL 資料結構
public class ACME_TASK_TL
{
    public static string AcmesqlSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
    private int _ID;
    private int _TpID;
    private int _TaskID;
    private string _Title;
    private int _SortID;
    private int _ParentID;
    private int _Ddays;

    public int ID { get { return _ID; } set { _ID = value; } }
    public int TpID { get { return _TpID; } set { _TpID = value; } }
    public int TaskID { get { return _TaskID; } set { _TaskID = value; } }
    public string Title { get { return _Title; } set { _Title = value; } }
    public int SortID { get { return _SortID; } set { _SortID = value; } }
    public int ParentID { get { return _ParentID; } set { _ParentID = value; } }
    public int Ddays { get { return _Ddays; } set { _Ddays = value; } }

    public ACME_TASK_TL(int ID, int TpID, int TaskID, string Title, int SortID, int ParentID, int Ddays)
    {
        _ID = ID;
        _TpID = TpID;
        _TaskID = TaskID;
        _Title = Title;
        _SortID = SortID;
        _ParentID = ParentID;
        _Ddays = Ddays;
    }
    public ACME_TASK_TL()
    {
    }
    // ACME_TASK_TL Insert
    public static void AddACME_TASK_TL(ACME_TASK_TL row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        SqlCommand command = new SqlCommand("Insert into ACME_TASK_TL(TpID,TaskID,Title,SortID,ParentID,Ddays) values(@TpID,@TaskID,@Title,@SortID,@ParentID,@Ddays)", connection);
        command.CommandType = CommandType.Text;
     
        command.Parameters.Add(new SqlParameter("@TpID", row.TpID));
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 50, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SortID", row.SortID));
        command.Parameters.Add(new SqlParameter("@ParentID", row.ParentID));
        command.Parameters.Add(new SqlParameter("@Ddays", row.Ddays));
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

    // ACME_TASK_TL Update
    public static void UpdateACME_TASK_TL(ACME_TASK_TL row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "UPDATE ACME_TASK_TL SET TpID = @TpID,TaskID = @TaskID,Title = @Title,SortID = @SortID,ParentID = @ParentID,Ddays = @Ddays WHERE ID=@ID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ID", row.ID));
        command.Parameters.Add(new SqlParameter("@TpID", row.TpID));
        command.Parameters.Add(new SqlParameter("@TaskID", row.TaskID));
        command.Parameters.Add(new SqlParameter("@Title", SqlDbType.VarChar, 50, "Title"));
        command.Parameters["@Title"].Value = row.Title;
        if (String.IsNullOrEmpty(row.Title))
        {
            command.Parameters["@Title"].IsNullable = true;
            command.Parameters["@Title"].Value = "";
        }
        command.Parameters.Add(new SqlParameter("@SortID", row.SortID));
        command.Parameters.Add(new SqlParameter("@ParentID", row.ParentID));
        command.Parameters.Add(new SqlParameter("@Ddays", row.Ddays));
        
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

    // ACME_TASK_TL Delete
    public static void DeleteACME_TASK_TL(ACME_TASK_TL row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "DELETE ACME_TASK_TL WHERE ID=@ID";
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

    // ACME_TASK_TL Select
    public static DataTable GetACME_TASK_TL(ACME_TASK_TL row)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,TpID,TaskID,Title,SortID,ParentID,Ddays FROM ACME_TASK_TL WHERE 1= 1  AND ID=@ID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ID", row.ID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_TL");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_TL"];
    }

    // ACME_TASK_TL Select
    public static DataTable GetACME_TASK_TL(int ID)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,TpID,TaskID,Title,SortID,ParentID,Ddays FROM ACME_TASK_TL WHERE 1= 1  AND ID=@ID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@ID", ID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_TL");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_TL"];
    }


    public static DataTable GetACME_TASK_TL_TpID(int TpID)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,TpID,TaskID,Title,SortID,ParentID,Ddays FROM ACME_TASK_TL WHERE  TpID=@TpID";
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        command.Parameters.Add(new SqlParameter("@TpID", TpID));
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_TL");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_TL"];
    }


    // Condition 版本
    public static DataTable GetACME_TASK_TL_Condition(string Condition)
    {
        SqlConnection connection = new SqlConnection(AcmesqlSP);
        string sql = "SELECT ID,TpID,TaskID,Title,SortID,ParentID,Ddays FROM ACME_TASK_TL WHERE 1= 1 ";
        sql += Condition;
        SqlCommand command = new SqlCommand(sql, connection);
        command.CommandType = CommandType.Text;
        SqlDataAdapter da = new SqlDataAdapter(command);
        DataSet ds = new DataSet();
        try
        {
            connection.Open();
            da.Fill(ds, "ACME_TASK_TL");
        }
        finally
        {
            connection.Close();
        }
        return ds.Tables["ACME_TASK_TL"];
    }

}

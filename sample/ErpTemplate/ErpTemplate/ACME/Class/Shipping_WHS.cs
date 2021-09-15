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
 /// Summary description for Shipping_WHS
 /// 作者:
 /// </summary>
// Shipping_WHS 資料結構
namespace ACME
{
    public class Shipping_WHS
    {
        private int _ID;
        private string _NUM;
        private string _WHSCODE;
        private string _WHS;
        private string _DESCRIPTION;
        private string _LOCATION;
        public int ID { get { return _ID; } set { _ID = value; } }
        public string NUM { get { return _NUM; } set { _NUM = value; } }
        public string WHS { get { return _WHS; } set { _WHS = value; } }
        public string WHSCODE { get { return _WHSCODE; } set { _WHSCODE = value; } }
        public string DESCRIPTION { get { return _DESCRIPTION; } set { _DESCRIPTION = value; } }
        public string LOCATION { get { return _LOCATION; } set { _LOCATION = value; } }
        public Shipping_WHS(int ID, string NUM, string WHS, string WHSCODE, string DESCRIPTION, string LOCATION)
        {
            _ID = ID;
            _NUM = NUM;
            _WHS = WHS;
            _WHSCODE = WHSCODE;
            _DESCRIPTION = DESCRIPTION;
            _LOCATION = LOCATION;
        }
        public Shipping_WHS()
        {
        }
        // Shipping_WHS Insert
        public static void AddShipping_WHS(Shipping_WHS row)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into Shipping_WHS(NUM,WHS,WHSCODE,DESCRIPTION,LOCATION) values(@NUM,@WHS,@WHSCODE,@DESCRIPTION,@LOCATION)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@NUM", SqlDbType.VarChar, 10, "NUM"));
            command.Parameters["@NUM"].Value = row.NUM;
            if (String.IsNullOrEmpty(row.NUM))
            {
                command.Parameters["@NUM"].IsNullable = true;
                command.Parameters["@NUM"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WHS", SqlDbType.NVarChar, 50, "WHS"));
            command.Parameters["@WHS"].Value = row.WHS;
            if (String.IsNullOrEmpty(row.WHSCODE))
            {
                command.Parameters["@WHS"].IsNullable = true;
                command.Parameters["@WHS"].Value = "";
            }

            command.Parameters.Add(new SqlParameter("@WHSCODE", SqlDbType.NVarChar, 100, "WHSCODE"));
            command.Parameters["@WHSCODE"].Value = row.WHSCODE;
            if (String.IsNullOrEmpty(row.WHSCODE))
            {
                command.Parameters["@WHSCODE"].IsNullable = true;
                command.Parameters["@WHSCODE"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 500, "DESCRIPTION"));
            command.Parameters["@DESCRIPTION"].Value = row.DESCRIPTION;
            if (String.IsNullOrEmpty(row.DESCRIPTION))
            {
                command.Parameters["@DESCRIPTION"].IsNullable = true;
                command.Parameters["@DESCRIPTION"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@LOCATION", SqlDbType.NVarChar, 500, "LOCATION"));
            command.Parameters["@LOCATION"].Value = row.LOCATION;
            if (String.IsNullOrEmpty(row.DESCRIPTION))
            {
                command.Parameters["@LOCATION"].IsNullable = true;
                command.Parameters["@LOCATION"].Value = "";
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

        // Shipping_WHS Update
        public static void UpdateShipping_WHS(Shipping_WHS row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "UPDATE Shipping_WHS SET NUM = @NUM,WHS = @WHS,WHSCODE = @WHSCODE,DESCRIPTION = @DESCRIPTION,LOCATION=@LOCATION WHERE ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            command.Parameters.Add(new SqlParameter("@NUM", SqlDbType.VarChar, 10, "NUM"));
            command.Parameters["@NUM"].Value = row.NUM;
            if (String.IsNullOrEmpty(row.NUM))
            {
                command.Parameters["@NUM"].IsNullable = true;
                command.Parameters["@NUM"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WHS", SqlDbType.NVarChar, 100, "50"));
            command.Parameters["@WHS"].Value = row.WHS;
            if (String.IsNullOrEmpty(row.WHS))
            {
                command.Parameters["@WHS"].IsNullable = true;
                command.Parameters["@WHS"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@WHSCODE", SqlDbType.NVarChar, 100, "WHSCODE"));
            command.Parameters["@WHSCODE"].Value = row.WHSCODE;
            if (String.IsNullOrEmpty(row.WHSCODE))
            {
                command.Parameters["@WHSCODE"].IsNullable = true;
                command.Parameters["@WHSCODE"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@DESCRIPTION", SqlDbType.NVarChar, 500, "DESCRIPTION"));
            command.Parameters["@DESCRIPTION"].Value = row.DESCRIPTION;
            if (String.IsNullOrEmpty(row.DESCRIPTION))
            {
                command.Parameters["@DESCRIPTION"].IsNullable = true;
                command.Parameters["@DESCRIPTION"].Value = "";
            }
            command.Parameters.Add(new SqlParameter("@LOCATION", SqlDbType.NVarChar, 500, "LOCATION"));
            command.Parameters["@LOCATION"].Value = row.LOCATION;
            if (String.IsNullOrEmpty(row.DESCRIPTION))
            {
                command.Parameters["@LOCATION"].IsNullable = true;
                command.Parameters["@LOCATION"].Value = "";
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

        // Shipping_WHS Delete
        public static void DeleteShipping_WHS(Shipping_WHS row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "DELETE Shipping_WHS WHERE ID=@ID";
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

        // Shipping_WHS Select
        public static DataTable GetShipping_WHS(Shipping_WHS row)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,NUM,WHS,WHSCODE,DESCRIPTION,LOCATION FROM Shipping_WHS WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", row.ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Shipping_WHS");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Shipping_WHS"];
        }

        // Shipping_WHS Select
        public static DataTable GetShipping_WHS(int ID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,NUM,WHS,WHSCODE,DESCRIPTION,LOCATION FROM Shipping_WHS WHERE 1= 1  AND ID=@ID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Shipping_WHS");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Shipping_WHS"];
        }
        // Condition 版本
        public static DataTable GetShipping_WHS_Condition(string Condition)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ID,NUM,WHS,WHSCODE,DESCRIPTION,LOCATION FROM Shipping_WHS WHERE 1= 1 ";
            sql += Condition;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Shipping_WHS");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Shipping_WHS"];
        }

    }
}
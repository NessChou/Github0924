using System;
using System.Collections.Generic;
using System.Text;
using System.IO;//
using System.Data;//
using System.Data.SqlClient; //SQL Server使用必要引入Namespace


namespace My
{
    public class MyDatabase
    {

        SqlConnection conn;
        SqlCommand cmd;
        SqlDataReader dr;
        SqlDataAdapter da;
        string errorMsg;

        string ConnString = My.MyGlobal.SQLConnectionString;

        #region 開啟DataTable的方法宣告

        /// <summary>
        /// 開啟DataTable的方法宣告
        /// </summary>
        /// <param name="DefaultUseDB">傳入使用的資料庫名稱有Access或SQLServer</param>
        /// <param name="SQLstr">SQL語法敘述</param>
        /// <param name="TableName">表格名稱</param>
        /// <returns></returns>
        public DataTable OpenDataTable(string DefaultUseDB, string SQLstr, string TableName)
        {

            DataSet ds = new DataSet();
            DataTable bufDataTable = new DataTable();


            conn = new SqlConnection(ConnString);
            conn.Open();
            da = new SqlDataAdapter(SQLstr, conn);
            da.Fill(ds, TableName);
            bufDataTable = ds.Tables[TableName];
            conn.Close();
            return bufDataTable;

        }

        #endregion


        #region 建立一個DataView的方法

        /// <summary>
        /// 建立一個DataView的方法
        /// </summary>
        /// <param name="DefaultUseDB">傳入使用的資料庫名稱有SQLServer</param>
        /// <param name="SQLstr">SQL語法敘述</param>
        /// <param name="TableName">表格名稱</param>
        /// <returns></returns>
        public DataView CreateDataView(string SQLstr, string TableName)
        {
            try
            {
                DataSet ds = new DataSet();
                DataView DVbuf = new DataView();

                conn = new SqlConnection(ConnString);
                conn.Open();
                da = new SqlDataAdapter(SQLstr, conn);
                da.Fill(ds, TableName);
                DVbuf = ds.Tables[TableName].DefaultView;
                conn.Close();
                return DVbuf;
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
                return null;
            }

            
           
        }

        #endregion


        #region 驗證主鍵值是否存在

        /// <summary>
        /// 驗證主鍵值是否存在
        /// </summary>
        /// <param name="PKval">要判斷此主鍵值是否存在</param>
        /// <param name="PKname">資料庫的主鍵值欄位名稱</param>
        /// <param name="TableName">資料表名稱</param>
        /// <returns></returns>
        public bool AuthPK(string PKval, string PKname, string TableName)
        {
           
            string selectCmd;
            string errorMsg;


            try
            {
                selectCmd = "SELECT * FROM " + TableName + " WHERE " + PKname + " ='" + PKval + "'";

                    conn = new SqlConnection(ConnString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        conn.Close();
                        return true;
                    }
                    else
                    {
                        conn.Close();
                        return false;
                    }

            }
            catch (Exception ex)
            {

                errorMsg = ex.Message;
                return false;
            }

        }

        #endregion


        #region "驗證雙主鍵值"

        /// <summary>
        /// 驗證雙主鍵值
        /// </summary>
        /// <param name="PKval1">第一個主鍵值</param>
        /// <param name="PKname1">第一個主鍵名稱</param>
        /// <param name="PKval2">第二個主鍵值</param>
        /// <param name="PKname2">第二個主鍵名稱</param>
        /// <param name="TableName"></param>
        /// <returns></returns>
        public bool AuthPK(string PKval1, string PKname1, string PKval2, string PKname2, string TableName)
        {
            
            string selectCmd;
            string errorMsg;


            try
            {
                selectCmd = "SELECT * FROM " + TableName + " WHERE " + PKname1 + " ='" + PKval1 + "' And " + PKname2 + "='" + PKval2 + "'";


                    conn = new SqlConnection(ConnString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        conn.Close();
                        return true;
                    }
                    else
                    {
                        conn.Close();
                        return false;
                    }

               

            }
            catch (Exception ex)
            {

                errorMsg = ex.Message;
                return false;
            }

        }

        #endregion


        #region 獲取資料庫表格欄位資料

        /// <summary>
        /// 獲取資料庫表格欄位資料
        /// </summary>
        /// <param name="TableName">資料表名稱</param>
        /// <param name="PKName">資料庫的主鍵值欄位名稱</param>
        /// <param name="PKValue">主鍵值</param>
        /// <param name="GetFieldName">要獲取的欄位名稱</param>
        /// <returns></returns>
        public string GetTableFieldData(string TableName, string PKName, string PKValue, string GetFieldName)
        {
           
            string selectCmd;
            string errorMsg;
            string bufstr;


            try
            {
                selectCmd = "SELECT " + GetFieldName + " FROM " + TableName + " WHERE " + PKName + "='" + PKValue + "'";

                    conn = new SqlConnection(ConnString);
                    conn.Open();
                    cmd = new SqlCommand(selectCmd, conn);
                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        bufstr = dr[GetFieldName].ToString();
                        conn.Close();
                        dr.Close();
                        cmd.Dispose();
                        return bufstr;

                    }
                    else
                    {
                        conn.Close();
                        dr.Close();
                        cmd.Dispose();
                        return "false";
                    }

            }
            catch (Exception ex)
            {

                errorMsg = ex.Message;
                return "false";
            }
        }

        #endregion


    }
}

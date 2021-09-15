using System;
using System.Collections.Generic;
using System.Text;
using System.IO;//
using System.Data;//
using System.Data.SqlClient; //SQL Server�ϥΥ��n�ޤJNamespace


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

        #region �}��DataTable����k�ŧi

        /// <summary>
        /// �}��DataTable����k�ŧi
        /// </summary>
        /// <param name="DefaultUseDB">�ǤJ�ϥΪ���Ʈw�W�٦�Access��SQLServer</param>
        /// <param name="SQLstr">SQL�y�k�ԭz</param>
        /// <param name="TableName">���W��</param>
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


        #region �إߤ@��DataView����k

        /// <summary>
        /// �إߤ@��DataView����k
        /// </summary>
        /// <param name="DefaultUseDB">�ǤJ�ϥΪ���Ʈw�W�٦�SQLServer</param>
        /// <param name="SQLstr">SQL�y�k�ԭz</param>
        /// <param name="TableName">���W��</param>
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


        #region ���ҥD��ȬO�_�s�b

        /// <summary>
        /// ���ҥD��ȬO�_�s�b
        /// </summary>
        /// <param name="PKval">�n�P�_���D��ȬO�_�s�b</param>
        /// <param name="PKname">��Ʈw���D������W��</param>
        /// <param name="TableName">��ƪ�W��</param>
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


        #region "�������D���"

        /// <summary>
        /// �������D���
        /// </summary>
        /// <param name="PKval1">�Ĥ@�ӥD���</param>
        /// <param name="PKname1">�Ĥ@�ӥD��W��</param>
        /// <param name="PKval2">�ĤG�ӥD���</param>
        /// <param name="PKname2">�ĤG�ӥD��W��</param>
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


        #region �����Ʈw��������

        /// <summary>
        /// �����Ʈw��������
        /// </summary>
        /// <param name="TableName">��ƪ�W��</param>
        /// <param name="PKName">��Ʈw���D������W��</param>
        /// <param name="PKValue">�D���</param>
        /// <param name="GetFieldName">�n��������W��</param>
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

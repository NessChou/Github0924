using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace ACME
{
    class util
    {

        public static string GetAutoNumber(SqlConnection connection,
             string NumberName)
        {

            int returnVal = 0;

            string sql = "SELECT CURRENTDIGITS FROM AUTONUM WHERE NUMBERNAME=@NUMBERNAME";

            SqlCommand command = new SqlCommand(sql, connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));


            //  int returnVal= (Int32)cmd.ExecuteScalar();

            try
            {
                SqlDataReader reader=null;
                if (connection.State == ConnectionState.Closed)
                  connection.Open();
                try
                {

                  reader = command.ExecuteReader();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //這種判定方法是有效果的
                if (reader.Read() == false)
                {

                    string InserSql = "Insert into AUTONUM(NUMBERNAME,CURRENTPREFIX,DIGITSWIDTH,CURRENTDIGITS,VALUEINTERVAL,MINVALUE,MAXVALUE)" +
                     "values(@NUMBERNAME,@CURRENTPREFIX,@DIGITSWIDTH,@CURRENTDIGITS,@VALUEINTERVAL,@MINVALUE,@MAXVALUE)";
                    command.CommandType = CommandType.Text;
                    reader.Close();//要先關閉 command.Dispose();

                    command.CommandText = InserSql;
                    command.Parameters.Clear();
                    command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));
                    command.Parameters.Add(new SqlParameter("@CURRENTPREFIX", ""));
                    command.Parameters.Add(new SqlParameter("@DIGITSWIDTH", 4));
                    command.Parameters.Add(new SqlParameter("@CURRENTDIGITS", 1));
                    command.Parameters.Add(new SqlParameter("@VALUEINTERVAL", 1));
                    command.Parameters.Add(new SqlParameter("@MINVALUE", 1));
                    command.Parameters.Add(new SqlParameter("@MAXVALUE", 9999));


                    command.ExecuteNonQuery();
                    returnVal = 1;
                }
                else
                {
                    string UpdateSql = "UPDATE AUTONUM " +
                       " Set CURRENTDIGITS=@CURRENTDIGITS " +
                       " Where NUMBERNAME=@NUMBERNAME and CURRENTDIGITS=@O_CURRENTDIGITS";

                    int CurrentVal = Convert.ToInt32(reader["CURRENTDIGITS"]);

                    command.CommandType = CommandType.Text;
                    reader.Close();//要先關閉 command.Dispose();

                    command.CommandText = UpdateSql;
                    command.Parameters.Clear();
                    command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));
                    command.Parameters.Add(new SqlParameter("@CURRENTDIGITS", CurrentVal + 1));
                    command.Parameters.Add(new SqlParameter("@O_CURRENTDIGITS", CurrentVal));

                    command.ExecuteNonQuery();
                    returnVal = CurrentVal + 1;
                }


            }
            finally
            {
                connection.Close();
            }

            return returnVal.ToString("000");

        }
    public static string GetSeqNo(int length, DataGridView gridview)
        {

            int iRecs;
        //  DataGridView gridview;
            iRecs = gridview.Rows.Count;

            string zeroLen = string.Empty;



            string s = "0000000000" + Convert.ToString(iRecs);

            return s.Substring(s.Length - length, length);
        }
        public static string GetAutoNumber1(SqlConnection connection,
         string NumberName)
        {

            int returnVal = 0;

            string sql = "SELECT CURRENTDIGITS FROM AUTONUM WHERE NUMBERNAME=@NUMBERNAME";

            SqlCommand command = new SqlCommand(sql, connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));


            //  int returnVal= (Int32)cmd.ExecuteScalar();

            try
            {
                SqlDataReader reader = null;
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                try
                {

                    reader = command.ExecuteReader();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //這種判定方法是有效果的
                if (reader.Read() == false)
                {

                    string InserSql = "Insert into AUTONUM(NUMBERNAME,CURRENTPREFIX,DIGITSWIDTH,CURRENTDIGITS,VALUEINTERVAL,MINVALUE,MAXVALUE)" +
                     "values(@NUMBERNAME,@CURRENTPREFIX,@DIGITSWIDTH,@CURRENTDIGITS,@VALUEINTERVAL,@MINVALUE,@MAXVALUE)";
                    command.CommandType = CommandType.Text;
                    reader.Close();//要先關閉 command.Dispose();

                    command.CommandText = InserSql;
                    command.Parameters.Clear();
                    command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));
                    command.Parameters.Add(new SqlParameter("@CURRENTPREFIX", ""));
                    command.Parameters.Add(new SqlParameter("@DIGITSWIDTH", 4));
                    command.Parameters.Add(new SqlParameter("@CURRENTDIGITS", 1));
                    command.Parameters.Add(new SqlParameter("@VALUEINTERVAL", 1));
                    command.Parameters.Add(new SqlParameter("@MINVALUE", 1));
                    command.Parameters.Add(new SqlParameter("@MAXVALUE", 9999));


                    command.ExecuteNonQuery();
                    returnVal = 1;
                }
                else
                {
                    string UpdateSql = "UPDATE AUTONUM " +
                       " Set CURRENTDIGITS=@CURRENTDIGITS " +
                       " Where NUMBERNAME=@NUMBERNAME and CURRENTDIGITS=@O_CURRENTDIGITS";

                    int CurrentVal = Convert.ToInt32(reader["CURRENTDIGITS"]);

                    command.CommandType = CommandType.Text;
                    reader.Close();//要先關閉 command.Dispose();

                    command.CommandText = UpdateSql;
                    command.Parameters.Clear();
                    command.Parameters.Add(new SqlParameter("@NUMBERNAME", NumberName));
                    command.Parameters.Add(new SqlParameter("@CURRENTDIGITS", CurrentVal + 1));
                    command.Parameters.Add(new SqlParameter("@O_CURRENTDIGITS", CurrentVal));

                    command.ExecuteNonQuery();
                    returnVal = CurrentVal + 1;
                }


            }
            finally
            {
                connection.Close();
            }

            return returnVal.ToString("00");

        }

        public static DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }

    }
}

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
using System.Reflection;
using System.Text.RegularExpressions;
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
    
            iRecs = gridview.Rows.Count;
            string zeroLen = string.Empty;
            string s = "0000000000" + Convert.ToString(iRecs);
            return s.Substring(s.Length - length, length);
        }
    public static decimal  CINT(string GS)
    {

        string g = GS;

        int n;
        if (!int.TryParse(g, out n))
        {
            g = "0";
        }

        return Convert.ToDecimal(g);
    }
    public static decimal CINT2(string GS)
    {

        string g = GS;

        int n;
        if (!int.TryParse(g, out n))
        {
            g = "1";
        }

        return Convert.ToDecimal(g);
    }
    public static decimal CINT3(string GS)
    {

        string g = GS;

        decimal  n;
        if (!decimal.TryParse(g, out n))
        {
            g = "0";
        }

        return Convert.ToDecimal(g);
    }
        public static string quarter(string length)
        {
    
            string g = length.Substring(4, 2);
            int q = 0;
            if (g == "01" || g == "02" || g == "03")
            {
                q = 1;
            }
            else if (g == "04" || g == "05" || g == "06")
            {
                q = 2;
            }
            else if (g == "07" || g == "08" || g == "09")
            {
                q = 3;
            }
            else if (g == "10" || g == "11" || g == "12")
            {
                q = 4;
            }
            return q.ToString();
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
        public static string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }

        public static string SPACE(string SPACE)
        {
            string trim = SPACE.Replace(" ", "");
            trim = trim.Replace("/r", "");
            trim = trim.Replace("/n", "");
            trim = trim.Replace("/t", "");
            SPACE = Regex.Replace(trim, @"\s", "");
            return SPACE;
        }
        public static string CARDNAME(string CARDNAME)
        {
            if (CARDNAME != "")
            {
                Regex rex = new Regex(@"^[A-Za-z0-9]+$");
                string ENG = CARDNAME.Substring(1, 1);
                Match ma = rex.Match(ENG);
                if (ma.Success)
                {
                    int t1 = CARDNAME.IndexOf(" ");
                    if (t1 != -1)
                    {
                        CARDNAME = CARDNAME.Substring(0, t1);
                    }
                }
                else
                {
                    CARDNAME = CARDNAME.Substring(0, 4);
                }

            }
            return CARDNAME;
        }
        public static SqlDataAdapter GetAdapter(object tableAdapter)
        {

            Type tableAdapterType = tableAdapter.GetType();

            SqlDataAdapter adapter = (SqlDataAdapter)tableAdapterType.GetProperty("Adapter", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(tableAdapter, null);

            return adapter;

        }


        public static string INVOTYPE(string F)
        {
            string s="";
            if (F == "0")
            {
                s = "0-三聯式發票/電子計算機發票";
            }
            else if (F == "1")
            {
                s = "1-三聯式收銀機發票/電子發票";
            }
            else if (F == "2")
            {
                s = "2-有稅憑證";
            }
            else if (F == "3")
            {
                s = "3-海關代徵稅";
            }
            else if (F == "4")
            {
                s = "4-免用統一發票/收據";
            }
            else if (F == "5")
            {
                s = "5-退折";
            }
            int iRecs;


            return s;
        }
        public static string EXDD(string MONTH)
        {
            string MM = "";
            if (MONTH == "1")
            {
                MM = "JAN";
            }
            if (MONTH == "2")
            {
                MM = "FEB";
            }
            if (MONTH == "3")
            {
                MM = "MAR";
            }
            if (MONTH == "4")
            {
                MM = "APR";
            }
            if (MONTH == "5")
            {
                MM = "MAY";
            }
            if (MONTH == "6")
            {
                MM = "JUN";
            }
            if (MONTH == "7")
            {
                MM = "JUL";
            }
            if (MONTH == "8")
            {
                MM = "AUG";
            }
            if (MONTH == "9")
            {
                MM = "SEP";
            }
            if (MONTH == "10")
            {
                MM = "OCT";
            }
            if (MONTH == "11")
            {
                MM = "NOV";
            }
            if (MONTH == "12")
            {
                MM = "DEC";
            }
            return MM;
        }
        public static string EXCEL(int F)
        {
            string s = "";
            if (F == 1)
            {
                s = "A";
            }
            else if (F == 2)
            {
                s = "B";
            }
            else if (F == 3)
            {
                s = "C";
            }
            else if (F == 4)
            {
                s = "D";
            }
            else if (F == 5)
            {
                s = "E";
            }
            else if (F == 6)
            {
                s = "F";
            }
            else if (F == 7)
            {
                s = "G";
            }
            else if (F == 8)
            {
                s = "H";
            }
            else if (F == 9)
            {
                s = "I";
            }
            else if (F == 10)
            {
                s = "J";
            }
            else if (F == 11)
            {
                s = "K";
            }
            else if (F == 12)
            {
                s = "L";
            }
            else if (F == 13)
            {
                s = "M";
            }
            else if (F == 14)
            {
                s = "N";
            }
            else if (F == 15)
            {
                s = "O";
            }
            else if (F == 16)
            {
                s = "P";
            }

            else if (F == 17)
            {
                s = "Q";
            }
            else if (F == 18)
            {
                s = "R";
            }
            else if (F == 19)
            {
                s = "S";
            }
            else if (F == 20)
            {
                s = "T";
            }
            return s;
        }
        public static string EXCEL2(int F)
        {
            string s = "";
            if (F == 1)
            {
                s = "A";
            }
            else if (F == 2)
            {
                s = "B";
            }
            else if (F == 3)
            {
                s = "C";
            }
            else if (F == 4)
            {
                s = "D";
            }
            else if (F == 5)
            {
                s = "E";
            }
            else if (F == 6)
            {
                s = "F";
            }
            else if (F == 7)
            {
                s = "G";
            }
            else if (F == 8)
            {
                s = "H";
            }
            else if (F == 9)
            {
                s = "I";
            }
            else if (F == 10)
            {
                s = "J";
            }
            else if (F == 11)
            {
                s = "K";
            }
            else if (F == 12)
            {
                s = "L";
            }
            else if (F == 13)
            {
                s = "M";
            }
            else if (F == 14)
            {
                s = "N";
            }
            else if (F == 15)
            {
                s = "O";
            }
            else if (F == 16)
            {
                s = "P";
            }

            else if (F == 17)
            {
                s = "Q";
            }
            else if (F == 18)
            {
                s = "R";
            }
            else if (F == 19)
            {
                s = "S";
            }
            else if (F == 20)
            {
                s = "T";
            }
            return s;
        }
        public static void AddOBTD(decimal LocTotal, int UserSign)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[OBTD](BatchNum,Status,NumOfTrans,DateID,LocTotal,FcTotal,SysTotal,UserSign) values(@BatchNum,@Status,@NumOfTrans,@DateID,@LocTotal,0,@SysTotal,@UserSign)", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM2().Rows[0][0].ToString();

            command.Parameters.Add(new SqlParameter("@BatchNum", T1));
            command.Parameters.Add(new SqlParameter("@Status", "O"));
            command.Parameters.Add(new SqlParameter("@NumOfTrans", 1));
            command.Parameters.Add(new SqlParameter("@DateID", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@LocTotal", LocTotal));

            command.Parameters.Add(new SqlParameter("@SysTotal", LocTotal));
            command.Parameters.Add(new SqlParameter("@UserSign", UserSign));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }


        public static void AddOBTF(string Ref2, decimal LocTotal, int FinncPriod, int USERSIGN, string U_SATT)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[OBTF](BatchNum,TransId,BtfStatus,TransType,RefDate,Memo,Ref1,Ref2,LocTotal,FcTotal,SysTotal,TransCode,TransRate,BtfLine,Project,DueDate,TaxDate,PCAddition,FinncPriod,DataSource,CreateDate,UserSign,UserSign2,RefndRprt,LogInstanc,ObjType,AdjTran,RevSource,AutoStorno,Corisptivi,StampTax,Series,AutoVAT,ReportEU,Report347,DocType,AttNum,GenRegNo,Printed,U_SATT) values(@BatchNum,@TransId,@BtfStatus,@TransType,@RefDate,@Memo,@Ref1,@Ref2,@LocTotal,0,@SysTotal,@TransCode,0,@BtfLine,@Project,@DueDate,@TaxDate,@PCAddition,@FinncPriod,@DataSource,@CreateDate,@UserSign,@UserSign2,@RefndRprt,0,@ObjType,@AdjTran,@RevSource,@AutoStorno,@Corisptivi,@StampTax,@Series,@AutoVAT,@ReportEU,@Report347,@DocType,0,@GenRegNo,@Printed,@U_SATT)", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM2().Rows[0][0].ToString();

            command.Parameters.Add(new SqlParameter("@BatchNum", T1));
            command.Parameters.Add(new SqlParameter("@TransId", 1));
            command.Parameters.Add(new SqlParameter("@BtfStatus", "O"));
            command.Parameters.Add(new SqlParameter("@TransType", -1));
            command.Parameters.Add(new SqlParameter("@RefDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@Memo", ""));
            command.Parameters.Add(new SqlParameter("@Ref1", ""));
            command.Parameters.Add(new SqlParameter("@Ref2", Ref2));
            command.Parameters.Add(new SqlParameter("@LocTotal", LocTotal));
            command.Parameters.Add(new SqlParameter("@SysTotal", LocTotal));
            command.Parameters.Add(new SqlParameter("@TransCode", ""));
            command.Parameters.Add(new SqlParameter("@BtfLine", 1));
            command.Parameters.Add(new SqlParameter("@Project", ""));
            command.Parameters.Add(new SqlParameter("@DueDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@TaxDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@PCAddition", "N"));
            command.Parameters.Add(new SqlParameter("@FinncPriod", FinncPriod));
            command.Parameters.Add(new SqlParameter("@DataSource", "I"));
            command.Parameters.Add(new SqlParameter("@CreateDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@UserSign", USERSIGN));
            command.Parameters.Add(new SqlParameter("@UserSign2", USERSIGN));
            command.Parameters.Add(new SqlParameter("@RefndRprt", "N"));
            command.Parameters.Add(new SqlParameter("@ObjType", "30"));
            command.Parameters.Add(new SqlParameter("@AdjTran", "N"));
            command.Parameters.Add(new SqlParameter("@RevSource", "N"));
            command.Parameters.Add(new SqlParameter("@AutoStorno", "N"));
            command.Parameters.Add(new SqlParameter("@Corisptivi", "N"));
            command.Parameters.Add(new SqlParameter("@StampTax", "N"));
            command.Parameters.Add(new SqlParameter("@Series", 14));
            command.Parameters.Add(new SqlParameter("@AutoVAT", "N"));
            command.Parameters.Add(new SqlParameter("@ReportEU", "N"));
            command.Parameters.Add(new SqlParameter("@Report347", "N"));
            command.Parameters.Add(new SqlParameter("@DocType", "00"));
            command.Parameters.Add(new SqlParameter("@GenRegNo", "N"));
            command.Parameters.Add(new SqlParameter("@Printed", "N"));
            command.Parameters.Add(new SqlParameter("@U_SATT", U_SATT));
            
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        public static  void AddBTF1(int Line_ID, string Account, decimal Debit, decimal Credit, string LineMemo, string TransType, string Ref2, string Project, string ProfitCode, int USERSIGN, int FinncPriod, string VatGroup, string VatLine, string DebCred)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[BTF1](TransId,Line_ID,Account,Debit,Credit,SYSCred,SYSDeb,FCDebit,FCCredit,FCCurrency,DueDate,ShortName,IntrnMatch,ExtrMatch,LineMemo,TransType,RefDate,Ref1,Ref2,Project,ProfitCode,TaxDate,SystemRate,ToMthSum,UserSign,BatchNum,FinncPriod,RelTransId,RelLineID,RelType,LogInstanc,VatGroup,BaseSum,VatRate,Indicator,AdjTran,RevSource,ObjType,SYSBaseSum,MultMatch,VatLine,VatAmount,SYSVatSum,Closed,GrossValue,LineType,DebCred,SequenceNr,MIEntry,MIVEntry,ClsInTP,CenVatCom,MatType,PstngType) values(@TransId,@Line_ID,@Account,@Debit,@Credit,@SYSCred,@SYSDeb,0,0,@FCCurrency,@DueDate,@ShortName,0,0,@LineMemo,@TransType,@RefDate,@Ref1,@Ref2,@Project,@ProfitCode,@TaxDate,0,0,@UserSign,@BatchNum,@FinncPriod,@RelTransId,@RelLineID,@RelType,0,@VatGroup,0,0,@Indicator,@AdjTran,@RevSource,@ObjType,0,0,@VatLine,0,0,@Closed,0,0,@DebCred,0,0,0,0,@CenVatCom,@MatType,0)", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM2().Rows[0][0].ToString();

            command.Parameters.Add(new SqlParameter("@TransId", 1));
            command.Parameters.Add(new SqlParameter("@Line_ID", Line_ID));
            command.Parameters.Add(new SqlParameter("@Account", Account));
            command.Parameters.Add(new SqlParameter("@Debit", Debit));
            command.Parameters.Add(new SqlParameter("@Credit", Credit));
            command.Parameters.Add(new SqlParameter("@SYSCred", Credit));
            command.Parameters.Add(new SqlParameter("@SYSDeb", Debit));
            command.Parameters.Add(new SqlParameter("@FCCurrency", ""));
            command.Parameters.Add(new SqlParameter("@DueDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@ShortName", Account));
            command.Parameters.Add(new SqlParameter("@LineMemo", LineMemo));
            command.Parameters.Add(new SqlParameter("@TransType", TransType));
            command.Parameters.Add(new SqlParameter("@RefDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@Ref1", ""));
            command.Parameters.Add(new SqlParameter("@Ref2", Ref2));
            command.Parameters.Add(new SqlParameter("@Project", Project));
            command.Parameters.Add(new SqlParameter("@ProfitCode", ProfitCode));
            command.Parameters.Add(new SqlParameter("@TaxDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@UserSign", USERSIGN));
            command.Parameters.Add(new SqlParameter("@BatchNum", T1));
            command.Parameters.Add(new SqlParameter("@FinncPriod", FinncPriod));
            command.Parameters.Add(new SqlParameter("@RelTransId", -1));
            command.Parameters.Add(new SqlParameter("@RelLineID", -1));
            command.Parameters.Add(new SqlParameter("@RelType", "N")); ;
            command.Parameters.Add(new SqlParameter("@VatGroup", VatGroup));
            command.Parameters.Add(new SqlParameter("@Indicator", ""));
            command.Parameters.Add(new SqlParameter("@AdjTran", "N"));
            command.Parameters.Add(new SqlParameter("@RevSource", "N"));
            command.Parameters.Add(new SqlParameter("@ObjType", 30));
            command.Parameters.Add(new SqlParameter("@VatLine", VatLine));
            command.Parameters.Add(new SqlParameter("@Closed", "N"));
            command.Parameters.Add(new SqlParameter("@DebCred", DebCred));
            command.Parameters.Add(new SqlParameter("@CenVatCom", -1));
            command.Parameters.Add(new SqlParameter("@MatType", -1));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        public static  System.Data.DataTable GetONNM2()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT AUTOKEY NUM,AUTOKEY NUM1 FROM ONNM WHERE OBJECTCODE='28'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static void ADDONNM()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE ONNM SET AUTOKEY=AUTOKEY+1 WHERE OBJECTCODE='28' ", connection);
            command.CommandType = CommandType.Text;

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        public static  System.Data.DataTable GETPACLS(string PLT)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CAST(SUBSTRING('" + PLT + "',CHARINDEX('-', '" + PLT + "')+1,2) AS INT)-CAST(SUBSTRING('" + PLT + "',CHARINDEX('PALLET', '" + PLT + "')+6,CHARINDEX('-', '" + PLT + "')-CHARINDEX('PALLET', '" + PLT + "')-6) AS INT)+1 PLT ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static  System.Data.DataTable GETPACL3(string InvoiceNo)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ltrim(substring(SayTotal,CHARINDEX('(', SayTotal)+1,CHARINDEX(')', SayTotal)-1-CHARINDEX('(', SayTotal))) PLT,SayTotal FROM rpa_packingH WHERE  InvoiceNo =@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL3B(string InvoiceNo)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ltrim(substring(SayTotal,0,CHARINDEX('PLT',SAYTOTAL))) PLT,REPLACE(ltrim(substring(SayTotal,CHARINDEX('(', SayTotal)+1,CHARINDEX(')', SayTotal)-1-CHARINDEX('(', SayTotal))),'CTNS','') CARTON  FROM rpa_packingH WHERE  InvoiceNo =@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL3B2(string InvoiceNo)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(PLT AS INT)) PLT,SUM(CAST(CARTON AS INT)) CARTON FROM rpa_packingD WHERE  InvoiceNo =@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static void DELPACK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete SHIPPING_PACK where users=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        public static System.Data.DataTable GetSHIPPACK9(string SHIPPINGCODE, string PLATENO)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SUM(CAST(CARTONNO AS INT)) FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND PLATENO2=@PLATENO GROUP BY PLATENO2  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLATENO", PLATENO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetWHPACK(string SHIPPINGCODE, string BLC, string CHE, string SB, string CAR)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (CHE == "TRUE")
            {
                sb.Append(" SELECT *,CAST(GW AS DECIMAL(10,2)) GW2 FROM WH_PACK2 WHERE SHIPPINGCODE IN (" + SB + "  ) ");
            }
            else
            {
                sb.Append(" SELECT *,CAST(GW AS DECIMAL(10,2)) GW2 FROM WH_PACK2 WHERE SHIPPINGCODE =@SHIPPINGCODE ");
            }
            if (!String.IsNullOrEmpty(BLC))
            {
                sb.Append(" AND BLC =@BLC ");
            }

            if (!String.IsNullOrEmpty(CAR))
            {
                sb.Append(" AND FLAG1 IN (" + CAR + "  ) ");
            }

            sb.Append("    ORDER BY SHIPPINGCODE,ID");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BLC", BLC));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }



        public static System.Data.DataTable GetSHIPOITM(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_ITEMNAME+' '+U_MODEL MODEL,' ('+CASE WHEN T1.U_GRADE='NN' THEN 'N' ELSE T1.U_GRADE END+' GRADE)' GRADE,T0.ES,T1.U_MODEL TMODEL   FROM SHIPPING_PACK T0 ");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" where users=@USERS  AND T0.ITEMCODE=@ITEMCODE   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static  System.Data.DataTable GetSHIPOITMES()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.ES   FROM SHIPPING_PACK T0 ");
            //    sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" where users=@USERS  AND ISNULL(T0.ES,'') <> ''   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            // command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPS3(string ShippingCode, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND ITEMCODE=@ITEMCODE  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACKQTY(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT QTY FROM SHIPPING_PACK where users=@USERS  AND ITEMCODE=@ITEMCODE ORDER BY CAST(QTY AS INT) DESC   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACK4(string ShippingCode, string MODEL)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND INDescription LIKE '%" + MODEL + "%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }


        public static System.Data.DataTable GetSHIPPACK5(string SER)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUM(CAST(CARTONQTY AS INT))　QTY, SUM(CAST(GW AS DECIMAL(10,2)))　GW, CAST(SUM(CAST(NW AS DECIMAL(10,4))) AS DECIMAL(10,3))　NW   FROM SHIPPING_PACK　WHERE SER=@SER AND USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACK6(string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  PAL_NW  FROM ACMESQL02.DBO.OITM  T1 ");
            sb.Append(" INNER JOIN CART T2 ON (T1.U_TMODEL=T2.MODEL_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS");
            sb.Append("  AND T1.U_VERSION =T2.MODEL_Ver COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE T1.ITEMCODE=@ITEMCODE  AND T2.PAL_QTY =@QTY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public static System.Data.DataTable GetSHIPPACK7(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  PAL_NW,PAL_QTY  FROM ACMESQL02.DBO.OITM  T1  ");
            sb.Append(" INNER JOIN CART T2 ON (T1.U_TMODEL=T2.MODEL_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS ");
            sb.Append(" AND T1.U_VERSION =T2.MODEL_Ver COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" WHERE T1.ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public static  System.Data.DataTable GetSHIPPACK()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT *,CASE LOACTION WHEN N'中国' THEN 'CHINA' ELSE LOACTION END  LOCATION FROM SHIPPING_PACK where users=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACK3(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT U_MODEL,U_GRADE,U_TMODEL,U_VERSION FROM OITM　WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACK2(string SER)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CAST(MIN(CAST(PLATENO AS INT)) AS VARCHAR) +'-'+CAST(MAX(CAST(PLATENO AS INT)) AS VARCHAR) PLATENO,CAST(MIN(CAST(SUBSTRING(CARTONNO2,1, (CASE CHARINDEX('~', CARTONNO2) WHEN 0 THEN 10 ELSE CHARINDEX('~', CARTONNO2)  END) -1) AS INT)) AS VARCHAR)+'~'+CAST(MAX(CAST(SUBSTRING(CARTONNO2, CHARINDEX('~', CARTONNO2)+1,5) AS INT)) AS VARCHAR) 　CARTONNO  FROM SHIPPING_PACK WHERE SER=@SER AND　USERS =@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static  System.Data.DataTable GetSHIPPACK4O(string ShippingCode, string MODEL, string GRADE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT  INDescription  　 FROM INVOICED T0 LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE　COLLATE  Chinese_Taiwan_Stroke_CI_AS) 　");
            sb.Append(" WHERE ShippingCode = @ShippingCode　AND INDescription LIKE '%" + MODEL + "%'  AND  CASE WHEN T1.U_GRADE='NN' THEN 'N' ELSE T1.U_GRADE END=@GRADE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static  System.Data.DataTable GetSHIPPS4(string ShippingCode, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select INDescription  from shipping_item T0");
            sb.Append(" LEFT JOIN InvoiceD T1 ON (T0.ShippingCode =T1.ShippingCode AND T0.linenum=T1.LINENUM)");
            sb.Append("  where T0.shippingcode=@shippingcode AND T0.ItemCode =@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetSHIPPACK4SQTY(string ShippingCode, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND ITEMCODE=@ITEMCODE AND INQTY=@QTY  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
     
        public static System.Data.DataTable GetSHIPPACK4S(string ShippingCode, string ITEMCODE, string ITEMNAME)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT INDescription  　 FROM INVOICED 　WHERE ShippingCode = @ShippingCode　AND ITEMCODE=@ITEMCODE  ");
            if (ITEMCODE == "ACMERMA01.RMA01" && ITEMNAME.Length > 4)
            {
                string IM = ITEMNAME.Substring(3, 2);
                sb.Append(" AND INDescription  LIKE '%" + IM + "%' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static void AddPACK(string PLATENO, string CARTONNO, string ITEMCODE, string QTY, string CARTONQTY, string NW, string GW, string L, string W, string H, string LOACTION, string SER, string CARTONNO2, string INVOICE, string ITEMNAME, string WHNO, string ES)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" Insert into SHIPPING_PACK(PLATENO,CARTONNO,ITEMCODE,QTY,CARTONQTY,NW,GW,L,W,H,LOACTION,USERS,SER,CARTONNO2,INVOICE,ITEMNAME,WHNO,ES) values(@PLATENO,@CARTONNO,@ITEMCODE,@QTY,@CARTONQTY,@NW,@GW,@L,@W,@H,@LOACTION,@USERS,@SER,@CARTONNO2,@INVOICE,@ITEMNAME,@WHNO,@ES)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PLATENO", PLATENO));
            command.Parameters.Add(new SqlParameter("@CARTONNO", CARTONNO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARTONQTY", CARTONQTY));
            command.Parameters.Add(new SqlParameter("@NW", NW));
            command.Parameters.Add(new SqlParameter("@GW", GW));
            command.Parameters.Add(new SqlParameter("@L", L));
            command.Parameters.Add(new SqlParameter("@W", W));
            command.Parameters.Add(new SqlParameter("@H", H));
            command.Parameters.Add(new SqlParameter("@LOACTION", LOACTION));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@CARTONNO2", CARTONNO2));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));
            command.Parameters.Add(new SqlParameter("@ES", ES));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public static  System.Data.DataTable GetSHIPPACK4S2(string SHIPPINGCODE, string ITEMNAME, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT ITEMCODE  FROM WH_PACK2  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  )  AND ITEMNAME=@ITEMNAME AND ITEMCODE <>@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static void UPPACKS(string SER)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE SHIPPING_PACK  SET SER=@SER  WHERE ID=(SELECT MAX(ID) FROM SHIPPING_PACK WHERE USERS=@USERS) ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public static System.Data.DataTable GETPACL2F(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
        //sb.Append("               select  T0.ID,T2.ID ID2,T2.ItemNo PARTNO,T0.InvoiceNo INV,T2.Qty  QTY,T2.NWeight  NW,GWeight  GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,REPLACE(T2.CBM,'*','x') CBM,REPLACE(REPLACE(T2.CartonNo,'PALLET',''),' ','')  PLT,T2.CARTON2,T2.PLT PLT2,CAST(REPLACE(REPLACE(T2.Qty,'@',''),',','') AS DECIMAL)   QTY2  from rpa_packingH T0  ");
        //sb.Append("               LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo) ");
        //sb.Append("               LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
        //sb.Append(" WHERE  T0.InvoiceNo=@InvoiceNo");

            sb.Append(" select  T0.ID,T2.ID ID2,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty,',','')    QTY,T2.NWeight2  NW,GWeight2  GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,REPLACE(T2.CBM,'*','x') CBM,");
            sb.Append(" T2.PLT2 PLT,T2.CARTON2,T2.PLT PLT2,REPLACE(T2.Qty2,',','')   QTY2  from rpa_packingH T0   ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)  ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append(" WHERE  T0.InvoiceNo=@InvoiceNo AND QTY <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GETPACL2F222(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   select  T0.ID,T2.ID ID2,T2.ItemNo PARTNO,T0.InvoiceNo INV,T2.Qty  QTY,CASE WHEN T2.Qty LIKE '%@%' THEN '@'+REPLACE(T2.NWeight,'@','') ELSE T2.NWeight  END  NW,CASE WHEN T2.Qty LIKE '%@%' THEN '@'+REPLACE(T2.GWeight,'@','') ELSE T2.GWeight  END   GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,REPLACE(T2.CBM,'*','x') CBM,T2.PLT2 PLT,T2.CARTON2,T2.PLT PLT2,CASE WHEN ISNULL(T2.QTY2,'') ='' THEN T2.Qty ELSE  REPLACE(T2.Qty2,',','')   END QTY2,T2.CARTON   from rpa_packingH T0    ");
      
            sb.Append("               LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo) ");
            sb.Append(" WHERE  T0.InvoiceNo=@InvoiceNo  AND  T2.CARTON <> ''    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GETPACL2F222LB(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.ID,T2.ID ID2,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty,',','')  QTY,REPLACE(T2.Qty2,',','')  QTY2,T2.NWeight2  NW,GWeight2  GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,'@'+REPLACE(T2.CBM,'*','x')  CBM,T2.PLT2  PLT,T2.CARTON,T2.CARTON2,T2.PLT PLT2  from rpa_packingH T0    ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            sb.Append(" WHERE  T0.InvoiceNo =@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACLAP(string U_Shipping_no)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("         select DISTINCT  T2.ItemNo PARTNO,T0.InvoiceNo INV,TotalQty QTY,TotalNW NW,TOTALGW GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,T0.CBM,T0.PLT,T0.CARTON,CBMM  from rpa_packingH T0  ");
            sb.Append("               LEFT JOIN rpa_packingD T2 ON (T0.ID=T2.ID) ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("               WHERE T1.U_Shipping_no  = @U_Shipping_no ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_Shipping_no", U_Shipping_no));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GETTEST(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MARK FROM RPA_PackingH  WHERE INVOICENO=@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETTEST2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT *  FROM Mark WHERE SHIPPINGCODE=@SHIPPINGCODE AND MARK LIKE '%TAIWAN%'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL2F2(string ID)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty,',','')  QTY,'@'+REPLACE(T2.NWeight,'@','') NW,'@'+REPLACE(T2.GWeight,'@','')  GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,'@'+REPLACE(T2.CBM,'*','x')  CBM,T2.PLT2 PLT,T2.CARTON,T2.CARTON2,T2.PLT PLT2 from rpa_packingH T0    ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            sb.Append(" WHERE  T2.ID=@ID ");
            sb.Append(" UNION ALL ");
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty2,',','')  QTY,T2.NWeight2  NW,GWeight2  GW,'SAY TOTAL:'+SAYTOTAL SAYTOTAL,'','',T2.CARTON,'',T2.PLT PLT2  from rpa_packingH T0    ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            sb.Append(" WHERE  T2.ID=@ID ");
            //sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,");
            //sb.Append(" '@'+CAST(CAST(REPLACE(T2.Qty,'@','') AS INT)*(CAST(T2.CARTON AS INT)/CAST(T2.PLT AS INT))  AS VARCHAR) QTY");
            //sb.Append(" ,'@'+CAST(CAST(REPLACE(T2.NWeight,'@','') AS DECIMAL(18,2))*(CAST(T2.CARTON AS INT)/CAST(T2.PLT AS INT))  AS VARCHAR) NW");
            //sb.Append(" ,'@'+CAST(CAST(REPLACE(T2.GWeight,'@','') AS DECIMAL(18,2))*(CAST(T2.CARTON AS INT)/CAST(T2.PLT AS INT))  AS VARCHAR) GW");
            //sb.Append(" ,'@'+REPLACE(T2.CBM,'*','x')  CBM,PLT2 PLT,T2.CARTON,T2.CARTON2,T2.PLT PLT2 from rpa_packingH T0     ");
            //sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)        ");
            //sb.Append(" WHERE  T2.ID=@ID  ");
            //sb.Append(" UNION ALL  ");
            //sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty2,',','')  QTY,T2.NWeight2  NW,GWeight2  GW,'','',T2.CARTON,'',T2.PLT PLT2  from rpa_packingH T0     ");
            //sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)    ");
            //sb.Append(" WHERE  T2.ID=@ID  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL2F22(string ID)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,'@'+CAST(T2.Qty2/CAST(T2.PLT AS INT) AS VARCHAR)  QTY,'@'+CAST(CAST(CAST(T2.NWeight AS decimal(18,3))/CAST(T2.PLT AS decimal) AS decimal(18,3)) AS VARCHAR)   NW,'@'+CAST(CAST(CAST(T2.GWeight AS decimal(18,3))/CAST(T2.PLT AS decimal) AS decimal(18,3)) AS VARCHAR)  GW,'@'+REPLACE(T2.CBM,'*','x')  CBM,REPLACE(REPLACE(T2.CartonNo,'PALLET',''),' ','') PLT,T0.CARTON,T2.CARTON2,T2.PLT PLT2 from rpa_packingH T0    ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            sb.Append(" WHERE  T2.ID=@ID ");
            sb.Append(" UNION ALL ");
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,T2.Qty2  QTY,T2.NWeight  NW,GWeight GW,'','',T0.CARTON,'',T2.PLT PLT2  from rpa_packingH T0    ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            sb.Append(" WHERE  T2.ID=@ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GETPACL2F2SS(string ID)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,T2.Qty QTY ");
            sb.Append(" ,'@'+CAST(CAST(REPLACE(T2.NWeight,'@','') AS DECIMAL(18,2)) AS VARCHAR) NW ");
            sb.Append(" ,'@'+CAST(CAST(REPLACE(T2.GWeight,'@','') AS DECIMAL(18,2)) AS VARCHAR) GW ");
            sb.Append(" ,'@'+REPLACE(T2.CBM,'*','x')  CBM,PLT2 PLT,T2.CARTON,T2.CARTON2,T2.PLT PLT2 from rpa_packingH T0      ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)         ");
            sb.Append(" WHERE  T2.ID=@ID ");
            sb.Append(" UNION ALL   ");
            sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,REPLACE(T2.Qty2,',','')  QTY,T2.NWeight2  NW,GWeight2  GW,'','',T2.CARTON,'',T2.PLT PLT2  from rpa_packingH T0      ");
            sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)     ");
            sb.Append(" WHERE  T2.ID=@ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
            //public static System.Data.DataTable GETPACL2F22(string ID)
            //{

            //    SqlConnection connection = globals.Connection;
            //    StringBuilder sb = new StringBuilder();
            //    sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,'@'+CAST(T2.Qty2/CAST(T2.PLT AS INT) AS VARCHAR)  QTY,'@'+CAST(CAST(CAST(T2.NWeight AS decimal(18,3))/CAST(T2.PLT AS decimal) AS decimal(18,3)) AS VARCHAR)   NW,'@'+CAST(CAST(CAST(T2.GWeight AS decimal(18,3))/CAST(T2.PLT AS decimal) AS decimal(18,3)) AS VARCHAR)  GW,'@'+REPLACE(T2.CBM,'*','x')  CBM,REPLACE(REPLACE(T2.CartonNo,'PALLET',''),' ','') PLT,T0.CARTON,T2.CARTON2,T2.PLT PLT2 from rpa_packingH T0    ");
            //    sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            //    sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            //    sb.Append(" WHERE  T2.ID=@ID ");
            //    sb.Append(" UNION ALL ");
            //    sb.Append(" select  T0.ID,T2.ItemNo PARTNO,T0.InvoiceNo INV,T2.Qty2  QTY,T2.NWeight  NW,GWeight GW,'','',T0.CARTON,'',T2.PLT PLT2  from rpa_packingH T0    ");
            //    sb.Append(" LEFT JOIN rpa_packingD T2 ON (T0.InvoiceNo =T2.InvoiceNo)   ");
            //    sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)    ");
            //    sb.Append(" WHERE  T2.ID=@ID ");

            //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
            //    command.CommandType = CommandType.Text;
            //    command.Parameters.Add(new SqlParameter("@ID", ID));
            //    SqlDataAdapter da = new SqlDataAdapter(command);

            //    DataSet ds = new DataSet();
            //    try
            //    {
            //        connection.Open();
            //        da.Fill(ds, "odln");
            //    }
            //    finally
            //    {
            //        connection.Close();
            //    }

            //    return ds.Tables[0];

            //}
        public static void GETCBM(string INV, string CBMM)
        {
            System.Data.DataTable K2 = null;
            string LB = INV.Substring(0, 2);
            //if (LB == "LB" || LB == "HS" || LB == "MS")
            //{
            //    K2 = GETPACL2B(INV);
            //}
            //else
            //{
            //    K2 = GETPACL2(INV);
            //}

            K2 = GETPACL2B(INV);
            decimal CBB = 0;

            if (K2.Rows.Count > 0)
            {
           //     System.Data.DataTable K3 = util.GETPACL3(INV);
                for (int i = 0; i <= K2.Rows.Count - 1; i++)
                {
                    string CBM = K2.Rows[i]["CBM"].ToString();
                    string DD = K2.Rows[i]["PLT"].ToString();
                    string[] CMS = CBM.ToUpper().Split(new Char[] { '*' });
                    int M1 = 0;
                    string L = "";
                    string W = "";
                    string H = "";
                    int T2 = -1;

                    foreach (string F in CMS)
                    {
                        M1++;
                        if (M1 == 1)
                        {
                            L = F;
                        }
                        if (M1 == 2)
                        {
                            W = F;
                        }
                        if (M1 == 3)
                        {
                            T2 = F.IndexOf("*");
                            if (T2 != -1)
                            {

                                H = F.Substring(0, T2);
                            }
                            else
                            {
                                H = F;
                            }
                        }
                    }

                    decimal n;

                    decimal GA = 1000000;


                    if (decimal.TryParse(L, out n) && decimal.TryParse(W, out n) && decimal.TryParse(H, out n) && decimal.TryParse(DD, out n))
                    {
                        if (DD == "0")
                        {
                            DD = "1";
                        }
                        decimal ff3 = (Convert.ToDecimal(L) * Convert.ToDecimal(W) * Convert.ToDecimal(H)) * Convert.ToDecimal(DD);
                        CBB += ff3 / GA;

                    }

                }
                string CARTON = "";
                string PLT = "";
                System.Data.DataTable K3 = null;

                if (LB == "LB" )
                {
                    K3 = util.GETPACL3B(INV);
                    CARTON = K3.Rows[0]["CARTON"].ToString();
                    PLT = K3.Rows[0]["PLT"].ToString();
                }
                else if (LB == "HS" || LB == "JJ" || LB == "MS")
                {
                    K3 = util.GETPACL3B2(INV);
                    CARTON = K3.Rows[0]["CARTON"].ToString();
                    PLT = K3.Rows[0]["PLT"].ToString();
                }
                else
                {
                    K3 = util.GETPACL3(INV);
                    if (K3.Rows.Count > 0)
                    {
                        string SAYTOTAL = K3.Rows[0][1].ToString();
                        if (SAYTOTAL != "")
                        {
                            int G1 = SAYTOTAL.LastIndexOf("(");
                            int G2 = SAYTOTAL.LastIndexOf(")");
                            CARTON = SAYTOTAL.Substring(G1 + 1, G2 - G1 - 1);
                            PLT = K3.Rows[0][0].ToString();
                            if (CARTON == PLT)
                            {
                                if (INV != "M021810499")
                                {
                                    PLT = "1";
                                }
                            }
                        }
                    }
                }
                if (String.IsNullOrEmpty(CARTON))
                {
                    K3 = util.GETPACL3(INV);
                    if (K3.Rows.Count > 0)
                    {
                        string SAYTOTAL = K3.Rows[0][1].ToString();
                        if (SAYTOTAL != "")
                        {
                            int G1 = SAYTOTAL.LastIndexOf("(");
                            int G2 = SAYTOTAL.LastIndexOf(")");
                            CARTON = SAYTOTAL.Substring(G1 + 1, G2 - G1 - 1);

                        }
                    }

                }
                decimal FF1 = Math.Round(CBB, 2, MidpointRounding.AwayFromZero);
                UPDATEPACK2(Convert.ToString(FF1), PLT, CARTON, INV, CBMM);
            }
        }
        public static System.Data.DataTable download21(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT [PATH] FROM download WHERE (FILENAME LIKE '%IP%' OR FILENAME LIKE '%AWB%' OR FILENAME LIKE '%BL%') AND  FILENAME NOT LIKE '%ZIP%' AND SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static System.Data.DataTable GETPACL2(string InvoiceNo)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select *,");
            sb.Append(" CASE WHEN InvoiceNo LIKE '%JJ%' THEN CAST(SUBSTRING(CartonNo,CHARINDEX('-', CartonNo)+1,2) AS INT)-CASE CAST(SUBSTRING(CartonNo,0,CHARINDEX('-', CartonNo)) AS INT) WHEN 0 THEN 1 ELSE CAST(SUBSTRING(CartonNo,0,CHARINDEX('-', CartonNo)) AS INT)  END+1");
            sb.Append(" ELSE ");
            sb.Append(" CAST(SUBSTRING(CartonNo,CHARINDEX('-', CartonNo)+1,2) AS INT)-CAST(SUBSTRING(CartonNo,CHARINDEX('PALLET', CartonNo)+6,CHARINDEX('-', CartonNo)-CHARINDEX('PALLET', CartonNo)-6) AS INT)+1 ");
            sb.Append(" END PLT");
            sb.Append(" from rpa_packingD");
            sb.Append(" WHERE INVOICENO=@InvoiceNo AND QTY <>'' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL2B(string InvoiceNo)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select *  from rpa_packingD WHERE InvoiceNo =@InvoiceNo");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static void UPDATEPACKH(string PLT, string invoiceno)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update RPA_PackingH set PLT=@PLT WHERE invoiceno=@invoiceno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACK(string CBM, string PLT, string invoiceno, string cartonno)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CBM=@CBM,PLT=@PLT WHERE invoiceno=@invoiceno AND cartonno=@cartonno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CBM", CBM));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("@cartonno", cartonno));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACK2(string CBM, string PLT, string invoiceno, string cartonno)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CBM=@CBM,PLT=@PLT WHERE invoiceno=@invoiceno AND   REPLACE(cartonno,' ','')=@cartonno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CBM", CBM));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("@cartonno", cartonno));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        public static void UPDATEPACKINGD(string SHIPPINGCODE, string LOCATION)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update PackingListD set LOCATION=@LOCATION WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACK2(string CARTON, string CARTON2, string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CARTON=@CARTON,CARTON2=@CARTON2 WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@CARTON2", CARTON2));
            command.Parameters.Add(new SqlParameter("@ID", ID));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        public static void UPDATEPACK22( string CARTON2, string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CARTON2=@CARTON2 WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CARTON2", CARTON2));
            command.Parameters.Add(new SqlParameter("@ID", ID));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACK222(string CARTON,string CARTON2,string PLT2, string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CARTON=@CARTON,CARTON2=@CARTON2,PLT2=@PLT2 WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));

            command.Parameters.Add(new SqlParameter("@CARTON2", CARTON2));
            command.Parameters.Add(new SqlParameter("@PLT2", PLT2));
            command.Parameters.Add(new SqlParameter("@ID", ID));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACK2222( string CARTON2, string PLT, string PLT2, string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CARTON2=@CARTON2,PLT=@PLT,PLT2=@PLT2 WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CARTON2", CARTON2));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@PLT2", PLT2));
            command.Parameters.Add(new SqlParameter("@ID", ID));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEPACKLB(string CBM, string PLT, string invoiceno, string CBM2)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingD set CBM=@CBM,PLT=@PLT WHERE invoiceno=@invoiceno AND CBM=@CBM2 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CBM", CBM));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("@CBM2", CBM2));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPH()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  UPDATE RPA_Packingh SET PLT=1 WHERE  SayTotal NOT LIKE '%PALLET%' AND PLT <> '1' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private static void UPDATEPACK2(string CBM, string PLT, string CARTON, string invoiceno, string CBMM)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingH set CBM=@CBM,PLT=@PLT,CARTON=@CARTON,CBMM=@CBMM WHERE invoiceno=@invoiceno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CBM", CBM));
            command.Parameters.Add(new SqlParameter("@PLT", PLT));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("@CBMM", CBMM));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static void UPDATEM(string TOTALQTY, string TotalGW, string TotalNW, string invoiceno)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update rpa_packingH set TOTALQTY=@TOTALQTY,TotalGW=@TotalGW,TotalNW=@TotalNW WHERE invoiceno=@invoiceno ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@TOTALQTY", TOTALQTY));
            command.Parameters.Add(new SqlParameter("@TotalGW", TotalGW));
            command.Parameters.Add(new SqlParameter("@TotalNW", TotalNW));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static System.Data.DataTable GETPACLS2(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @name varchar(100) ");
            sb.Append(" select @name =SUBSTRING(COALESCE(@name + '/',''),0,99) +PLT");
            sb.Append(" from  ( SELECT    CBM+ CASE WHEN PLT='1' THEN '' ELSE  '*'+PLT END PLT    FROM rpa_packingD WHERE InvoiceNo=@InvoiceNo AND ISNULL(CBM,'') <> '') PC");
            sb.Append(" SELECT @name CBMM");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACLS3(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(CAST(PLT AS INT)),0) PLT FROM RPA_PackingD WHERE InvoiceNo =@InvoiceNo ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GETPACLS3W(string InvoiceNo)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" 			  SELECT substring(SAYTOTAL,CHARINDEX('(', SAYTOTAL)+1,CHARINDEX(')', SAYTOTAL)-CHARINDEX('(', SAYTOTAL)-1) FROM ACMESQLSP.DBO.rpa_packingH WHERE SAYTOTAL like '%(%' AND SAYTOTAL like '%PALLET%' AND InvoiceNo =@InvoiceNo ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static  System.Data.DataTable GetITEMNAME(string SHIPPINGCODE, string U_PARTNO)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();

            sb.Append(" select t1.itemcode ITEMCODE,U_ITEMNAME COLLATE  Chinese_Taiwan_Stroke_CI_AS +' '+REPLACE(ISNULL(U_MODEL,''),'NON','') ITEMNAME,t1.Docentry DOC,linenum,VISORDER,T1.CURRENCY from shipping_item T1   ");
            sb.Append(" LEFT JOIN  ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append(" WHERE T1.SHIPPINGCODE=@SHIPPINGCODE   ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetITEMNAME2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();

            sb.Append(" select t1.itemcode ITEMCODE,U_ITEMNAME COLLATE  Chinese_Taiwan_Stroke_CI_AS +' '+REPLACE(ISNULL(U_MODEL,''),'NON','') ITEMNAME,t1.Docentry DOC,linenum,VISORDER,T1.CURRENCY,T1.Quantity QTY,T1.ItemPrice PRICE,T1.ItemAmount AMT from shipping_item T1   ");
            sb.Append(" LEFT JOIN  ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append(" WHERE T1.SHIPPINGCODE=@SHIPPINGCODE  ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetPACKCOUNT(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();

       
            sb.Append(" SELECT COUNT(*) C FROM PackingListD  ");
            sb.Append(" WHERE SHIPPINGCODE=@SHIPPINGCODE  ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetMARKCOUNT(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT COUNT(*) C FROM MARK WHERE SHIPPINGCODE=@SHIPPINGCODE   ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GetITEMNAME3(string SHIPPINGCODE, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();

            sb.Append(" 			  SELECT INDESCRIPTION ITEMNAME FROM INVOICED WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE  ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public static System.Data.DataTable GetITEMNAME4(string SHIPPINGCODE, string PARTNO)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT INDESCRIPTION ITEMNAME FROM INVOICED T0 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ItemCode COLLATE  Chinese_Taiwan_Stroke_CI_AS) WHERE  SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append(" AND T1.U_PARTNO =@PARTNO");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public static System.Data.DataTable GetITEMNAME5(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string aa = '"'.ToString();
            StringBuilder sb = new StringBuilder();


            sb.Append("			  SELECT INDESCRIPTION ITEMNAME FROM INVOICED T0   WHERE  SHIPPINGCODE=@SHIPPINGCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public static System.Data.DataTable GETPACLD(string INVOICENO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM RPA_PackingD   WHERE INVOICENO = @INVOICENO AND ISNULL(CBM,'') <> ''");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETLB1(string INVOICENO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select  ID,CARTONNO,ISNULL(PLT,'') PLT from RPA_PackingD  WHERE INVOICENO LIKE '%LB%' AND InvoiceNo=@INVOICENO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETLB2(string INVOICENO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select  ID,CARTONNO PLT,CARTON from RPA_PackingD  WHERE  InvoiceNo=@INVOICENO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GETPACL(string INVOICENO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select DISTINCT INVOICENO,InvoiceDate,TotalQty Qty,TotalNW NW,TotalGW GW,SayTotal,SIZE 版數,CBM,PLT,CARTON from rpa_packingH T0");
            sb.Append("  LEFT JOIN ACMESQL02.DBO.OPDN　T1 ON (T0.InvoiceNo=T1.U_ACME_INV  COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
         //   sb.Append("   WHERE  INVOICENO=@INVOICENO AND ISNULL(CBMM,'') ='' ");
            sb.Append("   WHERE  INVOICENO=@INVOICENO ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static void AddPACKMAIN(string ShippingCode, string PLNo, string PDate, string ForAccount, string ShippedBy, string Shipping_From, string Shipping_Per, string Shipping_To, string ShippedOn, string Bill_To, string Memo, string ColumnTotal, int Quantity, decimal Net, decimal Gross, string SayTotal, string CBM, string SayCTN)
        {
            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into PackingListM(ShippingCode,PLNo,PDate,ForAccount,ShippedBy,Shipping_From,Shipping_Per,Shipping_To,ShippedOn,Bill_To,Memo,ColumnTotal,Quantity,Net,Gross,SayTotal,CBM,SayCTN) values(@ShippingCode,@PLNo,@PDate,@ForAccount,@ShippedBy,@Shipping_From,@Shipping_Per,@Shipping_To,@ShippedOn,@Bill_To,@Memo,@ColumnTotal,@Quantity,@Net,@Gross,@SayTotal,@CBM,@SayCTN)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@PDate", PDate));
            command.Parameters.Add(new SqlParameter("@ForAccount", ForAccount));
            command.Parameters.Add(new SqlParameter("@ShippedBy", ShippedBy));
            command.Parameters.Add(new SqlParameter("@Shipping_From", Shipping_From));
            command.Parameters.Add(new SqlParameter("@Shipping_Per", Shipping_Per));
            command.Parameters.Add(new SqlParameter("@Shipping_To", Shipping_To));
            command.Parameters.Add(new SqlParameter("@ShippedOn", ShippedOn));
            command.Parameters.Add(new SqlParameter("@Bill_To", Bill_To));
            command.Parameters.Add(new SqlParameter("@Memo", Memo));
            command.Parameters.Add(new SqlParameter("@ColumnTotal", ColumnTotal));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Net", Net));
            command.Parameters.Add(new SqlParameter("@Gross", Gross));
            command.Parameters.Add(new SqlParameter("@SayTotal", SayTotal));
            command.Parameters.Add(new SqlParameter("@CBM", CBM));
            command.Parameters.Add(new SqlParameter("@SayCTN", SayCTN));

            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }

        public static void AddPACKD(string ShippingCode, string PLNo, string SeqNo, string PackageNo, string CNo, string DescGoods, string Quantity, string Net, string Gross, string MeasurmentCM, string PACKMARK)
        {
            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into PackingListD(ShippingCode,PLNo,SeqNo,PackageNo,CNo,DescGoods,Quantity,Net,Gross,MeasurmentCM,LOCATION,PACKMARK) values(@ShippingCode,@PLNo,@SeqNo,@PackageNo,@CNo,@DescGoods,@Quantity,@Net,@Gross,@MeasurmentCM,@LOCATION,@PACKMARK)", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@PackageNo", PackageNo));
            command.Parameters.Add(new SqlParameter("@CNo", CNo));
            command.Parameters.Add(new SqlParameter("@DescGoods", DescGoods));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Net", Net));
            command.Parameters.Add(new SqlParameter("@Gross", Gross));
            command.Parameters.Add(new SqlParameter("@MeasurmentCM", MeasurmentCM));
            command.Parameters.Add(new SqlParameter("@LOCATION", "CHINA"));
            command.Parameters.Add(new SqlParameter("@PACKMARK", PACKMARK));
            //PACKMARK

            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }

        public static System.Data.DataTable GetCART(string MODEL, string VER, string GRADE, int NW, string QQ)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT *  FROM acmesqlsp.dbo.CART  WHERE MODEL_NO =@MODEL AND MODEL_Ver =@VER AND TMEMO =@GRADE  AND ISNULL(CT_QTY,'') <> '' ");
            if (NW == 1)
            {
                sb.Append(" AND ISNULL(CT_QTY,'')=@QQ  ");
            }
            sb.Append("ORDER BY CREATE_DATE  DESC  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@QQ", QQ));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        public static System.Data.DataTable GetCARTK(string MODEL, string VER, string GRADE, string DOCTYPE, int NW, string QQ)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT *  FROM acmesqlsp.dbo.CART  WHERE MODEL_NO LIKE '%" + MODEL + "%' AND MODEL_Ver =@VER AND TMEMO =@GRADE  AND ISNULL(CT_QTY,'') <> ''");


            if (DOCTYPE == "O")
            {
                sb.Append(" AND DOCTYPE='Open Cell' ");
            }
            else
            {
                sb.Append(" AND DOCTYPE<>'Open Cell' ");
            }
            if (NW == 1)
            {
                sb.Append(" AND ISNULL(CT_QTY,'')=@QQ  ");
            }
            sb.Append("ORDER BY CREATE_DATE  DESC  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@QQ", QQ));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        public static System.Data.DataTable GetCARTJ(string MODEL, string VER, int NW, string QQ)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT *  FROM acmesqlsp.dbo.CART  WHERE MODEL_NO LIKE '%" + MODEL + "%' AND MODEL_Ver =@VER  AND ISNULL(CT_QTY,'') <> ''  AND ISNULL(CT_NW,'') <> '' ");
            if (NW == 1)
            {
                sb.Append(" AND ISNULL(CT_QTY,'')=@QQ  ");
            }

            sb.Append(" ORDER BY CREATE_DATE  DESC  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@QQ", QQ));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GetCARTL(string MODEL, string VER, int NW, string QQ)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT *  FROM acmesqlsp.dbo.CART  WHERE MODEL_NO = '" + MODEL + "' AND MODEL_Ver =@VER  AND ISNULL(CT_QTY,'') <> ''  AND ISNULL(CT_NW,'') <> ''   ");
            if (NW == 1)
            {
                sb.Append(" AND ISNULL(CT_QTY,'')=@QQ  ");
            }

            sb.Append(" ORDER BY CREATE_DATE  DESC  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@QQ", QQ));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        public static System.Data.DataTable GetITEMCODE(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE,T1.ITEMNAME,T1.U_VERSION VER  FROM ACMESQLSP.DBO.SHIPPING_ITEM T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("  WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GetOITML(string MODEL)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT U_TMODEL   FROM OITM WHERE U_TMODEL LIKE '%" + MODEL + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GetOITMW(string ITEMCODE)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUBSTRING(ITEMNAME,CHARINDEX('V.', ITEMNAME)+2,3) VER FROM OITM WHERE ITEMCODE=@ITEMCODE AND ITEMNAME LIKE '%V.%' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

    }
}

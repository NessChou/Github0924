using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;

//放置系統共用變數
namespace ACME
{
    class globals
    {
        public static ConnectionStringSettings settings;
        //資料連結字串
        public static string ConnectionString;

        public static SqlConnection Connection;

      
        //ship字串連結
        public static string shipConnectionString;

        public static SqlConnection shipConnection;

        //lp系統共用的連結
        public static SqlConnection lpConnection;
        public static ConnectionStringSettings lpSettings;
        //lp資料連結字串
        public static string lpConnectionString;

        public static string CHOICEConnectionString;
        public static SqlConnection CHOICEConnection;

        public static string EEPConnectionString;
        public static SqlConnection EEPConnection;


        //使用者編號
        public static string UserID;
        public static string GroupID;
        //錯誤處理
        public static string sErrMsg;
        public static int lErrCode;
        public static int lRetCode;

        public static string DBNAME;
        public static string SERVER;
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;

//��m�t�Φ@���ܼ�
namespace ACME
{
    class globals
    {
        public static ConnectionStringSettings settings;
        //��Ƴs���r��
        public static string ConnectionString;

        public static SqlConnection Connection;

      
        //ship�r��s��
        public static string shipConnectionString;

        public static SqlConnection shipConnection;

        //lp�t�Φ@�Ϊ��s��
        public static SqlConnection lpConnection;
        public static ConnectionStringSettings lpSettings;
        //lp��Ƴs���r��
        public static string lpConnectionString;

        public static string CHOICEConnectionString;
        public static SqlConnection CHOICEConnection;

        public static string EEPConnectionString;
        public static SqlConnection EEPConnection;


        //�ϥΪ̽s��
        public static string UserID;
        public static string GroupID;
        //���~�B�z
        public static string sErrMsg;
        public static int lErrCode;
        public static int lRetCode;

        public static string DBNAME;
        public static string SERVER;
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient; //�s�W�R�W�Ŷ� for SQL Server
using System.IO;
using System.Windows.Forms;
using System.Configuration;//

namespace My
{
    public class MyGlobal
    {
        public static string GlobalUserID;      //�O���n�J�ϥΪ̱b��
        public static string GlobalPassword;    //�O���n�J�ϥΪ̱K�X
        public static string GlobalHashPassword;//�O���n�J�ϥΪ̥[�K�L�᪺�K�X
        public static string GlobalRoleName;    //�O���n�J�ϥΪ̨���W��

        public static int GlobalLoginErrorCounter; //�O���n�J���~����
        public static bool GlobalSystemShutdown = false;  //���ܼƥD�n�ΨӧP�_�O�եΪ��ɶ���n�����t�Ρ@�٬O�@���`�ϥΤU�����t��
        public static string GlobalSysRegDefaultPath = @"Software\LMS\V1.0\"; //�n���ɳn����U���|

        //public static string DefaultMultiLanguage = "zh-tw";

        #region ***SQL Server 2005 EXPRESS***

        //***SQL Server 2005 ***
        public const string DbType = "SQL Server 2005";        //�w�]�ϥθ�Ʈw
        public const string DbName = "LMS";                    //�ϥθ�Ʈw�W��
        public const string SQLUserID = "sa";                  //SQL Server�ϥΪ�ID
        public const string SQLUserPwd = "123";                //SQL Server�ϥΪ̱K�X
        public const string DBServer = "intel";                //��Ʈw���A���W�� �]�O �q���W��
        public const string WorkStationID = "intel";           //�q���W��
        public const string ServerIP = "127.0.0.1";            //�����ҳ]�w��IP
        public const string ServerDNS = "localhost";           //�����A���W��
        public const string DataSource = ".\\SQLEXPRESS";
               

        private static string ConnString;
        public static string SQLConnectionString
        {
            get
            {
                //�ϥ����ε{���պA�ɳs�u�覡,�ݥ[�J�Ѧ�[System.configuration.dll]
                ConnectionStringSettings settings;
                //settings = ConfigurationManager.ConnectionStrings["SQLConnectionString"];

                //�Y�n����Ū���ɮ׫h�ϥΥH�U�y�k
                settings = ConfigurationManager.ConnectionStrings["SQLConnectionString"];
                ConnString = settings.ConnectionString;
                ConnString = ConnString.Replace("|DataDirectory|", System.Windows.Forms.Application.StartupPath);

                ConnString = settings.ConnectionString;
                return ConnString;
            }
        }

        #endregion

        

        #region ***�t�ζ}�o���ҳ]�w

        //***�t�ζ}�o���ҳ]�w
        public int MininumPasswordLength = 3;             //�̤p�K�X����
        public int MaxUserNameLength = 20;                //�̤j�ϥΪ̦W�٪���

        public const string GlobalSystemName = "LMS";
        public const string GlobalSystemTitle = "�ϮѺ޲z�t�� Library Management System";
        public const string GlobalSystemVersion = "V1.0";
        public const string GlobalUseLocale = "zh-tw";//�w�]�y�t      
        public const string GlobalDefaultLanguage = "C Sharp .NET 2.0"; //�O���t�ΨϥΦ�ص{���y���}�o

        

        #endregion

        

        #region ��Ʈw�s�u�]�w

        /// <summary>
        /// ��Ʈw�s�u�]�w
        /// </summary>
        /// <param name="AppPath">���ε{������Ʈw�Ҧb��m</param>
        /// <param name="DBType">��Ʈw�����G�i����Access��SQLServer</param>
        /// <returns>�^�ǳs�u�r���</returns>
        public string DBConnectionString(string AppPath, string DBType)
        {
            if (DBType == "SQL Server 2005")
            {
                //Builder["data Source"] = "192.168.12.185";
                //Builder["initial catalog"] = "company";
                //Builder["user id"] = "sa";
                //Builder["Password"] = "123";
                return "";
            }
            else //SQL Server 2005 Express
            {
                SqlConnectionStringBuilder Builder = new SqlConnectionStringBuilder();

                //Builder["data Source"] = "192.168.12.185";
                //Builder["initial catalog"] = "company";
                //Builder["user id"] = "sa";
                //Builder["Password"] = "123";
                Builder.DataSource = "192.168.0.1";
                Builder.InitialCatalog = "IPQC";
                Builder.UserID = "sa";
                Builder.Password = "123";
                return Builder.ConnectionString;


            }

        }

        #endregion


        #region ���������Ʈw�s�u�]�w

        /// <summary>
        /// ���������Ʈw�s�u�]�w
        /// </summary>
        /// <param name="AppPath">���ε{������Ʈw�Ҧb��m</param>
        /// <param name="DBType">��Ʈw�����G�i����Access��SQLServer</param>
        /// <returns>�^�ǳs�u�r���</returns>
        public string CRDBConnString(string AppPath, string DBType)
        {
            if (DBType == "SQL Server 2005")
            {
                //Builder["data Source"] = "192.168.12.185";
                //Builder["initial catalog"] = "company";
                //Builder["user id"] = "sa";
                //Builder["Password"] = "123";
                return "";
            }
            else //SQL Server 2005 Express
            {
                SqlConnectionStringBuilder Builder = new SqlConnectionStringBuilder();

                Builder.DataSource = "192.168.0.1";
                Builder.InitialCatalog = "IPQC";
                Builder.UserID = "sa";
                Builder.Password = "123";
                return Builder.ConnectionString;


            }

        }

        #endregion


        

    }
}

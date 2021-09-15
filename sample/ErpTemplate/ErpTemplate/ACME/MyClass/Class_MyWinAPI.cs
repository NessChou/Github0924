using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;//MessageBox.Show
using Microsoft.Win32; //for RegistryKey

///���X�����Ҧ��ƭȪ��`�p�j�p����W�L64K
///�G�i��� REG_BINARY ��l���G�i����,�i���ASCII���byte�s��H�g�J
///DWORD�ȡ@REG_DWORD �H4�Ӧ줸�ժ��רӪ�ܸ��(�ƭȫ��A)
///�r��� REG_SZ �T�w���ת���r��

namespace My
{
    public class MyWinAPI
    {
        static RegistryKey rkCR = Registry.ClassesRoot;
        static RegistryKey rkCC = Registry.CurrentConfig;
        static RegistryKey rkCU = Registry.CurrentUser;
        static RegistryKey rkLM = Registry.LocalMachine;
        static RegistryKey rkUsers = Registry.Users;
        static RegistryKey rk;
        
        #region "�ھ�[�s���X]���� �إ߷s��[�l���X] �ó]�w���"

        /// <summary>
        /// �ھ�[�s���X]���� �إ߷s��[�l���X] �ó]�w���
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <param name="KeyName">�ƭȦW��</param>
        /// <param name="KeyVal">�ƭȸ��</param>
        public static bool CreateSubKeyAndSetValue(string RegistryType, string AppName, string KeyName, object KeyVal)
        {
            //�Ϊk
            //CreateSubKeyAndSetValue("CurrentUser", @"eCRAM System\Login", "LoginCount", 2);
            //CreateSubKeyAndSetValue("CurrentUser", @"eCRAM System\Login", "LoginUser", "kevin");
            bool result = false;

            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        rkCR.CreateSubKey(AppName);
                        rkCR.OpenSubKey(AppName, true).SetValue(KeyName, KeyVal);
                        rkCR.Close();
                        result = true;
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        rkCC.CreateSubKey(AppName);
                        rkCC.OpenSubKey(AppName, true).SetValue(KeyName, KeyVal);
                        rkCC.Close();
                        result = true;
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        rkCU.CreateSubKey(AppName);
                        rkCU.OpenSubKey(AppName, true).SetValue(KeyName, KeyVal);
                        rkCU.Close();
                        result = true;
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        rkLM.CreateSubKey(AppName);
                        rkLM.OpenSubKey(AppName, true).SetValue(KeyName, KeyVal);
                        rkLM.Close();
                        result = true;
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        rkUsers.CreateSubKey(AppName);
                        rkUsers.OpenSubKey(AppName, true).SetValue(KeyName, KeyVal);
                        rkUsers.Close();
                        result = true;
                        break;
                    default:
                        result = false;
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("�إߤl���X�P�]�w��ȥ���,���~�T����:" + ex.Message, "Registry�@�~");
                return false;
            }


        }

        #endregion


        #region "����l���X���e��"

        /// <summary>
        /// ����l���X���e��
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <param name="KeyName">�ƭȦW��</param>
        /// <returns></returns>
        public static object GetSubKeyValue(string RegistryType, string AppName, string KeyName)
        {

            //�Ϊk(string)GetSubKeyValue("CurrentUser", @"eCRAM System\Login", "LoginUser")
            object result;

            string ErrorMessageString = "";

            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        result = rkCR.OpenSubKey(AppName, true).GetValue(KeyName, false); //�Y���X���e�Ȥ��s�b�h�|�^��false
                        rkCR.Close();
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        result = rkCC.OpenSubKey(AppName, true).GetValue(KeyName, false);
                        rkCC.Close();
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        result = rkCU.OpenSubKey(AppName, true).GetValue(KeyName, false);
                        rkCU.Close();
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        result = rkLM.OpenSubKey(AppName, true).GetValue(KeyName, false);
                        rkLM.Close();
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        result = rkUsers.OpenSubKey(AppName, true).GetValue(KeyName, false);
                        rkUsers.Close();
                        break;
                    default:
                        result = false;
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                //MessageBox.Show("����l���X���e�ȥ���,���~�T����:" + ex.Message, "Registry�@�~");
                ErrorMessageString = "����l���X���e�ȥ���,���~�T����:" + ex.Message;
                return false;
            }

        }

        #endregion


        #region "�P�_�l���X�O�_�s�b"

        /// <summary>
        /// �P�_�l���X�O�_�s�b
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <returns>�^�ǥ��L��</returns>
        public static bool SubKeyExist(string RegistryType, string AppName)
        {
            //�Ϊk
            //SubKeyExist("CurrentUser", @"eCRAM System\Login");

            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        rk = rkCR.OpenSubKey(AppName, true);
                        rkCR.Close();
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        rk = rkCC.OpenSubKey(AppName, true);
                        rkCC.Close();
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        rk = rkCU.OpenSubKey(AppName, true);
                        rkCU.Close();
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        rk = rkLM.OpenSubKey(AppName, true);
                        rkLM.Close();
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        rk = rkUsers.OpenSubKey(AppName, true);
                        rkUsers.Close();
                        break;
                    default:
                        break;
                }

                if (rk == null)
                {
                    rk.Close();
                    return false;
                }
                else
                {
                    rk.Close();
                    return true;
                }


            }

            catch (Exception ex)
            {
                MessageBox.Show("�P�_�l���X�O�_�s�b����,���~�T����:" + ex.Message, "Registry�@�~");
                return false;
            }


        }

        #endregion


        #region "�R�����w�l���X"

        /// <summary>
        /// �R�����w�l���X
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <param name="KeyName">�ƭȦW��</param>
        /// <returns></returns>
        public static object DeleteSpecSubKey(string RegistryType, string AppName, string KeyName)
        {

            //�Ϊkbool result = (bool)DeleteSpecSubKey("CurrentUser", "eCRAM System", "Login");
            object result;

            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        RegistryKey rk1 = rkCR.OpenSubKey(AppName, true);
                        rk1.DeleteSubKey(KeyName);
                        rk1.Close();
                        rkCR.Close();
                        result = true;
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        RegistryKey rk2 = rkCC.OpenSubKey(AppName, true);
                        rk2.DeleteSubKey(KeyName);
                        rk2.Close();
                        rkCC.Close();
                        result = true;
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        RegistryKey rk3 = rkCU.OpenSubKey(AppName, true);
                        rk3.DeleteSubKey(KeyName);
                        rk3.Close();
                        rkCU.Close();
                        result = true;
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        RegistryKey rk4 = rkLM.OpenSubKey(AppName, true);
                        rk4.DeleteSubKey(KeyName);
                        rk4.Close();
                        rkLM.Close();
                        result = true;
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        RegistryKey rk5 = rkUsers.OpenSubKey(AppName, true);
                        rk5.DeleteSubKey(KeyName);
                        rk5.Close();
                        rkUsers.Close();
                        result = true;
                        break;
                    default:
                        result = false;
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("�R���l���X���e�ȥ���,���~�T����:" + ex.Message, "Registry�@�~");
                return false;
            }

        }

        #endregion


        #region "�R���Ҧ��l���X"
        /// <summary>
        /// �R���Ҧ��l���X
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <returns></returns>
        public static object DeleteALLSubKey(string RegistryType, string AppName)
        {

            //�Ϊkbool result = (bool) DeleteALLSubKey("CurrentUser", @"eCRAM System\Login", "LoginUser");
            object result;

            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        rkCR.DeleteSubKeyTree(AppName);
                        rkCR.Close();
                        result = true;
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        rkCC.DeleteSubKeyTree(AppName);
                        rkCC.Close();
                        result = true;
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        rkCU.DeleteSubKeyTree(AppName);
                        rkCU.Close();
                        result = true;
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        rkLM.DeleteSubKeyTree(AppName);
                        rkLM.Close();
                        result = true;
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        rkUsers.DeleteSubKeyTree(AppName);
                        rkUsers.Close();
                        result = true;
                        break;
                    default:
                        result = false;
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("�R���l���X������l���X����,���~�T����:" + ex.Message, "Registry�@�~");
                return false;
            }

        }

        #endregion


        #region "�R���l���X���ƭȦW�٤��e��"

        /// <summary>
        /// �R���l���X���ƭȦW�٤��e��
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <param name="ValueName">�ƭȦW��</param>
        /// <returns></returns>
        public static bool DeleteSubKeyValue(string RegistryType, string AppName, string ValueName)
        {

            bool result = false;
            //�Ϊkbool result = (bool)DeleteSubKeyValue("CurrentUser", @"eCRAM System\Login", "LoginUser");
            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        RegistryKey rk1 = rkCR.OpenSubKey(AppName, true);
                        rk1.DeleteValue(ValueName);
                        rk1.Close();
                        rkCR.Close();
                        result = true;
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        RegistryKey rk2 = rkCC.OpenSubKey(AppName, true);
                        rk2.DeleteValue(ValueName);
                        rk2.Close();
                        rkCC.Close();
                        result = true;
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        RegistryKey rk3 = rkCU.OpenSubKey(AppName, true);
                        rk3.DeleteValue(ValueName);
                        rk3.Close();
                        rkCU.Close();
                        result = true;
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        RegistryKey rk4 = rkLM.OpenSubKey(AppName, true);
                        rk4.DeleteValue(ValueName);
                        rk4.Close();
                        rkLM.Close();
                        result = true;
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        RegistryKey rk5 = rkUsers.OpenSubKey(AppName, true);
                        rk5.DeleteValue(ValueName);
                        rk5.Close();
                        rkUsers.Close();
                        result = true;
                        break;
                    default:
                        result = false;
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("�R���ƭȦW�٤��e�ȥ���,���~�T����:" + ex.Message, "Registry�@�~");
                return false;
            }
        }

        #endregion



        #region "����@�~�t�Ϊ�����T"

        /// <summary>
        /// ����@�~�t�Ϊ�����T
        /// </summary>
        /// <param name="KeyName">�ƭȦW��
        /// �]�t:ProductName , �p: Microsoft Windows XP
        ///      CSDVersion , �p: Service Pack 2
        ///      CurrentBuild , �p: 1.511.1 () (Obsolete data - do not use)
        ///      CurrentVersion , �p: 5.1
        ///      RegisteredOrganization , �p:mis
        ///      RegisteredOwner , �p:kevin
        /// </param>
        /// <returns></returns>
        public static string GetWindowsXPProfessionalInfo(string KeyName)
        {
            string result = "";

            result = (string)GetSubKeyValue("LocalMachine", @"Software\Microsoft\Windows NT\CurrentVersion", KeyName);
            return result;

        }

        #endregion


        #region "������w���X�U���Ҧ��l���X�W��"

        /// <summary>
        /// ������w���X�U���Ҧ��l���X�W��
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���w���X�W��</param>
        /// <returns></returns>
        public static string[] GetAllSubKeyNames(string RegistryType, string AppName)
        {

            //�Ϊkstring[] test = GetAllSubKeyNames("CurrentUser", "eCRAM System");
            string[] result = new string[1000];
            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        RegistryKey rk1 = rkCR.OpenSubKey(AppName, true);
                        result = rk1.GetSubKeyNames();
                        rk1.Close();
                        rkCR.Close();
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        RegistryKey rk2 = rkCC.OpenSubKey(AppName, true);
                        result = rk2.GetSubKeyNames();
                        rk2.Close();
                        rkCC.Close();
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        RegistryKey rk3 = rkCU.OpenSubKey(AppName, true);
                        result = rk3.GetSubKeyNames();
                        rk3.Close();
                        rkCU.Close();
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        RegistryKey rk4 = rkLM.OpenSubKey(AppName, true);
                        result = rk4.GetSubKeyNames();
                        rk4.Close();
                        rkLM.Close();
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        RegistryKey rk5 = rkUsers.OpenSubKey(AppName, true);
                        result = rk5.GetSubKeyNames();
                        rk5.Close();
                        rkUsers.Close();
                        break;
                    default:
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("������w���X�U���Ҧ��l���X�W�٥���,���~�T����:" + ex.Message, "Registry�@�~");
                return result;
            }
        }

        #endregion


        #region "����Ҧ��l���X�U���ƭȦW��"

        /// <summary>
        /// ����Ҧ��l���X�U���ƭȦW��
        /// </summary>
        /// <param name="RegistryType">�s���X�����]�t�GClassesRoot,CurrentConfig,CurrentUser,LocalMachine,Users</param>
        /// <param name="AppName">���X�W��</param>
        /// <returns></returns>
        public static string[] GetAllValueNames(string RegistryType, string AppName)
        {

            //�Ϊkstring[] test = GetAllValueNames("CurrentUser", @"eCRAM System\Login");
            string[] result = new string[1000];
            try
            {
                switch (RegistryType)
                {
                    case "ClassesRoot"://�w�q��������H�γo�Ǭ����p���ݩ�
                        RegistryKey rk1 = rkCR.OpenSubKey(AppName, true);
                        result = rk1.GetValueNames();
                        rk1.Close();
                        rkCR.Close();
                        break;
                    case "CurrentConfig"://�]�t�ثe�q���n�w������պA�]�w
                        RegistryKey rk2 = rkCC.OpenSubKey(AppName, true);
                        result = rk2.GetValueNames();
                        rk2.Close();
                        rkCC.Close();
                        break;
                    case "CurrentUser"://�]�t�ثe�ϥΪ̭ӤH���n�����]�w��T(�������̱`�ϥ�)
                        RegistryKey rk3 = rkCU.OpenSubKey(AppName, true);
                        result = rk3.GetValueNames();
                        rk3.Close();
                        rkCU.Close();
                        break;
                    case "LocalMachine"://�P�q���������]�w��T,�]�t�@�~�t�λP�w�鵲�c���,�ΨӦs�񥻾��q���պA��Ƥ��Ӥl���X(Hardware,SAM,Security,Software,System)
                        RegistryKey rk4 = rkLM.OpenSubKey(AppName, true);
                        result = rk4.GetValueNames();
                        rk4.Close();
                        rkLM.Close();
                        break;
                    case "Users"://�]�t�Ҧ��ϥΪ̦b�q���W�Ҧ��۰ʸ��J���ϥΪ̳]�w��
                        RegistryKey rk5 = rkUsers.OpenSubKey(AppName, true);
                        result = rk5.GetValueNames();
                        rk5.Close();
                        rkUsers.Close();
                        break;
                    default:
                        break;
                }
                return result;
            }

            catch (Exception ex)
            {
                MessageBox.Show("����Ҧ��l���X�U���ƭȦW�٥���,���~�T����:" + ex.Message, "Registry�@�~");
                return result;
            }

        }

        #endregion


    }
}

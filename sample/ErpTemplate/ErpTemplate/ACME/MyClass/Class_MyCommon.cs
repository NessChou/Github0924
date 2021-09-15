using System;
using System.Collections.Generic;
using System.Text;
using System.IO;//
using System.Data;//
using System.Data.SqlClient; //SQL Server�ϥΥ��n�ޤJNamespace

namespace My
{
    public class MyCommon
    {
        #region "test"

        #endregion




        #region �P�_�ǤJ���r��O�_���ŭ�

        /// <summary>
        /// �P�_�ǤJ���r��O�_���ŭ�
        /// </summary>
        /// <param name="bufstr">�ǤJ�r��</param>
        /// <returns>�^�Ǧr���</returns>
        public static bool IsNullString(string bufstr)
        {
            if (bufstr == null)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        #endregion


        #region �P�_�O�_�����L��

        /// <summary>
        /// �P�_�O�_�����L��
        /// </summary>
        /// <param name="objbuf"></param>
        /// <returns>�^�ǥ��L��</returns>
        public static bool IsBoolean(object objbuf)
        {
            string bufstr = (string)objbuf;
            bufstr = Microsoft.VisualBasic.Strings.LCase(bufstr);

            if (bufstr == "false" || bufstr == "true")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        #endregion


        #region �N�ǤJ������ন���������

        /// <summary>
        /// �N�ǤJ������ন���������
        /// </summary>
        /// <param name="bufnum">��ƫ��A,
        /// 1 ���P���@ , 2���P���G...7���P����..�H������
        /// </param>
        /// <returns></returns>
        public static string numToWeek(int bufnum)
        {
            switch (bufnum)
            {
                case 1:
                    return "�P���@";
                //break;
                case 2:
                    return "�P���G";
                //break;
                case 3:
                    return "�P���T";
                //break;
                case 4:
                    return "�P���|";
                //break;
                case 5:
                    return "�P����";
                //break;
                case 6:
                    return "�P����";
                //break;
                case 7:
                    return "�P����";
                //break;
                default:
                    return "�ǤJ�ѼƦ��~";
                //break;
            }
        }

        #endregion


        #region ����Ȫ��p��

        /// <summary>
        /// ����Ȫ��p��
        /// </summary>
        /// <param name="InputNum"></param>
        /// <returns>�^�Ǿ�ƭ�</returns>
        public static int CustomAbs(int InputNum)
        {
            int result;
            result = (InputNum >= 0) ? InputNum : -InputNum;
            return result;
        }

        #endregion


        #region �N��Ʈw�������ȥ[��ComboBox�M�椺

        /// <summary>
        /// �N��Ʈw�������ȥ[��ComboBox�M�椺
        /// </summary>
        /// <param name="objCom">ComboBox����</param>
        /// <param name="DefaultUseDB">�ϥθ�Ʈw����</param>
        /// <param name="TableName">���W��</param>
        /// <param name="FieldID">�L�o���W��</param>
        /// <param name="FieldName">�n��J�����W��</param>
        /// <param name="WhereValue">�L�o�����</param>
        public static void GetComboBox(System.Windows.Forms.ComboBox objCom, string DefaultUseDB, string TableName, string FieldID, string FieldName, string WhereValue)
        {
            //cc.GetComboBox(DDL_Department, "SQLServer", "FactoryDept", "FactoryName", "Department", DDL_FactoryClass.SelectedItem.Text);
            string errorMsg = "";
            string selectCmd = "";
            int i;

            SqlConnection conn;
            SqlCommand cmd;
            SqlDataReader dr;

            objCom.Items.Clear();

            switch (TableName)
            {
                case "TableSchmea":
                    selectCmd = "SELECT Distinct " + FieldID + "," + FieldName + " FROM " + TableName + " Where " + FieldID + "='" + WhereValue + "' And IsValid=True  order by " + FieldID;
                    break;
                case "Code":
                    selectCmd = "SELECT Distinct " + FieldID + "," + FieldName + " FROM " + TableName + " Where " + FieldID + "='" + WhereValue + "' order by " + FieldID;
                    break;
                case "Membership":
                    if (WhereValue == "%")
                    {
                        selectCmd = "SELECT Distinct " + FieldID + "," + FieldName + " FROM " + TableName + " Where " + FieldID + " Like'" + WhereValue + "' order by " + FieldName;
                    }
                    break;
                case "SysRole":
                    if (WhereValue == "%")
                    {
                        selectCmd = "SELECT Distinct " + FieldID + "," + FieldName + " FROM " + TableName + " Where " + FieldID + " Like'" + WhereValue + "' order by " + FieldID;
                    }
                    break;
                case "ManufacMan":
                    if (WhereValue == "%")
                    {
                        selectCmd = "SELECT Distinct *  FROM " + TableName + " Where " + FieldID + " Like'" + WhereValue + "' And IsValid=True order by " + FieldID;
                    }
                    break;
                default:
                    break;
            }

            i = 0;

            try
            {

                string ConnString;
                ConnString = My.MyGlobal.SQLConnectionString;
                conn = new SqlConnection(ConnString);
                conn.Open();
                cmd = new SqlCommand(selectCmd, conn);
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    objCom.Items.Add(dr[FieldName].ToString());
                    //objCom.Items[i].Value = Srd[FieldName].ToString();
                    i = i + 1;
                }
                conn.Close();
                dr.Close();
                cmd.Dispose();
                
            }
            catch (Exception ex)
            {

                errorMsg = ex.Message;
            }
        }

        #endregion

    }
}

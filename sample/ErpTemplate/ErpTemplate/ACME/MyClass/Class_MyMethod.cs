using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;//
using System.Security.Cryptography;//�ϥ�����[�K�һݤޥΪ��R�W�Ŷ�

namespace My
{
    public class MyMethod
    {

        #region ���ͬy���s��

        /// <summary>
        /// ���ͬy���s��
        /// �Ҧp:095 12 20 12 53 20
        /// </summary>
        /// <param name="HeadStr">�y�����}�Y�r��</param>
        /// <returns></returns>
        public static string RunID(string HeadStr)
        {
            string IDStr;
            string NowStr;
            string YearStr;
            string MonthStr;
            string DayStr;
            string HourStr;
            string MinuteStr;
            string SecondStr;

            NowStr = DateTime.Now.ToString();
            YearStr = string.Format("{0:D3}", (int)DateTime.Now.Year - 1911); //D3��O�����T��ƫe���|��0
            MonthStr = string.Format("{0:D2}", (int)DateTime.Now.Month);
            DayStr = string.Format("{0:D2}", (int)DateTime.Now.Day);
            HourStr = string.Format("{0:D2}", (int)DateTime.Now.Hour);
            MinuteStr = string.Format("{0:D2}", (int)DateTime.Now.Minute);
            SecondStr = string.Format("{0:D2}", (int)DateTime.Now.Second); ;
            IDStr = HeadStr + YearStr + MonthStr + DayStr + HourStr + MinuteStr + SecondStr;
            return IDStr;
        }

        #endregion


        #region ���ͻ{�ҽX

        /// <summary>
        /// ���ͻ{�ҽX
        /// </summary>
        /// <param name="WordLen">�M�w���ͶüƽX����</param>
        /// <returns></returns>
        public static string GenerateAuthWord(int WordLen) //����WordLen���ת��@�նüƽX
        {
            int RanValue;
            string bufstr = "";
            Random rnd = new Random(DateTime.Now.Millisecond);

            for (int i = 0; i < WordLen; i++)
            {
                RanValue = (int)rnd.Next(0, 9);
                bufstr = bufstr + RanValue.ToString();
            }
            return bufstr;
        }

        #endregion


        #region �N����ഫ���Ѽ�

        /// <summary>
        /// �N����ഫ���Ѽ�
        /// </summary>
        /// <param name="second">���</param>
        /// <returns>�r��� D��H�p��M��S��</returns>
        public static string SecToDay(Int32 second)
        {
            int D, H, M, S;
            string bufstr;
            D = (int)(second / (60 * 60 * 24));
            H = (second / 3600) % 24;
            M = (int)((second % 3600) / 60);
            S = second % 60;
            bufstr = D + "��" + H + "�p��" + M + "��" + S + "��";
            return bufstr;
        }

        #endregion


        #region ��޸��ഫ�����޸�

        /// <summary>
        /// ��޸��ഫ�����޸�
        /// </summary>
        /// <param name="Inputstr">��J�i��]�t��޸��r��</param>
        /// <returns></returns>
        public static string quotates(string Inputstr)
        {
            string correctString = Inputstr.Replace("'", "''");
            return correctString;
        }

        #endregion


        #region �p��ɶ��t

        /// <summary>
        /// �p��ɶ��t
        /// �ΪkDateTimeDiff("2006�~4��1�� 18:00:00")
        /// </summary>
        /// <param name="EndDate">�]�w�פ�ɶ�</param>
        /// <returns></returns>
        public static int DateTimeDiff(string EndDate)
        {

            DateTime dt = Convert.ToDateTime(EndDate);
            int v1;


            DateTime DT1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            DateTime DT2 = dt;// new DateTime(2005, 10, 6, 18, 0, 0);
            TimeSpan TS1 = DT2.Subtract(DT1);

            v1 = (int)TS1.TotalSeconds;
            return v1;

        }

        #endregion



        #region �N���e�g�J����w���ɮ�

        /// <summary>
        /// �N���e�g�J����w���ɮ�
        /// </summary>
        /// <param name="FileContent">�n�g�J���ɮפ��e</param>
        /// <param name="FileName">�ɮצW��</param>
        public static void WriteContentToFile(string FileContent, string FileName)
        {
            FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.Unicode);
            sw.WriteLine(FileContent);
            sw.Close();
        }

        #endregion



        #region �q�ɮפ�Ū�����e�^�Ǥ��e�r��

        /// <summary>
        /// �q�ɮפ�Ū�����e�^�Ǥ��e�r��
        /// </summary>
        /// <param name="FileName">�ɮצW��</param>
        /// <returns></returns>
        public static string ReadFileToString(string FileName)
        {
            string bufstr = "";
            FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Read);
            StreamReader sr = new StreamReader(fs, Encoding.Unicode);
            sr.BaseStream.Seek(0, SeekOrigin.Begin);

            while (sr.Peek() > -1)
            {
                bufstr += sr.Read().ToString();
            }
            sr.Close();
            return bufstr;

        }

        #endregion



        #region ����[�K,����k�ݭn�f�t�@�ӡ@�줸�հ}�C ��r�ꪺ��k(�n�ۤv���g�{���X)

        /// <summary>
        /// ����[�K,����k�ݭn�f�t�@�ӡ@�줸�հ}�C ��r�ꪺ��k(�n�ۤv���g�{���X)
        /// MD5 �T���K�n5(Message Digest 5 , MD5)
        /// SHA1 �w������t��k(Secure Hashing Algorithm , SHA1)
        /// </summary>
        /// <param name="enCrypType">"MD5"��"SHA1"</param>
        /// <param name="bufstring">���i��[�K�r��</param>
        /// <returns>�^�ǥ[�K�r��</returns>
        public static string HashEncryption(string enCrypType, string bufstring)
        {
            //���k�]�i�H
            //HashAlgorithm sha = new SHA1CryptoServiceProvider(); //����j�p160�줸
            //HashAlgorithm md5 = new MD5CryptoServiceProvider();  //����j�p128�줸

            //����R�W�Ŷ�
            //System.Security.Cryptography.HashAlgorithm
            //System.Security.Cryptography.MD5 
            //System.Security.Cryptography.MD5CryptoServiceProvider(); 

            try
            {
                if (enCrypType == "MD5")
                {
                    MD5 md5 = new MD5CryptoServiceProvider();
                    byte[] dataArray = Encoding.UTF8.GetBytes(bufstring);
                    byte[] result = md5.ComputeHash(dataArray);
                    return byteArrayToString(result);
                }
                else if (enCrypType == "SHA1")
                {
                    SHA1 sha1 = new SHA1CryptoServiceProvider();
                    byte[] dataArray = Encoding.UTF8.GetBytes(bufstring);
                    byte[] result = sha1.ComputeHash(dataArray);
                    return byteArrayToString(result);
                }
                else
                {
                    return "error:�i��O�[�K���A���~";
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("���~�T��" + ex.Message.ToString(), "�o�ͨҥ~");
                return ex.Message.ToString();
            }




        }

        #endregion



        #region �N�줸�հ}�C���e�ഫ����r

        /// <summary>
        /// �N�줸�հ}�C���e�ഫ����r
        /// </summary>
        /// <param name="buf">�ǤJ�줸�հ}�C</param>
        /// <returns>��r</returns>
        public static string byteArrayToString(byte[] buf)
        {
            string result = "";
            foreach (byte var in buf)
            {
                result = result + var.ToString();
            }
            return result;
        }

        #endregion


        #region "����B�z"

        /// <summary>
        /// ����B�z
        /// </summary>
        /// <param name="SecInt">�ǤJ���</param>
        public static void DoSomeThing(int SecInt)
        {
            for (int i = 0; i < SecInt; i++)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion


        #region �p��ɶ��t

        public enum DateInterval
        {
            Second, Minute, Hour, Day
        }


        /// <summary>
        /// �p��ɶ��t
        /// </summary>
        /// <param name="StartDate">�_�l���</param>
        /// <param name="EndDate">�פ���</param>
        /// <param name="DateInterval">���w�^�ǳ��
        /// Second ��
        /// Minute ��
        /// Hour ��
        /// Day ��
        /// </param>
        /// <returns></returns>
        public static int DateDiff(DateTime StartDate, DateTime EndDate, DateInterval DI)
        {

            int v1;
            TimeSpan TS1 = EndDate.Subtract(StartDate);

            switch ((int)DI)
            {
                case (int)DateInterval.Second:
                    v1 = (int)TS1.TotalSeconds;
                    break;
                case (int)DateInterval.Minute:
                    v1 = (int)TS1.TotalMinutes;
                    break;
                case (int)DateInterval.Hour:
                    v1 = (int)TS1.TotalHours;
                    break;
                case (int)DateInterval.Day:
                    v1 = (int)TS1.TotalDays;
                    break;
                default:
                    v1 = (int)TS1.TotalSeconds;
                    break;
            }


            return v1;

        }

        #endregion


        #region "�r���PASCII�X���ഫ"

        public static char Chr(int N)
        {
            char C = Convert.ToChar(N);
            return C;
        }

        public static int ASC(string S)
        {
            int N = Convert.ToInt32(S[0]);
            return N;
        }

        public static int ASC(char C)
        {
            int N = Convert.ToInt32(C);
            return N;
        }

        #endregion


        #region "�ƭȧP�_"

        /// <summary>
        /// �ƭȧP�_
        /// </summary>
        /// <param name="num">�ǤJ�Ʀr�r��</param>
        /// <returns>�Y���ƭȫ��A�h�^��true,�_�h�^��false</returns>
        public bool IsNumeric(string num)
        {
            char c;
            bool symbol = false;
            string newNum = num.Trim();

            for (int i = 0; i < newNum.Length; i++)
            {
                c = Convert.ToChar(newNum.Substring(i, 1));
                if (char.IsNumber(c) == false)
                {
                    if (c == '.')
                    {
                        if (symbol == false)
                        {
                            symbol = true;
                        }
                        else
                        {
                            return false;
                        }

                    }
                    else
                    {
                        if (c == '+' || c == '-')
                        {
                            if (i != 0)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }

                }

            }
            return true;
        }


        #endregion

        
    }
}

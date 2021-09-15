using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;//�ϥ�ArrayList�һݥ[�J���R�W�Ŷ�

namespace My
{
    public class MyStrings
    {
        #region "�h���S�w�r��"

        /// <summary>
        /// �h���S�w�r��
        /// �Ҧp: aa  bc  de
        /// ���浲�G: aabcde
        /// toRidSpecChar(textBox1.Text," ")
        /// </summary>
        /// <param name="bufstr">�B�z�r��</param>
        /// <param name="toRidChar">�h���r�ꪺ�r��</param>
        /// <returns></returns>
        public static string toRidSpecChar(string bufstr, string toRidChar)
        {
            string newstr = "";
            int index = bufstr.IndexOf(toRidChar);
            do
            {
                newstr = newstr + bufstr.Substring(0, index);
                bufstr = bufstr.Substring(index);
                bufstr = bufstr.Trim();
                index = bufstr.IndexOf(toRidChar);
            } while (index != -1);

            return newstr + bufstr;

        }

        #endregion


        #region "�N�S�w�r�����j����s�JArrayList"

        /// <summary>
        /// �N�S�w�r�����j����s�JArrayList
        /// SplitSpecCharToArrayList((textBox1.Text).ToString()," ")[1].ToString()
        /// </summary>
        /// <param name="bufstr"></param>
        /// <param name="toRidChar"></param>
        /// <returns></returns>
        public static ArrayList SplitSpecCharToArrayList(string bufstr, string toRidChar)
        {
            ArrayList aList = new ArrayList();
            int i = 0;
            int index = bufstr.IndexOf(toRidChar);
            do
            {
                aList.Add(bufstr.Substring(0, index));
                bufstr = bufstr.Substring(index);
                bufstr = bufstr.Trim();
                index = bufstr.IndexOf(toRidChar);
                i++;
            } while (index != -1);

            aList.Add(bufstr);
            return aList;

        }

        #endregion


        #region "�N�S�w�r�����j����s�JArray"

        /// <summary>
        /// �N�S�w�r�����j����s�JArray
        /// SplitSpecCharToArray(textBox1.Text," ")[0].ToString()
        /// </summary>
        /// <param name="bufstr">�B�z�r��</param>
        /// <param name="toRidChar">�h���r�ꤤ���r��</param>
        /// <returns></returns>
        public static string[] SplitSpecCharToArray(string bufstr, string toRidChar)
        {
            string[] bufarray = new string[bufstr.Length];
            int i = 0;
            int index = bufstr.IndexOf(toRidChar);
            do
            {
                bufarray[i] = bufstr.Substring(0, index);
                bufstr = bufstr.Substring(index);
                bufstr = bufstr.Trim();
                index = bufstr.IndexOf(toRidChar);
                i = i + 1;
            } while (index != -1);

            bufarray[i + 1] = bufstr;
            return bufarray;

        }

        #endregion


        #region "��Ʀr�e���ɤW�Ů�\���k"

        /// <summary>
        /// ��Ʀr�e���ɤW�Ů�\���k
        ///  8
        ///  9
        /// 10
        /// </summary>
        /// <param name="bufNum">�ǤJ�Ʀr</param>
        /// <param name="blankspaceNum">�һݱ���Ʀr�`����</param>
        /// <returns></returns>
        public static string patchBlankSpace(int bufNum, int blankspaceNum)
        {
            int bufNumLength = bufNum.ToString().Length;

            if (bufNumLength >= blankspaceNum)
            {
                return bufNum.ToString();
            }
            else
            {
                string resultString = "";

                for (int i = 0; i < (blankspaceNum - bufNum.ToString().Length); i++)
                {
                    resultString = resultString + " ";
                }

                resultString = resultString + bufNum.ToString();
                return resultString;
            }

        }

        #endregion


        #region "�b��r�᭱�ɤW�Ů�"

        /// <summary>
        /// �b��r�᭱�ɤW�Ů�
        /// </summary>
        /// <param name="bufString">�ǤJ�r��</param>
        /// <param name="blankspaceNum">����r�����</param>
        /// <returns></returns>
        public static string patchBlankSpaceForString(string bufString, int blankspaceNum)
        {
            int bufNumLength = bufString.Length;

            if (bufNumLength >= blankspaceNum)
            {
                return bufString;
            }
            else
            {
                string resultString = "";

                for (int i = 0; i < (blankspaceNum - bufString.Length); i++)
                {
                    resultString = resultString + " ";
                }

                resultString = bufString + resultString;
                return resultString;
            }

        }

        #endregion
    }
}

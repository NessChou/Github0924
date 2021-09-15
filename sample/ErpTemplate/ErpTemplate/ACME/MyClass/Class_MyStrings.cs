using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;//ㄏノArrayList┮惠㏑丁

namespace My
{
    public class MyStrings
    {
        #region "埃疭﹚じ"

        /// <summary>
        /// 埃疭﹚じ
        /// ㄒ: aa  bc  de
        /// 磅︽挡狦: aabcde
        /// toRidSpecChar(textBox1.Text," ")
        /// </summary>
        /// <param name="bufstr">矪瞶﹃</param>
        /// <param name="toRidChar">埃﹃じ</param>
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


        #region "盢疭﹚じだ筳ぇArrayList"

        /// <summary>
        /// 盢疭﹚じだ筳ぇArrayList
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


        #region "盢疭﹚じだ筳ぇArray"

        /// <summary>
        /// 盢疭﹚じだ筳ぇArray
        /// SplitSpecCharToArray(textBox1.Text," ")[0].ToString()
        /// </summary>
        /// <param name="bufstr">矪瞶﹃</param>
        /// <param name="toRidChar">埃﹃いじ</param>
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


        #region "计玡干よ猭"

        /// <summary>
        /// 计玡干よ猭
        ///  8
        ///  9
        /// 10
        /// </summary>
        /// <param name="bufNum">肚计</param>
        /// <param name="blankspaceNum">┮惠北计羆</param>
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


        #region "ゅ干"

        /// <summary>
        /// ゅ干
        /// </summary>
        /// <param name="bufString">肚﹃</param>
        /// <param name="blankspaceNum">北﹃</param>
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

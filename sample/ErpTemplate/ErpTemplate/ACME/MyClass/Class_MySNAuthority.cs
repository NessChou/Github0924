using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;//新增命名空間

namespace My
{
    public class MySNAuthority
    {

        /// <summary>
        /// 產生註冊序號
        /// </summary>
        /// <param name="UserName">傳入使用者名稱</param>
        /// <returns>回傳字串序號</returns>
        public string GenerateKey(string UserName)
        {
            string CodeA, CodeB, CodeC, CodeD;

            if (UserName.Length < 3)
            {
                return "序號產生失敗,使用者名稱長度不可小於3";
            }
            else
            {
                CodeA = "SM" + AutoGenerateWord(3);
                CodeB = GetOSCode(My.Computer.Info.OSFullName);
                CodeC = UserNameCode(UserName);
                CodeD = CheckCode(CodeA, CodeB, CodeC);

                return CodeA + "-" + CodeB + "-" + CodeC + "-" + CodeD;
            }

       
        }

        /// <summary>
        /// 產生n個字元(0~9或a~z或A~Z)
        /// </summary>
        /// <param name="WordLen">字串長度</param>
        /// <returns></returns>
        public string AutoGenerateWord(int WordLen)
        {
            int RanValue = 0;
            string bufstr = "";

            Random rnd = new Random(DateTime.Now.Millisecond);


            for (int i = 0; i < WordLen; i++)
            {
                while (bufstr.Length != WordLen)
                {
                    RanValue = (int)rnd.Next(48, 122);
                    if ((RanValue >= 48 && RanValue <= 57) || (RanValue >= 65 && RanValue <= 90) || (RanValue >= 97 && RanValue <= 122))
                    {
                        bufstr = bufstr + Strings.Chr(RanValue);

                    }

                }

            }
            return bufstr;

        }

        /// <summary>
        /// 獲取作業系統代碼
        /// </summary>
        /// <param name="OperatingSystem">傳入作業系統名稱</param>
        /// <returns></returns>
        public string GetOSCode(string OperatingSystem)
        {
            string result = "";
            switch (OperatingSystem)
            {
                case "Microsoft Windows XP Professional":
                    result = "B39D3";
                    break;
                case "Microsoft Windows XP Home":
                    result = "B39D3";
                    break;
                case "Microsoft Windows 2000 Professional":
                    result = "CCW67";
                    break;
                case "Microsoft Windows 2000 Server":
                    result = "CCW67";
                    break;
                case "Microsoft Windows 2003":
                    result = "GFCW4";
                    break;
                case "Microsoft Windows ME":
                    result = "GWQY3";
                    break;
                default:
                    break;
            }
            return result;

        }




        #region "檢查ＳＮ序號"

        /// <summary>
        /// 檢查ＳＮ序號
        /// </summary>
        /// <param name="RegUserName">註冊的使用者名稱</param>
        /// <param name="SerialNumber">序號</param>
        /// <returns></returns>
        public bool checkSN(string RegUserName, string SerialNumber)
        {

            string[] bufcode = SerialNumber.Split('-');

            if (bufcode.Length != 4)
            {
                return false;
            }


            //step 1
            if (bufcode[0].Substring(0, 2) != "SM")
            {
                return false;
            }



            //step 2
            Microsoft.VisualBasic.Devices.ComputerInfo ComputerInfo = new Microsoft.VisualBasic.Devices.ComputerInfo();

            switch (ComputerInfo.OSFullName.ToString())
            {
                case "Microsoft Windows XP Professional":
                    if (bufcode[1] != "B39D3")
                    {
                        return false;
                    }
                    break;
                case "Microsoft Windows XP Home":
                    if (bufcode[1] != "B39D3")
                    {
                        return false;
                    }
                    break;
                case "Microsoft Windows 2000 Professional":
                    if (bufcode[1] != "CCW67")
                    {
                        return false;
                    }
                    break;
                case "Microsoft Windows 2000 Server":
                    if (bufcode[1] != "CCW67")
                    {
                        return false;
                    }
                    break;
                case "Microsoft Windows 2003":
                    if (bufcode[1] != "GFCW4")
                    {
                        return false;
                    }
                    break;
                case "Microsoft Windows ME":
                    if (bufcode[1] != "GWQY3")
                    {
                        return false;
                    }
                    break;
                default:
                    return false;
            }

            //step 3

            if (UserNameCode(RegUserName) != bufcode[2])
            {
                return false;
            }


            //step 4
            if (CheckCode(bufcode[0], bufcode[1], bufcode[2]) != bufcode[3])
            {
                return false;
            }

            return true; //若全部都符合則傳回True

        }

        #endregion


        #region "將傳入的使用者名稱進行編碼動作"

        /// <summary>
        /// 將傳入的使用者名稱進行編碼動作
        /// 使用者名稱UserName至少長度為5,最多為15
        /// </summary>
        /// <param name="UserName">使用者名稱</param>
        /// <returns></returns>
        public string UserNameCode(string UserName)
        {
            int[] bufarr = new int[15];
            int[] encode = new int[5];
            string NewCode = "";
            string NewUserName = UserName.ToUpper();


            for (int i = 0; i < UserName.Length; i++)
            {
                bufarr[i] = Strings.Asc(NewUserName.Substring(i, 1));

                encode[i % 5] = encode[i % 5] + bufarr[i];
                //MessageBox.Show(bufarr[i].ToString());
            }

            if (UserName.Length == 3)
            {
                encode[3] = encode[0] + encode[1];
                encode[4] = encode[2] + encode[3];
            }

            if (UserName.Length == 4)
            {
                encode[4] = encode[0] + encode[1] + encode[2] + encode[3];
            }


            for (int j = 0; j < 5; j++)
            {


                switch (j)
                {
                    case 0:
                        encode[j] = encode[j] + 1;
                        break;
                    case 1:
                        encode[j] = encode[j] + 3;
                        break;
                    case 2:
                        encode[j] = encode[j] + 5;
                        break;
                    case 3:
                        encode[j] = encode[j] + 2;
                        break;
                    case 4:
                        encode[j] = encode[j] + 4;
                        break;
                    default:
                        break;
                }

                //轉換成大寫字母範圍

                encode[j] = (encode[j] % 26) + 65;

                NewCode = NewCode + Strings.Chr(encode[j]);

            }

            return NewCode;

        }

        #endregion


        #region "檢查碼驗證"

        /// <summary>
        /// 檢查碼驗證
        /// </summary>
        /// <param name="PID">產品編號</param>
        /// <param name="OS">作業系統編號</param>
        /// <param name="UNC">使用者名稱編號</param>
        /// <returns></returns>
        public string CheckCode(string PID, string OS, string UNC)
        {
            int i = 0;

            int[] chkcode = new int[5];

            string bufstring = "";

            for (i = 0; i < 5; i++)
            {
                //將PID第n個字元與OS第n個字元與UNC第n個字元的ASCII碼相加
                chkcode[i] = Strings.Asc(PID.Substring(i, 1)) +
                             Strings.Asc(OS.Substring(i, 1)) +
                             Strings.Asc(UNC.Substring(i, 1));

                //轉換成大寫字母範圍
                chkcode[i] = (chkcode[i] % 26) + 65;

                bufstring = bufstring + Strings.Chr(chkcode[i]);
            }
            return bufstring;

        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    class ValidateUtils
    {

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

        public static bool IsDateString(string sDate)
        {

            try
            {
                UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
                UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
                UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));
                return true;
            }
            catch
            {
                return false;
            }


        }

        public static bool CheckDate(TextBox t)
        {

            if (t.Text.Trim() == "")
                return true;

            try
            {
                StrToDate(t.Text);
                return true;
            }
            catch
            {
                MessageBox.Show("¤é´Á¿é¤J¿ù»~");
                t.SelectAll();
                t.Focus();
                return false;
            }

        }

        public static bool IsNumeric(string number)
        {
            try
            {
                int.Parse(number);
                return true;
            }
            catch
            {
                return false;
            }
        }


    }
}

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Globalization;

namespace ACME
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
     {
         //using System.Globalization;

         //處理日期格式

         CultureInfo culture = new CultureInfo("zh-tw", false);

         culture.DateTimeFormat.DateSeparator = string.Empty;

         culture.DateTimeFormat.ShortDatePattern = "yyyyMMdd";

         culture.DateTimeFormat.LongDatePattern = "yyyyMMdd";

         System.Threading.Thread.CurrentThread.CurrentCulture = culture;

         System.Threading.Thread.CurrentThread.CurrentUICulture = culture;


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // Form1 測試用
          //  Application.Run(new Form1());
            //登入視窗
            fmLogin LoadForm = new fmLogin();
            //登入成功
      if (LoadForm.ShowDialog() == DialogResult.OK)
            //主畫面
            Application.Run(new MDIfrmMain());
        //    else
            //    Application.Exit();
        }
    }
}
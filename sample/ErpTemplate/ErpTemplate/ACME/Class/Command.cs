using System;
using System.Collections.Generic;
using System.Text;

namespace ACME
{
    class Command
    {

        /// <summary>
        /// 斷行
        /// </summary>
        public static byte[] NewLine = new byte[] { 10 };

        /// <summary>
        /// 跳頁
        /// </summary>
        public static byte[] NewPage = new byte[] { 12 };

        /// <summary>
        /// 立即切紙(剩1點沒切斷)
        /// </summary>
        public static byte[] CutReserve1 = new byte[] { 29, 86, 1 };

        /// <summary>
        /// 立即切紙(剩3點沒切斷)
        /// </summary>
        public static byte[] CutReserve3 = new byte[] { 29, 86, 2 };

        /// <summary>
        /// 跳到下一頁的開頭並切紙(剩1點沒切斷)
        /// </summary>
        public static byte[] NewPage_CutReserve1 = new byte[] { 29, 86, 66 };

        /// <summary>
        /// 跳到下一頁的開頭並切紙(剩3點沒切斷)
        /// </summary>
        public static byte[] NewPage_CutReserve3 = new byte[] { 29, 86, 67 };

        /// <summary>
        /// 將印字頭移到印字頭現在所在列的開頭處(沒有移到下一列喔!)
        /// <para>printer.WriteLine("test");</para>
        /// <para>printer.Write(Command.BackToFirst);</para>
        /// <para>printer.WriteLine("----");</para>
        /// <para>可得到刪除線效果</para>
        /// </summary>
        public static byte[] BackToFirst = new byte[] { 13 };

        /// <summary>
        /// 字元 2 倍寬
        /// </summary>
        public static byte[] DoubleWidth = new byte[] { 27, 33, 32 };

        /// <summary>
        /// 底線模式
        /// </summary>
        public static byte[] Underline = new byte[] { 27, 33, 128 };

        /// <summary>
        /// 取消「字元 2 倍寬」、「底線模式」
        /// </summary>
        public static byte[] Cancel = new byte[] { 27, 33, 0 };

        /// <summary>
        /// 初始化印表機(使印表機回復到剛開機時的狀態)
        /// </summary>
        public static byte[] ResetPrinter = new byte[] { 27, 64 };

        /// <summary>
        /// 選擇輸出目的地(僅印在存根聯)
        /// </summary>
        public static byte[] OnlyStub = new byte[] { 27, 99, 48, 1 };

        /// <summary>
        /// 選擇輸出目的地(僅印在收執聯)
        /// </summary>
        public static byte[] OnlyReceiver = new byte[] { 27, 99, 48, 2 };

        /// <summary>
        /// 選擇輸出目的地(印在存根聯及收據聯 - 內定值)
        /// </summary>
        public static byte[] StubAndReceiver = new byte[] { 27, 99, 48, 3 };

        /// <summary>
        /// 將資料印出並將印字頭往下移動 n 列
        /// </summary>
        public static byte[] MoveLines(byte n)
        {
            return new byte[] { 27, 100, n };
        }

        /// <summary>
        /// 在收執聯上蓋店章
        /// </summary>
        public static byte[] PrintMark = new byte[] { 27, 111 };

        /// <summary>
        /// 開啟錢櫃1
        /// </summary>
        public static byte[] OpenMoneyBox1 = new byte[] { 27, 112, 0, 50, 250 };

        /// <summary>
        /// 開啟錢櫃2
        /// </summary>
        public static byte[] OpenMoneyBox2 = new byte[] { 27, 112, 1, 50, 250 };


    }


}

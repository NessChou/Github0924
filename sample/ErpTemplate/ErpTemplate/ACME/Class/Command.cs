using System;
using System.Collections.Generic;
using System.Text;

namespace ACME
{
    class Command
    {

        /// <summary>
        /// �_��
        /// </summary>
        public static byte[] NewLine = new byte[] { 10 };

        /// <summary>
        /// ����
        /// </summary>
        public static byte[] NewPage = new byte[] { 12 };

        /// <summary>
        /// �ߧY����(��1�I�S���_)
        /// </summary>
        public static byte[] CutReserve1 = new byte[] { 29, 86, 1 };

        /// <summary>
        /// �ߧY����(��3�I�S���_)
        /// </summary>
        public static byte[] CutReserve3 = new byte[] { 29, 86, 2 };

        /// <summary>
        /// ����U�@�����}�Y�ä���(��1�I�S���_)
        /// </summary>
        public static byte[] NewPage_CutReserve1 = new byte[] { 29, 86, 66 };

        /// <summary>
        /// ����U�@�����}�Y�ä���(��3�I�S���_)
        /// </summary>
        public static byte[] NewPage_CutReserve3 = new byte[] { 29, 86, 67 };

        /// <summary>
        /// �N�L�r�Y����L�r�Y�{�b�Ҧb�C���}�Y�B(�S������U�@�C��!)
        /// <para>printer.WriteLine("test");</para>
        /// <para>printer.Write(Command.BackToFirst);</para>
        /// <para>printer.WriteLine("----");</para>
        /// <para>�i�o��R���u�ĪG</para>
        /// </summary>
        public static byte[] BackToFirst = new byte[] { 13 };

        /// <summary>
        /// �r�� 2 ���e
        /// </summary>
        public static byte[] DoubleWidth = new byte[] { 27, 33, 32 };

        /// <summary>
        /// ���u�Ҧ�
        /// </summary>
        public static byte[] Underline = new byte[] { 27, 33, 128 };

        /// <summary>
        /// �����u�r�� 2 ���e�v�B�u���u�Ҧ��v
        /// </summary>
        public static byte[] Cancel = new byte[] { 27, 33, 0 };

        /// <summary>
        /// ��l�ƦL���(�ϦL����^�_���}���ɪ����A)
        /// </summary>
        public static byte[] ResetPrinter = new byte[] { 27, 64 };

        /// <summary>
        /// ��ܿ�X�ت��a(�ȦL�b�s���p)
        /// </summary>
        public static byte[] OnlyStub = new byte[] { 27, 99, 48, 1 };

        /// <summary>
        /// ��ܿ�X�ت��a(�ȦL�b�����p)
        /// </summary>
        public static byte[] OnlyReceiver = new byte[] { 27, 99, 48, 2 };

        /// <summary>
        /// ��ܿ�X�ت��a(�L�b�s���p�Φ����p - ���w��)
        /// </summary>
        public static byte[] StubAndReceiver = new byte[] { 27, 99, 48, 3 };

        /// <summary>
        /// �N��ƦL�X�ñN�L�r�Y���U���� n �C
        /// </summary>
        public static byte[] MoveLines(byte n)
        {
            return new byte[] { 27, 100, n };
        }

        /// <summary>
        /// �b�����p�W�\����
        /// </summary>
        public static byte[] PrintMark = new byte[] { 27, 111 };

        /// <summary>
        /// �}�ҿ��d1
        /// </summary>
        public static byte[] OpenMoneyBox1 = new byte[] { 27, 112, 0, 50, 250 };

        /// <summary>
        /// �}�ҿ��d2
        /// </summary>
        public static byte[] OpenMoneyBox2 = new byte[] { 27, 112, 1, 50, 250 };


    }


}

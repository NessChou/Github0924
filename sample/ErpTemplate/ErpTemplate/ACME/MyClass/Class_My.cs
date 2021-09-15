using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;
//My.Application �һݤޥΩR�W�Ŷ�  My.User����
using Microsoft.VisualBasic.Logging; //My.Application.Log�һݤޥΩR�W�Ŷ�
using Microsoft.VisualBasic.Devices;
//My.Computer.Info ���� My.Computer.Keyboard ���� My.Computer.Mouse ���� 
//My.Computer.Network ���� My.Computer.Ports ����
using Microsoft.VisualBasic.MyServices;
//My.Computer.FileSystem ���� My.Computer.Registry ����

///****** C# My Object Implement ********
///Step 1 �[�J�Ѧ�[Microsoft.VisualBasic.dll]
///Step 2 ��@My�R�W�Ŷ�
namespace My
{
    //My.Application ���󴣨ѻP�ثe���ε{���������ݩʡB��k�M�ƥ�C
    public class Application
    {
        static Application()
        {
            MyApplication = new WindowsFormsApplicationBase();
        }

        //WindowsFormsApplicationBase�i�䴩�Ҧ�MyApplication���s��
        public readonly static WindowsFormsApplicationBase MyApplication;

        //My.Application.Log ���󴣨Ѥ@���ݩ� (Property) �M�h�Ӥ�k�A
        //�ΥH�N�ƥ�M�ҥ~���p��T�g�J���ε{�����O���ɱ�ť�{���C
        public static Log Log
        {

            get
            {
                Log Log = new Log();
                return Log;
            }
        }

        //My.Application.Info ����|�����ݩ� (Property)�A�Ω���o���ε{���ե�
        //��������T�A�Ҧp�������X�B�y�z�M���J���ե󵥵��C
        public static AssemblyInfo Info
        {
            get
            {
                AssemblyInfo Info = new
                    AssemblyInfo(System.Reflection.Assembly.GetExecutingAssembly());
                return Info;
            }
        }

    }

    //��@My.Computer
    public class Computer
    {
        static Computer()
        {
            //My.Computer.Info ���� ���Ѧs���q���O����,���J�ե�P�@�~�t�θ�T��
            ComputerInfo = new ComputerInfo();
        }

        public readonly static ComputerInfo ComputerInfo;

        //My.Computer.Audio ���� ���Ѽ��񭵮Ī��ݩʩM��k
        public static Audio Audio
        {
            get
            {
                Audio Ado = new Audio();
                return Ado;
            }
        }

        //My.Computer.Clock ���� ���Ѧs���ثe�������ɶ��M��ڼзǮɶ�
        public static Clock Clock
        {
            get
            {
                Clock clk = new Clock();
                return clk;
            }
        }

        //My.Computer.Info ���� ���Ѧs���q���O����,���J�ե�P�@�~�t�θ�T��
        public static ComputerInfo Info
        {
            get
            {
                ComputerInfo CInfo = new ComputerInfo();
                return CInfo;
            }
        }

        //My.Computer.Keyboard ���� ���Ѧs����L�ثe���A
        public static Keyboard Keyboard
        {
            get
            {
                Keyboard Keyb = new Keyboard();
                return Keyb;
            }
        }

        //My.Computer.Mouse ���� ���Ѧs�������ƹ��պA��T
        public static Mouse Mouse
        {
            get
            {
                Mouse Mos = new Mouse();
                return Mos;
            }
        }

        //My.Computer.Network ���� ���ѻP�q���ҳs�������������ݩʩM��k
        public static Network Network
        {
            get
            {
                Network Net = new Network();
                return Net;
            }
        }

        //My.Computer.Ports ���� �N�r��ǰe�ܹq�����ǦC��C
        public static Ports Ports
        {
            get
            {
                Ports Port = new Ports();
                return Port;
            }
        }

        //My.Computer.Clipboard ���� ���Ѻ޲z[�ŶKï]����k
        public static ClipboardProxy Clipboard
        {
            get
            {
                GetReturn MyReturn = new GetReturn();
                return MyReturn.ClipboardProxy;
            }
        }

        //My.Computer.FileSystem ���� ���ѹ�ϺХؿ��ɮת��s����k�P�ݩ�
        public static FileSystemProxy FileSystem
        {
            get
            {
                GetReturn MyReturn = new GetReturn();
                return MyReturn.FileSystemProxy;
            }
        }

        //My.Computer.Registry ���� �����ݩ� (Property) �M��k�Ӿާ@�n���C
        public static RegistryProxy Registry
        {
            get
            {
                GetReturn MyReturn = new GetReturn();
                return MyReturn.RegistryProxy;
            }
        }

        //����q���W��
        public static String Name
        {
            get
            {
                Microsoft.VisualBasic.Devices.Computer ComputerName = new Microsoft.VisualBasic.Devices.Computer();
                return ComputerName.Name;
            }
        }

        //����q���D�n��ܵe��
        public static System.Windows.Forms.Screen Screen
        {
            get
            {
                Microsoft.VisualBasic.Devices.Computer ComputerName = new Microsoft.VisualBasic.Devices.Computer();
                return ComputerName.Screen;
            }

        }


    }

    //��@My.User
    public class User
    {

        private static string UserName;
        private static int IndexPath;
        public static string Name
        {
            get
            {
                Microsoft.VisualBasic.ApplicationServices.User NowUser =
                    new Microsoft.VisualBasic.ApplicationServices.User();

                NowUser.InitializeWithWindowsUser();
                IndexPath = NowUser.Name.IndexOf("\\");
                UserName = NowUser.Name.Substring(IndexPath + 1);
                return UserName;
            }
        }
    }

    //���v�ŧi
    public class Copyrights
    {
        public static string GetCopyrightInfo
        {
            get
            {
                return "Copyright(c) 2006-2007 by Ching-Rung Shiu. All Rights Reserved.";

            }
        }

        public static string Author
        {
            get
            {
                return "Ching-Rung Shiu(�\�M�a)";

            }
        }

        public static void SayHello()
        {
            MessageBox.Show("Hello World!!");
        }

    }

}


class GetReturn
{
    public ClipboardProxy ClipboardProxy
    {
        get
        {
            Microsoft.VisualBasic.Devices.Computer ComputerName = new Microsoft.VisualBasic.Devices.Computer();
            return ComputerName.Clipboard;
        }

    }

    public FileSystemProxy FileSystemProxy
    {
        get
        {
            Microsoft.VisualBasic.Devices.Computer ComputerName = new Microsoft.VisualBasic.Devices.Computer();
            return ComputerName.FileSystem;
        }
    }

    public RegistryProxy RegistryProxy
    {
        get
        {
            Microsoft.VisualBasic.Devices.Computer ComputerName = new Microsoft.VisualBasic.Devices.Computer();
            return ComputerName.Registry;
        }
    }


}
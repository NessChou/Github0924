using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.IO;


namespace ACME
{

   

    public partial class LAB : Form
    {

        //
        private string Company_Man;
        private string Tel;
        private string Addr;
    
        //
        const uint IMAGE_BITMAP = 0;
        const uint LR_LOADFROMFILE = 16;
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern IntPtr LoadImage(IntPtr hinst, string lpszName, uint uType,
           int cxDesired, int cyDesired, uint fuLoad);
        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int DeleteObject(IntPtr ho);

         string szSavePath = "C:\\Argox";
         string szSaveFile = "C:\\Argox\\PPLA_Example.Prn";


        const string sznop1     = "nop_front\r\n";
        const string sznop2     = "nop_middle\r\n";
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_Maxi(int x, int y, int primary, int secondary, 
            int country, int service, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_Maxi_Ori(int x, int y, int ori, int primary, 
            int secondary, int country, int service, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_PDF417(int x, int y, int narrow, int width, char normal, 
            int security, int aspect, int row, int column, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_PDF417_Ori(int x, int y, int ori, int narrow, int width, 
            char normal, int security, int aspect, int row, int column, char mode, int numeric, 
            string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_DataMatrix(int x, int y, int rotation, int hor_mul, 
            int ver_mul, int ECC, int data_format, int num_rows, int num_col, char mode, 
            int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern void A_Clear_Memory();
        [DllImport("Winppla.dll")]
        private static extern void A_ClosePrn();
        [DllImport("Winppla.dll")]
        private static extern int A_CreatePrn(int selection, string filename);
        [DllImport("Winppla.dll")]
        private static extern int A_Del_Graphic(int mem_mode, string graphic);
        [DllImport("Winppla.dll")]
        private static extern int A_Draw_Box(char mode, int x, int y, int width, int height, 
            int top, int side);
        [DllImport("Winppla.dll")]
        private static extern int A_Draw_Line(char mode, int x, int y, int width, int height);
        [DllImport("Winppla.dll")]
        private static extern void A_Feed_Label();
        [DllImport("Winppla.dll")]
        private static extern IntPtr A_Get_DLL_Version(int nShowMessage);
        [DllImport("Winppla.dll")]
        private static extern int A_Get_DLL_VersionA(int nShowMessage);
        [DllImport("Winppla.dll")]
        private static extern int A_Get_Graphic(int x, int y, int mem_mode, char format, 
            string filename);
        [DllImport("Winppla.dll")]
        private static extern int A_Get_Graphic_ColorBMP(int x, int y, int mem_mode, char format, 
            string filename);
        [DllImport("Winppla.dll")]
        private static extern int A_Get_Graphic_ColorBMPEx(int x, int y, int nWidth, int nHeight,
            int rotate, int mem_mode, char format, string id_name, string filename);
        [DllImport("Winppla.dll")]
        private static extern int A_Get_Graphic_ColorBMP_HBitmap(int x, int y, int nWidth, int nHeight,
           int rotate, int mem_mode, char format, string id_name, IntPtr hbm);
        [DllImport("Winppla.dll")]
        private static extern int A_Initial_Setting(int Type, string Source);
        [DllImport("Winppla.dll")]
        private static extern int A_WriteData(int IsImmediate, byte[] pbuf, int length);
        [DllImport("Winppla.dll")]
        private static extern int A_ReadData(byte[] pbuf, int length, int dwTimeoutms);
        [DllImport("Winppla.dll")]
        private static extern int A_Load_Graphic(int x, int y, string graphic_name);
        [DllImport("Winppla.dll")]
        private static extern int A_Open_ChineseFont(string path);
        [DllImport("Winppla.dll")]
        private static extern int A_Print_Form(int width, int height, int copies, int amount, 
            string form_name);
        [DllImport("Winppla.dll")]
        private static extern int A_Print_Out(int width, int height, int copies, int amount);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Barcode(int x, int y, int ori, char type, int narrow, 
            int width, int height, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text(int x, int y, int ori, int font, int type, 
            int hor_factor, int ver_factor, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text_Chinese(int x, int y, int fonttype, string id_name, 
            string data, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text_TrueType(int x, int y, int FSize, string FType, 
            int Fspin, int FWeight, int FItalic, int FUnline, int FStrikeOut, string id_name, 
            string data, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text_TrueType_W(int x, int y, int FHeight, int FWidth, 
            string FType, int Fspin, int FWeight, int FItalic, int FUnline, int FStrikeOut, 
            string id_name, string data, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Backfeed(int back);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_BMPSave(int nSave, string pstrBMPFName);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Cutting(int cutting);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Darkness(int heat);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_DebugDialog(int nEnable);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Feed(char rate);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Form(string formfile, string form_name, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Margin(int position, int margin);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Prncomport(int baud, int parity, int data, int stop);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Prncomport_PC(int nBaudRate, int nByteSize, int nParity, 
            int nStopBits, int nDsr, int nCts, int nXonXoff);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Sensor_Mode(char type, int continuous);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Speed(char speed);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Syssetting(int transfer, int cut_peel, int length, 
            int zero, int pause);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Unit(char unit);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Gap(int gap);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_Logic(int logic);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_ProcessDlg(int nShow);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_ErrorDlg(int nShow);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_LabelVer(int centiInch);
        [DllImport("Winppla.dll")]
        private static extern int A_GetUSBBufferLen();
        [DllImport("Winppla.dll")]
        private static extern int A_EnumUSB(byte[] buf);
        [DllImport("Winppla.dll")]
        private static extern int A_CreateUSBPort(int nPort);
        [DllImport("Winppla.dll")]
        private static extern int A_CreatePort(int nPortType, int nPort, string filename);
        [DllImport("Winppla.dll")]
        private static extern int A_Clear_MemoryEx(int nMode);
        [DllImport("Winppla.dll")]
        private static extern void A_Set_Mirror();
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_RSS(int x, int y, int ori, int ratio, int height, 
            char rtype, int mult, int seg, string data1, string data2);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_QR_M(int x, int y, int ori, char mult, int value, 
            int model, char error, int mask, char dinput, char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_Bar2d_QR_A(int x, int y, int ori, char mult, int value, 
            char mode, int numeric, string data);
        [DllImport("Winppla.dll")]
        private static extern int A_GetNetPrinterBufferLen();
        [DllImport("Winppla.dll")]
        private static extern int A_EnumNetPrinter(byte[] buf);
        [DllImport("Winppla.dll")]
        private static extern int A_CreateNetPort(int nPort);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text_TrueType_Uni(int x, int y, int FSize, string FType, 
            int Fspin, int FWeight, int FItalic, int FUnline, int FStrikeOut, string id_name, 
            byte[] data, int format, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_Prn_Text_TrueType_UniB(int x, int y, int FSize, string FType, 
            int Fspin, int FWeight, int FItalic, int FUnline, int FStrikeOut, string id_name, 
            byte[] data, int format, int mem_mode);
        [DllImport("Winppla.dll")]
        private static extern int A_GetUSBDeviceInfo(int nPort, byte[] pDeviceName, 
            out int pDeviceNameLen, byte[] pDevicePath, out int pDevicePathLen);
        [DllImport("Winppla.dll")]
        private static extern int A_Set_EncryptionKey(string encryptionKey);
        [DllImport("Winppla.dll")]
        private static extern int A_Check_EncryptionKey(string decodeKey, string encryptionKey, 
            int dwTimeoutms);

        public LAB()
        {
            InitializeComponent();
        }

        public void ADD(string Company_ManS, string TelS, string AddrS)
        {
            //Test code start
            // open port.
            int nLen, ret, sw;
            byte[] pbuf = new byte[128];
            string strmsg;
            IntPtr ver;
            System.Text.Encoding encAscII = System.Text.Encoding.ASCII;
            System.Text.Encoding encUnicode = System.Text.Encoding.Unicode;

            // dll version.
            ver = A_Get_DLL_Version(0);

            // search port.
            //nLen = A_GetUSBBufferLen() + 1;
            nLen = A_GetUSBBufferLen() + 2;
            strmsg = "DLL ";
            strmsg += Marshal.PtrToStringAnsi(ver);
            strmsg += "\r\n";
            if (nLen > 1)
            {
                byte[] buf1, buf2;
                int len1 = 128, len2 = 128;
                buf1 = new byte[len1];
                buf2 = new byte[len2];
                A_EnumUSB(pbuf);
                A_GetUSBDeviceInfo(1, buf1, out len1, buf2, out len2);
                sw = 1;
                if (1 == sw)
                {
                    //  ret = A_CreatePrn(12, encAscII.GetString(buf2, 0, len2));// open usb.
                    //ret = A_CreatePrn(13, encAscII.GetString(buf2, 0, len2));// open usb.
                    ret = A_CreatePrn(10, "\\\\10.2.2.116\\cp-3140l");// open usb.
                    //Call A_CreatePrn(10, "\\\\allen\\Label")
                }
                else
                {
                    ret = A_CreateUSBPort(1);// must call A_GetUSBBufferLen() function fisrt.
                }
                if (0 != ret)
                {
                    strmsg += "Open USB fail!";
                }
                else
                {
                    strmsg += "Open USB:\r\nDevice name: ";
                    strmsg += encAscII.GetString(buf1, 0, len1);
                    strmsg += "\r\nDevice path: ";
                    strmsg += encAscII.GetString(buf2, 0, len2);
                    //sw = 2;
                    if (2 == sw)
                    {
                        //get printer status.
                        pbuf[0] = 0x01;
                        pbuf[1] = 0x46;
                        pbuf[2] = 0x0D;
                        pbuf[3] = 0x0A;
                        A_WriteData(1, pbuf, 4);//<SOH>F
                        ret = A_ReadData(pbuf, 2, 1000);
                    }
                }
            }
            else
            {
                System.IO.Directory.CreateDirectory(szSavePath);
                ret = A_CreatePrn(0, szSaveFile);// open file.
                strmsg += "Open ";
                strmsg += szSaveFile;
                if (0 != ret)
                {
                    strmsg += " file fail!";
                }
                else
                {
                    strmsg += " file succeed!";
                }
            }
         
            if (0 != ret)
            {
                MessageBox.Show(strmsg);
                return;
            }
            // sample setting.
            A_Set_DebugDialog(1);
            A_Set_Unit('n');
            A_Set_Syssetting(1, 0, 0, 0, 0);
            A_Set_Darkness(8);
            A_Del_Graphic(1, "*");// delete all picture.
            A_Clear_Memory();// clear memory.
            A_WriteData(0, encAscII.GetBytes(sznop2), sznop2.Length);
            A_WriteData(1, encAscII.GetBytes(sznop1), sznop1.Length);

            int LineLimit = 18;
            int fontSize = 50;
            if (Company_ManS.Length > LineLimit)
            {
                string Company_ManS1 = Company_ManS.Substring(0, LineLimit - 1);
                string Company_ManS2 = Company_ManS.Substring(LineLimit - 1, Company_ManS.Length - LineLimit + 1);

                A_Prn_Text_TrueType(20, 100, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A1", Company_ManS1, 1);
                A_Prn_Text_TrueType(20, 70, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A2", Company_ManS2, 1);
            }
            else
            {
                A_Prn_Text_TrueType(20, 100, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A1", Company_ManS, 1);
            }

            if (Company_ManS.Length > LineLimit)
            {
                A_Prn_Text_TrueType(20, 40, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A3", TelS, 1);
            }
            else
            {
                A_Prn_Text_TrueType(20, 70, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A2", TelS, 1);
            }

            //拆字成兩行..當長度大於...
            if (Company_ManS.Length > LineLimit)
            {
                if (AddrS.Length > LineLimit)
                {
                    string Addr1 = AddrS.Substring(0, LineLimit - 1);
                    string Addr2 = AddrS.Substring(LineLimit - 1, AddrS.Length - LineLimit + 1);

                    A_Prn_Text_TrueType(20, 20, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A4", Addr1, 1);
                    A_Prn_Text_TrueType(20, 0, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A5", Addr2, 1);


                }
                else
                {
                    // A_Prn_Text_TrueType_W(20, 40, 18, 18, "Times New Roman", 1, 400, 0, 0, 0, "A3", AddrS, 1);
                    A_Prn_Text_TrueType(20, 20, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A4", AddrS, 1);
                }
            }
            else

            {
                if (AddrS.Length > LineLimit)
                {
                    string Addr1 = AddrS.Substring(0, LineLimit - 1);
                    string Addr2 = AddrS.Substring(LineLimit - 1, AddrS.Length - LineLimit + 1);

                    A_Prn_Text_TrueType(20, 40, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A3", Addr1, 1);
                    A_Prn_Text_TrueType(20, 20, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A4", Addr2, 1);


                }
                else
                {
                    // A_Prn_Text_TrueType_W(20, 40, 18, 18, "Times New Roman", 1, 400, 0, 0, 0, "A3", AddrS, 1);
                    A_Prn_Text_TrueType(20, 40, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A3", AddrS, 1);
                }
            }


             A_Print_Out(1, 1, 1, 1);// copy 2.

            // close port.
             A_ClosePrn();


        }

      
        private void button1_Click(object sender, EventArgs e)
        {
            //Test code start
            // open port.
            int nLen, ret, sw;
            byte[] pbuf = new byte[128];
            string strmsg;
            IntPtr ver;
            System.Text.Encoding encAscII = System.Text.Encoding.ASCII;
            System.Text.Encoding encUnicode = System.Text.Encoding.Unicode;

            // dll version.
            ver = A_Get_DLL_Version(0);

            // search port.
           //nLen = A_GetUSBBufferLen() + 1;
            nLen = A_GetUSBBufferLen() + 2;
            strmsg = "DLL ";
            strmsg += Marshal.PtrToStringAnsi(ver);
            strmsg += "\r\n";
            if (nLen > 1)
            {
                byte[] buf1, buf2;
                int len1 = 128, len2 = 128;
                buf1 = new byte[len1];
                buf2 = new byte[len2];
                A_EnumUSB(pbuf);
                A_GetUSBDeviceInfo(1, buf1, out len1, buf2, out len2);
	            sw = 1;
	            if (1 == sw) {
		          //  ret = A_CreatePrn(12, encAscII.GetString(buf2, 0, len2));// open usb.
                    //ret = A_CreatePrn(13, encAscII.GetString(buf2, 0, len2));// open usb.
                    ret = A_CreatePrn(10, "\\\\10.2.2.116\\LabelDr2");// open usb.
                     //Call A_CreatePrn(10, "\\\\allen\\Label")
	            }
	            else {
		            ret = A_CreateUSBPort(1);// must call A_GetUSBBufferLen() function fisrt.
	            }
                if (0 != ret) {
                    strmsg += "Open USB fail!";
                }
                else {
                    strmsg += "Open USB:\r\nDevice name: ";
                    strmsg += encAscII.GetString(buf1, 0, len1);
                    strmsg += "\r\nDevice path: ";
                    strmsg += encAscII.GetString(buf2, 0, len2);
                    //sw = 2;
                    if (2 == sw)
                    {
                        //get printer status.
                        pbuf[0] = 0x01;
                        pbuf[1] = 0x46;
                        pbuf[2] = 0x0D;
                        pbuf[3] = 0x0A;
                        A_WriteData(1, pbuf, 4);//<SOH>F
                        ret = A_ReadData(pbuf, 2, 1000);
                    }
                }
            }
            else
            {
                System.IO.Directory.CreateDirectory(szSavePath);
                ret = A_CreatePrn(0, szSaveFile);// open file.
                strmsg += "Open ";
                strmsg += szSaveFile;
                if (0 != ret) {
                    strmsg += " file fail!";
                }
                else {
                    strmsg += " file succeed!";
                }
            }
         //   MessageBox.Show(strmsg);
            if (0 != ret)
            {
                MessageBox.Show(strmsg);
                return;
            }
            // sample setting.
            A_Set_DebugDialog(1);
            A_Set_Unit('n');
            A_Set_Syssetting(1, 0, 0, 0, 0);
            A_Set_Darkness(8);
            A_Del_Graphic(1, "*");// delete all picture.
            A_Clear_Memory();// clear memory.
            A_WriteData(0, encAscII.GetBytes(sznop2), sznop2.Length);
            A_WriteData(1, encAscII.GetBytes(sznop1), sznop1.Length);




        //    int fontSize = 36;
            int fontSize = 50;
            A_Prn_Text_TrueType(20, 100, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A1", Company_Man, 1);
            A_Prn_Text_TrueType(20, 70, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A2", Tel, 1);

            //拆字成兩行..當長度大於...
            int LineLimit = 18;
            if (Addr.Length > LineLimit)
            {
                string Addr1 = Addr.Substring(0, LineLimit-1);
                string Addr2 = Addr.Substring(LineLimit-1, Addr.Length - LineLimit );

                A_Prn_Text_TrueType(20, 40, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A3", Addr1, 1);
                A_Prn_Text_TrueType(20, 20, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A4", Addr2, 1);

                //A_Prn_Text_TrueType_W(20, 40, 18, 18, "Times New Roman", 1, 400, 0, 0, 0, "A3", Addr1, 1);
                //A_Prn_Text_TrueType_W(20, 20, 18, 18, "Times New Roman", 1, 400, 0, 0, 0, "A4", Addr2, 1);
            }
            else
            {
               // A_Prn_Text_TrueType_W(20, 40, 18, 18, "Times New Roman", 1, 400, 0, 0, 0, "A3", Addr, 1);
                A_Prn_Text_TrueType(20, 40, fontSize, "Times New Roman", 1, 400, 0, 0, 0, "A3", Addr, 1);
            }


            // output.
           // A_Print_Out(1, 1, 2, 1);// copy 2.
            A_Print_Out(1, 1, 1, 1);// copy 2.

            // close port.
            A_ClosePrn();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                
                if (dataGridView1.SelectedRows.Count > 0)
                {

                    for (int j = 0; j <= dataGridView1.SelectedRows.Count - 1; j++)
                    {
                 
                        sb.Append(dataGridView1.SelectedRows[j].Cells["DOCENTRY"].Value.ToString() + ",");
                    }
            
      
                    sb.Remove(sb.Length - 1, 1);

                }
  

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            System.Data.DataTable dt = null;
            string T1 = "";
            if (sb.ToString() == "")
            {
                T1 = "請確定是否要列印全部標籤";
                dt = GetAddress("1","");
            }
            else
            {
                T1 = "請確定是否要列印標籤 發票號碼 : "+sb.ToString();
                dt = GetAddress("2",sb.ToString());
            }

        
                                DialogResult result;
                                result = MessageBox.Show(T1, "YES/NO", MessageBoxButtons.YesNo);
                                if (result == DialogResult.Yes)
                                {

                                    if (dt.Rows.Count > 0)
                                    {
                                        for (int j = dt.Rows.Count - 1; j >= 0; j--)
                                        {


                                            Addr = Convert.ToString(dt.Rows[j]["ADDRESS"]);
                                            Company_Man = Convert.ToString(dt.Rows[j]["COMPANY"]);
                                            Tel = Convert.ToString(dt.Rows[j]["TEL"]);


                                        //System.Data.DataTable V = GetAddress2(Addr);
                                        //if (V.Rows.Count > 0)
                                        //{
                                        //    Company_Man = V.Rows[0][0].ToString();
                                        //}

                                            string H1 = Addr.Substring(0, 1);

                                            int n;
                                            string F1 = "";
                                            if (!int.TryParse(H1, out n))
                                            {

                                                DataTable dt2 = GetAddress2();
                                                for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                                                {
                                                    string ZIP = Convert.ToString(dt2.Rows[i]["ZIP"]);

                                                    string ADDRESS = Convert.ToString(dt2.Rows[i]["ADDRESS"]);
                                                    int F2 = Addr.IndexOf(ADDRESS);
                                                  
                                                        if (F2 != -1)
                                                        {
                                                            F1 = ZIP.Trim();
                                                        }
                                                    
                                                }
                                            }
                                            Addr = F1 + Addr;
                                            ADD(Company_Man, Tel, Addr);
                                        }

                                    }


                                }



        }

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("DOCENTRY", typeof(string));
            dt.Columns.Add("ADDRESS", typeof(string));
            dt.Columns.Add("PERSON", typeof(string));
            
            return dt;
        }


        private DataTable GetAddress(string TYPE, string cs)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT DISTINCT CAST(ISNULL(T2.building,T0.CARDNAME) AS VARCHAR)+'-'+CASE WHEN ISNULL(t8.pickrmrk,'') <> '' THEN  ISNULL(t8.pickrmrk,'') ELSE ISNULL(T2.ZIPCODE,'') END+'收'   COMPANY,T2.street+ISNULL(REPLACE(T2.COUNTY,'不同PO請務必分開開發票',''),'') ADDRESS,T2.block TEL     ");
            sb.Append("               FROM OINV T0     ");
            sb.Append("               LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND  T0.PayToCode=T2.[Address]  and T2.adrestype='B')   ");
            sb.Append("               LEFT JOIN  INV1 T1 on (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("               LEFT JOIN  dln1 t4 on (T1.baseentry=T4.docentry and  T1.baseline=t4.linenum  and T1.basetype='15')  ");
            sb.Append("               LEFT JOIN  rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')  ");
            sb.Append("               LEFT join ordr t8 on (t8.docentry=T5.docentry  )   ");
            if (TYPE == "1")
            {
                if (textBox1.Text == "")
                {
                    sb.Append("               where  substring(T0.CARDNAME,1,1) not like '%[A-Z]%' and isnull(T2.street,'') <> '' AND T0.U_IN_BSTY2 <> 2 ");
                    sb.Append("                  AND  Convert(varchar(8),T0.CREATEDATE,112) BETWEEN @AA AND @BB ");

                }
                if (textBox1.Text != "")
                {
                    sb.Append("               where  substring(T0.CARDNAME,1,1) not like '%[A-Z]%' and isnull(T2.street,'') <> '' AND U_IN_BSTY2 <> 2 ");
                    sb.Append(" AND  T0.DocEntry=@DocEntry ");
                }


            }
            if (TYPE == "2")
            {
                sb.Append(" WHERE  T0.DOCENTRY in ( " + cs + ")  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@AA", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox6.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
      

        private DataTable GetAddress1()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.DOCENTRY,T0.CARDNAME+' '+ISNULL(T2.street,'')+' '+ISNULL(REPLACE(T2.COUNTY,'不同PO請務必分開開發票',''),'') +' '  ");
            sb.Append(" +ISNULL(T2.block,'')+' '+ISNULL(T2.city,'') +ISNULL(T2.ZIPCODE,'')  ADDRESS FROM OINV T0    ");
            sb.Append(" LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PayToCode=T2.[Address]  and T2.adrestype='B')  ");
            sb.Append(" where  substring(T0.CARDNAME,1,1) not like '%[A-Z]%' and isnull(T2.street,'') <> '' AND T0.U_IN_BSTY2 <> 2  ");
                if (textBox1.Text == "")
                {
                    sb.Append("AND  Convert(varchar(8),T0.CREATEDATE,112) BETWEEN @AA AND @BB  ");

                }
                if (textBox1.Text != "")
                {
                    sb.Append(" AND  T0.DocEntry=@DocEntry ");
                }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@DocEntry", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@AA", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox6.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }

        private DataTable GetAddress1PER(string DOCENTRY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  DISTINCT T8.pickrmrk FROM inv1 T1");
            sb.Append(" left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
            sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
            sb.Append(" WHERE T1.DOCENTRY=@DOCENTRY AND ISNULL(T8.pickrmrk,'') <> ''");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private DataTable GetAddress2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT RTRIM(case isnull(ADD3,'') when '' then ADD1 ELSE ADD3 END) ADDRESS,ZIP FROM GB_ADD ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
           


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }


        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string szSavePath = GetExePath();
            string szSaveFile = GetExePath() +"\\PPLA_Example.Prn";

            textBox5.Text = GetMenu.Day();
            textBox6.Text = GetMenu.Day();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Data.DataTable TempDt = MakeTable();
            System.Data.DataTable dt = GetAddress1();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = TempDt.NewRow();
                string DOCENTRY = dt.Rows[i]["DOCENTRY"].ToString();
                dr["DOCENTRY"] = DOCENTRY;
                dr["ADDRESS"] = dt.Rows[i]["ADDRESS"].ToString();
                System.Data.DataTable G1 = GetAddress1PER(DOCENTRY);
                if (G1.Rows.Count > 0)
                {
                    dr["PERSON"] = G1.Rows[0][0].ToString();
                }
      
                TempDt.Rows.Add(dr);
            }

            dataGridView1.DataSource = TempDt;
        }

    }
}

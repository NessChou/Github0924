using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Reflection;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using System.Net;
using System.Linq.Expressions;

namespace ACME
{
    public partial class WH_main : ACME.fmBase1
    {
        string OutPutFileDRS = "";
        private SAPbobsCOM.Recordset oRecordSet;
        int QTYF = 0;
        public string PublicString;
        string DOCM = "";
        string JOBNO = "";

        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        System.Net.Mail.Attachment data = null;
        string df,ef;
        string fg = "";
        string bbs, hjj, DATE1, LOGINID = "";
        string DATE2 = GetMenu.Day();
        string MAILSUB = "";
     
        //
        private string Company_Man;
        private string Tel;
  

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


        const string sznop1 = "nop_front\r\n";
        const string sznop2 = "nop_middle\r\n";
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
        //
        public WH_main()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            wH_mainTableAdapter.Connection = MyConnection;

            wH_ItemTableAdapter.Connection = MyConnection;
            wH_Item2TableAdapter.Connection = MyConnection;
            wH_Item3TableAdapter.Connection = MyConnection;
            wH_Item4TableAdapter.Connection = MyConnection;
            wH_Item5TableAdapter.Connection = MyConnection;
            wH_LABTableAdapter.Connection = MyConnection;
            wH_FEETableAdapter.Connection = MyConnection;

        }
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            cardCodeTextBox.ReadOnly = true;
            cardNameTextBox.ReadOnly = true;
            boatNameTextBox.ReadOnly = true;
            boatCompanyTextBox.ReadOnly = true;
            textBox7.ReadOnly = false;
            button4.Enabled = true;
            button5.Enabled = true;
            button7.Enabled = true;
            button17.Enabled = true;
            button18.Enabled = true;
            btnSeqDelivery.Enabled = true;
            button8.Enabled = true;
            button34.Enabled = true;
            button13.Enabled = true;
            checkBox2.Enabled = true;
            button24.Enabled = true;
            checkBox3.Enabled = true;
            checkBox4.Enabled = true;
            checkBox5.Enabled = true;
            checkBox6.Enabled = true;

            receiveTypeComboBox.Enabled = true;

            forecastDayTextBox.ReadOnly =true;
            shipping_OBUTextBox.ReadOnly = true;
            contextMenuStrip3.Enabled = false;

            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
            button19.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;

            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;

            UPDATECHECK();
  
        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            cardCodeTextBox.ReadOnly = false;
            cardNameTextBox.ReadOnly = false;
            boatNameTextBox.ReadOnly = false;
            boatCompanyTextBox.ReadOnly = false;
            createNameTextBox.ReadOnly = false;
            modifyNameTextBox.ReadOnly = false;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            receiveTypeComboBox.SelectedIndex = -1;
            boardDeliverTextBox.Text = "";
        }
    
 

        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                wh.WH_main.RejectChanges();
                wh.WH_Item.RejectChanges();
                wh.WH_Item2.RejectChanges();
                wh.WH_Item3.RejectChanges();
                wh.WH_Item4.RejectChanges();
                wh.WH_FEE.RejectChanges();

            }
            catch
            { 
            }
            return true;

        }
        public override void EndEdit()
        {
            WW();
        }
        public override void SetInit()
        {
            MyBS = wH_mainBindingSource;
            MyTableName = "WH_main";
            MyIDFieldName = "ShippingCode";           
            UtilSimple.SetLookupBinding(boardCountNoComboBox, "boardCountNo", wH_mainBindingSource, "boardCountNo");
            UtilSimple.SetLookupBinding(receiveTypeComboBox, "receiveType", wH_mainBindingSource, "receiveType");


            MasterTable = wh.WH_main;
            DetailTables = new System.Data.DataTable[] { wh.WH_Item4 };
            DetailBindingSources = new BindingSource[] { wH_Item4BindingSource };

        }
        public override void AfterEdit()
        {


            string username = fmLogin.LoginID.ToString();

            if (globals.DBNAME == "宇豐")
            {
                System.Data.DataTable ff = GetOHEMAD(username);
                if (ff.Rows.Count > 0)
                {
                    modifyNameTextBox.Text = ff.Rows[0][0].ToString();

                }
                else
                {
                    modifyNameTextBox.Text = username;
                }
            }
            else
            {
                System.Data.DataTable ff = GetOHEM(username);
                if (ff.Rows.Count > 0)
                {
                    DataRow drw = ff.Rows[0];
                    string ss = drw["姓名"].ToString();
                    ss = ss.Replace("(", "");
                    ss = ss.Replace(")", "");
                    modifyNameTextBox.Text = ss;

                }
                else
                {
                    modifyNameTextBox.Text = username;
                }

            }
     
            forecastDayTextBox.ReadOnly = true;
            shippingCodeTextBox.ReadOnly = true;
            shipping_OBUTextBox.ReadOnly = true;
            contextMenuStrip3.Enabled = true;


    
        }
        public override void STOP()
        {
            if (forecastDayTextBox.Text == "")
            {
                MessageBox.Show("請輸入單據總類");
                this.SSTOPID = "1";
                forecastDayTextBox.Focus();
                return;
            }
            if (shipping_OBUTextBox.Text == "")
            {
                MessageBox.Show("請輸入倉庫");
                this.SSTOPID = "1";
                shipping_OBUTextBox.Focus();
                return;

            }

            if (!IsNumber(shipToDateTextBox.Text) && shipToDateTextBox.Text!="")
            {
                MessageBox.Show("進貨箱數請輸入數字");
                this.SSTOPID = "1";
                shipToDateTextBox.Focus();
                return;
            }
            if (!IsNumber(receiveDayTextBox.Text) && receiveDayTextBox.Text != "")
            {
                MessageBox.Show("出貨箱數請輸入數字");
                this.SSTOPID = "1";
                receiveDayTextBox.Focus();
                return;
            }

        }

        private void UPDATE()
        {
            if (wH_ItemDataGridView.Rows.Count > 1)
            {
                if (add4TextBox.Text == "")
                {
                    UpdateAPLC4();
                }

             
            }
            if (wH_Item2DataGridView.Rows.Count > 1)
            {
                if (add5TextBox.Text == "")
                {
                    UpdateAPLC5();
                }
            }
        }
        public override void AfterEndEdit()
        {

            UPDATE();

            try
            {


                StringBuilder sb = new StringBuilder();
                System.Data.DataTable dt = GetAUINV(shippingCodeTextBox.Text);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {

                        DataRow dd = dt.Rows[i];


                        sb.Append(dd["INV"].ToString() + "/");


                    }

                    sb.Remove(sb.Length - 1, 1);

                    sendGoodsTextBox.Text = sb.ToString();
                }

                StringBuilder sb2 = new StringBuilder();
                System.Data.DataTable dt2 = GetAUINV2(shippingCodeTextBox.Text);
                if (dt2.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                    {

                        DataRow dd = dt2.Rows[i];


                        sb2.Append(dd["INV"].ToString() + "/");


                    }

                    sb2.Remove(sb2.Length - 1, 1);

                    iNVOICENOTextBox.Text = sb2.ToString();
                }

                if (receiveMemoTextBox.Text == "")
                {
                    string FF4 = "";

                    string FF1 = "";
                    string FF2 = pQTYTextBox.Text;
                    string FF3 = kQTYTextBox.Text;

                    if (lIN20CheckBox.Checked)
                    {

                        FF4 = Environment.NewLine + "1*20''櫃:" + lINGATextBox.Text;
                    }

                    if (lIN40CheckBox.Checked)
                    {
                        FF4 = Environment.NewLine + "1*40''櫃:" + lINGATextBox.Text;
                    }
                    if (!String.IsNullOrEmpty(lINHTextBox.Text))
                    {
                        FF1 = lINHTextBox.Text + "=";
                    }
                    if (!String.IsNullOrEmpty(pQTYTextBox.Text))
                    {
                        FF2 = pQTYTextBox.Text + "PLTS";
                    }
                    if (!String.IsNullOrEmpty(kQTY2TextBox.Text))
                    {
                        FF3 = "=" + kQTY2TextBox.Text + "CTNS";
                    }

                    receiveMemoTextBox.Text = FF1 + FF2 + FF3 + FF4;
                }
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                {
                    System.Data.DataTable G5 = GetDI41();
                    StringBuilder sb5 = new StringBuilder();
                    if (G5.Rows.Count > 0)
                    {
                        for (int f = 0; f <= G5.Rows.Count - 1; f++)
                        {
                            DataRow dd = G5.Rows[f];
                            sb5.Append(dd["DOCENTRY"].ToString() + ",");
                        }

                        sb5.Remove(sb5.Length - 1, 1);
                        gPSPhoneTextBox.Text = sb5.ToString();
                        gPSPhoneTextBox1.Text = sb5.ToString();
                    }

                    if (quantityTextBox.Text == "")
                    {
                        if (forecastDayTextBox.Text == "銷售訂單")
                        {
                            System.Data.DataTable G1 = GetORDR();

                            if (G1.Rows.Count > 0)
                            {
                                string DOCENTRY = G1.Rows[0][0].ToString();
                                string LINENUM = G1.Rows[0][1].ToString();
                                System.Data.DataTable SHIPDATE = GetSHIPDATE(DOCENTRY, LINENUM);
                                if (SHIPDATE.Rows.Count > 0)
                                {
                                    quantityTextBox.Text = SHIPDATE.Rows[0][2].ToString();
                                }
                            }
                        }
                        if (forecastDayTextBox.Text == "庫存調撥-撥倉" || forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回")
                        {
                            System.Data.DataTable G1 = GetORDR();

                            if (G1.Rows.Count > 0)
                            {
                                string DOCENTRY = G1.Rows[0][0].ToString();
                                string LINENUM = G1.Rows[0][1].ToString();
                                System.Data.DataTable SHIPDATE = GetSHIPDATEWTR1(DOCENTRY, LINENUM);
                                if (SHIPDATE.Rows.Count > 0)
                                {
                                    quantityTextBox.Text = SHIPDATE.Rows[0][2].ToString();
                                }
                            }
                        }
                    }
                }
                if (quantityTextBox.Text.Length != 10 && quantityTextBox.Text != "")
                {
                    if (quantityTextBox.Text.Contains("/"))
                    {
                        DateTime date = Convert.ToDateTime(quantityTextBox.Text);
                        quantityTextBox.Text = string.Format("{0:yyyy\\/MM\\/dd}", date);
                    }
                    else
                    {
                        int year = Convert.ToInt32(quantityTextBox.Text.Substring(0, 4));
                        int month = Convert.ToInt32(quantityTextBox.Text.Substring(4, 2));
                        int day = Convert.ToInt32(quantityTextBox.Text.Substring(6, 2));
                        DateTime date = new DateTime(year,month,day);
                        quantityTextBox.Text = string.Format("{0:yyyy\\/MM\\/dd}", date);
                    }

                }

                wH_mainBindingSource.EndEdit();
                this.wH_mainTableAdapter.Update(wh.WH_main);
                wh.WH_main.AcceptChanges();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
    
        }
        public override void AfterAddNew()
        {
            WW();

            nTDollarsTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");
            contextMenuStrip3.Enabled = true;
        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "WH" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes  = NumberName + AutoNum + "X";

                if (globals.DBNAME == "達睿生")
                {
                    kyes = "DRS" + NumberName + AutoNum + "X";
                }
                if (globals.DBNAME == "CHOICE")
                {
                    kyes = NumberName + AutoNum + "X-CC";
                }
                if (globals.DBNAME == "INFINITE")
                {
                    kyes = NumberName + AutoNum + "X-IP";
                }
                if (globals.DBNAME == "TOP GARDEN")
                {
                    kyes = NumberName + AutoNum + "X-TG";
                }
                if (globals.DBNAME == "宇豐")
                {
                    kyes = NumberName + AutoNum + "X-AD";
                }
                if (globals.DBNAME == "禾中")
                {
                    kyes = NumberName + AutoNum + "X-GT";
                }
            }
            this.shippingCodeTextBox.Text = kyes;
            string username = fmLogin.LoginID.ToString();

            if (globals.DBNAME == "宇豐")
            {
                System.Data.DataTable ff = GetOHEMAD(username);
                if (ff.Rows.Count > 0)
                {
                    createNameTextBox.Text = ff.Rows[0][0].ToString();

                }
                else
                {
                    createNameTextBox.Text = username;
                }
            }
            else
            {
                System.Data.DataTable ff = GetOHEM(username);
                if (ff.Rows.Count > 0)
                {
                    DataRow drw = ff.Rows[0];
                    string ss = drw["姓名"].ToString();
                    ss = ss.Replace("(", "");
                    ss = ss.Replace(")", "");
                    createNameTextBox.Text = ss;
                }
                else
                {
                    createNameTextBox.Text = username;
                }
            }
    
        
            closeDayTextBox.Text = DateTime.Now.ToString("yyyyMMdd");

            buCardcodeCheckBox.Checked = false;
            soNoCheckBox.Checked = false;
     
            modifyDateCheckBox.Checked = false;
            s1CheckBox.Checked = false;
            s2CheckBox.Checked = false;
            s3CheckBox.Checked = false;
            s6CheckBox.Checked = false;
            s7CheckBox.Checked = false;
            this.wH_mainBindingSource.EndEdit();
            kyes = null;
        }
        public override void AfterCopy()
        {
            if (kyes == null)
            {

                string NumberName = "WH" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";

                if (globals.DBNAME == "達睿生")
                {
                    kyes = "DRS" + NumberName + AutoNum + "X";
                }
                if (globals.DBNAME == "CHOICE")
                {
                    kyes = NumberName + AutoNum + "X-CC";
                }
                if (globals.DBNAME == "INFINITE")
                {
                    kyes = NumberName + AutoNum + "X-IP";
                }
                if (globals.DBNAME == "TOP GARDEN")
                {
                    kyes = NumberName + AutoNum + "X-TG";
                }
                if (globals.DBNAME == "宇豐")
                {
                    kyes = NumberName + AutoNum + "X-AD";
                }
                if (globals.DBNAME == "禾中")
                {
                    kyes = NumberName + AutoNum + "X-GT";
                }
            }
        }
        public override void AfterCopy2()
        {

            tabControl1.SelectedIndex = 0;
            string username = fmLogin.LoginID.ToString();

            if (globals.DBNAME == "宇豐")
            {
                System.Data.DataTable ff = GetOHEMAD(username);
                if (ff.Rows.Count > 0)
                {
                    createNameTextBox.Text = ff.Rows[0][0].ToString();

                }
                else
                {
                    createNameTextBox.Text = username;
                }
            }
            else
            {
                System.Data.DataTable ff = GetOHEM(username);
                if (ff.Rows.Count > 0)
                {
                    DataRow drw = ff.Rows[0];
                    string ss = drw["姓名"].ToString();
                    ss = ss.Replace("(", "");
                    ss = ss.Replace(")", "");
                    createNameTextBox.Text = ss;
                }
                else
                {
                    createNameTextBox.Text = username;
                }
            }

            closeDayTextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            buCardcodeCheckBox.Checked = false;
            soNoCheckBox.Checked = false;
            s1CheckBox.Checked = false;
            s2CheckBox.Checked = false;
            s3CheckBox.Checked = false;

        }
        public override void FillData()
        {
            try
            {
                if (!String.IsNullOrEmpty(PublicString))
                {
                    MyID = PublicString.Trim();

                }

                wH_mainTableAdapter.Fill(wh.WH_main, MyID);
                wH_ItemTableAdapter.Fill(wh.WH_Item, MyID);
                wH_Item2TableAdapter.Fill(wh.WH_Item2, MyID);
                wH_Item3TableAdapter.Fill(wh.WH_Item3, MyID);
                wH_Item4TableAdapter.Fill(wh.WH_Item4, MyID);
                wH_Item5TableAdapter.Fill(wh.WH_Item5, MyID);
                wH_LABTableAdapter.Fill(wh.WH_LAB, MyID);
                wH_FEETableAdapter.Fill(wh.WH_FEE, MyID);

                ViewBatchPayment();

                SHIPNO();

                WHPACK();

                da();

                WHSUNNY();
                System.Data.DataTable G1 = GetFEE();
                if (G1.Rows.Count == 0)
                {
                    AddFEE();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void WHSUNNY()
        {

            System.Data.DataTable J1 = GETSUNNYS(shippingCodeTextBox.Text);
            if (J1.Rows.Count > 0)
            {
                string U_ACME_INV = J1.Rows[0]["U_ACME_INV"].ToString();
                string DOCENTRY = J1.Rows[0]["DOCENTRY"].ToString();

                UPODLN(U_ACME_INV, DOCENTRY);

            }
        }
            private void WHPACK()
        {
            string TYPE="";
            if (textBox7.Text == "")
            {
                TYPE = "A";
            }
            else
            {
                TYPE = "B";
            }
            System.Data.DataTable K5 = GetWHPACK2(shippingCodeTextBox.Text, textBox7.Text, TYPE);
            if (K5.Rows.Count > 0)
            {
                dataGridView2.DataSource = K5;
                System.Data.DataTable K6 = GetWHPACK3(shippingCodeTextBox.Text, textBox7.Text, TYPE);
                System.Data.DataTable K7 = GetWHPACK4(shippingCodeTextBox.Text, textBox7.Text, TYPE);
                if (K6.Rows.Count > 0)
                {
                    textBox1.Text = K7.Rows[0][0].ToString();
                    textBox2.Text = K6.Rows[0][1].ToString();
                    textBox4.Text = K6.Rows[0][2].ToString();
                    textBox5.Text = K6.Rows[0][3].ToString();
                    textBox6.Text = K6.Rows[0][4].ToString();

                }

            }
            else
            {
                dataGridView2.DataSource = GetWHPACK2("1234", textBox7.Text, "A");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
            }
        }
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {


                Validate();

                wH_Item4BindingSource.MoveFirst();

                for (int i = 1; i <= wH_Item4BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)wH_Item4BindingSource.Current;

                    row["SeqNo"] = i;



                    wH_Item4BindingSource.EndEdit();

                    wH_Item4BindingSource.MoveNext();
                }


                wH_ItemBindingSource.MoveFirst();
                for (int i = 1; i <= wH_ItemBindingSource.Count; i++)
                {
                    DataRowView row2 = (DataRowView)wH_ItemBindingSource.Current;

                    row2["SeqNo"] = i;



                    wH_ItemBindingSource.EndEdit();

                    wH_ItemBindingSource.MoveNext();
                }

                wH_Item2BindingSource.MoveFirst();
                
                for (int i = 1; i <= wH_Item2BindingSource.Count; i++)
                {
                    DataRowView row3 = (DataRowView)wH_Item2BindingSource.Current;

                    row3["SeqNo"] = i;



                    wH_Item2BindingSource.EndEdit();

                    wH_Item2BindingSource.MoveNext();
                }

                wH_Item3BindingSource.MoveFirst();
                for (int i = 1; i <= wH_Item3BindingSource.Count; i++)
                {
                    DataRowView row4 = (DataRowView)wH_Item3BindingSource.Current;

                    row4["SeqNo"] = i;

                    wH_Item3BindingSource.EndEdit();

                    wH_Item3BindingSource.MoveNext();
                }

                
                wH_LABBindingSource.MoveFirst();
                for (int i = 1; i <= wH_LABBindingSource.Count; i++)
                {
                    DataRowView row5 = (DataRowView)wH_LABBindingSource.Current;

                    row5["SeqNo"] = i;

                    wH_LABBindingSource.EndEdit();

                    wH_LABBindingSource.MoveNext();
                }

                wH_Item5BindingSource.MoveFirst();

                for (int i = 1; i <= wH_Item5BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)wH_Item5BindingSource.Current;

                    row["SeqNo"] = i;



                    wH_Item5BindingSource.EndEdit();

                    wH_Item5BindingSource.MoveNext();
                }




                wH_mainTableAdapter.Connection.Open();


                wH_mainBindingSource.EndEdit();
                wH_ItemBindingSource.EndEdit();
                wH_Item2BindingSource.EndEdit();
                wH_Item3BindingSource.EndEdit();
                wH_Item5BindingSource.EndEdit();
                wH_LABBindingSource.EndEdit();
                wH_FEEBindingSource.EndEdit();

                tx = wH_mainTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(wH_mainTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter1 = util.GetAdapter(wH_ItemTableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;
                
                SqlDataAdapter Adapter2 = util.GetAdapter(wH_Item2TableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;
             
                SqlDataAdapter Adapter3 = util.GetAdapter(wH_Item3TableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;
             
                SqlDataAdapter Adapter6 = util.GetAdapter(wH_LABTableAdapter);
                Adapter6.UpdateCommand.Transaction = tx;
                Adapter6.InsertCommand.Transaction = tx;
                Adapter6.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter7 = util.GetAdapter(wH_Item5TableAdapter);
                Adapter7.UpdateCommand.Transaction = tx;
                Adapter7.InsertCommand.Transaction = tx;
                Adapter7.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter8 = util.GetAdapter(wH_FEETableAdapter);
                Adapter8.UpdateCommand.Transaction = tx;
                Adapter8.InsertCommand.Transaction = tx;
                Adapter8.DeleteCommand.Transaction = tx;



                wH_mainTableAdapter.Update(wh.WH_main);
                wh.WH_main.AcceptChanges();

                wH_ItemTableAdapter.Update(wh.WH_Item);
                wh.WH_Item.AcceptChanges();

                wH_Item2TableAdapter.Update(wh.WH_Item2);
                wh.WH_Item2.AcceptChanges();

                wH_Item3TableAdapter.Update(wh.WH_Item3);
                wh.WH_Item3.AcceptChanges();

                wH_LABTableAdapter.Update(wh.WH_LAB);
                wh.WH_LAB.AcceptChanges();

                wH_Item5TableAdapter.Update(wh.WH_Item5);
                wh.WH_Item5.AcceptChanges();

                wH_FEETableAdapter.Update(wh.WH_FEE);
                wh.WH_FEE.AcceptChanges();



                //if (wh.WH_Item4.Rows.Count < 40)
                //{
                    wH_Item4BindingSource.EndEdit();

                    SqlDataAdapter Adapter4 = util.GetAdapter(wH_Item4TableAdapter);
                    Adapter4.UpdateCommand.Transaction = tx;
                    Adapter4.InsertCommand.Transaction = tx;
                    Adapter4.DeleteCommand.Transaction = tx;

                    wH_Item4TableAdapter.Update(wh.WH_Item4);
                    wh.WH_Item4.AcceptChanges();
            //    }


                tx.Commit();

                this.MyID = this.shippingCodeTextBox.Text;

                UpdateData = true;
            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }


                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;

            }
            finally
            {
                this.wH_mainTableAdapter.Connection.Close();

            }

            //if (wh.WH_Item4.Rows.Count >= 40)
            //{
            //    DELETETA();
            //    INSERTAA();
            //}

            return UpdateData;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" ||  globals.DBNAME == "測試區98" )
            {
                LookupValues = GetMenu.GetMenuList();
            }
            else if (globals.DBNAME == "CHOICE")
            {
              LookupValues = GetMenu.GetCHI();
            }
            else if (globals.DBNAME == "INFINITE")
            {
                LookupValues = GetMenu.GetCHI2();
            }
            else if (globals.DBNAME == "TOP GARDEN")
            {
                LookupValues = GetMenu.GetCHI4();
            }
            else if (globals.DBNAME == "宇豐")
            {
                LookupValues = GetMenu.GetCHI5();
            }
            else if (globals.DBNAME == "禾中")
            {
                LookupValues = GetMenu.GetCHI6();
            }
            if (LookupValues != null)
            {
                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }
        private void INSERTAA()
        {
            if (wh.WH_Item4.Rows.Count > 40)
            {
                for (int i = 0; i <= wH_Item4DataGridView.Rows.Count - 2; i++)
                {
                    DataGridViewRow row;

                    row = wH_Item4DataGridView.Rows[i];
                    string SeqNo4 = row.Cells["SeqNo4"].Value.ToString();
                    string Docentry = row.Cells["Docentry"].Value.ToString();
                    string linenum4 = row.Cells["linenum4"].Value.ToString();
                    string ShipDate4 = row.Cells["ShipDate4"].Value.ToString();
                    string ItemRemark = row.Cells["ItemRemark"].Value.ToString();
                    string ItemCode = row.Cells["ItemCode"].Value.ToString();
                    string Dscription = row.Cells["Dscription"].Value.ToString();
                    string PiNo = row.Cells["PiNo"].Value.ToString();
                    string Quantity4 = row.Cells["Quantity4"].Value.ToString();
                    string NowQty4 = row.Cells["NowQty4"].Value.ToString();
                    string Grade4 = row.Cells["Grade4"].Value.ToString();
                    string Ver4 = row.Cells["Ver4"].Value.ToString();
                    string INV4 = row.Cells["INV4"].Value.ToString();
                    string Invoice4 = row.Cells["Invoice4"].Value.ToString();
                    string Remark4 = row.Cells["Remark4"].Value.ToString();
                    string FrgnName = row.Cells["FrgnName"].Value.ToString();
                    string WHName = row.Cells["WHName"].Value.ToString();
                    string U_PAY = row.Cells["U_PAY"].Value.ToString();
                    string U_SHIPDAY = row.Cells["U_SHIPDAY"].Value.ToString();
                    string U_SHIPSTATUS = row.Cells["U_SHIPSTATUS"].Value.ToString();
                    string U_MARK = row.Cells["U_MARK"].Value.ToString();
                    string U_MEMO = row.Cells["U_MEMO"].Value.ToString();
                    string PO = row.Cells["PO"].Value.ToString();
                    string LOCATION4 = row.Cells["LOCATION4"].Value.ToString();
                    string CardCode = row.Cells["CardCode2"].Value.ToString();
                    string TREETYPE = row.Cells["TREETYPE"].Value.ToString();
                    INSERTA(shippingCodeTextBox.Text, SeqNo4, Docentry, linenum4, ItemRemark, ItemCode, Dscription, Quantity4, Remark4, INV4, PiNo, NowQty4, Ver4, Grade4, Invoice4, FrgnName, WHName, ShipDate4, CardCode, U_PAY, U_SHIPDAY, U_SHIPSTATUS, U_MARK, U_MEMO, PO, LOCATION4, TREETYPE);
                }
            }
        }
        private void wH_ItemDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = wH_ItemDataGridView.Rows.Count;
            e.Row.Cells["SeqNo"].Value = iRecs.ToString();

            e.Row.Cells["Quantity"].Value = 0;

        }



        private void WH_main_Load(object sender, EventArgs e)
        {
            UpdateAPLC6();
            DELETEFILE2();
            
            string connect = MyConnection.ConnectionString.ToString();
            int g = connect.ToUpper().IndexOf("TEST");
            if (g != -1)
            {
                this.Text = "倉管系統測試區";
            }
            LOGINID = fmLogin.LoginID.ToString() + "@acmepoint.com";
            WW();


            UtilSimple.SetLookupBinding(comboBox3, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, GetMenu.Month(), "DataValue", "DataValue");

            comboBox3.Text = DateTime.Now.ToString("yyyy");
            comboBox4.Text = Convert.ToString(Convert.ToInt16(DateTime.Now.ToString("MM")));

            if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
            {
                button16.Visible = false;
            }

            if (globals.DBNAME != "禾中")
            {
                textBox8.Visible = false;
                button27.Visible = false;
                label17.Visible = false;
                label18.Visible = false;
            }

            if (globals.DBNAME == "宇豐")
            {
                modifyDateCheckBox.Visible = true;
     
            }
        }


        private void buCardcodeCheckBox_Click(object sender, EventArgs e)
        {
            if (buCardcodeCheckBox.Checked)
            {

                buCardnameTextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            }
        }
 
        private void button9_Click(object sender, EventArgs e)
        {



            if (cardCodeTextBox.Text == "" && forecastDayTextBox.Text != "發貨單" && forecastDayTextBox.Text != "庫存調撥-撥倉" && forecastDayTextBox.Text != "收貨單")
            {
                MessageBox.Show("請輸入客戶編號");
                return;
            }

            else if (String.IsNullOrEmpty(forecastDayTextBox.Text))
            {
                MessageBox.Show("請選擇單據");
                return;
            }
            else if (String.IsNullOrEmpty(shipping_OBUTextBox.Text.ToString()))
            {
                MessageBox.Show("請選擇倉庫");
                return;
            }
            
                string check1="";
                if (checkBox1.Checked)
                {
                    check1 = "a";
                }
                else
                {
                    check1 = "b";
                }
                string aa = cardCodeTextBox.Text;
                object[] LookupValues = null;
                System.Data.DataTable h1 = null;
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98" )
                {
                     LookupValues = GetMenu.Getowtr1(aa, forecastDayTextBox.Text, check1);
                     h1 = GetOWHS(shipping_OBUTextBox.Text);
                }
                else if (globals.DBNAME == "CHOICE")
                {
                    LookupValues = GetMenu.GetowtrCHO1(aa, forecastDayTextBox.Text, check1);
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);
                }
                else if ( globals.DBNAME == "INFINITE")
                {
                    LookupValues = GetMenu.GetowtrCHO2(aa, forecastDayTextBox.Text, check1);
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);
                }
                else if (globals.DBNAME == "TOP GARDEN")
                {
                    LookupValues = GetMenu.GetowtrCHO4(aa, forecastDayTextBox.Text, check1);
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);

                    if (h1.Rows.Count == 0)
                    {

                        h1 = GetOWHSCHI2(shipping_OBUTextBox.Text);
                    }
                }
                else if (globals.DBNAME == "宇豐")
                {
                    LookupValues = GetMenu.GetowtrCHO5(aa, forecastDayTextBox.Text, check1);
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);
                }
                else if (globals.DBNAME == "禾中")
                {
                    LookupValues = GetMenu.GetowtrCHO6(aa, forecastDayTextBox.Text, check1);
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);

                    if (h1.Rows.Count == 0)
                    {

                        h1 = GetOWHSCHI2(shipping_OBUTextBox.Text);
                    }
                }
                if (LookupValues != null)
                {
                    string RR = Convert.ToString(LookupValues[0]);
                    string pino = RR;
                        pINOTextBox.Text = Convert.ToString(LookupValues[0]);
                        try
                        {


                            string dd = Convert.ToString(h1.Rows[0][0]);
                            System.Data.DataTable dt1 = null;
                            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                            {
                                dt1 = GetOrderData(RR, dd);
                            }
                            else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                            {
                                dt1 = GetOrderDataCHI(RR, dd, check1);
                            }

                            System.Data.DataTable dt2 = null;

                            dt2 = wh.WH_Item4;

                            int M1 = 0;
                            string MCODE = "";
                            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                            {
                                DataRow drw3 = dt1.Rows[0];
                            

                                    int g = drw3["工廠地址"].ToString().IndexOf("司");


                                    if (g == -1)
                                    {
                                        shipmentTextBox.Text = drw3["工廠地址"].ToString();

                                    }
                                    else
                                    {

                                        shipmentTextBox.Text = drw3["工廠地址"].ToString().Substring(g + 1).Trim();

                                    }

                                
                                arriveDayTextBox.Text = drw3["連絡人"].ToString();
                                cFSTextBox.Text = drw3["電話號碼"].ToString();
                                string g1 = drw3["業務"].ToString(); ;

                                buCntctPrsnTextBox.Text = drw3["業務"].ToString();

                         

                                DataRow drw = dt1.Rows[i];
                                DataRow drw2 = dt2.NewRow();
                                string D1 = drw["品名規格1"].ToString();

                                if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐")
                                {
                                    textBox3 .Text = drw["備註"].ToString();
                       
                                }
                                string 產品編號= drw["產品編號"].ToString();
                        
                                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                                drw2["Docentry"] = pINOTextBox.Text;
                                string 品名規格 = drw["品名規格"].ToString();
                              
                                drw2["Dscription"] = 品名規格;
                         
                                drw2["itemcode"] = 產品編號;
                                drw2["ItemRemark"] = forecastDayTextBox.Text;
                                drw2["WHName"] = shipping_OBUTextBox.Text.ToString();
                                decimal SS = Convert.ToDecimal(drw["數量"]);
                                string GH = Convert.ToDouble(SS).ToString();
                                drw2["Quantity"] = GH;
                                drw2["linenum"] = drw["欄號"];

                              
                                System.Data.DataTable QTY1 = GetQTYF(shipping_OBUTextBox.Text.ToString(), 產品編號);
                                if (QTY1.Rows.Count > 0)
                                {
                                    QTYF = Convert.ToInt32(QTY1.Rows[0][0]);
                                }
                                drw2["NowQty"] = Convert.ToInt32(drw["現有數量"]) - QTYF;
                                drw2["Ver"] = drw["版本"];
                                if (產品編號.Length > 2)
                                {
                                    string gg = 產品編號.Substring(0, 3).ToString().ToUpper();
                                    if (gg == "TAP")
                                    {

                                        int G1 = 品名規格.IndexOf(".");
                                        if (G1 != -1)
                                        {
                                            drw2["Ver"] = "V." + 品名規格.Substring(G1 + 1, 1);
                                        }
                                    }
                                }
                                drw2["Grade"] = drw["等級"];
                                string T1 = drw["單位"].ToString(); 
                                drw2["cardcode"] = drw["單位"];
                                drw2["ShipDate"] = drw["排程日期"];
                                if (forecastDayTextBox.Text == "採購單")
                                {


                                    System.Data.DataTable T1S = GetARRIVE(pINOTextBox.Text, drw["欄號"].ToString());
                                    if (T1S.Rows.Count > 0)
                                    {
                                        drw2["ShipDate"] = T1S.Rows[0][0].ToString();

                                        quantityTextBox.Text = T1S.Rows[0][1].ToString();

                                    }
                                }

                                drw2["U_PAY"] = drw["付款"];
                                drw2["U_SHIPDAY"] = drw["押出貨日"];
                                drw2["U_SHIPSTATUS"] = drw["貨況"];
                                drw2["U_MARK"] = drw["特殊嘜頭"];
                                drw2["U_MEMO"] = drw["注意事項"];
                                drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                                drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                                drw2["PO"] = drw["PO"];
                                drw2["LOCATION"] = drw["產地"];
                                string TREETYPE = drw["TREETYPE"].ToString();
                                drw2["TREETYPE"] = TREETYPE;
                                try
                                {
                                    hjj = "";
                     
                                    if (TREETYPE == "S")
                                    {
                                        hjj = "母料號";
                                        MCODE = 產品編號;
                                        M1 = 0;
                                    }
                                    else if (TREETYPE == "I")
                                    {

                                        M1++;
                                        hjj = MCODE + "-子料號-" + M1.ToString();
                                    }
                                    else
                                    {
                                        hjj = drw["PARTNO"].ToString();
                                    }

                                    drw2["pino"] = hjj;
                                }
                                catch
                                {

                                }
                                if (globals.DBNAME == "達睿生")
                                {
                                    if (forecastDayTextBox.Text == "採購單")
                                    {
                                        drw2["invoice"] = drw["INVOICE"].ToString();
                                    }

                                    drw2["FrgnName"] = drw["品名規格"];
                                }
                                else
                                {
                                    drw2["FrgnName"] = drw["品名規格1"];
                                    if (cardCodeTextBox.Text == "0017-00")
                                    {
                                        drw2["FrgnName"] = drw["品名規格1"] + "-" + drw["等級"];

                                    }
                                }

                                System.Data.DataTable L1 = GetCHIITEM(產品編號);
                                if (L1.Rows.Count > 0)
                                {
                                    if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐")
                                    {
                                        drw2["pino"] = L1.Rows[0]["PARTNO"].ToString();
                                 
                                    }
                                }
                                dt2.Rows.Add(drw2);

                            }

                            wH_Item4BindingSource.MoveFirst();

                            for (int i = 1; i <= wH_Item4BindingSource.Count; i++)
                            {
                                DataRowView row = (DataRowView)wH_Item4BindingSource.Current;

                                row["SeqNo"] = i;



                                wH_Item4BindingSource.EndEdit();

                                wH_Item4BindingSource.MoveNext();
                            }

                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    

                }

            wH_mainBindingSource.EndEdit();
            //wH_mainTableAdapter.Update(wh.WH_main);
            //wh.WH_main.AcceptChanges();

            wH_Item4BindingSource.EndEdit();
            //wH_Item4TableAdapter.Update(wh.WH_Item4);
            //wh.WH_Item4.AcceptChanges();

            button6_Click(sender, e);
        }
        private System.Data.DataTable   GetOrderData(string Doc_no, string whscode)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            switch (forecastDayTextBox.Text)
            {

                case "AR發票":
                    sb.Append(" SELECT ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),D.U_ACME_SHIPDAY,111) as 交貨日期,d.linenum as 欄號,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO, ");
                        sb.Append(" T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,d.U_CUSTITEMCODE,d.U_CUSTDOCENTRY  ");
                        sb.Append(" ,T1.street+ISNULL(T1.COUNTY,'') 工廠地址,os.slpname as 業務,os.MEMO as 流程, ");
                        sb.Append(" SALUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱, ");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO, ");
                        sb.Append("  '' as 發票方式, Convert(varchar(10),m.u_in_bsdat,111) as 發票日期,u_in_bsinv as 發票號碼,");
                        sb.Append(" 發票聯式 = case M.u_in_bsty1");
                        sb.Append(" when '0' then '三聯式發票'  when '1' then '三聯式收銀機發票' ");
                        sb.Append(" when '2' then '二聯式發票' when '3' then '二聯式收銀機發票'  ");
                        sb.Append(" when '4' then '電子計算機發票' when '5' then '免用統一發票' end, ");
                        sb.Append(" ISNULL(m.numatcard,'')+ISNULL(m.U_ACME_PAYGUI,'')+cast(isnull(m.u_acme_memo,'') as nvarchar) 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM oinv m");
                        sb.Append(" left join inv1 d on m.docentry=d.docentry");
                        sb.Append(" LEFT JOIN  CRD1 T1 ON (M.CARDCODE=T1.CARDCODE AND M.shiptocode=T1.ADDRESS and T1.adrestype='S')  ");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;
                    case "銷售訂單":

                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),D.U_ACME_SHIPDAY,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO, ");
                        sb.Append(" T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,d.U_CUSTITEMCODE,d.U_CUSTDOCENTRY   ");
                        sb.Append(" ,T1.street+ISNULL(T1.COUNTY,'')  工廠地址,os.slpname as 業務,os.MEMO as 流程, ");
                        sb.Append(" SALUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,ROHS='ROHS',AU='AUS', ");
                        sb.Append(" d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱, ");
                        sb.Append(" rtrim(ISNULL(m.numatcard,'')+ISNULL(m.U_ACME_PAYGUI,'')+cast(isnull(m.u_acme_memo,'') as nvarchar(1000))) 備註,oi.usertext 主要描述,oi.U_LOCATION 產地  FROM ordr m ");
                        sb.Append(" left join rdr1 d on m.docentry=d.docentry ");
                        sb.Append(" LEFT JOIN  CRD1 T1 ON (M.CARDCODE=T1.CARDCODE AND M.shiptocode=T1.ADDRESS and T1.adrestype='S')   ");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode ");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode ");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")  ");
                        if (checkBox1.Checked==false)
                        {
                            sb.Append("    and d.linestatus='O' ");
                        }
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;
                case "調撥單":
                case "庫存調撥-借出":
                    sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO, ");
                    sb.Append("  T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,m.cardname 客戶名稱 ,m.cardcode 客戶編號 ,m.cardname U_CUSTITEMCODE,m.cardcode U_CUSTDOCENTRY    ");
                    sb.Append(" ,m.ADDRESS 工廠地址,os.slpname as 業務,os.MEMO as 流程, ");
                    sb.Append(" INVNTRYUOM 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量, ");
                    sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION, OI.U_PARTNO PARTNO, ");
                    sb.Append(" case isnull(m.u_acme_serial,'') when '' then m.comments else  m.u_acme_serial end 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM owtr m ");
                    sb.Append(" left join wtr1 d on m.docentry=d.docentry ");
                    sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode ");
                    sb.Append(" LEFT JOIN  CRD1 T1 ON (oc.CARDCODE=T1.CARDCODE AND oc.shiptodef=T1.ADDRESS  and T1.adrestype='S')   ");
                    sb.Append(" left join oslp os on os.slpcode=m.slpcode ");
                    sb.Append(" left join oitm oi on oi.itemcode=d.itemcode ");
                    sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                    sb.Append(" WHERE m.u_acme_kind='1' ");
                    sb.Append(" and m.DOCENTRY in (" + Doc_no + ")");
                    sb.Append(" order by m.DOCENTRY,d.visorder ");
                    break;
                case "庫存調撥-撥倉":


                    sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                    sb.Append(" T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,m.cardname 客戶名稱 ,m.cardcode 客戶編號 ,m.cardname U_CUSTITEMCODE,m.cardcode U_CUSTDOCENTRY ");
                    sb.Append(" ,T1.street+ISNULL(T1.COUNTY,'') 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                    sb.Append(" INVNTRYUOM 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                    sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                    sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM owtr m");
                    sb.Append(" left join wtr1 d on m.docentry=d.docentry");
                    sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                    sb.Append(" LEFT JOIN  CRD1 T1 ON (oc.CARDCODE=T1.CARDCODE AND oc.shiptodef=T1.ADDRESS  and T1.adrestype='S')  ");
                    sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                    sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                    sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                    sb.Append(" WHERE  m.u_acme_kind='3'");
                    sb.Append(" and m.DOCENTRY in (" + Doc_no + ")");
                    sb.Append(" order by m.DOCENTRY,d.visorder ");

                    break;

                case "庫存調撥-借出還回":
                    sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                    sb.Append(" T1.zipcode+ISNULL(T1.U_USERNAME,'') as 連絡人,T1.block as 電話號碼,d.TREETYPE,m.cardname 客戶名稱 ,m.cardcode 客戶編號 ,m.cardname U_CUSTITEMCODE,m.cardcode U_CUSTDOCENTRY ");
                    sb.Append(" ,T1.street+ISNULL(T1.COUNTY,'') 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                    sb.Append(" INVNTRYUOM 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                    sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                    sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM owtr m");
                    sb.Append(" left join wtr1 d on m.docentry=d.docentry");
                    sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                    sb.Append(" LEFT JOIN  CRD1 T1 ON (oc.CARDCODE=T1.CARDCODE AND oc.shiptodef=T1.ADDRESS  and T1.adrestype='S')  ");
                    sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                    sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                    sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                    sb.Append(" WHERE m.u_acme_kind='2' ");
                    sb.Append(" and m.DOCENTRY in (" + Doc_no + ")");
                    sb.Append(" order by m.DOCENTRY,d.visorder ");

                    break;


                case "發貨單":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as  交貨日期,m.cardname as 客戶編號 ,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                        sb.Append(" o.[name] as 連絡人,'' as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,'' as 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" INVNTRYUOM 單位,'' 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM oige m");
                        sb.Append(" left join ige1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;
                  


                    case "採購報價":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                        sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" BUYUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地,d.u_acme_inv INVOICE FROM OPQT m");
                        sb.Append(" left join PQT1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append("  left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")  ");
                        if (checkBox1.Checked == false)
                        {
                            sb.Append("    and d.linestatus='O' ");
                        }
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;

                case "採購單":
                    sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                    sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                    sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                    sb.Append(" BUYUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                    sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                    sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地,d.u_acme_inv INVOICE FROM opor m");
                    sb.Append(" left join POR1 d on m.docentry=d.docentry");
                    sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                    sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                    sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                    sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                    sb.Append("  left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                    sb.Append(" where m.DOCENTRY in (" + Doc_no + ")  ");
                    if (checkBox1.Checked == false)
                    {
                        sb.Append("    and d.linestatus='O' ");
                    }
                    sb.Append(" order by m.DOCENTRY,d.visorder ");
                    break;

                case "收貨採購單":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                        sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" BUYUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,m.U_LOCATION 產地,d.u_acme_inv INVOICE FROM opdn m");
                        sb.Append(" left join PDN1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;

                    case "採購退貨":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                        sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" BUYUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,m.U_LOCATION 產地 FROM ORPD m");
                        sb.Append(" left join RPD1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;


                    case "AP貸項":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                        sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" BUYUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,m.U_LOCATION 產地 FROM ORPC m");
                        sb.Append(" left join RPC1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;

                    case "AR貸項通知單":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO,");
                        sb.Append(" o.[name] as 連絡人,oc.phone1 as 電話號碼,d.TREETYPE,d.U_CUSTITEMCODE,d.U_CUSTDOCENTRY");
                        sb.Append(" ,m.address 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" SALUNITMSR 單位,Convert(varchar(10),d.u_acme_work,111) 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append("  OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM orin m");
                        sb.Append(" left join rin1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join ocrd oc on oc.cardcode=m.cardcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;

                    case "收貨單":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as  交貨日期,m.cardname as 客戶編號 ,m.cardname 客戶名稱,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                        sb.Append(" o.[name] as 連絡人,'' as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,'' as 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" INVNTRYUOM 單位,'' 排程日期, d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM oign m");
                        sb.Append(" left join ign1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append("  left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;

                    case "生產訂單":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,'' as  交貨日期 ,'' as 付款,'' as 押出貨日,'' as 貨況,'' as PO,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                        sb.Append(" o.[name] as 連絡人,'' as 電話號碼,'' TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,'' as 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" INVNTRYUOM 單位,'' 排程日期,d.itemcode as 產品編號,oi.itemname as 品名規格,oi.frgnname as 品名規格1,d.plannedqty as 數量,m.cardcode 客戶編號,t1.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM OWOR m");
                        sb.Append(" left join WOR1 d on m.docentry=d.docentry");
                        sb.Append(" left join ORDR T0 on m.ORIGINNUM=T0.docentry");
                        sb.Append(" left join OCRD T1 on m.CARDCODE=T1.CARDCODE");
                        sb.Append(" left join ocpr o on o.cntctcode=T0.cntctcode");
                        sb.Append(" left join oslp os on os.slpcode=T0.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;


                    case "生產發貨":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as  交貨日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                        sb.Append(" o.[name] as 連絡人,'' as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.u_po_add2 as 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" INVNTRYUOM 單位,'' 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM oige m");
                        sb.Append(" left join ige1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;



                    case "生產收貨":
                        sb.Append(" SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as  交貨日期 ,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                        sb.Append(" o.[name] as 連絡人,'' as 電話號碼,d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY");
                        sb.Append(" ,m.u_po_add2 as 工廠地址,os.slpname as 業務,os.MEMO as 流程,");
                        sb.Append(" INVNTRYUOM 單位,'' 排程日期,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱,");
                        sb.Append(" OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO,");
                        sb.Append(" m.comments 備註,oi.usertext 主要描述,oi.U_LOCATION 產地 FROM OIGN m");
                        sb.Append(" left join IGN1 d on m.docentry=d.docentry");
                        sb.Append(" left join ocpr o on o.cntctcode=m.cntctcode");
                        sb.Append(" left join oslp os on os.slpcode=m.slpcode");
                        sb.Append(" left join oitm oi on oi.itemcode=d.itemcode");
                        sb.Append(" left join oitw ow on oi.itemcode=ow.itemcode and ow.whscode=@whscode ");
                        sb.Append(" where m.DOCENTRY in (" + Doc_no + ")");
                        sb.Append(" order by m.DOCENTRY,d.visorder ");
                        break;
            }



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@whscode", whscode));
          

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "new01");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderDataCHI(string Doc_no, string whscode,string STATUS)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            switch (forecastDayTextBox.Text)
            {


                case "銷售單":


sb.Append("                SELECT T0.BillNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T2.LinkMan 連絡人,T2.Telephone1 電話號碼,T0.CustAddress  工廠地址,T5.PersonName 業務, ");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
sb.Append("                                               END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,T2.ID 客戶編號,T2.FULLNAME 客戶名稱, ");
sb.Append("                                              'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY FROM OrdBillMain T0 ");
sb.Append("                      Inner Join OrdBillSub T1 On T0.Flag=T1.Flag And T0.BillNO=T1.BillNO  ");
sb.Append("                      Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID AND T2.Flag =1) ");
sb.Append("                      Inner Join comCustDesc T3 ON (T2.ID=T3.ID AND T3.Flag =1) ");
sb.Append("                      Left Join comPerson T5 ON (T0.SalesMan=T5.PersonID) ");
sb.Append("                      Left Join comProduct T6 ON (T1.ProdID =T6.ProdID) ");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and  T7.WareID=@whscode )");
                    sb.Append("  where t0.Flag =2 and T0.BillNO = @Doc_no  ");
                    if (STATUS == "b")
                    {
                        sb.Append(" AND T1.QtyRemain > 0 ");

                    }
                    break;

                case "採購單":


                    sb.Append(" SELECT T0.BillNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T2.LinkMan 連絡人,T2.Telephone1 電話號碼,T0.CustAddress  工廠地址,T5.PersonName 業務,");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
                    sb.Append("                          END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,T2.ID 客戶編號,T2.FULLNAME 客戶名稱,");
                    sb.Append("                         'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM OrdBillMain T0");
                    sb.Append(" Inner Join OrdBillSub T1 On T0.Flag=T1.Flag And T0.BillNO=T1.BillNO ");
                    sb.Append(" Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID AND T2.Flag =2)");
                    sb.Append(" Inner Join comCustDesc T3 ON (T2.ID=T3.ID AND T3.Flag =2)");
                    sb.Append(" Left Join comPerson T5 ON (T0.SalesMan=T5.PersonID)");
                    sb.Append(" Left Join comProduct T6 ON (T1.ProdID =T6.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("  where t0.Flag =4 and T0.BillNO = @Doc_no  ");
                    break;

                case "銷貨單":


                    sb.Append(" SELECT T8.ID 客戶編號,T8.FULLNAME 客戶名稱,T0.FundBillNo 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T0.ContactPerson 連絡人,T0.ContactPhone 電話號碼,T0.CustAddress  工廠地址,T2.PersonName 業務");
 sb.Append("              ,      case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
                    sb.Append("                                               END 等級,T3.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T3.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T3.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM comBillAccounts T0");
                    sb.Append(" LEFT JOIN comProdRec T1 ON (T0.FundBillNo=T1.BillNO AND T0.Flag=T1.Flag)");
                    sb.Append(" LEFT JOIN comPerson T2 ON (T0.SalesMan=T2.PersonID)");
                    sb.Append("           Left Join comProduct T3 ON (T1.ProdID =T3.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("                  Inner Join comCustomer T8 ON (T0.CustID =T8.ID)");
                    sb.Append("  WHERE T0.Flag=500 AND T0.FundBillNo = @Doc_no ");
                    if (STATUS == "b")
                    {
                        sb.Append(" AND T1.QtyRemain > 0 ");

                    }
                    break;

                case "銷退單":


                    sb.Append(" SELECT T8.ID 客戶編號,T8.FULLNAME 客戶名稱,T0.FundBillNo 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T0.ContactPerson 連絡人,T0.ContactPhone 電話號碼,T0.CustAddress  工廠地址,T2.PersonName 業務");
                    sb.Append("      ,case substring(T1.ProdID,11,1) ");
                    sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
                    sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
                    sb.Append("                                               when '4' then 'U' when '5' then 'NN'");
                    sb.Append("                                               END 等級,T3.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T3.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T3.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM comBillAccounts T0");
                    sb.Append(" LEFT JOIN comProdRec T1 ON (T0.FundBillNo=T1.BillNO AND T0.Flag=T1.Flag)");
                    sb.Append(" LEFT JOIN comPerson T2 ON (T0.SalesMan=T2.PersonID)");
                    sb.Append("           Left Join comProduct T3 ON (T1.ProdID =T3.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("                  Inner Join comCustomer T8 ON (T0.CustID =T8.ID)");
                    sb.Append("  WHERE T0.Flag=600 AND T0.FundBillNo = @Doc_no ");
                    if (STATUS == "b")
                    {
                        sb.Append(" AND T1.QtyRemain > 0 ");

                    }
                    break;


                case "調整單":

sb.Append(" SELECT '' 客戶編號,'' 客戶名稱,T0.AdjustNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,'' 連絡人,'' 電話號碼,''  工廠地址,'' 業務, ");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
sb.Append(" END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO, ");
sb.Append(" 'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM StkAdjustMain T0 ");
sb.Append(" LEFT Join comProdRec T1 On T0.Flag=T1.Flag And T0.AdjustNO =T1.BillNO  ");
sb.Append(" Left Join comProduct T6 ON (T1.ProdID =T6.ProdID) ");
sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID='A18' ) ");
sb.Append(" where T0.Flag =300  AND   T0.AdjustNO    = @Doc_no  ");

                    break;

                case "進貨單":


                    sb.Append(" SELECT T8.ID 客戶編號,T8.FULLNAME 客戶名稱,T0.FundBillNo 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T0.ContactPerson 連絡人,T0.ContactPhone 電話號碼,T0.CustAddress  工廠地址,T2.PersonName 業務");
               sb.Append("                  ,  case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
                    sb.Append("                                               END 等級,T3.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T3.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T3.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM comBillAccounts T0");
                    sb.Append(" LEFT JOIN comProdRec T1 ON (T0.FundBillNo=T1.BillNO AND T0.Flag=T1.Flag)");
                    sb.Append(" LEFT JOIN comPerson T2 ON (T0.SalesMan=T2.PersonID)");
                    sb.Append("           Left Join comProduct T3 ON (T1.ProdID =T3.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("                  Inner Join comCustomer T8 ON (T0.CustID =T8.ID)");
                    sb.Append("  WHERE T0.Flag=100 AND T0.FundBillNo = @Doc_no ");
                    break;

                case "調撥單":
                    sb.Append("              SELECT '' 客戶編號,'' 客戶名稱,T0.MoveNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,'' 連絡人,'' 電話號碼,''  工廠地址,'' 業務,");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
                    sb.Append("                                               END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                    sb.Append("                                              'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM StkMoveMain T0");
                    sb.Append("                      LEFT Join comProdRec T1 On T0.Flag=T1.Flag And T0.MoveNO=T1.BillNO ");
                    sb.Append("                      Left Join comProduct T6 ON (T1.ProdID =T6.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("                       where T0.Flag =400 AND   T0.MoveNO   = @Doc_no ");
                    break;

                case "借出單":
                    sb.Append("              SELECT T0.BorrowNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T2.LinkMan 連絡人,T2.Telephone1 電話號碼,T0.CustAddress  工廠地址,T5.PersonName 業務,");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
                    sb.Append("                                               END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO,");
                    sb.Append("                                              'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述,T2.ID 客戶編號,T2.FULLNAME 客戶名稱, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM StkBorrowMain T0");
                    sb.Append("                      Inner Join stkBorrowSub T1 On T0.Flag=T1.Flag And T0.BorrowNO=T1.BorrowNO ");
                    sb.Append("                      Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID )");
                    sb.Append("                      Left Join comPerson T5 ON (T0.SalesMan=T5.PersonID)");
                    sb.Append("                      Left Join comProduct T6 ON (T1.ProdID =T6.ProdID)");
                    sb.Append(" Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode )");
                    sb.Append("                       where t0.Flag =10 And T0.BorrowNO = @Doc_no ");
                    break;

                case "還回單":
sb.Append("                                  SELECT T0.ReturnNO 單號,ROWNO 欄號,ISNULL(CAST(T7.QUANTITY-T7.LENDQUAN AS INT),0) 現有數量,T2.LinkMan 連絡人,T2.Telephone1 電話號碼,T0.CustAddress  工廠地址,T5.PersonName 業務, ");
sb.Append("                    case WHEN substring(T1.ProdID,1,1) IN  ('K','T','U','C','A') THEN 'N/A' ELSE       case substring(T1.ProdID,11,1)  ");
sb.Append("                                               when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
sb.Append("                                               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
sb.Append("                                               when '4' then 'U' when '5' then 'NN'  END");
sb.Append("                                                                     END 等級,T6.Unit 單位,'' 排程日期,'' as 付款,'' as 押出貨日,'' as 貨況,'' as 特殊嘜頭,'' as 注意事項,'' as PO, ");
sb.Append("                                                                    'V.'+ substring(T1.ProdID,12,1) 版本,'' 交貨日期,T1.ProdID 產品編號,T1.ProdName 品名規格,T6.InvoProdName 品名規格1,T1.Quantity 數量,T0.REMARK 備註,T6.ProdDesc 主要描述,T2.ID 客戶編號,T2.FULLNAME 客戶名稱, CASE SUBSTRING(T1.PRODID,LEN(T1.PRODID),1) WHEN '1' THEN 'Taiwan' WHEN '2' THEN 'China' END 產地,'' TREETYPE,'' PARTNO,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY  FROM stkReturnMain T0 ");
sb.Append("                                            Inner Join stkReturnSub T1 On T0.Flag=T1.Flag And T0.ReturnNO=T1.ReturnNO  ");
sb.Append("                                            Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID ) ");
sb.Append("                                            Left Join comPerson T5 ON (T0.SalesMan=T5.PersonID) ");
sb.Append("                                            Left Join comProduct T6 ON (T1.ProdID =T6.ProdID) ");
sb.Append("                       Left Join comWareAmount T7 ON (T1.ProdID =T7.ProdID and T7.WareID=@whscode ) ");
sb.Append("                                             where t0.Flag =12 And T0.ReturnNO = @Doc_no  ");
                    break;
            }



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Doc_no", Doc_no));
            command.Parameters.Add(new SqlParameter("@whscode", whscode));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "new01");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button5_Click(object sender, EventArgs e)
        {
            
            DELETEFILE2();
            
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            if (boardCountNoComboBox.Text != "內銷" || globals.DBNAME != "進金生")
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單.xlsx";
                }
            }
            else
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2.xlsx";
                }
            }

            if (globals.DBNAME == "宇豐")
            {
                if (modifyDateCheckBox.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\AD\\備貨通知單太陽能.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\AD\\備貨通知單2.xlsx";
                }
            }
            string prepare = shippingCodeTextBox.Text;
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" )
            {
                if (pINOTextBox.Text == "")
                {
                    System.Data.DataTable G1 = Getwh(shippingCodeTextBox.Text);
                    if (G1.Rows.Count > 0)
                    {
                        pINOTextBox.Text = G1.Rows[0][0].ToString();
                    }
                }
                System.Data.DataTable h1 = GetOWHS(shipping_OBUTextBox.Text);
                if (h1.Rows.Count > 0)
                {
                    string dd = Convert.ToString(h1.Rows[0][0]);
                    System.Data.DataTable dt1 = GetOrderData(pINOTextBox.Text, dd);

                    if (dt1.Rows.Count > 0)
                    {
                        DataRow drw3 = dt1.Rows[0];

                        if (shipmentTextBox.Text == "")
                        {
                            int g = drw3["工廠地址"].ToString().IndexOf("司");


                            if (g == -1)
                            {
                                shipmentTextBox.Text = drw3["工廠地址"].ToString();

                            }
                            else
                            {

                                shipmentTextBox.Text = drw3["工廠地址"].ToString().Substring(g + 1).Trim();

                            }
                        }
                        if (arriveDayTextBox.Text == "")
                        {
                            arriveDayTextBox.Text = drw3["連絡人"].ToString();
                        }
                        if (cFSTextBox.Text == "")
                        {
                            cFSTextBox.Text = drw3["電話號碼"].ToString();
                        }
                
                        wH_mainBindingSource.EndEdit();
                        this.wH_mainTableAdapter.Update(wh.WH_main);
                        wh.WH_main.AcceptChanges();
                    }
                }
            }

            DOCMS();

            System.Data.DataTable OrderData = Getprepare(prepare, bbs, DOCM);


            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = "";
            string OWHS = shipping_OBUTextBox.Text.Trim().Replace("(", "").Replace(")", "");
     
            //20130225博豐借出單

            string CARD = cardNameTextBox.Text.Trim().Replace("*", "");
            if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
            {
                CARD = CARD.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");
                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                shipping_OBUTextBox.Text + "備貨通知單(" + CARD + ")--" + pINOTextBox.Text.Trim() + ".xlsx";
            }
            else
            {
                System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                System.Data.DataTable H2 = GetQTY2();
                string G1 = "";

                if (forecastDayTextBox.Text.Trim() == "庫存調撥-借出")
                {
                    G1 = "借出";
                }
                    if (H2.Rows.Count > 0)
                    {

                        string QTY = H2.Rows[0][0].ToString();
                        if (H1.Rows.Count > 0)
                        {
                            int LEN = OWHS.Length;
                            string OWHS1 = shipping_OBUTextBox.Text.Trim().Replace("倉", "").Replace("-", "");
                    

                            CARD = CARD.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");
                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                              OWHS1.Replace("(", "").Replace(")", "") + "備貨通知單(" + CARD + ")" + G1 + "--" + QTY + "PCS.xlsx";

                        }
                        else
                        {
                            StringBuilder sb = new StringBuilder();
                            System.Data.DataTable dt = GetSunny2(shippingCodeTextBox.Text);
                            for (int i = 0; i <= dt.Rows.Count - 1; i++)
                            {

                                DataRow d = dt.Rows[i];


                                sb.Append(d["docentry"].ToString() + " ");


                            }

                            sb.Remove(sb.Length - 1, 1);
                            string fg = sb.ToString();
                            CARD = CARD.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                           "備'" + fg + CARD + G1 + "--" + QTY + "片.xlsx";




                        }
                    }
              //  }
            }
            //

            int GG1 = OutPutFile.LastIndexOf("\\");
            int LL = OutPutFile.Length;
            string FFILE = OutPutFile.Substring(GG1+1, LL - GG1-1);


            int IF = 0;
            if (checkBox6.Checked == true)
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 15;
                }
                else
                {
                    IF = 13;
                }
            }
            else
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 13;
                }
                else
                {
                    IF = 11;
                }
            }
            ExcelReport.APPLE(OrderData, ExcelTemplate, OutPutFile, FFILE, fmLogin.LoginID.ToString().ToLower(), globals.DBNAME,IF,"Y",fmLogin.LoginID.ToString().ToUpper());
            if (fmLogin.LoginID.ToString().ToUpper() == "APPLECHEN")
            {
                if (boardCountNoComboBox.Text == "三角")
                {
                    //序號檔.xlsx 裝箱明细範本.xls

                    string F1 = lsAppDir + "\\Excel\\wh\\序號檔.xlsx";
                    string F2 = lsAppDir + "\\Excel\\wh\\裝箱明细範本.xls";
                    System.Data.DataTable DTOUT = GETOUT(shippingCodeTextBox.Text);

                    if (DTOUT.Rows.Count > 0)
                    {
                        //
                        string O1 = lsAppDir + "\\Excel\\temp\\" +
                DateTime.Now.ToString("yyyyMMddHHmmss") + "序號檔.xlsx";
                        string O2 = lsAppDir + "\\Excel\\temp\\" +
       DateTime.Now.ToString("yyyyMMddHHmmss") + "裝箱明细範本.xls";
                        ExcelReport.ExcelReportOutput(DTOUT, F1, O1, "N");
                        ExcelReport.ExcelReportOutput(DTOUT, F2, O2, "N");

                    }
                }
            }

            if (add4TextBox.Text == "")
            {
                UpdateAPLC4();
            }
        }
        private void DOCMS()
        {
            StringBuilder sbS = new StringBuilder();
            System.Data.DataTable dtS = Getwh(shippingCodeTextBox.Text);
            for (int i = 0; i <= dtS.Rows.Count - 1; i++)
            {

                DataRow dd = dtS.Rows[i];


                sbS.Append(dd["docentry"].ToString() + ",");


            }

            sbS.Remove(sbS.Length - 1, 1);

            DOCM = sbS.ToString();
        }

        private void UPODLN(string U_ACME_INV,string DOCENTRY)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update odln set U_ACME_INV=@U_ACME_INV where DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            add4TextBox.Text = DateTime.Now.ToString();
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        private void UpdateAPLC4()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update wh_main set add4=@aa where shippingcode=@bb");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@aa", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@bb", shippingCodeTextBox.Text));
            add4TextBox.Text = DateTime.Now.ToString();
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdateINVOICE(string INVOICE, string INV, string DOCENTRY1)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update wh_item set INVOICE=@INVOICE,INV=@INV where DOCENTRY1=@DOCENTRY1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@DOCENTRY1", DOCENTRY1));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdateOWTR(string GPSPHONE,string SHIPPINGCODE)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update wh_main set GPSPHONE=@GPSPHONE where shippingcode=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@GPSPHONE", GPSPHONE));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdatINV(string sendGoods, string shippingcode)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update WH_MAIN set sendGoods=@sendGoods where shippingcode=@shippingcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@sendGoods", sendGoods));
               command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdITEM4(string Invoice,string  Docentry1)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update WH_ITEM4 set Invoice=@Invoice  where Docentry1=@Docentry1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@Invoice", Invoice));

            command.Parameters.Add(new SqlParameter("@Docentry1", Docentry1));
            add5TextBox.Text = DateTime.Now.ToString();
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdateAPLC5()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update wh_main set add5=@aa where shippingcode=@bb");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@aa", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@bb", shippingCodeTextBox.Text));
            add5TextBox.Text = DateTime.Now.ToString();
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        private void UpdateAPLC6()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  WH_MAIN SET ADD4=GETDATE() WHERE SHIPPINGCODE IN (SELECT DISTINCT SHIPPINGCODE  FROM WH_Item )");
            sb.Append(" AND ISNULL(ADD4,'') = '' AND SUBSTRING(SHIPPINGCODE,3,4)='2019'");
            sb.Append(" UPDATE  WH_MAIN SET ADD5=GETDATE() WHERE SHIPPINGCODE IN (SELECT DISTINCT SHIPPINGCODE  FROM WH_Item2 )");
            sb.Append(" AND ISNULL(ADD5,'') = '' AND SUBSTRING(SHIPPINGCODE,3,4)='2019'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            add5TextBox.Text = DateTime.Now.ToString();
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public static System.Data.DataTable GETOUT(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT '' B,ITEMCODE 產品編號,Invoice INV,GRADE,T0.PINO 產品名稱,VER,T0.QUANTITY QTY,UPPER(LOCATION) LOC");
            sb.Append("  ,T0.ShippingCode JOBNO, convert(varchar, CAST(T1.CLOSEDAY AS DATETIME),102) 日期,T1.CARDNAME 客戶,T0.DSCRIPTION 品名規格  FROM WH_ITEM T0");
            sb.Append("  LEFT JOIN WH_MAIN T1 ON (T0.ShippingCode =T1.ShippingCode)");
            sb.Append("   WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable Getprepare(string docentry, string bb, string BBS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select itemremark 單據類別,''''+ '" + BBS + "'   單號,Convert(varchar(10),GETDATE(),111)  日期,t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話");
            sb.Append(" ,t0.shipment 送貨地址,t0.bucntctprsn 業務人員,t1.SEQNO 項次,t1.itemcode 產品編號,");
            if (globals.DBNAME == "宇豐")
            {
                sb.Append("t1.FrgnName 品名規格,");
            }
            else
            {
                sb.Append("t1.dscription 品名規格,");
            }
            sb.Append(" t1.pino 料號,t1.quantity 出貨數量,t1.nowqty 現有數量,RTRIM(ISNULL(t0.add1,'')) 備註, REPLACE(CreateName,'KIKI','Lily')   倉管,shipping_obu 倉別,isnull(shipping_obu,'')+");
            sb.Append(" '備貨通知單'+t0.shippingcode 文件名稱,t1.invoice INV,t0.ARType 發票方式,t0.ARTyp2 發票聯式,ARDate 發票日期,ARNumber 發票號碼,t1.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("   '客戶名稱:'+t0.cardname 銷售客戶名稱,t1.grade 等級,t1.ver 版本,T1.INV INVDATE,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,公司=@bb,t1.cardcode 單位,T1.LOCATION 產地,t1.TREETYPE,T0.PACKMEMO 嘜頭,T1.CARDNAME 帳冊,t1.FrgnName 太陽能品名規格,t1.dscription 太陽能料號,''''+U_CUSTITEMCODE 客戶料號,U_CUSTDOCENTRY 客戶PO,t1.NOTICE 注意事項   ");
            sb.Append(" from wh_main t0 left join wh_item t1 on (t0.shippingcode=t1.shippingcode) where t0.shippingcode=@aa ORDER BY SEQNO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
          
            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@bb", bb));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GetDEZUTAO(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select itemcode 產品編號,SUM(CAST(quantity AS DECIMAL)) 出貨數量 from wh_item t0 where shippingcode=@shippingcode GROUP BY itemcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GetOWTR(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T1.WhsName   FROM WTR1 T0 LEFT JOIN OWHS T1 ON (T0.WHSCODE=T1.WHSCODE) WHERE DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable Getprepare2(string docentry, string cc, string dd, string ee, string AR發票, string bb,string DATE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select itemremark 單據類別,t1.docentry 單號,@DATE 日期,t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話");
            sb.Append(" ,送貨地址=@dd,t0.bucntctprsn 業務人員,t1.SEQNO 項次,t1.itemcode 產品編號,case when itmsgrpcod =1032 AND SUBSTRING(t1.itemcode,1,4) <> 'ACME' then t1.frgnname else t1.dscription end 品名規格,");
            sb.Append(" t1.pino 料號,t1.quantity 出貨數量,t1.nowqty 現有數量,T0.unloadCargo 備註,'製單: '+createName 倉管,shipping_obu 倉別,");
            sb.Append(" @ee+t0.shippingcode 文件名稱,t1.invoice INV,t0.ARType 發票方式,t0.ARTyp2 發票聯式,ARDate 發票日期,ARNumber 發票號碼,t1.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("           '客戶名稱:'+t0.cardname 銷售客戶名稱,t1.grade 等級,t1.ver 版本,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,AR=@cc,單據=@ee,AR發票=@AR發票,公司=@bb ");
            sb.Append("   ,''''+U_CUSTITEMCODE  PO料號,U_CUSTDOCENTRY PO");
            sb.Append("  from wh_main t0 left join wh_item2 t1 on (t0.shippingcode=t1.shippingcode) left join acmesql02.dbo.oitm t2 on (t1.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS) where t0.shippingcode=@aa ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@dd", dd));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GETSUNNYS(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.U_ACME_INV,T2.DOCENTRY  FROM AcmeSqlSP.DBO.WH_MAIN T0");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OPDN  T1 ON (SUBSTRING(ADD1,1,14)=T1.U_Shipping_no COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.ODLN  T2 ON (T0.SHIPPINGCODE=T2.U_WH_NO COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE ADD1 LIKE '%整進整出%' ");
            sb.Append(" AND ADD1 LIKE '%SI%' AND BoardCountNo ='內銷' ");
            sb.Append(" AND ISNULL(T2.U_ACME_INV,'') ='' AND ISNULL(T2.U_WH_NO,'') <>''");
            sb.Append(" AND T0.SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE",SHIPPINGCODE ));
;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetPO1(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("       SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(U_CUSTITEMCODE,'')+ISNULL(U_CUSTDOCENTRY,'') <> '' ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetSALES(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT '業務-'+substring(buCntctPrsn,1,CHARINDEX('(', buCntctPrsn)-1)  SALES FROM WH_MAIN  WHERE buCntctPrsn LIKE '%(%' AND SHIPPINGCODE=@SHIPPINGCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetPO2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT U_CUSTITEMCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(U_CUSTITEMCODE,'')<> ''  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GetPO3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT U_CUSTITEMCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(U_CUSTDOCENTRY,'')<> ''  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable Getprepare2S(string docentry, string cc, string dd, string ee, string AR發票, string bb, string DATE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select itemremark 單據類別,t1.docentry 單號,@DATE 日期,t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話");
            sb.Append(" ,送貨地址=@dd,t0.bucntctprsn 業務人員,t1.SEQNO 項次,t1.itemcode 產品編號,");

            if (globals.DBNAME == "宇豐")
            {
                sb.Append("t1.FrgnName 品名規格,");
            }
            else
            {
                sb.Append("case when itmsgrpcod =1032 AND SUBSTRING(t1.itemcode,1,4) <> 'ACME' then (CASE WHEN ISNULL(t1.frgnname,'')='' THEN T2.frgnname ELSE t1.frgnname END COLLATE Chinese_Taiwan_Stroke_CI_AS)  else t1.dscription end 品名規格,");
            }
            sb.Append(" t1.pino 料號,t1.quantity 出貨數量,t1.nowqty 現有數量,T0.unloadCargo 備註,'製單: '+createName 倉管,shipping_obu 倉別,");
            sb.Append(" @ee+t0.shippingcode 文件名稱,t1.invoice INV,t0.ARType 發票方式,t0.ARTyp2 發票聯式,ARDate 發票日期,ARNumber 發票號碼,t1.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("   '客戶名稱:'+t0.cardname 銷售客戶名稱,t1.grade 等級,t1.ver 版本,T1.INV INVDATE,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,AR=@cc,單據=@ee,AR發票=@AR發票,公司=@bb ");
            sb.Append("  ,''  PO料號,   '' PO,T1.U_MEMO 注意事項");
            sb.Append("  from wh_main t0 left join wh_item2 t1 on (t0.shippingcode=t1.shippingcode) left join acmesql02.dbo.oitm t2 on (t1.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS) where t0.shippingcode=@aa ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@dd", dd));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable Getprepare22(string docentry, string bb, string ee, string ff, string gg, string AR發票, string DATE, string 船務,string 備註,string FAX,string SeqDelivery = "")
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select t1.itemremark 單據類別,@DATE 日期,t0.cardname 客戶名稱,t0.arriveday 聯絡人員,t0.cfs 聯絡電話");
            sb.Append(" ,t0.shipment 送貨地址,t0.bucntctprsn 業務人員,t1.SEQNO 項次,t1.itemcode 產品編號,");
            if (globals.DBNAME == "宇豐")
            {
                sb.Append("t1.FrgnName 品名規格,");
            }
            else
            {
                sb.Append("case when itmsgrpcod =1032 AND SUBSTRING(t1.itemcode,1,4) <> 'ACME' then (CASE WHEN ISNULL(t1.frgnname,'')='' THEN T2.frgnname ELSE t1.frgnname END COLLATE Chinese_Taiwan_Stroke_CI_AS)  else t1.dscription end 品名規格,");
            }
            sb.Append(" t1.pino 料號,t1.quantity 出貨數量,t1.nowqty 現有數量,備註=@備註,createName 倉管,shipping_obu 倉別,");
            sb.Append(" '放貨單'+t0.shippingcode 文件名稱,t1.invoice INV,t0.ARType 發票方式,t0.ARTyp2 發票聯式,t0.ARDate 發票日期,t0.ARNumber 發票號碼,t1.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("   '客戶名稱:'+t0.cardname 銷售客戶名稱,t1.grade 等級,t1.ver 版本,T1.INV INVDATE,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,表頭=@dd,T1=@ee,T2=@ff,T3=@gg,單號=@AR發票,船務=@船務,FAX=@FAX ");
            sb.Append(" from wh_main t0 left join wh_item2 t1 on (t0.shippingcode=t1.shippingcode)  ");
            if (SeqDelivery != "") 
            {
                sb.Append(" left join wh_item t3 on (t1.shippingcode=t3.shippingcode and t1.SeqNo = t3.SeqNo)  ");
            }
            sb.Append(" left join acmesql02.dbo.oitm t2 on (t1.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS) where t0.shippingcode=@aa ");
            if (SeqDelivery != "")
            {
                sb.Append(" and t3.SeqDelivery = @SeqDelivery ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@dd", bb));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@ff", ff));
            command.Parameters.Add(new SqlParameter("@gg", gg));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            command.Parameters.Add(new SqlParameter("@船務", 船務));
            command.Parameters.Add(new SqlParameter("@備註", 備註));
            command.Parameters.Add(new SqlParameter("@FAX", FAX));
            if (SeqDelivery != "0")
            {
                command.Parameters.Add(new SqlParameter("@SeqDelivery", SeqDelivery));
            }
            StringBuilder sb2 = new StringBuilder();

            System.Data.DataTable dt2 = wh.WH_Item2;
            string cd = "";
            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {
                DataRow dd = dt2.Rows[i];
              
                if (dd["invoice"].ToString() != cd)
                {

                    sb2.Append(dd["invoice"].ToString()+"/");
                }
               
                cd = dd["invoice"].ToString();
            }
            ef = sb2.ToString();
           
            command.Parameters.Add(new SqlParameter("@bb", ef));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private int GetDeliveryCount(string shippingcode) 
        {
            int i = 0;
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select MAX(IsNull(SeqDelivery,0)) SeqDelivery from wh_item where shippingcode = @shippingcode");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            i = Convert.ToInt32(ds.Tables["wh_main"].Rows[0]["SeqDelivery"]);
            return i;
        }
        public  System.Data.DataTable Getprepare3(string docentry, string dd, string 公司)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select t1.itemremark 單據類別,t1.docentry 單號,t0.quantity 日期,t0.receiveCard 客戶名稱,t1.arriveday 聯絡人員,t1.cfs 聯絡電話");
            sb.Append(" ,t0.shipment 送貨地址,t0.closeday 申請日期,t0.bucntctprsn 業務人員,RANK() OVER (ORDER BY t1.itemcode,T1.SEQNO ) AS 項次,t1.itemcode 產品編號,");
            if (globals.DBNAME == "宇豐")
            {
                sb.Append("t1.A1  品名規格,");
            }
            else
            {
                sb.Append("t1.dscription 品名規格,");
            }
            sb.Append(" case isnull(t1.pino,'') when '' then t1.dscription else t1.pino end COLLATE Chinese_Taiwan_Stroke_CI_AS 料號,t1.quantity 出貨數量,t1.nowqty 現有數量, REPLACE(CreateName,'KIKI','Lily') 倉管,shipping_obu 倉別,isnull(shipping_obu,'')+");
            sb.Append(" '收貨通知單---'+t0.shippingcode 文件名稱,t1.cardcode PCS,ROHS='ROHS',AU='AUS',''''+t0.receiveMemo 備註,");
            if (checkBox3.Checked)
            {
                sb.Append(" T1.DeCust 預出客戶,");
            }
            else
            {
                sb.Append(" '' 預出客戶,");
            }
            if (checkBox4.Checked)
            {
                sb.Append(" REPLACE(t1.ver,'V.','') 版本,");
            }
            else
            {
                sb.Append(" t1.ver 版本,");
            }
            sb.Append(" t1.grade 等級,T1.INV INVDATE,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,DeCust 預進客戶,BoxCheck 外箱檢查,receiveType 國內收貨,closeDay 收貨日期,bb=@bb,公司=@公司,T1.FRGNNAME SHI,T1.LOCATION 產地 ");
            sb.Append(" from wh_main t0 left join wh_item3 t1 on (t0.shippingcode=t1.shippingcode) where t0.shippingcode=@aa ");
            if (checkBox5.Checked)
            {
                sb.Append(" order by BoxCheck ");
            }
            else
            {
                sb.Append(" order by t1.itemcode ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@bb", dd));
            command.Parameters.Add(new SqlParameter("@公司", 公司));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable Getwhitem(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from wh_item where shippingcode=@shippingcode ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable Getwhitem4(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               select SeqNo,T0.Docentry,T0.ItemCode,linenum,ItemRemark,WHName,Dscription,Quantity,Remark,PiNo,NowQty,Ver,Grade,INV,invoice,T0.frgnname,shipdate,u_pay, ");
            sb.Append("               T0.cardcode,U_SHIPDAY,U_SHIPSTATUS,U_MARK,U_MEMO,po,U_CUSTITEMCODE,U_CUSTDOCENTRY,  ");
            sb.Append("                LOCATION,T0.TREETYPE,T1.U_TMODEL MODEL,T1.U_VERSION VERSION,T0.Docentry1 from wh_item4 T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE  COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("  where shippingcode=@shippingcode ORDER BY SEQNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable Getwhitem42(string shippingcode, string DOC)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               select SeqNo,T0.Docentry,T0.ItemCode,linenum,ItemRemark,WHName,Dscription,Quantity,Remark,PiNo,NowQty,Ver,Grade,INV,invoice,T0.frgnname,shipdate,u_pay, ");
            sb.Append("               T0.cardcode,U_SHIPDAY,U_SHIPSTATUS,U_MARK,U_MEMO,po,U_CUSTITEMCODE,U_CUSTDOCENTRY,  ");
            sb.Append("                LOCATION,T0.TREETYPE,T1.U_TMODEL MODEL,T1.U_VERSION VERSION from wh_item4 T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE  COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("  where shippingcode=@shippingcode");
            sb.Append(" AND T0.DOCENTRY1 IN (" + DOC + ") ");
            sb.Append("  ORDER BY SEQNO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable Getwhitem4QTY(string SHIPPINGCODE, string ITEMCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SUM(CAST(QUANTITY AS INT)) QTY,ITEMCODE FROM wh_item4 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE GROUP BY ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GetSHIPDATE(string DOCENTRY, string LINENUM)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              select U_SHIPDAY,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 112 ) SHIPDATE,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 111 ) SHIPDATE2  from RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GetSHIPDATEWTR1(string DOCENTRY, string LINENUM)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              select U_SHIPDAY,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 112 ) SHIPDATE,CONVERT(VARCHAR(10) ,U_ACME_SHIPDAY, 111 ) SHIPDATE2  from WTR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETPACK(string ITEMCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ISNULL(MAX(CAST(CARTONQTY AS INT)),0), ISNULL(MAX(CAST(QTY AS INT)),0)     FROM ACMESQLSP.DBO.WH_PACK2 WHERE ITEMCODE=@ITEMCODE AND QTY <> '空箱'  HAVING  ISNULL(MAX(CAST(CARTONQTY AS INT)),0) <> 0");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable Getwhitem4LAB(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T0.SEQNO,T0.ITEMCODE,T0.DSCRIPTION,T1.U_LOCATION FROM WH_ITEM4 T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where shippingcode=@shippingcode");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable Getwhitem44(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select SeqNo,Docentry,ItemCode,linenum,ItemRemark,WHName,Dscription,Quantity,Remark,PiNo,NowQty,Ver,Grade,INV,invoice,frgnname,shipdate,");
            sb.Append("               cardcode,U_MEMO,CASE WHEN SUBSTRING(T0.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                       SUBSTRING(T0.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                       SUBSTRING(T0.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                      AND SUBSTRING(T0.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T0.[ItemCode],1,9)  ELSE  ");
            sb.Append("               Substring (T0.[ItemCode],2,8) END MODEL,Substring(T0.[ItemCode],12,1) VERSION,TREETYPE,Invoice,LOCATION  from wh_item T0 where shippingcode=@shippingcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GetALL(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.Quantity QTY,T1.PiNo AUNO,T1.CARDCODE UNIT,* FROM ACMESQLSP.DBO.WH_MAIN  T0");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.WH_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable Getwhitem3(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from wh_item3 where shippingcode=@shippingcode ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GetPO(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select distinct po,quantity from wh_item where shippingcode=@shippingcode and isnull(po,'') <> ''  ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                
                DELETEFILE2();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                if (globals.DBNAME == "宇豐")
                {
                    FileName = lsAppDir + "\\Excel\\AD\\宇豐放貨單國內.xls";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\放貨單.xls";
                }
            
                string cc = "";
                string dd = "";
                string ee = "";
                string AR發票 = "AR發票:";
                if (forecastDayTextBox.Text == "庫存調撥-借出")
                {
                    cc = pINOTextBox.Text.ToString();
                    dd = shipmentTextBox.Text.ToString();
                    AR發票 = "借出單號:";
                    ee = "借出單";
                }
                else if (forecastDayTextBox.Text == "庫存調撥-撥倉")
                {
                    cc = pINOTextBox.Text.ToString();
                    dd = shipmentTextBox.Text.ToString();
                    AR發票 = "調撥單號:";
                    ee = "放貨單";
                }
                else
                {
                    cc = add2TextBox.Text.ToString();
                    dd = add3TextBox.Text.ToString();
                    if (dd == "")
                    {
                        dd = shipmentTextBox.Text.ToString();
                    }
                    ee = "放貨單";
                }

                string prepare = shippingCodeTextBox.Text;
                if (checkBox2.Checked == true)
                {
                    if (wH_Item2DataGridView.Rows.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        System.Data.DataTable dt = GetSunny(shippingCodeTextBox.Text);
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {

                            DataRow d = dt.Rows[i];


                            sb.Append(d["docentry"].ToString() + "/");


                        }

                        sb.Remove(sb.Length - 1, 1);
                        fg = sb.ToString();
                        cc = fg;
                        if (forecastDayTextBox.Text == "生產發貨")
                        {
                            AR發票 = "生產發貨:";
                        }
                        else
                        {
                            AR發票 = "銷售訂單:";
                        }
                    }
                }

                wH_mainBindingSource.EndEdit();
                wH_mainTableAdapter.Update(wh.WH_main);
                wh.WH_main.AcceptChanges();

                ViewDATE();
                string CARD = cardNameTextBox.Text.Trim().Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");
                string OWHS = shipping_OBUTextBox.Text.Trim().Replace("(", "").Replace(")", "");
                string DOCTYPE = forecastDayTextBox.Text.Trim();

                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = "";
                if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                {
           

                    OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  shipping_OBUTextBox.Text + "放貨單(" + CARD + ")--" + pINOTextBox.Text.Trim() + ".xls";
                }
                else
                {
                    System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                    System.Data.DataTable H2 = GetQTY();
                    if (DOCTYPE == "庫存調撥-借出")
                    {
                        OutPutFile = lsAppDir + "\\Excel\\temp\\" +
        DateTime.Now.ToString("yyyyMMdd") + CARD + "借出單.xls";
                    }
                    else
                    {
                        if (H1.Rows.Count > 0)
                        {

                            if (H2.Rows.Count > 0)
                            {
                                int LEN = OWHS.Length;
                                string OWHS1 = "";
                                if (LEN <= 3)
                                {
                                    OWHS1 = shipping_OBUTextBox.Text.Trim();
                                }
                                else
                                {
                                    OWHS1 = shipping_OBUTextBox.Text.Trim().Replace("倉", "").Replace("-", "");
                                }

                                string QTY = H2.Rows[0][0].ToString();
                                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                  OWHS1.Replace("(", "").Replace(")", "") + "放貨單(" + CARD + ")" + QTY + "PCS.xls";
                            }
                        }
                        else
                        {
                            string QTY = "";
                            if (H2.Rows.Count > 0)
                            {
                                QTY = H2.Rows[0][0].ToString();
                                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                  // DATE2 + CARD + "放貨單--" + QTY + ".xls";
                                      quantityTextBox.Text.Replace("/","") + CARD + "放貨單--" + QTY + ".xls";
                            }
                        }
                    }
                }

                string B2 = "//acmew08r2ap//table//放貨單//APPLECHEN.JPG";
                string JOJO = "//acmew08r2ap//table//放貨單//JOJOHSU.JPG";
                string S = dtSPLIT().Rows[0][0].ToString();
                System.Data.DataTable PO1 = GetPO1(shippingCodeTextBox.Text);
                System.Data.DataTable S1 = GetSALES(shippingCodeTextBox.Text);
                string SALES = "//acmew08r2ap//table//放貨單//JOJOHSU.JPG";
                if (S1.Rows.Count > 0)
                {
                    SALES = "//acmew08r2ap//table//SIGN//SALES//" + S1.Rows[0][0].ToString() + ".JPG";
                }
                if (PO1.Rows.Count > 0)
                {
                    System.Data.DataTable PO2= GetPO2(shippingCodeTextBox.Text);
                    System.Data.DataTable PO3 = GetPO3(shippingCodeTextBox.Text);
                    string AA = cardNameTextBox.Text.Substring(0, 2);
                    System.Data.DataTable OrderData = Getprepare2(prepare, cc, dd, ee, AR發票, bbs, DATE1);

               
                    if (PO2.Rows.Count == 0)
                    {

                        FileName = lsAppDir + "\\Excel\\wh\\放貨單南京2.xls";
                        if (cTOPTextBox.Text.Trim() == "Checked")
                        {
                            ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "B", "Y", "Y");
                        }
                        else
                        {
                            ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "B", S, "Y");
                        }
                    }
                    else if (PO3.Rows.Count == 0)
                    {

                        FileName = lsAppDir + "\\Excel\\wh\\放貨單南京3.xls";
                        if (cTOPTextBox.Text.Trim() == "Checked")
                        {
                            ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "C", "Y", "Y");
                        }
                        else
                        {
                            ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "C", S, "Y");
                        }
                    }
                    else
                    {
           
                        FileName = lsAppDir + "\\Excel\\wh\\放貨單南京.xls";
                        if (cTOPTextBox.Text.Trim() == "Checked")
                        {
                            ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "A", "Y", "Y");
                        }
                        else
                        {
                            ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "A", S, "Y");
                        }
                    }
                }
                else
                {

                    System.Data.DataTable OrderDataS = Getprepare2S(prepare, cc, dd, ee, AR發票, bbs, DATE1);
                    string FLAG = "";
                    if (globals.DBNAME == "宇豐")
                    {
                        FLAG = "N";
                        ExcelReport.ExcelReportOutputSUNNY(OrderDataS, ExcelTemplate, OutPutFile, FLAG, "Y");
                    }
                    else
                    {
                        if (cTOPTextBox.Text.Trim() == "Checked")
                        {
                            ExcelReport.ExcelReportOutputLA(OrderDataS, ExcelTemplate, OutPutFile, JOJO, SALES, "Y");
                        }
                        else
                        {
                            ExcelReport.ExcelFUNHOUR2(OrderDataS, ExcelTemplate, OutPutFile, B2, S, "Y");
                        }
             
                    }
             
           
                }
                if (add5TextBox.Text == "")
                {
                    UpdateAPLC5();
                }

        
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void wH_Item2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["dataGridViewTextBoxColumn9"].Value = util.GetSeqNo(2, wH_Item2DataGridView);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt1 = Getwhitem(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = wh.WH_Item2;

                if (dt1.Rows.Count == 0)
                {
                    MessageBox.Show("來源無資料，請先存檔");

                    tabControl1.SelectedIndex = 1;
                   
                }

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                    drw2["Docentry"] = drw["Docentry"];
                    drw2["linenum"] = drw["linenum"];
                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["Remark"] = drw["Remark"];
                    drw2["INV"] = drw["INV"];
                    drw2["PiNo"] = drw["PiNo"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["Invoice"] = drw["Invoice"];
                    drw2["BaseDoc"] = drw["Docentry1"];
                    drw2["FrgnName"] = drw["FrgnName"];
                    drw2["Shipdate"] = drw["Shipdate"];
                    drw2["cardcode"] = drw["cardcode"];
                    drw2["U_MEMO"] = drw["U_MEMO"];
                    drw2["FrgnName1"] = drw["FrgnName1"];
                    drw2["TREETYPE"] = drw["TREETYPE"];
                    drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                    drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                    dt2.Rows.Add(drw2);
                }
                System.Data.DataTable ty = GetPO(shippingCodeTextBox.Text);

                StringBuilder sb = new StringBuilder();
                string j4 = "";
                if (ty.Rows.Count > 0)
                {
                    for (int i = 0; i <= ty.Rows.Count - 1; i++)
                    {

                        DataRow d = ty.Rows[i];


                        sb.Append("PO#"+d["po"].ToString() + "*" + d["quantity"].ToString() + "片/");


                    }

                    sb.Remove(sb.Length - 1, 1);
                    j4 = sb.ToString();
                }
                if (forecastDayTextBox.Text == "庫存調撥-借出")
                {
                    string gj = "31.5'S/N:1G0A72W2WDZZ-NS0201" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "上/下午快遞/派車" +
              Environment.NewLine + "客戶自取/業務自送";

                    unloadCargoTextBox.Text = gj;
                }
                else
                {

                    string gj = j4 +
Environment.NewLine + "第1項為FOC" +
Environment.NewLine + "AUS-INVOICE NO#";
                    

                    unloadCargoTextBox.Text = gj;
                }
                if (add5TextBox.Text == "")
                {
                    UpdateAPLC5();
                }
                dollarsKindTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");


                wH_Item2BindingSource.EndEdit();
                this.wH_Item2TableAdapter.Update(wh.WH_Item2);
                wh.WH_Item2.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {
                boatNameTextBox.Text = Convert.ToString(LookupValues[0]);
                boatCompanyTextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (wH_ItemDataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("請輸入項目");
                    return;
                }

                if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                {
                    System.Data.DataTable dt1CHO = GetMenu.GetCHO(cardCodeTextBox.Text);
                    DataRow drw = dt1CHO.Rows[0];

                    oBUShipToTextBox.Text = drw["shipbuilding"].ToString() +
                            Environment.NewLine + drw["shipstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();
                    
                    if (arriveDayTextBox.Text == "")
                    {
                        arriveDayTextBox.Text = drw["shipzipcode"].ToString();
                    }
                    if (cFSTextBox.Text == "")
                    {
                        cFSTextBox.Text = drw["shipblock"].ToString();
                    }

                    oBUBillToTextBox.Text = drw["billbuilding"].ToString() +
                    Environment.NewLine + drw["billstreet"].ToString() +
                    Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                    Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                    Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();
                }
                else
                {



                    if (boatNameTextBox.Text == "")
                    {
                        DataGridViewRow rowt;
                        rowt = wH_Item4DataGridView.Rows[0];
                        string aas = rowt.Cells["ItemRemark"].Value.ToString();
                        System.Data.DataTable dt1s = GetMenu.Getaddress(cardCodeTextBox.Text);
                        DataRow drw1 = dt1s.Rows[0];
                        System.Data.DataTable dt1sar = GetMenu.Getocrdnew2(pINOTextBox.Text, aas);
                        if (pINOTextBox.Text == "" || aas == "發貨單")
                        {
                            oBUShipToTextBox.Text = drw1["mailaddres"].ToString();
                            oBUBillToTextBox.Text = drw1["cardname"].ToString() +
                            Environment.NewLine + drw1["mailaddres"].ToString() +
                            Environment.NewLine + "TEL:" + drw1["phone1"].ToString() +
                            Environment.NewLine + "FAX:" + drw1["fax"].ToString() +
                            Environment.NewLine + "ATTN:" + drw1["cntctprsn"].ToString();

                        }

                        else if (aas == "銷售訂單" || aas == "AR發票")
                        {



                            DataRow drw = dt1sar.Rows[0];

                            oBUShipToTextBox.Text = drw["shipbuilding"].ToString() +
                                    Environment.NewLine + drw["shipstreet"].ToString() +
                                    Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                                    Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                                    Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();


                            oBUBillToTextBox.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();





                        }

                        else
                        {
                            oBUBillToTextBox.Text = dt1sar.Rows[0][0].ToString().Replace("\r", System.Environment.NewLine);
                            oBUShipToTextBox.Text = dt1sar.Rows[0][1].ToString().Replace("\r", System.Environment.NewLine);
                        }
                    }
                    else if (pINOTextBox.Text == "")
                    {
                        MessageBox.Show("請輸入單號");
                    }
                    else
                    {
                        DataGridViewRow rowt;
                        rowt = wH_Item4DataGridView.Rows[0];
                        string aas = rowt.Cells["ItemRemark"].Value.ToString();
                        System.Data.DataTable dt1s = GetMenu.Getaddress(boatNameTextBox.Text);
                        DataRow drw1 = dt1s.Rows[0];
                        System.Data.DataTable dt1sar = GetMenu.Getocrdnew2(pINOTextBox.Text, aas);
                        if (pINOTextBox.Text == "" || aas == "發貨單")
                        {
                            oBUShipToTextBox.Text = drw1["mailaddres"].ToString();
                            oBUBillToTextBox.Text = drw1["cardname"].ToString() +
                            Environment.NewLine + drw1["mailaddres"].ToString() +
                            Environment.NewLine + "TEL:" + drw1["phone1"].ToString() +
                            Environment.NewLine + "FAX:" + drw1["fax"].ToString() +
                            Environment.NewLine + "ATTN:" + drw1["cntctprsn"].ToString();

                        }

                        else if (aas == "銷售訂單" || aas == "AR發票")
                        {



                            DataRow drw = dt1sar.Rows[0];



                            oBUShipToTextBox.Text = drw["shipbuilding"].ToString() +
                                      Environment.NewLine + drw["shipstreet"].ToString() +
                                      Environment.NewLine + "TEL:" + drw["shipblock"].ToString() +
                                      Environment.NewLine + "FAX:" + drw["shipcity"].ToString() +
                                      Environment.NewLine + "ATTN:" + drw["shipzipcode"].ToString();


                            oBUBillToTextBox.Text = drw["billbuilding"].ToString() +
                            Environment.NewLine + drw["billstreet"].ToString() +
                            Environment.NewLine + "TEL:" + drw["billblock"].ToString() +
                            Environment.NewLine + "FAX:" + drw["billcity"].ToString() +
                            Environment.NewLine + "ATTN:" + drw["billzipcode"].ToString();



                        }
                        else
                        {
                            oBUShipToTextBox.Text = dt1sar.Rows[0][0].ToString().Replace("\r", System.Environment.NewLine);
                            oBUBillToTextBox.Text = dt1sar.Rows[0][1].ToString().Replace("\r", System.Environment.NewLine);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DELETEFILE2();
            
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string bb = "";
            string ee = "";
            string ff = "";
            string gg = "";

            if ( globals.DBNAME == "CHOICE" )
            {
                bb = "CHOICE CHANNEL CO.,LTD.";
                ee = "60 market squarem p.o. box364,Belize";
                ff = "BELIZE CITY,BELIZE";
                gg = "";
            }
            else if (globals.DBNAME == "INFINITE")
            {
                bb = "Infinite Power Group Inc.";
                ee = "60 market squarem p.o. box364,Belize";
                ff = "BELIZE CITY,BELIZE";
                gg = "";
            }
            else if (globals.DBNAME == "TOP GARDEN")
            {
                bb = "TOP GARDEN INT'L LTD";
                ee = "60 market squarem p.o. box364,Belize";
                ff = "BELIZE CITY,BELIZE";
                gg = "";
            }
            else if (globals.DBNAME == "禾中")
            {

                bb = "GeTogether Technology Co., Limited";
                ee = "5/F DAH SING LIFE BLDG 99-105";
                ff = "DES VOEUX RD CENTRAL,HONG KONG";
                gg = "";

            }
            else if (globals.DBNAME == "達睿生")
            {

                bb = "达睿生科技发展（深圳）有限公司";
                ee = "Room 407&408，Block 213th Tairan Science Park，";
                ff = "Tairan Ninth Road, Chegongmiao Futian District,SHENZHEN, China";
                gg = "TEL:755-25911195 FAX:0755-25911201";
                
            }
            else if (globals.DBNAME == "宇豐")
            {

                bb = "宇豐光電股份有限公司";
                ee = "4F., No.39, Ln. 76, Ruiguang Rd.,";
                ff = "Neihu Dist., Taipei City 11466, Taiwan";
                gg = "TEL:886-2-8791-2868 FAX:886-02-8791-2869";

            }
            else
            {
                bb = "進金生實業股份有限公司";
                ee = "5F.-3, No.257, Sinhu 2nd Rd.,";
                ff = "Nei-hu District, Taipei Taiwan";
                gg = "TEL:886-2-8791-2868 FAX:886-02-8791-2869";
            }
            if (globals.DBNAME == "宇豐")
            {
                FileName = lsAppDir + "\\Excel\\AD\\宇豐放貨單國外.xls";
            }
            else if (globals.DBNAME == "達睿生")
            {
                string NAME = fmLogin.LoginID.ToString().ToLower();
                if (NAME == "wendyzhang")
                {
                    FileName = lsAppDir + "\\Excel\\DRS\\達睿生放貨.xls";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\DRS\\達睿生放貨2.xls";
                }
            }

            else
            {

                FileName = lsAppDir + "\\Excel\\wh\\進金生.xls";
            }
            string prepare = shippingCodeTextBox.Text;



            if (wH_Item2DataGridView.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                System.Data.DataTable dt = GetSunny(shippingCodeTextBox.Text);
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    DataRow d = dt.Rows[i];


                    sb.Append(d["docentry"].ToString() + "/");


                }

                sb.Remove(sb.Length - 1, 1);
                fg = "'"+sb.ToString();
        
            }
            System.Data.DataTable N1G = Getprepare2G(shippingCodeTextBox.Text);
            StringBuilder sb5 = new StringBuilder();
            string T5 = "";
            if (N1G.Rows.Count > 0)
            {

                for (int i = 0; i <= N1G.Rows.Count - 1; i++)
                {

                    DataRow d = N1G.Rows[i];



                    sb5.Append(d["INV"].ToString() + "/");
                }

                sb5.Remove(sb5.Length - 1, 1);
                T5 = "AUS-INV#" + sb5.ToString();

            }


            wH_mainBindingSource.EndEdit();
            wH_mainTableAdapter.Update(wh.WH_main);
            wh.WH_main.AcceptChanges();


            ViewDATE();
            string FAX = "";
            if (globals.DBNAME == "達睿生")
            {
                FAX = "*請於收到貨後,簽回傳真至FAX:0755-25911201,謝謝!";
            }
            else
            {
                FAX = "*請於收到貨後,簽回傳真至FAX:+886-2-8791-2869,謝謝!";
            }
            string Seq = "";
            if (GetDeliveryCount(shippingCodeTextBox.Text) != 0) 
            {
                Seq = GetMenu.GetWhMainDelivery(shippingCodeTextBox.Text);
            }

            System.Data.DataTable OrderData = Getprepare22(prepare, bb, ee, ff, gg, fg, DATE1, JOBNO, T5, FAX, Seq);


            string CARD = cardNameTextBox.Text.Trim().Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");
            string OWHS = shipping_OBUTextBox.Text.Trim().Replace("(", "").Replace(")", "");
            string DOCTYPE = forecastDayTextBox.Text.Trim();
         
            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = "";
            if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
            {
                if (DOCTYPE == "庫存調撥-借出")
                {
                    OutPutFile = lsAppDir + "\\Excel\\temp\\" +
    DateTime.Now.ToString("yyyyMMdd") + CARD + "借出單.xls";
                }
                else
                {
                    System.Data.DataTable H2 = GetQTY();
                    if (H2.Rows.Count > 0)
                    {
                        string OWHS1 = "";
                        if (shipping_OBUTextBox.Text.Length <= 4)
                        {
                            OWHS1 = shipping_OBUTextBox.Text.Trim();
                        }
                        else
                        {
                            OWHS1 = shipping_OBUTextBox.Text.Trim().Replace("倉", "").Replace("-", "");
                        }

                        string QTY = H2.Rows[0][0].ToString();
                        CARD = CARD.Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

                        OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                        OWHS1 + "放貨單(" + CARD + ")" + QTY + "PCS.xls";
                    }
                    else
                    {
                        OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                            DateTime.Now.ToString("yyyyMMdd") + CARD + Path.GetFileName(FileName);
                    }
                }
            }
            else
            {
                System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                System.Data.DataTable H2 = GetQTY();
                if (DOCTYPE == "庫存調撥-借出")
                {
                    OutPutFile = lsAppDir + "\\Excel\\temp\\" +
    DateTime.Now.ToString("yyyyMMdd") + CARD + "借出單.xls";
                }
                else
                {
                    if (H1.Rows.Count > 0)
                    {

                        if (H2.Rows.Count > 0)
                        {
                            string OWHS1 = "";
                            if (shipping_OBUTextBox.Text.Length <= 4)
                            {
                                OWHS1 = shipping_OBUTextBox.Text.Trim();
                            }
                            else
                            {
                                OWHS1 = shipping_OBUTextBox.Text.Trim().Replace("倉", "").Replace("-", "");
                            }
     
                            string QTY = H2.Rows[0][0].ToString();
                            CARD = CARD.Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                            OWHS1.Replace("(", "").Replace(")", "") + "放貨單(" + CARD + ")" + QTY + "PCS.xls";
                        }
                    }
                    else
                    {
                        OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                   DateTime.Now.ToString("yyyyMMdd") + CARD + Path.GetFileName(FileName);
                    }
                }
            }
            string COMPANY = "ACME";
            if (cACMECheckBox.Checked)
            {
                COMPANY = "ACME";
            }
             if (cIPGICheckBox.Checked)
            {
                COMPANY = "IPGI";
            }
              if (cCHOICECheckBox.Checked)
            {
                COMPANY = "CHOICE";
            }
              if (cTOPTextBox.Text.Trim() == "Checked")
              {
                  COMPANY = "異常" + COMPANY;
              }
            string B2 = "//acmew08r2ap//table//放貨單//";
            string S = dtSPLIT().Rows[0][0].ToString();

            if (globals.DBNAME == "宇豐")
            {
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
           else if (globals.DBNAME == "達睿生")
            {
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                ExcelReport.ExcelFUNHOUR(OrderData, ExcelTemplate, OutPutFile, B2 + COMPANY + ".jpg", S);
            }
            // ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            //string B2 = "//acmew08r2ap//table//SIGN//SALESOUT//";
            //string CNAME = "翁若婷Vivi";
            //if (createNameTextBox.Text == "耿玲玲Milly")
            //{
            //    CNAME = "耿玲玲Milly";
            //}
            //ExcelReport.ExcelReportOutputLA2(OrderData, ExcelTemplate, OutPutFile, B2 + CRENAME + "放貨單" + CNAME + ".JPG");
     
            if (add5TextBox.Text == "")
            {
                UpdateAPLC5();
            }

        }
        private void button78()
        {
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable dt = GetOrderDataAP();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
               
                DataRow dd = dt.Rows[i];
               

                    sb.Append(dd["總類"].ToString() + "-"+ dd["單號"].ToString());
                
              
            }

            System.Data.DataTable dt2 = GetOrderDataAP2();
            if (dt2.Rows.Count > 0)
            {

                sb.Append(" 採購單號-");

                for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                {

                    DataRow dd = dt2.Rows[i];


                    sb.Append(dd[0].ToString() + ",");


                }
                sb.Remove(sb.Length - 1, 1);
            }
            df = sb.ToString();
           
        }

        private void CI1()
        {
            string FileName = string.Empty;
            string FileName1 = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);



            FileName1 = lsAppDir + "\\Excel\\wh\\地址條.xlsx";
            string prepare = shippingCodeTextBox.Text;

            DOCMS();


            System.Data.DataTable OrderData1 = dtcost(bbs);
            //Excel的樣版檔
            string ExcelTemplate = FileName;


            string CARD = cardNameTextBox.Text.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

   
            string MAILNAME = shipping_OBUTextBox.Text + "備貨通知單" + CARD + "--" + shippingCodeTextBox.Text;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" + MAILNAME + ".xlsx";
            OutPutFileDRS = OutPutFile;
            System.Data.DataTable J1 = GetDEZUTAO(prepare);
        

      
                try
                {

                    FileName = lsAppDir + "\\Excel\\wh\\放貨單.xls";
                    string cc = "";
                    string dd = "";
                    string ee = "";
                    string AR發票 = "AR發票:";
                    if (forecastDayTextBox.Text == "庫存調撥-借出")
                    {
                        cc = pINOTextBox.Text.ToString();
                        dd = shipmentTextBox.Text.ToString();
                        AR發票 = "借出單號:";
                        ee = "借出單";
                    }
                    else if (forecastDayTextBox.Text == "庫存調撥-撥倉")
                    {
                        cc = pINOTextBox.Text.ToString();
                        dd = shipmentTextBox.Text.ToString();
                        AR發票 = "調撥單號:";
                        ee = "放貨單";
                    }
                    else
                    {
                        cc = add2TextBox.Text.ToString();
                        dd = add3TextBox.Text.ToString();
                        if (dd == "")
                        {
                            dd = shipmentTextBox.Text.ToString();
                        }
                        ee = "放貨單";
                    }


                    if (checkBox2.Checked == true)
                    {
                        if (wH_Item2DataGridView.Rows.Count > 0)
                        {
                            StringBuilder sb = new StringBuilder();
                            System.Data.DataTable dt = GetSunny(shippingCodeTextBox.Text);
                            for (int i = 0; i <= dt.Rows.Count - 1; i++)
                            {

                                DataRow d = dt.Rows[i];


                                sb.Append(d["docentry"].ToString() + "/");


                            }

                            sb.Remove(sb.Length - 1, 1);
                            fg = sb.ToString();
                            cc = fg;
                            if (forecastDayTextBox.Text == "生產發貨")
                            {
                                AR發票 = "生產發貨:";
                            }
                            else
                            {
                                AR發票 = "銷售訂單:";
                            }
                        }
                    }

                    wH_mainBindingSource.EndEdit();
                    wH_mainTableAdapter.Update(wh.WH_main);
                    wh.WH_main.AcceptChanges();

                    ViewDATE();

                    string OWHS = shipping_OBUTextBox.Text.Trim().Replace("(", "").Replace(")", "");
                    string DOCTYPE = forecastDayTextBox.Text.Trim();


                    if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                    {

                        MAILNAME = shipping_OBUTextBox.Text + "放貨單(" + CARD + ")--" + pINOTextBox.Text.Trim();
                        OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                   MAILNAME + ".xls";
                    }
                    else
                    {
                        string MM = "";
                        int LEN = OWHS.Length;

                        string OWHS1 = OWHS.Trim().Replace("倉", "").Replace("-", "");
                        if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                        {

                            if (forecastDayTextBox.Text == "庫存調撥-借出")
                            {
                                MM = "-借出";
                            }
                            if (forecastDayTextBox.Text == "庫存調撥-借出還回")
                            {
                                MM = "-借出還回";
                            }
                            if (forecastDayTextBox.Text == "庫存調撥-撥倉")
                            {
                                System.Data.DataTable GH1 = GetOWTR(pINOTextBox.Text);

                                string OWHS2 = OWHS.Trim().Replace("-", "") + "調撥回" + GH1.Rows[0][0].ToString().Replace("倉", "");
                                MM = OWHS2;
                            }
                        }

                        System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                        System.Data.DataTable H2 = GetQTY();
                        if (DOCTYPE == "庫存調撥-借出")
                        {

                            OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
            DateTime.Now.ToString("yyyyMMdd") + CARD + MM + "借出單.xls";
                        }
                        else
                        {
                            if (H1.Rows.Count > 0)
                            {

                                if (H2.Rows.Count > 0)
                                {
                              

                                    string QTYB = H2.Rows[0][0].ToString();
                                    MAILNAME = OWHS1.Replace("(", "").Replace(")", "") + "放貨單(" + CARD +MM+ ")" + QTYB + "PCS";
                                    OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
                             MAILNAME + ".xls";
                                }
                            }
                            else
                            {
                                string QTYB = "";
                                if (H2.Rows.Count > 0)
                                {
                                    QTYB = H2.Rows[0][0].ToString();

                                    MAILNAME = quantityTextBox.Text.Replace("/", "") + CARD + MM + "放貨單--" + QTYB;
                                    OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
                                 MAILNAME + ".xls";
                                }
                            }
                        }
                    }

                    string B2 = "//acmew08r2ap//table//放貨單//APPLECHEN.JPG";

                    string JOJO = "//acmew08r2ap//table//放貨單//JOJOHSU.JPG";
                    string S = dtSPLIT().Rows[0][0].ToString();
                    System.Data.DataTable PO1 = GetPO1(shippingCodeTextBox.Text);
                    System.Data.DataTable S1 = GetSALES(shippingCodeTextBox.Text);
                    string SALES = "//acmew08r2ap//table//放貨單//JOJOHSU.JPG";
                    if (S1.Rows.Count > 0)
                    {
                        SALES = "//acmew08r2ap//table//SIGN//SALES//" + S1.Rows[0][0].ToString() + ".JPG";
                    }
                    if (PO1.Rows.Count > 0)
                    {
                        System.Data.DataTable PO2 = GetPO2(shippingCodeTextBox.Text);
                        System.Data.DataTable PO3 = GetPO3(shippingCodeTextBox.Text);
                        string AA = cardNameTextBox.Text.Substring(0, 2);
                        System.Data.DataTable OrderData = Getprepare2(prepare, cc, dd, ee, AR發票, bbs, DATE1);


                        if (PO2.Rows.Count == 0)
                        {

                            FileName = lsAppDir + "\\Excel\\wh\\放貨單南京2.xls";
                            if (cTOPTextBox.Text.Trim() == "Checked")
                            {
                                ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "B", "Y", "N");
                            }
                            else
                            {
                                ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "B", S, "N");
                            }
                        }
                        else if (PO3.Rows.Count == 0)
                        {

                            FileName = lsAppDir + "\\Excel\\wh\\放貨單南京3.xls";
                            if (cTOPTextBox.Text.Trim() == "Checked")
                            {
                                ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "C", "Y", "N");
                            }
                            else
                            {
                                ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "C", S, "N");
                            }
                        }
                        else
                        {

                            FileName = lsAppDir + "\\Excel\\wh\\放貨單南京.xls";
                            if (cTOPTextBox.Text.Trim() == "Checked")
                            {
                                ExcelReport.ExcelHelenPIC(OrderData, FileName, OutPutFile, AA, JOJO, SALES, "A", "Y", "N");
                            }
                            else
                            {
                                ExcelReport.ExcelFUNHOUR3(OrderData, FileName, OutPutFile, AA, B2, "A", S, "N");
                            }
                        }
                    }
                    else
                    {
                        FileName = lsAppDir + "\\Excel\\wh\\放貨單.xls";
                        System.Data.DataTable OrderDataS = Getprepare2S(prepare, cc, dd, ee, AR發票, bbs, DATE1);
                        if (cTOPTextBox.Text.Trim() == "Checked")
                        {
                            ExcelReport.ExcelReportOutputLA(OrderDataS, FileName, OutPutFile, JOJO, SALES, "N");
                        }
                        else
                        {
                            ExcelReport.ExcelFUNHOUR2(OrderDataS, FileName, OutPutFile, B2, S, "N");
                        }


                    }



                    string sourcexlsx = OutPutFile;
                    // PDF 儲存位置
                    string targetpdf = lsAppDir + "\\Excel\\temp\\wh\\"  +
                    MAILNAME + ".PDF";

                    //建立 Excel application instance
                    Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
                    //開啟 Excel 檔案
                    var xlsxDocument = appExcel.Workbooks.Open(sourcexlsx);
                    //匯出為 pdf
                    xlsxDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetpdf);
                    //關閉 Excel 檔
                    xlsxDocument.Close();
                    //結束 Excel
                    appExcel.Quit();
                File.Delete(sourcexlsx);
               // DELETEFILEXLS();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            
        }
        private void button79(string QTY)
        {

            string FileName = string.Empty;
            string FileName1 = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            if (boardCountNoComboBox.Text != "內銷" || globals.DBNAME != "進金生")
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單.xlsx";
                }
            }
            else
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2.xlsx";
                }
            }

            FileName1 = lsAppDir + "\\Excel\\wh\\地址條.xlsx";
            string prepare = shippingCodeTextBox.Text;

            DOCMS();
            System.Data.DataTable OrderData = Getprepare(prepare, bbs, DOCM);

            System.Data.DataTable OrderData1 = dtcost(bbs);
            //Excel的樣版檔
            string ExcelTemplate = FileName;
            string ExcelTemplate1 = FileName1;

            int IF = 0;
            if (checkBox6.Checked == true)
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 15;
                }
                else
                {
                    IF = 13;
                }
            }
            else
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 13;
                }
                else
                {
                    IF = 11;
                }
            }


            string CARD = cardNameTextBox.Text.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

            string OutPutFile1 = lsAppDir + "\\Excel\\temp\\wh\\" +
             CARD + "地址條--" + QTY + "片.xlsx";

            string MAILNAME = shipping_OBUTextBox.Text + "備貨通知單" + CARD + "--" + shippingCodeTextBox.Text;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" + MAILNAME + ".xlsx";
            OutPutFileDRS = OutPutFile;
            System.Data.DataTable J1 = GetDEZUTAO(prepare);
            ExcelReport.APPLE(OrderData, ExcelTemplate, OutPutFile, "N", fmLogin.LoginID.ToString().ToLower(), globals.DBNAME, IF, "N", fmLogin.LoginID.ToString().ToUpper());

            //地址條
            if (boardCountNoComboBox.Text == "內銷")
            {
                ExcelReport.ExcelReportOutpuwh(OrderData1, ExcelTemplate1, OutPutFile1, "Y", J1, shippingCodeTextBox.Text);
            }

        }
        private void button799(string QTY)
        {

            string FileName = string.Empty;
            string FileName1 = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            if (boardCountNoComboBox.Text != "內銷")
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單.xlsx";
                }
            }
            else
            {
                if (checkBox6.Checked)
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2PO.xlsx";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\wh\\備貨通知單2.xlsx";
                }
            }

            FileName1 = lsAppDir + "\\Excel\\wh\\地址條.xlsx";
            string prepare = shippingCodeTextBox.Text;

            DOCMS();
            System.Data.DataTable OrderData = Getprepare(prepare, bbs, DOCM);

            System.Data.DataTable OrderData1 = dtcost(bbs);
            //Excel的樣版檔
            string ExcelTemplate = FileName;
            string ExcelTemplate1 = FileName1;

            int IF = 0;
            if (checkBox6.Checked == true)
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 15;
                }
                else
                {
                    IF = 13;
                }
            }
            else
            {
                if (boardCountNoComboBox.Text != "內銷")
                {
                    IF = 13;
                }
                else
                {
                    IF = 11;
                }
            }


            string CARD = cardNameTextBox.Text.Replace("/", "").Replace("/", "").Replace(".", "").Replace("“", "").Replace("”", "").Replace(":", "");

            string OutPutFile1 = lsAppDir + "\\Excel\\temp\\wh\\" +
             CARD + "地址條--" + QTY + "片.xlsx";
            string MM = "";
            string OWHS = shipping_OBUTextBox.Text;
            int LEN = OWHS.Length;


            if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
            {

                if (forecastDayTextBox.Text == "庫存調撥-借出")
                {
                    MM = "-借出";
                }
                if (forecastDayTextBox.Text == "庫存調撥-借出還回")
                {
                    MM = "-借出還回";
                }
                if (forecastDayTextBox.Text == "庫存調撥-撥倉")
                {
                    System.Data.DataTable GH1 = GetOWTR(pINOTextBox.Text);

                    string OWHS1 = OWHS.Trim().Replace("-", "") + "調撥回" + GH1.Rows[0][0].ToString().Replace("倉", "");

                    MM = "-" + OWHS1;
                }
            }


            string MAILNAME = shipping_OBUTextBox.Text + "備貨通知單" + CARD + MM + "--" + shippingCodeTextBox.Text + "--" + QTY + "PCS";

            string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" + MAILNAME + ".xlsx";
            OutPutFileDRS = OutPutFile;
            System.Data.DataTable J1 = GetDEZUTAO(prepare);
            if (s1CheckBox.Checked)
            {

                ExcelReport.APPLE(OrderData, ExcelTemplate, OutPutFile, "N", fmLogin.LoginID.ToString().ToLower(), globals.DBNAME, IF, "N", fmLogin.LoginID.ToString().ToUpper());
            }
            //快遞
            if (s2CheckBox.Checked)
            {

                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }

                CI1();

                ExcelReport.ExcelReportOutpuwh(OrderData1, ExcelTemplate1, OutPutFile1, "Y", J1, shippingCodeTextBox.Text);
            }

            //派車
            if (s3CheckBox.Checked)
            {
                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }


                CI1();

            }

            //派車
            if (s6CheckBox.Checked)
            {
                ExcelReport.APPLE(OrderData, ExcelTemplate, OutPutFile, "N", fmLogin.LoginID.ToString().ToLower(), globals.DBNAME, IF, "N", fmLogin.LoginID.ToString().ToUpper());

                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }

                CI1();

                ExcelReport.ExcelReportOutpuwh(OrderData1, ExcelTemplate1, OutPutFile1, "Y", J1, shippingCodeTextBox.Text);
            }

            //備加快遞
            if (s7CheckBox.Checked)
            {
                ExcelReport.APPLE(OrderData, ExcelTemplate, OutPutFile, "N", fmLogin.LoginID.ToString().ToLower(), globals.DBNAME, IF, "N", fmLogin.LoginID.ToString().ToUpper());

                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }

                CI1();

           
            }
        }

        private System.Data.DataTable GetOrderDataAP()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct itemremark 總類,cast(docentry as nvarchar) 單號  from WH_Item3 where shippingcode=@aa  and docentry is not null");
           SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOrderDataAP2()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T1.BaseEntry   from WH_Item3 T0 INNER JOIN ACMESQL02.DBO.PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) where shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button8_Click(object sender, EventArgs e)
        {
            button80("Y");
           
        }
        private void button80(string SHOW)
        {
            button78();


            string OutPutFile = "";
            System.Data.DataTable H2 = GetQTY3();
            if (H2.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                FileName = lsAppDir + "\\Excel\\wh\\收貨單.xls";
                string prepare = shippingCodeTextBox.Text;
                System.Data.DataTable OrderData = Getprepare3(prepare, df, bbs);

                string ExcelTemplate = FileName;
                string OWHS = shipping_OBUTextBox.Text;
                int LEN = OWHS.Length;

                string OWHS1 = "";
                if (LEN <= 3)
                {
                    OWHS1 = OWHS.Trim();
                }
                else
                {
                    OWHS1 = OWHS.Trim().Replace("倉", "").Replace("-", "");
                }

                string QTY = H2.Rows[0][0].ToString();
                OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
                  OWHS1 + "收貨通知單---" + prepare + "--" + QTY + "片.xls";
                string SH = "";
                if (!String.IsNullOrEmpty(boardCountTextBox.Text))
                {
                    SH = "--" + boardCountTextBox.Text;
                }

                string INV = "";
                if (!String.IsNullOrEmpty(sendGoodsTextBox.Text))
                {
                    INV = "--" + sendGoodsTextBox.Text;
                }
                MAILSUB = OWHS1 + "收貨通知單---" + prepare + "--" + QTY + "片" + SH + INV;
                if (globals.DBNAME == "宇豐")
                {
                    if (modifyDateCheckBox.Checked)
                    {
                        FileName = lsAppDir + "\\Excel\\AD\\收貨單太陽能.xlsx";
                        ExcelReport.ExcelReportOutputLEMONFIT(OrderData, FileName, OutPutFile, SHOW);
                    }
                    else
                    {
                        if (Getwh2(prepare).Rows.Count > 0)
                        {
                            DOCMS();
                            FileName = lsAppDir + "\\Excel\\AD\\收貨單2.xlsx";
                            System.Data.DataTable OrderData2 = Getprepare(prepare, bbs, DOCM);
                            ExcelReport.ExcelAD(OrderData, FileName, OutPutFile, SHOW, OrderData2);
                        }
                        else
                        {
                            FileName = lsAppDir + "\\Excel\\AD\\收貨單3.xlsx";
                            ExcelReport.ExcelReportOutputLEMONFIT(OrderData, FileName, OutPutFile, SHOW);
                        }
                    }
                }
                else
                {
                    ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, SHOW);
                }
            }

        }
        public System.Data.DataTable GETPAUINV(string INV)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT FPATH,FNAME FROM WH_AUINV　WHERE FNAME LIKE '%PK%' AND FNAME LIKE '%" + INV + "%' ");
           

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OWTR"];
        }

        public System.Data.DataTable dtcost(string COM)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select cardname 公司,shipment 地址,cFS 電話,arriveDay 姓名,");
            if (globals.DBNAME == "CHOICE")
            {
                sb.Append("  'CHOICE'   COM");
    
            }
            else
            {
                sb.Append("''''+ '" + COM + "'   COM");
            }
            sb.Append(" ,CASE WHEN T0.CARDCODE='0060-00' THEN '' ELSE substring(buCntctPrsn,1,3) END 業務,CASE WHEN T0.CARDCODE='0060-00' THEN '' ELSE '#'+officeext END 分機  from  wh_main t0 ");
            sb.Append(" left join acmesql02.dbo.ohem t1 on (substring(t0.buCntctPrsn,1,3)=t1.pager COLLATE Chinese_Taiwan_Stroke_CI_AS) where t0.shippingcode=@aa");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@aa",shippingCodeTextBox.Text ));

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable dtSPLIT()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" Declare @name varchar(100) ");
            sb.Append(" SELECT @name =SUBSTRING(COALESCE(@name + '/',''),0,99)+S   FROM (");
            sb.Append(" SELECT '1一放' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND C1FUN ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '2二放' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND C2FUN ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '3三放' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND C3FUN ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '4可報關' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CKBOU ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '5不可報關' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CBKBOU ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '6虛擬移倉' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CCE ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '7移倉' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CE ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '8提供文件' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CTKWG ='CHECKED'");
            sb.Append(" UNION");
            sb.Append(" SELECT '9已入帳' S FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE AND CRUZAN ='CHECKED'");
            sb.Append(" )AS A");
            sb.Append(" SELECT @name AS A");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private void wH_Item4DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = wH_ItemDataGridView.Rows.Count;

            e.Row.Cells["SeqNo4"].Value = iRecs.ToString();

            e.Row.Cells["Quantity4"].Value = 0;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                int JK1 = 0;
                string memo = "";
                System.Data.DataTable dt3 = GetSHIP(shippingCodeTextBox.Text);

                if (dt3.Rows.Count > 0)
                {
                    if (dt3.Rows[0]["ITEMREMARK"].ToString() == "銷售訂單")
                    {

                        StringBuilder sb2 = new StringBuilder();
                        StringBuilder sb3 = new StringBuilder();
                        for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                        {

                    
                            string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString();
                            string LINENUM = dt3.Rows[i]["LINENUM"].ToString();
                            string ITEMCODE = dt3.Rows[i]["ITEMCODE"].ToString();
                            sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                            string OWHS = shipping_OBUTextBox.Text;
                            if (boardDeliverTextBox.Text == "")
                            {
                                System.Data.DataTable dtG = GetGROSS(DOCENTRY, LINENUM);
                                if (dtG.Rows.Count > 0 && forecastDayTextBox.Text == "銷售訂單")
                                {
                                    MessageBox.Show("單號 " + DOCENTRY + " Lineno " + LINENUM + " 毛利為負數無法存檔，請填寫原因");
                                    JK1 = 1;
                                }

                                System.Data.DataTable dtG2 = GetSTOCK(ITEMCODE, OWHS);

                                if (dtG2.Rows.Count > 0 && forecastDayTextBox.Text == "銷售訂單")
                                {
                                    MessageBox.Show("單號 " + DOCENTRY + " Lineno " + LINENUM + " 庫存數量小於0無法存檔，請填寫原因");
                                    JK1 = 1;
                                }

                            }
                        }
                    }
                }


                if (JK1 == 1)
                {
                    return;
                }



                System.Data.DataTable dt1 = Getwhitem4(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = wh.WH_Item;
                if (dt1.Rows.Count == 0)
                {
                    MessageBox.Show("來源無資料，請先存檔");

                    tabControl1.SelectedIndex = 0;

                }
                int h = 0;
            string DOC="";
                              string LINE ="";
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                     DOC = drw["Docentry"].ToString();
                     LINE = drw["linenum"].ToString();
                    drw2["Docentry"] = DOC;
                    drw2["linenum"] = LINE;
                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    string ITEMCODE = drw["ItemCode"].ToString();
                    drw2["ItemCode"] = ITEMCODE;
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["Remark"] = drw["Remark"];
                    drw2["INV"] = drw["INV"];
                    drw2["PiNo"] = drw["PiNo"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["Invoice"] = drw["Invoice"];
                    drw2["FrgnName"] = drw["FrgnName"];
                    drw2["Shipdate"] = drw["Shipdate"];
                    drw2["cardcode"] = drw["cardcode"];

                    memo = drw["U_MEMO"].ToString();
                    
                    drw2["U_PAY"] = drw["U_PAY"];
                    drw2["U_SHIPDAY"] = drw["U_SHIPDAY"];
                    drw2["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
                    drw2["U_MARK"] = drw["U_MARK"];
                 //   ShipDate2
                    drw2["U_MEMO"] = memo;
                    drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                    drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                    drw2["PO"] = drw["PO"];
                    drw2["TREETYPE"] = drw["TREETYPE"];
                    if (drw["U_PAY"].ToString().Trim() == "FOC")
                    {
                        h = i;
                    
                    }
                    //PQTY5
                    drw2["FrgnName1"] = drw["FrgnName"];
                    drw2["LOCATION"] = drw["LOCATION"];
                    if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                    {
                        System.Data.DataTable SHIPDATE = GetSHIPDATE(DOC, LINE);
                        if (SHIPDATE.Rows.Count > 0)
                        {
                            drw2["ShipDate2"] = SHIPDATE.Rows[0][1].ToString();
                        }
                    }

                    int X1 = ITEMCODE.IndexOf("ACME");
                    if (X1 == -1)
                    {
                        System.Data.DataTable GE1 = GETPACK(ITEMCODE);
                        if (GE1.Rows.Count > 0)
                        {
                            int QTY = Convert.ToInt32(drw["Quantity"]);
                            int FF1 = Convert.ToInt32(GE1.Rows[0][0]);
                            int FF2 = Convert.ToInt32(GE1.Rows[0][1]);
                            int mod = QTY % FF1;
                            int mod2 = QTY % FF2;
                            int mod3 = QTY / FF2;
                            if (mod2 > 0)
                            {
                                mod3 = mod3 + 1;
                            }
                            drw2["PQTY5"] = GE1.Rows[0][0].ToString();
                            drw2["PQTY1"] = GE1.Rows[0][1].ToString();
                            drw2["LPRINT"] = mod3.ToString();
                            drw2["PQTY2"] = mod2.ToString();
                            drw2["PQTY6"] = mod.ToString();
                            //LPRINT
                        }
                    }
                    dt2.Rows.Add(drw2);
                }


                if (memo != "")
                {
                    memo = memo + "\r\n";
                }

                string TDATE = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                {
                    System.Data.DataTable SHIPDATE = GetSHIPDATE(DOC, LINE);
                    if (SHIPDATE.Rows.Count > 0)
                    {
                        TDATE = SHIPDATE.Rows[0][0].ToString();
                    }

                    string USER = fmLogin.LoginID.ToString().ToUpper();

                    if (USER == "EVAHSU")
                    {
                        if (add1TextBox.Text == "")
                        {
                            string gh = memo + TDATE + "派車/下午快/借出/客戶自取/提供序號資料/需貼麥頭/需每箱貼麥頭/需貼箱麥,不需貼板麥/需貼箱麥+需貼板麥/請打三條束帶--提供照片" +
                              Environment.NewLine + TDATE + "請做包裝明細--請打棧板--請打邊條--等貼完提單及麥頭--請提供照片" +
                              Environment.NewLine + TDATE + "借出--請RMA出貨即除帳,產生費用請掛RMA" +
                              Environment.NewLine + TDATE + "下快派車/提供序號資料/新得利司機自提/聯倉司機自提--請提供長寬高及重量" +
                              Environment.NewLine + TDATE + "下快派車/提供序號資料--需至聯倉取貨一起派車/需至新得利取貨一起派車" +
                              Environment.NewLine + TDATE + "下快請併車--提供併車價--請提供長寬高及重量" +
                              Environment.NewLine + "※ 假如沒原箱換,請找合適的空箱,以安全方式換箱,有任何問題再提出討論,謝謝~~" +
                              Environment.NewLine + "※此票調撥為出口,請找箱子無破凹友達箱備貨" +
                              Environment.NewLine + "※送貨單及貨, 請不要有進金生字樣~~快遞單及貨及送貨單, 請不要有進金生字樣~~" +
                              Environment.NewLine + "※請做包裝明細-需打木箱，請提供打木箱前後的照片，請單箱捆膠膜，並提供照片，貨代會安排木箱行去打木箱，打完木箱後請在提供包裝明細(長寬高重量)";
                            add1TextBox.Text = gh;
                        }
                    }
                    else
                    {
                    

                            string gh = memo + TDATE + "派車/下午快/借出/客戶自取/提供序號資料/需貼麥頭/需每箱貼麥頭/需貼箱麥,不需貼板麥/需貼箱麥+需貼板麥/請打三條束帶--提供照片" +
                                Environment.NewLine + TDATE + "請做包裝明細--請打棧板--請打邊條--等貼完提單及麥頭--請提供照片" +
                                Environment.NewLine + TDATE + "借出--請RMA出貨即除帳,產生費用請掛RMA" +
                                Environment.NewLine + TDATE + "下快派車/提供序號資料/新得利司機自提/聯倉司機自提--請提供長寬高及重量" +
                                Environment.NewLine + TDATE + "下快派車/提供序號資料--需至聯倉取貨一起派車/需至新得利取貨一起派車" +
                                Environment.NewLine + TDATE + "下快請併車--提供併車價--請提供長寬高及重量" +
                                Environment.NewLine + "※ 假如沒原箱換,請找合適的空箱,以安全方式換箱,有任何問題再提出討論,謝謝~~" +
                                Environment.NewLine + "※此票調撥為出口,請找箱子無破凹友達箱備貨" +
                                Environment.NewLine + "※送貨單及貨, 請不要有進金生字樣~~快遞單及貨及送貨單, 請不要有進金生字樣~~" +
                                Environment.NewLine + "※請做包裝明細-需打木箱，請提供打木箱前後的照片，請單箱捆膠膜，並提供照片，貨代會安排木箱行去打木箱，打完木箱後請在提供包裝明細(長寬高重量)";
                            add1TextBox.Text = gh;
                        
                    }
              
                }
                    if (add4TextBox.Text == "")
                    {
                        UpdateAPLC4();
                    }
                
                wH_mainBindingSource.EndEdit();
                wH_mainTableAdapter.Update(wh.WH_main);
                wh.WH_main.AcceptChanges();

                wH_ItemBindingSource.EndEdit();
                wH_ItemTableAdapter.Update(wh.WH_Item);
                wh.WH_Item.AcceptChanges();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string OPDN = "";
            try
            {

                receiveCardTextBox.Text = cardNameTextBox.Text;
                System.Data.DataTable dt1 = Getwhitem4(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = wh.WH_Item3;
                if (dt1.Rows.Count == 0)
                {
                    MessageBox.Show("來源無資料，請先存檔");

                    tabControl1.SelectedIndex = 0;

                }

                string gj = "1*20'櫃號:1*40'櫃號:板/箱/貨代:和達/大榮/東南亞/友福/DHL/中菲行/驊洲/聯倉/航通--請告知幾板幾箱及INV";
       
                receiveMemoTextBox.Text = gj;
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                    string DOC = drw["Docentry"].ToString();
             
                    drw2["Docentry"] = DOC;
                    drw2["linenum"] = drw["linenum"].ToString();
                    OPDN = DOC;
  
                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    string Docentry1 = drw["Docentry1"].ToString();
                    string QQ = drw["Quantity"].ToString();
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["Remark"] = drw["Remark"];

                    drw2["PiNo"] = drw["PiNo"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["TREETYPE"] = drw["TREETYPE"];
      

                    drw2["cardcode"] = drw["cardcode"];
                    drw2["LOCATION"] = drw["LOCATION"];
                    drw2["A1"] = drw["frgnname"];
                    if (forecastDayTextBox.Text == "收貨採購單")
                    {
                        string ITEMCODE = drw["ItemCode"].ToString();
                        string QTY = drw["Quantity"].ToString();
                        string one = ITEMCODE.Substring(0, 4);
                        if (one.ToUpper() == "ACME")
                        {
                         
                                System.Data.DataTable dts = GetSHIPOPOR(DOC, drw["Dscription"].ToString());
                                if (dts.Rows.Count > 0)
                                {
                                    for (int J = 0; J <= dts.Rows.Count - 1; J++)
                                    {
                                        string A1 = dts.Rows[0][0].ToString();
                                        string A2 = dts.Rows[0][1].ToString();
                                        string A3 = dts.Rows[0][2].ToString();


                                        drw2["FrgnName"] = A1;
                                        drw2["DeCust"] = A2;
                                        drw2["BoxCheck"] = A3;
                                        UpdITEM4(A3, Docentry1);
                                    }
                                }
                                else
                                {

                                    MessageBox.Show("跟船務料號不同,請確認");
                                }
                      
                        }
                        else
                        {
                            int BE = 0;
                                System.Data.DataTable L1 = GetOITM(ITEMCODE);
                                if (L1.Rows.Count == 0)
                                {
                                    string QTY2 = "";
                                    System.Data.DataTable dt1F = Getwhitem4QTY(shippingCodeTextBox.Text, ITEMCODE);
                                    if (dt1F.Rows.Count > 0)
                                    {
                                        QTY = dt1F.Rows[0][0].ToString();
                                    }

                                    System.Data.DataTable dt1FQ = GetSHIPOPF4(ITEMCODE);
                                    if (dt1FQ.Rows.Count > 0)
                                    {

                                        QTY2 = dt1FQ.Rows[0][0].ToString();
                                    }

                                    System.Data.DataTable dts = GetSHIPOP(DOC, ITEMCODE, QTY);

                                    System.Data.DataTable dtsF2 = GetSHIPOPF2(DOC, ITEMCODE, QQ);
                                    if (dts.Rows.Count == 0)
                                    {
                                        dts = GetSHIPOPF(DOC, ITEMCODE, QQ);
                                    }
                                    if (dtsF2.Rows.Count == 0)
                                    {
                                        BE = 1;
                                        dtsF2 = GetSHIPOPF(DOC, ITEMCODE, QTY2);
                                    }
                                    if (dts.Rows.Count > 0)
                                    {
                                        string A1 = dts.Rows[0][0].ToString();
                                        string A2 = dts.Rows[0][1].ToString();
                                        string A3 = dts.Rows[0][2].ToString();

                                        drw2["FrgnName"] = A1;
                                        drw2["DeCust"] = A2;
                                        drw2["BoxCheck"] = A3;

                                        UpdITEM4(A3, Docentry1);


                                    }
                                    else if (dtsF2.Rows.Count > 0)
                                    {
                                        StringBuilder sb = new StringBuilder();
                                        string A1 = "";
                                        string A2 = "";
                                        string A3 = "";
                                        A1 = dtsF2.Rows[0][0].ToString();
                                        A2 = dtsF2.Rows[0][1].ToString();
                                        A3 = dtsF2.Rows[0][2].ToString();
                                        System.Data.DataTable GBG = GetSHIPOPF3(DOC, ITEMCODE);
                                        if (GBG.Rows.Count > 0)
                                        {
                                            for (int J = 0; J <= GBG.Rows.Count - 1; J++)
                                            {
                                                A2 = GBG.Rows[0][0].ToString();
                                                sb.Append(A2 + "/");
                                            }
                                        }

                                        sb.Remove(sb.Length - 1, 1);

                                        drw2["FrgnName"] = A1;
                                        if (BE == 1)
                                        {
                                            drw2["DeCust"] = A2;
                                        }
                                        else
                                        {

                                            drw2["DeCust"] = sb.ToString();
                                        }
                                        drw2["BoxCheck"] = A3;


                                    }
                                    else
                                    {

                                        MessageBox.Show(ITEMCODE + " 跟船務料號數量不同,請確認");
                                    }
                                }
                                else
                                {
                                    System.Data.DataTable dts = GetSHIPOPOR3(DOC);
                                    if (dts.Rows.Count > 0)
                                    {
                                        string A1 = dts.Rows[0][0].ToString();
                                        drw2["BoxCheck"] = A1;
                                    }

                                }
                        }
                    }


                    if (forecastDayTextBox.Text == "採購單")
                    {
                        System.Data.DataTable dts = GetSHIPOP2(DOC, drw["MODEL"].ToString(), drw["VERSION"].ToString());
                        if (dts.Rows.Count > 0)
                        {
                            for (int J = 0; J <= dts.Rows.Count - 1; J++)
                            {
                                string A1 = dts.Rows[J][0].ToString();
                                string A2 = dts.Rows[J][1].ToString();
                                string A4 = dts.Rows[J][2].ToString();
                                string A5 = dts.Rows[J][3].ToString();

                                drw2["FrgnName"] = A1;
                                drw2["DeCust"] = A2;
                            }
                        }
                    }

   
                        string FF1 = drw2["BoxCheck"].ToString();
                        if (String.IsNullOrEmpty(FF1))
                        {
                            drw2["BoxCheck"] = drw["Invoice"];
                        }
                    
                    dt2.Rows.Add(drw2);
                }
            
                dollarsKindTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");
                if (globals.DBNAME == "達睿生")
                {
                    if (forecastDayTextBox.Text == "採購單")
                    {
                        System.Data.DataTable G1 = GetOPORCUS(pINOTextBox.Text);
                        if (G1.Rows.Count > 0)
                        {
                            receiveMemoTextBox.Text = "对应采购订单#" + pINOTextBox.Text + "（" + G1.Rows[0][0].ToString() + "），共计1PLT（1CTNS）,AU厦门4/8离厂";
                        }
                    }
                }

                if (globals.DBNAME == "宇豐")
                {
                    if (modifyDateCheckBox.Checked)
                    {
                        string gh = "4/30上午/下午1點 友達進貨63板，請於收貨單上簽收回傳宇豐。" +

                        Environment.NewLine + "請注意:" +
                        Environment.NewLine + "一、若實物與收貨單訊息不符(型號/數量), 請向我司反應&確認" +
                        Environment.NewLine + "二、收貨過程中如有凹箱破損…等異常狀況, 請寫在運輸派送簽收單上, 並請將派送單掃描/拍照回傳我司" +
                        Environment.NewLine + "三、若有貨物異常, 請拍照,包括1)破損箱號2)破損那板嘜頭 3)破損處近照";
                        receiveMemoTextBox.Text = gh;
                    }
                }



                wH_Item3BindingSource.EndEdit();
                if (globals.GroupID.ToString().Trim() != "EEP")
                {
                    this.wH_Item3TableAdapter.Update(wh.WH_Item3);
                    wh.WH_Item3.AcceptChanges();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.wH_Item3BindingSource.EndEdit();
        }



        public static System.Data.DataTable Getprepare2G(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT invoice INV FROM wh_item WHERE SHIPPINGCODE=@ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private System.Data.DataTable GetSH(string DocEntry, string ITEMREMARK)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from shipping_item T0 left join shipping_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK=@ITEMREMARK and ISNULL(t1.quantity,'') <> '取消'  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMREMARK", ITEMREMARK));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetSH2(string DocEntry, string ITEMREMARK)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T0.SHIPPINGCODE CODE from shipping_item T0 left join shipping_main t1 on (t0.SHIPPINGCODE=t1.SHIPPINGCODE) ");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND t0.ITEMREMARK=@ITEMREMARK and t1.quantity <> '取消' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMREMARK", ITEMREMARK));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetWHPACK2(string SHIPPINGCODE, string ITEMCODE, string TYPE)
        {

            if (globals.DBNAME == "禾中")
            {
                strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlSPALL;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            }
            if (globals.DBNAME == "達睿生")
            {
                strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlSPDRS;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select SHIPPINGCODE,PLATENO,CARTONNO,AUNO,ITEMCODE,ITEMNAME,GRADE,VER,QTY,CARTONQTY,NW,GW,L,W,H,MATERIAL,LOACTION,FLAG1 from WH_PACK2  WHERE 1=1 ");
            if (TYPE == "A")
            {
                sb.Append(" AND SHIPPINGCODE=@SHIPPINGCODE ");
            }
            if (TYPE == "B")
            {
                sb.Append(" AND ITEMCODE=@ITEMCODE ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetWHPACK3(string SHIPPINGCODE, string ITEMCODE, string TYPE)
        {
            if (globals.DBNAME == "禾中")
            {
                strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlSPALL;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select MAX(CAST(PLATENO AS INT)) PLATENO,SUM(CAST(ISNULL(CASE CARTONNO WHEN '' THEN '0' ELSE CARTONNO END,0) AS decimal))　CARTONNO, ");
            sb.Append(" SUM(CAST(ISNULL(CASE CARTONQTY WHEN '' THEN '0' ELSE CARTONQTY END,0) AS decimal))　CARTONQTY,");
            sb.Append(" SUM(CAST(ISNULL(CASE NW WHEN '' THEN '0' ELSE NW END,0) AS decimal(18,2)))　NW,");
            sb.Append(" SUM(CAST(ISNULL(CASE GW WHEN '' THEN '0' ELSE GW END,0) AS decimal(18,2)))　GW");
            sb.Append("  from WH_PACK2  WHERE 1=1 ");
            if (TYPE == "A")
            {
                sb.Append(" AND SHIPPINGCODE=@SHIPPINGCODE ");
            }
            if (TYPE == "B")
            {
                sb.Append(" AND ITEMCODE=@ITEMCODE ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];



        }


        private System.Data.DataTable GetWHPACK4(string SHIPPINGCODE, string ITEMCODE, string TYPE)
        {
            if (globals.DBNAME == "禾中")
            {
                strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlSPALL;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            }
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(PLATENO) PLATENO FROM (     select MAX(CAST(PLATENO AS INT)) PLATENO");
            sb.Append(" from WH_PACK2  WHERE 1=1 ");
            if (TYPE == "A")
            {
                sb.Append(" AND SHIPPINGCODE=@SHIPPINGCODE ");
            }
            if (TYPE == "B")
            {
                sb.Append(" AND ITEMCODE=@ITEMCODE ");
            }

            sb.Append("   GROUP BY FLAG1 ) AS A");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];



        }
        private System.Data.DataTable GetOWHS(string WHSNAME)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT WHSCODE FROM OWHS WHERE street=@WHSNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHSNAME", WHSNAME));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetCHIITEM(string ITEMCODE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT U_PARTNO PARTNO FROM OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetQTY()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) FROM wH_Item2 WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetQTY2()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) FROM wH_Item WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetPACK()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ShippingCode 工單號碼,Convert(varchar(10),GETDATE(),111)  日期,T0.CardName 客戶名稱,BoardCountNo 貿易形式,INVOICENO INVOICE,T1.ITEMCODE 產品編號,T1.Dscription 產品名稱,GRADE 等級");
            sb.Append(" ,VER 版本,T1.Quantity QTY,'' 高,T1.LOCATION 產地    FROM WH_MAIN T0 LEFT JOIN WH_ITEM T1 ON (T0.ShippingCode=T1.ShippingCode)");
            sb.Append(" WHERE T0.ShippingCode =@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetQTY3()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) FROM wH_Item3 WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GETD1()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  MAX(T0.quantity) DDATE, (T1.ITEMCODE) ITEMCODE, MAX(T1.DSCRIPTION) DSCRIPTION,SUM(CAST(T1.Quantity AS INT)) QTY   ");
            sb.Append("  FROM WH_MAIN T0 LEFT JOIN WH_ITEM3 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append("  WHERE T0.SHIPPINGCODE=@SHIPPINGCODE GROUP BY ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text ));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GETD2(string ITEMCODE)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   Declare @N1 varchar(200) ");
            sb.Append("  select @N1 =SUBSTRING(COALESCE(@N1 + '/',''),0,99) + U_CUSTITEMCODE ");
            sb.Append("  from   (  SELECT  BoxCheck U_CUSTITEMCODE FROM  WH_ITEM3 ");
            sb.Append("  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE) pc");
            sb.Append("  SELECT @N1 A");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GETD3(string ITEMCODE)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  MAX(T0.quantity) DDATE, (T1.ITEMCODE) ITEMCODE, MAX(T1.DSCRIPTION) DSCRIPTION,SUM(CAST(T1.Quantity AS INT)) QTY   ");
            sb.Append("  FROM WH_MAIN T0 LEFT JOIN WH_ITEM3 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append("  WHERE T0.SHIPPINGCODE=@SHIPPINGCODE GROUP BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetOWHSCHI(string ShortName)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("select WareHouseID  from comWareHouse WHERE ShortName =@ShortName ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShortName", ShortName));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetOWHSCHI2(string ShortName)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("select WareHouseID  from comWareHouse WHERE WareHouseName =@ShortName ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ShortName", ShortName));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetFEE()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select * from WH_FEE WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetOHEM(string hometel)
        {
            SqlConnection MyConnection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();
            sb.Append("  select lastname+firstname 姓名 from ohem  where hometel = @hometel ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@hometel", hometel));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetOHEMAD(string hometel)
        {
            SqlConnection MyConnection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();
            sb.Append("  select CASE MOBILE WHEN 'Kiki Lee' THEN 'Lily Lee' ELSE MOBILE END USERNAME  from ohem  where hometel  = @hometel ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@hometel", hometel));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetQTYF(string WH, string ItemCode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(CAST(CAST(T1.QUANTITY AS DECIMAL) AS INT)),0) QTY  FROM WH_MAIN T0");
            sb.Append(" LEFT JOIN WH_Item T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("  WHERE SUBSTRING(T0.SHIPPINGCODE,3,8)=Convert(varchar(8),Getdate(),112)");
            sb.Append("  AND T0.Shipping_OBU =@WH AND T1.ItemCode=@ItemCode");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WH", WH));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

     

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt2 = wh.WH_Item;
            DataRow newCustomersRow = dt2.NewRow();
            int i = wH_ItemDataGridView.CurrentRow.Index;

            DataRow drw = dt2.Rows[i];
            string sa = drw["shippingcode"].ToString();
            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["Docentry"] = drw["Docentry"];
            newCustomersRow["linenum"] = drw["linenum"];
            newCustomersRow["ItemRemark"] = drw["ItemRemark"];
            newCustomersRow["ItemCode"] = drw["ItemCode"];
            newCustomersRow["Dscription"] = drw["Dscription"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Remark"] = drw["Remark"];
            newCustomersRow["INV"] = drw["INV"];
            newCustomersRow["PiNo"] = drw["PiNo"];
            newCustomersRow["NowQty"] = drw["NowQty"];
            newCustomersRow["Ver"] = drw["Ver"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["Invoice"] = drw["Invoice"];
            newCustomersRow["FrgnName"] = drw["FrgnName"];
            newCustomersRow["CardCode"] = drw["CardCode"];
            newCustomersRow["CardName"] = drw["CardName"];
            newCustomersRow["WHName"] = drw["WHName"];
            newCustomersRow["Shipdate"] = drw["Shipdate"];
            newCustomersRow["U_PAY"] = drw["U_PAY"];
            newCustomersRow["U_SHIPDAY"] = drw["U_SHIPDAY"];
            newCustomersRow["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
            newCustomersRow["U_MARK"] = drw["U_MARK"];
            newCustomersRow["U_MEMO"] = drw["U_MEMO"];
            newCustomersRow["PO"] = drw["PO"];
            newCustomersRow["FrgnName1"] = drw["FrgnName1"];
            newCustomersRow["LOCATION"] = drw["LOCATION"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, wH_ItemDataGridView.CurrentRow.Index);

                for (int j = 0; j <= wH_ItemDataGridView.Rows.Count - 2; j++)
                {
                    wH_ItemDataGridView.Rows[j].Cells[0].Value = (j + 1).ToString();
                }

             this.wH_ItemBindingSource.EndEdit();
                this.wH_ItemTableAdapter.Update(wh.WH_Item);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                if (pINOTextBox.Text == "")
                {
                    MessageBox.Show("請輸入單號");
                    pINOTextBox.Focus();
                    return;
                }

                string check1 = "";
                if (checkBox1.Checked)
                {
                    check1 = "a";
                }
                else
                {
                    check1 = "b";
                }

                System.Data.DataTable h1 = null;
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" ||  globals.DBNAME == "測試區98")
                {
                    h1 = GetOWHS(shipping_OBUTextBox.Text);
                }
                else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                {
                    h1 = GetOWHSCHI(shipping_OBUTextBox.Text);
                 
                }

                string docentry = pINOTextBox.Text;
               // h1 = GetOWHS(shipping_OBUTextBox.Text);
                string dd = Convert.ToString(h1.Rows[0][0]);
                System.Data.DataTable dt1 = null;
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" ||  globals.DBNAME == "測試區98")
                {
                    dt1 = GetOrderData(docentry, dd);
                }
                else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
                {
                    dt1 = GetOrderDataCHI(docentry, dd, check1);
                }

                System.Data.DataTable dt2 = null;

                dt2 = wh.WH_Item4;
                int M1 = 0;
                string MCODE = "";
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw3 = dt1.Rows[0];
              


                        int g = drw3["工廠地址"].ToString().IndexOf("司");


                        if (g == -1)
                        {
                            shipmentTextBox.Text = drw3["工廠地址"].ToString();

                        }
                        else
                        {

                            shipmentTextBox.Text = drw3["工廠地址"].ToString().Substring(g + 1).Trim();

                        }

                    
                    if (cardCodeTextBox.Text == "")
                    {
                        cardCodeTextBox.Text = drw3["客戶編號"].ToString();
                        cardNameTextBox.Text = drw3["客戶名稱"].ToString();
                    }
                    arriveDayTextBox.Text = drw3["連絡人"].ToString();
                    cFSTextBox.Text = drw3["電話號碼"].ToString();
    
                    buCntctPrsnTextBox.Text = drw3["業務"].ToString();

                    if (forecastDayTextBox.Text == "採購單" || forecastDayTextBox.Text == "收貨採購單" || forecastDayTextBox.Text == "採購退貨" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "AR貸項通知單")
                    {
                        receiveMemoTextBox.Text = drw3["備註"].ToString();
                    }
                    //20210720
                    //else
                    //{
                    //    if (fmLogin.LoginID.ToString().ToUpper() != "ESTHERYEH")
                    //    {
                    //        textBox3.Text = drw3["備註"].ToString();
                    //    }
              
                    //}

                   // quantityTextBox.Text = drw3["交貨日期"].ToString();
              
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["seqNo"] = "0";
                    drw2["Docentry"] = pINOTextBox.Text;
                    drw2["Dscription"] = drw["品名規格"];
      
                    drw2["itemcode"] = drw["產品編號"];
                    drw2["ShipDate"] = drw["排程日期"];
                    
                    drw2["ItemRemark"] = forecastDayTextBox.Text;
                    drw2["WHName"] = shipping_OBUTextBox.Text;
                    string d = drw["品名規格"].ToString();
                    decimal SS = Convert.ToDecimal(drw["數量"]);
                    string GH = Convert.ToDouble(SS).ToString();
                    drw2["Quantity"] = GH;

                    drw2["linenum"] = drw["欄號"];
                    if (fmLogin.LoginID.ToString().ToUpper() == "ESTHERYEH")
                    {
                        System.Data.DataTable GE = GetORDRE(docentry, drw["欄號"].ToString());
                        if (GE.Rows.Count > 0)
                        {


                            string U_SHIPDAY = GE.Rows[0]["U_SHIPDAY"].ToString().Trim().ToUpper();
                            string U_MARK = GE.Rows[0]["U_MARK"].ToString().Trim().ToUpper();
                            if (U_MARK == "X")
                            {
                                U_MARK = "";
                            }
                            else
                            {
                                U_MARK = "請貼" + U_MARK;
                            }


                            textBox3.Text = U_SHIPDAY +
                                Environment.NewLine + U_MARK;
                        }
                    }
                    string 產品編號 = drw["產品編號"].ToString();
                    System.Data.DataTable QTY1 = GetQTYF(shipping_OBUTextBox.Text.ToString(), 產品編號);
               
                        QTYF = Convert.ToInt32(QTY1.Rows[0][0]);
                    
                    drw2["NowQty"] = Convert.ToInt32(drw["現有數量"]) - QTYF;

                    drw2["Ver"] = drw["版本"];
                    drw2["WHSCODE"] = dd;
                    string 品名規格 = drw["品名規格"].ToString();
                    if (forecastDayTextBox.Text == "採購單")
                    {


                        System.Data.DataTable T1 = GetARRIVE(pINOTextBox.Text, drw["欄號"].ToString());
                        if (T1.Rows.Count > 0)
                        {
                            drw2["ShipDate"] = T1.Rows[0][0].ToString();

                            quantityTextBox.Text = T1.Rows[0][1].ToString();
              
                        }
                    }

      
                    string gg = 產品編號.Substring(0, 3).ToString().ToUpper();
                    if (gg == "TAP")
                    {

                        int G1 = 品名規格.IndexOf(".");
                        if (G1 != -1)
                        {
                            drw2["Ver"] = "V." + 品名規格.Substring(G1 + 1, 1);
                        }
                    }
                    drw2["Grade"] = drw["等級"];
                    drw2["cardcode"] = drw["單位"];

                    drw2["U_PAY"] = drw["付款"];
                    drw2["U_SHIPDAY"] = drw["押出貨日"];
                    drw2["U_SHIPSTATUS"] = drw["貨況"];
                    drw2["U_MARK"] = drw["特殊嘜頭"];
                    drw2["U_MEMO"] = drw["注意事項"];
                    drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                    drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                    drw2["PO"] = drw["PO"];
                    drw2["LOCATION"] = drw["產地"];
                    drw2["TREETYPE"] = drw["TREETYPE"];

                    string TREETYPE = drw["TREETYPE"].ToString();
                    drw2["TREETYPE"] = TREETYPE;
                    try
                    {
                        hjj = "";

                        if (TREETYPE == "S")
                        {
                            hjj = "母料號";
                            MCODE = 產品編號;
                            M1 = 0;
                        }
                        else if (TREETYPE == "I")
                        {

                            M1++;
                            hjj = MCODE + "-子料號-" + M1.ToString();
                        }
                        else
                        {
                            hjj = drw["PARTNO"].ToString();
                        }

                        drw2["pino"] = hjj;
                    }
                    catch
                    {

                    }

                    System.Data.DataTable L1 = GetCHIITEM(產品編號);
                    if (L1.Rows.Count > 0)
                    {
                        if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐")
                        {
                            drw2["pino"] = L1.Rows[0]["PARTNO"].ToString();

                        }
                    }
                    if (globals.DBNAME == "達睿生")
                    {
                        if (forecastDayTextBox.Text == "採購單")
                        {
                            string invoice = drw["INVOICE"].ToString();
                            drw2["invoice"] = invoice;
                        }

                        drw2["FrgnName"] = drw["品名規格"];
                    }
                    else
                    {
                        drw2["FrgnName"] = drw["品名規格1"];
                        if (cardCodeTextBox.Text == "0017-00")
                        {
                            drw2["FrgnName"] = drw["品名規格1"] + "-" + drw["等級"];

                        }
                    }

                    dt2.Rows.Add(drw2);



                }

                for (int j = 0; j <= wH_Item4DataGridView.Rows.Count - 2; j++)
                {
                    wH_Item4DataGridView.Rows[j].Cells[0].Value = j.ToString();
                }
               
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            wH_mainBindingSource.EndEdit();
            //wH_mainTableAdapter.Update(wh.WH_main);
            //wh.WH_main.AcceptChanges();

            wH_Item4BindingSource.EndEdit();
            button6_Click(sender, e);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string ss = cardCodeTextBox.Text.ToString();
             string tt = forecastDayTextBox.Text;
            object[] LookupValues = GetCardList(ss, tt);



            if (LookupValues != null)
            {


                       StringBuilder sb = new StringBuilder();


                for (int i = 0; i <= LookupValues.Length - 1; i++)
                {

                    sb.Append("'" + Convert.ToString(LookupValues[i]) + "',");

                }
                sb.Remove(sb.Length - 1, 1);
                string ds = sb.ToString();
               
               
                    try
                    {
                        string docentry = pINOTextBox.Text;
                        System.Data.DataTable h1 = GetOWHS(shipping_OBUTextBox.Text);
                        string dd = Convert.ToString(h1.Rows[0][0]);
                        System.Data.DataTable dt1 = GetOrderData(ds, dd);

                        System.Data.DataTable dt2 = null;

                        dt2 = wh.WH_Item4;

                        for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                        {
                            DataRow drw3 = dt1.Rows[0];
                    


                                int g = drw3["工廠地址"].ToString().IndexOf("司");


                                if (g == -1)
                                {
                                    shipmentTextBox.Text = drw3["工廠地址"].ToString();

                                }
                                else
                                {

                                    shipmentTextBox.Text = drw3["工廠地址"].ToString().Substring(g + 1).Trim();

                                }

                            
                            arriveDayTextBox.Text = drw3["連絡人"].ToString();
                            cFSTextBox.Text = drw3["電話號碼"].ToString();
                            buCntctPrsnTextBox.Text = drw3["業務"].ToString();

                            if (forecastDayTextBox.Text == "採購單" || forecastDayTextBox.Text == "收貨採購單" || forecastDayTextBox.Text == "採購退貨" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "AR貸項通知單")
                            {
                                receiveMemoTextBox.Text = drw3["備註"].ToString();
                            }
                            //20210720
                            //else
                            //{
                            //    add1TextBox.Text = drw3["備註"].ToString();
                            //}

                      //      quantityTextBox.Text = drw3["交貨日期"].ToString();

                            DataRow drw = dt1.Rows[i];
                            DataRow drw2 = dt2.NewRow();
                            drw2["ShippingCode"] = shippingCodeTextBox.Text;
                            drw2["seqNo"] = "0";
                            drw2["Docentry"] = drw["單號"];
                            drw2["Dscription"] = drw["品名規格"];
                            drw2["itemcode"] = drw["產品編號"];
                            drw2["ItemRemark"] = forecastDayTextBox.Text;
                            drw["單號"] = drw["單號"];
                            drw2["linenum"] = drw["欄號"];
                            string 產品編號 = drw["產品編號"].ToString();
                            System.Data.DataTable QTY1 = GetQTYF(shipping_OBUTextBox.Text.ToString(), 產品編號);
                            if (QTY1.Rows.Count > 0)
                            {
                                QTYF = Convert.ToInt32(QTY1.Rows[0][0]);
                            }
                            drw2["NowQty"] = Convert.ToInt32(drw["現有數量"]) - QTYF;
                            drw2["ShipDate"] = drw["排程日期"];
                            drw2["Ver"] = drw["版本"];

                            string 品名規格 = drw["品名規格"].ToString();
                            string gg = 產品編號.Substring(0, 3).ToString().ToUpper();
                            if (gg == "TAP")
                            {

                                int G1 = 品名規格.IndexOf(".");
                                if (G1 != -1)
                                {
                                    drw2["Ver"] = "V." + 品名規格.Substring(G1 + 1, 1);
                                }
                            }
                            drw2["Grade"] = drw["等級"];
                            drw2["cardcode"] = drw["單位"];
                            drw2["U_PAY"] = drw["付款"];
                            drw2["U_SHIPDAY"] = drw["押出貨日"];
                            drw2["U_SHIPSTATUS"] = drw["貨況"];
                            drw2["U_MARK"] = drw["特殊嘜頭"];
                            drw2["U_MEMO"] = drw["注意事項"];
                            drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                            drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                            drw2["PO"] = drw["PO"];
                            drw2["LOCATION"] = drw["產地"];
                            drw2["TREETYPE"] = drw["TREETYPE"];
                            drw2["pino"] = drw["PARTNO"];
                            if (globals.DBNAME == "達睿生")
                            {
                                if (forecastDayTextBox.Text == "採購單")
                                {
                                    drw2["invoice"] = drw["INVOICE"].ToString();
                                }

                                drw2["FrgnName"] = drw["品名規格"];
                            }
                            else
                            {
                                drw2["FrgnName"] = drw["品名規格1"];
                                if (cardCodeTextBox.Text == "0017-00")
                                {
                                    drw2["FrgnName"] = drw["品名規格1"] + "-" + drw["等級"];

                                }
                            }
                            dt2.Rows.Add(drw2);



                        }

                        for (int j = 0; j <= wH_Item4DataGridView.Rows.Count - 2; j++)
                        {
                            wH_Item4DataGridView.Rows[j].Cells[0].Value = j.ToString();
                        }
                      
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                

            }

            wH_mainBindingSource.EndEdit();
            //wH_mainTableAdapter.Update(wh.WH_main);
            //wh.WH_main.AcceptChanges();

            wH_Item4BindingSource.EndEdit();
            //wH_Item4TableAdapter.Update(wh.WH_Item4);
            //wh.WH_Item4.AcceptChanges();
      
        }
        public static System.Data.DataTable GERCARD(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection ;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT CARDCODE,CARDNAME FROM ORDR WHERE DOCENTRY=@DOCENTRY ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GERCARD3(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT CARDCODE,CARDNAME FROM OWTR WHERE DOCENTRY=@DOCENTRY ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GERCARD2(string BillNO)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CustomerID CARDCODE,T2.FullName CARDNAME FROM OrdBillMain T0 ");
            sb.Append(" Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID AND T2.Flag =1) ");
            sb.Append(" where t0.Flag =2 and T0.BillNO = @BillNO   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private object[] GetCardList(string aa, string bb)
        {

            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };


            string SqlScript = "";
            if (bb == "AR發票")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111)  as  過帳日期  FROM oinv T0 where T0.cardcode='" + aa + "' ";

            }
            else if (bb == "銷售訂單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM ORDR T0 where  T0.cardcode='" + aa + "' ";

            }
            else if (bb == "庫存調撥-借出")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='1' and T0.cardcode='" + aa + "' ";

            }
            else if (bb == "發貨單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Oige T0  ";

            }
            else if (bb == "庫存調撥-撥倉")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='3' ";

            }
            else if (bb == "採購單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Opor T0 where  T0.cardcode='" + aa + "' ";

            }
            else if (bb == "收貨採購單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Opdn T0 where  T0.cardcode='" + aa + "' ";

            }
            else if (bb == "採購退貨")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM ORPD T0 where  T0.cardcode='" + aa + "' ";

            }
            else if (bb == "庫存調撥-借出還回")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='2' and T0.cardcode='" + aa + "' ";

            }
            else if (bb == "AR貸項通知單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Orin T0 where  T0.cardcode='" + aa + "' ";

            }
            else if (bb == "收貨單")
            {
                SqlScript = "SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Oign T0  ";

            }




            MultiValueDialog dialog = new MultiValueDialog();



            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;

            dialog.KeyFieldName = "單號";



            dialog.SqlScript = SqlScript;

            try
            {





                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;

                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {

                DELETEFILE();
                DialogResult result;
                result = MessageBox.Show("收件人地址為" + LOGINID + "是否要寄出", "Close", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    System.Data.DataTable H2 = GetQTY2();
                    if (H2.Rows.Count > 0)
                    {
                        string QTY = H2.Rows[0][0].ToString();
                        button79(QTY);


                        string template;
                        StreamReader objReader;
                        string FileName = string.Empty;
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                        FileName = lsAppDir + "\\MailTemplates\\wh.htm";
                        objReader = new StreamReader(FileName);

                        template = objReader.ReadToEnd();
                        objReader.Close();
                        objReader.Dispose();



                        StringWriter writer = new StringWriter();
                        HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

                        string Html = GetTODO_USERDataSource2();
                        template = template.Replace("##Content##", Html);

                        string h = fmLogin.LoginID.ToString();

                        try
                        {
                            System.Data.DataTable dt1 = GetMenu.Getemployee(h);
                            DataRow drw = dt1.Rows[0];
                            if ((dt1.Rows.Count) > 0)
                            {
                                string a1 = drw["pager"].ToString();
                                string a2 = drw["mobile"].ToString();
                                template = template.Replace("##eng##", a2);
                                template = template.Replace("##name##", a1);
                                template = template.Replace("##mail##", h + "@acmepoint.com");
                            }
                        }
                        catch
                        {
                            template = template.Replace("##eng##", "");
                            template = template.Replace("##name##", "");
                            template = template.Replace("##mail##", h + "@acmepoint.com");
                        }

                        MailMessage message = new MailMessage();

                        message.To.Add(LOGINID);

                        StringBuilder sb = new StringBuilder();
                        System.Data.DataTable dt = Getwh(shippingCodeTextBox.Text);
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {

                            DataRow dd = dt.Rows[i];


                            sb.Append(dd["docentry"].ToString() + "/");


                        }

                        sb.Remove(sb.Length - 1, 1);
                        fg = sb.ToString();
             
                        message.Subject = shipping_OBUTextBox.Text + "--備貨單通知--" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                        message.Body = template;
                        if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98" || globals.DBNAME == "宇豐" || globals.DBNAME == "INFINITE" || globals.DBNAME ==  "CHOICE")
                        {
                        string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                        string[] filenames = Directory.GetFiles(OutPutFile);
                        foreach (string file in filenames)
                        {

                            string m_File = "";

                            m_File = file;
                            data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                            //附件资料
                            ContentDisposition disposition = data.ContentDisposition;


                            // 加入邮件附件
                            message.Attachments.Add(data);

                            int G1 = file.IndexOf("備貨通知單");
                            if (G1 != -1)
                            {
                                UPLOAD(file);
                            }
                        }

                 
                            string aaa = boardCountTextBox.Text;

                            if (!String.IsNullOrEmpty(aaa))
                            {
                                string[] arrurla = aaa.Split(new Char[] { ',' });

                                foreach (string i in arrurla)
                                {

                                    System.Data.DataTable GG1 = download21(i);
                                    if (GG1.Rows.Count > 0)
                                    {
                                        for (int s = 0; s <= GG1.Rows.Count - 1; s++)
                                        {
                                            string PATH = GG1.Rows[s][0].ToString();
                                            string m_File = "";

                                            m_File = PATH;
                                            data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);
                                            ContentDisposition disposition = data.ContentDisposition;
                                            message.Attachments.Add(data);
                                        }
                                    }
                                }
                            }
                        }


                        message.IsBodyHtml = true;

                        SmtpClient client = new SmtpClient();
                        client.Send(message);
                        if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98" || globals.DBNAME == "宇豐" || globals.DBNAME == "INFINITE")
                        {
                            data.Dispose();
                            message.Attachments.Dispose();
                        }

                        DELETEFILE();
                        MessageBox.Show("寄信成功");
                    }
                }
            }
            catch (Exception ex)
            {
                DELETEFILE();
                MessageBox.Show(ex.Message);
            }
        }


        private void DELETEFILE()
        {
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void DELETEFILEXLS()
        {
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {
                    if (file.IndexOf("xls") != -1)
                    {

                        File.Delete(file);
                    }

                }
            }
            catch { }
        }
        private void DELETEFILE2()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private string GetTODO_USERDataSource2()
        {
            System.Data.DataTable dtEvent = GetMenu.GetMail2(shippingCodeTextBox.Text);

            string html = string.Empty;
            string DateGroup = string.Empty;
            string Docentry1 = "單號";
            string itemcode1 = "產品編號";
            string pino1 = "PART NO";
            string Quantity1 = "數量";
            html = html + " <table><tr height=14 style='height:10.25pt'><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + Docentry1 + "</td><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + itemcode1 + "</td><td nowrap style='width:30%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + pino1 + "</td><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + Quantity1 + "</td></tr>";

            foreach (DataRow row in dtEvent.Rows)
            {
                string Docentry = Convert.ToString(row["Docentry"]);
                string itemcode = Convert.ToString(row["itemcode"]);
                string pino = Convert.ToString(row["pino"]);
                string Quantity = Convert.ToString(row["Quantity"]);
                html = html + " <tr height=14 style='height:10.25pt'><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + Docentry + "</td><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + itemcode + "</td><td nowrap style='width:30%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + pino + "</td><td nowrap style='width:14%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 0cm 0cm 0cm;height:10pt' bgcolor='#CCFFCC'><font size=2 face=Arial>" + Quantity + "</td></tr>";
            }
            html = html + "</table>";
            string Remarks = "<dl>";
            Remarks = Remarks + "<dt>1.	請依附檔安排備貨,並回傳 明細& 出貨序號  & 照片  給我司 , 待我司放貨單再行出貨</dt>";
            Remarks = Remarks + "<dt>2.	理貨資料發出給Acme時間</dt>";
            Remarks = Remarks + "<dt>3.	麥頭:如備貨單所示</dt>";
            Remarks = Remarks + "<dt>4.	備註:" + add1TextBox.Text + "</dt>";
            Remarks = Remarks + "</dl>";
            html += Remarks;

            return html;
        }


        public System.Data.DataTable Getwh(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct docentry from wh_item where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable Getwh2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct docentry from wh_item where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable GetAUINV(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct BoxCheck INV from wh_item3 where shippingcode=@Docentry AND ISNULL(BoxCheck,'') <> '' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetORDR()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DOCENTRY,LINENUM FROM WH_ITEM4 WHERE SHIPPINGCODE=@SHIPPINGCODE ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable GetORDRE(string DOCENTRY, string LINENUM)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "SELECT  U_SHIPDAY,U_MARK  FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM AND ISNULL(U_SHIPDAY,'') <> '' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetAUINV2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct Invoice INV from wh_item where shippingcode=@Docentry AND ISNULL(Invoice,'') <> ''";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

   


        public System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM WH_ITEM4 T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMREMARK in ('銷售訂單','銷售單') ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetSHIPDIAO(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM WH_ITEM4 T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND  itemremark like '%調撥%' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable GetSHITSAIGO(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM,ITEMCODE FROM WH_ITEM4 T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND  itemremark ='採購單' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetSHIPF(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT FrgnName FROM WH_ITEM3　WHERE SHIPPINGCODE=@SHIPPINGCODE AND  SUBSTRING(FrgnName,1,2)='SH' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetGROSS(string docentry, string LINENUM)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select  docentry,price,(grossbuypr/ isnull(case rate when 0 then 1 else rate end,1)) from rdr1 where price-(grossbuypr/ isnull(case rate when 0 then 1 else rate end,1)) < 0 AND docentry=@docentry AND LINENUM=@LINENUM ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable GetSTOCK(string itemcode, string WHSNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select onhand from oitw T0 LEFT JOIN OWHS T1 ON (T0.WHSCODE=T1.WHSCODE) WHERE itemcode=@itemcode and WHSNAME=@WHSNAME and onhand <= 0 ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@itemcode", itemcode));
            command.Parameters.Add(new SqlParameter("@WHSNAME", WHSNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetSHIP2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT DOCENTRY,LINENUM FROM WH_ITEM T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMREMARK='銷售訂單'";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetSunny(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct docentry from wh_item2 where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }

        public System.Data.DataTable GetSunny2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct docentry from wh_item where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
 
        private bool CheckSerial(string sData, ref string FieldValue,System.Data.DataTable dt)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(dt.Rows[0][FieldName]);
                return true;
            }
            return false;
        }



        private bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }

            //}
            return false;
        }



        private void SetRow(int iRow, string sData, ref string FieldValue, System.Data.DataTable dt,string LINE)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);

                int QTY = Convert.ToInt32(dt.Rows[iRow]["數量"]);

                decimal PPRICE = Convert.ToDecimal(dt.Rows[iRow]["PRICE"]);
                if (FieldName == "AMT")
                {

                    FieldValue = Convert.ToString(QTY * PPRICE);

                }
                else if (FieldName == "NO")
                {

                    FieldValue = LINE;

                }
                else
                {
                    FieldValue = Convert.ToString(dt.Rows[iRow][FieldName]);
                }
            }

        }
        public static System.Data.DataTable GetPRICE(string DOCENTRY, string LINENUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T1.PRICE FROM ACMESQL02.DBO.PDN1 T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.POR1 T1 ON (T0.BASEENTRY=T1.DOCENTRY AND T0.BASELINE=T1.LINENUM)");
            sb.Append("  WHERE T0.DOCENTRY=@DOCENTRY AND T0.LINENUM=@LINENUM AND T0.BASETYPE=22");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetPRICE2(string DOCENTRY, string LINENUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PRICE FROM ACMESQL02.DBO.POR1 T0");
            sb.Append("  WHERE T0.DOCENTRY=@DOCENTRY AND T0.LINENUM=@LINENUM");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetARRIVE(string DOCENTRY, string LINENUM)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT arriveDay DAY1,substring(arriveDay,1,4)+'/'+substring(arriveDay,5,2)+'/'+substring(arriveDay,7,2) DAY2  FROM lcInstro1 T0 ");
            sb.Append(" LEFT JOIN SHIPPING_MAIN T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetBU(string aa)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT PARAM_DESC as DataText FROM RMA_PARAMS where param_kind=@aa  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GetORDRDATE(string DOCENTRY, string LINENUM)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  select CASE U_ACME_WORKDAY WHEN '內銷' THEN  CONVERT(VARCHAR(10) ,u_ACME_SHIPDAY, 111 )  WHEN '進口轉內銷' THEN  CONVERT(VARCHAR(10) ,u_ACME_SHIPDAY, 111 ) ELSE CONVERT(VARCHAR(10) ,GETDATE(), 111 ) END DATE FROM RDR1    WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

   

        private void button14_Click(object sender, EventArgs e)
        {
            if (boatNameTextBox.Text == "")
            {
                MessageBox.Show("請輸入OBU客戶資訊");
            }
            else
            {
                System.Data.DataTable dt1 = GetMenu.Getaddress(boatNameTextBox.Text);
                DataRow drw = dt1.Rows[0];
                oBUBillToTextBox.Text = drw["cardname"].ToString() + "\r\n" + drw["address"].ToString() + "\r\n" + "TEL:" + drw["phone1"].ToString() + "\r\n" + "FAX:" + drw["fax"].ToString() + "\r\n" + "ATTN:" + drw["cntctprsn"].ToString();
 
            }
        }

        private void wH_Item4DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
         
            }
        }

        private void wH_ItemDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
            
            }
        }

        private void wH_Item2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

            }
        }

        private void wH_Item3DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
              
            }
        }

        private void wH_Car2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
         
            }
        }


        private void ViewBatchPayment()
        {
            if (globals.DBNAME == "CHOICE")
            {
                bbs = "CHOICE CHANNEL CO.,LTD. ";
            }
            else if (globals.DBNAME == "INFINITE")
            {
                bbs = "Infinite Power Group Inc.";
            }
            else if (globals.DBNAME == "TOP GARDEN")
            {
                bbs = "TOP GARDEN INT'L LTD";
            }
         
            else if (globals.DBNAME == "達睿生")
            {
                bbs = "达睿生科技发展（深圳）有限公司";
            }
            else if (globals.DBNAME == "宇豐")
            {
                bbs = "宇豐光電股份有限公司";
            }
            else if (globals.DBNAME == "禾中")
            {
                bbs = "GeTogether Technology Co., Limited";
            }
            else 
            {

                    bbs = "進金生實業股份有限公司";
            }

        }


        private void ViewDATE()
        {
            //if (forecastDayTextBox.Text.Trim() == "銷售訂單")
            //{
            //    DataGridViewRow row = wH_ItemDataGridView.Rows[0];

            //    string dd = row.Cells["dataGridViewTextBoxColumn55"].Value.ToString();
            //    System.Data.DataTable N1 = GetORDRDATE(pINOTextBox.Text, dd);

            //    if (N1.Rows.Count > 0)
            //    {
            //        DataRow drw3 = N1.Rows[0];
            //        DATE1 = drw3["DATE"].ToString();
            //    }
            //    else
            //    {
            //        DATE1 = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            //    }

            //}
            //else
            //{
            //    DATE1 = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            //}
            DATE1 = quantityTextBox.Text;
            DATE2 = DATE1.Replace("/", "");
        }
      

        private void wH_ItemDataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            //GetMenu.InsertLog(fmLogin.LoginID.ToString(), "wH_ItemDataGridView_UserDeletingRow", "單號" + shippingCodeTextBox.Text, DateTime.Now.ToString("yyyyMMddHHmmss"));
               
                    if (!(e.Row.IsNewRow))
                    {
                        if (add4TextBox.Text != "")
                        {
                            if (MessageBox.Show("已發備貨您確定要刪除？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            {

                                e.Cancel = true;
                            }
                            else
                            {

                              //  GetMenu.InsertLog(fmLogin.LoginID.ToString(), "EVA1", shippingCodeTextBox.Text, DateTime.Now.ToString("yyyyMMddHHmmss"));
                            }
                        }
                     
                    }
        }

        private void wH_Item2DataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (!(e.Row.IsNewRow))
            {
                if (add5TextBox.Text != "")
                {
                    if (MessageBox.Show("已發放貨您確定要刪除？", "信息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {

                        e.Cancel = true;
                    }
                    else
                    {
                       // GetMenu.InsertLog(fmLogin.LoginID.ToString(), "EVA2", shippingCodeTextBox.Text, DateTime.Now.ToString("yyyyMMddHHmmss"));

                    }
                }

            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            
            DELETEFILE();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\wh\\裝箱明細.xlsx";
      
            System.Data.DataTable OrderData = GetPACK();


            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  "裝箱明細 "+cardNameTextBox.Text.Trim()+".xlsx";
            //
            //產生 Excel Report
            ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            forecastDayTextBox.Text = comboBox1.Text;

  
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            shipping_OBUTextBox.Text = comboBox2.Text;
       
        }



        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = null;
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" ||  globals.DBNAME == "測試區98" )
            {

                dt3 = GetBU("whs");

            }
            else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
            {
                dt3 = GetBU("wcho1");

            }
            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = null;
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生"  || globals.DBNAME == "測試區98")
            {
                dt4 = GetMenu.Getwarehouse();
            }
            else if ( globals.DBNAME == "宇豐")
            {
                dt4 = GetMenu.GetwarehouseAD();
            }
            else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "禾中")
            {
                dt4 = GetMenu.GetwarehouseCHI();
            }

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][1]));
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {

            System.Data.DataTable T1 = GetCARTON();
            if (T1.Rows.Count > 0)
            {
                dataGridView1.DataSource = T1;
                ExcelReport.GridViewToExcel(dataGridView1);
            }

            System.Data.DataTable T2 = GetWHLIST2();
            if (T2.Rows.Count > 0)
            {
                dataGridView1.DataSource = T2;
                ExcelReport.GridViewToExcel(dataGridView1);
            }
       


        }
        public static bool IsNumber(string strNumber)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"^\d+(\.)?\d*$");
            return r.IsMatch(strNumber);
        }

        public  System.Data.DataTable GetCARTON()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2) 日期,SUM(CAST(ISNULL(CASE shipToDate WHEN '' THEN 0 ELSE shipToDate END,0) AS DECIMAL)) 進貨箱數,SUM(CAST(ISNULL(CASE receiveDay WHEN '' THEN 0 ELSE receiveDay END,0) AS DECIMAL)) 出貨箱數 FROM WH_MAIN ");
            sb.Append(" WHERE  (ISNULL(receiveDay,0) <> 0 OR ISNULL(shipToDate,0) <> 0)");
            sb.Append(" AND SUBSTRING(SHIPPINGCODE,3,4)=@YEAR AND SUBSTRING(SHIPPINGCODE,7,2)=@MONTH ");
            sb.Append(" GROUP BY  SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2)");
            sb.Append(" ORDER BY  SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR",comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", Convert.ToInt16(comboBox4.Text)));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetCARTON2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2) 日期,SUM(CAST(ISNULL(CASE shipToDate WHEN '' THEN 0 ELSE shipToDate END,0) AS DECIMAL)) 進貨箱數,SUM(CAST(ISNULL(CASE receiveDay WHEN '' THEN 0 ELSE receiveDay END,0) AS DECIMAL)) 出貨箱數 FROM WH_MAIN ");
            sb.Append(" WHERE  (ISNULL(receiveDay,0) <> 0 OR ISNULL(shipToDate,0) <> 0)");
            sb.Append(" AND SUBSTRING(SHIPPINGCODE,3,4)=@YEAR ");
            sb.Append(" GROUP BY  SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2)");
            sb.Append(" ORDER BY  SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR", comboBox3.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public void UPDATECHECK()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE WH_MAIN SET CACME='Unchecked' WHERE ISNULL(CACME,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CCHOICE='Unchecked' WHERE ISNULL(CCHOICE,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CIPGI='Unchecked' WHERE ISNULL(CIPGI,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CTOP='Unchecked' WHERE ISNULL(CTOP,'')=''");
            sb.Append(" UPDATE WH_MAIN SET C1FUN='Unchecked' WHERE ISNULL(C1FUN,'')=''");
            sb.Append(" UPDATE WH_MAIN SET C2FUN='Unchecked' WHERE ISNULL(C2FUN,'')=''");
            sb.Append(" UPDATE WH_MAIN SET C3FUN='Unchecked' WHERE ISNULL(C3FUN,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CKBOU='Unchecked' WHERE ISNULL(CKBOU,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CBKBOU='Unchecked' WHERE ISNULL(CBKBOU,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CE='Unchecked' WHERE ISNULL(CE,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CCE='Unchecked' WHERE ISNULL(CCE,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CTKWG='Unchecked' WHERE ISNULL(CTKWG,'')=''");
            sb.Append(" UPDATE WH_MAIN SET CRUZAN='Unchecked' WHERE ISNULL(CRUZAN,'')=''");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void INSERTDOWNLOAD2(string shippingcode, string seq, string filename, string path, string STATUS)
        {
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand(" Insert into Download2(shippingcode,seq,filename,path,STATUS) values(@shippingcode,@seq,@filename,@path,@STATUS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@STATUS", STATUS));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private void INSERTA(string ShippingCode, string SeqNo, string Docentry, string linenum, string ItemRemark, string ItemCode, string Dscription, string Quantity
         , string Remark, string INV, string PiNo, string NowQty, string Ver, string Grade, string Invoice, string FrgnName, string WHName
         , string ShipDate, string CardCode, string U_PAY, string U_SHIPDAY, string U_SHIPSTATUS, string U_MARK, string U_MEMO
         , string PO, string LOCATION, string TREETYPE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO  [dbo].[WH_Item4]");
            sb.Append("            ([ShippingCode],[SeqNo],[Docentry],[linenum],[ItemRemark],[ItemCode],[Dscription],[Quantity]");
            sb.Append("            ,[Remark],[INV],[PiNo],[NowQty],[Ver],[Grade],[Invoice],[FrgnName],[WHName]");
            sb.Append("            ,[ShipDate],[CardCode],[U_PAY],[U_SHIPDAY],[U_SHIPSTATUS],[U_MARK],[U_MEMO]");
            sb.Append("            ,[PO],[LOCATION],[TREETYPE])");
            sb.Append("      VALUES (@ShippingCode,@SeqNo,@Docentry,@linenum,@ItemRemark,@ItemCode,@Dscription,@Quantity");
            sb.Append("            ,@Remark,@INV,@PiNo,@NowQty,@Ver,@Grade,@Invoice,@FrgnName,@WHName");
            sb.Append("            ,@ShipDate,@CardCode,@U_PAY,@U_SHIPDAY,@U_SHIPSTATUS,@U_MARK,@U_MEMO");
            sb.Append("            ,@PO,@LOCATION,@TREETYPE)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@Dscription", Dscription));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
            command.Parameters.Add(new SqlParameter("@INV", INV));
            command.Parameters.Add(new SqlParameter("@PiNo", PiNo));
            command.Parameters.Add(new SqlParameter("@NowQty", NowQty));
            command.Parameters.Add(new SqlParameter("@Ver", Ver));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@Invoice", Invoice));
            command.Parameters.Add(new SqlParameter("@FrgnName", FrgnName));
            command.Parameters.Add(new SqlParameter("@WHName", WHName));
            command.Parameters.Add(new SqlParameter("@ShipDate", ShipDate));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@U_PAY", U_PAY));
            command.Parameters.Add(new SqlParameter("@U_SHIPDAY", U_SHIPDAY));
            command.Parameters.Add(new SqlParameter("@U_SHIPSTATUS", U_SHIPSTATUS));
            command.Parameters.Add(new SqlParameter("@U_MARK", U_MARK));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@TREETYPE", TREETYPE));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }

        private void DELETETA()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE  [dbo].[WH_Item4] WHERE SHIPPINGCODE=@SHIPPINGCODE ");
           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode",shippingCodeTextBox.Text));
          

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
        public System.Data.DataTable GetWHLIST()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select T0.SHIPPINGCODE JOBNO,T0.CARDNAME 客戶名稱,SHIPPING_OBU 倉別,T1.SeqNo 序號,T1.Docentry 單號");
            sb.Append(" ,ShipDate 排程日期,ItemRemark 單據總類,ItemCode 產品編號,Dscription 品名規格,T1.PiNo 料號,T1.Quantity 數量");
            sb.Append(" ,T1.Grade 等級,T1.Ver 版本,INV 原廠INVOCE日期,Invoice 原廠INVOCE FROM WH_MAIN T0");
            sb.Append(" LEFT JOIN  WH_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE  SUBSTRING(T0.SHIPPINGCODE,3,4)=@YEAR AND SUBSTRING(T0.SHIPPINGCODE,7,2)=@MONTH ORDER BY T0.SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", Convert.ToInt16(comboBox4.Text)));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetWHLIST2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             select T0.SHIPPINGCODE JOBNO,T0.CARDNAME 客戶名稱,SHIPPING_OBU 倉別,T1.SeqNo 序號,T1.Docentry 單號");
            sb.Append("             ,ShipDate 排程日期,ItemRemark 單據總類,ItemCode 產品編號,Dscription 品名規格,T1.PiNo 料號,T1.Quantity 數量");
            sb.Append("             ,T1.Grade 等級,T1.Ver 版本,INV 原廠INVOCE日期,Invoice 原廠INVOCE FROM WH_MAIN T0");
            sb.Append("             LEFT JOIN  WH_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("   WHERE              SUBSTRING(T0.SHIPPINGCODE,3,4)=@YEAR AND SUBSTRING(T0.SHIPPINGCODE,7,2)=@MONTH  AND (ISNULL(receiveDay,0) <> 0 OR ISNULL(shipToDate,0) <> 0)");
            sb.Append("  AND SUBSTRING(T0.SHIPPINGCODE,3,8)  IN (");
            sb.Append("     SELECT SUBSTRING(T0.SHIPPINGCODE,3,8) 日期 FROM WH_MAIN ");
            sb.Append("             WHERE  (ISNULL(receiveDay,0) <> 0 OR ISNULL(shipToDate,0) <> 0)");
            sb.Append("             AND SUBSTRING(SHIPPINGCODE,3,4)=@YEAR AND SUBSTRING(SHIPPINGCODE,7,2)=@MONTH ");
            sb.Append("             GROUP BY  SUBSTRING(SHIPPINGCODE,3,4)+'/'+SUBSTRING(SHIPPINGCODE,7,2)+'/'+SUBSTRING(SHIPPINGCODE,9,2)");
            sb.Append("                 HAVING SUM(CAST(ISNULL(CASE shipToDate WHEN '' THEN 0 ELSE shipToDate END,0) AS DECIMAL)) >= 10 OR  SUM(CAST(ISNULL(CASE receiveDay WHEN '' THEN 0 ELSE receiveDay END,0) AS DECIMAL)) >= 10) ORDER BY T0.SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", Convert.ToInt16(comboBox4.Text)));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOP(string DOCENTRY, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                 select  t0.shippingcode CODE,T0.Remark CUST,T2.U_ACME_INV INV,T0.ITEMCODE,SUM(T0.QUANTITY) QTY from shipping_item t0  ");
            sb.Append("                       LEFT JOIN ACMESQL02.DBO.OPDN T2 ON (T0.SHIPPINGCODE=T2.U_SHIPPING_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("                          where t0.itemremark IN ('採購訂單','採購報價')  and cast(t2.docentry as varchar)=@DOCENTRY  AND T0.ITEMCODE=@ITEMCODE ");
            sb.Append("                    GROUP BY t0.shippingcode ,T0.Remark ,T2.U_ACME_INV,T0.ITEMCODE   HAVING SUM(T0.QUANTITY)=@QTY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPF2(string DOCENTRY, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                 select  t0.shippingcode CODE,T2.U_ACME_INV INV,T0.ITEMCODE,SUM(T0.QUANTITY) QTY from shipping_item t0  ");
            sb.Append("                       LEFT JOIN ACMESQL02.DBO.OPDN T2 ON (T0.SHIPPINGCODE=T2.U_SHIPPING_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("                          where t0.itemremark IN ('採購訂單','採購報價')  and cast(t2.docentry as varchar)=@DOCENTRY  AND T0.ITEMCODE=@ITEMCODE ");
            sb.Append("                    GROUP BY t0.shippingcode ,T2.U_ACME_INV,T0.ITEMCODE   HAVING SUM(T0.QUANTITY)=@QTY ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPF4(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT SUM(CAST(QUANTITY AS INT)) QTY FROM WH_ITEM4 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE AND ISNULL(Invoice,'') = '' ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

       
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPF3(string DOCENTRY, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  T0.Remark CUST from shipping_item t0    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OPDN T2 ON (T0.SHIPPINGCODE=T2.U_SHIPPING_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS)   ");
            sb.Append(" where t0.itemremark IN ('採購訂單','採購報價')  and cast(t2.docentry as varchar)=@DOCENTRY  AND T0.ITEMCODE=@ITEMCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPF(string DOCENTRY, string ITEMCODE, string QTY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     select  t0.shippingcode CODE,T0.Remark CUST,T2.U_ACME_INV INV,T0.ITEMCODE,(T0.QUANTITY) QTY from shipping_item t0   ");
            sb.Append("                                     LEFT JOIN ACMESQL02.DBO.OPDN T2 ON (T0.SHIPPINGCODE=T2.U_SHIPPING_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                                        where t0.itemremark  IN ('採購訂單','採購報價') and cast(t2.docentry as varchar)=@DOCENTRY  AND T0.ITEMCODE=@ITEMCODE AND T0.QUANTITY=@QTY  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetIP()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                      SELECT  '日期:'+Convert(varchar(10),GETDATE(),126)  日期,T0.CARDNAME 公司名稱  ");
            sb.Append("                             ,REPLACE(T2.U_MODEL,'V','') MODEL,SUM(CAST(T1.QUANTITY AS INT)) 數量,T1.BOXCHECK INVOICE,U_LOCATION LOC, ");
            sb.Append("               CASE WHEN SUBSTRING(T1.ITEMCODE,1,1)='O' THEN '液晶顯示板(不含背光模組) ('+T2.U_SIZE+'吋)'");
            sb.Append(" WHEN T2.U_GROUP='117-T-Con Board' THEN '同步控制板'");
            sb.Append("  ELSE '液晶顯示屏(含背光模組) ('+T2.U_SIZE+'吋)' END ITEMNAME ,");
            sb.Append(" CASE WHEN T2.U_GROUP='117-T-Con Board' THEN '''9013902000'");
            sb.Append("            WHEN T2.U_GROUP IN ('100-Panel','180-Open cell') THEN ");
            sb.Append(" CASE WHEN CAST(T2.U_SIZE AS DECIMAL) <=10.1 THEN '''9013803010'   ");
            sb.Append("  WHEN CAST(T2.U_SIZE AS DECIMAL)  >10.1 AND CAST(T2.U_SIZE AS DECIMAL) <=32 THEN '''9013803020'  ELSE '''9013803090' END");
            sb.Append("        END 稅則 ,CASE WHEN ITEMREMARK='收貨採購單' THEN T3.PRICE ELSE T4.PRICE END PRICE FROM WH_MAIN T0   ");
            sb.Append("                             LEFT JOIN WH_ITEM3 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append("                             LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                           LEFT JOIN ( SELECT T0.DOCENTRY,T0.LINENUM,T1.PRICE FROM ACMESQL02.DBO.PDN1 T0  ");
            sb.Append("                             LEFT JOIN ACMESQL02.DBO.POR1 T1 ON (T0.BASEENTRY=T1.DOCENTRY AND T0.BASELINE=T1.LINENUM)  ");
            sb.Append("                              WHERE  T0.BASETYPE=22 ) T3 ON (T3.DOCENTRY=T1.DOCENTRY AND T3.LINENUM=T1.LINENUM) ");
            sb.Append("               LEFT JOIN (SELECT T0.DOCENTRY,T0.LINENUM,PRICE FROM ACMESQL02.DBO.POR1 T0 ) T4  ON (T4.DOCENTRY=T1.DOCENTRY AND T4.LINENUM=T1.LINENUM) ");
            sb.Append("                            WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append("               GROUP BY T0.CARDNAME ,REPLACE(T2.U_MODEL,'V',''),T1.BOXCHECK,U_LOCATION, ");
            sb.Append(" CASE WHEN SUBSTRING(T1.ITEMCODE,1,1)='O' THEN '液晶顯示板(不含背光模組) ('+T2.U_SIZE+'吋)'");
            sb.Append(" WHEN T2.U_GROUP='117-T-Con Board' THEN '同步控制板'");
            sb.Append("  ELSE '液晶顯示屏(含背光模組) ('+T2.U_SIZE+'吋)' END");
            sb.Append("                     ,CASE WHEN ITEMREMARK='收貨採購單' THEN T3.PRICE ELSE T4.PRICE END,");
            sb.Append(" CASE WHEN T2.U_GROUP='117-T-Con Board' THEN '''9013902000'");
            sb.Append("            WHEN T2.U_GROUP IN ('100-Panel','180-Open cell') THEN ");
            sb.Append(" CASE WHEN CAST(T2.U_SIZE AS DECIMAL) <=10.1 THEN '''9013803010'   ");
            sb.Append("  WHEN CAST(T2.U_SIZE AS DECIMAL)  >10.1 AND CAST(T2.U_SIZE AS DECIMAL) <=32 THEN '''9013803020'  ELSE '''9013803090' END");
            sb.Append("        END ");

            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPOR(string DOCENTRY, string DSCRIPTION)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("   select t0.shippingcode CODE,T0.Remark CUST,T2.U_ACME_INV INV,T0.SEQNO,T0.ITEMCODE from shipping_item t0 ");
            sb.Append("         LEFT JOIN ACMESQL02.DBO.OPDN T2 ON (T0.SHIPPINGCODE=T2.U_SHIPPING_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("         LEFT JOIN ACMESQL02.DBO.PDN1 T3 ON (T2.DOCENTRY=T3.DOCENTRY)");
            sb.Append("            where t0.itemremark IN ('採購訂單','採購報價')  and cast(t2.docentry as varchar)=@DOCENTRY ");
            sb.Append(" AND T0.DSCRIPTION=@DSCRIPTION ");
           
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@DSCRIPTION", DSCRIPTION));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOPOR3(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT U_ACME_INV FROM OPDN WHERE DOCENTRY=@DOCENTRY ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOITM(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT * FROM OITM WHERE ITMSGRPCOD='102' AND ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
           
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetSHIPOP2(string DOC, string MODEL,string VER)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               select t1.shippingcode CODE,Remark CUS,T0.LINENUM,T0.ITEMCODE from shipping_item t0  ");
            sb.Append("                left join shipping_main t1 on (t0.shippingcode=t1.shippingcode) ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T0.ITEMCODE=T2.ITEMCODE  COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where t0.itemremark='採購訂單' and cast(t0.docentry as varchar)=@DOC   ");
            sb.Append("               AND T2.U_TMODEL =@MODEL and T2.U_VERSION=@VER ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOC", DOC));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOPORCUS(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_ACME_CUS FROM OPOR WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void button20_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetCARTON2();
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void wH_ItemDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (globals.DBNAME != "宇豐" && globals.DBNAME != "禾中")
                {
                    if (wH_ItemDataGridView.Columns[e.ColumnIndex].Name == "Invoice")
                    {
                        string Invoice = Convert.ToString(this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["Invoice"].Value);

                        System.Data.DataTable T1 = GetOPDN(Invoice);
                        if (T1.Rows.Count == 0 && Invoice.Length > 10)
                        {
                            string INV = Invoice.Substring(0, 10);
                            T1 = GetOPDN(INV);
                        }
                        if (T1.Rows.Count > 0)
                        {
                            this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["LOCATION"].Value = T1.Rows[0][0].ToString();
                            this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["INV"].Value = T1.Rows[0][1].ToString();
                        }

                    }
                }
                if (wH_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemCode2" || wH_ItemDataGridView.Columns[e.ColumnIndex].Name == "Quantity")
                {
                    string ITEMCODE = Convert.ToString(this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["ItemCode2"].Value);
                    int QTY = Convert.ToInt32(this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    System.Data.DataTable GE1 = GETPACK(ITEMCODE);
                    if (GE1.Rows.Count > 0)
                    {

                        int FF1 = Convert.ToInt32(GE1.Rows[0][0]);
                        int FF2 = Convert.ToInt32(GE1.Rows[0][1]);

                        int mod2 = QTY % FF2;
                        int mod3 = QTY / FF2;
                        if (mod2 > 0)
                        {
                            mod3 = mod3 + 1;
                        }

                        this.wH_ItemDataGridView.Rows[e.RowIndex].Cells["LPRINT"].Value = mod3.ToString();
     
                    }

                }
            }
            catch { }
   
        }


        public System.Data.DataTable GetOPDN(string INV)
        {
            SqlConnection MyConnection = new SqlConnection(strCn02);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_LOCATION,Convert(varchar(10),U_ACME_Invoice,112) INVDATE   FROM OPDN  WHERE U_ACME_INV LIKE '%" + INV + "%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetDI(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item3 where shippingcode=@shippingcode  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetDIM(string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item3 WHERE DOCENTRY IN ( " + cs + ")  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        //            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item3 WHERE DOCENTRY IN ( " + cs + ")  ");
        public System.Data.DataTable GetDIB(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item where shippingcode=@shippingcode  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetDIBM2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item where shippingcode=@shippingcode  AND ITEMCODE='ACMERMA01.RMA01' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetDIBM(string cs)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select ITEMCODE,Quantity QTY,WhName WHNAME  from wH_Item WHERE DOCENTRY1 IN ( " + cs + ")  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetDI2(string WhsName)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select WHSCODE   from ACMESQL02.DBO.OWHS  WHERE STREET =@WhsName ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WhsName", WhsName));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetDI3(string WhsName)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @name3 varchar(100) ");
            sb.Append(" select @name3 =SUBSTRING(COALESCE(@name3 + '/',''),0,99) + INVOICENO ");
            sb.Append(" FROM (");
            sb.Append(" select DISTINCT BoxCheck  INVOICENO   from wH_Item3 where shippingcode=@ShippingCode AND ISNULL(BoxCheck,'') <> '' ");
            sb.Append(" ) A");
            sb.Append(" SELECT @name3 INVOICENO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetDI4()
        {
            
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OWTR");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetDI41()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DOCENTRY FROM OWTR WHERE  JRNLMEMO=@ShippingCode");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void button21_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetWHLIST();
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void wH_Item4DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= wH_Item4DataGridView.Rows.Count-1)
                return;
            DataGridViewRow dgr = wH_Item4DataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["TREETYPE"].Value.ToString() == "S")
                {

                    dgr.DefaultCellStyle.ForeColor  = Color.Red;
                }
                if (dgr.Cells["TREETYPE"].Value.ToString() == "I")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Gray;
                }
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 

        }

        private void wH_ItemDataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= wH_ItemDataGridView.Rows.Count - 1)
                return;
            DataGridViewRow dgr = wH_ItemDataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["TREETYPE1"].Value.ToString() == "S")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }
                if (dgr.Cells["TREETYPE1"].Value.ToString() == "I")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Gray;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }  
        }

        private void wH_Item2DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= wH_Item2DataGridView.Rows.Count - 1)
                return;
            DataGridViewRow dgr = wH_Item2DataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["TREETYPE2"].Value.ToString() == "S")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }
                if (dgr.Cells["TREETYPE2"].Value.ToString() == "I")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Gray;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }  
        }

        private void wH_Item3DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= wH_Item3DataGridView.Rows.Count - 1)
                return;
            DataGridViewRow dgr = wH_Item3DataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["TREETYPE3"].Value.ToString() == "S")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }
                if (dgr.Cells["TREETYPE3"].Value.ToString() == "I")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Gray;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            string OPDN = "";
            try
            {

                receiveCardTextBox.Text = cardNameTextBox.Text;
                System.Data.DataTable dt1 = Getwhitem44(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = wh.WH_Item3;
                if (dt1.Rows.Count == 0)
                {
                    MessageBox.Show("來源無資料，請先存檔");

                    tabControl1.SelectedIndex = 0;

                }

                string gj = "1*20'櫃號:1*40'櫃號:板/箱/貨代:和達/大榮/東南亞/友福/DHL/中菲行/驊洲/聯倉/航通--請告知幾板幾箱及INV";

                receiveMemoTextBox.Text = gj;
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {

                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                    string DOC = drw["Docentry"].ToString();

                    drw2["Docentry"] = DOC;
                    drw2["linenum"] = drw["linenum"].ToString();
                    OPDN = DOC;

                    //採購單

                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["Remark"] = drw["Remark"];
                    //DeCust
                    drw2["PiNo"] = drw["PiNo"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["TREETYPE"] = drw["TREETYPE"];


                    drw2["cardcode"] = drw["cardcode"];
                    drw2["LOCATION"] = drw["LOCATION"];
                    drw2["BoxCheck"] = drw["Invoice"];


                    if (forecastDayTextBox.Text == "採購單")
                    {
                        System.Data.DataTable dts = GetSHIPOP2(DOC, drw["MODEL"].ToString(), drw["VERSION"].ToString());
                        if (dts.Rows.Count > 0)
                        {
                            for (int J = 0; J <= dts.Rows.Count - 1; J++)
                            {
                                string A1 = dts.Rows[J][0].ToString();
                                string A2 = dts.Rows[J][1].ToString();
                                string A4 = dts.Rows[J][2].ToString();
                                string A5 = dts.Rows[J][3].ToString();

                                drw2["FrgnName"] = A1;
                                drw2["DeCust"] = A2;
                            }
                        }
                    }
                    else
                    {
                        drw2["FrgnName"] = boardCountTextBox.Text;
                    }

                    dt2.Rows.Add(drw2);
                }

                dollarsKindTextBox.Text = DateTime.Now.ToString("yyyyMMddHHmmss");

                wH_Item3BindingSource.EndEdit();
                this.wH_Item3TableAdapter.Update(wh.WH_Item3);
                wh.WH_Item3.AcceptChanges();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.wH_Item3BindingSource.EndEdit();
        }


        private void button23_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = Getwhitem4LAB(shippingCodeTextBox.Text);
            System.Data.DataTable dt2 = wh.WH_LAB;
            if (dt1.Rows.Count == 0)
            {
                MessageBox.Show("來源無資料，請先存檔");

                tabControl1.SelectedIndex = 0;

            }
            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["SeqNo"] = drw["SEQNO"];
                drw2["Dscription"] = drw["DSCRIPTION"];
                drw2["ItemCode"] = drw["ITEMCODE"];
                drw2["LOCATION"] = drw["U_LOCATION"];
                dt2.Rows.Add(drw2);
            }

             wH_LABBindingSource.EndEdit();
             wH_LABTableAdapter.Update(wh.WH_LAB);
             wh.WH_LAB.AcceptChanges();

        }

        private void button24_Click(object sender, EventArgs e)
        {

            DialogResult result;
            result = MessageBox.Show("請確認是否要列印", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                wH_LABBindingSource.EndEdit();
                wH_LABTableAdapter.Update(wh.WH_LAB);
                wh.WH_LAB.AcceptChanges();

                if (wH_LABDataGridView.Rows.Count > 1)
                {
                    for (int h = 0; h <= wH_LABDataGridView.Rows.Count - 2; h++)
                    {

                        DataGridViewRow row1 = wH_LABDataGridView.Rows[h];
                        string  DESP = Convert.ToString(row1.Cells["LDESP"].Value);
                        string QQ = Convert.ToString(row1.Cells["LQTY"].Value);
                        string QTY = "   ";
                        if (!String.IsNullOrEmpty(QQ))
                        {
                            QTY = "QTY :" + QQ;
                        }
                        string LOCATION = "MADE IN " + Convert.ToString(row1.Cells["LLOCATION"].Value);
                        int  PQTY = Convert.ToInt16(row1.Cells["PQTY"].Value);

                        if (PQTY > 0)
                        {
                            for (int i = 0; i <= PQTY-1; i++)
                            {
                                ADD(DESP, QTY, LOCATION);

                            }

                        }
                    }

                }


            }
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
                    ret = A_CreatePrn(10, "\\\\10.2.2.116\\LabelDr2");// open usb.
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

            int LineLimit = 180;
            int fontSize = 36;
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

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt2 = wh.WH_Item3;
            DataRow newCustomersRow = dt2.NewRow();
            int i = wH_Item3DataGridView.CurrentRow.Index;

            DataRow drw = dt2.Rows[i];
            string sa = drw["shippingcode"].ToString();
            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["Docentry"] = drw["Docentry"];
            newCustomersRow["linenum"] = drw["linenum"];
            newCustomersRow["ItemRemark"] = drw["ItemRemark"];
            newCustomersRow["ItemCode"] = drw["ItemCode"];
            newCustomersRow["Dscription"] = drw["Dscription"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Remark"] = drw["Remark"];
            newCustomersRow["INV"] = drw["INV"];
            newCustomersRow["PiNo"] = drw["PiNo"];
            newCustomersRow["NowQty"] = drw["NowQty"];
            newCustomersRow["Ver"] = drw["Ver"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["Invoice"] = drw["Invoice"];
            newCustomersRow["DeCust"] = drw["DeCust"];
            newCustomersRow["BoxCheck"] = drw["BoxCheck"];
            newCustomersRow["FrgnName"] = drw["FrgnName"];
            newCustomersRow["WhName"] = drw["WhName"];
            newCustomersRow["Shipdate"] = drw["Shipdate"];
            newCustomersRow["CardCode"] = drw["CardCode"];
            newCustomersRow["A1"] = drw["A1"];
            newCustomersRow["A2"] = drw["A2"];
            newCustomersRow["A3"] = drw["A3"];
            newCustomersRow["A4"] = drw["A4"];
            newCustomersRow["A5"] = drw["A5"];
            newCustomersRow["A6"] = drw["A6"];
            newCustomersRow["A7"] = drw["A7"];
            newCustomersRow["LOCATION"] = drw["LOCATION"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, wH_Item3DataGridView.Rows.Count + 1);

                for (int j = 0; j <= wH_Item3DataGridView.Rows.Count - 2; j++)
                {
                    wH_Item3DataGridView.Rows[j].Cells[0].Value = (j + 1).ToString();
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void wH_Item5DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["dataGridViewTextBoxColumn44"].Value = util.GetSeqNo(2, wH_Item5DataGridView);
        }
        public void SHIPNO()
        {

            System.Data.DataTable dt3 = GetSHIP(shippingCodeTextBox.Text);
            if (dt3.Rows.Count > 0)
            {
                if (dt3.Rows[0]["ITEMREMARK"].ToString() == "銷售訂單")
                {

                    StringBuilder sb2 = new StringBuilder();
                    StringBuilder sb3 = new StringBuilder();
                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString();
                        string LINENUM = dt3.Rows[i]["LINENUM"].ToString();
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                    }
                    sb2.Remove(sb2.Length - 1, 1);
                    string A = sb2.ToString();
                    System.Data.DataTable SS = GetSH(A, "銷售訂單");
                    if (SS.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SS.Rows.Count - 1; i++)
                        {
                            string CODE = SS.Rows[i]["CODE"].ToString();

                            sb3.Append(CODE + ",");

                        }
                        sb3.Remove(sb3.Length - 1, 1);
                        JOBNO = sb3.ToString();
                        boardCountTextBox.Text = sb3.ToString();
                    }

                }
            }
        

            System.Data.DataTable dt3DIAO = GetSHIPDIAO(shippingCodeTextBox.Text);
            if (dt3DIAO.Rows.Count > 0)
            {
                StringBuilder sb2 = new StringBuilder();
                StringBuilder sb3 = new StringBuilder();
                for (int i = 0; i <= dt3DIAO.Rows.Count - 1; i++)
                {
                    string DOCENTRY = dt3DIAO.Rows[i]["DOCENTRY"].ToString();
                    string LINENUM = dt3DIAO.Rows[i]["LINENUM"].ToString();
                    sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                }
                sb2.Remove(sb2.Length - 1, 1);
                string A = sb2.ToString();
                System.Data.DataTable SS = GetSH(A, "調撥單");
                if (SS.Rows.Count > 0)
                {
                    for (int i = 0; i <= SS.Rows.Count - 1; i++)
                    {

                        string CODE = SS.Rows[i]["CODE"].ToString();
                        if (quantityTextBox.Text == "")
                        {
                            if (boardCountNoComboBox.Text == "出口" || boardCountNoComboBox.Text == "三角")
                            {
                                System.Data.DataTable K1 = GETARRIVE2(CODE);
                                if (K1.Rows.Count > 0)
                                {
                                    quantityTextBox.Text = K1.Rows[0][0].ToString();
                                }
                            }
                        }
                        sb3.Append(CODE + ",");

                    }
                    sb3.Remove(sb3.Length - 1, 1);
                    JOBNO = sb3.ToString();
                    boardCountTextBox.Text = sb3.ToString();
                }
            }
            if (cardCodeTextBox.Text.Trim() == "S0028")
            {
                System.Data.DataTable dt3TSAIGO = GetSHITSAIGO(shippingCodeTextBox.Text);
                if (dt3TSAIGO.Rows.Count > 0)
                {
                    StringBuilder sb2 = new StringBuilder();
                    StringBuilder sb3 = new StringBuilder();
       
                    for (int i = 0; i <= dt3TSAIGO.Rows.Count - 1; i++)
                    {
                        string DOCENTRY = dt3TSAIGO.Rows[i]["DOCENTRY"].ToString();
                        string LINENUM = dt3TSAIGO.Rows[i]["LINENUM"].ToString();
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                    }
                    sb2.Remove(sb2.Length - 1, 1);
                    string A = sb2.ToString();
                    System.Data.DataTable SS = GetSH(A, "採購訂單");
                    if (SS.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SS.Rows.Count - 1; i++)
                        {
                            string CODE = SS.Rows[i]["CODE"].ToString();

                            sb3.Append(CODE + ",");

                        }
                        sb3.Remove(sb3.Length - 1, 1);
                        JOBNO = sb3.ToString();
                        boardCountTextBox.Text = sb3.ToString();
                    }
                }
            }

            string CRENAME = "";
            //船務jobnoCHOICE
            if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN" || globals.DBNAME == "宇豐" || globals.DBNAME == "禾中")
            {
                if (dt3.Rows.Count > 0)
                {
                    StringBuilder sb2 = new StringBuilder();
                    StringBuilder sb3 = new StringBuilder();
                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString().Trim();
                        string LINENUM = dt3.Rows[i]["LINENUM"].ToString().Trim();
                        sb2.Append("'" + DOCENTRY + ' ' + LINENUM + "',");

                    }
                    sb2.Remove(sb2.Length - 1, 1);
                    string A = sb2.ToString();
                    System.Data.DataTable SS = null;

                    if (globals.DBNAME == "CHOICE")
                    {
                        SS = GetSH2(A, "Choice");
                        CRENAME = "CHO";
                    }
                    if (globals.DBNAME == "INFINITE")
                    {
                        SS = GetSH2(A, "Infinite");
                        CRENAME = "INF";
                    }
                    if (globals.DBNAME == "TOP GARDEN")
                    {
                        SS = GetSH2(A, "TOP GARDEN");
                        CRENAME = "TOP";
                    }
                    if (globals.DBNAME == "禾中")
                    {
                        SS = GetSH2(A, "禾中");
                        CRENAME = "禾中";
                    }
                    if (globals.DBNAME == "宇豐")
                    {
                        SS = GetSH2(A, "宇豐");
                    }
                    if (SS.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SS.Rows.Count - 1; i++)
                        {
                            string CODE = SS.Rows[i]["CODE"].ToString();

                            sb3.Append(CODE + ",");

                        }
                        sb3.Remove(sb3.Length - 1, 1);
                        JOBNO = sb3.ToString();
                        boardCountTextBox.Text = sb3.ToString();
                    }
                }
            }
            if (String.IsNullOrEmpty(boardCountTextBox.Text))
            {
                StringBuilder sb3 = new StringBuilder();
                System.Data.DataTable SS = GetSHIPF(shippingCodeTextBox.Text);
                if (SS.Rows.Count > 0)
                {
                    for (int i = 0; i <= SS.Rows.Count - 1; i++)
                    {
                        string CODE = SS.Rows[i][0].ToString();

                        sb3.Append(CODE + ",");

                    }
                    sb3.Remove(sb3.Length - 1, 1);
                    JOBNO = sb3.ToString();
                    boardCountTextBox.Text = sb3.ToString();
                }
            
            }

            wH_mainBindingSource.EndEdit();
            wH_mainTableAdapter.Update(wh.WH_main);
            wh.WH_main.AcceptChanges();
        
        }
        private System.Data.DataTable GETARRIVE2(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT SUBSTRING(ARRIVEDAY,1,4)+'/'+SUBSTRING(ARRIVEDAY,5,2)+'/'+SUBSTRING(ARRIVEDAY,7,2) DDATE FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private void wH_Item5DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "colEdit")
                {
                    wH_Item5BindingSource.EndEdit();
                    wH_Item5TableAdapter.Update(wh.WH_Item5);

                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;
                    if (row != null)
                    {
                        string ID = Convert.ToString(row["ID"]);
                        WHOHEM form = new WHOHEM(ID);
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            wH_Item5TableAdapter.Fill(wh.WH_Item5, MyID);
                            try
                            {
                                (sender as DataGridView).CurrentCell = (sender as DataGridView)[0, scrollPosition];
                            }
                            catch
                            {

                            }
                        }

                    }
                }

            }
        }

     



        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                DELETEFILE();

                SHIPNO();
                string SEMAIL  = "";
                string CEMAIL = "";
                DialogResult result;
                result = MessageBox.Show("請確認是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                  
                    System.Data.DataTable GETMAIL = GetMenu.GetWHNAIL(shipping_OBUTextBox.Text);
                    if (GETMAIL.Rows.Count > 0)
                    {
                        SEMAIL = GETMAIL.Rows[0]["SEMAIL"].ToString();
                        CEMAIL = GETMAIL.Rows[0]["CEMAIL"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("沒有收貨EMAIL");
                        return;
                    }
                    button80("N");

                        string template;
                        StreamReader objReader;
                        string FileName = string.Empty;
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                        FileName = lsAppDir + "\\MailTemplates\\WHMAIN.htm";
                        objReader = new StreamReader(FileName);

                        template = objReader.ReadToEnd();
                        objReader.Close();
                        objReader.Dispose();



                        StringWriter writer = new StringWriter();
                        HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);



                        string h = fmLogin.LoginID.ToString();
                        template = template.Replace("##Content1##", "請依收貨工單對照型號/等級/數量/產地/版本，並將ACME１５碼打進庫存表 & 提供進貨序號以利核對 , 謝謝。");
                        template = template.Replace("##Content2##", " P.S. ");
                        template = template.Replace("##Content3##", "1.請務必於點收進貨確認無誤後將收貨工單簽名回覆");
                        template = template.Replace("##Content4##", "2.進貨如有異常請於回傳收貨工單時寫清楚INV#/異常幾板／幾箱");
                        template = template.Replace("##Content5##", "3.貨代送貨到時,請當場對點完於簽收單上簽名,如有異常請載明異常點,以利確認責任歸屬!!!");
                        try
                        {
                            System.Data.DataTable dt1 = GetMenu.Getemployee(h);
                            DataRow drw = dt1.Rows[0];
                            if ((dt1.Rows.Count) > 0)
                            {
                                string a1 = drw["pager"].ToString();
                                string a2 = drw["mobile"].ToString();
                                template = template.Replace("##eng##", a2);
                                template = template.Replace("##name##", a1);
                                template = template.Replace("##mail##", h + "@acmepoint.com");
                            }
                        }
                        catch
                        {
                            template = template.Replace("##eng##", "");
                            template = template.Replace("##name##", "");
                            template = template.Replace("##mail##", h + "@acmepoint.com");
                        }

                        MailMessage message = new MailMessage();
                        if (globals.GroupID.ToString().Trim() == "EEP")
                        {
                            message.To.Add("LLEYTONCHEN@ACMEPOINT.COM");
                        }
                        else
                        {
                            string[] arrurl = SEMAIL.Replace("\r", "").Replace("\n", "").Split(new Char[] { ',' });

                            foreach (string i in arrurl)
                            {
                                if (!String.IsNullOrEmpty(i))
                                {
                                    message.To.Add(i);
                                }
                            }

                            string[] arrurl2 = CEMAIL.Replace("\r", "").Replace("\n", "").Split(new Char[] { ',' });

                            foreach (string i in arrurl2)
                            {
                                if (!String.IsNullOrEmpty(i))
                                {

                                    message.CC.Add(i);

                                }
                            }
                        }
                        message.Subject = MAILSUB;
                        message.Body = template;

                        string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                        string[] filenames = Directory.GetFiles(OutPutFile);
                        foreach (string file in filenames)
                        {

                            string m_File = "";

                            m_File = file;
                            data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);
                            ContentDisposition disposition = data.ContentDisposition;
                            message.Attachments.Add(data);

                        }
                        if (shipping_OBUTextBox.Text == "新得利倉")
                        {
                            System.Data.DataTable dt2 = GetAUINV(shippingCodeTextBox.Text);
                            if (dt2.Rows.Count > 0)
                            {
                                for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                                {

                                    DataRow dd = dt2.Rows[i];


                                    string INV = dd["INV"].ToString();
                                    string DIR = "//acmesrv01//Public//採購進貨_出貨文件//PdfBak";

                                    string[] filenamesD = Directory.GetFiles(DIR);
                                    foreach (string file in filenamesD)
                                    {
                                        FileInfo info = new FileInfo(file);
                                        string NAME = info.Name.ToString().Trim().Replace(" ", "");


                                        if (NAME != "Thumbs.db")
                                        {
                                            int J1 = NAME.LastIndexOf(".");
                                            string M2 = NAME.Substring(0, J1).ToUpper().Replace("SHIPDOC", "").Replace("_", "").Replace("PK", "");

                                            int G2 = M2.IndexOf(INV);
                                            if (G2 == -1)
                                            {
                                                 G2 = INV.IndexOf(M2);
                                            }
                                            int G3 = NAME.IndexOf("Pk");

                                            if (G2 != -1)
                                            {
                                                if (G3 != -1)
                                                {
                                                    //result = MessageBox.Show("請確認是否要夾附件 " + NAME, "YES/NO", MessageBoxButtons.YesNo);
                                                    //if (result == DialogResult.Yes)
                                                    //{

                                                    data = new System.Net.Mail.Attachment(file, MediaTypeNames.Application.Octet);
                                                    ContentDisposition disposition = data.ContentDisposition;
                                                    message.Attachments.Add(data);
                                                    // }
                                                }
                                            }

                                        }
                                    }

                                }
                            }

                        }
                        message.IsBodyHtml = true;

                        SmtpClient client = new SmtpClient();
                         client.Send(message);
                        data.Dispose();
                        message.Attachments.Dispose();

                        DELETEFILE();
                        MessageBox.Show("寄信成功");
                    
                }
            }
            catch (Exception ex)
            {
                DELETEFILE();
                MessageBox.Show(ex.Message);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {

            try
            {

                receiveCardTextBox.Text = cardNameTextBox.Text;
                System.Data.DataTable dt1 = GetALL(textBox8.Text);
                System.Data.DataTable dt2 = wh.WH_Item;
                if (dt1.Rows.Count == 0)
                {
                    MessageBox.Show("來源無資料，請先存檔");

                    tabControl1.SelectedIndex = 0;

                }


                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {

                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    boardDeliverTextBox.Text = drw["boardDeliver"].ToString();
                    add1TextBox.Text = drw["add1"].ToString();
                    iNVOICENOTextBox.Text = drw["iNVOICENO"].ToString();
                    pACKMEMOTextBox.Text = drw["pACKMEMO"].ToString();

                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];

                    drw2["Docentry"] = drw["Docentry"] ;
                    drw2["linenum"] = drw["linenum"];
                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["Quantity"] = drw["QTY"];
                    drw2["Remark"] = drw["Remark"];
                    drw2["INV"] = drw["INV"];
                    drw2["PiNo"] = drw["AUNO"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["Invoice"] = drw["Invoice"];
                    drw2["FrgnName"] = drw["FrgnName"];
                    drw2["Shipdate"] = drw["Shipdate"];
                    drw2["cardcode"] = drw["UNIT"];


                    drw2["U_PAY"] = drw["U_PAY"];
                    drw2["U_SHIPDAY"] = drw["U_SHIPDAY"];
                    drw2["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
                    drw2["U_MARK"] = drw["U_MARK"];

                    drw2["U_MEMO"] = drw["U_MEMO"].ToString(); ;
                    drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                    drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                    drw2["PO"] = drw["PO"];
                    drw2["TREETYPE"] = drw["TREETYPE"];
              

                    drw2["FrgnName1"] = drw["FrgnName"];
                    drw2["LOCATION"] = drw["LOCATION"];

                    dt2.Rows.Add(drw2);
                }

              

                wH_ItemBindingSource.EndEdit();
                this.wH_ItemTableAdapter.Update(wh.WH_Item);
                wh.WH_Item.AcceptChanges();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.wH_Item3BindingSource.EndEdit();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            WHPACK();
        }

        private void pINOTextBox_TextChanged(object sender, EventArgs e)
        {
        
            AUTOCARD();
        }

        private void AUTOCARD()
        {
            try
            {
                string USER = fmLogin.LoginID.ToString();
                string F1 = add4TextBox.Text;
                if (USER.ToUpper() != "JOYCHEN")
                {
                    if (globals.DBNAME != "宇豐")
                    {
                        System.Data.DataTable G1 = null;


                        if (forecastDayTextBox.Text == "銷售訂單" || forecastDayTextBox.Text == "銷售單")
                        {
                            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                            {
                                G1 = GERCARD(pINOTextBox.Text.Trim());
                                if (G1.Rows.Count > 0)
                                {
                                    cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                                    cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                                }
                            }

                            if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE")
                            {
                                G1 = GERCARD2(pINOTextBox.Text.Trim());
                                if (G1.Rows.Count > 0)
                                {
                                    cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                                    cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                                }
                            }
                        }

                        if (forecastDayTextBox.Text == "庫存調撥-借出")
                        {
                            if (globals.DBNAME == "達睿生")
                            {
                                G1 = GERCARD3(pINOTextBox.Text.Trim());
                                if (G1.Rows.Count > 0)
                                {
                                    cardCodeTextBox.Text = G1.Rows[0]["CARDCODE"].ToString();
                                    cardNameTextBox.Text = G1.Rows[0]["CARDNAME"].ToString();
                                }
                            }

                        }

                       
                      
                        

                        
                      
                    }

                }
            }
            catch { }
        }
        

        private void forecastDayTextBox_TextChanged(object sender, EventArgs e)
        {
            AUTOCARD();
        }

        private void cACMECheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (cACMECheckBox.Checked)
            {
                cIPGICheckBox.Checked = false;
                cCHOICECheckBox.Checked = false;
            }
        }

        private void cIPGICheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (cIPGICheckBox.Checked)
            {
                cCHOICECheckBox.Checked = false;
                cACMECheckBox.Checked = false;
            }
        }

        private void cCHOICECheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (cCHOICECheckBox.Checked)
            {
                cIPGICheckBox.Checked = false;
                cACMECheckBox.Checked = false;
            }
        }
        public static System.Data.DataTable download2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select ISNULL(MAX(CAST(SEQ AS INT))+1,0) SEQ  from download2 WHERE SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }

        public static System.Data.DataTable download21(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select [PATH]  from download2 WHERE [STATUS] ='嘜頭' AND SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        private void UPLOAD(string FILE)
        {
            string PATH = @"\\acmesrv01\SAP_Share\shipping\";
            string DIR = "//acmesrv01//SAP_Share//shipping//";
            string dd = DateTime.Now.ToString("yyyyMM");
            string server = DIR + dd + "//";

            string filename = Path.GetFileName(FILE);
            System.Data.DataTable dt2 = GetMenu.download2(filename);

            if (dt2.Rows.Count > 0)
            {
                GetMenu.DELdownload2(filename);
            }


                string file = FILE;
                bool FF1 = getrma.UploadFile(file, server, false);
                if (FF1 == false)
                {
                    return;
                }
                      string aa = boardCountTextBox.Text;

                    string[] arrurl = aa.Split(new Char[] { ',' });

                    foreach (string i in arrurl)
                    {


                        System.Data.DataTable GG1 = download2(i);
                        string SEQ = GG1.Rows[0][0].ToString();
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        INSERTDOWNLOAD2(i, SEQ, filename, PATH + de + filename, "");
                    }

        }
        private void button25_Click(object sender, EventArgs e)
        {
            string PATH = @"\\acmesrv01\SAP_Share\shipping\";
            string  DIR = "//acmesrv01//SAP_Share//shipping//";
            string dd = DateTime.Now.ToString("yyyyMM");
            string server = DIR + dd + "//";
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            string filename = Path.GetFileName(opdf.FileName);
            System.Data.DataTable dt2 = GetMenu.download2(filename);

            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("檔案名稱重複,請修改檔名");
            }
            else
            {
                if (result == DialogResult.OK)
                {

                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }

                    string aa = boardCountTextBox.Text;

                    string[] arrurl = aa.Split(new Char[] { ',' });

                    foreach (string i in arrurl)
                    {

                        System.Data.DataTable GG1 = download2(i);
                        string  SEQ = GG1.Rows[0][0].ToString();
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        INSERTDOWNLOAD2(i, SEQ, filename, PATH + de + filename, "嘜頭");
                    }


                    MessageBox.Show("上傳成功");
                }

            }
        }

        private void boardCountTextBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {


                string MEMOT = boardCountTextBox.Text;
                string MEMO = "";
                int G1 = MEMOT.IndexOf("SH20");
                string H1 = MEMOT.Substring(G1, MEMOT.Length - G1);
                if (G1 != -1)
                {
                    string[] arrurl = H1.Split(new Char[] { ',' });

                    foreach (string i in arrurl)
                    {
                        MEMO = i.Substring(0, 14);

                        int T1 = MEMO.IndexOf("SH");

                        if (T1 != -1)
                        {
                            fmShip a = new fmShip();
                            a.PublicString = MEMO;
                            a.Show();
                        }
                    }

                }
            }
            catch { }
        }

        private void button28_Click(object sender, EventArgs e)
        {



                       DialogResult resultS;
                       resultS = MessageBox.Show("請確認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
                       if (resultS == DialogResult.Yes)
                       {
                           System.Data.DataTable G1 = GetDI(shippingCodeTextBox.Text);


                           if (G1.Rows.Count > 0)
                           {

                               System.Data.DataTable G2 = GetDI2(shipping_OBUTextBox.Text);
                               string WHSCODE = "";

                               if (G2.Rows.Count > 0)
                               {
                                   SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                                   oCompany = new SAPbobsCOM.Company();

                                   oCompany.Server = "acmesap";
                                   oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                                   oCompany.UseTrusted = false;
                                   oCompany.DbUserName = "sapdbo";
                                   oCompany.DbPassword = "@rmas";
                                   oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                                   int i = 0; //  to be used as an index

                                   oCompany.CompanyDB = "acmesql02";
                                   oCompany.UserName = "A01";
                                   oCompany.Password = "89206602";
                                   int result = oCompany.Connect();
                                   if (result == 0)
                                   {

                                       SAPbobsCOM.StockTransfer oStock = null;
                                       oStock = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                                 


                                       WHSCODE = G2.Rows[0][0].ToString();
                                       oStock.CardCode = "";
                                       oStock.FromWarehouse = "OT001";
                                       oStock.JournalMemo = shippingCodeTextBox.Text;
                                       System.Data.DataTable G3 = GetDI3(shippingCodeTextBox.Text);

                                       if (G3.Rows.Count > 0)
                                       {
                                           string INV = G3.Rows[0][0].ToString();
                                           oStock.UserFields.Fields.Item("U_ACME_INV").Value = INV;
                                       }
       
                                       for (int s = 0; s <= G1.Rows.Count - 1; s++)
                                       {
                                           string ITEMCODE = G1.Rows[s]["ITEMCODE"].ToString();
                                           string QTY = G1.Rows[s]["QTY"].ToString();

                                           oStock.Lines.ItemCode = ITEMCODE;
                                           oStock.Lines.Quantity = Convert.ToDouble(QTY);
                                           oStock.Lines.WarehouseCode = WHSCODE;
                                           oStock.Lines.Add();
                                       }

                                       int res = oStock.Add();
                                       if (res != 0)
                                       {
                                           MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                                       }
                                       else
                                       {
                                           System.Data.DataTable G4 = GetDI4();
                                           string OWTR = G4.Rows[0][0].ToString();
                                           UpdateOWTR(gPSPhoneTextBox.Text,shippingCodeTextBox.Text);
                              

                                           MessageBox.Show("上傳成功 調撥單號 : " + OWTR);

                                       }
                                   }
                                   else
                                   {
                                       MessageBox.Show(oCompany.GetLastErrorDescription());

                                   }
                               }


                        


                           }
                      
                       }
        }

        private void button29_Click(object sender, EventArgs e)
        {

        }

        private void button26_Click(object sender, EventArgs e)
        {
            if (boardCountNoComboBox.Text != "出口")
            {
                MessageBox.Show("貿易形式出口才可匯入");
                return;
            
            }
            DialogResult resultS;
            resultS = MessageBox.Show("請確認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (resultS == DialogResult.Yes)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = "acmesql02";
                oCompany.UserName = "A01";
                oCompany.Password = "89206602";
                int result = oCompany.Connect();
                if (result == 0)
                {

                    SAPbobsCOM.StockTransfer oStock = null;
                    oStock = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    System.Data.DataTable G1 = GetDIB(shippingCodeTextBox.Text);

                    if (G1.Rows.Count > 0)
                    {

                        string WHNAME = G1.Rows[0]["WHNAME"].ToString();
                        System.Data.DataTable G2 = GetDI2(WHNAME);
                        string WHSCODE = "";

                        if (G2.Rows.Count > 0)
                        {
                            WHSCODE = G2.Rows[0][0].ToString();
                            oStock.CardCode = "";
                            oStock.FromWarehouse = WHSCODE;
                            oStock.JournalMemo = shippingCodeTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_reason").Value = "直接" + boardCountNoComboBox.Text + "-" + cardNameTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_USER").Value = createNameTextBox.Text;
                            for (int s = 0; s <= G1.Rows.Count - 1; s++)
                            {
                                string ITEMCODE = G1.Rows[s]["ITEMCODE"].ToString();
                                string QTY = G1.Rows[s]["QTY"].ToString();

                                oStock.Lines.ItemCode = ITEMCODE;
                                oStock.Lines.Quantity = Convert.ToDouble(QTY);
                                oStock.Lines.WarehouseCode = "TW006";
                                oStock.Lines.Add();
                            }


                        }

                    }

                    int res = oStock.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
                        gPSPhoneTextBox1.Text = OWTR;
                        UpdateOWTR(gPSPhoneTextBox1.Text,shippingCodeTextBox.Text);
                        MessageBox.Show("上傳成功 調撥單號 : " + OWTR);

                    }


                }
                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }
            }
        }

        private void pINOTextBox_Enter(object sender, EventArgs e)
        {
            pINOTextBox.ImeMode = ImeMode.OnHalf;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (wH_Item4DataGridView.SelectedRows.Count > 0)
            
            {
                DataGridViewRow row;
                StringBuilder sb = new StringBuilder();

                for (int i = wH_Item4DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = wH_Item4DataGridView.SelectedRows[i];

                    sb.Append("'" + row.Cells["Docentry1"].Value.ToString() + "',");
                }
                sb.Remove(sb.Length - 1, 1);

                System.Data.DataTable dt1 = Getwhitem42(shippingCodeTextBox.Text, sb.ToString());
                System.Data.DataTable dt2 = wh.WH_Item;

                int h = 0;
                string DOC = "";
                string LINE = "";
                int G1 = wH_ItemDataGridView.Rows.Count;
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = (G1 + i).ToString();
                    DOC = drw["Docentry"].ToString();
                    LINE = drw["linenum"].ToString();
                    drw2["Docentry"] = DOC;
                    drw2["linenum"] = LINE;
                    drw2["ItemRemark"] = drw["ItemRemark"];
                    drw2["WHName"] = drw["WHName"];
                    drw2["ItemCode"] = drw["ItemCode"];
                    drw2["Dscription"] = drw["Dscription"];
                    drw2["Quantity"] = drw["Quantity"];
                    drw2["Remark"] = drw["Remark"];
                    drw2["INV"] = drw["INV"];
                    drw2["PiNo"] = drw["PiNo"];
                    drw2["NowQty"] = drw["NowQty"];
                    drw2["Ver"] = drw["Ver"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["Invoice"] = drw["Invoice"];
                    drw2["FrgnName"] = drw["FrgnName"];
                    drw2["Shipdate"] = drw["Shipdate"];
                    drw2["cardcode"] = drw["cardcode"];


                    drw2["U_PAY"] = drw["U_PAY"];
                    drw2["U_SHIPDAY"] = drw["U_SHIPDAY"];
                    drw2["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
                    drw2["U_MARK"] = drw["U_MARK"];

                    drw2["U_MEMO"] = drw["U_MEMO"].ToString();
                    drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                    drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                    drw2["PO"] = drw["PO"];
                    drw2["TREETYPE"] = drw["TREETYPE"];
                    if (drw["U_PAY"].ToString().Trim() == "FOC")
                    {
                        h = i;

                    }

                    drw2["FrgnName1"] = drw["FrgnName"];
                    drw2["LOCATION"] = drw["LOCATION"];
                    dt2.Rows.Add(drw2);
                }

                wH_ItemBindingSource.EndEdit();
                wH_ItemTableAdapter.Update(wh.WH_Item);
                wh.WH_Item.AcceptChanges();
            }
            else
            {
                MessageBox.Show("請選擇項目");
            }

        }




        public System.Data.DataTable GetDI5()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT JrnlMemo WHNO,DOCENTRY  FROM OWTR WHERE SUBSTRING(JrnlMemo,1,2)='WH' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetFEE1()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(MIN(T1.U_SIZE),0) SIZE   FROM WH_ITEM4 T0 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE  COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetFEE1S()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT CASE WHEN oBUShipTo  LIKE '%台南%' THEN 1 ");
            sb.Append(" WHEN oBUShipTo  LIKE '%高雄%' THEN 1 ");
            sb.Append(" WHEN oBUShipTo  LIKE '%屏東%' THEN 1 ");

            sb.Append(" WHEN oBUShipTo  LIKE '%嘉義%' THEN 1 ");
            sb.Append(" WHEN oBUShipTo  LIKE '%南投%' THEN 1 ");
            sb.Append(" WHEN oBUShipTo  LIKE '%宜蘭%' THEN 2");
            sb.Append(" WHEN oBUShipTo  LIKE '%花蓮%' THEN 3 ");
            sb.Append(" WHEN oBUShipTo  LIKE '%台東%' THEN 4 ELSE 5");
            sb.Append(" END");
            sb.Append("   FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetOPTW(string U_ACME_INV)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱 from oclg t2     ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)     ");
            sb.Append(" inner join opdn t4 on(cast(t2.docentry as varchar)=cast(t4.docentry as varchar))                 ");
            sb.Append(" where  T2.DOCTYPE='20'  ");
            //    sb.Append(" and   t4.U_ACME_INV=@U_ACME_INV");
            sb.Append(" and   t4.U_ACME_INV   LIKE '%" + U_ACME_INV + "%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOPTWT(string U_ACME_INV)
        {
          
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],''+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱 ");
            sb.Append("     from ATC1 T3");
            sb.Append("     inner join opdn t4 on(T3.ABSENTRY=t4.ATCENTRY)             ");
            sb.Append(" WHERE   t4.U_ACME_INV   LIKE '%" + U_ACME_INV + "%' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void K1()
        {
            try
            {
                decimal CARTON = util.CINT(kQTYTextBox.Text.Trim());
                if (kQTY3TextBox.Text != "")
                {
                    CARTON = util.CINT(kQTY3TextBox.Text.Trim());
                }
                decimal TFEE = 0;
                if (k1TextBox.Text == "航通" || k1TextBox.Text == "昕航")
                {
                    if (shippingCodeTextBox.Text != "")
                    {
                        System.Data.DataTable F1 = GetFEE1();

                        if (F1.Rows.Count > 0)
                        {
                            decimal SIZE = Convert.ToDecimal(F1.Rows[0][0]);

                            decimal FEE = 0;
                            if (SIZE <= 32)
                            {
                                FEE = 90 * CARTON;
                                System.Data.DataTable F1S = GetFEE1S();
                                string FE=F1S.Rows[0][0].ToString();
                                if (FE == "1")
                                {
                                    if (CARTON > 10)
                                    {
                                        FEE = 95 * 10 + 70 * (CARTON - 10);
                                    }
                                    else
                                    {
                                        FEE = 95 * CARTON;
                                    }
                                }
                                else if (FE == "5")
                                {
                                    if (CARTON > 10)
                                    {
                                        FEE = 90 * 10 + 70 * (CARTON - 10);
                                    }
                                    else
                                    {
                                        FEE = 90 * CARTON;
                                    }
                                }
                                else  if (FE == "2")
                                {
                                    FEE = 200 * CARTON;
                                }
                                else if (FE == "3")
                                {
                                    FEE = 250 * CARTON;
                                }
                                else if (FE == "4")
                                {
                                    FEE = 300 * CARTON;
                                }
                            }
                            else if (SIZE > 32 && SIZE <= 37)
                         //   if (SIZE <= 32)
                            {
                                if (CARTON > 0 && CARTON <= 10)
                                {
                                    FEE = 200 * CARTON + 100;
                                    // FEE = 85 * CARTON;
                                }
                                if (CARTON > 10)
                                {
                                    FEE = 2000 + (70 * (CARTON - 10)) + 100;
                                }


                            }
                            else if (SIZE > 37 && SIZE <= 42)
                            {
                                FEE = 350* CARTON;

                            }
                            else
                            {
                                FEE = 450 * CARTON;

                            }
                            TFEE = FEE ;
                            //kQTYTextBox
                        }
                    }
                }
                if (k1TextBox.Text == "新竹")
                {

                    if (kQTYTextBox.Text.Trim() == "1")
                    {

                        TFEE = 95;
                    }
                    else
                    {
                        TFEE = 95 + (CARTON - 1) * 75;
                    }
                }

                decimal kW = util.CINT(kWTextBox.Text);
                decimal kP = util.CINT(kPTextBox.Text);
                decimal kF = util.CINT(kFTextBox.Text);
                decimal kZ = util.CINT(kZTextBox.Text);
                decimal kH1 = util.CINT(kH1TextBox.Text);
                decimal kE = util.CINT(kETextBox.Text);
                //kETextBox
                kTTextBox.Text = (TFEE + kE + kW + kP + kF + kZ + kH1).ToString();


            }
            catch { }
        }
        private void CARS()
        {
            try
            {
                decimal FA = 0;
                decimal FA2 = 0;
                decimal FA3 = 0;
                if (!String.IsNullOrEmpty(cTTextBox.Text) )
                {
                    System.Data.DataTable FF = GETFEE(cTTextBox.Text);
          
                    if (FF.Rows.Count > 0)
                    {
                        FA = Convert.ToDecimal(FF.Rows[0][0]);
                    }

                    decimal FADD = 0;
                    decimal FADD2 = 0;
                    decimal FADD3 = 0;
                    decimal CB = 0;
       
                    if (cT2TextBox.Text != "")
                    {
                        System.Data.DataTable FFT2 = GETFEE(cT2TextBox.Text);
                        if (FFT2.Rows.Count > 0)
                        {
                            FA2 = Convert.ToDecimal(FFT2.Rows[0][0]);

                        }
                    }

                    if (cT3TextBox.Text != "")
                    {
                        System.Data.DataTable FFT3 = GETFEE(cT3TextBox.Text);
                        if (FFT3.Rows.Count > 0)
                        {
                            FA3 = Convert.ToDecimal(FFT3.Rows[0][0]);

                        }
                    }

                    if (cGADCheckBox.Checked)
                    {
                        System.Data.DataTable FF2 = GETFEE2(cTTextBox.Text);
                        if (FF2.Rows.Count > 0)
                        {
                            FADD = Convert.ToDecimal(FF2.Rows[0][0]);
                        }

                        if (cT2TextBox.Text != "")
                        {
                            System.Data.DataTable FF3 = GETFEE2(cT2TextBox.Text);
                            if (FF3.Rows.Count > 0)
                            {
                                FADD2 = Convert.ToDecimal(FF3.Rows[0][0]);
                            }
                        }

                        if (cT3TextBox.Text != "")
                        {
                            System.Data.DataTable FF4 = GETFEE2(cT3TextBox.Text);
                            if (FF4.Rows.Count > 0)
                            {
                                FADD3 = Convert.ToDecimal(FF4.Rows[0][0]);
                            }
                        }
                    }


                    //聯倉
                    if (c1TextBox.Text == "聯倉")
                    {
                        if (cBCheckBox.Checked)
                        {
                            cBFTextBox.Text = CAR1().ToString();
                        }
                    }

                    cGADFTextBox.Text = (FADD + FADD2 + FADD3).ToString();
                  
                }

                decimal CBF = util.CINT(cBFTextBox.Text);
                decimal CGADF = util.CINT(cGADFTextBox.Text);
                decimal cU = util.CINT(cUTextBox.Text);
                decimal cLI = util.CINT(cLITextBox.Text);
                decimal cGAB = util.CINT(cGABTextBox.Text);
                decimal cYA = util.CINT(cYATextBox.Text);
                decimal cRE = util.CINT(cRETextBox.Text);
                decimal cQO = util.CINT(cQOTextBox.Text);

                decimal FS = util.CINT(cETextBox.Text);

                //cETextBox
                kT2TextBox.Text = (FS + FA + FA2 + FA3 + CGADF + CBF + cU + cLI + cGAB + cYA + cRE + cQO).ToString();

            }
            catch { }
       
        }
        private void COU()
        {
            try
            {
                decimal lOU = 0;
                decimal lOUS = 0;
                decimal lOUB = 0;
                decimal lOUSOR = 0;
                decimal LOUTS = 0;
                decimal lOU20 = 0;
                decimal lOU40 = 0;

                //lOUETextBox
                decimal CARTON = util.CINT(kQTYTextBox.Text.Trim());
                decimal PLATE = util.CINT(pQTYTextBox.Text.Trim());
                decimal CP = PLATE;
                if (CARTON > 0)
                {
                    CP = PLATE + 1;
                }
                decimal lOUD = util.CINT(lOUDTextBox.Text.Trim()) * 150;
                decimal lOUE = util.CINT(lOUETextBox.Text.Trim());
                decimal lOUG = util.CINT(lOUGTextBox.Text.Trim());
             
                if (lOUCTextBox.Text == "聯倉")
                {
                    decimal lOUT = util.CINT(lOUTTextBox.Text.Trim());
                    decimal lOUB2 = util.CINT(lOUBTextBox.Text.Trim());

                    lOU = 5 * lOUT;

                   
                    if (lOUSCheckBox.Checked)
                    {
                        lOUS = 5 * CARTON + 5 * PLATE;
                    }
                    //打邊條
                    lOUB = 100  * lOUB2;
         
                    //Sorting費
                    if (lOUSORCheckBox.Checked)
                    {
                        lOUSOR = 400 * PLATE;
                    }
                    //拆板費
                    if (lOUTSCheckBox.Checked)
                    {
                        LOUTS = 150 * PLATE;
                    }

                    if (lOU20CheckBox.Checked)
                    {
                        lOU20 = 2000;
                    }

                    if (lOU40CheckBox.Checked)
                    {
                        lOU40 = 3000;
                    }

                    //lINCheckBox
                }

                if (lOUCTextBox.Text == "新得利倉")
                {
                    if (lOU20CheckBox.Checked)
                    {
                        lOU20 = 1500;
                    }

                    if (lOU40CheckBox.Checked)
                    {
                        lOU40 = 3000;
                    }


                
                }
                if (lOUCTextBox.Text == "大發倉")
                {
                    if (lOU20CheckBox.Checked)
                    {
                        lOU20 = 2500;
                    }

                    if (lOU40CheckBox.Checked)
                    {
                        lOU40 = 3000;
                    }

                    if (lOUCheckBox.Checked)
                    {
                        lOU = 130 * CP;
                    }
                }
                lOUFTextBox.Text = (lOU20 + lOU40).ToString();
                decimal lOUF = util.CINT(lOUFTextBox.Text.Trim());

                //        lOUBFTextBox.Text 
                //
                lOUFEETextBox.Text = (lOU + lOUS + lOUB + lOUSOR + LOUTS + lOUF + lOUD + lOUE + lOUG).ToString();
            }
            catch { }
        }

        private void CIN()
        {
            try
            {
                decimal lIN = 0;
                decimal lIN20 = 0;
                decimal lIN40 = 0;

                //lINGTextBox
                decimal CARTON = util.CINT(kQTYTextBox.Text.Trim());
                decimal PLATE = util.CINT(pQTYTextBox.Text.Trim());
                decimal lING = util.CINT(lINGTextBox.Text.Trim());
                decimal lINE = util.CINT(lINETextBox.Text.Trim());
                
                decimal CP = PLATE;
                if (CARTON > 0)
                {
                    CP = PLATE + 1;
                }

                if (lINCTextBox.Text == "聯倉")
                {

                    if (lIN20CheckBox.Checked)
                    {
                        lIN20 = 1000;
                    }

                    if (lIN40CheckBox.Checked)
                    {
                        lIN40 = 2000;
                    }

                    //lINCheckBox
                }

                if (lINCTextBox.Text == "新得利")
                {
                    if (lIN20CheckBox.Checked)
                    {
                        lIN20 = 1500;
                    }

                    if (lIN40CheckBox.Checked)
                    {
                        lIN40 = 3000;
                    }



                }
                if (lINCTextBox.Text == "大發倉")
                {
                    if (lIN20CheckBox.Checked)
                    {
                        lIN20 = 2500;
                    }

                    if (lIN40CheckBox.Checked)
                    {
                        lIN40 = 3000;
                    }


           
                }
                if (lINCheckBox.Checked)
                {
                    lIN = 150 * CP;
                }
                // 新得利倉
                //
                lINFEETextBox.Text = (lIN20 + lIN40 + lIN + lINE + lING).ToString();
            }
            catch { }
        }
        public void AddFEE()
        {
            SqlConnection Connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into WH_FEE(ShippingCode,CB,CGAD,LOUS,LOUSOR,LOU20,LOU40,LOUTS,LIN,LOU,LIN20,LIN40,LINZ) values(@ShippingCode,'Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked','Unchecked')", Connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));



            try
            {

                try
                {
                    Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                Connection.Close();
            }

        }


        private System.Data.DataTable GETFEE(string WEIGHT)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT AMT FROM WH_CARFEE WHERE ([WEIGHT] =@WEIGHT OR REPLACE([WEIGHT],'T','')=REPLACE(@WEIGHT,'T','')) AND LOCATION =@LOCATION AND CARNAME =@CARNAME");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WEIGHT", WEIGHT));
            command.Parameters.Add(new SqlParameter("@LOCATION", cLOCTextBox.Text));
            command.Parameters.Add(new SqlParameter("@CARNAME", c1TextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GETFEE2(string WEIGHT)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql;


            sql = "SELECT AMT FROM WH_CARFEE WHERE [WEIGHT] =@WEIGHT AND LOCATION ='加點' AND CARNAME =@CARNAME";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WEIGHT", WEIGHT));
            command.Parameters.Add(new SqlParameter("@CARNAME", c1TextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        private void kTTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void k1ComboBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }
        private decimal CAR1()
        {
            decimal g = 0;
            decimal CARTON = util.CINT(kQTYTextBox.Text.Trim());
            if (CARTON >= 15 && CARTON <= 19)
            {
                g = 300;
            }
            else if (CARTON >= 20 && CARTON <= 24)
            {
                g = 400;
            }
            else if (CARTON >= 25 && CARTON <= 30)
            {
                g = 500;
            }
            else if (CARTON >= 31)
            {
                g = 600;
            }
            return g;
        }

        public System.Data.DataTable GetBU1()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT REPLACE([WEIGHT],'T','')+'T' T FROM WH_CARFEE WHERE CARNAME =@CARNAME ORDER BY REPLACE([WEIGHT],'T','')+'T' ";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARNAME", c1TextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }
        public System.Data.DataTable GetBU2()
        {
            SqlConnection MyConnection = globals.Connection;
            string sql;


            sql = "SELECT DISTINCT LOCATION FROM WH_CARFEE WHERE CARNAME =@CARNAME AND LOCATION <> '加點費用'";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARNAME", c1TextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

    

        private void cLOCComboBox_MouseClick(object sender, MouseEventArgs e)
        {
       
        }

        private void cTComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cLOCComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cGADCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cBCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void lINCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lIN20CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lIN40CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lINGTextBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lINETextBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void cUTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cLITextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cGABTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cYATextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cRETextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cQOTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cETextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void kWTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kPTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kFTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kZTextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kH1TextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kETextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void kQTYTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
            CIN();
            K1();
            COU();
        }

        private void lOUBCheckBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUCheckBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUDTextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUSORCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUBCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUTCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUTSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOU20CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOU40CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUGTextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUETextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lINCComboBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void button29_Click_1(object sender, EventArgs e)
        {
            //RCH_2.xls -> 兩欄要資料一樣 儲存格設為 =PN =PO =Quantity

            System.Data.DataTable dtData = wH_ItemDataGridView.DataSource as System.Data.DataTable;

            //產生 QrCode 資料
            // MakeQrCodeValue(dtData);

            string Dir = GetExePath() + "\\Output\\";
            string DirTemplate = GetExePath() + "\\XlsTemplate\\";

            if (!Directory.Exists(Dir))
            {
                Directory.CreateDirectory(Dir);
            }

            if (!Directory.Exists(DirTemplate))
            {
                Directory.CreateDirectory(DirTemplate);
            }



            string FileName = GetExePath() + "\\" + shippingCodeTextBox.Text + ".xls";

            string Template = GetExePath() + "\\XlsTemplate\\" + "RCH.xls";

            ExcelRch(dtData, Template, FileName);
        }
        public bool ExcelRch(System.Data.DataTable dt,
string Template, string SaveFileName, int Interval = 6, Int32 PageBreak = 4, string EndCell = "A5")
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;

            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                //{
                //    MessageBox.Show(nm.NameLocal);
                //}


                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;


                //依  dt 筆數產生分頁
                Int32 PrintQty = 1;

                DataRow dr;

                Microsoft.Office.Interop.Excel.Range cell = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    //複製範本
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);

                    //xlSht.Copy(Type.Missing, xlWb.Sheets[xlWb.Sheets.Count]); // copy
                    //xlWb.Sheets[xlWb.Sheets.Count].Name = "NEW SHEET";
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[i+1];
                    //excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                    excelSheet.Name = "Q" + (i + 1).ToString();

                    //處理範本
                    string BarcodeFile = "";
                    int shapeCount = excelSheet.Shapes.Count;

                    string[] shapeNames = new string[shapeCount];

                    for (int iShape = 1; iShape <= shapeCount; iShape++)
                    {
                        shapeNames[iShape - 1] = excelSheet.Shapes.Item(iShape).Name;
                    }

                    for (int iShape = 0; iShape <= shapeNames.Length - 1; iShape++)
                    {
                        //  string ShapeName = SheetTemplate.Shapes.Item(iShape).Name;
                        string ShapeName = shapeNames[iShape];

                        if (ShapeName == "Logo" || ShapeName == "Group") continue;

                        //if (ShapeName == "QrCode")
                        //{
                        //    GetQrCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        //    BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        //    UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                        //    continue;
                        //}
                        //更換圖片
                        GetBarCode(ShapeName, Convert.ToString(dr[ShapeName]));
                        BarcodeFile = GetExePath() + "\\Output\\" + ShapeName + ".jpg";
                        UpdatePicture(excelSheet, ShapeName, BarcodeFile);
                    }

                    //Dscription = Convert.ToString(dr["Dscription"]);
                    ////SetCellValue(excelSheet, "A4", Dscription);
                    //Microsoft.Office.Interop.Excel.Range cell = excelSheet.Evaluate("Dscription") as Microsoft.Office.Interop.Excel.Range;
                    //if (cell != null) cell.Value = Dscription;

                    //產地 = Convert.ToString(dr["產地"]);
                    //cell = excelSheet.Evaluate("產地") as Microsoft.Office.Interop.Excel.Range;
                    //if (cell != null) cell.Value = 產地;

                    foreach (Microsoft.Office.Interop.Excel.Name nm in excelworkBook.Names)
                    {
                        //MessageBox.Show(nm.NameLocal);
                        string FieldName = nm.NameLocal;

                        try
                        {
                            cell = excelSheet.Evaluate(FieldName) as Microsoft.Office.Interop.Excel.Range;
                            if (cell != null) cell.Value = Convert.ToString(dr[FieldName]);
                        }
                        catch
                        {

                        }
                    }

                    //excelSheet.Activate();

                    // Range r = SheetTemplate.get_Range("A1,A6", Type.Missing) as Range;
                    //更換拷貝區域
                    //string EndCell = "A5";
                    Range r = excelSheet.get_Range("A1", EndCell) as Range;
                    // r.EntireRow.Select();
                    r.EntireRow.Copy(Type.Missing);

                    Range d;
                    PrintQty = Convert.ToInt32(dt.Rows[i]["PrintQty"]);

                    // int Interval = 6;


                    //Int32 PageBreak = 4;
                    for (int q = 1; q <= PrintQty - 1; q++)
                    {
                        d = excelSheet.get_Range("A" + (Interval * q + 1).ToString(), Type.Missing) as Range;
                        d.Select();
                        excelSheet.Paste(Type.Missing);
                        //sheet.HPageBreaks.Add(sheet.Range["A11"]);
                        if (q % PageBreak == 0)
                        {
                            excelSheet.HPageBreaks.Add(excelSheet.Range["A" + (Interval * q + 1).ToString()]);
                        }
                    }




                }


                SheetTemplate.Delete();

                excelworkBook.SaveAs(SaveFileName); ;
                excelworkBook.Close();
                //excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (excelSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(SheetTemplate);


                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                SheetTemplate = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
            }
        }


        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void GetBarCode(string CodeName, string Data)
        {
            string Url = "https://generator.barcodetools.com/barcode.png?gen=0&data={0}&bcolor=FFFFFF&fcolor=000000&tcolor=000000&fh=14&bred=0&w2n=2.5&xdim=2&w=&h=120&debug=1&btype=7&angle=0&quiet=1&balign=2&talign=2&guarg=1&text=1&tdown=1&stst=1&schk=0&cchk=1&ntxt=1&c128=0";

            Url = string.Format(Url, Data);
            string PicFile = GetExePath() + "\\Output\\" + CodeName + ".jpg";
            GetUrlPicture(Url, PicFile);
        }

        private void GetUrlPicture(string Url, string PicName)
        {
            WebResponse response = default(WebResponse);
            Stream remoteStream = default(Stream);
            StreamReader readStream = default(StreamReader);
            WebRequest request = WebRequest.Create(Url);
            response = request.GetResponse();
            remoteStream = response.GetResponseStream();
            readStream = new StreamReader(remoteStream);

            System.Drawing.Image img = System.Drawing.Image.FromStream(remoteStream);

            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                byte[] byteImage = ms.ToArray();
                // imgBarCode.ImageUrl = "data:image/png;base64," + Convert.ToBase64String(byteImage);
            }
            img.Save(PicName);

            response.Close();
            remoteStream.Close();
            readStream.Close();
        }

        private void UpdatePicture(Microsoft.Office.Interop.Excel.Worksheet excelSheet, string ShapeName, string QrFileName)
        {
            //取得來源圖片的位置大小
            float iLeft = 0;
            float iTop = 0;
            float iWidth = 0;
            float iHeight = 0;

            Shape x = excelSheet.Shapes.Item(ShapeName);

            iLeft = x.Left;
            iTop = x.Top;
            iWidth = x.Width;
            iHeight = x.Height;

            x.Delete();

            x = excelSheet.Shapes.AddPicture(QrFileName,
                Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue, iLeft, iTop, iWidth, iHeight);

            x.Name = ShapeName;
        }

        private void cGTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void lOUBTextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            lOUCTextBox.Text = comboBox5.Text;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            c1TextBox.Text = comboBox6.Text;
        }

        private void comboBox7_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU1();

            comboBox7.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox7.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            cTTextBox.Text = comboBox7.Text;
        }

        private void comboBox8_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU1();

            comboBox8.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox8.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox9_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU1();

            comboBox9.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox9.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            cT2TextBox.Text = comboBox8.Text;
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            cT3TextBox.Text = comboBox9.Text;
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            lINCTextBox.Text = comboBox11.Text;
        }

        private void cTTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cT2TextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cT3TextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void k1TextBox_TextChanged(object sender, EventArgs e)
        {
            K1();
        }

        private void pQTYTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
            CIN();
            K1();
            COU();
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            k1TextBox.Text = comboBox10.Text;
        }

        private void lOUTTextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }

        private void lOUCTextBox_TextChanged(object sender, EventArgs e)
        {
            COU();
        }


        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            lINHTextBox.Text = comboBox12.Text;
        }

        private void wH_ItemDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "Invoice")
                {
                    //DocentryS
                    string Invoice = wH_ItemDataGridView.CurrentRow.Cells["Invoice"].Value.ToString();
                    string ItemCode2 = wH_ItemDataGridView.CurrentRow.Cells["ItemCode2"].Value.ToString();
                    string DocentryS = wH_ItemDataGridView.CurrentRow.Cells["DocentryS"].Value.ToString();
                    if (!String.IsNullOrEmpty(Invoice))
                    {
                        System.Data.DataTable gg1 = GetOPTWT(Invoice);
                        if (gg1.Rows.Count == 0)
                        {
                            gg1 = GetOPTW(Invoice);

                        }
                  
   
                        if (gg1.Rows.Count == 0 && Invoice.Length >10)
                        {
                            string INV = Invoice.Substring(0, 10);
                            gg1 = GetOPTW(INV);
                        }
                        if (gg1.Rows.Count > 0)
                        {
                            for (int i2 = 0; i2 <= gg1.Rows.Count - 1; i2++)
                            {
                                string path = gg1.Rows[i2]["path"].ToString();
                                string 路徑 = gg1.Rows[i2]["路徑"].ToString();
                                string 檔案名稱 = gg1.Rows[i2]["檔案名稱"].ToString();

                                string aa = path + "\\" + 路徑;

                                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                                string filename = 檔案名稱;
                                string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                                System.IO.File.Copy(aa, NewFileName, true);

                                System.Diagnostics.Process.Start(NewFileName);
                            }

                        }
               


                    }
                    else
                    {
                   
                            APOWHS frm1 = new APOWHS();
                            frm1.ITEMCODE = ItemCode2;

                            if (frm1.ShowDialog() == DialogResult.OK)
                            {
                                try
                                {
                                    string INV = frm1.a;
                                    if (!String.IsNullOrEmpty(INV))
                                    {
                                        string INVOICE = "";
                                        string INVOICEDATE = "";
                                        int ii = 0;
                                        string[] arrurl = INV.Split(new Char[] { '/' });

                                        foreach (string i in arrurl)
                                        {
                                            ii++;
                                            if (ii == 1)
                                            {
                                                INVOICE = i.ToString();
                                            }

                                            if (ii == 2)
                                            {
                                                INVOICEDATE = i.ToString();
                                            }
                                        }
                                        UpdateINVOICE(INVOICE, INVOICEDATE, DocentryS);
                                        wH_ItemTableAdapter.Fill(wh.WH_Item, MyID);
                                    }

                                }
                                catch { }
                            }
                        
                    }
                }
            }
            catch { }
        }

  
        


        private void wH_ItemDataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            //Docentry3
            try
            {
                if (wH_ItemDataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    listBox1.Items.Clear();
                    for (int i = wH_ItemDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = wH_ItemDataGridView.SelectedRows[i];

                        listBox1.Items.Add(row.Cells["DocentryS"].Value.ToString());

                    }

                }
                else
                {
                    listBox1.Items.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            if (boardCountNoComboBox.Text != "出口")
            {
                MessageBox.Show("貿易形式出口才可匯入");
                return;

            }
            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("請選擇列");
                return;
            }
            DialogResult resultS;
            resultS = MessageBox.Show("請確認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (resultS == DialogResult.Yes)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = "acmesql02";
                oCompany.UserName = "A01";
                oCompany.Password = "89206602";
                int result = oCompany.Connect();
                if (result == 0)
                {
                    ArrayList al = new ArrayList();

                    for (int i2 = 0; i2 <= listBox1.Items.Count - 1; i2++)
                    {
                        al.Add(listBox1.Items[i2].ToString());
                    }
                    StringBuilder sb = new StringBuilder();



                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);
                    SAPbobsCOM.StockTransfer oStock = null;
                    oStock = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    System.Data.DataTable G1 = GetDIBM(sb.ToString());

                    if (G1.Rows.Count > 0)
                    {

                        string WHNAME = G1.Rows[0]["WHNAME"].ToString();
                        System.Data.DataTable G2 = GetDI2(WHNAME);
                        string WHSCODE = "";

                        if (G2.Rows.Count > 0)
                        {
                            WHSCODE = G2.Rows[0][0].ToString();
                            oStock.CardCode = "";
                            oStock.FromWarehouse = WHSCODE;
                            oStock.JournalMemo = shippingCodeTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_reason").Value = "直接" + boardCountNoComboBox.Text + "-" + cardNameTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_USER").Value = createNameTextBox.Text;
                            for (int s = 0; s <= G1.Rows.Count - 1; s++)
                            {
                                string ITEMCODE = G1.Rows[s]["ITEMCODE"].ToString();
                                string QTY = G1.Rows[s]["QTY"].ToString();

                                oStock.Lines.ItemCode = ITEMCODE;
                                oStock.Lines.Quantity = Convert.ToDouble(QTY);
                                oStock.Lines.WarehouseCode = "TW006";
                                oStock.Lines.Add();
                            }


                        }

                    }

                    int res = oStock.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
                        gPSPhoneTextBox1.Text = OWTR;
                        UpdateOWTR(gPSPhoneTextBox1.Text, shippingCodeTextBox.Text);
                        MessageBox.Show("上傳成功 調撥單號 : " + OWTR);

                    }


                }
                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            if (boardCountNoComboBox.Text != "出口")
            {
                MessageBox.Show("貿易形式出口才可匯入");
                return;

            }

     
            DialogResult resultS;
            resultS = MessageBox.Show("請確認是否要匯入", "YES/NO", MessageBoxButtons.YesNo);
            if (resultS == DialogResult.Yes)
            {
                SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

                oCompany = new SAPbobsCOM.Company();

                oCompany.Server = "acmesap";
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = "sapdbo";
                oCompany.DbPassword = "@rmas";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                int i = 0; //  to be used as an index

                oCompany.CompanyDB = "acmesql02";
                oCompany.UserName = "A01";
                oCompany.Password = "89206602";
                int result = oCompany.Connect();
                if (result == 0)
                {
   
                    SAPbobsCOM.StockTransfer oStock = null;
                    oStock = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    System.Data.DataTable G1 = GetDIBM2(shippingCodeTextBox.Text);

                    if (G1.Rows.Count > 0)
                    {

                            oStock.CardCode = "";
                            oStock.FromWarehouse = "Z0005";
                            oStock.JournalMemo = shippingCodeTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_reason").Value = "直接" + boardCountNoComboBox.Text + "-" + cardNameTextBox.Text;
                            oStock.UserFields.Fields.Item("U_ACME_USER").Value = createNameTextBox.Text;
                            for (int s = 0; s <= G1.Rows.Count - 1; s++)
                            {
                                string ITEMCODE = G1.Rows[s]["ITEMCODE"].ToString();
                                string QTY = G1.Rows[s]["QTY"].ToString();

                                oStock.Lines.ItemCode = ITEMCODE;
                                oStock.Lines.Quantity = Convert.ToDouble(QTY);
                                oStock.Lines.WarehouseCode = "TW006";
                                oStock.Lines.Add();
                            }


                        

                    }

                    int res = oStock.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        System.Data.DataTable G4 = GetDI4();
                        string OWTR = G4.Rows[0][0].ToString();
                        gPSPhoneTextBox1.Text = OWTR;
                        UpdateOWTR(gPSPhoneTextBox1.Text, shippingCodeTextBox.Text);
                        MessageBox.Show("上傳成功 調撥單號 : " + OWTR);

                    }


                }
                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }
            }
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            cLOCTextBox.Text = comboBox13.Text;
        }

        private void cLOCTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void comboBox13_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU2();

            comboBox13.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox13.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void lINCTextBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lINHTextBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void lINGATextBox_TextChanged(object sender, EventArgs e)
        {
            CIN();
        }

        private void cBFTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void cGADFTextBox_TextChanged(object sender, EventArgs e)
        {
            CARS();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            UpdateWhNoAuto(shippingCodeTextBox.Text, "Y");
        }


        private void UpdateWhNoAuto(string WhNo, string Flag)
        {
            string Sql = "update Wh_Main set PrintFlag='{1}' where ShippingCode='{0}'";
            Sql = string.Format(Sql, WhNo, Flag);

            UpdateData(Sql);

        }
        public void UpdateData(string Sql)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand(Sql, connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@DocType", DocType));
            //command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            //command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            //command.Parameters.Add(new SqlParameter("@Msg", Msg));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            string[] arrurl = iNVOICENO2TextBox.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                string INV = i.ToString();

                VIVI(INV);
            }

        
        }
        public System.Data.DataTable GETVIVI1(string U_ACME_INV)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                         SELECT m.docnum as 單號,d.linenum as 欄號,ISNULL(ow.onhand,0) as 現有數量,Convert(varchar(10),m.docduedate,111) as 交貨日期,d.U_PAY as 付款,d.U_SHIPDAY as 押出貨日,d.U_SHIPSTATUS as 貨況,d.U_MARK as 特殊嘜頭,d.U_MEMO as 注意事項,m.NUMATCARD as PO, ");
            sb.Append("                          d.TREETYPE,'' U_CUSTITEMCODE,'' U_CUSTDOCENTRY ");
            sb.Append("                          ,m.address 工廠地址,Convert(varchar(10),d.u_acme_work,111) 排程日期,");
            sb.Append("                          BUYUNITMSR 單位,d.itemcode as 產品編號,d.dscription as 品名規格,oi.frgnname as 品名規格1,d.quantity as 數量,m.cardcode 客戶編號,m.cardname 客戶名稱, ");
            sb.Append("                          OI.U_GRADE 等級, 版本='V.'+OI.U_VERSION,OI.U_PARTNO PARTNO, ");
            sb.Append("                          m.comments 備註,oi.usertext 主要描述,m.U_LOCATION 產地,M.CARDCODE,M.CARDNAME FROM opdn m ");
            sb.Append("                          left join PDN1 d on m.docentry=d.docentry ");
            sb.Append("                          left join oitm oi on oi.itemcode=d.itemcode ");
            sb.Append("                          left join oitw ow on oi.itemcode=ow.itemcode AND d.WHSCODE=OW.WHSCODE");
            sb.Append("                   WHERE M.U_ACME_INV LIKE '%" + U_ACME_INV + "%' ");
            sb.Append("                          order by m.DOCENTRY,d.visorder  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_INV", U_ACME_INV));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void VIVI(string INV)
        {




            try
            {



                System.Data.DataTable dt1 = null;
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                {
                    dt1 = GETVIVI1(INV);
                }
                if (dt1.Rows.Count > 0)
                {

                    System.Data.DataTable dt2 = null;

                    dt2 = wh.WH_Item4;

                    int M1 = 0;
                    string MCODE = "";
                    DataRow drw3 = dt1.Rows[dt1.Rows.Count - 1];
                    cardCodeTextBox.Text = drw3["CARDCODE"].ToString();
                    cardNameTextBox.Text = drw3["CARDNAME"].ToString();
                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
  

                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        string D1 = drw["品名規格1"].ToString();


                        string 產品編號 = drw["產品編號"].ToString();

                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["Docentry"] = drw["單號"].ToString();
                        string 品名規格 = drw["品名規格"].ToString();

                        drw2["Dscription"] = 品名規格;

                        drw2["itemcode"] = 產品編號;
                        drw2["ItemRemark"] = "收貨採購單";
                        drw2["WHName"] = shipping_OBUTextBox.Text.ToString();
                        decimal SS = Convert.ToDecimal(drw["數量"]);
                        string GH = Convert.ToDouble(SS).ToString();
                        drw2["Quantity"] = GH;
                        drw2["linenum"] = drw["欄號"];


                        System.Data.DataTable QTY1 = GetQTYF(shipping_OBUTextBox.Text.ToString(), 產品編號);
                        if (QTY1.Rows.Count > 0)
                        {
                            QTYF = Convert.ToInt32(QTY1.Rows[0][0]);
                        }
                        drw2["NowQty"] = Convert.ToInt32(drw["現有數量"]) - QTYF;
                        drw2["Ver"] = drw["版本"];
                        if (產品編號.Length > 2)
                        {
                            string gg = 產品編號.Substring(0, 3).ToString().ToUpper();
                            if (gg == "TAP")
                            {

                                int G1 = 品名規格.IndexOf(".");
                                if (G1 != -1)
                                {
                                    drw2["Ver"] = "V." + 品名規格.Substring(G1 + 1, 1);
                                }
                            }
                        }
                        drw2["Grade"] = drw["等級"];
                        string T1 = drw["單位"].ToString();
                        drw2["cardcode"] = drw["單位"];
                        drw2["ShipDate"] = drw["排程日期"];


                        drw2["U_PAY"] = drw["付款"];
                        drw2["U_SHIPDAY"] = drw["押出貨日"];
                        drw2["U_SHIPSTATUS"] = drw["貨況"];
                        drw2["U_MARK"] = drw["特殊嘜頭"];
                        drw2["U_MEMO"] = drw["注意事項"];
                        drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                        drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                        drw2["PO"] = drw["PO"];
                        drw2["LOCATION"] = drw["產地"];
                        string TREETYPE = drw["TREETYPE"].ToString();
                        drw2["TREETYPE"] = TREETYPE;
                        try
                        {
                            hjj = "";

                            if (TREETYPE == "S")
                            {
                                hjj = "母料號";
                                MCODE = 產品編號;
                                M1 = 0;
                            }
                            else if (TREETYPE == "I")
                            {

                                M1++;
                                hjj = MCODE + "-子料號-" + M1.ToString();
                            }
                            else
                            {
                                hjj = drw["PARTNO"].ToString();
                            }

                            drw2["pino"] = hjj;
                        }
                        catch
                        {

                        }
                        if (globals.DBNAME == "達睿生")
                        {
                            if (forecastDayTextBox.Text == "採購單")
                            {
                                drw2["invoice"] = drw["INVOICE"].ToString();
                            }

                            drw2["FrgnName"] = drw["品名規格"];
                        }
                        else
                        {
                            drw2["FrgnName"] = drw["品名規格1"];
                            if (cardCodeTextBox.Text == "0017-00")
                            {
                                drw2["FrgnName"] = drw["品名規格1"] + "-" + drw["等級"];

                            }
                        }

                        dt2.Rows.Add(drw2);

                    }


                    wH_Item4BindingSource.MoveFirst();

                    for (int i = 1; i <= wH_Item4BindingSource.Count; i++)
                    {
                        DataRowView row = (DataRowView)wH_Item4BindingSource.Current;

                        row["SeqNo"] = i;



                        wH_Item4BindingSource.EndEdit();

                        wH_Item4BindingSource.MoveNext();
                    }

                    wH_mainBindingSource.EndEdit();
                    this.wH_mainTableAdapter.Update(wh.WH_main);
                    wh.WH_main.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



          

            wH_mainBindingSource.EndEdit();
 
            wH_Item4BindingSource.EndEdit();
            wH_Item4TableAdapter.Update(wh.WH_Item4);
            wh.WH_Item4.AcceptChanges();

            string memo = "";
            System.Data.DataTable dtt1 = Getwhitem4(shippingCodeTextBox.Text);
            System.Data.DataTable dtt2 = wh.WH_Item;
     
            int h = 0;
            string DOC = "";
            string LINE = "";
            for (int i = 0; i <= dtt1.Rows.Count - 1; i++)
            {
                DataRow drw = dtt1.Rows[i];
                DataRow drw2 = dtt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["SeqNo"] = drw["SeqNo"];
                DOC = drw["Docentry"].ToString();
                LINE = drw["linenum"].ToString();
                drw2["Docentry"] = DOC;
                drw2["linenum"] = LINE;
                drw2["ItemRemark"] = drw["ItemRemark"];
                drw2["WHName"] = drw["WHName"];
                string ITEMCODE = drw["ItemCode"].ToString();
                drw2["ItemCode"] = ITEMCODE;
                drw2["Dscription"] = drw["Dscription"];
                drw2["Quantity"] = drw["Quantity"];
                drw2["Remark"] = drw["Remark"];
                drw2["INV"] = drw["INV"];
                drw2["PiNo"] = drw["PiNo"];
                drw2["NowQty"] = drw["NowQty"];
                drw2["Ver"] = drw["Ver"];
                drw2["Grade"] = drw["Grade"];
                drw2["Invoice"] = drw["Invoice"];
                drw2["FrgnName"] = drw["FrgnName"];
                drw2["Shipdate"] = drw["Shipdate"];
                drw2["cardcode"] = drw["cardcode"];

                memo = drw["U_MEMO"].ToString();

                drw2["U_PAY"] = drw["U_PAY"];
                drw2["U_SHIPDAY"] = drw["U_SHIPDAY"];
                drw2["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
                drw2["U_MARK"] = drw["U_MARK"];
                //   ShipDate2
                drw2["U_MEMO"] = memo;
                drw2["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
                drw2["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
                drw2["PO"] = drw["PO"];
                drw2["TREETYPE"] = drw["TREETYPE"];
                if (drw["U_PAY"].ToString().Trim() == "FOC")
                {
                    h = i;

                }
                //PQTY5
                drw2["FrgnName1"] = drw["FrgnName"];
                drw2["LOCATION"] = drw["LOCATION"];
                if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
                {
                    System.Data.DataTable SHIPDATE = GetSHIPDATE(DOC, LINE);
                    if (SHIPDATE.Rows.Count > 0)
                    {
                        drw2["ShipDate2"] = SHIPDATE.Rows[0][1].ToString();
                    }
                }

                int X1 = ITEMCODE.IndexOf("ACME");
                if (X1 == -1)
                {
                    System.Data.DataTable GE1 = GETPACK(ITEMCODE);
                    if (GE1.Rows.Count > 0)
                    {
                        int QTY = Convert.ToInt32(drw["Quantity"]);
                        int FF1 = Convert.ToInt32(GE1.Rows[0][0]);
                        int FF2 = Convert.ToInt32(GE1.Rows[0][1]);
                        int mod = QTY % FF1;
                        int mod2 = QTY % FF2;
                        int mod3 = QTY / FF2;
                        if (mod2 > 0)
                        {
                            mod3 = mod3 + 1;
                        }
                        drw2["PQTY5"] = GE1.Rows[0][0].ToString();
                        drw2["PQTY1"] = GE1.Rows[0][1].ToString();
                        drw2["LPRINT"] = mod3.ToString();
                        drw2["PQTY2"] = mod2.ToString();
                        drw2["PQTY6"] = mod.ToString();
                        //LPRINT
                    }
                }
                dtt2.Rows.Add(drw2);
            }


            if (memo != "")
            {
                memo = memo + "\r\n";
            }

            string TDATE = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
            {
                System.Data.DataTable SHIPDATE = GetSHIPDATE(DOC, LINE);
                if (SHIPDATE.Rows.Count > 0)
                {
                    TDATE = SHIPDATE.Rows[0][0].ToString();
                }


                string gh = memo + TDATE + "派車/下午快/借出/客戶自取/提供序號資料/需貼箱麥,不需貼板麥/需貼箱麥、需貼箱麥/需每箱貼麥頭/請打三條束帶--提供照片" +
                    Environment.NewLine + TDATE + "派車/提供序號資料/新得利司機自提/聯倉司機自提--請提供長寬高及重量" +
                    Environment.NewLine + TDATE + "派車/提供序號資料--需至聯倉取貨一起派車/需至新得利取貨一起派車" +
                    Environment.NewLine + TDATE + "請併車--提供併車價--請提供長寬高及重量" +
                    Environment.NewLine + "請做包裝明細--請打棧板--請打邊條--等貼完提單及麥頭--請提供照片" +
                    Environment.NewLine + "需貼板麥、需貼箱麥,請告知各板分佈,謝謝~" +
                    Environment.NewLine + "需每箱貼麥頭--需用A4紙列印貼,一次只用一種顏色~~" +
                    Environment.NewLine + "借出--請RMA出貨即除帳,產生費用請掛RMA" +
                    Environment.NewLine + "※ 假如沒原箱換,請找合適的空箱,以安全方式換箱,有任何問題再提出討論,謝謝~~" +
                    Environment.NewLine + "此票調撥為出口,請找箱子無破凹友達箱備貨" +
                    Environment.NewLine + "送貨單及貨, 請不要有進金生字樣~~快遞單及貨及送貨單, 請不要有進金生字樣~~" +
                    Environment.NewLine + "請做包裝明細-需打木箱，請提供打木箱前後的照片，請單箱捆膠膜，並提供照片，貨代會安排木箱行去打木箱，打完木箱後請在提供包裝明細(長寬高重量)";

      

            }



            wH_ItemBindingSource.EndEdit();
            wH_ItemTableAdapter.Update(wh.WH_Item);
            wh.WH_Item.AcceptChanges();

        
        }


        private void button34_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = MakeTableCombine();
            string OutPutFile = "";
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            FileName = lsAppDir + "\\Excel\\wh\\庫存單.xls";
                OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" + "庫存單---"+
                  shippingCodeTextBox.Text + ".xls";
            System.Data.DataTable H2 = GETD1();
            if (H2.Rows.Count > 0)
            {

                DataRow dr = null;

                for (int i = 0; i <= H2.Rows.Count - 1; i++)
                {

       
                    DataRow dd = H2.Rows[i];
                    dr = dtCost.NewRow();

                    dr["DDATE"] = dd["DDATE"].ToString();
                    string ITEMCODE = dd["ITEMCODE"].ToString();
                    dr["ITEMCODE"] = ITEMCODE;
                    dr["DSCRIPTION"] = dd["DSCRIPTION"].ToString();
                    dr["QTY"] = dd["QTY"].ToString();
                    dr["DDATE"] = dd["DDATE"].ToString();
                    dr["INVOICE"] = GETD2(ITEMCODE).Rows[0][0].ToString();
                    dtCost.Rows.Add(dr);
                }
                ExcelReport.ExcelDAVID(dtCost, FileName, OutPutFile);

            }
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("DDATE", typeof(string));
            dt.Columns.Add("ITEMCODE", typeof(string));
            dt.Columns.Add("DSCRIPTION", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("INVOICE", typeof(string));

            return dt;
        }
        string 放貨單 = "";
        string 地址條 = "";
        string 備貨單 = "";
        private void button35_Click(object sender, EventArgs e)
        {
            try
            {

                DELETEFILE();
                DELETEFILE2();
                DialogResult result;
                result = MessageBox.Show("是否要寄出", "Close", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    System.Data.DataTable H2 = GetQTY2();
                    if (H2.Rows.Count > 0)
                    {
                        string QTY = H2.Rows[0][0].ToString();
                        button799(QTY);
                        string h = fmLogin.LoginID.ToString();

                        SUNN("2",h, QTY);
               
                        DELETEFILE();
                        MessageBox.Show("寄信成功");
                    }
                }
            }
            catch (Exception ex)
            {
                DELETEFILE();
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSeqDelivery_Click(object sender, EventArgs e)
        {
            int SeqDelivery = GetMaxSeqDelivery() + 1;//最大值+1
            foreach (DataGridViewRow row in wH_ItemDataGridView.Rows) 
            {
                if (Convert.ToBoolean(row.Cells["ColWhItemCheckBox"].Value) == true) 
                {
                    row.Cells["SeqDelivery"].Value = SeqDelivery;
                    row.Cells["ColWhItemCheckBox"].Value = false;
                }
            }
        }
        private int GetMaxSeqDelivery() 
        {
            int i = 0;
            int j = 0;
            foreach (DataGridViewRow row in wH_ItemDataGridView.Rows)
            {
                if (int.TryParse(Convert.ToString(row.Cells["SeqDelivery"].Value),out j))
                {
                    if (i < j) 
                    {
                        i = j;//找最大值回傳
                    }
                }
            }
            return i;
        }

        private void SUNN(string DOCTYPE,string h, string QTYY)
        {

            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\wh2.htm";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();



            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);



            try
            {
                System.Data.DataTable dt1 = GetMenu.Getemployee(h);
                DataRow drw = dt1.Rows[0];
                if ((dt1.Rows.Count) > 0)
                {
                    string a1 = drw["pager"].ToString();
                    string a2 = drw["mobile"].ToString();
                    string officeext = drw["officeext"].ToString();
                    template = template.Replace("##eng##", a2);
                    template = template.Replace("##name##", a1);
                    template = template.Replace("##officeext##", officeext);
                    template = template.Replace("##mail##", h + "@acmepoint.com");
                }
            }
            catch
            {
                template = template.Replace("##eng##", "");
                template = template.Replace("##name##", "");
                template = template.Replace("##officeext##", "");
                template = template.Replace("##mail##", h + "@acmepoint.com");
            }
         
            MailMessage message = new MailMessage();
            if (DOCTYPE == "1")
            {
                message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            }
            if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
            {
                string[] arrurl = s5TextBox.Text.Split(new Char[] { ';' });

                foreach (string i in arrurl)
                {
                    message.To.Add(i.ToString());
                }
            }
            message.To.Add(LOGINID);
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable dt = Getwh(shippingCodeTextBox.Text);
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                DataRow dd = dt.Rows[i];


                sb.Append(dd["docentry"].ToString() + "/");


            }

            sb.Remove(sb.Length - 1, 1);
            fg = sb.ToString();
            string OWHS = shipping_OBUTextBox.Text;
            int LEN = OWHS.Length;

            string OWHS1 = OWHS.Trim().Replace("倉", "").Replace("-", "");

            string MM = "";
            if (forecastDayTextBox.Text == "庫存調撥-借出")
            {
                MM = "借出-國內-";
            }
            if (forecastDayTextBox.Text == "庫存調撥-借出還回")
            {
                MM = "借出還回-國內-";
            }
            if (forecastDayTextBox.Text == "庫存調撥-撥倉")
            {
                System.Data.DataTable GH1 = GetOWTR(pINOTextBox.Text);

                string OWHS2 = OWHS.Trim().Replace("-", "") + "調撥回" + GH1.Rows[0][0].ToString().Replace("倉", "");
                MM = OWHS2 + "-";
            }
            if (s1CheckBox.Checked)
            {

                if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                {

                    message.Subject = shipping_OBUTextBox.Text + "--備貨單通知--"  + cardNameTextBox.Text +"-"+MM + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                }
                else
                {
                    message.Subject = shipping_OBUTextBox.Text + "--備貨單通知--國內-" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                }
                //在途倉備貨單--國內—客戶名—工單號—數量PCS—日期
                if (OWHS1 == "在途" || OWHS1 == "平鎮")
                {
                    message.Subject =  shipping_OBUTextBox.Text + "備貨單--國內—" + cardNameTextBox.Text + "—" + shippingCodeTextBox.Text + "—" + QTYY + "PCS—" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                    s4TextBox.Text = "";
                    if (OWHS1 == "在途")
                    {
                        string gh = "Dear Maggie / Vivi/Lulu" +

                              Environment.NewLine +
           Environment.NewLine + "麻煩請安排" + quantityTextBox.Text + "提貨直送" + cardNameTextBox.Text + "，謝謝！";
                        s4TextBox.Text = gh;
                    }

                    if (OWHS1 == "平鎮")
                    {
                        //quantityTextBox
                       // string gh = "請安排" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "快遞 謝謝";
                        string gh = "請安排" + quantityTextBox.Text + "快遞 謝謝";
                        s4TextBox.Text = gh;
                    }
                }
            }
            if (s2CheckBox.Checked)
            {
                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }
                if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                {

                    message.Subject = shipping_OBUTextBox.Text + "--地址條+放貨單通知--" + cardNameTextBox.Text + "-" + MM + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                }
                else
                {
                    message.Subject = shipping_OBUTextBox.Text + "--地址條+放貨單通知--" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
                }
              
            }
            if (s3CheckBox.Checked)
            {
                if (wH_Item2DataGridView.Rows.Count == 1)
                {

                    MessageBox.Show("請輸入放貨單");
                    return;
                }
                message.Subject = shipping_OBUTextBox.Text + "--放貨單通知--" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            }

            if (s6CheckBox.Checked)
            {
                string SUB1 = "";
       
                if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                {
                    SUB1 = shipping_OBUTextBox.Text + "--備貨單+地址條+放貨單通知--" + cardNameTextBox.Text + "-" + MM + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
   
                }
                else
                {
                    SUB1 = shipping_OBUTextBox.Text + "--備貨單+地址條+放貨單通知--" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
          
                }
                message.Subject = SUB1;

            }

            if (s7CheckBox.Checked)
            {
                string SUB1 = "";

                if (forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                {
                    SUB1 = shipping_OBUTextBox.Text + "--備貨單+放貨單通知--" + cardNameTextBox.Text + "-" + MM + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");

                }
                else
                {
                    SUB1 = shipping_OBUTextBox.Text + "--備貨單+放貨單通知--" + cardNameTextBox.Text + "--" + shippingCodeTextBox.Text + "--" + QTYY + "PCS--" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");

                }
                message.Subject = SUB1;

            }

            template = template.Replace("##Content##", s4TextBox.Text.Replace(System.Environment.NewLine, "<br>"));


            message.Body = template;
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98" || globals.DBNAME == "宇豐" || globals.DBNAME == "INFINITE")
            {
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {

                    string m_File = "";

                    m_File = file;
                    data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                    //附件资料
                    ContentDisposition disposition = data.ContentDisposition;


                    // 加入邮件附件
                    message.Attachments.Add(data);


                }



            }


            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Send(message);
            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98" || globals.DBNAME == "宇豐" || globals.DBNAME == "INFINITE")
            {
                data.Dispose();
                message.Attachments.Dispose();
            }
        }

        private void da()
        {
            StringBuilder sb = new StringBuilder();
            if (tabControl1.SelectedTab == tabPage6)
            {
                if (s5TextBox.Text == "")
                {

                    if (forecastDayTextBox.Text == "銷售訂單")
                    {
                        System.Data.DataTable G1 = GetMenu.GetSA(pINOTextBox.Text.Trim());
                        if (G1.Rows.Count > 0)
                        {

                            string SA信箱 = G1.Rows[0]["SA信箱"].ToString();


                            if (!String.IsNullOrEmpty(SA信箱))
                            {
                                sb.Append(SA信箱 + ";");
                            }
                        }
                    }

                    if ( forecastDayTextBox.Text == "庫存調撥-借出" || forecastDayTextBox.Text == "庫存調撥-借出還回" || forecastDayTextBox.Text == "庫存調撥-撥倉")
                    {
                        System.Data.DataTable G1 = GetMenu.GetSA2(shippingCodeTextBox.Text);
                        if (G1.Rows.Count > 0)
                        {

                            string SA信箱 = G1.Rows[0][1].ToString();

                            if (!String.IsNullOrEmpty(SA信箱))
                            {
                                sb.Append(SA信箱 + ";");
                            }
                 
                        }
                    }
                    if (shipping_OBUTextBox.Text == "新得利倉")
                    {
                        sb.Append("syang.dejye@msa.hinet.net;");
                        sb.Append("sdl.tw@msa.hinet.net;");

                    }
                    if (shipping_OBUTextBox.Text == "聯揚倉")
                    {
                        sb.Append("8248@mail.lcwebs.com.tw;");
                        sb.Append("8329@mail.lcwebs.com.tw;");
                    }
                 
                    if (shipping_OBUTextBox.Text == "在途倉")
                    {
                        sb.Append("viviweng@acmepoint.com;");
                        sb.Append("maggieweng@acmepoint.com;");
                        sb.Append("luluhsieh@acmepoint.com;");
                        sb.Append("mimichen@acmepoint.com;");
                    }
                    if (shipping_OBUTextBox.Text == "平鎮倉")
                    {
                        sb.Append("syoutali@aresopto.com;");
                        sb.Append("nicelin@aresopto.com;");
                        sb.Append("sylviashih@aresopto.com;");
                        sb.Append("thomashung@aresopto.com;");
                        sb.Append("timmychou@aresopto.com;");
                    }

                    sb.Append("davidhuang@acmepoint.com;");
                    sb.Append("bettytseng@acmepoint.com;");
                    sb.Append("SunnyWang@acmepoint.com;");
                    sb.Append("jingdong@acmepoint.com;");
                    if (sb.Length > 1)
                    {
                        sb.Remove(sb.Length - 1, 1);

                        s5TextBox.Text = sb.ToString();
                    }
                    else
                    {
                        s5TextBox.Text = "";
                    }
                }
            }
        }

        private void s4TextBox_Click(object sender, EventArgs e)
        {
            if (s4TextBox.Text == "")
            {
                string DATE = quantityTextBox.Text.Substring(5, 5);
                string gh = "" + DATE + "放貨單及地址條," +
                       Environment.NewLine + "請安排" + DATE + "下午快,謝謝~~" +
                                          Environment.NewLine +
                       Environment.NewLine + "Dear蕭小姐:" +
                       Environment.NewLine + "請安排" + DATE + "早上11點前送達，謝謝~~" +
                       Environment.NewLine + "請注意客戶中午12點-13點不收貨" +
                       Environment.NewLine + "若超過4.5頓請安排下午3點前送達(請勿大於8.8頓)" +
                       Environment.NewLine +
                       Environment.NewLine + "Dear蕭小姐:" +
                       Environment.NewLine + "請司機記得帶放貨單+進料驗收單給客戶,謝謝~" +
                       Environment.NewLine + "請安排" + DATE + "早上10點前送達，謝謝~~" +
                       Environment.NewLine + "請注意客戶中午12點-13點不收貨" +
                                 Environment.NewLine +
                       Environment.NewLine + "Dear夏先生/陳先生:" +
                       Environment.NewLine + "請單箱派0.6T車,安排" + DATE + "早上10點前送達，謝謝~~" +
                       Environment.NewLine + "請安排" + DATE + "早上10點前送達，謝謝~~" +
                       Environment.NewLine + "請注意客戶中午12點-13點不收貨" +
                                                    Environment.NewLine +
                       Environment.NewLine + "Dear夏先生/陳先生:" +
                       Environment.NewLine + "此票請併車至高雄未稅NTD:978 安排" + DATE + "下午5點前到貨,謝謝~~" +
                       Environment.NewLine + "請注意客戶中午12點-13點不收貨";
                s4TextBox.Text = gh;
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = wh.WH_Item4;
            DataRow newCustomersRow = dt2.NewRow();
            int i = wH_Item4DataGridView.CurrentRow.Index;

            DataRow drw = dt2.Rows[i];
            string sa = drw["shippingcode"].ToString();
            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["Docentry"] = drw["Docentry"];
            newCustomersRow["linenum"] = drw["linenum"];
            newCustomersRow["ItemRemark"] = drw["ItemRemark"];
            newCustomersRow["ItemCode"] = drw["ItemCode"];
            newCustomersRow["Dscription"] = drw["Dscription"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Remark"] = drw["Remark"];
            newCustomersRow["INV"] = drw["INV"];
            newCustomersRow["PiNo"] = drw["PiNo"];
            newCustomersRow["NowQty"] = drw["NowQty"];
            newCustomersRow["Ver"] = drw["Ver"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["Invoice"] = drw["Invoice"];
            newCustomersRow["FrgnName"] = drw["FrgnName"];
            newCustomersRow["WHName"] = drw["WHName"];
            newCustomersRow["Shipdate"] = drw["Shipdate"];
            newCustomersRow["CardCode"] = drw["CardCode"];
            newCustomersRow["U_PAY"] = drw["U_PAY"];
            newCustomersRow["U_SHIPDAY"] = drw["U_SHIPDAY"];
            newCustomersRow["U_SHIPSTATUS"] = drw["U_SHIPSTATUS"];
            newCustomersRow["U_MARK"] = drw["U_MARK"];
            newCustomersRow["U_MEMO"] = drw["U_MEMO"];
            newCustomersRow["PO"] = drw["PO"];
            newCustomersRow["LOCATION"] = drw["LOCATION"];
            newCustomersRow["TREETYPE"] = drw["TREETYPE"];
            newCustomersRow["U_CUSTITEMCODE"] = drw["U_CUSTITEMCODE"];
            newCustomersRow["U_CUSTDOCENTRY"] = drw["U_CUSTDOCENTRY"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, wH_Item4DataGridView.CurrentRow.Index);

                for (int j = 0; j <= wH_Item4DataGridView.Rows.Count - 2; j++)
                {
                    wH_Item4DataGridView.Rows[j].Cells[0].Value = (j + 1).ToString();
                }

                this.wH_Item4BindingSource.EndEdit();
                this.wH_Item4TableAdapter.Update(wh.WH_Item4);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
    }
}


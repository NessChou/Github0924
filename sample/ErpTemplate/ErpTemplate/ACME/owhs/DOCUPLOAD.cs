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
using System.Collections;
using System.Net.Mime;
namespace ACME
{
    public partial class DOCUPLOAD : Form
    {
        string WAR = "";
        int s1 = 0;
        int inint = 0;
        int ouint = 0;
        string DOC = "";
        string CARD = "";
        string intname = "";
        string ountname = "";
        string DIR;
        string PATH;
        string OBJ = "";
        string FF = "";
        string strCn = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public DOCUPLOAD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                 if (GetMenu.Getdata("SAP").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "SAP");
                }
                else
                {
                    GetMenu.UP(t1, "SAP");
                }

                textBox1.Text = t1;
            }
        }



        private void DOCUPLOAD_Load(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GetMenu.Getdata("SAP");
            if (G1.Rows.Count > 0)
            {
                textBox1.Text = G1.Rows[0][0].ToString();
            }

            System.Data.DataTable G2 = GetMenu.Getdata("WHPACK2");
            if (G2.Rows.Count > 0)
            {
                textBox2.Text = G2.Rows[0][0].ToString();
            }
            System.Data.DataTable G3 = GetMenu.Getdata("NESS1");
            if (G3.Rows.Count > 0)
            {
                textBox3.Text = G3.Rows[0][0].ToString();
            }
            DIR = "//acmesrv01//SAP_Share//shipping//";
            PATH = @"\\acmesrv01\SAP_Share\shipping\";

            if (globals.DBNAME == "達睿生")
            {
                DIR = "//acmesrv01//SAP_Share//shipping達睿生//";
                PATH = @"\\acmesrv01\SAP_Share\shipping達睿生\";
            }

            if (globals.DBNAME == "禾中")
            {
                groupBox1.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                inint = 0;
                s1 = 0;
                string d = textBox1.Text;

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {

                    FF = file;
                    FileInfo info = new FileInfo(file);
                    int IN1 = info.Name.ToString().ToUpper().IndexOf("DRS");
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");
                    if (globals.DBNAME == "達睿生")
                    {
                        if (IN1 == -1)
                        {
                            MessageBox.Show(info.Name.ToString()+" 檔名沒有DRS無法上傳");
                            return;
                        }
                    }
                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");
                        string M1 = NAME.Substring(J1 + 1, NAME.Length - J1 - 1);
                        string M2 = NAME.Substring(0, J1);
                        string DOC = NAME.Substring(1, J1 - 1);
                        string DOC2 = NAME.Substring(0, 1);
                  

                        int f1 = DOC.IndexOf("-");
                        int n;

                        if (f1 != -1)
                        {
                            DOC = NAME.Substring(1, f1);
                        
                        }


                        int F2 = NAME.IndexOf("INV-");
                        if (F2 != -1)
                        {
                            string DD = NAME.Replace("INV-", "");
                            int F3 = DD.IndexOf("-");
                            DOC = DD.Substring(0, F3);

                        }

                        if (globals.DBNAME == "達睿生")
                        {
                            DOC2 = NAME.Substring(3, 1);

                            if (f1 != -1)
                            {
                                DOC = NAME.Substring(4, f1 - 3);

                            }
                        }
                        string DOCENTRY = "";

                        if (DOC2 == "放" || DOC2 == "序" || DOC2 == "備" || DOC2 == "簽" || DOC2 == "备" || DOC2 == "签" || DOC2 == "I")
                        {
                            OBJ = "15";
                        }
                        else if (DOC2 == "進" || DOC2 == "进" ||　DOC2 == "收" )
                        {
                            OBJ = "20";
                        }
                        else if (DOC2 == "調" || DOC2 == "借")
                        {
                            OBJ = "67";
                        }
                        else if (DOC2 == "生")
                        {
                            OBJ = "202";
                        }
                        string CARDCODE = "";
                        string TEL1 = "";
                        string CNTCTCODE = "";
                        string SLPCODE = "";

                        int h1 = DOC.IndexOf("_");
                        if (h1 != -1)
                        {
                            DOC = DOC.Substring(0, DOC.IndexOf("_"));
                        
                        }


                        if (int.TryParse(DOC, out n))
                        {
                            DOCENTRY = DOC;
                            System.Data.DataTable T1 = GetODLN("10000000");
                            if (OBJ == "15")
                            {
                                T1 = GetODLN(DOC);
                            }
                            if (OBJ == "20")
                            {
                                T1 = GetOPDN(DOC);
                            }
                            if (OBJ == "202")
                            {
                                T1 = GetOWOR(DOC);
                            }
                            if (T1.Rows.Count > 0 || OBJ == "67")
                            {
                                System.Data.DataTable T2 = GetMAXOCLG2(OBJ, DOC);

                                if (OBJ == "67")
                                {
                                    CARDCODE = "0001-00";
                                    TEL1 = "03-407-8800";
                                    CNTCTCODE = "1282";
                                    SLPCODE = "-1";
                                }
                                else
                                {
                                    DataRow dd = T1.Rows[0];
                                    CARDCODE = dd["CARDCODE"].ToString();
                                    TEL1 = dd["TEL1"].ToString();
                                    CNTCTCODE = dd["CNTCTCODE"].ToString();
                                    SLPCODE = dd["SLPCODE"].ToString();
                                }


                                int d1 = Convert.ToInt32(GetMAXOCLG().Rows[0][0].ToString());
                                int m2 = Convert.ToInt32(GetMAXOATC().Rows[0][0].ToString());
                                DateTime now = DateTime.Now;
                                int d2 = Convert.ToInt16(DateTime.Now.ToString("HHmm"));

                                string ATT = textBox1.Text;
                                string UATT = textBox1.Text + @"\" + NAME;
                                string ATT2 = @"\\ACMEW08R2AP\SAPFILES2\Attachments" + DateTime.Now.ToString("yyyy") + "\\ATT" + DateTime.Now.ToString("yyyyMM");
                                bool FF1 = getrma.UploadFile(UATT, ATT2, false);
                                if (FF1 == false)
                                {
                                    return;
                                }
                                if (T2.Rows.Count == 0)
                                {
                                    AddOCLG(d1, CARDCODE, now, d2, now, "N", TEL1, -1, "N", OBJ, DOC, DOC, ATT, "l", 1,
                                        Convert.ToInt32(CNTCTCODE), 1, Convert.ToInt32(SLPCODE), "C", -1, d2, "M", "1", "N", 15, "M", "N", 0, "N", "N", "N", m2, now.AddDays(1), d2);
                                    AddOACT(m2);
                                    AddATC1(m2, 1, ATT, ATT2, M2, M1, now, 1, "Y", "Y");
                                    UPONNM(d1 + 1, "33");
                                    UPONNM(m2 + 1, "221");
                                    if (OBJ == "20")
                                    {
                                        UPOPDN(m2, DOC);
                                    }
                                }
                                else
                                {
                                    DataRow dd2 = T2.Rows[0];
                                    string ATCENTRY = dd2["ATCENTRY"].ToString();
                                    if (String.IsNullOrEmpty(ATCENTRY))
                                    {
                                        AddOACT(m2);
                                        AddATC1(m2, 1, ATT, ATT2, M2, M1, now, 1, "Y", "Y");
                                        UPOCLG(m2, OBJ, DOC);
                                        UPONNM(m2 + 1, "221");
                                        if (OBJ == "20")
                                        {
                                            UPOPDN(m2, DOC);
                                        }
                             
                                    }
                                    else
                                    {

                                        System.Data.DataTable H1 = GetATC1S(ATCENTRY);
                                        int g1 = 0;
                                        if (H1.Rows.Count > 0)
                                        {
                                            g1 = Convert.ToInt32(GetATC1(ATCENTRY).Rows[0][0].ToString());

                                        }
                                        else
                                        {
                                            g1 = 1;
                                        }

                                        AddATC1(Convert.ToInt32(ATCENTRY), g1, ATT, ATT2, M2, M1, now, 1, "Y", "Y");
                                        if (OBJ == "20")
                                        {
                                            UPOPDN(Convert.ToInt32(ATCENTRY), DOC);
                                        }
                                    }
                                }
                             
    
                                System.GC.Collect();
                                System.GC.WaitForPendingFinalizers();
                                File.Delete(file);

                                s1 = 1;
                            }
                       

                        }
               

                    }
                }

                if (s1 == 1)
                {
                    MessageBox.Show("上傳成功");
                }
                else
                {
                    MessageBox.Show("沒有檔案匯入");
                }

            }
            catch (Exception ex)
            {

                //MAILFILE(intname + ex.Message.ToString());
                MessageBox.Show(FF + ex.Message);
            }





        }
        public void UPONNM(int AUTOKEY, string ObjectCode)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE ONNM SET AUTOKEY=@AUTOKEY WHERE ObjectCode=@ObjectCode", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AUTOKEY", AUTOKEY));
            command.Parameters.Add(new SqlParameter("@ObjectCode", ObjectCode));

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
        public void UPOPDN(int AtcEntry, string DocEntry)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("update OPDN set AtcEntry=@AtcEntry  where DocEntry=@DocEntry", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AtcEntry", AtcEntry));
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));

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


        public void UPOCLG(int atcentry, string DOCTYPE, string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("UPDATE OCLG SET atcentry=@atcentry where DOCTYPE=@DOCTYPE AND DOCENTRY=@DOCENTRY ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@atcentry", atcentry));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable GetMAXOCLG()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(CLGCODE)+1 ID FROM OCLG ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public void AddOACT(int AbsEntry)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into OATC(AbsEntry) values(@AbsEntry)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));

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
        public void AddOCLG(int ClgCode, string CardCode, DateTime CntctDate, int CntctTime, DateTime Recontact, string Closed, string Tel, int CntctSbjct, string Transfered, string DocType, string DocNum, string DocEntry, string Attachment, string DataSource, int AttendUser, int CntctCode, int UserSign, int SlpCode, string Action, int CntctType, int BeginTime, string DurType, string Priority, string Reminder, int RemQty, string RemType, string RemSented, int Instance, string personal, string inactive, string tentative, int AtcEntry, DateTime endDate, int ENDTime)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into OCLG(ClgCode,CardCode,CntctDate,CntctTime,Recontact,Closed,Tel,CntctSbjct,Transfered,DocType,DocNum,DocEntry,Attachment,DataSource,AttendUser,CntctCode,UserSign,SlpCode,Action,CntctType,BeginTime,DurType,Priority,Reminder,RemQty,RemType,RemSented,Instance,personal,inactive,tentative,AtcEntry,duration,endDate,ENDTime) values(@ClgCode,@CardCode,@CntctDate,@CntctTime,@Recontact,@Closed,@Tel,@CntctSbjct,@Transfered,@DocType,@DocNum,@DocEntry,@Attachment,@DataSource,@AttendUser,@CntctCode,@UserSign,@SlpCode,@Action,@CntctType,@BeginTime,@DurType,@Priority,@Reminder,@RemQty,@RemType,@RemSented,@Instance,@personal,@inactive,@tentative,@AtcEntry,@duration,@endDate,@ENDTime)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ClgCode", ClgCode));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@CntctDate", CntctDate));
            command.Parameters.Add(new SqlParameter("@CntctTime", CntctTime));
            command.Parameters.Add(new SqlParameter("@Recontact", Recontact));
            command.Parameters.Add(new SqlParameter("@Closed", Closed));
            command.Parameters.Add(new SqlParameter("@Tel", Tel));
            command.Parameters.Add(new SqlParameter("@CntctSbjct", CntctSbjct));
            command.Parameters.Add(new SqlParameter("@Transfered", Transfered));
            command.Parameters.Add(new SqlParameter("@DocType", DocType));
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@Attachment", Attachment));
            command.Parameters.Add(new SqlParameter("@DataSource", DataSource));
            command.Parameters.Add(new SqlParameter("@AttendUser", AttendUser));
            command.Parameters.Add(new SqlParameter("@CntctCode", CntctCode));
            command.Parameters.Add(new SqlParameter("@UserSign", UserSign));
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            command.Parameters.Add(new SqlParameter("@Action", Action));
            command.Parameters.Add(new SqlParameter("@CntctType", CntctType));
            command.Parameters.Add(new SqlParameter("@BeginTime", BeginTime));
            command.Parameters.Add(new SqlParameter("@DurType", DurType));
            command.Parameters.Add(new SqlParameter("@Priority", Priority));
            command.Parameters.Add(new SqlParameter("@Reminder", Reminder));
            command.Parameters.Add(new SqlParameter("@RemQty", RemQty));
            command.Parameters.Add(new SqlParameter("@RemType", RemType));
            command.Parameters.Add(new SqlParameter("@RemSented", RemSented));
            command.Parameters.Add(new SqlParameter("@Instance", Instance));
            command.Parameters.Add(new SqlParameter("@personal", personal));
            command.Parameters.Add(new SqlParameter("@inactive", inactive));
            command.Parameters.Add(new SqlParameter("@tentative", tentative));
            command.Parameters.Add(new SqlParameter("@AtcEntry", AtcEntry));
            command.Parameters.Add(new SqlParameter("@duration", 1.000000));
            command.Parameters.Add(new SqlParameter("@endDate", endDate));
            command.Parameters.Add(new SqlParameter("@ENDTime", ENDTime));

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
        public void AddATC1(int AbsEntry, int Line, string srcPath, string trgtPath, string FileName, string FileExt, DateTime Date, int UsrID, string Copied, string Override)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into ATC1(AbsEntry,Line,srcPath,trgtPath,FileName,FileExt,Date,UsrID,Copied,Override) values(@AbsEntry,@Line,@srcPath,@trgtPath,@FileName,@FileExt,@Date,@UsrID,@Copied,@Override)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));
            command.Parameters.Add(new SqlParameter("@Line", Line));
            command.Parameters.Add(new SqlParameter("@srcPath", srcPath));
            command.Parameters.Add(new SqlParameter("@trgtPath", trgtPath));
            command.Parameters.Add(new SqlParameter("@FileName", FileName));
            command.Parameters.Add(new SqlParameter("@FileExt", FileExt));
            command.Parameters.Add(new SqlParameter("@Date", Date));
            command.Parameters.Add(new SqlParameter("@UsrID", UsrID));
            command.Parameters.Add(new SqlParameter("@Copied", Copied));
            command.Parameters.Add(new SqlParameter("@Override", Override));
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

        public void UPLOADOPDN(int AbsEntry, int Line, string srcPath, string trgtPath, string FileName, string FileExt, DateTime Date, int UsrID, string Copied, string Override)
        {
            SqlConnection connection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into ATC1(AbsEntry,Line,srcPath,trgtPath,FileName,FileExt,Date,UsrID,Copied,Override) values(@AbsEntry,@Line,@srcPath,@trgtPath,@FileName,@FileExt,@Date,@UsrID,@Copied,@Override)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));
            command.Parameters.Add(new SqlParameter("@Line", Line));
            command.Parameters.Add(new SqlParameter("@srcPath", srcPath));
            command.Parameters.Add(new SqlParameter("@trgtPath", trgtPath));
            command.Parameters.Add(new SqlParameter("@FileName", FileName));
            command.Parameters.Add(new SqlParameter("@FileExt", FileExt));
            command.Parameters.Add(new SqlParameter("@Date", Date));
            command.Parameters.Add(new SqlParameter("@UsrID", UsrID));
            command.Parameters.Add(new SqlParameter("@Copied", Copied));
            command.Parameters.Add(new SqlParameter("@Override", Override));
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

        public void AddADOWNLOAD(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into Download(shippingcode,seq,filename,path) values(@shippingcode,@seq,@filename,@path)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));

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
        public void AddADOWNLOAD2(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into Download2(shippingcode,seq,filename,path) values(@shippingcode,@seq,@filename,@path)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));

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
        private System.Data.DataTable GetODLN(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CARDCODE,TEL1,ISNULL(T2.CNTCTCODE,0) CNTCTCODE,SLPCODE FROM ODLN T0 ");
            sb.Append(" LEFT JOIN OCPR T2 ON (T0.CNTCTCODE=T2.CNTCTCODE)");
            sb.Append(" WHERE T0.DOCENTRY=@DOC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETF1()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT DOCENTRY,AtcEntry FROM OPDN WHERE DOCENTRY > 48167 AND ISNULL(AtcEntry,'') <> '' ");
            sb.Append(" SELECT DOCENTRY FROM OPDN WHERE  ISNULL(AtcEntry,'') ='' AND DOCENTRY >40000 AND YEAR(DOCDATE)=2021");
            sb.Append(" AND CAST(DOCENTRY AS VARCHAR) IN (");
            sb.Append(" SELECT DOCENTRY FROM OCLG WHERE DocType ='20'");
            sb.Append(" )");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETF2(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT ABSENTRY, SUBSTRING(FILENAME, 2, 5), FILENAME FROM ATC1 WHERE SUBSTRING(FILENAME, 1, 1) = '進' AND SUBSTRING(FILENAME, 2, 5) = @DOC ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETF3(string AbsEntry)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM OATC WHERE AbsEntry =@AbsEntry ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AbsEntry", AbsEntry));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOPDN(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CARDCODE,TEL1,ISNULL(T2.CNTCTCODE,0) CNTCTCODE,SLPCODE FROM OPDN T0 ");
            sb.Append(" LEFT JOIN OCPR T2 ON (T0.CNTCTCODE=T2.CNTCTCODE)");
            sb.Append(" WHERE T0.DOCENTRY=@DOC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOWOR(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT");
            sb.Append(" CASE WHEN ISNULL(T0.CARDCODE,'')='' THEN '0001-00' ELSE  T0.CARDCODE END CARDCODE,");
            sb.Append(" CASE WHEN ISNULL(T0.CARDCODE,'')='' THEN '03-407-8800' ELSE　T2.TEL1 END TEL1,");
            sb.Append(" CASE WHEN ISNULL(T0.CARDCODE,'')='' THEN '1282' ELSE　ISNULL(T2.CNTCTCODE,0)  END CNTCTCODE,");
            sb.Append(" CASE WHEN ISNULL(T0.CARDCODE,'')='' THEN '-1' ELSE　 T1.SLPCODE  END SLPCODE");
            sb.Append(" FROM OWOR T0  ");
            sb.Append(" LEFT JOIN OCRD T1 ON (T0.CARDCODE=T1.CARDCODE)");
            sb.Append(" LEFT JOIN OCPR T2 ON (T1.CntctPrsn =T2.NAME) ");
            sb.Append(" WHERE T0.DOCENTRY=@DOC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOC", DOC));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void MAILFILE(string aa)
        {


            MailMessage message = new MailMessage();


            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");

            message.To.Add(new MailAddress("lleytonchen@acmepoint.com"));

            message.Subject = aa;


            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            try
            {
                client.Send(message);



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }

        private System.Data.DataTable GetWH_PACK2(string SHIPPINGCODE, string FLAG)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND SER <>'0' AND SER <>'' AND QTY <>'空箱'  ");

            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetWH_PACK2S(string SHIPPINGCODE, string FLAG)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND SER2 <>'0' AND SER2 <>'' AND QTY <>'空箱'  ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSHIPPINGCODE(string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM AcmeSqlSPCHOICE.DBO.WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT SHIPPINGCODE FROM AcmeSqlSPINFINITE.DBO.WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT SHIPPINGCODE FROM AcmeSqlSP.DBO.WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT SHIPPINGCODE FROM AcmeSqlSPDRS.DBO.WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetWH_PACK2SB(string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND QTY <>'空箱'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetWH_PACK2SB2(string SHIPPINGCODE,string ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT   ISNULL((SUM(CAST(GW AS DECIMAL(10,3)))),0) FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND QTY <>'空箱' AND ID <> @ID    ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetWH_PACK2N(string SHIPPINGCODE, string FLAG)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE   AND QTY <>'空箱'  ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    
        private System.Data.DataTable GetMAXOCLG2(string DOCTYPE, string DOCENTRY)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM OCLG WHERE DOCTYPE=@DOCTYPE AND DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAXOATC()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(ABSENTRY)+1 ID FROM OATC ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetATC1S(string ABSENTRY)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ABSENTRY FROM ATC1 WHERE ABSENTRY=@ABSENTRY   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ABSENTRY", ABSENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetATC1(string ABSENTRY)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT max(LINE)+1 FROM ATC1 WHERE ABSENTRY=@ABSENTRY   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ABSENTRY", ABSENTRY));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAXID(string ShippingCode, string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(ID) ID FROM WH_PACK2 WHERE ShippingCode =@ShippingCode   ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETOITM(string ITEMCODE)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ITEMCODE FROM OITM WHERE ITEMCODE=@ITEMCODE");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT PRODID COLLATE  Chinese_Taiwan_Stroke_CI_AS　FROM　otherDB.CHIComp21.DBO.comProduct　 WHERE PRODID=@ITEMCODE");
            sb.Append("  UNION ALL");
            sb.Append("  SELECT PRODID COLLATE  Chinese_Taiwan_Stroke_CI_AS　FROM　otherDB.CHIComp22.DBO.comProduct　 WHERE PRODID=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETOITM2(string PARAM_NO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT PARAM_DESC   FROM PARAMS WHERE PARAM_KIND ='WHLOCATION' AND PARAM_NO =@PARAM_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1(string ShippingCode, string SER, string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(cast(GW as decimal(10,2)))  GW,SUM(CAST(CARTONQTY AS INT)) CARTONQTY,SER FROM WH_PACK2 WHERE ShippingCode =@ShippingCode AND SER =@SER    ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }

            sb.Append("  GROUP BY SER   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1FS(string ShippingCode, string SER,string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(NW)  NW,SUM(CAST(CARTONQTY AS INT)) CARTONQTY,SER FROM WH_PACK2 WHERE ShippingCode =@ShippingCode AND SER =@SER AND ITEMCODE NOT LIKE '%ACME%'     ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }

            sb.Append("  GROUP BY SER   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1FS2(string ShippingCode, string SER,string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT *  FROM WH_PACK2 WHERE ShippingCode =@ShippingCode AND SER =@SER  AND NW <>''  ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1S(string ShippingCode, string SER, string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(ID) ID,MAX(cast(GW as decimal(10,2)))  GW,SUM(CAST(CARTONQTY AS INT)) CARTONQTY,SER2,SUM(CAST(NW AS DECIMAL(10,2))) NW   FROM WH_PACK2 WHERE ShippingCode =@ShippingCode AND SER2 =@SER2     ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            sb.Append("   GROUP BY SER2   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SER2", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1S2(string ShippingCode, string SER, string FLAG,string ID)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(cast(GW as decimal(10,2))) GW FROM WH_PACK2 WHERE ShippingCode =@ShippingCode AND SER2 =@SER2  AND ID <> @ID     ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            sb.Append("   GROUP BY SER2   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@SER2", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1SB(string SHIPPINGCODE,decimal GAN)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (SUM(CAST(GW AS DECIMAL(10,3)))+@GAN) FW,SUM(CAST(NW AS DECIMAL(10,3))) FNW FROM WH_PACK2  WHERE SHIPPINGCODE=@SHIPPINGCODE  AND QTY <> '空箱'    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@GAN", GAN));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


    
        private System.Data.DataTable GetPACK1F(string SHIPPINGCODE,string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(MAX(CARTONNO2) AS DECIMAL(10,2)) -SUM(CAST(GW AS decimal(10,2))) GW,SER  FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(SER,0) <> 0 ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            sb.Append("  GROUP BY SER");
            sb.Append("  HAVING SUM(CAST(GW AS decimal(10,2))) <> CAST(MAX(CARTONNO2) AS DECIMAL(10,2)) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetPACK1F2(string SHIPPINGCODE, string SER, string FLAG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT TOP 1 ID,GW    FROM WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND SER=@SER ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDOWNSEQ(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT COUNT(*) SEQ FROM Download WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDOWNSEQ2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT COUNT(*) SEQ FROM Download2 WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDOWNSEQS(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOUT1(string MODEL,string GRADE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM CART WHERE MODEL_NO =@MODEL AND TMEMO=@GRADE ORDER BY UPDATE_DATE DESC  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("WHPACK2").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "WHPACK2");
                }
                else
                {
                    GetMenu.UP(t1, "WHPACK2");
                }

                textBox2.Text = t1;
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("NESS1").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "NESS1");
                }
                else
                {
                    GetMenu.UP(t1, "NESS1");
                }

                textBox3.Text = t1;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            WAR = "";
            string d = textBox2.Text;


            if (!String.IsNullOrEmpty(d))
            {
                string[] filenames = Directory.GetFiles(d);
                int M = 0;
                foreach (string file in filenames)
                {
                    FileInfo info = new FileInfo(file);
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");
                    string SH = NAME.Substring(0, 2);
                    string SH2 = NAME.Substring(0, 5);
                    M++;
                    if (SH == "WH")
                    {
                        WriteExcel(file, "三角", NAME);
                        File.Delete(file);
                    }

                    if (globals.DBNAME == "達睿生")
                    {
                        if (SH2 == "DRSWH")
                        {
                            WriteExcel(file, "三角", NAME);
                            File.Delete(file);
                        }
                    }
                }
                if (WAR == "")
                {
                    MessageBox.Show("上傳成功" + M.ToString());
                }

            }
            
        }
        public void WriteExcel(string ExcelFile, string FLAG,string FILE)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            int G1=FILE.LastIndexOf(".");
            string FILENAME = FILE.Substring(0, G1);
            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            int B = excelBook.Sheets.Count;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            if (iRowCnt > 500)
            {
                iRowCnt = 500;
            }


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string WHNO;
                string PLATENO;
                string PLATENO2N;
                string PLATENO2ND = "";
                string CARTONNO;
                string AUNO;
                string ITEMCODE = "";
                string ITEMCODED = "";
                string AUNOD = "";
                string GRADE;
                string ITEMNAME = "";

                string ITEMNAMED = "";
                string ITEMNAMED2 = "";
                string VER;
                string QTY;
                string CARTONQTY;
                string NW;
                string GW;
                string L;
                string W;
                string H;
                string MATERIAL;
                string LOACTION;
                int QTY2 = 0;
                int QTY2D = 0;
                int CARTONQTY2 = 0;
                int CARTONQTY2D = 0;
                string BLC;
                string GW2 = "";
                string SER = "";
                string MATERIAL2 = "";
                string LOACTION2 = "";
                string PLATENO2 = "";
                string GF = "";
                string GFN = "";
                string ITEMCODE2 = "";//特殊料號
                string ES = "";
                string CART = "";
                string CartonNoTemp = "";
                int IG1 = FILE.IndexOf("第");
                int IG2 = FILE.IndexOf("車");
                if (IG1 != -1 && IG2 != -1)
                {
                    CART = FILE.Substring(IG1 + 1, IG2 - IG1 - 1);
                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, 2]);
                range.Select();
                WHNO = range.Text.ToString().Trim();

                System.Data.DataTable SH1 = GetSHIPPINGCODE(FILENAME);
                if (SH1.Rows.Count > 0)
                {
                    if (WHNO != FILENAME)
                    {
                        MessageBox.Show("上傳失敗 工單號碼 檔名跟內文不一樣");
                        WAR = "1";
                        return;

                    }
                    WHNO = FILENAME;
                }

                DELETETA(WHNO, CART);
                decimal GAN = 0;

                int GGF = 0;
                int GGFN = 0;
                StringBuilder sb = new StringBuilder();
                for (int iRecord = 7; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    PLATENO = range.Text.ToString().Trim();
                    PLATENO = PLATENO.Replace("--", "-");

                    PLATENO2N = PLATENO;


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CARTONNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    BLC = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    AUNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    if (ITEMCODE.Length == 15)
                    {
                        System.Data.DataTable GF1S = GETOITM(ITEMCODE);
                        if (GF1S.Rows.Count == 0)
                        {
                            MessageBox.Show("上傳失敗 " + FILE + "料號 " + ITEMCODE + " SAP無此料號");
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                            range = null;
                            excelApp = null;
                            excelBook = null;
                            excelSheet = null;

                            System.GC.Collect();
                            System.GC.WaitForPendingFinalizers();

                            return;
                        }
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    ITEMNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    VER = range.Text.ToString().Trim().ToUpper().Replace("V.", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    QTY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    CARTONQTY = range.Text.ToString().Trim();


                    //if (FLAG == "出口")
                    //{
                    if (String.IsNullOrEmpty(PLATENO2N))
                    {
                        PLATENO2N = PLATENO2ND;
                    }

                    if (!String.IsNullOrEmpty(PLATENO))
                    {
                        PLATENO2ND = PLATENO;
                    }



                    if (!String.IsNullOrEmpty(CARTONNO))
                    {
                        QTY2 = 0;
                        CARTONQTY2 = 0;
                        Clear(sb);
                        CartonNoTemp = CARTONNO;
                    }
                    else 
                    {
                        //上下合併儲存格的時候,有值先記錄在CartonNoTemp沒值再取出
                        CARTONNO = CartonNoTemp;
                    }
                    int nn;
                    if (int.TryParse(QTY, out nn) && int.TryParse(CARTONQTY, out nn))
                    {

                        QTY2 += Convert.ToInt32(QTY);
                        CARTONQTY2 += Convert.ToInt32(CARTONQTY);

                        //if (ITEMCODED != ITEMCODE && !String.IsNullOrEmpty(ITEMCODED))
                        //{
                        //    QTY2 -= Convert.ToInt32(QTY2D);
                        //    CARTONQTY2 -= Convert.ToInt32(CARTONQTY2D);
                        //}


                    }
                    if (AUNOD != AUNO)
                    {
                        sb.Append(AUNO + "/");
                    }
                    // }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    NW = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    GW = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    L = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    W = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    H = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    MATERIAL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    LOACTION = range.Text.ToString().Trim();

                    if (LOACTION != "")
                    {
                        if (ITEMCODE.Length == 15)
                        {

                            string R1 = ITEMCODE.Substring(14, 1);
                            int ITEMS = ITEMCODE.ToUpper().IndexOf("ACME");
                            if (ITEMS == -1)
                            {
                                System.Data.DataTable FF = GETOITM2(R1);
                                if (FF.Rows.Count > 0)
                                {
                                    string P1 = FF.Rows[0][0].ToString().Trim();

                                    if (P1.ToUpper() != LOACTION.ToUpper())
                                    {

                                        MessageBox.Show(ITEMCODE + " 產地錯誤 無法匯入");
                                        WAR = "1";
                                        return;
                                    }

                                }
                            }

                        }
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    ITEMCODE2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    ES = range.Text.ToString().Trim();

                    string FLA = "N";

                    if (FLAG == "出口")
                    {
                        if (ITEMCODE.Length > 10)
                        {
                            VER = ITEMCODE.Substring(11, 1);
                        }

                        int L1 = GW.IndexOf("*");
                        if (L1 != -1)
                        {
                            string A1 = GW.Substring(0, L1);
                            string A2 = GW.Substring(L1 + 1, GW.Length - L1 - 1).ToUpper().Replace("C", "");
                            decimal GWF = 0;
                            decimal n;
                            if (decimal.TryParse(A1, out n) && decimal.TryParse(A2, out n))
                            {
                                GWF = Convert.ToDecimal(A1) * Convert.ToDecimal(A2);
                                GW = GWF.ToString();
                            }
                        }
                    }

                    System.Data.DataTable GF1 = GETOITM(ITEMCODE);
                    if (GF1.Rows.Count > 0)
                    {
                        FLA = "Y";
                    }

                    if (CARTONQTY == "空箱")
                    {
                        QTY = "空箱";
                        CARTONQTY = "0";
                    }
                    if (QTY == "空箱")
                    {
                        FLA = "Y";
                    }
                    if (ITEMCODE != "")
                    {
                        FLA = "Y";
                    }
                    if (FLA == "Y")
                    {
                        int PLA2 = PLATENO.LastIndexOf("-");
                        int start2 = 0;
                        int end2 = 0;
                        if (PLA2 != -1)
                        {
                            string D1 = PLATENO.Substring(0, PLA2);
                            start2 = Convert.ToInt16(D1);
                            end2 = Convert.ToInt16(PLATENO.Substring(PLA2 + 1, PLATENO.Length - PLA2 - 1));

                            if (start2 == end2)
                            {
                                PLATENO = D1;
                            }
                        }

                        int PLA = PLATENO.LastIndexOf("-");
                        int start = 0;
                        int end = 0;
                        if (PLA != -1)
                        {
                            start = Convert.ToInt16(PLATENO.Substring(0, PLA));
                            end = Convert.ToInt16(PLATENO.Substring(PLA + 1, PLATENO.Length - PLA - 1));
                            int f1 = (end - start) + 1;
                            CARTONQTY = (Convert.ToDecimal(CARTONQTY) / f1).ToString();
                            NW = (Convert.ToDecimal(NW) / f1).ToString();
                            if (String.IsNullOrEmpty(GW))
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord - 1, 11]);
                                range.Select();
                                GW = range.Text.ToString().Trim();
                            }
                            GW = Math.Round((float.Parse(GW) / f1), 2).ToString();//毛重取到小數第二位
                            CARTONNO = (Convert.ToDecimal(CARTONNO) / f1).ToString();
                            for (int i = start; i <= end; i++)
                            {

                                AddWHPACK2(WHNO, i.ToString(), CARTONNO, AUNO, ITEMCODE, GRADE, ITEMNAME, VER, QTY, CARTONQTY, NW, GW, L, W, H, MATERIAL, LOACTION, "", i.ToString(), fmLogin.LoginID.ToString(), "", BLC, CART, ITEMCODE2, ES);
                            }
                        }
                        else
                        {

                            if (String.IsNullOrEmpty(GW) && (QTY != "空箱"))
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord - 1, 11]);
                                range.Select();
                                string LF = range.Text.ToString().Trim();
                                if (LF != "")
                                {
                                    GGFN++;
                                    GFN = GGFN.ToString();
                                    System.Data.DataTable JJ1 = GetMAXID(WHNO, CART);
                                    if (JJ1.Rows.Count > 0)
                                    {

                                        UPWHPACK2N(JJ1.Rows[0][0].ToString(), GGFN.ToString().Trim());
                                    }
                                }


                                GW = GW2;
                                if (String.IsNullOrEmpty(MATERIAL))
                                {
                                    MATERIAL = MATERIAL2;
                                }

                                if (String.IsNullOrEmpty(LOACTION))
                                {
                                    LOACTION = LOACTION2;
                                }
                                L = "";
                                W = "";
                                H = "";
                                //  PLATENO = "";
                            }
                            else
                            {
                                GF = "0";
                                GFN = "0";
                            }

                            if (!String.IsNullOrEmpty(GW))
                            {

                                GW2 = GW;
                                MATERIAL2 = MATERIAL;
                                LOACTION2 = LOACTION;
                            }


                            //int ITEMS = ITEMCODE.ToUpper().IndexOf("ACME");
                            //if (String.IsNullOrEmpty(CARTONNO) && (ITEMNAMED2 == ITEMNAME) && (ITEMS == -1))
                            //{
                            //    System.Data.DataTable JJ1 = GetMAXID(WHNO, CART);
                            //    if (JJ1.Rows.Count > 0)
                            //    {
                            //        sb.Remove(sb.Length - 1, 1);
                            //        UPWHPACK3N2(JJ1.Rows[0][0].ToString(), QTY2.ToString(), CARTONQTY2.ToString(), sb.ToString());
                            //    }
                            //}
                            //else
                            //{

                            //    AddWHPACK2(WHNO, PLATENO, CARTONNO, AUNO, ITEMCODE, GRADE, ITEMNAME, VER, QTY, CARTONQTY, NW, GW, L, W, H, MATERIAL, LOACTION, GF, PLATENO2N, fmLogin.LoginID.ToString(), GFN, BLC,CART);
                            //}

                            AddWHPACK2(WHNO, PLATENO, CARTONNO, AUNO, ITEMCODE, GRADE, ITEMNAME, VER, QTY, CARTONQTY, NW, GW, L, W, H, MATERIAL, LOACTION, GF, PLATENO2N, fmLogin.LoginID.ToString(), GFN, BLC, CART, ITEMCODE2, ES);
                        }
                    }
                    AUNOD = AUNO;
                    ITEMCODED = ITEMCODE;
                    ITEMNAMED2 = ITEMNAME;
                    QTY2D = QTY2;
                    CARTONQTY2D = CARTONQTY2;
                    if (ITEMNAME.Length > 7)
                    {
                        ITEMNAMED = ITEMNAME.Substring(0, 8);
                    }
                    //棧板
                    if (NW == "栈板" || NW == "邊條" || NW == "空箱" || NW == "棧板")
                    {
                        decimal DD;
                        if (decimal.TryParse(GW, out DD))
                        {
                            GAN += Convert.ToDecimal(GW);
                        }
                    }
                }
                //if (FLAG == "出口")
                //{
                System.Data.DataTable H1FF = GetWH_PACK2(WHNO, CART);
                decimal PNW = 0;
                for (int i = 0; i <= H1FF.Rows.Count - 1; i++)
                {
                    string ID = H1FF.Rows[i]["ID"].ToString();
                    string SERS = H1FF.Rows[i]["SER"].ToString().Trim();
                    decimal FCARTONQTY = Convert.ToDecimal(H1FF.Rows[i]["CARTONQTY"]);
                    string NWN = H1FF.Rows[i]["NW"].ToString();
                    if (String.IsNullOrEmpty(NWN))
                    {
                        System.Data.DataTable HF = GetPACK1FS(WHNO, SERS, CART);
                        System.Data.DataTable HF2 = GetPACK1FS2(WHNO, SERS, CART);
                        if (HF.Rows.Count > 0 && HF2.Rows.Count > 0)
                        {
                            if (!String.IsNullOrEmpty(HF.Rows[0]["NW"].ToString()))
                            {

                                if (i == 0)
                                {
                                    PNW = Convert.ToDecimal(HF.Rows[0]["NW"]);
                                }
                                decimal PCARTONQTY = Convert.ToDecimal(HF.Rows[0]["CARTONQTY"]);


                                UPWHPACK3N(ID.ToString(), ((FCARTONQTY / PCARTONQTY) * PNW).ToString("0.000"));
                            }
                        }
                    }
                }
                //   }
                //if (FLAG == "出口")
                //{

                System.Data.DataTable H1N = GetWH_PACK2N(WHNO, CART);

                for (int i = 0; i <= H1N.Rows.Count - 1; i++)
                {
                    string ID = H1N.Rows[i]["ID"].ToString();
                    string ITEMCODEN = H1N.Rows[i]["ITEMCODE"].ToString();
                    string QQ = H1N.Rows[i]["QTY"].ToString();
                    string ITE = "";
                    if (ITEMCODEN.Length > 0)
                    {
                        ITE = ITEMCODEN.Substring(0, 1);
                    }
                    //if (ITEMCODE == "M215HAN01.50D07")
                    //{
                    //    MessageBox.Show("A");
                    //}
                    string ITEMNAMEN = H1N.Rows[i]["ITEMNAME"].ToString();
                    string NWN = H1N.Rows[i]["NW"].ToString();
                    string NGW = H1N.Rows[i]["GW"].ToString();
                    if (String.IsNullOrEmpty(NWN))
                    {
                        string VERN = H1N.Rows[i]["VER"].ToString();
                        string QTYNF = H1N.Rows[i]["QTY"].ToString();
                        if (String.IsNullOrEmpty(QTYNF))
                        {
                            QTYNF = "0";
                        }
                        decimal QTYN = Convert.ToDecimal(QTYNF);
                        string CARTNO = H1N.Rows[i]["CARTONNO"].ToString();
                        if (String.IsNullOrEmpty(CARTNO))
                        {
                            CARTNO = "1";
                        }
                        decimal CARTONNON = Convert.ToDecimal(CARTNO);

                        int FF1 = ITEMCODEN.IndexOf(".");

                        if (FF1 != -1)
                        {
                            string MODEL = ITEMCODEN.Substring(0, ITEMCODEN.IndexOf("."));
                            string MODEL2 = ITEMCODEN.Substring(1, ITEMCODEN.IndexOf(".") - 1);
                            string GRADEN = "";
                            int IN1 = ITEMNAMEN.IndexOf(".");
                            if (ITEMNAMEN.Length > 0 && IN1 != -1)
                            {
                                string ITEMM = ITEMNAMEN.Substring(IN1 + 1, ITEMNAMEN.Length - IN1 - 1);
                                if (ITEMM.Length > 2)
                                {
                                    GRADEN = ITEMNAMEN.Substring(ITEMNAMEN.IndexOf(".") + 1, 3);
                                }
                            }
                            System.Data.DataTable ITEM1 = util.GetOITMW(ITEMCODEN);
                            if (ITEM1.Rows.Count > 0)
                            {
                                GRADEN = ITEM1.Rows[0][0].ToString();
                            }

                            //SELECT SUBSTRING(ITEMNAME,CHARINDEX('V.', ITEMNAME)+2,3) VER FROM OITM WHERE ITEMCODE='TG121SN01.04022'
                            System.Data.DataTable J1 = null;

                            J1 = util.GetCART(MODEL, VERN, GRADEN, 0, QQ);

                            if (J1.Rows.Count == 0)
                            {
                                J1 = util.GetCARTK(MODEL2, VERN, GRADEN, ITE, 0, QQ);
                            }
                            if (J1.Rows.Count == 0)
                            {
                                J1 = util.GetCARTL(MODEL, VERN, 0, QQ);
                            }
                            if (J1.Rows.Count == 0)
                            {
                                J1 = util.GetCARTJ(MODEL2, VERN, 0, QQ);
                            }
                            if (J1.Rows.Count == 0)
                            {
                                System.Data.DataTable TG = util.GetOITML(MODEL2);
                                if (TG.Rows.Count > 0)
                                {
                                    string TTMODEL = TG.Rows[0][0].ToString();
                                    int TT1 = TTMODEL.IndexOf("/");
                                    if (TT1 != -1)
                                    {
                                        string[] s = TTMODEL.Split('/');
                                        string H1 = s[1];
                                        J1 = util.GetCARTJ(H1, VERN, 0, QQ);
                                    }
                                }
                            }

                            if (J1.Rows.Count > 0)
                            {
                                decimal CT_QTY = Convert.ToDecimal(J1.Rows[0]["CT_QTY"]);
                                decimal CT_NW = Convert.ToDecimal(J1.Rows[0]["CT_NW"]);

                                UPWHPACK3N(ID.ToString(), (((QTYN / CT_QTY) * CT_NW) * CARTONNON).ToString("0.000"));
                            }
                            else
                            {
                                MessageBox.Show(ITEMCODEN + " 包裝規格沒有資料 上傳失敗");
                                WAR = "1";
                                return;

                            }
                        }
                    }
                    else
                    {

                        string VERN = H1N.Rows[i]["VER"].ToString();
                        string QTYNF = H1N.Rows[i]["QTY"].ToString();
                        string MODEL = ITEMCODEN.Substring(0, ITEMCODEN.IndexOf("."));
                        string MODEL2 = ITEMCODEN.Substring(1, ITEMCODEN.IndexOf(".") - 1);
                        string GRADEN = "";
                        int IN1 = ITEMNAMEN.IndexOf(".");
                        if (ITEMNAMEN.Length > 0 && IN1 != -1)
                        {
                            string ITEMM = ITEMNAMEN.Substring(IN1 + 1, ITEMNAMEN.Length - IN1 - 1);
                            if (ITEMM.Length > 2)
                            {
                                GRADEN = ITEMNAMEN.Substring(ITEMNAMEN.IndexOf(".") + 1, 3);
                            }
                        }
                        System.Data.DataTable J1 = null;

                        J1 = util.GetCART(MODEL, VERN, GRADEN, 1, QQ);
                        if (J1.Rows.Count == 0)
                        {
                            J1 = util.GetCARTK(MODEL2, VERN, GRADEN, ITE, 1, QQ);
                        }
                        if (J1.Rows.Count == 0)
                        {
                            J1 = util.GetCARTL(MODEL, VERN, 1, QQ);
                        }
                        if (J1.Rows.Count == 0)
                        {
                            J1 = util.GetCARTJ(MODEL2, VERN, 1, QQ);
                        }
                        //if (J1.Rows.Count > 0)
                        //{
                        //    string CARTNO = H1N.Rows[i]["CARTONNO"].ToString();
                        //    decimal QTYN = Convert.ToDecimal(QTYNF);
                        //    decimal CARTONNON = Convert.ToDecimal(CARTNO);
                        //    decimal CT_QTY = Convert.ToDecimal(J1.Rows[0]["CT_QTY"]);
                        //    decimal CT_NW = Convert.ToDecimal(J1.Rows[0]["CT_NW"]);

                        //    UPWHPACK3N(ID.ToString(), (((QTYN / CT_QTY) * NWN) * CARTONNON).ToString("0.00"));
                        //}
                        //else
                        //{
                        //    MessageBox.Show(ITEMCODEN + " 包裝規格沒有資料 上傳失敗");
                        //    WAR = "1";
                        //    return;

                        //}
                        int L1 = 0;
                        decimal n;


                        //if (J1.Rows.Count > 0)
                        //{
                        //    if (decimal.TryParse(J1.Rows[0]["PAL_GW"].ToString(), out n) && decimal.TryParse(J1.Rows[0]["PAL_CTNS"].ToString(), out n) && decimal.TryParse(NGW, out n))
                        //    {
                        //        string CARTNO = H1N.Rows[i]["CARTONNO"].ToString();
                        //        decimal CARTONNON = Convert.ToDecimal(CARTNO);
                        //        decimal PAL_CTNS = Convert.ToDecimal(J1.Rows[0]["PAL_CTNS"]);
                        //        if (CARTONNON == PAL_CTNS)
                        //        {
                        //            decimal PAL_GW = Convert.ToDecimal(J1.Rows[0]["PAL_GW"]);
                        //            decimal GGW = Convert.ToDecimal(NGW);
                        //            decimal GS1 = GGW + 10;
                        //            decimal GS2 = GGW - 10;

                        //            if (PAL_GW > GS1)
                        //            {
                        //                L1 = 1;
                        //            }
                        //            if (PAL_GW < GS2)
                        //            {
                        //                L1 = 1;
                        //            }
                        //        }
                        //    }
                        //}


                        //if (L1 == 1)
                        //{
                        //    MessageBox.Show("料號 : " + ITEMCODEN + " 毛重異常");
                        //}
                    }

                    //}
                }

                //System.Data.DataTable H1 = GetWH_PACK2(WHNO, CART);

                //for (int i = 0; i <= H1.Rows.Count - 1; i++)
                //{
                //    string ID = H1.Rows[i]["ID"].ToString();
                //    string SERS = H1.Rows[i]["SER"].ToString().Trim();
                //    decimal FCARTONQTY = Convert.ToDecimal(H1.Rows[i]["CARTONQTY"]);
                //    System.Data.DataTable HF = GetPACK1(WHNO, SERS, CART);
                //    if (HF.Rows.Count > 0)
                //    {
                //        decimal PGW = Convert.ToDecimal(HF.Rows[0]["GW"]);
                //        decimal PCARTONQTY = Convert.ToDecimal(HF.Rows[0]["CARTONQTY"]);
                //        UPWHPACK3(ID.ToString(), ((FCARTONQTY / PCARTONQTY) * PGW).ToString("0.00"), PGW.ToString());
                //    }
                //}

                System.Data.DataTable H2 = GetWH_PACK2S(WHNO, CART);
                string SERSD = "";
                for (int i = 0; i <= H2.Rows.Count - 1; i++)
                {
                    string ID = H2.Rows[i]["ID"].ToString();
                    string SERS = H2.Rows[i]["SER2"].ToString().Trim();
                    decimal FNW = Convert.ToDecimal(H2.Rows[i]["NW"]);
                    System.Data.DataTable HF = GetPACK1S(WHNO, SERS, CART);
                    if (HF.Rows.Count > 0)
                    {

                        string IDM = HF.Rows[0]["ID"].ToString();
                        decimal PGW = Convert.ToDecimal(HF.Rows[0]["GW"]);
                        decimal PNWF = Convert.ToDecimal(HF.Rows[0]["NW"]);
                        decimal PCARTONQTY = Convert.ToDecimal(HF.Rows[0]["CARTONQTY"]);
                        decimal PP1 = Convert.ToDecimal(((PGW / PNWF) * FNW).ToString("0.00"));
                        UPWHPACK3(ID.ToString(), ((PGW / PNWF) * FNW).ToString("0.00"), PGW.ToString());
                        if (ID == IDM)
                        {
                            System.Data.DataTable HF2 = GetPACK1S2(WHNO, SERS, CART, IDM);
                            if (HF2.Rows.Count > 0)
                            {
                                decimal HFGW = Convert.ToDecimal(HF2.Rows[0]["GW"]);
                                decimal PP2 = PGW - HFGW;
                                UPWHPACK3(ID.ToString(), (PP2).ToString("0.00"), PGW.ToString());
                            }
                        }

                    }

                }



                UPWHPACK4(WHNO);

                System.Data.DataTable H1F = GetPACK1F(WHNO, CART);
                if (H1F.Rows.Count > 0)
                {
                    for (int i = 0; i <= H1F.Rows.Count - 1; i++)
                    {
                        decimal FGW = Convert.ToDecimal(H1F.Rows[i]["GW"]);
                        string FSER = H1F.Rows[i]["SER"].ToString();
                        System.Data.DataTable H1F2 = GetPACK1F2(WHNO, FSER, CART);
                        string ID = H1F2.Rows[0]["ID"].ToString();
                        decimal FGW2 = Convert.ToDecimal(H1F2.Rows[0]["GW"]) + Convert.ToDecimal(H1F.Rows[i]["GW"]);
                        UPWHPACK5(ID, FGW2.ToString());
                    }
                }

                if (GAN != 0)
                {
                    System.Data.DataTable HF = GetPACK1SB(WHNO, GAN);
                    decimal FFS = 0;
                    decimal FW = 0;
                    if (HF.Rows.Count > 0)
                    {
                        FW = Convert.ToDecimal(HF.Rows[0]["FW"]);
                        decimal FNW = Convert.ToDecimal(HF.Rows[0]["FNW"]);
                        FFS = FW / FNW;
                    }
                    System.Data.DataTable H2B = GetWH_PACK2SB(WHNO);

                    for (int i = 0; i <= H2B.Rows.Count - 1; i++)
                    {
                        string ID = H2B.Rows[i]["ID"].ToString();
                        decimal NWS = Convert.ToDecimal(H2B.Rows[i]["NW"]);
                        decimal f1 = (FFS) * NWS;
                        if (i == H2B.Rows.Count - 1)
                        {
                            System.Data.DataTable H2B2 = GetWH_PACK2SB2(WHNO, ID);
                            f1 = FW - Convert.ToDecimal(H2B2.Rows[0][0]);
                        }

                        UPWHPACK5(ID.ToString(), (f1).ToString("0.00"));


                    }
                }
                if (B > 1)
                {
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                    excelSheet2.Activate();
                    int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;

                    for (int iRecord2 = 1; iRecord2 <= iRowCnt2; iRecord2++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord2, 1]);
                        range.Select();
                        string MARK = range.Text.ToString().Trim();
                        if (!String.IsNullOrEmpty(MARK))
                        {
                            AddWHPACK3(WHNO, MARK);
                        }
                    }
                }


            }
            catch (Exception ex) 
            {

            }
            finally
            {



                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();





            }



        }

        public  void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }
        public void AddWHPACK2(string ShippingCode, string PLATENO, string CARTONNO, string AUNO, string ITEMCODE, string GRADE, string ITEMNAME, string VER,
            string QTY, string CARTONQTY, string NW, string GW, string L, string W, string H, string MATERIAL, string LOACTION, string SER, string PLATENO2, string USERS, string SER2, string BLC,string FLAG1, string ITEMCODE2, string ES)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into WH_PACK2(ShippingCode,PLATENO,CARTONNO,AUNO,ITEMCODE,GRADE,ITEMNAME,VER,QTY,CARTONQTY,NW,GW,L,W,H,MATERIAL,LOACTION,SER,PLATENO2,USERS,SER2,BLC,FLAG1,ES) values(@ShippingCode,@PLATENO,@CARTONNO,@AUNO,@ITEMCODE,@GRADE,@ITEMNAME,@VER,@QTY,@CARTONQTY,@NW,@GW,@L,@W,@H,@MATERIAL,@LOACTION,@SER,@PLATENO2,@USERS,@SER2,@BLC,@FLAG1,@ES)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@PLATENO", PLATENO));
            command.Parameters.Add(new SqlParameter("@CARTONNO", CARTONNO));
            command.Parameters.Add(new SqlParameter("@AUNO", AUNO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARTONQTY", CARTONQTY));
            command.Parameters.Add(new SqlParameter("@GW", GW));
            command.Parameters.Add(new SqlParameter("@NW", NW));
            command.Parameters.Add(new SqlParameter("@L", L));
            command.Parameters.Add(new SqlParameter("@W", W));
            command.Parameters.Add(new SqlParameter("@H", H));
            command.Parameters.Add(new SqlParameter("@MATERIAL", MATERIAL));
            command.Parameters.Add(new SqlParameter("@LOACTION", LOACTION));
            command.Parameters.Add(new SqlParameter("@SER", SER));
            command.Parameters.Add(new SqlParameter("@PLATENO2", PLATENO2));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@SER2", SER2));
            command.Parameters.Add(new SqlParameter("@BLC", BLC));
            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG1));
            command.Parameters.Add(new SqlParameter("@ITEMCODE2", ITEMCODE2));//特殊料號
            command.Parameters.Add(new SqlParameter("@ES", ES));
            

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
        public void AddWHPACK3(string SHIPPINGCODE, string MARK)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into WH_PACK3(SHIPPINGCODE,MARK) values(@SHIPPINGCODE,@MARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@MARK", MARK));


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
        public void UPWHPACK2( string ID, string SER)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET SER=@SER WHERE ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@SER", SER));


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
        public void UPWHPACK2N(string ID, string SER)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET SER2=@SER WHERE ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@SER", SER));


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
        public void UPWHPACK3(string ID, string GW, string CARTONNO2)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET GW=@GW,CARTONNO2=@CARTONNO2 WHERE  ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@GW", GW));
            command.Parameters.Add(new SqlParameter("@CARTONNO2", CARTONNO2));

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
    
        public void UPWHPACK3N(string ID, string NW)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET NW=@NW WHERE  ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@NW", NW));


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
        public void UPWHPACK3NF(string ID, string NW, string CARTONNO2)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET NW=@NW,CARTONNO2=@CARTONNO2 WHERE  ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@NW", NW));
            command.Parameters.Add(new SqlParameter("@CARTONNO2", CARTONNO2));

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
        public void UPWHPACK3N2(string ID, string QTY, string CARTONQTY, string AUNO)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET QTY =@QTY,CARTONQTY =@CARTONQTY,AUNO=@AUNO WHERE ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CARTONQTY", CARTONQTY));
            command.Parameters.Add(new SqlParameter("@AUNO", AUNO));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        public void UPWHPACK4(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET GW='0' WHERE SHIPPINGCODE=@SHIPPINGCODE AND QTY='空箱' ", connection);
            command.CommandType = CommandType.Text;
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

        public void UPWHPACK5(string ID,string  GW)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE WH_PACK2 SET GW=@GW WHERE ID=@ID ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@GW", GW));

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
        private void DELETETA(string SHIPPINGCODE,string FLAG)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            sb.Append(" DELETE WH_PACK3 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            if (!String.IsNullOrEmpty(FLAG))
            {
                sb.Append(" AND FLAG1=@FLAG1 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            command.Parameters.Add(new SqlParameter("@FLAG1", FLAG));
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

        private void button5_Click(object sender, EventArgs e)
        {
            WAR = "";
            string d = textBox2.Text;

            if (!String.IsNullOrEmpty(d))
            {
                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {
                    string USER = fmLogin.LoginID.ToString();
                    FileInfo info = new FileInfo(file);
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");
                    WriteExcel(file, "出口", NAME);
                    if (USER.ToUpper() != "LLEYTONCHEN")
                    {
                        File.Delete(file);
                    }
                }
                if (WAR == "")
                {
                    MessageBox.Show("上傳成功");
                }

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GetOUT1();

            if (G1.Rows.Count > 0)
            {
                for (int i = 0; i <= G1.Rows.Count-1; i++)
                {
                    DataRow drw = G1.Rows[i];

                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    string aa = drw["path"].ToString() + "\\" + drw["路徑"].ToString();

                    string filename = drw["檔案名稱"].ToString();
                    string NewFileName = lsAppDir + "\\AA\\" + filename;

                    System.IO.File.Copy(aa, NewFileName, true);
                }
            }
        }
        private System.Data.DataTable GetOUT1()
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            
            sb.Append(" SELECT '香港冠榮電子有限公司' 公司,cast(TRGTPATH as nvarchar(80))  [path],FILENAME+'.'+Fileext 檔案名稱,'\'+CAST([FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑 FROM ATC1　WHERE ABSENTRY IN (SELECT ATCENTRY FROM OCLG WHERE CARDCODE='1594-04')");
            sb.Append(" AND (FILENAME LIKE '%簽%' OR FileName LIKE '%签%')");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '廣州冠榮電子科技有限公司',cast(TRGTPATH as nvarchar(80))  [path],FILENAME+'.'+Fileext 檔案名稱,'\'+CAST([FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑　FROM ATC1　WHERE ABSENTRY IN (SELECT ATCENTRY FROM OCLG WHERE CARDCODE='1594-04')");
            sb.Append(" AND (FILENAME LIKE '%簽%' OR FileName LIKE '%签%')");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }



        private void btnExUpload_Click(object sender, EventArgs e)
        {
            WAR = "";
            string d = textBox3.Text;


            if (!String.IsNullOrEmpty(d))
            {
                string[] filenames = Directory.GetFiles(d);
                int M = 0;
                foreach (string file in filenames)
                {
                    AUINV a = new AUINV();
                    a.readExcel(file);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            System.Data.DataTable  G1 = GETF1();
            for (int i = 0; i <= G1.Rows.Count - 1; i++)
            {
                //AtcEntry
                string DOCENTRY = G1.Rows[i][0].ToString();
               // string DOC = G1.Rows[i][1].ToString();
                System.Data.DataTable G2 = GETF2(DOCENTRY);
                if (G2.Rows.Count > 0)
                {
                  //  AddOACT(Convert.ToInt32(DOC));
                    UPOPDN(Convert.ToInt32(G2.Rows[0][0]), DOCENTRY);
                }
            }

            }
    }
}
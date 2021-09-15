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
    public partial class DOCUPLOADS : Form
    {
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
        public DOCUPLOADS()
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


            }
        }

        private void DOCUPLOAD_Load(object sender, EventArgs e)
        {

             System.Data.DataTable G2 = GetMenu.Getdata("SHIPYES");
            if (G2.Rows.Count > 0)
            {
                textBox2.Text = G2.Rows[0][0].ToString();
            }
            System.Data.DataTable G3 = GetMenu.Getdata("SHIPNO");
            if (G3.Rows.Count > 0)
            {
                textBox3.Text = G3.Rows[0][0].ToString();
            }

            System.Data.DataTable G4 = GetMenu.Getdata("SHIPOUT");
            if (G4.Rows.Count > 0)
            {
                textBox1.Text = G4.Rows[0][0].ToString();
            }

            System.Data.DataTable G5 = GetMenu.Getdata("BIN");
            if (G5.Rows.Count > 0)
            {
                textBox4.Text = G5.Rows[0][0].ToString();
            }
            DIR = "//acmesrv01//SAP_Share//shipping//";
            PATH = @"\\acmesrv01\SAP_Share\shipping\";

            if (globals.DBNAME == "達睿生")
            {
                DIR = "//acmesrv01//SAP_Share//shipping達睿生//";
                PATH = @"\\acmesrv01\SAP_Share\shipping達睿生\";
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
        public void deleteDOWNLOAD(string filename)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("DELETE Download WHERE filename=@filename", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@filename", filename));


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
        public void AddADOWNLOADCAR(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into Shipping_CARDownload(shippingcode,seq,filename,path) values(@shippingcode,@seq,@filename,@path)", connection);
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
        private System.Data.DataTable GETODLNF(string U_WH_NO)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY DOC  FROM ODLN WHERE U_Shipping_no like '%" + U_WH_NO + "%' AND U_Shipping_no LIKE '%SH%' ");

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
        private System.Data.DataTable GetODLN2(string FILENAME, string DOCENTRY)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT T3.[FILENAME]   from oclg t2   ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)   ");
            sb.Append(" where  t2.doctype='15'   ");
            sb.Append(" and T3.[FILENAME] =@FILENAME AND T2.DOCENTRY=@DOCENTRY");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
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
        private System.Data.DataTable GetDOWNSEQCAR(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT COUNT(*) SEQ FROM Shipping_CARDownload WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

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
        private System.Data.DataTable GetDOWNSEQSCAR(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE FROM Shipping_CAR WHERE SHIPPINGCODE=@SHIPPINGCODE   ");

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
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                inint = 0;
                s1 = 0;
                string d = textBox2.Text;

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {


                    FileInfo info = new FileInfo(file);
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");
   
                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");
                        int J2 = NAME.IndexOf("-");
                        if (J2 == -1)
                        {
                            J2 = NAME.IndexOf("_");
                        }

                        string M2 = NAME.Substring(0, J1);
                        string SHIP = M2.Substring(0, J2);
                        if (globals.DBNAME == "禾中")
                        {
                            SHIP = M2.Substring(0, 17);
                        }
                        string SEQ = "0";
                        string f = "c";
                        string[] filebType = Directory.GetDirectories(DIR);
                        string dd = DateTime.Now.ToString("yyyyMM");
                        string tt = DIR + dd;
                        foreach (string fileaSize in filebType)
                        {

                            if (fileaSize == tt)
                            {
                                f = "d";

                            }

                        }
                        if (f == "c")
                        {
                            Directory.CreateDirectory(tt);
                        }
                        try
                        {
                            string server = DIR + dd + "//";
                            string server2 = PATH + dd + "\\" + NAME;
                            string USER = fmLogin.LoginID.ToString();
                                System.Data.DataTable O2 = GetMenu.GetSHIPOHEM(USER);
                                if (O2.Rows.Count > 0)
                                {
                                    deleteDOWNLOAD(NAME);
                                }


                            System.Data.DataTable dt2 = GetMenu.download(NAME);

                            if (dt2.Rows.Count > 0)
                            {
                                MessageBox.Show("檔案名稱重複,請修改檔名");
                            }
                            else
                            {
                                System.Data.DataTable dt3 = GetDOWNSEQ(SHIP);
                                System.Data.DataTable dt4 = GetDOWNSEQS(SHIP);
                                if (dt4.Rows.Count > 0)
                                {
                                    if (dt3.Rows.Count > 0)
                                    {
                                        SEQ = dt3.Rows[0][0].ToString();
                                    }
                                    string UATT = textBox2.Text + @"\" + NAME;

                                    bool FF1 = getrma.UploadFile(UATT, server, false);
                                    if (FF1 == false)
                                    {
                                        return;
                                    }
                                    AddADOWNLOAD(SHIP, SEQ, NAME, server2);

                                    System.GC.Collect();
                                    System.GC.WaitForPendingFinalizers();
                                    File.Delete(file);

                                    s1 = 1;
                                }

                              
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
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
                MessageBox.Show(ex.Message);
            }

        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();
                if (GetMenu.Getdata("SHIPYES").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "SHIPYES");
                }
                else
                {
                    GetMenu.UP(t1, "SHIPYES");
                }

                textBox2.Text = t1;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("SHIPNO").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "SHIPNO");
                }
                else
                {
                    GetMenu.UP(t1, "SHIPNO");
                }

                textBox3.Text = t1;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                inint = 0;
                s1 = 0;
                string d = textBox3.Text;

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {


                    FileInfo info = new FileInfo(file);
   
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");
          
           
                        if (NAME != "Thumbs.db")
                        {

                            int J1 = NAME.IndexOf(".");
                            int J2 = NAME.IndexOf("-");
                            if (J2 == -1)
                            {
                                J2 = NAME.IndexOf("_");
                            }
                            string M2 = NAME.Substring(0, J1);
                            string SHIP = M2.Substring(0, J2);
                            if (globals.DBNAME == "禾中")
                            {
                                SHIP = M2.Substring(0, 17);
                            }
                            string SEQ = "0";
                            string f = "c";
                            string[] filebType = Directory.GetDirectories(DIR);
                            string dd = DateTime.Now.ToString("yyyyMM");
                            string tt = DIR + dd;
                            foreach (string fileaSize in filebType)
                            {

                                if (fileaSize == tt)
                                {
                                    f = "d";

                                }

                            }
                            if (f == "c")
                            {
                                Directory.CreateDirectory(tt);
                            }
                            try
                            {
                                string server = DIR + dd + "//";
                                string server2 = PATH + dd + "\\" + NAME;
                         

                                System.Data.DataTable dt2 = GetMenu.download2(NAME);

                                if (dt2.Rows.Count > 0)
                                {


                                    MessageBox.Show("檔案名稱重複,請修改檔名");
                                }
                                else
                                {
                                    System.Data.DataTable dt3 = GetDOWNSEQ2(SHIP);
                                    System.Data.DataTable dt4 = GetDOWNSEQS(SHIP);
                                    if (dt4.Rows.Count > 0)
                                    {
                                        if (dt3.Rows.Count > 0)
                                        {
                                            SEQ = dt3.Rows[0][0].ToString();
                                        }
                                        string UATT = textBox3.Text + @"\" + NAME;


                                        bool FF1 = getrma.UploadFile(UATT, server, false);
                                        if (FF1 == false)
                                        {
                                            return;
                                        }
                                        AddADOWNLOAD2(SHIP, SEQ, NAME, server2);

                                        System.GC.Collect();
                                        System.GC.WaitForPendingFinalizers();
                                        File.Delete(file);

                                        s1 = 1;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
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
                MessageBox.Show(ex.Message);
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

                    DirectoryInfo S = new DirectoryInfo(d);


                    FileInfo info = new FileInfo(file);
                    string DIRS = S.Name.ToString().Trim().Replace(" ", "");
                    string NAME = info.Name.ToString().Trim();

                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");
                        int J2 = NAME.IndexOf("-");
                        if (J2 == -1)
                        {
                            J2 = NAME.IndexOf("_");
                        }
                        string M1 = NAME.Substring(J1 + 1, NAME.Length - J1 - 1);
                        string FILENAME = NAME.Substring(0, J1);
                        string M2 = NAME.Substring(0, J1);
                        System.Data.DataTable dt4F = GetDOWNSEQS(DIRS);
                        string SHIP = "";
                        if (dt4F.Rows.Count > 0)
                        {
                            SHIP = DIRS;
                        }
                        else
                        {
                             SHIP = M2.Substring(0, J2);
                        }
                    
                        string SEQ = "0";
                        string f = "c";
                        string[] filebType = Directory.GetDirectories(DIR);
                        string dd = DateTime.Now.ToString("yyyyMM");
                        string tt = DIR + dd;
                        foreach (string fileaSize in filebType)
                        {

                            if (fileaSize == tt)
                            {
                                f = "d";

                            }

                        }
                        if (f == "c")
                        {
                            Directory.CreateDirectory(tt);
                        }
                        try
                        {
                            System.Data.DataTable dt4 = GetDOWNSEQS(SHIP);


                            string server = DIR + dd + "//";
                            string server2 = PATH + dd + "\\" + NAME;
                            string USER = fmLogin.LoginID.ToString();
                            System.Data.DataTable O2 = GetMenu.GetSHIPOHEM(USER);
                            if (O2.Rows.Count > 0)
                            {
                                deleteDOWNLOAD(NAME);
                            }

                            System.Data.DataTable dtSU = GetMenu.GETSUIP();
                            int P1 = 0;
                            if (dtSU.Rows.Count > 0)
                            {
                                for (int i = 0; i <= dtSU.Rows.Count - 1; i++)
                                {
                                    string SUIP = dtSU.Rows[i][0].ToString().Trim();
                                    int P2 = NAME.ToUpper().IndexOf(SUIP);
                                    if (P2 != -1)
                                    {
                                        P1 = 1;
                                    }
                                }
                            }

                            int P3 = NAME.ToUpper().IndexOf("線上簽收");
                            int P4 = NAME.ToUpper().IndexOf("簽收單");

                            int P5  = NAME.ToUpper().IndexOf("IPPC");
                            if (P5 != -1)
                            {
                                P1 = 0;
                            }
                            System.Data.DataTable dt2 = null;
                            if (P1 == 1)
                            {
                                dt2 = GetMenu.download(NAME);
                            }
                            else
                            {
                                dt2 = GetMenu.download2(NAME);
                            }

                            if (dt2.Rows.Count > 0)
                            {
                                MessageBox.Show("檔案名稱重複,請修改檔名");
                            }
                            else
                            {

                                System.Data.DataTable dt3 = null;

                                if (P1 == 1)
                                {
                                    dt3 = GetDOWNSEQ(SHIP);
                                }
                                else
                                {
                                    dt3 = GetDOWNSEQ2(SHIP);
                                }
                                if (dt4.Rows.Count > 0)
                                {
                                    if (dt3.Rows.Count > 0)
                                    {
                                        SEQ = dt3.Rows[0][0].ToString();
                                    }
                                    string UATT = textBox1.Text + @"\" + NAME;

                                    bool FF1 = getrma.UploadFile(UATT, server, false);
                                    if (FF1 == false)
                                    {
                                        return;
                                    }
                                    if (P1 == 1)
                                    {
                                        AddADOWNLOAD(SHIP, SEQ, NAME, server2);
                                    }
                                    else
                                    {
                                        AddADOWNLOAD2(SHIP, SEQ, NAME, server2);
                                    }

                                    if (P3 != -1 || P4 != -1)
                                    {
                                        System.Data.DataTable L1 = GETODLNF(SHIP);

                                        if (L1.Rows.Count > 0)
                                        {
                                            for (int i = 0; i <= L1.Rows.Count - 1; i++)
                                            {
                                                int n;
                                                string CARDCODE = "";
                                                string TEL1 = "";
                                                string CNTCTCODE = "";
                                                string SLPCODE = "";
                                                string SAPDOC = L1.Rows[i][0].ToString();
                                                if (int.TryParse(SAPDOC, out n))
                                                {
                                                    System.Data.DataTable L2 = GetODLN2(FILENAME,SAPDOC);
                                                    if (L2.Rows.Count == 0)
                                                    {
                                                        System.Data.DataTable T1 = GetODLN("10000000");
                                                        T1 = GetODLN(SAPDOC);
                                                        if (T1.Rows.Count > 0)
                                                        {


                                                            System.Data.DataTable T2 = GetMAXOCLG2("15", SAPDOC);

                                                            DataRow dd1 = T1.Rows[0];
                                                            CARDCODE = dd1["CARDCODE"].ToString();
                                                            TEL1 = dd1["TEL1"].ToString();
                                                            CNTCTCODE = dd1["CNTCTCODE"].ToString();
                                                            SLPCODE = dd1["SLPCODE"].ToString();


                                                            int d1 = Convert.ToInt32(GetMAXOCLG().Rows[0][0].ToString());
                                                            int m2 = Convert.ToInt32(GetMAXOATC().Rows[0][0].ToString());
                                                            DateTime now = DateTime.Now;
                                                            int d2 = Convert.ToInt16(DateTime.Now.ToString("HHmm"));

                                                            string ATT = @"C:\Program Files\SAP\SAP Business One\Attachments";
                                                            string UATT2 = PATH;
                                                            string ATT2 = @"\\ACMEW08R2AP\SAPFILES2\Attachments" + DateTime.Now.ToString("yyyy") + "\\ATT" + DateTime.Now.ToString("yyyyMM");
                                                            bool FF12 = getrma.UploadFile(server2, ATT2, false);
                                                            if (FF12 == false)
                                                            {
                                                                return;
                                                            }
                                                            if (T2.Rows.Count == 0)
                                                            {
                                                                AddOCLG(d1, CARDCODE, now, d2, now, "N", TEL1, -1, "N", OBJ, SAPDOC, SAPDOC, ATT, "l", 1,
                                                                    Convert.ToInt32(CNTCTCODE), 1, Convert.ToInt32(SLPCODE), "C", -1, d2, "M", "1", "N", 15, "M", "N", 0, "N", "N", "N", m2, now.AddDays(1), d2);
                                                                AddOACT(m2);
                                                                AddATC1(m2, 1, ATT, ATT2, FILENAME, M1, now, 1, "Y", "Y");
                                                                UPONNM(d1 + 1, "33");
                                                                UPONNM(m2 + 1, "221");
                                                            }
                                                            else
                                                            {
                                                                DataRow dd2 = T2.Rows[0];
                                                                string ATCENTRY = dd2["ATCENTRY"].ToString();
                                                                if (String.IsNullOrEmpty(ATCENTRY))
                                                                {
                                                                    AddOACT(m2);
                                                                    AddATC1(m2, 1, ATT, ATT2, FILENAME, M1, now, 1, "Y", "Y");
                                                                    UPOCLG(m2, OBJ, SAPDOC);
                                                                    UPONNM(m2 + 1, "221");
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

                                                                    AddATC1(Convert.ToInt32(ATCENTRY), g1, ATT, ATT2, FILENAME, M1, now, 1, "Y", "Y");

                                                                }
                                                            }


                                                            System.GC.Collect();
                                                            System.GC.WaitForPendingFinalizers();


                                                            s1 = 1;
                                                        }
                                                    }
                                                }
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
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
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
                MessageBox.Show(ex.Message);
            }
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
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();
                if (GetMenu.Getdata("SHIPOUT").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "SHIPOUT");
                }
                else
                {
                    GetMenu.UP(t1, "SHIPOUT");
                }

                textBox1.Text = t1;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                inint = 0;
                s1 = 0;
                string d = textBox4.Text;

                string[] filenames = Directory.GetFiles(d);
                foreach (string file in filenames)
                {


                    FileInfo info = new FileInfo(file);
                    string NAME = info.Name.ToString().Trim().Replace(" ", "");

                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");
                        int J2 = NAME.IndexOf("-");
                        if (J2 == -1)
                        {
                            J2 = NAME.IndexOf("_");
                        }

                        string M2 = NAME.Substring(0, J1);
                        string SHIP = M2.Substring(0, J2);
                 
                        string SEQ = "0";
                        string f = "c";
                        string[] filebType = Directory.GetDirectories(DIR);
                        string dd = DateTime.Now.ToString("yyyyMM");
                        string tt = DIR + dd;
                        foreach (string fileaSize in filebType)
                        {

                            if (fileaSize == tt)
                            {
                                f = "d";

                            }

                        }
                        if (f == "c")
                        {
                            Directory.CreateDirectory(tt);
                        }
                        try
                        {
                            string server = DIR + dd + "//";
                            string server2 = PATH + dd + "\\" + NAME;
                            string USER = fmLogin.LoginID.ToString();
           

                            System.Data.DataTable dt2 = GetMenu.downloadCAR(NAME);

                            if (dt2.Rows.Count > 0)
                            {
                                MessageBox.Show("檔案名稱重複,請修改檔名");
                            }
                            else
                            {
                                System.Data.DataTable dt3 = GetDOWNSEQCAR(SHIP);
                                System.Data.DataTable dt4 = GetDOWNSEQSCAR(SHIP);
                                if (dt4.Rows.Count > 0)
                                {
                                    if (dt3.Rows.Count > 0)
                                    {
                                        SEQ = dt3.Rows[0][0].ToString();
                                    }
                                    string UATT = textBox4.Text + @"\" + NAME;

                                    bool FF1 = getrma.UploadFile(UATT, server, false);
                                    if (FF1 == false)
                                    {
                                        return;
                                    }
                                    AddADOWNLOADCAR(SHIP, SEQ, NAME, server2);

                                    System.GC.Collect();
                                    System.GC.WaitForPendingFinalizers();
                                    File.Delete(file);

                                    s1 = 1;
                                }


                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
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
                MessageBox.Show(ex.Message);
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();
                if (GetMenu.Getdata("BIN").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "BIN");
                }
                else
                {
                    GetMenu.UP(t1, "BIN");
                }

                textBox4.Text = t1;
            }
        }

     

 
    }
}
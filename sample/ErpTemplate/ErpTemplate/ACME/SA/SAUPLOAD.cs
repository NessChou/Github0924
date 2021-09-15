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
    public partial class SAUPLOAD : Form
    {
        string USER = "";
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
        public SAUPLOAD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                 if (Getdata(USER, "SAPDRS").Rows.Count == 0)
                {
                    Add(USER, t1, "SAPDRS");
                }
                else
                {
                    UP(USER, t1, "SAPDRS");
                }

                textBox1.Text = t1;
            }
        }
        public void Add(string USERS, string DOCPATH, string DOCKIND)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO [WH_SAPDOC]");
            sb.Append("            (USERS,DOCPATH,DOCKIND)");
            sb.Append("      VALUES(@USERS,@DOCPATH,@DOCKIND)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@DOCPATH", DOCPATH));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
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
        public void UP(string USERS, string DOCPATH, string DOCKIND)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE [WH_SAPDOC] SET DOCPATH=@DOCPATH WHERE USERS=@USERS AND DOCKIND=@DOCKIND");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@DOCPATH", DOCPATH));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
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
        private System.Data.DataTable Getdata(string USERS, string DOCKIND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCPATH from WH_SAPDOC where USERS=@USERS AND DOCKIND=@DOCKIND ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void DOCUPLOAD_Load(object sender, EventArgs e)
        {
          
                 
             USER = fmLogin.LoginID.ToString().Trim();
            System.Data.DataTable G1 = Getdata(USER, "SAPDRS");
            if (G1.Rows.Count > 0)
            {
                textBox1.Text = G1.Rows[0][0].ToString();
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


                    FileInfo info = new FileInfo(file);
                    int IN1 = info.Name.ToString().ToUpper().IndexOf("DRS");
                    string NAME = info.Name.ToString().ToUpper().Trim().Replace(" ", "");

                    if (globals.DBNAME == "達睿生")
                    {
                        if (IN1 == -1)
                        {
                            MessageBox.Show(info.Name.ToString() + " 檔名沒有DRS無法上傳");
                            return;
                        }
                    }
                    if (NAME != "Thumbs.db")
                    {

                        int J1 = NAME.IndexOf(".");
                
            
                            string M1 = NAME.Substring(J1 + 1, NAME.Length - J1 - 1);
                            string M2 = NAME.Substring(0, J1);
                            string DOC = NAME.Substring(0, J1);

                            int f1 = DOC.IndexOf("-");
                            int n;

                            if (f1 != -1)
                            {
                                DOC = NAME.Substring(0, f1);

                            }
                            if (globals.DBNAME == "達睿生")
                            {

                                if (f1 != -1)
                                {
                                    DOC = NAME.Substring(3, f1 - 3);

                                }
                            }

                            string DOCENTRY = "";

                            OBJ = "17";

                            string CARDCODE = "";
                            string TEL1 = "";
                            string CNTCTCODE = "";
                            string SLPCODE = "";

                            if (int.TryParse(DOC, out n))
                            {
                                DOCENTRY = DOC;
                                System.Data.DataTable T1 = GetORDR(DOC);
                        
                                if (T1.Rows.Count > 0 )
                                {
                                    System.Data.DataTable T2 = GetMAXOCLG2(OBJ, DOC);

                                    DataRow dd = T1.Rows[0];
                                    CARDCODE = dd["CARDCODE"].ToString();
                                    TEL1 = dd["TEL1"].ToString();
                                    CNTCTCODE = dd["CNTCTCODE"].ToString();
                                    SLPCODE = dd["SLPCODE"].ToString();


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
                                        }
                                        else
                                        {

                                            System.Data.DataTable H1 = GetATC1(ATCENTRY);
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
                MessageBox.Show(ex.Message);
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
        private System.Data.DataTable GetORDR(string DOC)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CARDCODE,TEL1,ISNULL(T2.CNTCTCODE,0) CNTCTCODE,SLPCODE FROM ORDR T0 ");
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



 
    }
}
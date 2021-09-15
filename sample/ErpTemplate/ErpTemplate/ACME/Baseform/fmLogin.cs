using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.Devices;
using System.IO;
using System.Collections;
namespace ACME
{
    public partial class fmLogin : Form
    {
        string strCn = "Data Source=10.10.1.61;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        public fmLogin()
        {
            InitializeComponent();
            _loginLimit = 2;
            _currentTimes = 0;
        }

        //全域變數
        public static string LoginID = null;
        public static string PWD = null;
        public static string HostName = null;

        public string aa;
        private int _loginLimit;
        private int _currentTimes;

        //private string _HostName;

        public int LoginLimit
        {
            get
            {
                return _loginLimit;
            }
            set
            {
                if (value > 0)
                    _loginLimit = value;
                else
                {
                    _loginLimit = 0;
                    _currentTimes = 0;
                }
            }
        }

        public int CurrentTimes
        {
            get
            {
                return _currentTimes;
            }
        }



        private void PerformLoginTimes()
        {
            if (_loginLimit > 0 && _currentTimes <= _loginLimit)
                _currentTimes++;
        }




       

        private void button1_Click(object sender, EventArgs e)
        {

            //||COM== "PC-013910288"
            Microsoft.VisualBasic.Devices.Computer computer2 = new Computer();
            PerformLoginTimes();
            string COM = computer2.Name.ToString().ToUpper();
            if (COM == "PC-0003861" || COM == "ACMEW08R2AP1" || COM == "PC-013910265" || COM == "PC-013910288")
            {
    
                LoginID = textBox1.Text;
                globals.UserID = LoginID;
                globals.DBNAME = comboBox1.Text;

                aa = textBox1.Text;
                DialogResult = DialogResult.OK;

            }
            else
            {
                if (textBox1.Text.Trim() != "" && textBox2.Text != "")
                {


                    LoginID = textBox1.Text;
                    PWD = textBox2.Text;



                    //if ((AcmeLdapUtils.CheckAdUser(LoginID, PWD)))
                    //{
                    System.Data.DataTable PWDD = GETPWD(textBox1.Text, textBox2.Text);
                    if (PWDD.Rows.Count > 0)
                    {
                        globals.UserID = LoginID;
                        globals.DBNAME = comboBox1.Text;

                        aa = textBox1.Text;
                        DialogResult = DialogResult.OK;


                    }
                    else
                    {
                        if ((AcmeLdapUtils.CheckAdUser(LoginID, PWD)))
                        {

                            globals.UserID = LoginID;
                            globals.DBNAME = comboBox1.Text;

                            aa = textBox1.Text;
                            DialogResult = DialogResult.OK;
                        }
                        else
                        {

                            MessageBox.Show("用戶名或密碼不正確！", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            textBox1.Focus();
                            textBox1.SelectAll();
                        }
                    }

                }
                else
                {
                    MessageBox.Show("請輸入完整！", "信息提示");
                }

                if (CurrentTimes >= LoginLimit)
                {
                    DialogResult = DialogResult.Cancel;
                }
                try
                {
                    Microsoft.VisualBasic.Devices.Computer computer = new Computer();
                    System.Data.DataTable dt1 = GetOrdr2(computer.Name.ToString());

                    if (dt1.Rows.Count >= 1)
                    {

                    }
                    else
                    {

                        UpdateMasterSQL(computer.Name.ToString(), textBox1.Text, textBox2.Text);
                    }
                    System.Data.DataTable dt2 = GetOrdr3(computer.Name.ToString(), textBox1.Text);
                    System.Data.DataTable dt3 = GetOrdr4(computer.Name.ToString(), textBox2.Text);
                    if (dt2.Rows.Count >= 1)
                    {

                    }
                    else
                    {

                        UpdateMasterSQL1(computer.Name.ToString(), textBox1.Text, textBox2.Text);
                    }
                    if (dt3.Rows.Count >= 1)
                    {

                    }
                    else
                    {

                        UpdateMasterSQL1(computer.Name.ToString(), textBox1.Text, textBox2.Text);
                    }


                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("無法連線到資料庫");
                }

        

            }
        }

        //private void COPYJ()
        //{
        //    string OutPutFile = @"\\acmew08r2ap\SAPUPLOAD\ExtensionProperty.tdc";
        //    //C:\Program Files (x86)\SAP\SAP Business One\AddOns\TADC\TGUI_CXN_882_20201123D
        //    string FileName = string.Empty;
        //    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
        //    //string OutPutFile = lsAppDir + "\\" + "ExtensionProperty.tdc";
        //    string OUTPUT2 = @"C:\Program Files (x86)\\SAP\\SAP Business One\\AddOns\\TADC\\TGUI_CXN_882_20201123D\\"  +"ExtensionProperty.tdc";


        //    File.Copy(OutPutFile, OUTPUT2, true);
        
        //}
        public DataTable GetPATH()
        {
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "SELECT PARAM_NO  FROM RMA_PARAMS WHERE PARAM_KIND='COPYPATH2'";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }
        private void UPLOAD2()
        {

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string NewFileName = lsAppDir + "\\更新程式.exe";

            ArrayList ar = GetFileList(lsAppDir);
            foreach (string Fil in ar)
            {
                FileInfo info = new FileInfo(Fil);
                string G1 = info.Name.ToString();
                int T2 = G1.IndexOf("更新程式");
                if (T2 != -1)
                {
                    System.IO.File.Delete(Fil);
                }
            }

            string OutPutFile = GetPATH().Rows[0][0].ToString() + "\\更新程式.exe";
            System.IO.File.Copy(OutPutFile, NewFileName, true);


            System.Diagnostics.Process.Start(NewFileName);
            Close();

        }
        private void frmLoad_Load(object sender, EventArgs e)
        {
            try
            {
                //try
                //{
                //    COPYJ();
                //}
                //catch { }
                Microsoft.VisualBasic.Devices.Computer computer = new Computer();
                string COM = computer.Name.ToString().ToUpper();
                //try
                //{
               

                //    System.Data.DataTable BUT = GETBUY(COM);
                //    if (BUT.Rows.Count > 0)
                //    {

                //        string lsAppDirF = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                //        string NewFileNameF = lsAppDirF + "\\ACME.exe";
                //        string NewFileNameF2 = @"\\acmew08r2ap\SAPUPLOAD\ACME.exe";
                //        FileInfo filessF = new FileInfo(NewFileNameF);
                //        FileInfo filessF2 = new FileInfo(NewFileNameF2);
                //        string FileDateF = filessF.LastWriteTime.ToString();
                //        string FileDateF2 = filessF2.LastWriteTime.ToString();
                //        if (FileDateF != FileDateF2)
                //        {
                //            UPLOAD2();
                //        }


                //    }
                //}
                //catch { }
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string NewFileName = lsAppDir + "\\ACME.exe";
                FileInfo filess = new FileInfo(NewFileName);
                string FileDate = filess.LastWriteTime.ToString();
                Text = "SAP輔助系統 登入者-" + fmLogin.LoginID + " 登入公司-" + globals.DBNAME + " 版本-" + FileDate;




                System.Data.DataTable dt1 = GetOrdr2(computer.Name.ToString());

                if (COM == "PC-000386" || COM == "PC-013910265" || COM == "PC-013910288")
                {
                    textBox1.Text = "lleytonchen";
                }


                if (COM != "ACMEW08R2RDP")
                {
                    if (dt1.Rows.Count >= 1)
                    {
                        DataRow drw = dt1.Rows[0];
                        textBox1.Text = drw["name"].ToString();
                        textBox2.Text = drw["pwd"].ToString();
                    }

                }
                EXEC();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static ArrayList GetFileList(string Dir)
        {
            ArrayList fils = new ArrayList();
            bool Empty = true;
            foreach (string file in Directory.GetFiles(Dir))
            {
                fils.Add(file);
                Empty = false;
            }

            if (Empty)
            {
                if (Directory.GetDirectories(Dir).Length == 0)
                    fils.Add(Dir + @"/");

            }

            foreach (string dirs in Directory.GetDirectories(Dir))
            {
                foreach (object obj in GetFileList(dirs))
                {
                    fils.Add(obj);
                }
            }
            return fils;
        }
        public DataTable GetOrdr4(string hostname, string pwd)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "select * from hostname where hostname=@hostname and pwd=@pwd";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@hostname", hostname));
            command.Parameters.Add(new SqlParameter("@pwd", pwd));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
       
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }
        public DataTable GetOrdr2(string hostname)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "select * from hostname where hostname=@hostname";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@hostname", hostname));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }
        public  System.Data.DataTable GETBUY(string HOSTNAME)
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT HOSTNAME FROM HOSTNAME  T0 ");
            sb.Append(" LEFT JOIN [RIGHT] T1 ON (T0.[name] =T1.Username COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
          //  sb.Append("               WHERE (T1.Category IN ('ACC','SA') OR T1.Username IN ('APPLECHEN','SUNNYWANG','NANCYTSAI','THOMASLIU','jojohsu'))  AND hostname NOT LIKE '%DRS%'   ");
            sb.Append("               WHERE (T1.Category IN ('SHIPBUY','ACC','SA') OR T1.Username IN ('APPLECHEN','SUNNYWANG','THOMASLIU','jojohsu'))  AND hostname NOT LIKE '%DRS%'   ");
            sb.Append(" AND HOSTNAME=@HOSTNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOSTNAME", HOSTNAME));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["shipping_item"];
        }
 
        private void UpdateMasterSQL(string hostname, string name, string pwd)
        {



            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO Hostname (hostname,name,pwd) VALUES (@hostname,@name,@pwd)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@hostname", hostname));
            command.Parameters.Add(new SqlParameter("@name", name));
            command.Parameters.Add(new SqlParameter("@pwd", pwd));

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
       
       



     

        private void UpdateMasterSQL1(string hostname, string name, string pwd)
        {



            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" update Hostname set name=@name,pwd=@pwd where hostname=@hostname");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@hostname", hostname));
            command.Parameters.Add(new SqlParameter("@name", name));
            command.Parameters.Add(new SqlParameter("@pwd", pwd));

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

        public DataTable GetOrdr3(string hostname, string name)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "select * from hostname where hostname=@hostname and name=@name";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@hostname", hostname));
            command.Parameters.Add(new SqlParameter("@name", name));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }
        public DataTable GETPWD(string NAME, string PWD)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "SELECT * FROM HOSTNAME WHERE NAME=@NAME AND PWD=@PWD";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@NAME", NAME));
            command.Parameters.Add(new SqlParameter("@PWD", PWD));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }
        public DataTable GETRIGHT(string USERNAME)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "SELECT RGROUP  FROM RIGHT_GROUP WHERE USERNAME=@USERNAME";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }
        public DataTable GETRIGHT2(string USERNAME)
        {
            SqlConnection connection = new SqlConnection(strCn);
            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='ACOMPANY' ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERNAME", USERNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "hostname");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["hostname"];
        }


        public  void EXEC()
        {
            Microsoft.VisualBasic.Devices.Computer computer3 = new Computer();
            string COM = computer3.Name.ToString().ToUpper();
            if (COM == "PC-000386" || COM == "PC-013910265")
            {
                System.Data.DataTable dt3 = GETRIGHT(textBox1.Text);
                if (dt3.Rows.Count == 0)
                {


                    dt3 = GETRIGHT2("ACOMPANY");
                    comboBox1.Items.Clear();


                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
                    }
                    comboBox1.Text = Convert.ToString(dt3.Rows[0][0]);
                }

              //  comboBox1.Text = "船務測試區";
            }
            else
            {

                if (textBox1.Text.Trim() != "" && textBox2.Text != "")
                {


                    LoginID = textBox1.Text;
                    PWD = textBox2.Text;



                    System.Data.DataTable PWDD = GETPWD(textBox1.Text, textBox2.Text);
                    if (PWDD.Rows.Count > 0)
                    {
                        System.Data.DataTable dt3 = GETRIGHT(textBox1.Text);
                        if (dt3.Rows.Count == 0)
                        {


                            dt3 = GETRIGHT2("ACOMPANY");

                        }

                        comboBox1.Items.Clear();


                        for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                        {
                            comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
                        }
                        comboBox1.Text = Convert.ToString(dt3.Rows[0][0]);
                    }

   
                }
            }

        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            EXEC();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            EXEC();
        }
      


       
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class fmFTP : Form
    {
        String host = "60.248.27.111";
        String username = "tl_system";
        String password = "GUxaXvyM";

        //ftp://tl_system:GUxaXvyM@202.3.189.166
        
        public fmFTP()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text.ToUpper();



            if (textBox1.Text == "")
            {
                MessageBox.Show("請輸入JOB NO");
                return;
            
            }
            int G2 = textBox1.Text.IndexOf(";");
            if (G2 != -1)
            {
                MessageBox.Show("JOB NO分隔符號請用 , ");
                return;
            }

            if (listBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("請選擇報單");
                return;
            }
            try
            {
                FTPclient ftp = new FTPclient(host, username, password);

                string a0 = listBox1.SelectedItems[0].ToString();
                a0 = a0.Replace("chienshing/", "");
                int G1 = a0.IndexOf(".");

                if (G1 - 13 < 0)
                {
                    MessageBox.Show("報單檔名有問題");
                    return;
                }

                string T1 = a0.Substring(13, G1 - 13);
                string a1 = "/chienshing/" + a0;
                string a2 = @"\\acmesrv01\SAP_Share\shipping\" + DateTime.Now.ToString("yyyyMM") + "\\" + a0;

                ftp.Download(a1, a2, true);

                string strVariable = textBox1.Text.Trim().Replace("\r\n", "");
                string[] keyword = new string[] { "," };
                string[] result = strVariable.Split(keyword, StringSplitOptions.RemoveEmptyEntries);
                foreach (string item in result)
                {
                                int SC = item.IndexOf("SC");
                                if (SC != -1)
                                {
                                    System.Data.DataTable TSC = GetSC(item);
                                    if (TSC.Rows.Count > 0)
                                    {
                                        for (int i = 0; i <= TSC.Rows.Count - 1; i++)
                                        {
                                            DataRow dd = TSC.Rows[i];
                                            string JOBNO = dd["JOBNO"].ToString();
                                     
                                            int R1 = JOBNO.ToUpper().IndexOf("RMA");
                                        
                                            string aa = "";
                                            DataTable t1 = Getdd(JOBNO);
                                            if (R1 != -1)
                                            {
                                                t1 = GetddRMA(JOBNO);
                                            }
                                            if (t1.Rows.Count > 0)
                                            {
                                                aa = t1.Rows[0]["seq"].ToString();
                                            }
                                            else
                                            {
                                                aa = "0";
                                            }

                                            DataTable t2 = GetSHIP(JOBNO);
                                            if (R1 != -1)
                                            {
                                                t2 = GetSHIPRMA(JOBNO);
                                            }
                                            if (t2.Rows.Count == 0)
                                            {
                                                MessageBox.Show("船務系統無此工單號碼");
                                                return;
                                            }
                                            if (R1 != -1)
                                            {
                                                AddRMA(JOBNO, aa, a0, a2);
                                            }
                                            else
                                            {
                                                Add(JOBNO, aa, a0, a2);
                                            }

                                            if (T1.Length == 13)
                                            {
                                                string F2 = T1.Substring(0, 2);
                                                string F3 = T1.Substring(3, 10);
                                                T1 = F2 + "  " + F3;
                                            }
                                            if (R1 != -1)
                                            {
                                                UPDATERMA(JOBNO, T1);
                                            }
                                            else
                                            {
                                                UPDATE(JOBNO, T1);
                                            }
                                        }
                                      
                                    }


                                }
                                else
                                {
                                    string aa = "";
                                    int R1 = item.ToUpper().IndexOf("RMA");
                                    DataTable t1 = Getdd(item);
                                    if (R1 != -1)
                                    {
                                        t1 = GetddRMA(item);
                                    }
                                    if (t1.Rows.Count > 0)
                                    {
                                        aa = t1.Rows[0]["seq"].ToString();
                                    }
                                    else
                                    {
                                        aa = "0";
                                    }
                       
                                    DataTable t2 = GetSHIP(item);
                                    if(R1 != -1)
                                    {
                                        t2 = GetSHIPRMA(item);
                                    }
                                    if (t2.Rows.Count == 0)
                                    {
                                        MessageBox.Show("船務系統無此工單號碼");
                                        return;
                                    }

                                    if (R1 != -1)
                                    {
                                        AddRMA(item, aa, a0, a2);
                                    }
                                    else
                                    {
                                        Add(item, aa, a0, a2);
                                    }

                                    if (T1.Length == 13)
                                    {
                                        string F2 = T1.Substring(0, 2);
                                        string F3 = T1.Substring(3, 10);
                                        T1 = F2 + "  " + F3;
                                    }

                                    if (R1 != -1)
                                    {
                                        UPDATERMA(item, T1);
                                    }
                                    else
                                    {
                                        UPDATE(item, T1);
                                    }
                                }
          
                }




                if (ftp.FtpFileExists(a1))
                {

                    ftp.FtpDelete(a1);
                }

                textBox1.Text = "";
                LOAD();


                MessageBox.Show("OK");
            }
            catch (Exception ex)
            {
                MessageBox.Show("無法連線 " + ex.Message);
            }




        }

        private void button2_Click(object sender, EventArgs e)
        {
            FTPclient ftp = new FTPclient(host, username, password);

            string a1 = "/chienshing/" + listBox1.SelectedItems[0].ToString();
            if (ftp.FtpFileExists(a1))
            {
                //'rename a file
                //ftp.FtpRename("/pub/upload.exe", "/pub/newname.exe")
                //'delete a file
                ftp.FtpDelete(a1);
            }



            MessageBox.Show("OK");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FTPclient ftp = new FTPclient(host, username, password);
            string a1 = "/chienshing/" + listBox1.SelectedItems[0].ToString();

            if (ftp.FtpFileExists(a1))
            {

                ftp.FtpDelete(a1);
            }


          

            MessageBox.Show("OK");
        }

        private void button4_Click(object sender, EventArgs e)
        {


       

        }

        public void Add(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into DOWNLOAD2(shippingcode,seq,[filename],[path],MARK) values(@shippingcode,@seq,@filename,@path,@MARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@MARK", "1"));

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

        public void AddRMA(string shippingcode, string seq, string filename, string path)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Rma_Download(shippingcode,seq,[filename],[path],MARK) values(@shippingcode,@seq,@filename,@path,@MARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@MARK", "1"));

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
        public void UPDATE(string SHIPPINGCODE, string ADD9)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE SHIPPING_MAIN SET ADD9=@ADD9 WHERE SHIPPINGCODE=@SHIPPINGCODE ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ADD9", ADD9));


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
        public void UPDATERMA(string SHIPPINGCODE, string ADD1)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE RMA_MAIN SET ADD1=@ADD1 WHERE SHIPPINGCODE=@SHIPPINGCODE ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ADD1", ADD1));


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
        public System.Data.DataTable GetSC(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT JOBNO FROM shipping_CAR2 WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable Getdd(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT seq+1 seq FROM DOWNLOAD2 where shippingcode=@shippingcode order by cast(seq as int) desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable GetddRMA(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT seq+1 seq FROM RMA_DOWNLOAD where shippingcode=@shippingcode order by cast(seq as int) desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable GetSHIP(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT shippingcode FROM SHIPPING_MAIN where shippingcode=@shippingcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable GetSHIPRMA(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT shippingcode FROM RMA_MAIN where shippingcode=@shippingcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        private void fmFTP_Load(object sender, EventArgs e)
        {

            LOAD();
          
        }

        private void LOAD()
        {
            try
            {
                FTPclient ftp = new FTPclient(host, username, password);


                List<string> l = ftp.ListDirectory("/chienshing");

                listBox1.Items.Clear();

                for (int i = 0; i <= l.Count - 1; i++)
                {
                    listBox1.Items.Add(l[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "請擷取錯誤訊息圖片，連絡MIS");
                Close(); 
            }
        }

  



       
    }
}




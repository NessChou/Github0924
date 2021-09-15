using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Web.UI;
using System.Net.Mail;
using System.Net.Mime;
namespace ACME
{
    public partial class WH_INVOICE : Form
    {
        public WH_INVOICE()
        {
            InitializeComponent();
        }

        private void F1(string DOCTYPE,DataGridView  D1)
        {
            System.Data.DataTable TempDt = MakeTable();
            System.Data.DataTable dt = null;
            if (DOCTYPE == "3")
            {
                dt = Getdata3();
            }
            else
            {
                dt = Getdata2(DOCTYPE);
            }
            DataRow dr = null;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = TempDt.NewRow();
                //            sb.Append("							  ,T1.U_EMAILYES 寄送電子發票紙本,T1.U_EMAIL2 多組接受發票EMAIL");

                dr["單號"] = dt.Rows[i]["單號"].ToString();
                dr["發票日期"] = dt.Rows[i]["發票日期"].ToString();
                dr["客戶編號"] = dt.Rows[i]["客戶編號"].ToString();
                dr["統一編號"] = dt.Rows[i]["統一編號"].ToString();
                dr["SA"] = dt.Rows[i]["SA"].ToString();
                dr["寄送電子發票紙本"] = dt.Rows[i]["寄送電子發票紙本"].ToString();
                dr["多組接受發票EMAIL"] = dt.Rows[i]["多組接受發票EMAIL"].ToString();
                string INV = dt.Rows[i]["發票號碼"].ToString();
                dr["發票號碼"] = INV;
                dr["客戶名稱"] = dt.Rows[i]["客戶名稱"].ToString();
                dr["EMAIL"] = dt.Rows[i]["EMAIL"].ToString();
                dr["中華票服日期"] = dt.Rows[i]["中華票服日期"].ToString();
                string TIME = dt.Rows[i]["中華票服時間"].ToString();
                string TIME2 = TIME.Substring(0, 2) + ":" + TIME.Substring(2, 2);
                dr["中華票服時間"] = TIME2;
                dr["中華票服狀態"] = dt.Rows[i]["中華票服狀態"].ToString();

                string Url = "https://api.cxn.com.tw/get_invoice_status_incsv.php";
                string postString = "id=89206602&user=89206602&passwd=1qaz2wsx&invoice=" + INV;
                byte[] postData = Encoding.UTF8.GetBytes(postString);
                WebClient client = new WebClient();
                client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                byte[] responseData = client.UploadData(Url, "POST", postData);
                string fa = Encoding.UTF8.GetString(responseData);
                      
                 int  sa = fa.IndexOf("status=");
                 string fa2 = fa.Substring(sa + 7, fa.Length - sa - 7);
                 string fa3 = Encoding.UTF8.GetString(Convert.FromBase64String(fa2));

                 String[] split = fa3.Split('\"');
           
                        StringBuilder sb = new StringBuilder();
                        int t1 = 0;
                        foreach (String F in split)
                        {
                            t1++;
                            string V1 = "";
                            if (t1 == 4)
                            {
                                 V1 = F.ToString();
                                 if (F.ToString() == "C0401")
                                 {
                                     V1 = "開立";
                                 }
                                 if (F.ToString() == "C0501")
                                 {
                                     V1 = "作廢";
                                 }
                                 if (F.ToString() == "C0701")
                                 {
                                     V1 = "註銷";
                                 }
                                 if (F.ToString() == "D0401")
                                 {
                                     V1 = "開立折讓";
                                 }
                                 if (F.ToString() == "D0501")
                                 {
                                     V1 = "作廢折讓";
                                 }
                                 dr["財政部發票類別"] = V1;
                            }
                            if (t1 == 8)
                            {

                                dr["財政部發票狀態"] = F.ToString();
                            }

                        }
                TempDt.Rows.Add(dr);
            }

            D1.DataSource = TempDt;

 
        }
        private void WH_INVOICE_Load(object sender, EventArgs e)
        {
            FS();
 
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
        }
        private System.Data.DataTable Getdata2(string DTYPE)
        {

            SqlConnection connection = globals.shipConnection ;

            StringBuilder sb = new StringBuilder();
            sb.Append("                                               SELECT T0.DOCENTRY 單號, T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,U_IN_BSINV  發票號碼, Convert(varchar(8),T0.DOCDATE,112) 發票日期,T0.U_GUI_EMAIL EMAIL,     ");
            sb.Append("                                                        Convert(varchar(10),U_PLTUPDDATE,111)   中華票服日期,   ");
            sb.Append("                                       				 CASE WHEN (LEN(CAST(U_PLTUPDTIME AS VARCHAR))=3)  THEN '0'+CAST(U_PLTUPDTIME AS VARCHAR) ELSE CAST(U_PLTUPDTIME AS VARCHAR) END    ");
            sb.Append("                                       				  中華票服時間     ");
            sb.Append("                                                        ,CASE T0.U_CXNUploadStatus WHEN 0 THEN '未上傳' WHEN '1' THEN '上傳成功'  WHEN '2' THEN '上傳失敗'  WHEN '3' THEN '上傳成功' END 中華票服狀態      ");
            sb.Append("                          							  ,T1.U_EMAILYES 寄送電子發票紙本,CASE WHEN  ISNULL(T0.U_EMAIL2,'')<>'' THEN  T0.U_EMAIL2 ELSE  T1.U_EMAIL2 END  多組接受發票EMAIL,T1.LICTRADNUM 統一編號, lastname+firstname SA ");
            sb.Append("                                                        FROM OINV  T0       ");
            sb.Append("                          							  LEFT JOIN OCRD T1 ON (T0.CARDCODE=T1.CARDCODE)  ");
            sb.Append("             										  iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("                                                        WHERE T0.DOCENTRY>43237  AND U_IN_BSINV <>'__________'      ");
            sb.Append("                                                        AND ISNULL(U_PLTUPDTIME,'') <> ''     ");

            if (DTYPE == "1")
            {
                sb.Append(" AND ISNULL(T1.LICTRADNUM,'') <>'' ");
            }
            else
            {
                sb.Append(" AND ISNULL(T1.LICTRADNUM,'') ='' ");
            }
   
                sb.Append(" AND Convert(varchar(8),T0.DOCDATE,112) BETWEEN @DOCDATE AND @DOCDATE1 ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T0.CARDNAME  like '%" + textBox3.Text + "%'   ");
            }
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                sb.Append(" AND  (T0.[cardcode] >=@CARD1 AND  T0.[cardcode] <=@CARD2 )  ");
            }

            if (textBox6.Text != "")
            {
                sb.Append(" AND T0.DOCENTRY =@DOCENTRY  ");
            }

            if (textBox7.Text != "")
            {
                sb.Append(" AND T0.U_IN_BSINV =@U_IN_BSINV  ");
            }
            if (textBox8.Text != "")
            {
                sb.Append(" AND T0.DOCENTRY IN (SELECT DOCENTRY FROM INV1 WHERE BaseType =15 AND BASEREF=@BASEREF) ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CARD1", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@CARD2", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@U_IN_BSINV", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BASEREF", textBox8.Text));
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
        private System.Data.DataTable Getdata3()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                                               SELECT T0.DOCENTRY 單號, T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,U_BSIDF  發票號碼, Convert(varchar(8),T0.DOCDATE,112) 發票日期,T0.U_GUI_EMAIL EMAIL,     ");
            sb.Append("                                                        Convert(varchar(10),U_PLTUPDDATE,111)   中華票服日期,   ");
            sb.Append("                                       				 CASE WHEN (LEN(CAST(U_PLTUPDTIME AS VARCHAR))=3)  THEN '0'+CAST(U_PLTUPDTIME AS VARCHAR) ELSE CAST(U_PLTUPDTIME AS VARCHAR) END    ");
            sb.Append("                                       				  中華票服時間     ");
            sb.Append("                                                        ,CASE T0.U_CXNUploadStatus WHEN 0 THEN '未上傳' WHEN '1' THEN '上傳成功'  WHEN '2' THEN '上傳失敗'  WHEN '3' THEN '上傳成功' END 中華票服狀態      ");
            sb.Append("                          							  ,T1.U_EMAILYES 寄送電子發票紙本,T1.U_EMAIL2 多組接受發票EMAIL,T1.LICTRADNUM 統一編號, lastname+firstname SA ");
            sb.Append("                                                        FROM ORIN  T0       ");
            sb.Append("                          							  LEFT JOIN OCRD T1 ON (T0.CARDCODE=T1.CARDCODE)  ");
            sb.Append("             										  iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("                                                        WHERE  U_BSIDF <>'__________'      ");
            sb.Append("                                           AND ISNULL(U_PLTUPDTIME,'') <> ''     ");




            sb.Append(" AND Convert(varchar(8),T0.DOCDATE,112) BETWEEN @DOCDATE AND @DOCDATE1 ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T0.CARDNAME  like '%" + textBox3.Text + "%'   ");
            }
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                sb.Append(" AND  (T0.[cardcode] >=@CARD1 AND  T0.[cardcode] <=@CARD2 )  ");
            }



            if (textBox7.Text != "")
            {
                sb.Append(" AND T0.U_BSIDF =@U_IN_BSINV  ");
            }
   
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CARD1", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@CARD2", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@U_IN_BSINV", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BASEREF", textBox8.Text));
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
        string NewFileName = "";
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {

                    string 發票號碼 = dataGridView1.CurrentRow.Cells["發票號碼"].Value.ToString();
          
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                     NewFileName = lsAppDir + "\\EXCEL\\temp\\" + 發票號碼+".PDF";

           

                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;
                    
                    string Url = "https://api.cxn.com.tw/get_invoice_pdf.php";
                    string postString = "id=89206602&user=89206602&passwd=1qaz2wsx&invoice=" + 發票號碼 + "&type=2";
                    byte[] postData = Encoding.UTF8.GetBytes(postString);
                    WebClient client = new WebClient();
                    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] responseData = client.UploadData(Url, "POST", postData);
                    FileStream fs = new FileStream(NewFileName, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(responseData);
                    bw.Close();
                    fs.Close();


                    System.Diagnostics.Process.Start(NewFileName);
                }

                if (dgv.Columns[e.ColumnIndex].Name == "SEND")
                {

                    string EMAIL = dataGridView1.CurrentRow.Cells["EMAIL"].Value.ToString();

                    string 發票號碼 = dataGridView1.CurrentRow.Cells["發票號碼"].Value.ToString();

                    if (!String.IsNullOrEmpty(EMAIL))
                    {

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                        NewFileName = lsAppDir + "\\EXCEL\\temp\\進金生實業" + 發票號碼 + ".PDF";



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;

                        string Url = "https://api.cxn.com.tw/get_invoice_pdf.php";
                        string postString = "id=89206602&user=89206602&passwd=1qaz2wsx&invoice=" + 發票號碼 + "&type=2";
                        byte[] postData = Encoding.UTF8.GetBytes(postString);
                        WebClient client = new WebClient();
                        client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                        byte[] responseData = client.UploadData(Url, "POST", postData);
                        FileStream fs = new FileStream(NewFileName, FileMode.Create);
                        BinaryWriter bw = new BinaryWriter(fs);
                        bw.Write(responseData);
                        bw.Close();
                        fs.Close();

                        string template;
                        StreamReader objReader;
                        string FileName = string.Empty;


                        FileName = lsAppDir + "\\MailTemplates\\SA.htm";
                        objReader = new StreamReader(FileName);

                        template = objReader.ReadToEnd();
                        objReader.Close();
                        objReader.Dispose();

                        StringWriter writer = new StringWriter();
                        HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                        template = template.Replace("##Content##", "");




                        MailMessage message = new MailMessage();

                        int G = 0;

                        //G = 1;

                        if (G == 1)
                        {
                            EMAIL = "LLEYTONCHEN@ACMEPOINT.COM";
                        }
                        message.To.Add(new MailAddress(EMAIL));

                        message.Subject = "進金生實業統一電子發票";
                        message.Body = template;
                        string m_File = NewFileName;
                        Attachment data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                        //附件资料
                        ContentDisposition disposition = data.ContentDisposition;


                        // 加入邮件附件
                        message.Attachments.Add(data);
                        message.IsBodyHtml = true;

                        SmtpClient client2 = new SmtpClient();
                        try
                        {
                            client2.Send(message);

                            MessageBox.Show("寄信成功");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK2")
                {

                    string 發票號碼 = dataGridView2.CurrentRow.Cells["發票號碼2"].Value.ToString();

                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    NewFileName = lsAppDir + "\\EXCEL\\temp\\" + 發票號碼 + ".PDF";



                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;

                    string Url = "https://api.cxn.com.tw/get_allowance_pdf.php";
                    string postString = "id=89206602&user=89206602&passwd=1qaz2wsx&allowance=" + 發票號碼 + "&type=2";
                    byte[] postData = Encoding.UTF8.GetBytes(postString);
                    WebClient client = new WebClient();
                    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] responseData = client.UploadData(Url, "POST", postData);
                    FileStream fs = new FileStream(NewFileName, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(responseData);
                    bw.Close();
                    fs.Close();


                    System.Diagnostics.Process.Start(NewFileName);
                }

                if (dgv.Columns[e.ColumnIndex].Name == "SEND2")
                {

                    string EMAIL = dataGridView2.CurrentRow.Cells["EMAIL2"].Value.ToString();

                    string 發票號碼 = dataGridView2.CurrentRow.Cells["發票號碼2"].Value.ToString();

                    if (!String.IsNullOrEmpty(EMAIL))
                    {

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                        NewFileName = lsAppDir + "\\EXCEL\\temp\\進金生實業" + 發票號碼 + ".PDF";



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;

                        string Url = "https://api.cxn.com.tw/get_allowance_pdf.php";
                        string postString = "id=89206602&user=89206602&passwd=1qaz2wsx&allowance=" + 發票號碼 + "&type=2";
                        byte[] postData = Encoding.UTF8.GetBytes(postString);
                        WebClient client = new WebClient();
                        client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                        byte[] responseData = client.UploadData(Url, "POST", postData);
                        FileStream fs = new FileStream(NewFileName, FileMode.Create);
                        BinaryWriter bw = new BinaryWriter(fs);
                        bw.Write(responseData);
                        bw.Close();
                        fs.Close();

                        string template;
                        StreamReader objReader;
                        string FileName = string.Empty;


                        FileName = lsAppDir + "\\MailTemplates\\SA.htm";
                        objReader = new StreamReader(FileName);

                        template = objReader.ReadToEnd();
                        objReader.Close();
                        objReader.Dispose();

                        StringWriter writer = new StringWriter();
                        HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                        template = template.Replace("##Content##", "");




                        MailMessage message = new MailMessage();

                        int G = 0;

                        //G = 1;

                        if (G == 1)
                        {
                            EMAIL = "LLEYTONCHEN@ACMEPOINT.COM";
                        }
                        message.To.Add(new MailAddress(EMAIL));

                        message.Subject = "進金生實業統一電子發票";
                        message.Body = template;
                        string m_File = NewFileName;
                        Attachment data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                        //附件资料
                        ContentDisposition disposition = data.ContentDisposition;


                        // 加入邮件附件
                        message.Attachments.Add(data);
                        message.IsBodyHtml = true;

                        SmtpClient client2 = new SmtpClient();
                        try
                        {
                            client2.Send(message);

                            MessageBox.Show("寄信成功");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        //private void button2_Click(object sender, EventArgs e)
        //{
        //    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
        //    string inv = "AV82000051";
        //    NewFileName = lsAppDir + "\\EXCEL\\temp\\" + inv + ".csv";
        //    string s1 = lsAppDir + "\\EXCEL\\temp2\\1.csv";
        //    string Url = "https://api-test.cxn.com.tw/get_invoice_status_incsv.php";
        //    string postString = "id=89206602&user=admin&passwd=89206602&invoice=" + inv;
        //    byte[] postData = Encoding.UTF8.GetBytes(postString);
        //    WebClient client = new WebClient();
        //    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
        //    byte[] responseData = client.UploadData(Url, "POST", postData);
        //    string fa = Encoding.UTF8.GetString(responseData);
        //    //FileStream fs = new FileStream(NewFileName, FileMode.Create);
        //    //BinaryWriter bw = new BinaryWriter(fs);
        //    //bw.Write(responseData);
        //    //bw.Close();
        //    //fs.Close();


        //    //System.Diagnostics.Process.Start(fa);
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            FS();
        }
        private void FS()
        {
            F1("1", dataGridView1);
            F1("3", dataGridView2);
            F1("2", dataGridView3);
        }
        private System.Data.DataTable MakeTable()
        {


            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("單號", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));

            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("統一編號", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("EMAIL", typeof(string));
            dt.Columns.Add("中華票服日期", typeof(string));
            dt.Columns.Add("中華票服時間", typeof(string));
            dt.Columns.Add("中華票服狀態", typeof(string));
            dt.Columns.Add("財政部發票類別", typeof(string));
            dt.Columns.Add("財政部發票狀態", typeof(string));
            dt.Columns.Add("寄送電子發票紙本", typeof(string));
            dt.Columns.Add("多組接受發票EMAIL", typeof(string));
            dt.Columns.Add("SA", typeof(string));

            return dt;
        }

        private void linkLabel4_Click(object sender, EventArgs e)
        {
            string d = @"\\acmesrv01\Public\SAP電子發票測試\操作SOP";

        
                string[] filenames = Directory.GetFiles(d);
      
                foreach (string file in filenames)
                {
                    System.Diagnostics.Process.Start(file);
                }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView1.Rows.Count - 1)
                return;
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
            try
            {

                //財政部發票狀態
                if (dgr.Cells["中華票服狀態"].Value.ToString() == "上傳失敗")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
        }

   
    }
}

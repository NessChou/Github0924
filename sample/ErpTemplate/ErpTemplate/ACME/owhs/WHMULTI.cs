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
namespace ACME
{

    public partial class WHMULTI : Form
    {
        string MAILSUB = "";
        System.Net.Mail.Attachment data = null;
        public WHMULTI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            AA("Y");
        }
        private void AA(string FLAG)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            try
            {
                try
                {
                    ArrayList al = new ArrayList();

                    for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                    {
                        al.Add(listBox2.Items[i].ToString());
                    }

                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                        sb2.Append("" + v + "+");
                    }
                    sb.Remove(sb.Length - 1, 1);
                    sb2.Remove(sb2.Length - 1, 1);
                }
                catch { }

                if (sb.Length > 0)
                {
                    string OWHS = comboBox2.Text;
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
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    FileName = lsAppDir + "\\Excel\\wh\\收貨單M.xls";

                    System.Data.DataTable OrderData = Getprepare3(sb.ToString(), sb2.ToString(), OWHS1);

                    string ExcelTemplate = FileName;


                    string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\" +
                       OWHS1 + "收貨通知單---" + sb2.ToString() + ".xls";
                           string INV="";
                    string INV2 = GETINV(sb.ToString()).Rows[0][0].ToString();
                    if (!String.IsNullOrEmpty(INV2))
                    {
                        INV = "--" + INV2;
                    }
                    MAILSUB = OWHS1 + "收貨通知單---" + sb2.ToString() + "--" + GETQTY(sb.ToString()).Rows[0][0].ToString() + "片" + INV;

                    ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, FLAG);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public System.Data.DataTable Getprepare3(string docentry, string DOCNAME,string OWHS1)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" select Convert(varchar(10),Getdate(),111) 日期,RANK() OVER (ORDER BY t0.shippingcode,t1.itemcode,T1.SEQNO ) AS 項次,t1.itemcode 產品編號,t1.dscription 品名規格,");
            //sb.Append(" case isnull(t1.pino,'') when '' then t1.dscription else t1.pino end COLLATE Chinese_Taiwan_Stroke_CI_AS 料號,t1.quantity 出貨數量,'" + OWHS1 + "'+");
            //sb.Append(" '收貨通知單---'+@DOCNAME 文件名稱,T0.CARDNAME 預出客戶,t0.closeday 申請日期,");
            //sb.Append(" t1.grade 等級,t1.ver 版本,BOXCHECK 外箱檢查,t0.boardCount SHI,T1.LOCATION 產地 ");
            //sb.Append(" from wh_main t0 left join wh_item t1 on (t0.shippingcode=t1.shippingcode) where t0.shippingcode IN (" + docentry + ") order by t0.shippingcode,t1.itemcode ");

            sb.Append(" select Convert(varchar(10),Getdate(),111) 日期,RANK() OVER (ORDER BY t0.shippingcode,t1.itemcode,T1.SEQNO ) AS 項次,t1.itemcode 產品編號,t1.dscription 品名規格,");
            sb.Append(" case isnull(t1.pino,'') when '' then t1.dscription else t1.pino end COLLATE Chinese_Taiwan_Stroke_CI_AS 料號,t1.quantity 出貨數量,'" + OWHS1 + "'+");
            sb.Append(" '收貨通知單---'+@DOCNAME 文件名稱,T0.CARDNAME 預出客戶,t0.closeday 申請日期,");
            sb.Append(" t1.grade 等級,t1.ver 版本,Invoice  外箱檢查,t0.boardCount SHI,T1.LOCATION 產地 ");
            sb.Append(" from wh_main t0 left join wh_item t1 on (t0.shippingcode=t1.shippingcode) where t0.shippingcode IN (" + docentry + ") order by t0.shippingcode,t1.itemcode ");
            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));

            command.Parameters.Add(new SqlParameter("@DOCNAME", DOCNAME));
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


        public System.Data.DataTable GETINV(string docentry)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @name varchar(100) ");
            sb.Append(" select @name =SUBSTRING(COALESCE(@name + '/',''),0,99) + SendGoods ");
            sb.Append(" from   (select SendGoods  from wh_main t0 left join wh_item3 t1 on (t0.shippingcode=t1.shippingcode) ");
            sb.Append(" where t0.shippingcode IN (" + docentry + ") ");
            sb.Append(" AND ISNULL(SendGoods,'') <> '') pc");
            sb.Append(" SELECT ISNULL(@name,'') INV");

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
        public System.Data.DataTable GETQTY(string docentry)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) QTY   from  wh_item t1 where T1.shippingcode IN (" + docentry + ")   ");


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

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text == "")
            {
                MessageBox.Show("請選擇倉庫");
                return;
            }
            
            System.Data.DataTable T1 = GetSHIP();

            if (T1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }
        }

        private System.Data.DataTable GetSHIP()
        {
            SqlConnection MyConnection = globals.Connection;
                        StringBuilder sb = new StringBuilder();

                        sb.Append(" SELECT DISTINCT t0.SHIPPINGCODE FROM WH_MAIN  t0 INNER join wh_item t1 on (t0.shippingcode=t1.shippingcode)   WHERE SUBSTRING(t0.SHIPPINGCODE,3,8) BETWEEN @AA AND @BB ORDER BY t0.SHIPPINGCODE");
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }


        private void SHIPMULTI_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();

            System.Data.DataTable dt4 = null;

            if (globals.DBNAME == "進金生" || globals.DBNAME == "達睿生" || globals.DBNAME == "測試區98")
            {
                dt4 = GetMenu.Getwarehouse();
            }
            else if (globals.DBNAME == "宇豐")
            {
                dt4 = GetMenu.GetwarehouseAD();
            }
            else if (globals.DBNAME == "CHOICE" || globals.DBNAME == "INFINITE" || globals.DBNAME == "TOP GARDEN")
            {
                dt4 = GetMenu.GetwarehouseCHI();
            }
            comboBox2.Items.Clear();

            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][1]));
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                DELETEFILE();


                string SEMAIL = "";
                string CEMAIL = "";

                System.Data.DataTable GETMAIL = GetMenu.GetWHNAIL(comboBox2.Text);
                if (GETMAIL.Rows.Count > 0)
                {
                    SEMAIL = GETMAIL.Rows[0]["SEMAIL"].ToString();
                    CEMAIL = GETMAIL.Rows[0]["CEMAIL"].ToString();
                }
                else
                {

                    return;
                }
                AA("N");
                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\WHMAIN.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();
                MailMessage message = new MailMessage();
                message.From = new MailAddress("workflow@acmepoint.com", "系統發送");

                //      message.To.Add(new MailAddress("lleytonchen@acmepoint.com"));
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

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);


                template = template.Replace("##Content1##", "請依收貨工單對照型號/等級/數量/產地/版本，並將ACME15碼打進庫存表 & 提供進貨序號以利核對 , 謝謝。");
                template = template.Replace("##Content2##", " P.S. ");
                template = template.Replace("##Content3##", "1.請務必於點收進貨確認無誤後將收貨工單簽名回覆");
                template = template.Replace("##Content4##", "2.進貨如有異常請於回傳收貨工單時寫清楚INV#/異常幾板／幾箱");
                template = template.Replace("##Content5##", "3.貨代送貨到時,請當場對點完於簽收單上簽名,如有異常請載明異常點,以利確認責任歸屬!!!");

                template = template.Replace("##eng##", "");
                template = template.Replace("##name##", "");
                template = template.Replace("##mail##", "");



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


                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                client.Host = "ms.mailcloud.com.tw";
                client.UseDefaultCredentials = true;
                string pwd = "@cmeworkflow";

                client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);



                try
                {
                    client.Send(message);
                    data.Dispose();
                    message.Attachments.Dispose();

                    DELETEFILE();
                    //  MessageBox.Show("寄信成功");
                }
                catch (SmtpFailedRecipientsException ex)
                {

                }
                catch (Exception ex)
                {

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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}

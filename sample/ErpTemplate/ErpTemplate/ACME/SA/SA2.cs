using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;
namespace ACME
{
    public partial class SA2 : Form
    {
        System.Net.Mail.Attachment data = null;
        public SA2()
        {
            InitializeComponent();
        }

        private void SA2_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GETORTT();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result;
                result = MessageBox.Show("請確認是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        DataGridViewRow row;

                        row = dataGridView1.SelectedRows[i];

     
                        string MAILTO = row.Cells["EMAIL"].Value.ToString();

                        string USER = fmLogin.LoginID.ToString();
                
                        if (checkBox1.Checked)
                        {
                            MAILTO = USER + "@acmepoint.com";
                        }
                        MailTest("進金生電子發票通知函(此為系統自動通知信，請勿直接回信)", MAILTO, textBox1.Text);
                        MessageBox.Show("信件已寄出");
                    }
                }
            }
        }

        private void MailTest(string strSubject, string MailAddress, string template)
        {
            MailMessage message = new MailMessage();
            string USER = fmLogin.LoginID.ToString();
            string FORM = USER + "@acmepoint.com";
            message.From = new MailAddress(FORM, "進金生發送");
            message.To.Add(new MailAddress(MailAddress));


            message.Subject = strSubject;
            message.Body = template;
            message.IsBodyHtml = true;


            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;


            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            string pwd = "@cmeworkflow";
            string OutPutFile = lsAppDir + "\\Excel\\電子發票通知函(新).docx";

            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            string m_File = OutPutFile;

            data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

            //附件资料
            ContentDisposition disposition = data.ContentDisposition;


            // 加入邮件附件
            message.Attachments.Add(data);

            try
            {
                client.Send(message);


            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        MessageBox.Show("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        MessageBox.Show(String.Format("Failed to deliver message to {0}",
                            ex.InnerExceptions[i].FailedRecipient));

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Exception caught in RetryIfBusy(): {0}",
                        ex.ToString()));

            }

        }
        private System.Data.DataTable GETORTT()
        {

            SqlConnection connection = globals.shipConnection;
            string USER = fmLogin.LoginID.ToString().ToUpper();
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT t0.cardcode 客戶編號,t0.cardname 客戶名稱,T0.LicTradNum 統編, t2.slpname 業務 ,T1.HOMETEL 業助,T0.ADDRESS 收款方地址 ,T0.mailaddres 收貨方地址,phone1,phone2,TS.E_MailL EMAIL FROM OCRD T0");
            sb.Append("  LEFT JOIN OSLP T2 ON (T0.SlpCode = T2.SlpCode)");
            sb.Append("  LEFT JOIN OCRG T3 ON (T0.GROUPCODE = T3.GROUPCODE)");
            sb.Append("  LEFT JOIN OCPR TS ON (T0.CntctPrsn  = TS.Name AND T0.CARDCODE=TS.CARDCODE )");
            sb.Append("  LEFT JOIN (SELECT CARDCODE,MAX(OWNERCODE) OWNERCODE FROM ORDR  where OWNERCODE in (select empid from ohem where isnull(termdate,'') =  '' )");
            sb.Append("  GROUP BY CARDCODE) T4 ON (T0.CARDCODE=T4.CARDCODE)");
            sb.Append("  LEFT JOIN OHEM T1 ON (T0.DfTcnician=T1.EMPID)");
            sb.Append("  where cardtype='c'  and t0.CARDCODE in (SELECT distinct CARDCODE FROM OINV WHERE YEAR(DOCDATE) between '2018' and '2020')");
            if (USER == "PATTYLIU")
            {
                sb.Append("  and  SUBSTRING(T3.GROUPNAME,4,15)='ESCO' AND T0.LicTradNum <>''");
            }
            else
            {
                sb.Append("  and  SUBSTRING(T3.GROUPNAME,4,15)='TFT' AND T0.LicTradNum <>''");
            }
            sb.Append("   ORDER BY T1.HOMETEL,t0.cardcode");


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

    }
}

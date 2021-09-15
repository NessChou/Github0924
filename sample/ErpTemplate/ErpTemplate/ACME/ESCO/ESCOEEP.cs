using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
namespace ACME
{
    public partial class ESCOEEP : Form
    {
        Attachment data = null;
        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public ESCOEEP()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetOrderDataAPL()
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                  SELECT FLOW_DESC 流程,LTRIM(LTRIM(REPLACE(REPLACE(REPLACE(FORM_PRESENTATION,'=',''),'DOCENTRY',''),'ID',''))) 單號, ");
            sb.Append("               D_STEP_ID 流程階段,MAILHEAD,MAILTEMP,MAILTO ");
            sb.Append("               FROM SYS_TODOLIST T0  ");
            sb.Append("                             LEFT JOIN ACME_MAIL_BACKUP2 T1 ON (T0.LISTID=T1.MAILDOC AND T0.D_STEP_ID=T1.FLOWTYPE)  ");
            sb.Append("                              WHERE STATUS='N'   and FLOW_DESC IN(");
            sb.Append(" SELECT CAPTION+'(ESCO)'  FROM MENUTABLE WHERE FORM LIKE '%ESCO%' AND  PACKAGE='FL' ) ORDER BY FLOW_DESC  ");


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

        private void SAEEP_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOrderDataAPL();
        }


        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result;
                result = MessageBox.Show("請確認是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[0];
                    string MAILHEAD = "[提醒簽核]" + row.Cells["MAILHEAD"].Value.ToString();
                    string MAILTO = row.Cells["MAILTO"].Value.ToString();
                    string MAILTEMP = row.Cells["MAILTEMP"].Value.ToString();
                    string F1 = "";

                    F1 = "1";
                    if (F1 == "1")
                    {
                        MAILTO = "LLEYTONCHEN@ACMEPOINT.COM";
                    }
                    MailTest(MAILHEAD, MAILTO, MAILTEMP);
                    MessageBox.Show("信件已寄出");
                }
            }
        }


        private void MailTest(string strSubject, string MailAddress, string template)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));



            message.Subject = strSubject;
            message.Body = template;
            message.IsBodyHtml = true;


            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";


            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

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
    }
}

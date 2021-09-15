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
    public partial class HREEP : Form
    {
        Attachment data = null;
        string strCn02 = "Data Source=10.10.1.47;Initial Catalog=AcmeEEP;Persist Security Info=True;User ID=NewType;Password=NewType";
        public HREEP()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetOrderDataAPL()
        {
            SqlConnection connection = new SqlConnection(strCn02);
            StringBuilder sb = new StringBuilder();
            sb.Append("                  SELECT FLOW_DESC 流程,APPLICANT 申請人,LTRIM(LTRIM(REPLACE(REPLACE(REPLACE(FORM_PRESENTATION,'=',''),'DOCENTRY',''),'ID',''))) 單號,  ");
            sb.Append("                           D_STEP_ID 卡關階段");
            sb.Append("                           FROM SYS_TODOLIST T0   ");
            sb.Append("");
            sb.Append("             							                        ");
            sb.Append("                                          WHERE STATUS='N'   AND FLOW_DESC  IN ('職工福利申請表','外訓申請') AND REPLACE(SUBSTRING(UPDATE_DATE,0,8),'-','')>'202012' ");

       
            sb.Append("               ORDER BY FLOW_DESC ");


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
            SS();
        }

        private void SS()
        {
            System.Data.DataTable TempDt = MakeTable();

            System.Data.DataTable dtemp5 = GetOrderDataAPL();

            dataGridView1.DataSource = dtemp5;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

      
            dt.Columns.Add("流程", typeof(string));
            dt.Columns.Add("單號", typeof(string));
                 dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("流程階段", typeof(string));
            dt.Columns.Add("MAILHEAD", typeof(string));
            dt.Columns.Add("MAILTEMP", typeof(string));
            dt.Columns.Add("MAILTO", typeof(string));
            dt.Columns.Add("支付通知單備註", typeof(string));
            
            return dt;
        }
        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
           
        }


        private void MailTest(string strSubject, string MailAddress, string template)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));
            //if (fmLogin.LoginID.ToString().ToUpper() == "SERENAWU")
            //{
            //    message.CC.Add("SERENAWU@acmepoint.com");
            //}

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

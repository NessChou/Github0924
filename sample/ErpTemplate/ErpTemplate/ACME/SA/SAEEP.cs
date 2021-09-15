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
    public partial class SAEEP : Form
    {
        Attachment data = null;

        public SAEEP()
        {
            InitializeComponent();
        }
        private System.Data.DataTable GET1(string ID)
        {
            SqlConnection connection = globals.EEPConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    			   SELECT  DISTINCT DOCENTRY from acme_itt1 WHERE  ID=@ID");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
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
        private System.Data.DataTable GetOrderDataAPL(string APPLICANT)
        {
            SqlConnection connection = globals.EEPConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT FLOW_DESC 流程,LTRIM(LTRIM(REPLACE(REPLACE(REPLACE(FORM_PRESENTATION,'=',''),'DOCENTRY',''),'ID',''))) 單號, ");
            sb.Append("              D_STEP_ID 流程階段,MAILHEAD,MAILTEMP,MAILTO ,T2.MEMO 支付通知單備註 ");
            sb.Append("              FROM SYS_TODOLIST T0  ");
            sb.Append("                            INNER JOIN ACME_MAIL_BACKUP2 T1 ON (T0.LISTID=T1.MAILDOC AND T0.D_STEP_ID=T1.FLOWTYPE)  ");
            sb.Append("							                            LEFT JOIN ACME_OITT T2 ON ( REPLACE(LTRIM(LTRIM(REPLACE(REPLACE(REPLACE(FORM_PRESENTATION,'=',''),'DOCENTRY',''),'ID',''))),'''','')=T2.ID)  ");
            sb.Append("                             WHERE STATUS='N'  AND SUBSTRING(UPDATE_DATE,1,6)>'202001'   AND FLOW_DESC <> '銷貨單流程(TFT)'   ");

            if (globals.GroupID.ToString().Trim() != "EEP")
            {
                if (fmLogin.LoginID.ToString().ToUpper() == "KIKILEE")
                {
                    sb.Append("      AND APPLICANT IN ('KIKILEE','LILYLEE') ");
                }
                else
                {


                    sb.Append("      AND APPLICANT=@APPLICANT ");

                }
            }
            sb.Append("               ORDER BY FLOW_DESC ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@APPLICANT", APPLICANT));
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

            System.Data.DataTable dtemp5 = GetOrderDataAPL(fmLogin.LoginID.ToString());
                            DataRow dr= null;
                            for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
                            {

                                dr = TempDt.NewRow();
                                string 流程 = dtemp5.Rows[i]["流程"].ToString();
                                dr["流程"] = 流程;
                                string ID = dtemp5.Rows[i]["單號"].ToString();
                                dr["單號"] = ID;
                                string ff = ID.Substring(1, 1);
                                //if (ff != "P")
                                //{
                                    System.Data.DataTable G1 = GET1(ID);
                                    StringBuilder sb = new StringBuilder();

                                    if (流程 == "支付通知單(服務類)")
                                    {
                                        if (G1.Rows.Count > 0)
                                        {
                                            for (int i2 = 0; i2 <= G1.Rows.Count - 1; i2++)
                                            {


                                                sb.Append("" + G1.Rows[i2][0].ToString() + "/");


                                            }
                                            sb.Remove(sb.Length - 1, 1);

                                            dr["採購單號"] = sb.ToString();
                                        }
                                    }
                                    dr["流程階段"] = dtemp5.Rows[i]["流程階段"].ToString();
                                    dr["MAILHEAD"] = dtemp5.Rows[i]["MAILHEAD"].ToString();
                                    dr["MAILTEMP"] = dtemp5.Rows[i]["MAILTEMP"].ToString();
                                    dr["MAILTO"] = dtemp5.Rows[i]["MAILTO"].ToString();
                                    dr["支付通知單備註"] = dtemp5.Rows[i]["支付通知單備註"].ToString();
                                    TempDt.Rows.Add(dr);
                             //   }
                            }
                            dataGridView1.DataSource = TempDt;
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
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result;
                result = MessageBox.Show("請確認是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[0];
                    //                    string MAILHEAD = "[提醒簽核]" + row.Cells["MAILHEAD"].Value.ToString();
                    string MAILHEAD = row.Cells["MAILHEAD"].Value.ToString();
                    string MAILTO = row.Cells["MAILTO"].Value.ToString();
                    string MAILTEMP = row.Cells["MAILTEMP"].Value.ToString();
                
                    string USER = fmLogin.LoginID.ToString();
                  //  F1 = "1";
                    if (checkBox1.Checked)
                    {
                        MAILTO = USER + "@acmepoint.com";
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

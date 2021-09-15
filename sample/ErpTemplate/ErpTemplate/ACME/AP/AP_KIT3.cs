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
    public partial class AP_KIT3 : Form
    {
        string FA = "acmesql98";
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public string q2;
        string STATUS = "";
        public string q;
        public AP_KIT3()
        {
            InitializeComponent();
        }





        private void AP_KIT_Load(object sender, EventArgs e)
        {
            toolTip1.AutomaticDelay = 100;
            toolTip1.AutoPopDelay = 1000;
            toolTip1.ReshowDelay = 100;
            aP_KIT9DataGridView.ShowCellToolTips = false;

            this.aP_KIT9TableAdapter.Fill(this.lC.AP_KIT9);
            if (globals.DBNAME == "進金生")
            {
                strCn98 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                FA = "acmesql02";
            }
            aP_KIT3TableAdapter.Connection = globals.Connection;
            aP_KIT4TableAdapter.Connection = globals.Connection;
            aP_KIT5TableAdapter.Connection = globals.Connection;
            aP_KIT6TableAdapter.Connection = globals.Connection;
            aP_KIT7TableAdapter.Connection = globals.Connection;
            aP_KIT9TableAdapter.Connection = globals.Connection;
            this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text.Trim(), textITEMNAME.Text.Trim(), STATUS);
            this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text.Trim(), textITEMNAME.Text.Trim(), STATUS);
            this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text.Trim(), textITEMNAME.Text.Trim(), STATUS);
            this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text.Trim(), textITEMNAME.Text.Trim(), STATUS);
            this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text.Trim(), textITEMNAME.Text.Trim(), STATUS);
            this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text.Trim(), STATUS);
            comboBox2.Text = "未結";

            comboBox1.Text = "廠商詢價";

            contextTV.Items["TV"].Visible = false;
            contextTV.Items["PID"].Visible = true;
            contextTV.Items["GD"].Visible = true;
            contextTV.Items["DT"].Visible = true;
            contextTV.Items["NB"].Visible = true;

            //tabControl1.SelectedIndex = 1;
            //tabControl1.SelectedIndex = 0;
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT4BindingSource.EndEdit();
            this.aP_KIT4TableAdapter.Update(this.lC.AP_KIT4);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex != 5)
            {
                if (tabControl1.SelectedIndex == 0)
                {
                    if (aP_KIT3DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    if (aP_KIT4DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 2)
                {
                    if (aP_KIT5DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 3)
                {
                    if (aP_KIT6DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 4)
                {
                    if (aP_KIT7DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }

                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }


                StringBuilder sb = new StringBuilder();



                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);


                q = sb.ToString();

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\OPCH\\KIT3.xls";


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                System.Data.DataTable G1 = null;


                if (tabControl1.SelectedIndex == 0)
                {
                    G1 = GetEXCEL(q, "AP_KIT3", "TV");
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    G1 = GetEXCEL(q, "AP_KIT4", "PID");

                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    G1 = GetEXCEL(q, "AP_KIT5", "GD");

                }
                else if (tabControl1.SelectedIndex == 3)
                {
                    G1 = GetEXCEL(q, "AP_KIT6", "DT");

                }
                else if (tabControl1.SelectedIndex == 4)
                {
                    G1 = GetEXCEL(q, "AP_KIT7", "NB");

                }
      
                    ExcelReport.ExcelReportOutput(G1, ExcelTemplate, OutPutFile, "N");
                

            }
            else
            {
                ExcelReport.GridViewToExcel(aP_KIT9DataGridView);
            }

        }
        private void Updatepath(string filename, string path, string ID, string ID2)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (ID2 == "A")
            {
                sb.Append(" update AP_KIT set filename=@filename,[path]=@path where ID=@ID");
            }
            if (ID2 == "B")
            {
                sb.Append(" update AP_KIT2 set filename=@filename,[path]=@path where ID=@ID");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@ID", ID));
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
        private void UPKIT9(string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("UPDATE AP_KIT9  SET  EBU=CASE ISNULL(EBU,0) WHEN 0 THEN 1 WHEN 1 THEN 2 WHEN 2 THEN 0  END  WHERE ID=@ID");
            

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
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

        private void aP_KITDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void aP_KIT2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void aP_KIT2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KITDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  查無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }


            strB.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border-collapse:collapse;'>");
            strB.AppendLine("<tr class='HeaderBorder'>");
            //cteate table header
            for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
            {
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                //if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                //{



                //    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                //    tmpKeyValue = KeyValue;
                //}
                //else
                //{
                //    tmpKeyValue = "";
                //}
                KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                tmpKeyValue = KeyValue;

                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }



                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                }


                for (int d = 1; d <= dg.Rows[i].Cells.Count - 1; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    // HttpUtility.HtmlDecode("&nbsp;&nbsp;&nbsp;")

                    if (dgvc.ValueType == typeof(Int32))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Int32 x = Convert.ToInt32(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }

                    else if (dgvc.ValueType == typeof(Decimal) || dgvc.ValueType == typeof(Double))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Decimal x = Convert.ToDecimal(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0.00") + "</td>");
                        }


                    }
                    else
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {

                            if (dg.Columns[dgvc.ColumnIndex].HeaderText.IndexOf("日期") >= 0)
                            {
                                if (dgvc.Value.ToString() == "0")
                                {
                                    strB.AppendLine("<td>&nbsp;</td>");
                                }
                                else
                                {

                                    string sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                 dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                 dgvc.Value.ToString().Substring(6, 2);


                                    strB.AppendLine("<td>" + sDate + "</td>");
                                }
                            }
                            else
                            {
                                strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                            }
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string TABLE = "";
            string DOCTYPE = "";
            try
            {

                if (tabControl1.SelectedIndex == 0)
                {
                    TABLE = "AP_KIT3";
                    DOCTYPE = "TV";
                    if (aP_KIT3DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    TABLE = "AP_KIT4";
                    DOCTYPE = "PID";
                    if (aP_KIT4DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 2)
                {
                    TABLE = "AP_KIT5";
                    DOCTYPE = "GD";
                    if (aP_KIT5DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 3)
                {
                    TABLE = "AP_KIT6";
                    DOCTYPE = "DT";
                    if (aP_KIT6DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 4)
                {
                    TABLE = "AP_KIT7";
                    DOCTYPE = "NB";
                    if (aP_KIT7DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }


                StringBuilder sb = new StringBuilder();



                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);


                q = sb.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            string MAIL = "";
            string SUBJECT = "";
            string SA = "";
            System.Data.DataTable G1 = null;
            //if (comboBox1.Text == "進貨通知")
            //{
            //    MAIL = "\\MailTemplates\\KIT1.htm";
            //    if (tabControl1.SelectedIndex == 1)
            //    {
            //        MAIL = "\\MailTemplates\\KIT4.htm";
            //    }
            //    else
            //    {
            //        MAIL = "\\MailTemplates\\KIT1.htm";
            //    }
            //    G1 = Getbb(q, TABLE);
            //    SUBJECT = G1.Rows[0]["廠商"].ToString() + "進貨內湖，請查收，謝謝!!";
            //    SA = G1.Rows[0]["SA"].ToString();
            //}
            //if (comboBox1.Text == "廠商詢價")
            //{
            //    G1 = Getbb2(q, TABLE);
            //    MAIL = "\\MailTemplates\\KIT2.htm";
            //    SUBJECT = "請幫忙提供報價單 (" + G1.Rows[0]["廠商"].ToString() + ")";
            //}
            if (comboBox1.Text == "廠商詢價")
            {

                //ACME PO#27682-瑞威
                G1 = GetEXCEL(q, TABLE, DOCTYPE);
                MAIL = "\\MailTemplates\\KIT3.htm";
                SUBJECT = "進金生詢價";
            }

            if (G1.Rows.Count > 0)
            {

                string CARDNAME = G1.Rows[0]["客戶名稱"].ToString();
                dataGridView1.DataSource = G1;
                string GG = htmlMessageBody(dataGridView1).ToString();
                string EMAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                MailTest2(SUBJECT, EMAIL, GG, CARDNAME, MAIL, SA);
                MessageBox.Show("寄信成功");
            }
        }
        private void MailTest2(string strSubject, string MailAddress, string MailContent, string CUST, string MAIL, string SA)
        {
            MailMessage message = new MailMessage();
            string FROM = fmLogin.LoginID.ToString() + "@acmepoint.com";
            message.From = new MailAddress(FROM, "系統發送");
            //   MailAddress = "";
            if (MailAddress == "")
            {
                MailAddress = "LLEYTONCHEN@ACMEPOINT.COM";
            }
            message.To.Add(new MailAddress(MailAddress));

            string template;
            StreamReader objReader;
            string GetExePath = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //"\\MailTemplates\\KIT1.htm"
            objReader = new StreamReader(GetExePath + MAIL);

            template = objReader.ReadToEnd();
            objReader.Close();

            template = template.Replace("##Content##", MailContent);
            template = template.Replace("##CUST##", CUST);
            template = template.Replace("##CUST2##", "進貨內湖，請查收，謝謝!!");
            template = template.Replace("##CUST3##", "進貨內湖，請安排檢測，");
            template = template.Replace("##CUST3F##", "進貨內湖，請安排檢測及備品，");
            template = template.Replace("##CUST4##", "檢測完成後，煩請轉交給David入庫，謝謝。");
            template = template.Replace("##CUST5##", "FQA報告如附檔,謝謝!");
            template = template.Replace("##CUST6##", "HI!   " + SA);
            template = template.Replace("##KIT2##", "附上規格書及圖面，請幫忙提供報價單，謝謝。");
            template = template.Replace("##KIT3##", "請查收附檔po並回覆交期，謝謝!~");
            string USER = fmLogin.LoginID.ToString();
            System.Data.DataTable T1 = GETOHEM(USER);
            if (T1.Rows.Count > 0)
            {
                template = template.Replace("##EMP##", T1.Rows[0]["EMP"].ToString());
                template = template.Replace("##TEL##", "Tel : 886-2-87912868 *" + T1.Rows[0]["分機"].ToString());

                string EMAIL = T1.Rows[0]["EMAIL"].ToString();
                //StringBuilder sb = new StringBuilder();
                //sb.Append(" <a href='mailto:	" + EMAIL + "' ");
                //sb.Append("                 title='blocked::mailto:" + EMAIL + "' '>" + EMAIL + "</a></span></font></p>");

                template = template.Replace("##EMAIL##", EMAIL);
            }
            message.Subject = strSubject;
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            //string OutPutFile = lsAppDir + "\\Excel\\temp";
            //string[] filenames = Directory.GetFiles(OutPutFile);
            //foreach (string file in filenames)
            //{

            //    string m_File = "";

            //    m_File = file;
            //    data = new Attachment(m_File, MediaTypeNames.Application.Octet);

            //    //附件资料
            //    ContentDisposition disposition = data.ContentDisposition;


            //    // 加入邮件附件
            //    message.Attachments.Add(data);

            //}

            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";

            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

            try
            {
                client.Send(message);
                //   data.Dispose();
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {

                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }

                }
            }
            catch (Exception ex)
            {

            }

        }


        public System.Data.DataTable Getbb(string cs, string TABLE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                   SELECT  PO 採購單號,PODATE 採購回覆, ");
            sb.Append(" ITEMCODE 進金生15碼料號,ITEMNAME 型號品名,ITEMCODES 小料號,GRADE 等級,QTY '問貨況(數量)' ,");
            sb.Append(" F1 '1月FCST',F2 '2月FCST',F3 '3月FCST',F4 '4月FCST',F5 '5月FCST',F6 '6月FCST',F7 '7月FCST',F8 '8月FCST',F9 '9月FCST',F10 '10月FCST',F11 '11月FCST',F12 '12月FCST',");
            sb.Append(" SO 銷售訂單,SALES,SA,CARDNAME 客戶名稱,PRICE 賣價,USD '完稅 or 美金交易',MEMO SA備註");
            sb.Append("  FROM " + TABLE + " WHERE ID in ( " + cs + ") ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable GETOHEM(string EMP)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = globals.shipConnection;
            sb.Append(" SELECT MOBILE EMP,EMAIL,OFFICEEXT 分機 FROM OHEM WHERE HOMETEL=@EMP ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EMP", EMP));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }


        public System.Data.DataTable Getbb3(string cs, string TABLE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT CARDNAME 廠商,PO 採購單號,'' 交期,ITEMCODE 進金生料號,ITEMNAME 購買品名 ");
            sb.Append("               ,QTY 採購數量,Currency,AMT 採購金額 FROM " + TABLE + "  WHERE ID in ( " + cs + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable GetEXCEL(string cs, string TABLE, string DOCTYPE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("                   SELECT  '" + DOCTYPE + "'  類別 ,CASE STATUS WHEN 'True' then 'Close' else '' end Status,PO 採購單號,PODATE 採購回覆, ");
            sb.Append(" ITEMCODE 進金生15碼料號,ITEMNAME 型號品名,ITEMCODES 小料號,GRADE 等級,QTY 數量 ,");
            sb.Append(" F1 '1月FCST',F2 '2月FCST',F3 '3月FCST',F4 '4月FCST',F5 '5月FCST',F6 '6月FCST',F7 '7月FCST',F8 '8月FCST',F9 '9月FCST',F10 '10月FCST',F11 '11月FCST',F12 '12月FCST',");
            sb.Append(" SO 銷售訂單,SALES,SA,CARDNAME 客戶名稱,PRICE 賣價,USD 美金,MEMO SA備註,DESCRIPTION 接單數量,DOCDATE 日期 ");
            sb.Append(" FROM  " + TABLE + "  ");
            sb.Append("   WHERE ID in ( " + cs + ") ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
     
        public System.Data.DataTable GetPATH(string ID)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [PATH] FROM AP_KIT WHERE ID=@ID");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));

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


        private void button4_Click(object sender, EventArgs e)
        {
            textPO1.Text = "";

            comboBox2.Text = "全部";
            this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);
            this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
        }


        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "全部")
            {
                STATUS = "";
            }
            if (comboBox2.Text == "已結")
            {
                STATUS = "True";
            }

            if (comboBox2.Text == "未結")
            {
                STATUS = "False";
            }
        }



        private void aP_KITBindingNavigatorSaveItem_Click_5(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT3BindingSource.EndEdit();
            this.aP_KIT3TableAdapter.Update(this.lC.AP_KIT3);
        }

        private void aP_KIT3DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STATUS1"].Value = "False";
            e.Row.Cells["DOCDATE1"].Value = GetMenu.Day();

        }


        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT5BindingSource.EndEdit();
            this.aP_KIT5TableAdapter.Update(this.lC.AP_KIT5);
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT6BindingSource.EndEdit();
            this.aP_KIT6TableAdapter.Update(this.lC.AP_KIT6);
        }

        private void toolStripButton28_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT7BindingSource.EndEdit();
            this.aP_KIT7TableAdapter.Update(this.lC.AP_KIT7);
        }

        private void aP_KIT3DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KIT4DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KIT5DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KIT6DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KIT7DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
        private void aP_KIT9DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
        private void aP_KIT4DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STATUS2"].Value = "False";
            e.Row.Cells["DOCDATE2"].Value = GetMenu.Day();
        }

        private void aP_KIT5DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STATUS3"].Value = "False";
            e.Row.Cells["DOCDATE3"].Value = GetMenu.Day();
        }

        private void aP_KIT6DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STATUS4"].Value = "False";
            e.Row.Cells["DOCDATE4"].Value = GetMenu.Day();
        }

        private void aP_KIT7DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["STATUS5"].Value = "False";
            e.Row.Cells["DOCDATE5"].Value = GetMenu.Day();
        }

        private void aP_KIT3DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT3DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT3DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT3DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID1"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aP_KIT4DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT4DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT4DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT4DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID2"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aP_KIT5DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT5DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT5DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT5DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID3"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aP_KIT6DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT6DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT6DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT6DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID4"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aP_KIT7DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT7DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT7DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT7DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID5"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CalcTotals1();
        }

        private void CalcTotals1()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT3DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT3DataGridView.SelectedRows[iRecs];


                    //DESCRIPTION
                    string DESCRIPTION1 = row.Cells["DESCRIPTION1"].Value.ToString();
                    if (String.IsNullOrEmpty(DESCRIPTION1))
                    {
                        DESCRIPTION1 = "0";
                    }
                    string QTY1 = row.Cells["QTY1"].Value.ToString();
                    if (String.IsNullOrEmpty(QTY1))
                    {
                        QTY1 = "0";
                    }
                    string F1G1 = row.Cells["F1G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F1G1))
                    {
                        F1G1 = "0";
                    }
                    string F2G1 = row.Cells["F2G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F2G1))
                    {
                        F2G1 = "0";
                    }
                    string F3G1 = row.Cells["F3G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F3G1))
                    {
                        F3G1 = "0";
                    }
                    string F4G1 = row.Cells["F4G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F4G1))
                    {
                        F4G1 = "0";
                    }
                    string F5G1 = row.Cells["F5G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F5G1))
                    {
                        F5G1 = "0";
                    }
                    string F6G1 = row.Cells["F6G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F6G1))
                    {
                        F6G1 = "0";
                    }
                    string F7G1 = row.Cells["F7G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F7G1))
                    {
                        F7G1 = "0";
                    }
                    string F8G1 = row.Cells["F8G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F8G1))
                    {
                        F8G1 = "0";
                    }
                    string F9G1 = row.Cells["F9G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F9G1))
                    {
                        F9G1 = "0";
                    }
                    string F10G1 = row.Cells["F10G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F10G1))
                    {
                        F10G1 = "0";
                    }
                    string F11G1 = row.Cells["F11G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F11G1))
                    {
                        F11G1 = "0";
                    }
                    string F12G1 = row.Cells["F12G1"].Value.ToString();
                    if (String.IsNullOrEmpty(F12G1))
                    {
                        F12G1 = "0";
                    }
                    iTotal += Convert.ToInt32(DESCRIPTION1)
                        + Convert.ToInt32(QTY1)
                        + Convert.ToInt32(F1G1)
                        + Convert.ToInt32(F2G1)
                        + Convert.ToInt32(F3G1)
                        + Convert.ToInt32(F4G1)
                        + Convert.ToInt32(F5G1)
                        + Convert.ToInt32(F6G1)
                        + Convert.ToInt32(F7G1)
                        + Convert.ToInt32(F8G1)
                        + Convert.ToInt32(F9G1)
                        + Convert.ToInt32(F10G1)
                        + Convert.ToInt32(F11G1)
                        + Convert.ToInt32(F12G1);

                }

                textBox1.Text = iTotal.ToString("0");


            }
            catch { }


        }
        private void CalcTotals2()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT4DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT4DataGridView.SelectedRows[iRecs];
                    string DESCRIPTION2 = row.Cells["DESCRIPTION2"].Value.ToString();
                    if (String.IsNullOrEmpty(DESCRIPTION2))
                    {
                        DESCRIPTION2 = "0";
                    }
                    string QTY2 = row.Cells["QTY2"].Value.ToString();
                    if (String.IsNullOrEmpty(QTY2))
                    {
                        QTY2 = "0";
                    }
                    string F1G2 = row.Cells["F1G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F1G2))
                    {
                        F1G2 = "0";
                    }
                    string F2G2 = row.Cells["F2G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F2G2))
                    {
                        F2G2 = "0";
                    }
                    string F3G2 = row.Cells["F3G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F3G2))
                    {
                        F3G2 = "0";
                    }
                    string F4G2 = row.Cells["F4G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F4G2))
                    {
                        F4G2 = "0";
                    }
                    string F5G2 = row.Cells["F5G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F5G2))
                    {
                        F5G2 = "0";
                    }
                    string F6G2 = row.Cells["F6G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F6G2))
                    {
                        F6G2 = "0";
                    }
                    string F7G2 = row.Cells["F7G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F7G2))
                    {
                        F7G2 = "0";
                    }
                    string F8G2 = row.Cells["F8G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F8G2))
                    {
                        F8G2 = "0";
                    }
                    string F9G2 = row.Cells["F9G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F9G2))
                    {
                        F9G2 = "0";
                    }
                    string F10G2 = row.Cells["F10G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F10G2))
                    {
                        F10G2 = "0";
                    }
                    string F11G2 = row.Cells["F11G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F11G2))
                    {
                        F11G2 = "0";
                    }
                    string F12G2 = row.Cells["F12G2"].Value.ToString();
                    if (String.IsNullOrEmpty(F12G2))
                    {
                        F12G2 = "0";
                    }
                    iTotal += Convert.ToInt32(DESCRIPTION2)
                        + Convert.ToInt32(QTY2)
                        + Convert.ToInt32(F1G2)
                        + Convert.ToInt32(F2G2)
                        + Convert.ToInt32(F3G2)
                        + Convert.ToInt32(F4G2)
                        + Convert.ToInt32(F5G2)
                        + Convert.ToInt32(F6G2)
                        + Convert.ToInt32(F7G2)
                        + Convert.ToInt32(F8G2)
                        + Convert.ToInt32(F9G2)
                        + Convert.ToInt32(F10G2)
                        + Convert.ToInt32(F11G2)
                        + Convert.ToInt32(F12G2);

                }

                textBox2.Text = iTotal.ToString("0");


            }
            catch { }


        }
        private void CalcTotals3()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT5DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT5DataGridView.SelectedRows[iRecs];

                    string DESCRIPTION3 = row.Cells["DESCRIPTION3"].Value.ToString();
                    if (String.IsNullOrEmpty(DESCRIPTION3))
                    {
                        DESCRIPTION3 = "0";
                    }

                    string QTY3 = row.Cells["QTY3"].Value.ToString();
                    if (String.IsNullOrEmpty(QTY3))
                    {
                        QTY3 = "0";
                    }
                    string F1G3 = row.Cells["F1G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F1G3))
                    {
                        F1G3 = "0";
                    }
                    string F2G3 = row.Cells["F2G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F2G3))
                    {
                        F2G3 = "0";
                    }
                    string F3G3 = row.Cells["F3G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F3G3))
                    {
                        F3G3 = "0";
                    }
                    string F4G3 = row.Cells["F4G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F4G3))
                    {
                        F4G3 = "0";
                    }
                    string F5G3 = row.Cells["F5G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F5G3))
                    {
                        F5G3 = "0";
                    }
                    string F6G3 = row.Cells["F6G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F6G3))
                    {
                        F6G3 = "0";
                    }
                    string F7G3 = row.Cells["F7G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F7G3))
                    {
                        F7G3 = "0";
                    }
                    string F8G3 = row.Cells["F8G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F8G3))
                    {
                        F8G3 = "0";
                    }
                    string F9G3 = row.Cells["F9G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F9G3))
                    {
                        F9G3 = "0";
                    }
                    string F10G3 = row.Cells["F10G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F10G3))
                    {
                        F10G3 = "0";
                    }
                    string F11G3 = row.Cells["F11G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F11G3))
                    {
                        F11G3 = "0";
                    }
                    string F12G3 = row.Cells["F12G3"].Value.ToString();
                    if (String.IsNullOrEmpty(F12G3))
                    {
                        F12G3 = "0";
                    }
                    iTotal += Convert.ToInt32(DESCRIPTION3)
                        + Convert.ToInt32(QTY1)
                        + Convert.ToInt32(F1G3)
                        + Convert.ToInt32(F2G3)
                        + Convert.ToInt32(F3G3)
                        + Convert.ToInt32(F4G3)
                        + Convert.ToInt32(F5G3)
                        + Convert.ToInt32(F6G3)
                        + Convert.ToInt32(F7G3)
                        + Convert.ToInt32(F8G3)
                        + Convert.ToInt32(F9G3)
                        + Convert.ToInt32(F10G3)
                        + Convert.ToInt32(F11G3)
                        + Convert.ToInt32(F12G3);

                }

                textBox3.Text = iTotal.ToString("0");


            }
            catch { }


        }
        private void CalcTotals4()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT6DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT6DataGridView.SelectedRows[iRecs];

                    string DESCRIPTION4 = row.Cells["DESCRIPTION4"].Value.ToString();
                    if (String.IsNullOrEmpty(DESCRIPTION4))
                    {
                        DESCRIPTION4 = "0";
                    }

                    string QTY4 = row.Cells["QTY4"].Value.ToString();
                    if (String.IsNullOrEmpty(QTY4))
                    {
                        QTY4 = "0";
                    }
                    string F1G4 = row.Cells["F1G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F1G4))
                    {
                        F1G4 = "0";
                    }
                    string F2G4 = row.Cells["F2G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F2G4))
                    {
                        F2G4 = "0";
                    }
                    string F3G4 = row.Cells["F3G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F3G4))
                    {
                        F3G4 = "0";
                    }
                    string F4G4 = row.Cells["F4G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F4G4))
                    {
                        F4G4 = "0";
                    }
                    string F5G4 = row.Cells["F5G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F5G4))
                    {
                        F5G4 = "0";
                    }
                    string F6G4 = row.Cells["F6G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F6G4))
                    {
                        F6G4 = "0";
                    }
                    string F7G4 = row.Cells["F7G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F7G4))
                    {
                        F7G4 = "0";
                    }
                    string F8G4 = row.Cells["F8G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F8G4))
                    {
                        F8G4 = "0";
                    }
                    string F9G4 = row.Cells["F9G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F9G4))
                    {
                        F9G4 = "0";
                    }
                    string F10G4 = row.Cells["F10G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F10G4))
                    {
                        F10G4 = "0";
                    }
                    string F11G4 = row.Cells["F11G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F11G4))
                    {
                        F11G4 = "0";
                    }
                    string F12G4 = row.Cells["F12G4"].Value.ToString();
                    if (String.IsNullOrEmpty(F12G4))
                    {
                        F12G4 = "0";
                    }
                    iTotal += Convert.ToInt32(DESCRIPTION4)
                        + Convert.ToInt32(QTY1)
                        + Convert.ToInt32(F1G4)
                        + Convert.ToInt32(F2G4)
                        + Convert.ToInt32(F3G4)
                        + Convert.ToInt32(F4G4)
                        + Convert.ToInt32(F5G4)
                        + Convert.ToInt32(F6G4)
                        + Convert.ToInt32(F7G4)
                        + Convert.ToInt32(F8G4)
                        + Convert.ToInt32(F9G4)
                        + Convert.ToInt32(F10G4)
                        + Convert.ToInt32(F11G4)
                        + Convert.ToInt32(F12G4);

                }

                textBox4.Text = iTotal.ToString("0");


            }
            catch { }


        }
        private void CalcTotals5()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT7DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT7DataGridView.SelectedRows[iRecs];

                    string DESCRIPTION5 = row.Cells["DESCRIPTION5"].Value.ToString();
                    if (String.IsNullOrEmpty(DESCRIPTION5))
                    {
                        DESCRIPTION5 = "0";
                    }


                    string QTY5 = row.Cells["QTY5"].Value.ToString();
                    if (String.IsNullOrEmpty(QTY5))
                    {
                        QTY5 = "0";
                    }
                    string F1G5 = row.Cells["F1G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F1G5))
                    {
                        F1G5 = "0";
                    }
                    string F2G5 = row.Cells["F2G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F2G5))
                    {
                        F2G5 = "0";
                    }
                    string F3G5 = row.Cells["F3G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F3G5))
                    {
                        F3G5 = "0";
                    }
                    string F4G5 = row.Cells["F4G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F4G5))
                    {
                        F4G5 = "0";
                    }
                    string F5G5 = row.Cells["F5G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F5G5))
                    {
                        F5G5 = "0";
                    }
                    string F6G5 = row.Cells["F6G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F6G5))
                    {
                        F6G5 = "0";
                    }
                    string F7G5 = row.Cells["F7G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F7G5))
                    {
                        F7G5 = "0";
                    }
                    string F8G5 = row.Cells["F8G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F8G5))
                    {
                        F8G5 = "0";
                    }
                    string F9G5 = row.Cells["F9G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F9G5))
                    {
                        F9G5 = "0";
                    }
                    string F10G5 = row.Cells["F10G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F10G5))
                    {
                        F10G5 = "0";
                    }
                    string F11G5 = row.Cells["F11G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F11G5))
                    {
                        F11G5 = "0";
                    }
                    string F12G5 = row.Cells["F12G5"].Value.ToString();
                    if (String.IsNullOrEmpty(F12G5))
                    {
                        F12G5 = "0";
                    }
                    iTotal += Convert.ToInt32(DESCRIPTION5)
                        + Convert.ToInt32(QTY5)
                        + Convert.ToInt32(F1G5)
                        + Convert.ToInt32(F2G5)
                        + Convert.ToInt32(F3G5)
                        + Convert.ToInt32(F4G5)
                        + Convert.ToInt32(F5G5)
                        + Convert.ToInt32(F6G5)
                        + Convert.ToInt32(F7G5)
                        + Convert.ToInt32(F8G5)
                        + Convert.ToInt32(F9G5)
                        + Convert.ToInt32(F10G5)
                        + Convert.ToInt32(F11G5)
                        + Convert.ToInt32(F12G5);

                }

                textBox5.Text = iTotal.ToString("0");


            }
            catch { }


        }
        private void button5_Click(object sender, EventArgs e)
        {
            CalcTotals2();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CalcTotals5();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            CalcTotals4();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            CalcTotals3();
        }

        private void aP_KIT3DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (aP_KIT3DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE1")
            {
                string ITEM = aP_KIT3DataGridView.Rows[e.RowIndex].Cells["ITEMCODE1"].Value.ToString();

                System.Data.DataTable T1 = GetITEM(ITEM);
                if (T1.Rows.Count > 0)
                {

                    this.aP_KIT3DataGridView.Rows[e.RowIndex].Cells["ITEMNAME1"].Value = T1.Rows[0][0].ToString();
                }

            }
        }
        public System.Data.DataTable GetITEM(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMNAME FROM OITM WHERE SUBSTRING(ITEMCODE,1,4) <> 'ACME' AND ITEMCODE=@ITEMCODE ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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

        private void aP_KIT4DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (aP_KIT4DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE2")
            {
                string ITEM = aP_KIT4DataGridView.Rows[e.RowIndex].Cells["ITEMCODE2"].Value.ToString();

                System.Data.DataTable T1 = GetITEM(ITEM);
                if (T1.Rows.Count > 0)
                {

                    this.aP_KIT4DataGridView.Rows[e.RowIndex].Cells["ITEMNAME2"].Value = T1.Rows[0][0].ToString();
                }

            }
        }

        private void aP_KIT5DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (aP_KIT5DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE3")
            {
                string ITEM = aP_KIT5DataGridView.Rows[e.RowIndex].Cells["ITEMCODE3"].Value.ToString();

                System.Data.DataTable T1 = GetITEM(ITEM);
                if (T1.Rows.Count > 0)
                {

                    this.aP_KIT5DataGridView.Rows[e.RowIndex].Cells["ITEMNAME3"].Value = T1.Rows[0][0].ToString();
                }

            }
        }

        private void aP_KIT6DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (aP_KIT6DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE4")
            {
                string ITEM = aP_KIT6DataGridView.Rows[e.RowIndex].Cells["ITEMCODE4"].Value.ToString();

                System.Data.DataTable T1 = GetITEM(ITEM);
                if (T1.Rows.Count > 0)
                {

                    this.aP_KIT6DataGridView.Rows[e.RowIndex].Cells["ITEMNAME4"].Value = T1.Rows[0][0].ToString();
                }

            }
        }

        private void aP_KIT7DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (aP_KIT7DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE5")
            {
                string ITEM = aP_KIT7DataGridView.Rows[e.RowIndex].Cells["ITEMCODE5"].Value.ToString();

                System.Data.DataTable T1 = GetITEM(ITEM);
                if (T1.Rows.Count > 0)
                {

                    this.aP_KIT7DataGridView.Rows[e.RowIndex].Cells["ITEMNAME5"].Value = T1.Rows[0][0].ToString();
                }

            }
        }
        private void TV_Click(object sender, EventArgs e)
        {

            if (tabControl1.SelectedIndex == 1)
            {
                ADD1(lC.AP_KIT4, lC.AP_KIT3, aP_KIT4DataGridView, aP_KIT3DataGridView, "AP_KIT4");
                this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ADD1(lC.AP_KIT5, lC.AP_KIT3, aP_KIT5DataGridView, aP_KIT3DataGridView, "AP_KIT5");
                this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                ADD1(lC.AP_KIT6, lC.AP_KIT3, aP_KIT6DataGridView, aP_KIT3DataGridView, "AP_KIT6");
                this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 4)
            {
                ADD1(lC.AP_KIT7, lC.AP_KIT3, aP_KIT7DataGridView, aP_KIT3DataGridView, "AP_KIT7");
                this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);

            }
            this.aP_KIT3BindingSource.EndEdit();
            this.aP_KIT3TableAdapter.Update(this.lC.AP_KIT3);
        }
        private void PID_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ADD1(lC.AP_KIT3, lC.AP_KIT4, aP_KIT3DataGridView, aP_KIT4DataGridView, "AP_KIT3");
                this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ADD1(lC.AP_KIT5, lC.AP_KIT4, aP_KIT5DataGridView, aP_KIT4DataGridView, "AP_KIT5");
                this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                ADD1(lC.AP_KIT6, lC.AP_KIT4, aP_KIT6DataGridView, aP_KIT4DataGridView, "AP_KIT6");
                this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 4)
            {
                ADD1(lC.AP_KIT7, lC.AP_KIT4, aP_KIT7DataGridView, aP_KIT4DataGridView, "AP_KIT7");
                this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            this.aP_KIT4BindingSource.EndEdit();
            this.aP_KIT4TableAdapter.Update(this.lC.AP_KIT4);
        }

        private void GD_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ADD1(lC.AP_KIT3, lC.AP_KIT5, aP_KIT3DataGridView, aP_KIT5DataGridView, "AP_KIT3");
                this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ADD1(lC.AP_KIT4, lC.AP_KIT5, aP_KIT4DataGridView, aP_KIT5DataGridView, "AP_KIT4");
                this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                ADD1(lC.AP_KIT6, lC.AP_KIT5, aP_KIT6DataGridView, aP_KIT5DataGridView, "AP_KIT6");
                this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 4)
            {
                ADD1(lC.AP_KIT7, lC.AP_KIT5, aP_KIT7DataGridView, aP_KIT5DataGridView, "AP_KIT7");
                this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);
            }

            this.aP_KIT5BindingSource.EndEdit();
            this.aP_KIT5TableAdapter.Update(this.lC.AP_KIT5);
        }
        private void DT_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ADD1(lC.AP_KIT3, lC.AP_KIT6, aP_KIT3DataGridView, aP_KIT6DataGridView, "AP_KIT3");
                this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ADD1(lC.AP_KIT4, lC.AP_KIT6, aP_KIT4DataGridView, aP_KIT6DataGridView, "AP_KIT4");
                this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ADD1(lC.AP_KIT5, lC.AP_KIT6, aP_KIT5DataGridView, aP_KIT6DataGridView, "AP_KIT5");
                this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 4)
            {
                ADD1(lC.AP_KIT7, lC.AP_KIT6, aP_KIT7DataGridView, aP_KIT6DataGridView, "AP_KIT7");
                this.aP_KIT7TableAdapter.FillBy(this.lC.AP_KIT7, textPO1.Text, textITEMNAME.Text, STATUS);
            }

            this.aP_KIT6BindingSource.EndEdit();
            this.aP_KIT6TableAdapter.Update(this.lC.AP_KIT6);
        }

        private void NB_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ADD1(lC.AP_KIT3, lC.AP_KIT7, aP_KIT3DataGridView, aP_KIT6DataGridView, "AP_KIT3");
                this.aP_KIT3TableAdapter.FillBy(this.lC.AP_KIT3, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ADD1(lC.AP_KIT4, lC.AP_KIT7, aP_KIT4DataGridView, aP_KIT6DataGridView, "AP_KIT4");
                this.aP_KIT4TableAdapter.FillBy(this.lC.AP_KIT4, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ADD1(lC.AP_KIT5, lC.AP_KIT7, aP_KIT5DataGridView, aP_KIT6DataGridView, "AP_KIT5");
                this.aP_KIT5TableAdapter.FillBy(this.lC.AP_KIT5, textPO1.Text, textITEMNAME.Text, STATUS);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                ADD1(lC.AP_KIT6, lC.AP_KIT7, aP_KIT6DataGridView, aP_KIT6DataGridView, "AP_KIT6");
                this.aP_KIT6TableAdapter.FillBy(this.lC.AP_KIT6, textPO1.Text, textITEMNAME.Text, STATUS);
            }

            this.aP_KIT7BindingSource.EndEdit();
            this.aP_KIT7TableAdapter.Update(this.lC.AP_KIT7);
        }

        private void ADD1(System.Data.DataTable dt1, System.Data.DataTable dt2, DataGridView FROM, DataGridView T0, string FORMTABLE)
        {
            try
            {
                DataRow newCustomersRow = dt2.NewRow();
                int i = FROM.CurrentRow.Index;
                DataRow drw = dt1.Rows[i];
                int ID = Convert.ToInt32(drw["ID"]);

                newCustomersRow["STATUS"] = drw["STATUS"];
                newCustomersRow["PO"] = drw["PO"];
                newCustomersRow["PODATE"] = drw["PODATE"];
                newCustomersRow["ITEMCODE"] = drw["ITEMCODE"];
                newCustomersRow["ITEMNAME"] = drw["ITEMNAME"];
                newCustomersRow["ITEMCODES"] = drw["ITEMCODES"];
                newCustomersRow["GRADE"] = drw["GRADE"];
                newCustomersRow["QTY"] = drw["QTY"];
                newCustomersRow["F1"] = drw["F1"];
                newCustomersRow["F2"] = drw["F2"];
                newCustomersRow["F3"] = drw["F3"];
                newCustomersRow["F4"] = drw["F4"];
                newCustomersRow["F5"] = drw["F5"];
                newCustomersRow["F6"] = drw["F6"];
                newCustomersRow["F7"] = drw["F7"];
                newCustomersRow["F8"] = drw["F8"];
                newCustomersRow["F9"] = drw["F9"];
                newCustomersRow["F10"] = drw["F10"];
                newCustomersRow["F11"] = drw["F11"];
                newCustomersRow["F12"] = drw["F12"];
                newCustomersRow["SO"] = drw["SO"];
                newCustomersRow["SALES"] = drw["SALES"];
                newCustomersRow["SA"] = drw["SA"];
                newCustomersRow["CARDNAME"] = drw["CARDNAME"];
                newCustomersRow["PRICE"] = drw["PRICE"];
                newCustomersRow["USD"] = drw["USD"];
                newCustomersRow["MEMO"] = drw["MEMO"];
                newCustomersRow["DESCRIPTION"] = drw["DESCRIPTION"];
                newCustomersRow["DOCDATE"] = GetMenu.Day();

                dt2.Rows.InsertAt(newCustomersRow, T0.Rows.Count);
                Delete(FORMTABLE, ID);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }
        private void ADD2(System.Data.DataTable dt1, System.Data.DataTable dt2, DataGridView FROM, DataGridView T0)
        {
            try
            {
                DataRow newCustomersRow = dt2.NewRow();
                int i = FROM.CurrentRow.Index;
                DataRow drw = dt1.Rows[i];
                newCustomersRow["MODEL"] = drw["ITEMNAME"];
                newCustomersRow["GRADE"] = drw["GRADE"];
     

                dt2.Rows.InsertAt(newCustomersRow, T0.Rows.Count);

                this.aP_KIT9BindingSource.EndEdit();
                this.aP_KIT9TableAdapter.Update(this.lC.AP_KIT9);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }


        private void Delete(string TABLE, int ID)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" delete " + TABLE + "  WHERE ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ID", ID));


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

        private void tabControl1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                contextTV.Items["TV"].Visible = false;
                contextTV.Items["PID"].Visible = true;
                contextTV.Items["GD"].Visible = true;
                contextTV.Items["DT"].Visible = true;
                contextTV.Items["NB"].Visible = true;
            }

            if (tabControl1.SelectedIndex == 1)
            {
                contextTV.Items["TV"].Visible = true;
                contextTV.Items["PID"].Visible = false;
                contextTV.Items["GD"].Visible = true;
                contextTV.Items["DT"].Visible = true;
                contextTV.Items["NB"].Visible = true;

            }
            if (tabControl1.SelectedIndex == 2)
            {
                contextTV.Items["TV"].Visible = true;
                contextTV.Items["PID"].Visible = true;
                contextTV.Items["GD"].Visible = false;
                contextTV.Items["DT"].Visible = true;
                contextTV.Items["NB"].Visible = true;
            }
            if (tabControl1.SelectedIndex == 3)
            {
                contextTV.Items["TV"].Visible = true;
                contextTV.Items["PID"].Visible = true;
                contextTV.Items["GD"].Visible = true;
                contextTV.Items["DT"].Visible = false;
                contextTV.Items["NB"].Visible = true;

            }
            if (tabControl1.SelectedIndex == 4)
            {
                contextTV.Items["TV"].Visible = true;
                contextTV.Items["PID"].Visible = true;
                contextTV.Items["GD"].Visible = true;
                contextTV.Items["DT"].Visible = true;
                contextTV.Items["NB"].Visible = false;

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT3DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT3DataGridView.SelectedRows[i];

                row.Cells[2].Value = "True";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT3DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT3DataGridView.SelectedRows[i];

                row.Cells[2].Value = "False";
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT4DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT4DataGridView.SelectedRows[i];

                row.Cells[2].Value = "True";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT4DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT4DataGridView.SelectedRows[i];

                row.Cells[2].Value = "False";
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT5DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT5DataGridView.SelectedRows[i];

                row.Cells[2].Value = "True";
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT5DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT5DataGridView.SelectedRows[i];

                row.Cells[2].Value = "False";
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT6DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT6DataGridView.SelectedRows[i];

                row.Cells[2].Value = "True";
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT6DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT6DataGridView.SelectedRows[i];

                row.Cells[2].Value = "False";
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT7DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT7DataGridView.SelectedRows[i];

                row.Cells[2].Value = "True";
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT7DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT7DataGridView.SelectedRows[i];

                row.Cells[2].Value = "False";
            }
        }
        public System.Data.DataTable GetTOSAP(string cs, string DB)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME,ITEMCODE,ISNULL([DESCRIPTION],0)+ISNULL(QTY,0)+ISNULL(F1,0)+ISNULL(F2,0) +ISNULL(F3,0) +ISNULL(F4,0) +ISNULL(F5,0) +ISNULL(F6,0) +ISNULL(F7,0) +ISNULL(F8,0) +ISNULL(F9,0) +ISNULL(F10,0) +ISNULL(F11,0) +ISNULL(F12,0)   QTY,PRICE,USD CURRENCY,convert(varchar,CAST(MEMO AS DATETIME), 111)  CUSTDATE, ");
            sb.Append(" ISNULL(CARDNAME,'')+ISNULL(USD,'')+ISNULL(PRICE,'')+ISNULL(MEMO,'')  MEMO   FROM  " + DB + "	   WHERE ID in ( " + cs + ") ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable GETCARCODE(string CARDNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDCODE FROM OCRD WHERE CARDNAME LIKE  '%" + CARDNAME + "%' AND SUBSTRING(CardCode,1,1)='S' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable GetDI4()
        {
            SqlConnection connection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OPOR");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetOPOR7()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT EMPID FROM OHEM WHERE homeTel =@homeTel ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@homeTel", fmLogin.LoginID.ToString()));
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
        private void button20_Click(object sender, EventArgs e)
        {

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

            int i = 0; //  to be used as an index

            oCompany.CompanyDB = FA;
            oCompany.UserName = "A02";
            oCompany.Password = "6500";
            int result = oCompany.Connect();
            if (result == 0)
            {

                if (tabControl1.SelectedIndex == 0)
                {
                    if (aP_KIT3DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    if (aP_KIT4DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 2)
                {
                    if (aP_KIT5DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 3)
                {
                    if (aP_KIT6DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 4)
                {
                    if (aP_KIT7DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                ArrayList al = new ArrayList();

                for (int i2 = 0; i2 <= listBox1.Items.Count - 1; i2++)
                {
                    al.Add(listBox1.Items[i2].ToString());
                }


                StringBuilder sb = new StringBuilder();

                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);
                q2 = sb.ToString();
                System.Data.DataTable G1 = null;

                string CARDCODE = "";

                if (tabControl1.SelectedIndex == 0)
                {
                    G1 = GetTOSAP(q2, "AP_KIT3");
                    CARDCODE = "S0001-TV";
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    G1 = GetTOSAP(q2, "AP_KIT4");
                    CARDCODE = "S0623-PID";
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    G1 = GetTOSAP(q2, "AP_KIT5");
                    CARDCODE = "S0623-GD";
                }
                else if (tabControl1.SelectedIndex == 3)
                {
                    G1 = GetTOSAP(q2, "AP_KIT6");
                    CARDCODE = "S0001-DD";
                }
                else if (tabControl1.SelectedIndex == 4)
                {
                    G1 = GetTOSAP(q2, "AP_KIT7");
                    CARDCODE = "S0001-NB";
                }


                if (G1.Rows.Count > 0)
                {
                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                    string CURRENCY = G1.Rows[0]["CURRENCY"].ToString();

                    if (G1.Rows.Count > 0)
                    {

                        oPURCH.CardCode = CARDCODE;
                            oPURCH.VatPercent = 5;
                            oPURCH.DocCurrency = CURRENCY;

                            System.Data.DataTable G7 = GetOPOR7();
                            if (G7.Rows.Count > 0)
                            {
                                oPURCH.DocumentsOwner = Convert.ToInt32(G7.Rows[0][0]);
                            }
                            oPURCH.SalesPersonCode = 65;
                            for (int s = 0; s <= G1.Rows.Count - 1; s++)
                            {


                                string ITEMCODE = G1.Rows[s]["ITEMCODE"].ToString();
                                double QTY = Convert.ToDouble(G1.Rows[s]["QTY"]);
                                double PRICE = Convert.ToDouble(G1.Rows[s]["PRICE"]);
                                string MEMO = G1.Rows[s]["MEMO"].ToString();
                                string CUSTDATE = G1.Rows[s]["CUSTDATE"].ToString();

                                oPURCH.Lines.WarehouseCode = "TW001";
                                oPURCH.Lines.ItemCode = ITEMCODE;
                                oPURCH.Lines.Quantity = QTY;
                                oPURCH.Lines.Price = PRICE;
                                oPURCH.Lines.VatGroup = "AP5%";
                                oPURCH.Lines.CostingCode = "11111";
                                oPURCH.Lines.Currency = CURRENCY;
                                if (!String.IsNullOrEmpty(CUSTDATE))
                                {
                                    DateTime S1 = Convert.ToDateTime(CUSTDATE);
                                    oPURCH.Lines.ShipDate = S1;
                                }
                                oPURCH.Lines.UserFields.Fields.Item("U_MEMO").Value = MEMO;
                                oPURCH.Lines.Add();

                            }

                            int res = oPURCH.Add();
                            if (res != 0)
                            {
                                MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                            }
                            else
                            {
                                System.Data.DataTable G4 = GetDI4();
                                string OWTR = G4.Rows[0][0].ToString();
                                MessageBox.Show("上傳成功 採購單號 : " + OWTR);


                            }
                        
                    }

                }
            }



        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT9BindingSource.EndEdit();
            this.aP_KIT9TableAdapter.Update(this.lC.AP_KIT9);
        }

        private void button21_Click(object sender, EventArgs e)
        {
                      OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result2 = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                GD6(opdf.FileName);
                this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
            }
        }
        private void GD6(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            int T1 = 0;
            int T2 = 0;
            int T3 = 0;
            int T4 = 0;
            int T5 = 0;
            int T6 = 0;
            int T7 = 0;
            int T8 = 0;
            int T9 = 0;
            int T10 = 0;
            int T11 = 0;
            int T12 = 0;

            for (int b = 1; b <= 20; b++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                range.Select();
                string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");

                if (id == "1月FCST")
                {
                    T1 = b;
                }

                if (id == "2月FCST")
                {
                    T2 = b;
                }
                if (id == "3月FCST")
                {
                    T3 = b;
                }
                if (id == "4月FCST")
                {
                    T4 = b;
                }
                if (id == "5月FCST")
                {
                    T5 = b;
                }
                if (id == "6月FCST")
                {
                    T6 = b;
                }
                if (id == "7月FCST")
                {
                    T7 = b;
                }
                if (id == "8月FCST")
                {
                    T8 = b;
                }
                if (id == "9月FCST")
                {
                    T9 = b;
                }
                if (id == "10月FCST")
                {
                    T10 = b;
                }
                if (id == "11月FCST")
                {
                    T11 = b;
                }
                if (id == "12月FCST")
                {
                    T12 = b;

                }


            }


            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string BU;
                string MODEL;
                string TYPE;
                string GRADE;
                string ITEMCODE;
                string F1 = "";
                string F2 = "";
                string F3 = "";
                string F4 = "";
                string F5 = "";
                string F6 = "";
                string F7 = "";
                string F8 = "";
                string F9 = "";
                string F10 = "";
                string F11 = "";
                string F12 = "";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                BU = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                MODEL = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                TYPE = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                GRADE = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                ITEMCODE = range.Text.ToString().Trim();

                if (T1 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T1]);
                    range.Select();
                    F1 = range.Text.ToString().Trim();
                }

                if (T2 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T2]);
                    range.Select();
                    F2 = range.Text.ToString().Trim();
                }

                if (T3 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T3]);
                    range.Select();
                    F3 = range.Text.ToString().Trim();
                }

                if (T4 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T4]);
                    range.Select();
                    F4 = range.Text.ToString().Trim();
                }


                if (T5 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T5]);
                    range.Select();
                    F5 = range.Text.ToString().Trim();
                }

                if (T6 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T6]);
                    range.Select();
                    F6 = range.Text.ToString().Trim();
                }

                if (T7 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T7]);
                    range.Select();
                    F7 = range.Text.ToString().Trim();
                }

                if (T8 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T8]);
                    range.Select();
                    F8 = range.Text.ToString().Trim();
                }
                if (T9 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T9]);
                    range.Select();
                    F9 = range.Text.ToString().Trim();
                }
                if (T10 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T10]);
                    range.Select();
                    F10 = range.Text.ToString().Trim();
                }
                if (T11 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T11]);
                    range.Select();
                    F11 = range.Text.ToString().Trim();
                }

                if (T12 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T12]);
                    range.Select();
                    F12 = range.Text.ToString().Trim();
                }

                try
                {
                    if (!String.IsNullOrEmpty(BU))
                    {
                        ADDOPOR(BU, MODEL, TYPE, GRADE, ITEMCODE, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12);
                    }

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
            //Quit
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            range = null;
            excelApp = null;
            excelBook = null;
            excelSheet = null;

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();


        }
        private void GD7(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            int T1 = 0;
            int T2 = 0;
            int T3 = 0;
            int T4 = 0;
            int T5 = 0;
            int T6 = 0;
            int T7 = 0;
            int T8 = 0;
            int T9 = 0;
            int T10 = 0;
            int T11 = 0;
            int T12 = 0;

            for (int b = 1; b <= 30; b++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                range.Select();
                string id = range.Text.ToString().Trim().ToUpper().Replace(" ", "");


                if (id == "1月FCST")
                {
                    T1 = b;
                }

                if (id == "2月FCST")
                {
                    T2 = b;
                }
                if (id == "3月FCST")
                {
                    T3 = b;
                }
                if (id == "4月FCST")
                {
                    T4 = b;
                }
                if (id == "5月FCST")
                {
                    T5 = b;
                }
                if (id == "6月FCST")
                {
                    T6 = b;
                }
                if (id == "7月FCST")
                {
                    T7 = b;
                }
                if (id == "8月FCST")
                {
                    T8 = b;
                }
                if (id == "9月FCST")
                {
                    T9 = b;
                }
                if (id == "10月FCST")
                {
                    T10 = b;
                }
                if (id == "11月FCST")
                {
                    T11 = b;
                }
                if (id == "12月FCST")
                {
                    T12 = b;

                }


            }


            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string BU;
                string MODEL;
                string TYPE;
                string GRADE;
                string ITEMCODE;
                string F1 = "";
                string F2 = "";
                string F3 = "";
                string F4 = "";
                string F5 = "";
                string F6 = "";
                string F7 = "";
                string F8 = "";
                string F9 = "";
                string F10 = "";
                string F11 = "";
                string F12 = "";

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                BU = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                MODEL = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                TYPE = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                GRADE = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                ITEMCODE = range.Text.ToString().Trim();

                if (T1 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T1]);
                    range.Select();
                    F1 = range.Text.ToString().Trim();
                }

                if (T2 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T2]);
                    range.Select();
                    F2 = range.Text.ToString().Trim();
                }

                if (T3 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T3]);
                    range.Select();
                    F3 = range.Text.ToString().Trim();
                }

                if (T4 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T4]);
                    range.Select();
                    F4 = range.Text.ToString().Trim();
                }


                if (T5 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T5]);
                    range.Select();
                    F5 = range.Text.ToString().Trim();
                }

                if (T6 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T6]);
                    range.Select();
                    F6 = range.Text.ToString().Trim();
                }

                if (T7 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T7]);
                    range.Select();
                    F7 = range.Text.ToString().Trim();
                }

                if (T8 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T8]);
                    range.Select();
                    F8 = range.Text.ToString().Trim();
                }
                if (T9 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T9]);
                    range.Select();
                    F9 = range.Text.ToString().Trim();
                }
                if (T10 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T10]);
                    range.Select();
                    F10 = range.Text.ToString().Trim();
                }
                if (T11 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T11]);
                    range.Select();
                    F11 = range.Text.ToString().Trim();
                }

                if (T12 != 0)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, T12]);
                    range.Select();
                    F12 = range.Text.ToString().Trim();
                }

                try
                {
                    if (!String.IsNullOrEmpty(BU))
                    {
                        ADDOPOR(BU, MODEL, TYPE, GRADE, ITEMCODE, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12);
                    }

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
            //Quit
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            range = null;
            excelApp = null;
            excelBook = null;
            excelSheet = null;

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();


        }
        private void GD6DD(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;
            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            foreach (Microsoft.Office.Interop.Excel.Worksheet excelSheet in excelBook.Worksheets)
            {
                excelSheet.Activate();
                string NAME = excelSheet.Name.ToString();
                int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
  



                object SelectCell = "A1";
                range = excelSheet.get_Range(SelectCell, SelectCell);

                int T1 = 0;
                int T2 = 0;
                int T3 = 0;
                int T4 = 0;
                int T5 = 0;
                int T6 = 0;
                int T7 = 0;
                int T8 = 0;
                int T9 = 0;
                int T10 = 0;
                int T11 = 0;
                int T12 = 0;

        
                for (int i = 2; i <= iRowCnt; i++)
                {


                    if (iRowCnt > 500)
                    {
                        iRowCnt = 500;
                    }

                    string MODEL;
                    string TYPE;
                    string GRADE;
                    string ITEMCODE;
                    string F1 = "";
                    string F2 = "";
                    string F3 = "";
                    string F4 = "";
                    string F5 = "";
                    string F6 = "";
                    string F7 = "";
                    string F8 = "";
                    string F9 = "";
                    string F10 = "";
                    string F11 = "";
                    string F12 = "";


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    TYPE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    GRADE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                     T1 = NAME.IndexOf(".01");
                     T2 = NAME.IndexOf(".02");
                     T3 = NAME.IndexOf(".03");
                     T4 = NAME.IndexOf(".04");
                     T5 = NAME.IndexOf(".05");
                     T6 = NAME.IndexOf(".06");
                     T7 = NAME.IndexOf(".07");
                     T8 = NAME.IndexOf(".08");
                     T9 = NAME.IndexOf(".09");
                     T10 = NAME.IndexOf(".10");
                     T11 = NAME.IndexOf(".11");
                     T12 = NAME.IndexOf(".12");

                     range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                     range.Select();
                     if (T1 != -1)
                     {
                         F1 = range.Text.ToString().Trim();
                     }

                     if (T2 != -1)
                    {

                        F2 = range.Text.ToString().Trim();
                    }

                     if (T3 != -1)
                    {

                        F3 = range.Text.ToString().Trim();
                    }

                     if (T4 != -1)
                    {
                        F4 = range.Text.ToString().Trim();
                    }


                     if (T5 != -1)
                    {
                        F5 = range.Text.ToString().Trim();
                    }

                     if (T6 != -1)
                    {
                        F6 = range.Text.ToString().Trim();
                    }

                     if (T7 != -1)
                    {
                        F7 = range.Text.ToString().Trim();
                    }

                     if (T8 != -1)
                    {
                        F8 = range.Text.ToString().Trim();
                    }
                    if (T9 != -1)
                    {
                        F9 = range.Text.ToString().Trim();
                    }
                    if (T10 != -1)
                    {
                        F10 = range.Text.ToString().Trim();
                    }
                    if (T11 != -1)
                    {
                        F11 = range.Text.ToString().Trim();
                    }

                    if (T12 != -1)
                    {
                        F12 = range.Text.ToString().Trim();
                    }

                    try
                    {
                        if (!String.IsNullOrEmpty(MODEL))
                        {
                            ADDOPOR("DT", MODEL, TYPE, GRADE, ITEMCODE, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12);
                        }

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

            }




       
       

            
        //Quit
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
     
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            range = null;
            excelApp = null;
            excelBook = null;
     

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();


        }


        public void ADDOPOR(string BU, string MODEL, string TYPE, string GRADE, string ITEMCODE, string F1, string F2, string F3, string F4, string F5, string F6, string F7, string F8, string F9, string F10, string F11, string F12)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_KIT9(BU,MODEL,TYPE,GRADE,ITEMCODE,F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,USERS,DOCDATE) values(@BU,@MODEL,@TYPE,@GRADE,@ITEMCODE,@F1,@F2,@F3,@F4,@F5,@F6,@F7,@F8,@F9,@F10,@F11,@F12,@USERS,@DOCDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@F1", F1));
            command.Parameters.Add(new SqlParameter("@F2", F2));
            command.Parameters.Add(new SqlParameter("@F3", F3));
            command.Parameters.Add(new SqlParameter("@F4", F4));
            command.Parameters.Add(new SqlParameter("@F5", F5));
            command.Parameters.Add(new SqlParameter("@F6", F6));
            command.Parameters.Add(new SqlParameter("@F7", F7));
            command.Parameters.Add(new SqlParameter("@F8", F8));
            command.Parameters.Add(new SqlParameter("@F9", F9));
            command.Parameters.Add(new SqlParameter("@F10", F10));
            command.Parameters.Add(new SqlParameter("@F11", F11));
            command.Parameters.Add(new SqlParameter("@F12", F12));
           
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DateTime.Now.ToString("MMdd")));
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

        public void ADDAP_KIT9D2(int MID, int COMUMN, int ROW)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_KIT9D2(MID,COMUMN,ROW,USERS) values(@MID,@COMUMN,@ROW,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MID", MID));
            command.Parameters.Add(new SqlParameter("@COMUMN", COMUMN));
            command.Parameters.Add(new SqlParameter("@ROW", ROW));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public void DELAP_KIT9D2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_KIT9D2 WHERE USERS=@USERS", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public void UPAP_KIT9(string COLUMN, string ID,int S1)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE AP_KIT9 SET  " + COLUMN + "=@S1 WHERE ID=@ID", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@S1", S1));
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
        public System.Data.DataTable GETKIT9D(string COLUMN,string ID)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = globals.Connection;
            sb.Append(" select " + COLUMN + " from AP_KIT9D WHERE ID=@ID");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT4DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT4DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT4DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dDPIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ADD2(lC.AP_KIT4, lC.AP_KIT9, aP_KIT4DataGridView, aP_KIT9DataGridView);

            this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
        }

        private void aP_KIT9DataGridView_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {

        }

        private void aP_KIT9DataGridView_DoubleClick(object sender, EventArgs e)
        {
            if (aP_KIT9DataGridView.SelectedRows.Count > 0)
            {

                string da = aP_KIT9DataGridView.SelectedRows[0].Cells["ID9"].Value.ToString();

                AP_KIT3D a = new AP_KIT3D();
                a.PublicString = da;

                a.ShowDialog();
            }
        }
        //int TTS = 0;
        //int TTE = 0;
        //int TT1 = 0;
        //int TT2 = 0;
        //int TT3 = 0;
        private void aP_KIT9DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
           // TTE = TT1;

           // for (int i = 0; i <= TTE - 1; i++)
           //{

           //}
           //TT3 = 0;

           //DELAP_KIT9D2();
        }

        private void aP_KIT9DataGridView_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {

          //  try
          //  {
          //      if (TT3 == 1)
          //      {
          //          TT1 = e.RowIndex;
          //          TTS = e.RowIndex;
          //          TT2 = e.ColumnIndex;
          //          string ID = aP_KIT9DataGridView.Rows[e.RowIndex].Cells[0].Value.ToString();
          //          //ADDAP_KIT9D2(Convert.ToInt32(ID), TT2, TT1);
          //          aP_KIT9DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Yellow;

          //      }
          ////      MessageBox.Show("B");
          //  }
          //  catch { }
        }

        private void aP_KIT9DataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
        
          
        }

        private void aP_KIT9DataGridView_MouseClick(object sender, MouseEventArgs e)
        {
          
        }
        private void SS()
        {

            int n;
            if (int.TryParse(textBox6.Text, out n) && int.TryParse(textBox7.Text, out n))
            {
                int T1 = Convert.ToInt16(textBox6.Text);
                int T2 = Convert.ToInt16(textBox7.Text);
                for (int i = 1; i <= 12; i++)
                {
                    string F1 = "F" + i.ToString();
 
                    aP_KIT9DataGridView.Columns[F1].Visible = false;

                }

                for (int i = T1; i <= T2; i++)
                {
                    string F1 = "F" + i.ToString();
           
                    aP_KIT9DataGridView.Columns[F1].Visible = true;

                }
            }
        }

        private void SS2(string S1,string S2,DataGridView DG1)
        {

            int n;
            if (int.TryParse(S1, out n) && int.TryParse(S2, out n))
            {
                int T1 = Convert.ToInt16(S1);
                int T2 = Convert.ToInt16(S2);
                //F1G1 aP_KIT3DataGridView
                int DG = Convert.ToInt16(DG1.Name.ToString().Substring(6, 1)) - 2;
                for (int i = 1; i <= 12; i++)
                {
                    string F1 = "F" + i.ToString() + "G" + DG.ToString();
                    DG1.Columns[F1].Visible = false;
                }

                for (int i = T1; i <= T2; i++)
                {
                    string F1 = "F" + i.ToString() + "G" + DG.ToString();
                    DG1.Columns[F1].Visible = true;
                }
            }
        }
        private void aP_KIT9DataGridView_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex == -1 || e.ColumnIndex == -1) return;
            //var cell = aP_KIT9DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
      
            //    if (cell.Value != null)
            //    {
            //        if (!String.IsNullOrEmpty(cell.Value.ToString()))
            //        {
            //            string COLUMN = "";
            //            if (e.ColumnIndex == 1)
            //            {
            //                COLUMN = "BU";
            //            }
            //            else if (e.ColumnIndex == 2)
            //            {
            //                COLUMN = "MODEL";
            //            }
            //            else if (e.ColumnIndex == 3)
            //            {
            //                COLUMN = "TYPE";
            //            }
            //            else if (e.ColumnIndex == 4)
            //            {
            //                COLUMN = "GRADE";
            //            }
            //            else if (e.ColumnIndex == 5)
            //            {
            //                COLUMN = "ITEMCODE";
            //            }
            //            else if (e.ColumnIndex == 6)
            //            {
            //                COLUMN = "F1";
            //            }
            //            else if (e.ColumnIndex == 7)
            //            {
            //                COLUMN = "F2";
            //            }
            //            else if (e.ColumnIndex == 8)
            //            {
            //                COLUMN = "F3";
            //            }
            //            else if (e.ColumnIndex == 9)
            //            {
            //                COLUMN = "F4";
            //            }
            //            else if (e.ColumnIndex == 10)
            //            {
            //                COLUMN = "F5";
            //            }
            //            else if (e.ColumnIndex == 11)
            //            {
            //                COLUMN = "F6";
            //            }
            //            else if (e.ColumnIndex == 12)
            //            {
            //                COLUMN = "F7";
            //            }
            //            else if (e.ColumnIndex == 13)
            //            {
            //                COLUMN = "F8";
            //            }
            //            else if (e.ColumnIndex == 14)
            //            {
            //                COLUMN = "F9";
            //            }
            //            else if (e.ColumnIndex == 15)
            //            {
            //                COLUMN = "F10";
            //            }
            //            else if (e.ColumnIndex == 16)
            //            {
            //                COLUMN = "F11";
            //            }
            //            else if (e.ColumnIndex == 17)
            //            {
            //                COLUMN = "F12";
            //            }
            //            else if (e.ColumnIndex == 18)
            //            {
            //                COLUMN = "F1S";
            //            }
            //            else if (e.ColumnIndex == 19)
            //            {
            //                COLUMN = "F2S";
            //            }
            //            else if (e.ColumnIndex == 20)
            //            {
            //                COLUMN = "F3S";
            //            }
            //            else if (e.ColumnIndex == 21)
            //            {
            //                COLUMN = "F4S";
            //            }
            //            else if (e.ColumnIndex == 22)
            //            {
            //                COLUMN = "F5S";
            //            }
            //            else if (e.ColumnIndex == 23)
            //            {
            //                COLUMN = "F6S";
            //            }
            //            else if (e.ColumnIndex == 24)
            //            {
            //                COLUMN = "F7S";
            //            }
            //            else if (e.ColumnIndex == 25)
            //            {
            //                COLUMN = "F8S";
            //            }
            //            else if (e.ColumnIndex == 26)
            //            {
            //                COLUMN = "F9S";
            //            }
            //            else if (e.ColumnIndex == 27)
            //            {
            //                COLUMN = "F10S";
            //            }
            //            else if (e.ColumnIndex == 28)
            //            {
            //                COLUMN = "F11S";
            //            }
            //            else if (e.ColumnIndex == 29)
            //            {
            //                COLUMN = "F12S";
            //            }

            //        toolTip1 = new ToolTip();

            //        string ID9 = aP_KIT9DataGridView.Rows[e.RowIndex].Cells["ID9"].Value.ToString();
            //        if (COLUMN != "")
            //        {
            //            System.Data.DataTable gf = GETKIT9D(COLUMN, ID9);
            //            if (gf.Rows.Count > 0)
            //            {
            //                toolTip1.SetToolTip(aP_KIT9DataGridView, gf.Rows[0][0].ToString());
            //            }

            //        }
            //    }
            //}
        }

        private void aP_KIT9DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
                 //DataGridView dgv = (DataGridView)sender;

                 //if (dgv.Columns[e.ColumnIndex].Name == "BU")
                 //{
                 //    string ID9 = aP_KIT9DataGridView.CurrentRow.Cells["ID9"].Value.ToString();

                 //    UPKIT9(ID9);
                 //    this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
                 //}
        }

        private void aP_KIT9DataGridView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //TT3 = 1;
            //TTS = 99;
            //DELAP_KIT9D2();
        }

        private void aP_KIT9DataGridView_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
        
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            SS();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SS();
        }

        private void 黑ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }


        private void aP_KIT9DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= aP_KIT9DataGridView.Rows.Count)
                return;
            DataGridViewRow dgr = aP_KIT9DataGridView.Rows[e.RowIndex];
            try
            {

                if (dgr.Cells["EBU"].Value.ToString() == "0")
                {
                    aP_KIT9DataGridView.Rows[e.RowIndex].Cells["BU"].Style.BackColor = Color.White;
                }
                if (dgr.Cells["EBU"].Value.ToString() == "1")
                {
                    aP_KIT9DataGridView.Rows[e.RowIndex].Cells["BU"].Style.BackColor = Color.Yellow;
                }
                if (dgr.Cells["EBU"].Value.ToString() == "2")
                {
                    aP_KIT9DataGridView.Rows[e.RowIndex].Cells["BU"].Style.BackColor = Color.Red;
                }
            }
            catch 
            {
 
            } 
        }

        private void textTV1_TextChanged(object sender, EventArgs e)
        {
            SS2(textTV1.Text, textTV2.Text, aP_KIT3DataGridView);
        }

        private void textTV2_TextChanged(object sender, EventArgs e)
        {
            SS2(textTV1.Text, textTV2.Text, aP_KIT3DataGridView);
        }

        private void textPID1_TextChanged(object sender, EventArgs e)
        {
            SS2(textPID1.Text, textPID2.Text, aP_KIT4DataGridView);
        }

        private void textPID2_TextChanged(object sender, EventArgs e)
        {
            SS2(textPID1.Text, textPID2.Text, aP_KIT4DataGridView);
        }

        private void textGD1_TextChanged(object sender, EventArgs e)
        {
            SS2(textGD1.Text, textGD2.Text, aP_KIT5DataGridView);
        }

        private void textGD2_TextChanged(object sender, EventArgs e)
        {
            SS2(textGD1.Text, textGD2.Text, aP_KIT5DataGridView);
        }

        private void textDT1_TextChanged(object sender, EventArgs e)
        {
            SS2(textDT1.Text, textDT2.Text, aP_KIT6DataGridView);
        }

        private void textDT2_TextChanged(object sender, EventArgs e)
        {
            SS2(textDT1.Text, textDT2.Text, aP_KIT6DataGridView);
        }

        private void textNB1_TextChanged(object sender, EventArgs e)
        {
            SS2(textNB1.Text, textNB2.Text, aP_KIT7DataGridView);
        }

        private void textNB2_TextChanged(object sender, EventArgs e)
        {
            SS2(textNB1.Text, textNB2.Text, aP_KIT7DataGridView);
        }



        private void button22_Click_1(object sender, EventArgs e)
        {
            CalcTotals5S();
        }
        private void CalcTotals5S()
        {

            try
            {
                Int32 iTotal = 0;
                DataGridViewRow row;

                int i = this.aP_KIT9DataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    row = aP_KIT9DataGridView.SelectedRows[iRecs];


                    string F1 = row.Cells["F1"].Value.ToString();
                    if (String.IsNullOrEmpty(F1))
                    {
                        F1 = "0";
                    }
                    string F2 = row.Cells["F2"].Value.ToString();
                    if (String.IsNullOrEmpty(F2))
                    {
                        F2 = "0";
                    }
                    string F3 = row.Cells["F3"].Value.ToString();
                    if (String.IsNullOrEmpty(F3))
                    {
                        F3 = "0";
                    }
                    string F4 = row.Cells["F4"].Value.ToString();
                    if (String.IsNullOrEmpty(F4))
                    {
                        F4 = "0";
                    }
                    string F5 = row.Cells["F5"].Value.ToString();
                    if (String.IsNullOrEmpty(F5))
                    {
                        F5 = "0";
                    }
                    string F6 = row.Cells["F6"].Value.ToString();
                    if (String.IsNullOrEmpty(F6))
                    {
                        F6 = "0";
                    }
                    string F7 = row.Cells["F7"].Value.ToString();
                    if (String.IsNullOrEmpty(F7))
                    {
                        F7 = "0";
                    }
                    string F8 = row.Cells["F8"].Value.ToString();
                    if (String.IsNullOrEmpty(F8))
                    {
                        F8 = "0";
                    }
                    string F9 = row.Cells["F9"].Value.ToString();
                    if (String.IsNullOrEmpty(F9))
                    {
                        F9 = "0";
                    }
                    string F10 = row.Cells["F10"].Value.ToString();
                    if (String.IsNullOrEmpty(F10))
                    {
                        F10 = "0";
                    }
                    string F11 = row.Cells["F11"].Value.ToString();
                    if (String.IsNullOrEmpty(F11))
                    {
                        F11 = "0";
                    }
                    string F12 = row.Cells["F12"].Value.ToString();
                    if (String.IsNullOrEmpty(F12))
                    {
                        F12 = "0";
                    }
           
                    iTotal += Convert.ToInt32(F1)
                        + Convert.ToInt32(F2)
                        + Convert.ToInt32(F3)
                        + Convert.ToInt32(F4)
                        + Convert.ToInt32(F5)
                        + Convert.ToInt32(F6)
                        + Convert.ToInt32(F7)
                        + Convert.ToInt32(F8)
                        + Convert.ToInt32(F9)
                        + Convert.ToInt32(F10)
                        + Convert.ToInt32(F11)
                        + Convert.ToInt32(F12);

                }

                textBox11.Text = iTotal.ToString("0");


            }
            catch { }


        }

        private void button23_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result2 = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                GD6DD(opdf.FileName);
                this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
            }
        }

        private void aP_KIT9DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {

        }


        private void button24_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result2 = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                GD7(opdf.FileName);
                this.aP_KIT9TableAdapter.FillBy(this.lC.AP_KIT9, textITEMNAME.Text, STATUS);
            }
        }


       
    

     
    }
}

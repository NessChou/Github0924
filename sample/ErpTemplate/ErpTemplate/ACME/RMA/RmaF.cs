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
    public partial class RmaF : ACME.fmBase1
    {

        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        int scrollPosition = 0;
        public RmaF()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            rma_mainFTableAdapter.Connection = MyConnection;
            rma_InvoiceFTableAdapter.Connection = MyConnection;

        }


        public override void AfterEdit()
        {


            uUSERTextBox.Text = fmLogin.LoginID.ToString();
            shippingCodeTextBox.ReadOnly = true;
            cUSERTextBox.ReadOnly = true;
            uUSERTextBox.ReadOnly = true;

        }

        public override void AfterCancelEdit()
        {
            Control();

        }
        private void Control()
        {

            cUSERTextBox.ReadOnly = true;
            uUSERTextBox.ReadOnly = true;
            shippingCodeTextBox.ReadOnly = true;

            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;

            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button6.Enabled = true;
            comboBox1.Enabled = true;
            textBox3.Enabled = true;
        }

        public override void AfterAddNew()
        {
            Control();
        }
        public override void EndEdit()
        {
            Control();
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                rm.Rma_InvoiceF.RejectChanges();

            }
            catch
            {
            }

            return true;
        }
        public override void SetInit()
        {

            MyBS = rma_mainFBindingSource;
            MyTableName = "Rma_mainF";
            MyIDFieldName = "ShippingCode";


        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "RMF" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            kyes = this.shippingCodeTextBox.Text;
            cUSERTextBox.Text = fmLogin.LoginID.ToString();

            this.rma_mainFBindingSource.EndEdit();
            kyes = null;
        }
        public override void FillData()
        {
            try
            {

                rma_mainFTableAdapter.Fill(rm.Rma_mainF, MyID);
                rma_InvoiceFTableAdapter.Fill(rm.Rma_InvoiceF, MyID);

                string g = rm.Rma_InvoiceF.Compute("Sum(QTY)", null).ToString();
                string g2 = rm.Rma_InvoiceF.Compute("Sum(RQTY)", null).ToString();
                string g3 = rm.Rma_InvoiceF.Compute("Sum(VQTY)", null).ToString();
                if (String.IsNullOrEmpty(g))
                {
                    g = "0";
                }
                if (String.IsNullOrEmpty(g2))
                {
                    g2 = "0";
                }
                if (String.IsNullOrEmpty(g3))
                {
                    g3 = "0";
                }
                decimal sh = Convert.ToDecimal(g);
                decimal sh2 = Convert.ToDecimal(g2);
                decimal sh3 = Convert.ToDecimal(g3);

                label5.Text = "AU本次還貨數量 : "+sh.ToString("#,##0");
                label6.Text = "原退數量 : " + sh2.ToString("#,##0");
                label7.Text = "AU已還數量 : " + sh3.ToString("#,##0");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public override bool UpdateData()
        {
            bool UpdateData;

            try
            {
                Validate();

                rma_InvoiceFBindingSource.MoveFirst();
                for (int i = 0; i <= rma_InvoiceFBindingSource.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)rma_InvoiceFBindingSource.Current;
                    row["SEQNO"] = i;
                    rma_InvoiceFBindingSource.EndEdit();
                    rma_InvoiceFBindingSource.MoveNext();
                }



                rma_mainFTableAdapter.Connection.Open();


                rma_mainFBindingSource.EndEdit();
                rma_InvoiceFBindingSource.EndEdit();

                rma_mainFTableAdapter.Update(rm.Rma_mainF);
                rma_InvoiceFTableAdapter.Update(rm.Rma_InvoiceF);

                UpdateData = true;
            }
            catch (Exception ex)
            {

                ////NG
                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.rma_mainFTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        public override void AfterEndEdit()
        {
            //try
            //{


            //    if (rma_InvoiceFDataGridView.Rows.Count > 1)
            //    {

         

            //        System.Data.DataTable dt1 = rm.Rma_InvoiceD;
            //        try
            //        {
            //            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            //            {
            //                DateTime R1 = Convert.ToDateTime(GetMenu.DayS(dOCDATETextBox.Text));

            //                DataRow drw = dt1.Rows[i];
            //                string RMANO = drw["RMANO"].ToString();
            //                string QTY = drw["QTY"].ToString();
            //                UPDATEJOBNO(R1, QTY, RMANO);
            //            }
            //        }
            //        catch { }
            //    }




            //    rma_InvoiceFBindingSource.EndEdit();
            //    rma_mainFTableAdapter.Update(rm.Rma_mainF);

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }
   
        private void button1_Click(object sender, EventArgs e)
        {
            string strCollected = string.Empty;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    if (strCollected == string.Empty)
                    {
                        strCollected = checkedListBox1.GetItemText(
         checkedListBox1.Items[i]);
                    }
                    else
                    {
                        strCollected = strCollected + checkedListBox1.
         GetItemText(checkedListBox1.Items[i]);
                    }
                }
            }
            string FD = strCollected;
            if (FD == "")
            {
                MessageBox.Show("請選擇公司");
                return;
            }

      
            RmaNo frm1 = new RmaNo();
            frm1.q1 = FD;
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Data.DataTable dt1 = GetAR2(frm1.q, frm1.q2);
                    System.Data.DataTable dt2 = rm.Rma_InvoiceF;

                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();

                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["COMPANY"] = drw["COMPANY"].ToString().Trim();
                        drw2["RMANO"] = drw["U_RMA_NO"].ToString().Trim();
                        drw2["MODEL"] = drw["U_RMODEL"].ToString().Trim();
                        drw2["SEQNO"] = "0";
                        drw2["VENDER"] = drw["U_AUO_RMA_NO"].ToString().Trim();
                        drw2["CARDNAME"] = drw["U_cusname_s"].ToString().Trim();
                        drw2["VER"] = drw["U_Rver"].ToString().Trim();
                        drw2["GRADE"] = drw["U_Rgrade"].ToString().Trim();
                        drw2["SORTES2"] = drw["U_RENGINEER"].ToString().Trim();
                        drw2["CONTRACTID"] = drw["CONTRACTID"].ToString().Trim();
                        string K1 = drw["U_Rquinity"].ToString();
                        if (ValidateUtils.IsNumeric(K1))
                        {
                            drw2["RQTY"] = K1;
                        }
                        else
                        {
                            drw2["RQTY"] = "0";
                        }

                        int g = drw["U_Rquinity"].ToString().IndexOf("+");
                        int t = drw["U_Rquinity"].ToString().LastIndexOf("+");
                        string h;
                        string s;


                        if (g == -1)
                        {
                            s = drw["U_Rquinity"].ToString();
                           // drw2["QTY"] = s;
                        }
                        else
                        {
                            s = drw["U_Rquinity"].ToString().Substring(g + 1);

                            try
                            {

                                if (drw["U_Rquinity"].ToString().Substring(3, 1) != "+")
                                {
                                    h = drw["U_Rquinity"].ToString().Substring(0, 2);
                                }
                                else
                                {
                                    h = drw["U_Rquinity"].ToString().Substring(0, 1);
                                }
                                int a = Convert.ToInt16(s.ToString());
                                int b = Convert.ToInt16(h.ToString());
                              //  drw2["QTY"] = (a + b).ToString();
                            }
                            catch (Exception ex)
                            {
                                h = drw["U_Rquinity"].ToString().Substring(0, 1);
                               // drw2["QTY"] = h.ToString();
                            }

                        }


                        if (drw["U_ACME_QBACK"].ToString() == "")
                        {
                            drw2["VQTY"] = "0";
                        }
                        else 
                        {

                        int g2 = drw["U_ACME_QBACK"].ToString().IndexOf("+");
                        int t2 = drw["U_ACME_QBACK"].ToString().LastIndexOf("+");
                        string h2;
                        string s2;


                        if (g == -1)
                        {
                            s2 = drw["U_ACME_QBACK"].ToString();
                            drw2["VQTY"] = s2;
                        }
                        else
                        {
                            s2 = drw["U_ACME_QBACK"].ToString().Substring(g2 + 1);

                            try
                            {

                                if (drw["U_ACME_QBACK"].ToString().Substring(3, 1) != "+")
                                {
                                    h2 = drw["U_ACME_QBACK"].ToString().Substring(0, 2);
                                }
                                else
                                {
                                    h2 = drw["U_ACME_QBACK"].ToString().Substring(0, 1);
                                }
                                int a2 = Convert.ToInt16(s2.ToString());
                                int b2 = Convert.ToInt16(h2.ToString());

                                drw2["VQTY"] = (a2 + b2).ToString();
                            }
                            catch (Exception ex)
                            {
                                h2 = drw["U_ACME_QBACK"].ToString().Substring(0, 1);
                                drw2["VQTY"] = h2.ToString();
                            }

                        }
                        }

           

                        dt2.Rows.Add(drw2);
                    }
                    for (int j = 0; j <= rma_InvoiceFDataGridView.Rows.Count - 2; j++)
                    {
                        rma_InvoiceFDataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {

            }
        }

        public System.Data.DataTable GetAR2(string DocEntry, string DocEntry2)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            string aa = '"'.ToString();
            if (!String.IsNullOrEmpty(DocEntry) && !String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("select '進金生' COMPANY,CONTRACTID,U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,U_ACME_QBACK U_ACME_QBACK,aa=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END,U_RENGINEER from acmesql02.DBO.octr where Contractid IN (" + DocEntry + ") ");
                sb.Append("UNION ALL ");
                sb.Append("select '達睿生' COMPANY,CONTRACTID,U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,U_ACME_QBACK U_ACME_QBACK,aa=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END,U_RENGINEER from acmesql05.DBO.octr where Contractid IN (" + DocEntry2 + ") ");
            }
            if (!String.IsNullOrEmpty(DocEntry) && String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("select '進金生' COMPANY,CONTRACTID,U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,U_ACME_QBACK U_ACME_QBACK,aa=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END,U_RENGINEER from acmesql02.DBO.octr where Contractid IN (" + DocEntry + ") ");
            }
            if (String.IsNullOrEmpty(DocEntry) && !String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("select '達睿生' COMPANY,CONTRACTID,U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,U_ACME_QBACK U_ACME_QBACK,aa=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END,U_RENGINEER from acmesql05.DBO.octr where Contractid IN (" + DocEntry2 + ") ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " octr ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" octr "];
        }

        private void RmaF_Load(object sender, EventArgs e)
        {
            Control();
            checkedListBox1.SetItemChecked(0, true);
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            textBox4.Text = GetMenu.DFirst();
            textBox5.Text = GetMenu.DLast();
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string GlobalMailContent = "";
                FileName = lsAppDir + "\\Excel\\RMA\\通知覆判明細.xls";


                System.Data.DataTable OrderData = GetOrderData3();


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                //請安排復判AU還回RMA _日期  
                string DATE = dOCDATETextBox.Text.Substring(4, 2) + "/" + dOCDATETextBox.Text.Substring(6, 2);
                string SUBJECT = "請安排復判AU還回RMA _" + DATE;
                string DATA = DATE + " AU还回如下RMA，请安排复判，谢谢！";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                if (OrderData.Rows.Count  > 0)
                {
                    dataGridView1.DataSource = OrderData;
                    GlobalMailContent = htmlMessageBody(dataGridView1).ToString();
                    //  ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile,"N");
                    MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, GlobalMailContent, DATA);

                    MessageBox.Show("信件已寄送");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  今日無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }

            //create html & table
            //strB.AppendLine("<html><body><center><table border='1' cellpadding='0' cellspacing='0'>");
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

                if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                {

                    //if (i != 0)
                    //{
                    //    strB.AppendLine("<tr class='HeaderBorder'>");
                    //    for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                    //    {
                    //        strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                    //    }
                    //    strB.AppendLine("</tr>");
                    //}

                    //處理鍵值




                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;

                    if (KeyValue.IndexOf("{") >= 0)
                    {
                        tmpKeyValue = "";
                    }
                }
                else
                {
                    tmpKeyValue = "";
                }


                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }


                DataGridViewCell dgvc;


                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td align='center'>" + tmpKeyValue + "</td>");
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
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
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
                                    string sDate = "";
                                    if (dgvc.Value.ToString() != "小計")
                                    {
                                        sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                    dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                    dgvc.Value.ToString().Substring(6, 2);
                                        strB.AppendLine("<td>" + sDate + "</td>");
                                    }
                                    else
                                    {
                                        strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                                    }


                                }
                            }

                            else
                            {
                                strB.AppendLine("<td align='center'>" + dgvc.Value.ToString() + "</td>");
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
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent, string DATA)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("LleytonChen@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));
            message.To.Add(new MailAddress("leogeng@acmepoint.com"));
            message.To.Add(new MailAddress("ianli@acmepoint.com"));
            message.To.Add(new MailAddress("erinchou@acmepoint.com"));
            //erinchou@acmepoint.com
            string template;
            StreamReader objReader;

            objReader = new StreamReader(GetExePath() + "\\MailTemplates\\RMALEMON.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##Content##", MailContent);
            template = template.Replace("##DATA##", DATA);
            message.Subject = strSubject;
            message.Body = template;
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {


                message.Attachments.Add(new Attachment(file));

            }

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
                //data.Dispose();
                //message.Dispose();
                //DELETEFILE();
                foreach (Attachment item in message.Attachments)
                {
                    item.Dispose();   //一定要释放该对象,否则无法删除附件
                }
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        //  SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        // SetMsg(String.Format("Failed to deliver message to {0}",
                        // ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                //        ex.ToString()));
            }

        }
        private System.Data.DataTable GetOrderData3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT RMANO 'RMA NO',CARDNAME 客戶簡稱,Model,Ver,Grade,QTY AU本次還貨,VENDER 'Vender RMA NO.',RQTY '原退數量',VQTY 'AU已還數量',OK 'OK（pcs）',NG 'NG（pcs）',Remark,SORTES2 '原Sorting ES.' FROM dbo.Rma_InvoiceF");
            sb.Append(" WHERE SHIPPINGCODE = @SHIPPINGCODE ");
            sb.Append(" UNION ALL ");
            sb.Append("SELECT '','','','','TOTAL:',SUM(CAST(QTY AS INT)) ,'',SUM(CAST(RQTY AS INT)),SUM(CAST(VQTY AS INT)),SUM(CAST(OK AS INT)) ,SUM(CAST(NG AS INT)),'',''  FROM dbo.Rma_InvoiceF");
            sb.Append(" WHERE SHIPPINGCODE = @SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetNOTRETURN()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT VQTY,RMANO,MODEL,VER,GRADE,RQTY,''''+VENDER VENDER,QTY,OK,NG,REMARK,CARDNAME,DOCDATE 收貨日期,SORTES2 ES FROM dbo.Rma_InvoiceF T0");
            sb.Append("         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T1.DOCDATE BETWEEN @A1 AND @A2  AND  ( ISNULL(OK,'') ='' AND  ISNULL(NG,'') ='' AND  ISNULL(SORTDATE,'') ='' AND  ISNULL(SORTES,'') =''   AND  ISNULL(T1.RMASELECT,'') ='' )  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox5.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData4(string MODEL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT VQTY,RMANO,MODEL,VER,GRADE,RQTY,''''+VENDER VENDER,QTY,OK,NG,REMARK,CARDNAME,SORTDATE,SORTES,DOCDATE 收貨日期,T0.RMASELECT FROM dbo.Rma_InvoiceF T0");
            sb.Append(" LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T1.DOCDATE BETWEEN @A1 AND @A2   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }

            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }

    
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData43(string MODEL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CONVERT(VARCHAR(10) ,CAST(DOCDATE AS DATETIME), 111 )  收貨日期,CENTER 維修廠,RMANO,CARDNAME,MODEL,''''+VER VER,''''+VENDER VENDER,U_S_seq  SN");
            sb.Append(" , ''''+U_U_month_seq  WC, U_U_iqc IQC,U_U_C_complain COMPLAIN");
            sb.Append(" , U_U_acme_confirm CONFIRM, U_U_Acme_judge JUDGE, U_U_PLACE_1  產地, U_rRemark  REMARK");
            sb.Append(" ,SORTES,T1.RMASELECT 檢測方式 FROM RMA_MAINF T0");
            sb.Append(" LEFT JOIN RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" LEFT JOIN RMA_CTR1 T2 ON (T1.RMANO=T2.U_RMA_NO COLLATE  Chinese_Taiwan_Stroke_CI_AS AND T0.DOCDATE=T2.ManufSN )  ");
            sb.Append(" WHERE T0.DOCDATE BETWEEN @A1 AND @A2 AND U_U_Acme_judge='NG'   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T0.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T1.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN  T1.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }

        
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetC1(string H1,string H2,string DD,string MODEL)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" Select  [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("             SELECT SUM(ISNULL(QTY,0)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" UNION ALL");
            sb.Append(" Select [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("              SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END ,0) AS INT))+SUM(CAST(CASE NG WHEN '' THEN '0' ELSE NG END AS INT)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" UNION ALL");
            sb.Append(" Select [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("              SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS INT)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" UNION ALL");
            sb.Append(" Select [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("              SELECT SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS INT)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" UNION ALL");
            sb.Append(" Select [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("                   SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) END C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");
            sb.Append(" UNION ALL");
            sb.Append(" Select [" + DD + "01] AS '1',[" + DD + "02] AS '2',[" + DD + "03] AS '3',[" + DD + "04] AS '4',[" + DD + "05] AS '5',[" + DD + "06] AS '6',[" + DD + "07] AS '7',[" + DD + "08] AS '8',[" + DD + "09] AS '9',[" + DD + "10] AS '10',[" + DD + "11] AS '11',[" + DD + "12] AS '12'");
            sb.Append(" from (");
            sb.Append("                               SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) END C1,SUBSTRING(DOCDATE,1,6)  DOCDATE  FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("              LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE SUBSTRING(DOCDATE,1,6) BETWEEN @A1 AND @A2   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" GROUP BY SUBSTRING(DOCDATE,1,6) ");
            sb.Append(" ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(C1)");
            sb.Append(" FOR DOCDATE  IN");
            sb.Append(" ( [" + DD + "01],[" + DD + "02],[" + DD + "03],[" + DD + "04],[" + DD + "05],[" + DD + "06],[" + DD + "07],[" + DD + "08],[" + DD + "09],[" + DD + "10],[" + DD + "11],[" + DD + "12] )");
            sb.Append(" ) AS pvt");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", H1));
            command.Parameters.Add(new SqlParameter("@A2", H2));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetBU(string aa)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT PARAM_DESC as DataText FROM RMA_PARAMS where param_kind=@aa  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));

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
        public System.Data.DataTable GetY(string CONTRACTID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT U_ACME_QBACK FROM ACMESQL02.DBO.OCTR WHERE CONTRACTID =@CONTRACTID ");
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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

        public System.Data.DataTable GetY5(string CONTRACTID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT U_ACME_QBACK FROM ACMESQL05.DBO.OCTR WHERE CONTRACTID =@CONTRACTID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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
        public  System.Data.DataTable GetF(string MODEL)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT DISTINCT SUBSTRING(DOCDATE,1,6) DOCDATE   FROM dbo.Rma_MAINF T0  LEFT JOIN Rma_InvoiceF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.DOCDATE BETWEEN @AA AND @BB ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T0.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T1.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T1.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
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
        public  System.Data.DataTable GetF2(string AA,string MODEL)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT SUM(ISNULL(QTY,0)) C1, SUBSTRING(DOCDATE,1,6) DOCDATE   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
       
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS INT))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS INT)) C1, SUBSTRING(DOCDATE,1,6) DOCDATE   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS INT)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS INT)) C1,SUBSTRING(DOCDATE,1,6)  DOCDATE   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            sb.Append(" UNION ALL");
            sb.Append("        SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0'  ELSE NG END,0) AS DECIMAL))) END C1,SUBSTRING(DOCDATE,1,6)  DOCDATE   FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            sb.Append(" UNION ALL");
            sb.Append("                              SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL)))  WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL)))  END C1,SUBSTRING(DOCDATE,1,6) DOCDATE   FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE SUBSTRING(DOCDATE,1,6) = @AA  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append("           GROUP BY SUBSTRING(DOCDATE,1,6)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", AA));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
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

        public  System.Data.DataTable GetF3(string MODEL)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT SUM(ISNULL(QTY,0)) C1   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE DOCDATE BETWEEN @AA AND @BB  ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK  END,0) AS INT))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS INT)) C1   FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("    WHERE  DOCDATE BETWEEN @AA AND @BB");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK END ,0) AS INT)) C1  FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("    WHERE  DOCDATE BETWEEN @AA AND @BB   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("               SELECT SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS INT)) C1  FROM dbo.Rma_InvoiceF T0");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("    WHERE  DOCDATE BETWEEN @AA AND @BB   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("        SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0' ELSE OK END ,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))) END C1  FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("    WHERE  DOCDATE BETWEEN @AA AND @BB   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("                              SELECT CASE (SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL)))  WHEN 0 THEN 0 ELSE SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL))/(SUM(CAST(ISNULL(CASE OK WHEN '' THEN '0'  ELSE OK END ,0) AS DECIMAL))+SUM(CAST(ISNULL(CASE NG WHEN '' THEN '0' ELSE NG END,0) AS DECIMAL)))  END C1  FROM dbo.Rma_InvoiceF T0 ");
            sb.Append("                         LEFT JOIN dbo.Rma_MAINF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("    WHERE  DOCDATE BETWEEN @AA AND @BB   ");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND  T1.CENTER = @A3   ");
            }
            if (MODEL != "")
            {
                sb.Append(" AND CASE  WHEN T0.MODEL LIKE '%OPEN%' THEN 'OPEN CELL'   WHEN T0.MODEL LIKE '%T CON%' THEN 'T CON'  ELSE 'MODULE' END   in ( " + MODEL + " )  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@A3", comboBox1.Text));
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

        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;

            comboBox2.SelectedIndex = -1;

        }



        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text.Substring(0, 4) != textBox2.Text.Substring(0, 4))
                {
                    MessageBox.Show("收貨日期請輸入同個年度");
                    return;
                }



                string T1 = textBox1.Text.Substring(0, 4);

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                if (comboBox1.Text == "BVC(景智)")
                {

                    FileName = lsAppDir + "\\Excel\\RMA\\AU還回RMA完修率統計BVC.xls";
                }
                else
                {
                    FileName = lsAppDir + "\\Excel\\RMA\\AU還回RMA完修率統計.xls";
                }

                string H1 = textBox1.Text.Substring(0, 6);
                string H2 = textBox2.Text.Substring(0, 6);
                System.Data.DataTable OrderData = GetOrderData4(GetSeqNo());
                System.Data.DataTable OrderData2 = GetOrderData43(GetSeqNo());
                System.Data.DataTable DT = GetC1(H1, H2, T1, GetSeqNo());
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                if (OrderData.Rows.Count > 0)
                {
                    //產生 Excel Report
                    ExcelReport.ODLNN2(OrderData, ExcelTemplate, OutPutFile, DT, OrderData2, T1);
                }
                else
                {
                    MessageBox.Show("沒有資料");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {

                string T1 = textBox1.Text.Substring(0, 4);

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\AU還回RMA完修率統計2.xls";

                string H1 = textBox1.Text.Substring(0, 6);
                string H2 = textBox2.Text.Substring(0, 6);
                System.Data.DataTable OrderData = GetOrderData4(GetSeqNo());
            
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ODLNN3(OrderData, ExcelTemplate, OutPutFile, T1);
                }
                else
                {
                    MessageBox.Show("沒有資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public  void ODLNN3(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string T1)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            object SelectCell = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(OrderData, sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }


                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            break;
                        }

                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();


                            FieldValue = "";
                            SetRow(OrderData, aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;

                        }


                        DetailRow++;


                    }

                }



                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                excelSheet.Activate();
                System.Data.DataTable TH = GetF(GetSeqNo());
                int L1 = 0;
                if (TH.Rows.Count > 0)
                {

                    for (int X = 0; X <= TH.Rows.Count - 1; X++)
                    {
                        string DOCDATE = TH.Rows[X][0].ToString();


                        System.Data.DataTable TH2 = GetF2(DOCDATE, GetSeqNo());

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[2, X + 2]);
                        range.Select();
                        range.Value2 = "'" + DOCDATE;

                        string J1 = TH2.Rows[0][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, X + 2]);
                        range.Select();
                        range.Value2 = J1;

                        string J2 = TH2.Rows[1][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[4, X + 2]);
                        range.Select();
                        range.Value2 = J2;

                        string J3 = TH2.Rows[2][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[5, X + 2]);
                        range.Select();
                        range.Value2 = J3;

                        string J4 = TH2.Rows[3][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[6, X + 2]);
                        range.Select();
                        range.Value2 = J4;

                        string J5 = TH2.Rows[4][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[7, X + 2]);
                        range.Select();
                        range.Value2 = J5;

                        string J6 = TH2.Rows[5][0].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[8, X + 2]);
                        range.Select();
                        range.Value2 = J6;

                        if (X != (TH.Rows.Count - 1))
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 3 + L1]);
                            range.Select();


                            range.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                          oMissing);

                            L1++;
                        }
                        else
                        {

                            System.Data.DataTable TH3 = GetF3(GetSeqNo());
                            string J7 = TH3.Rows[0][0].ToString();
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[3, 3 + L1]);
                            range.Select();
                            range.Value2 = J7;

                            string J8 = TH3.Rows[1][0].ToString();
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[4, 3 + L1]);
                            range.Select();
                            range.Value2 = J8;

                            string J9 = TH3.Rows[2][0].ToString();
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[5, 3 + L1]);
                            range.Select();
                            range.Value2 = J9;


                            string J10 = TH3.Rows[3][0].ToString();
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[6, 3 + L1]);
                            range.Select();
                            range.Value2 = J10;
                        }
                    }
                }

                excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            }
            finally
            {

                try
                {
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                string Msg = string.Empty;
                string Mo;

                System.Diagnostics.Process.Start(OutPutFile);

            }

        }
        public  string GetSeqNo()
        {

            ArrayList al = new ArrayList();

            for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
            {
                al.Add(listBox1.SelectedItems[i].ToString());
            }


            if (listBox1.SelectedItems.Count != 0)
            {
                StringBuilder sb2 = new StringBuilder();

                foreach (string v in al)
                {
                    sb2.Append("'" + v + "',");
                }

                sb2.Remove(sb2.Length - 1, 1);
                return sb2.ToString();
            }
            else
            {
                return "";
            }
        }
        public static void SetRow(System.Data.DataTable OrderData, int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[iRow][FieldName]);
            }

        }
        public static bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }
            //}
            return false;
        }
        public static bool CheckSerial(System.Data.DataTable OrderData, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }



        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            cENTERTextBox.Text = comboBox2.Text;
        }

        private void comboBox2_MouseClick_1(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU("RMACENTER");

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void rma_InvoiceFDataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "CELL")
                {
                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                    if (row != null)
                    {
                     
          
                        RmaF2 frm1 = new RmaF2();
                        frm1.q1 = Convert.ToString(row["RMANO"]).Trim();
                        frm1.q2 = dOCDATETextBox.Text;
                        frm1.Show();

                    }
                }

             
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = comboBox1.Text;
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetBU("RMACENTER");

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\AU還回RMA復判未結案明細.xls";


                System.Data.DataTable OrderData = GetNOTRETURN();

                if (OrderData.Rows.Count == 0)
                {

                    MessageBox.Show("沒有資料");
                    return;
                }

                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void rma_InvoiceFDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i2 = 0; i2 <= rma_InvoiceFDataGridView.Rows.Count - 2; i2++)
            {

                DataGridViewRow row;

                row = rma_InvoiceFDataGridView.Rows[i2];
                //COMPANY
                string COMPANY = row.Cells["COMPANY"].Value.ToString();
                string CONTRACTID = row.Cells["CONTRACTID"].Value.ToString();
                int 本次還貨 = Convert.ToInt16(row.Cells["QTY"].Value);
                DateTime ACME收貨日 = Convert.ToDateTime(GetMenu.DayS(dOCDATETextBox.Text));
                System.Data.DataTable Y1 = null;
                if (!String.IsNullOrEmpty(COMPANY))
                {
                    if (COMPANY == "進金生")
                    {
                        Y1 = GetY(CONTRACTID);
                    }
                    if (COMPANY == "達睿生")
                    {
                        Y1 = GetY5(CONTRACTID);
                    }
                    int n;
                    int H1 = 0;
                    if (Y1.Rows.Count > 0)
                    {

                        if (int.TryParse(Y1.Rows[0][0].ToString(), out n))
                        {
                            H1 = Convert.ToInt32(Y1.Rows[0][0]);
                        }
                    }

                    if (int.TryParse(row.Cells["QTY"].Value.ToString(), out n))
                    {
                        本次還貨 = 本次還貨 + H1;

                    }
                    if (COMPANY == "進金生")
                    {
                        UPDATEJOBNO(ACME收貨日, 本次還貨.ToString(), CONTRACTID);
                    }

                    if (COMPANY == "達睿生")
                    {
                        UPDATEJOBNO5(ACME收貨日, 本次還貨.ToString(), CONTRACTID);
                    }
                  
                }
              
            }
        }
        public void UPDATEJOBNO(DateTime U_ACME_BackDate, string U_ACME_QBACK, string CONTRACTID)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE acmesql02.DBO.OCTR SET U_ACME_BackDate=@U_ACME_BackDate,U_ACME_QBACK=@U_ACME_QBACK WHERE CONTRACTID =@CONTRACTID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate", U_ACME_BackDate));
            command.Parameters.Add(new SqlParameter("@U_ACME_QBACK", U_ACME_QBACK));
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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

        public void UPDATEJOBNO5(DateTime U_ACME_BackDate, string U_ACME_QBACK, string CONTRACTID)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE acmesql05.DBO.OCTR SET U_ACME_BackDate=@U_ACME_BackDate,U_ACME_QBACK=@U_ACME_QBACK WHERE CONTRACTID =@CONTRACTID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate", U_ACME_BackDate));
            command.Parameters.Add(new SqlParameter("@U_ACME_QBACK", U_ACME_QBACK));
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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

    }
}


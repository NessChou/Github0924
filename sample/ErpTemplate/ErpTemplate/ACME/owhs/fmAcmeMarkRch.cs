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

namespace ACME
{
    public partial class fmAcmeMarkRch : Form
    {

        private string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
        private string ShipConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
        //private string SapConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";

        string AutoFlag = "Y";

        public fmAcmeMarkRch()
        {
            InitializeComponent();
        }

        public System.Data.DataTable GetData(string Sql)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);


            SqlCommand command = new SqlCommand();
            command.Connection = connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(Sql);



            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }

        public System.Data.DataTable GetData(string ConnectiongString, string Sql)
        {
            SqlConnection connection = new SqlConnection(ConnectiongString);


            SqlCommand command = new SqlCommand();
            command.Connection = connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(Sql);



            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();

            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }

        public System.Data.DataTable GetACME_MAILLIST_DocType(string DocType)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            string sql = "SELECT UserCode,UserMail FROM ACME_ARES_MAIL where SysCode='{0}' and  Active='Y' ";
            sql = string.Format(sql, DocType);
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            //command.Parameters.Add(new SqlParameter("@DocType", DocType));
            //command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MAIL");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MAIL"];
        }

        private string GetTotalQty(string WhNo)
        {

            string Sql = @"SELECT  sum(Convert(int,Quantity)) Qty 
FROM WH_Item 
where ShippingCode='{0}' 
";

            Sql = string.Format(Sql, WhNo);
            System.Data.DataTable dt = GetData(Sql);
            if (dt.Rows.Count > 0)
            {
                return Convert.ToString(dt.Rows[0][0]);
            }
            else
            {
                return "";
            }
        }


        private void SendMail2(string FileName, string WhNo, string CardName)
        {
            DataRow dr;

            string strSubject;
            string UserCode;
            string UserMail;
            string MailContent = "";
            string MailDate = DateTime.Now.ToString("yyyyMMdd");

            string DocType = "ShipMark2";

            System.Data.DataTable dt = GetACME_MAILLIST_DocType(DocType);

            string Qty = GetTotalQty(WhNo);

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                //產生檔案

                //發送郵件
                //[工單] WH20191122041X--宜春-10PCS 嘜頭列印訊息通知
                //總數量
                //MAIL主旨請 SHOW “WH工單號 + 客戶名稱 + 數量 + “箱嘜””
                if (AutoFlag == "Y")
                {
                    strSubject = string.Format("{0} {1} {2}PCS 箱嘜", WhNo, CardName, Qty);
                }
                else
                {
                    strSubject = string.Format("[重送] {0} {1} {2}PCS 箱嘜", WhNo, CardName, Qty);
                }
                //if (cb.Checked)
                //{
                //    strSubject = string.Format("[工單重送] {0} 箱片序號檔案上傳結果訊息通知", WhNo);
                //}

                dr = dt.Rows[i];
                UserCode = Convert.ToString(dr["UserCode"]);

                //SetMsg("[郵寄] " + UserCode);

                UserMail = Convert.ToString(dt.Rows[i]["UserMail"]);

               // UserMail = "terrylee@acmepoint.com";

                if (string.IsNullOrEmpty(UserMail))
                {
                    UserMail = "terrylee@acmepoint.com";
                }

                MailContent = string.Format("客戶:{0} <br> 工單:{1}  嘜頭列印<br>", CardName, WhNo);

                //if (CheckMsg == "")
                //{

                //}
                //else
                //{
                //    MailContent += string.Format("異常訊息:{0} </br>", CheckMsg);
                //}

                MailTest(strSubject, UserCode, UserMail, MailContent, FileName);


                //Stage
                AddACME_MAIL_LOG(DocType, MailDate, UserCode, WhNo);

            }
        }

        public void AddACME_MAIL_LOG(string DocType, string MailDate, string UserCode, string Msg)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            SqlCommand command = new SqlCommand("Insert into ACME_MAIL_LOG(DocType,MailDate,UserCode,Msg) values(@DocType,@MailDate,@UserCode,@Msg)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocType", DocType));
            command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            command.Parameters.Add(new SqlParameter("@Msg", Msg));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent, string FileName)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));

            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\MailTemplates\\MarkMail.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;

            if (!string.IsNullOrEmpty(FileName))
            {
                message.Attachments.Add(new Attachment(FileName));
            }



            //message.Attachments.Add(new Attachment(Chart));

            //bettytseng@acmepoint.com
            //davidhuang@acmepoint.com
            //20191008
            if (SlpName.ToLower() == "sunny")
            {
                message.CC.Add(new MailAddress("bettytseng@acmepoint.com"));
                message.CC.Add(new MailAddress("davidhuang@acmepoint.com"));
            }


            SmtpClient client = new SmtpClient();
            client.Host = "smtp.acmepoint.com";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";
            //group-acmepoint@acmepoint.com
            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
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
                        // SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        //SetMsg(String.Format("Failed to deliver message to {0}",
                        //   ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                // SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                //  ex.ToString()));
            }

        }

        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }

        private void GridViewAutoSize(DataGridView dgv)
        {

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            // dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i <= dgv.Columns.Count - 1; i++)
            {
                int colw = dgv.Columns[i].Width;
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dgv.Columns[i].Width = colw;
            }
        }

        public bool ExcelRchColumn(System.Data.DataTable dt,
string Template, string SaveFileName, int Interval = 6, Int32 PageBreak = 4, string EndCell = "A5")
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;

            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Open(Template, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //第一個當作範本
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;



                Range rDetail = SheetTemplate.get_Range("A1", "D5") as Range;

                //依  dt 筆數產生分頁
                Int32 PrintQty = 1;
                DataRow dr;

                Microsoft.Office.Interop.Excel.Range d = null;
                Microsoft.Office.Interop.Excel.Range cell = null;




                Int32 CurrentLine = 1;
                Int32 CartonNo = 0;

                Int32 PrintCol = 2;
                string PrintColKey = "B";
                string PrintColValue = "C";

                string PO = "";
                string PN = "";
                //一筆一頁式
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];

                    CurrentLine = 1;
                    CartonNo = 0;

                    //    //複製範本
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];
                    excelSheet.Name = "Q" + (i + 1).ToString();
                    try
                    {
                        string s = Convert.ToString(dr["ItemCode"]) + "_"
                            + Convert.ToString(dr["Quantity"]);
                        s = s.Replace("AU", "").Replace(" ", "");
                        excelSheet.Name = s;
                    }
                    catch
                    {
                    }


                    PO = Convert.ToString(dr["PO"]);
                    PN = Convert.ToString(dr["PN"]);
                    PrintQty = Convert.ToInt32(dr["PrintQty"]);

                    //尾箱
                    string CartonFlag = "N";

                    if (Convert.ToString(dr["尾箱片數"]) != "0")
                    {
                        CartonFlag = "Y";
                    }

                    for (int q = 1; q <= PrintQty; q++)
                    {
                        PrintColKey = "A";
                        PrintColValue = "C";
                        if ((CartonNo + 1) % 2 == 0)
                        {
                            PrintColKey = "E";
                            PrintColValue = "G";
                        }

                        //最後一筆
                        //if (q == PrintQty - 1 && checkBox2.Checked == false)
                        if (q == PrintQty)
                        {
                            //是否有尾箱
                            if (CartonFlag == "Y")
                            {
                                //string cellx = "D" + (Interval * q + 12).ToString();
                                //SetCellValue(excelSheet, cellx, "零數");
                                rDetail.Copy();
                                d = excelSheet.get_Range(PrintColKey + (CurrentLine).ToString(), Type.Missing) as Range;
                                d.Select();
                                excelSheet.Paste(Type.Missing);

                                //Value---------------------------------------------------------------------------------
                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 1).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["PO"]);

                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 2).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["PN"]);

                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 3).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["尾箱片數"]);
                                //Value---------------------------------------------------------------------------------


                                if ((CartonNo + 1) % 2 == 0)
                                {
                                    CurrentLine = CurrentLine + Interval;
                                }

                                CartonNo = CartonNo + 1;
                            }
                            else
                            {
                                rDetail.Copy();
                                d = excelSheet.get_Range(PrintColKey + (CurrentLine).ToString(), Type.Missing) as Range;
                                d.Select();
                                excelSheet.Paste(Type.Missing);
                                //Value---------------------------------------------------------------------------------
                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 1).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["PO"]);

                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 2).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["PN"]);

                                cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 3).ToString(), Type.Missing) as Range;
                                cell.Value2 = Convert.ToString(dr["滿箱片數"]);
                                //Value---------------------------------------------------------------------------------



                                if ((CartonNo + 1) % 2 == 0)
                                {
                                    CurrentLine = CurrentLine + Interval;
                                }
                                CartonNo = CartonNo + 1;
                            }
                        }
                        else
                        {
                            rDetail.Copy();
                            d = excelSheet.get_Range(PrintColKey + (CurrentLine).ToString(), Type.Missing) as Range;
                            d.Select();
                            excelSheet.Paste(Type.Missing);

                            //Value---------------------------------------------------------------------------------
                            cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 1).ToString(), Type.Missing) as Range;
                            cell.Value2 = Convert.ToString(dr["PO"]);

                            cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 2).ToString(), Type.Missing) as Range;
                            cell.Value2 = Convert.ToString(dr["PN"]);

                            cell = excelSheet.get_Range(PrintColValue + (CurrentLine + 3).ToString(), Type.Missing) as Range;
                            cell.Value2 = Convert.ToString(dr["滿箱片數"]);
                            //Value---------------------------------------------------------------------------------


                            if ((CartonNo + 1) % 2 == 0)
                            {
                                CurrentLine = CurrentLine + Interval;
                            }
                            CartonNo = CartonNo + 1;

                        }

                        //sheet.HPageBreaks.Add(sheet.Range["A11"]);
                        if (CartonNo % (PageBreak * 2) == 0)
                        {
                            // CurrentLine = CurrentLine + 1;
                            //20200106 //簡易法 不同者加入分頁
                            // if (cbPallet.Checked == false)
                            //{
                            excelSheet.HPageBreaks.Add(excelSheet.Range["A" + (CurrentLine).ToString()]);
                            //}
                        }

                        //  CurrentLine = CurrentLine + Interval;
                    }




                    // MessageBox.Show(CartonNo.ToString());


                }//for 


                SheetTemplate.Delete();

                excelworkBook.SaveAs(SaveFileName); ;
                excelworkBook.Close();
                //excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (excelSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(SheetTemplate);


                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                SheetTemplate = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
            }
        }


        private void button30_Click(object sender, EventArgs e)
        {
            dgData.DataSource = null;

            //1854-00 
            //因為兩個 Column ,列印數 要 除以 2

            string Sql = @"declare @ShippingCode  nvarchar(30)
set @ShippingCode ='{0}'
SELECT    
substring(itemCode,1,11) ItemCode,
U_CUSTITEMCODE PN,
Quantity,
U_CUSTDOCENTRY PO,
--LPrint as PrintQty,
case when (LPrint='' or LPrint is null) then 1 else Convert(int,LPrint) end as PrintQty,
isnull(PQty1,'10') 滿箱片數,
isnull(PQty2,'0') 尾箱片數
FROM WH_Item 
where ShippingCode=@ShippingCode  
";
            //and U_CUSTDOCENTRY='PO19HK00100895'
            Sql = string.Format(Sql, txtRCH.Text);

            System.Data.DataTable dt = GetData(Sql);
            dgData.DataSource = dt;

            System.Data.DataTable dtCopy = dt.Clone();

            Int32 PrintQty = 0;
            DataRow dr;
            DataRow drNew;

            //if (cbQty.Checked)
            //{
            //    //滿箱
            //    for (int i = 0; i <= dt.Rows.Count - 1; i++)
            //    {
            //        dr = dt.Rows[i];
            //        dr.BeginEdit();
            //        try
            //        {
            //            //計算數字的整數部分
            //            PrintQty = Convert.ToInt32(Math.Truncate(
            //                       Convert.ToDouble(Convert.ToInt32(dr["Quantity"]) / Convert.ToInt32(dr["滿箱片數"]))));
            //            dr["Quantity"] = Convert.ToString(dr["滿箱片數"]);
            //            dr["PrintQty"] = PrintQty;
            //        }
            //        catch
            //        {
            //        }
            //        dr.EndEdit();

            //        CopyRow(dr, dtCopy);
            //    }



            //    Int32 RecCount = dt.Rows.Count;
            //    for (int i = 0; i <= RecCount - 1; i++)
            //    {
            //        dr = dt.Rows[i];


            //        if (Convert.ToString(dr["尾箱片數"]) != "0")
            //        {
            //            drNew = dt.NewRow();
            //            for (int j = 0; j <= dt.Columns.Count - 1; j++)
            //            {
            //                drNew[j] = dr[j];
            //            }

            //            drNew["Quantity"] = Convert.ToString(dr["尾箱片數"]);
            //            drNew["PrintQty"] = 1;

            //            dt.Rows.Add(drNew);
            //            //dt.Rows.InsertAt
            //            CopyRow(drNew, dtCopy);
            //        }
            //    }
            //    dgData.DataSource = dtCopy;
            //}
            //else
            //{
            //    dgData.DataSource = dt;
            //}


            //if (checkBox4.Checked)
            //{
            //    Int32 RecCount = dt.Rows.Count;
            //    for (int i = 0; i <= RecCount - 1; i++)
            //    {
            //        dr = dt.Rows[i];


            //        if (Convert.ToString(dr["尾箱片數"]) != "0")
            //        {



            //            drNew = dt.NewRow();
            //            for (int j = 0; j <= dt.Columns.Count - 1; j++)
            //            {
            //                drNew[j] = dr[j];
            //            }

            //            drNew["滿箱片數"] = Convert.ToString(dr["尾箱片數"]);
            //            drNew["尾箱片數"] = 0;
            //            drNew["Quantity"] = Convert.ToString(dr["尾箱片數"]);
            //            drNew["PrintQty"] = 1;

            //            dt.Rows.Add(drNew);
            //            //dt.Rows.InsertAt
            //            CopyRow(drNew, dtCopy);

            //            //20200113
            //            dr.BeginEdit();
            //            dr["尾箱片數"] = 0;
            //            dr["PrintQty"] = Convert.ToInt32(dr["PrintQty"]) - 1;
            //            dr.EndEdit();
            //        }
            //    }
            //}
            GridViewAutoSize(dgData);


            dgData.Columns["滿箱片數"].DefaultCellStyle.BackColor = Color.Yellow;
            dgData.Columns["尾箱片數"].DefaultCellStyle.BackColor = Color.Yellow;
        }

        private void button63_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtData = dgData.DataSource as System.Data.DataTable;

            //產生 QrCode 資料
            // MakeQrCodeValue(dtData);

            string Dir = GetExePath() + "\\Output\\";
            string DirTemplate = GetExePath() + "\\Excel\\Mark\\";

            if (!Directory.Exists(Dir))
            {
                Directory.CreateDirectory(Dir);
            }

            if (!Directory.Exists(DirTemplate))
            {
                Directory.CreateDirectory(DirTemplate);
            }



            string FileName = GetExePath() + "\\Output\\" + txtRCH.Text + ".xls";

            string Template = GetExePath() + "\\Excel\\Mark\\" + "RCHColumn.xls";

            ExcelRchColumn(dtData, Template, FileName);

           // if (Environment.UserName.ToLower() == "terrylee")
           // {
                System.Diagnostics.Process.Start(FileName);
            //}
        }

        private void button28_Click(object sender, EventArgs e)
        {
            string WhNo = txtRCH.Text;
            string FileName = GetExePath() + "\\Output\\" + WhNo + ".xls";

            SendMail2(FileName, WhNo, "RCH Asia Limited");
        }
    }
}

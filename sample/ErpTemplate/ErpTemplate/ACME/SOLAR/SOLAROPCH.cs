using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mime;
using System.Net.Mail;
using System.Security.Cryptography;
namespace ACME
{
    public partial class SOLAROPCH : ACME.fmBase1
    {
        System.Data.DataTable dtCost = null;
        decimal SIZE = 0;
        Attachment data = null;
        string P1 = "";
        string P2 = "";
        public SOLAROPCH()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sOLAR_OPCHTableAdapter.Connection = MyConnection;
            sOLAR_OPCH1TableAdapter.Connection = MyConnection;
            sOLAR_OPCH2TableAdapter.Connection = MyConnection;
        }
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;

            button2.Enabled = true;
            textBox2.Enabled = false;
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            dOCDATETextBox.ReadOnly = false;
            createNameTextBox.ReadOnly = false;
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                sOLAR.SOLAR_OPCH.RejectChanges();
                sOLAR.SOLAR_OPCH1.RejectChanges();
                sOLAR.SOLAR_OPCH2.RejectChanges();
            }
            catch
            {
            }
            return true;

        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override void EndEdit()
        {
            WW();

        }
        public override void STOP()
        {


        }
        public override void AfterEdit()
        {
            shippingCodeTextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
        }
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = sOLAR_OPCHBindingSource;
            MyTableName = "SOLAR_OPCH";
            MyIDFieldName = "ShippingCode";

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "BP" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            string username = fmLogin.LoginID.ToString();
            dOCDATETextBox.Text = GetMenu.Day();
            createNameTextBox.Text = username;
            this.sOLAR_OPCHBindingSource.EndEdit();
            kyes = null;
        }
        public override void FillData()
        {
            try
            {


                sOLAR_OPCHTableAdapter.Fill(sOLAR.SOLAR_OPCH, MyID);
                sOLAR_OPCH1TableAdapter.Fill(sOLAR.SOLAR_OPCH1, MyID);
                sOLAR_OPCH2TableAdapter.Fill(sOLAR.SOLAR_OPCH2, MyID);
                decimal M = 0;
                for (int i = 0; i <= sOLAR.SOLAR_OPCH2.Rows.Count - 1; i++)
                {
                    string PATH = sOLAR.SOLAR_OPCH2.Rows[i]["PATH"].ToString();

                    FileInfo filess = new FileInfo(PATH);
                    string size = filess.Length.ToString();
                    M += Convert.ToInt32(size);
                }
                M = M / 1000000;
                SIZE = M;
                textBox2.Text = M.ToString("#,##0.00") + "M";


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {


                Validate();

                sOLAR_OPCH1BindingSource.MoveFirst();

                for (int i = 0; i <= sOLAR_OPCH1BindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)sOLAR_OPCH1BindingSource.Current;

                    row3["NO"] = i;



                    sOLAR_OPCH1BindingSource.EndEdit();

                    sOLAR_OPCH1BindingSource.MoveNext();

                }


                sOLAR_OPCH2BindingSource.MoveFirst();

                for (int i = 0; i <= sOLAR_OPCH2BindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)sOLAR_OPCH2BindingSource.Current;

                    row3["NO"] = i;

                    sOLAR_OPCH2BindingSource.EndEdit();
                    sOLAR_OPCH2BindingSource.MoveNext();

                }

                sOLAR_OPCHTableAdapter.Connection.Open();

                sOLAR_OPCHBindingSource.EndEdit();
                sOLAR_OPCH1BindingSource.EndEdit();
                sOLAR_OPCH2BindingSource.EndEdit();

                tx = sOLAR_OPCHTableAdapter.Connection.BeginTransaction();

                SqlDataAdapter Adapter = util.GetAdapter(sOLAR_OPCHTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(sOLAR_OPCH1TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(sOLAR_OPCH2TableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;

                sOLAR_OPCHTableAdapter.Update(sOLAR.SOLAR_OPCH);
                sOLAR.SOLAR_OPCH.AcceptChanges();

                sOLAR_OPCH1TableAdapter.Update(sOLAR.SOLAR_OPCH1);
                sOLAR.SOLAR_OPCH1.AcceptChanges();

                sOLAR_OPCH2TableAdapter.Update(sOLAR.SOLAR_OPCH2);
                sOLAR.SOLAR_OPCH2.AcceptChanges();

                tx.Commit();

                this.MyID = this.shippingCodeTextBox.Text;

                UpdateData = true;
            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }


                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;

            }
            finally
            {
                this.sOLAR_OPCHTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("請輸入採購單");

                return;
            }
            object[] LookupValues = GetOPORList(textBox1.Text);

            if (LookupValues != null)
            {
                StringBuilder sb = new StringBuilder();


                for (int i = 0; i <= LookupValues.Length - 1; i++)
                {

                    sb.Append("'" + Convert.ToString(LookupValues[i]) + "',");

                }
                sb.Remove(sb.Length - 1, 1);
                string ds = sb.ToString();
                try
                {

                    System.Data.DataTable dt1 = GetOrderData(sb.ToString());
                    System.Data.DataTable dt4 = GetOPDN();
                    System.Data.DataTable dt2 = sOLAR.SOLAR_OPCH1;
                    System.Data.DataTable dt3 = sOLAR.SOLAR_OPCH2;
                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();

                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["NO"] = "0";
                        drw2["CARDNAME"] = drw["廠商"];
                        drw2["ITEMCODE"] = drw["料號"];
                        drw2["QTY"] = drw["數量"];
                        drw2["PAYAMT"] = drw["請款金額"];
                        drw2["AMT"] = drw["總金額"];
                        drw2["PRJCODE"] = drw["專案"];
                        drw2["DOCENTRY"] = drw["單號"];
                        drw2["PROJECT"] = drw["PROJECT"];
                        drw2["LINENUM"] = drw["LINE"];
                        dt2.Rows.Add(drw2);
                    }

                    if (dt4.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                        {
                            DataRow drw3 = dt3.NewRow();
                            DataRow drw4 = dt4.Rows[i];
                            string FILENAME = drw4["檔案名稱"].ToString();
                            drw3["ShippingCode"] = shippingCodeTextBox.Text;
                            drw3["DOCENTRY"] = textBox1.Text;
                            drw3["NO"] = "0";
                            drw3["PATH"] = drw4["path"].ToString() + "\\" + drw4["路徑"].ToString();
                            drw3["FILENAME"] = FILENAME;
                            if (!String.IsNullOrEmpty(FILENAME))
                            {
                                dt3.Rows.Add(drw3);
                            }

                        }
                    }

                    for (int j = 0; j <= sOLAR_OPCH1DataGridView.Rows.Count - 2; j++)
                    {
                        sOLAR_OPCH1DataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }

                    for (int j = 0; j <= sOLAR_OPCH2DataGridView.Rows.Count - 2; j++)
                    {
                        sOLAR_OPCH2DataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            textBox1.Text = "";
        }

        public void UPDATESAP()
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE SOLAR_OPCH1 SET OPCHECK='True' where SHIPPINGCODE=@SHIPPINGCODE ", connection);


            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));




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

        public void UPDATESAP2()
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;

            command = new SqlCommand("UPDATE SOLAR_OPCH2 SET OPCHECK='True' where SHIPPINGCODE=@SHIPPINGCODE ", connection);


            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));




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
        private System.Data.DataTable GetOPDN()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select t4.docentry 採購單號,t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,[FILENAME] 檔案名稱 from oclg t2");
            sb.Append("   LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append("  inner join OPOR t4 on(t2.docentry=t4.docentry)");
            sb.Append("  where  t2.doctype='22'  and t4.docentry=@docentry");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        private System.Data.DataTable GetMAILINFO()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT [NO] 編號,CARDNAME 廠商,ITEMCODE 料號,QTY 數量,AMT 總金額,PAYAMT  請款金額,PRJCODE 專案,ISNULL(MEMO,'') 備註,CAST(T0.DOCENTRY AS VARCHAR) 採購單號,'' 收貨採購單,'' 生產發貨單,LINENUM,PROJECT  FROM sOLAR_OPCH1  T0 ");
            sb.Append(" WHERE SHIPPINGCODE=@SHIPPINGCODE  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }

        private System.Data.DataTable GetMAILINFO2S(string PROJECTCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
     
            sb.Append("      SELECT DOCENTRY,U_PROJECTCODE PROJECTCODE  FROM OWOR WHERE  STATUS <> 'C' and U_PROJECTCODE=@PROJECTCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECTCODE", PROJECTCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        private System.Data.DataTable GetMAILINFO22(string DOCENTRY, string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL((CAST(PLANNEDQTY AS DECIMAL(10,4))),0),ISNULL((CAST(ISSUEDQTY AS DECIMAL(10,4))),0),LINENUM+1 LINENUM FROM WOR1  WHERE DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        private System.Data.DataTable GetMAILINFO2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DISTINCT COUNT(*) FROM sOLAR_OPCH1  WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            //if (!checkBox1.Checked)
            //{
            //    sb.Append("   AND OPCHECK <> 'True'  ");
            //}
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }

        private System.Data.DataTable GetMAILINFO3(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT  [FILENAME] FROM SOLAR_OPCH2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        private object[] GetOPORList(string aa)
        {

            string[] FieldNames = new string[] { "廠商", "專案", "料號", "數量", "請款金額", "總金額", "LINE" };

            string[] Captions = new string[] { "廠商", "專案", "料號", "數量", "請款金額", "總金額", "LINE" };

            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT T1.LINENUM LINE,T0.CARDNAME 廠商,T1.ITEMCODE 料號,CAST(T1.QUANTITY AS DECIMAL(10,2)) 數量,CAST(T1.LINETOTAL  AS INT) 請款金額,CAST(T0.DOCTOTAL-T0.VATSUM AS INT) 總金額");
            sb.Append("              ,T1.PROJECT+U_MEMO 專案 FROM OPOR T0 ");
            sb.Append("              LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("              LEFT JOIN OPRJ T2 ON (T1.PROJECT =T2.PRJCODE) ");
            sb.Append(" WHERE T0.DOCENTRY='" + aa + "' AND LINESTATUS='O' ");


            MultiValueDialog dialog = new MultiValueDialog();

            dialog.Captions = Captions;

            dialog.FieldNames = FieldNames;

            dialog.LookUpConnection = MyConnection;
            dialog.KeyFieldName = "LINE";
            dialog.SqlScript = sb.ToString();

            try
            {

                if (dialog.ShowDialog() == DialogResult.OK)
                {


                    object[] LookupValues = dialog.LookupValues;

                    return LookupValues;



                }

                else
                {

                    return null;

                }

            }

            finally
            {

                dialog.Dispose();

            }

        }
        private System.Data.DataTable GetOrderData(string Doc_no)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("              SELECT T1.LINENUM LINE,T0.CARDNAME 廠商,T1.ITEMCODE 料號,CAST(T1.QUANTITY AS DECIMAL(10,2)) 數量,CAST(T1.LINETOTAL  AS INT) 請款金額,CAST(T0.DOCTOTAL-T0.VATSUM AS INT) 總金額");
            sb.Append("              ,T1.PROJECT+ISNULL(U_MEMO,'') 專案,T0.DOCENTRY 單號,T1.PROJECT  FROM OPOR T0 ");
            sb.Append("              LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("              LEFT JOIN OPRJ T2 ON (T1.PROJECT =T2.PRJCODE) ");
            sb.Append("    WHERE T0.DOCENTRY=@DOCENTRY AND T1.LINENUM in (" + Doc_no + ")");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", textBox1.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "new01");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent)
        {
            MailMessage message = new MailMessage();
            message.From = new MailAddress("LleytonChen@acmepoint.com", "系統發送");
            message.To.Add(new MailAddress(MailAddress));
            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\SOLAR.htm";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##Content##", MailContent);
            template = template.Replace("##AA##", "請幫忙做系統收/發貨。 謝謝！");
            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;
            for (int i = 0; i <= sOLAR_OPCH2DataGridView.Rows.Count - 2; i++)
            {
                DataGridViewRow row = sOLAR_OPCH2DataGridView.Rows[i];

                string PATH = row.Cells["PATH"].Value.ToString().Trim();

                // string OPCHECK = row.Cells["OPCHECK2"].Value.ToString().Trim();

                //if (!checkBox1.Checked)
                //{
                //    if (OPCHECK != "True")
                //    {

                //        string m_File = "";

                //        m_File = PATH;

                //        data = new Attachment(m_File, MediaTypeNames.Application.Octet);
                //        ContentDisposition disposition = data.ContentDisposition;
                //        message.Attachments.Add(data);

                //    }
                //}
                //else
                //{
                string m_File = "";

                m_File = PATH;

                data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                ContentDisposition disposition = data.ContentDisposition;
                message.Attachments.Add(data);
                //  }

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
                data.Dispose();

            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        //SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        //SetMsg(String.Format("Failed to deliver message to {0}",
                        //    ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                //        ex.ToString()));
            }

        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("編號", typeof(string));
            dt.Columns.Add("廠商", typeof(string));
            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("計劃數量", typeof(string));
            dt.Columns.Add("已發貨", typeof(string));
            dt.Columns.Add("總金額", typeof(decimal));
            dt.Columns.Add("請款金額", typeof(decimal));
            dt.Columns.Add("專案", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("附件", typeof(string));
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("採購列號碼", typeof(string));
            dt.Columns.Add("生產單號", typeof(string));
            dt.Columns.Add("生產訂單列號碼", typeof(string));
            dt.Columns.Add("收貨採購單", typeof(string));
            dt.Columns.Add("生產發貨單", typeof(string));

            return dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            P1 = "";
            P2 = "";
            if (SIZE > 15)
            {
                MessageBox.Show("附件容量過大無法寄信");
                return;
            }

            TOTAL2();
            if (P2 == "1")
            {
                DialogResult result;
                result = MessageBox.Show("採購數量與生產訂單計劃數量不同，是否要寄出?", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.No)
                {
                    return;
                }
            }
            if (P1 == "")
            {
                System.Data.DataTable J2 = GetMAILINFO2();
                if (dtCost.Rows.Count > 0)
                {
                    string strSubject = "FW: [Solar] 請協助做系統進貨，共" + J2.Rows[0][0].ToString() + "筆";

                    string MailAddress = fmLogin.LoginID.ToString() + "@acmepointes.com";
                    string SlpName = fmLogin.LoginID.ToString();


                    dataGridView1.DataSource = dtCost;
                    string MailContent = htmlMessageBody(dataGridView1).ToString();
                    MailTest(strSubject, SlpName, MailAddress, MailContent);
                    //UPDATESAP();
                    //UPDATESAP2();
                    MessageBox.Show("寄信成功");
                    //sOLAR_OPCH1TableAdapter.Fill(sOLAR.SOLAR_OPCH1, MyID);
                    //sOLAR_OPCH2TableAdapter.Fill(sOLAR.SOLAR_OPCH2, MyID);
                }
                else
                {

                    MessageBox.Show("沒有資料");

                }
            }
        }
        private void TOTAL2()
        {
            dtCost = MakeTableCombine();

            System.Data.DataTable DT1 = GetMAILINFO();
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string ID = DT1.Rows[i]["編號"].ToString().Trim();
                string ITEM = DT1.Rows[i]["料號"].ToString().Trim();
                string 專案 = DT1.Rows[i]["PROJECT"].ToString().Trim();
                dr["編號"] = ID;
                dr["廠商"] = DT1.Rows[i]["廠商"].ToString().Trim();
                dr["料號"] = ITEM;
                dr["數量"] = DT1.Rows[i]["數量"].ToString().Trim();
                dr["總金額"] = DT1.Rows[i]["總金額"].ToString().Trim();
                dr["請款金額"] = DT1.Rows[i]["請款金額"].ToString().Trim();
                dr["專案"] = 專案;
                dr["備註"] = DT1.Rows[i]["備註"].ToString().Trim();
                string DOC = DT1.Rows[i]["採購單號"].ToString().Trim();
                dr["採購單號"] = DOC;
                dr["採購列號碼"] = DT1.Rows[i]["LINENUM"].ToString().Trim();
                //  
                StringBuilder sb = new StringBuilder();
                StringBuilder sb21 = new StringBuilder();
                StringBuilder sb22 = new StringBuilder();
                StringBuilder sb23 = new StringBuilder();
                StringBuilder sb24 = new StringBuilder();
                System.Data.DataTable DT = GetMAILINFO3(DOC);
                if (DT.Rows.Count > 0)
                {
                    for (int S = 0; S <= DT.Rows.Count - 1; S++)
                    {

                        DataRow dd = DT.Rows[S];
                        sb.Append(dd["FILENAME"].ToString() + "/");
                    }

                    sb.Remove(sb.Length - 1, 1);
                    dr["附件"] = sb.ToString();
                }
                string LINE = "";

                System.Data.DataTable DT2S = GetMAILINFO2S(專案);
                if (DT2S.Rows.Count > 0)
                {
                    for (int S2 = 0; S2 <= DT2S.Rows.Count - 1; S2++)
                    {
                        DataRow dd2 = DT2S.Rows[S2];
                        string 生產單號 = dd2["DOCENTRY"].ToString();
                        //            sb2.Append(dd2["FILENAME"].ToString() + "/");
                        System.Data.DataTable DT2 = GetMAILINFO22(生產單號, ITEM);

                        if (DT2.Rows.Count > 0)
                        {
                            if (DT2.Rows[0][0].ToString() != "0")
                            {
                                string 計劃數量 = DT2.Rows[0][0].ToString();
                                string 已發貨 = DT2.Rows[0][1].ToString();
                                LINE = DT2.Rows[0][2].ToString();
                                sb21.Append(計劃數量 + "/");
                                sb22.Append(已發貨 + "/");
                                sb23.Append(LINE + "/");
                                sb24.Append(生產單號 + "/");

                                string Q1 = DT1.Rows[i]["數量"].ToString().Trim();
                                string Q2 = DT2.Rows[0][1].ToString().Trim();
                                if (Q1 != Q2)
                                {
                                    P2 = "1";

                                }
                            }
                        }



                    }

                }
                if (sb21.Length != 0)
                {
                    sb21.Remove(sb21.Length - 1, 1);
                    dr["計劃數量"] = sb21.ToString();
                }
                if (sb22.Length != 0)
                {
                    sb22.Remove(sb22.Length - 1, 1);
                    dr["已發貨"] = sb22.ToString();
                }

                if (sb23.Length != 0)
                {
                    sb23.Remove(sb23.Length - 1, 1);
                    dr["生產訂單列號碼"] = sb23.ToString();
                }

                if (sb24.Length != 0)
                {
                    sb24.Remove(sb24.Length - 1, 1);
                    dr["生產單號"] = sb24.ToString();
                }

                if (sb24.Length != 0 && sb23.Length == 0)
                {
                    P1 = "1";
                    MessageBox.Show("生產單號 :" + sb24.ToString() + " 料號 :" + ITEM + " 生產訂單列號碼為空白，無法寄信");
                    return;
                }




                dr["收貨採購單"] = "";
                dr["生產發貨單"] = "";
                dtCost.Rows.Add(dr);
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






                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;
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

        private void SOLAROPCH_Load(object sender, EventArgs e)
        {
            WW();
        }

    }
}

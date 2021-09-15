using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Security.Cryptography;
using System.IO;
using System.Net.Mime;
using CarlosAg.ExcelXmlWriter;
namespace ACME
{
    public partial class RmaCarton : Form
    {
        Attachment data = null;
        public RmaCarton()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.rMA_CARTONTableAdapter.Fill(this.rm.RMA_CARTON, new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox1.Text, typeof(int))))), new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox2.Text, typeof(int))))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void rMA_CARTONDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["DOCDATE"].Value = rMA_CARTONDataGridView.Rows.Count.ToString();
            e.Row.Cells["DOCYEAR"].Value = comboBox1.Text;
            e.Row.Cells["DOCMONTH"].Value = comboBox2.Text;
         
        }

        private void RmaCarton_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Month(), "DataValue", "DataValue");

            comboBox1.Text = DateTime.Now.ToString("yyyy");
            comboBox2.Text = Convert.ToString(Convert.ToInt16(DateTime.Now.ToString("MM")));
            textBox1.Text = GetMenu.Day();
            try
            {
                this.rMA_CARTONTableAdapter.Fill(this.rm.RMA_CARTON, new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox1.Text, typeof(int))))), new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox2.Text, typeof(int))))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rMA_CARTONBindingSource.EndEdit();
            this.rMA_CARTONTableAdapter.Update(this.rm.RMA_CARTON);

            MessageBox.Show("¦sÀÉ¦¨¥\");
        }

        private void button3_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = GetYear();
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        public System.Data.DataTable GetMonth()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select cast(DOCYEAR as varchar)+'/'+cast(DOCMONTH as varchar)+'/'+cast(DOCDATE as varchar) ¤é´Á ,AUIN AU¦¬³f,AUOUT AU©ñ³f,CUSTIN «È¤á¦¬³f,CUSTOUT «È¤á©ñ³f from rMA_CARTON where docyear=@docyear and docmonth=@docmonth ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docyear", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@docmonth", comboBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetMonth2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("            select T0.SHIPPINGCODE JOBNO,T0.CARDNAME «È¤á¦WºÙ,SHIPPING_OBU ­Ü§O,T1.SeqNo §Ç¸¹,T1.Docentry ³æ¸¹");
            sb.Append("            ,ShipDate ±Æµ{¤é´Á,ItemRemark ³æ¾ÚÁ`Ãþ,ItemCode ²£«~½s¸¹,Dscription «~¦W³W®æ,T1.PiNo ®Æ¸¹,T1.Quantity ¼Æ¶q");
            sb.Append("            ,T1.Grade µ¥¯Å,T1.Ver ª©¥»,INV ­ì¼tINVOCE¤é´Á,Invoice ­ì¼tINVOCE FROM WH_MAIN T0");
            sb.Append("            LEFT JOIN  WH_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append("            WHERE  SUBSTRING(T0.SHIPPINGCODE,3,8) IN (select Convert(varchar(10),CAST(cast(DOCYEAR as varchar)+'/'+cast(DOCMONTH as varchar)+'/'+cast(DOCDATE as varchar) AS DATETIME),112) «È¤á©ñ³f from rMA_CARTON ");
            sb.Append(" WHERE (AUIN >9 OR AUOUT > 9 OR CUSTIN > 9 OR CUSTOUT > 9) AND docyear=@docyear and docmonth=@docmonth) ORDER BY T0.SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docyear", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@docmonth", comboBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public System.Data.DataTable GetYear()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select cast(DOCYEAR as varchar)+'/'+cast(DOCMONTH as varchar)+'/'+cast(DOCDATE as varchar) ¤é´Á ,AUIN AU¦¬³f,AUOUT AU©ñ³f,CUSTIN «È¤á¦¬³f,CUSTOUT «È¤á©ñ³f from rMA_CARTON where docyear=@docyear  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docyear", comboBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        public  System.Data.DataTable download3()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT RANK() OVER (ORDER BY T0.U_RMA_NO DESC) AS §Ç¸¹,T0.U_RMA_NO RMANO,T0.U_CUSNAME_S «È¤áÂ²ºÙ,T0.U_RMODEL MODEL,T0.U_RVER VER,T0.U_RQUINITY QTY");
            sb.Append(" ,''''+T0.U_AUO_RMA_NO VENDER,Convert(varchar(10),CASE WHEN ISNULL(T0.U_VENREDATE,'')='' THEN T0.U_RTORECEIVING");
            sb.Append(" ELSE T0.U_VENREDATE END,111) VENDERÁÙ³f¤é,CASE WHEN ISNULL(T0.U_VENREDATE,'')='' THEN T0.U_RQUINITY");
            sb.Append(" ELSE T0.U_VENREQTY  END VENDER¤wÁÙ¼Æ¶q FROM OCTR T0");
            sb.Append(" LEFT JOIN CTR1 T1 ON (T0.CONTRACTID=T1.CONTRACTID)");
            sb.Append(" WHERE Convert(varchar(10),CASE WHEN ISNULL(T0.U_VENREDATE,'')='' THEN T0.U_RTORECEIVING");
            sb.Append(" ELSE T0.U_VENREDATE END,112)=@AA");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }

        public System.Data.DataTable download4()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.U_RMA_NO RMANO,T0.U_CUSNAME_S «È¤áÂ²ºÙ,T0.U_RMODEL MODEL,T0.U_RVER VER,T0.U_RQUINITY QTY");
            sb.Append(" ,Convert(varchar(10),T0.U_RTORECEIVING,111)  ACME¦¬³f¤é FROM OCTR T0");
            sb.Append(" LEFT JOIN CTR1 T1 ON (T0.CONTRACTID=T1.CONTRACTID)");
            sb.Append(" WHERE Convert(varchar(10),T0.U_RTORECEIVING,112)=@AA");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private void button4_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = GetMonth();


            dataGridView2.DataSource = GetMonth2();
           


            CarlosAg.ExcelXmlWriter.Workbook book = new CarlosAg.ExcelXmlWriter.Workbook();
            WorksheetStyle headerStyle = book.Styles.Add("headerStyleID");
            headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            headerStyle.Alignment.WrapText = true;
            headerStyle.Interior.Color = "#284775";
            headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;
            headerStyle.Font.Color = "white";
            headerStyle.Font.Bold = true;

            WorksheetStyle defaultStyle = book.Styles.Add("workbookStyleID");
            defaultStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            defaultStyle.Alignment.WrapText = true;
            defaultStyle.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");
            WH(book, dataGridView1, "²Î­p");

            WH(book, dataGridView2, "¦¬¥X¶W¹L10½c");

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
    DateTime.Now.ToString("yyyyMMddHHmmss") + "½c¼ÆºÞ²z.xls";
            book.Save(OutPutFile);
            System.Diagnostics.Process.Start(OutPutFile);
        }

        private void button5_Click(object sender, EventArgs e)
        {

            try
            {
                string SlpName = "LleytonChen";
                string MailAddress = "LleytonChen@acmepoint.com";
                DELETEFILE();
                System.Data.DataTable dt = GetACME_MAILLIST("RMA_TO");
                DataRow dr;
             
                        System.Data.DataTable dt7 = download4();
                        if (dt7.Rows.Count > 0)
                        {

                            string FileName = string.Empty;
                            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                            FileName = lsAppDir + "\\Excel\\RMA\\¦¬³f«È¤á.xls";

                            string ExcelTemplate = FileName;
                            string OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("MMdd") + "¦¬³f«È¤á" + ".xls";
                            ExcelReport.ExcelReportTONY(dt7, ExcelTemplate, OutPutFile, "N");


                        }

                        System.Data.DataTable dt6 = download3();
                        if (dt6.Rows.Count > 0)
                        {

                            string FileName = string.Empty;
                            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                            FileName = lsAppDir + "\\Excel\\RMA\\AUÁÙ¦^.xls";

                            string ExcelTemplate = FileName;
                            string OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("MMdd") + "AUÁÙ¦^" + ".xls";
                            ExcelReport.ExcelReportTONY(dt6, ExcelTemplate, OutPutFile, "N");


                        }
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {
                            if (dt.Rows.Count > 0)
                            {
                                dr = dt.Rows[i];


                                SlpName = Convert.ToString(dr["UserCode"]);





                                MailAddress = Convert.ToString(dt.Rows[i]["UserMail"]);

                                MailTest("¨C¤é¦¬³f", SlpName, MailAddress, "");
                            }
                        }
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                MessageBox.Show("±H«H¦¨¥\");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }

        }
        public System.Data.DataTable GetACME_MAILLIST(string SysCode)
        {
            SqlConnection connection = globals.Connection;
            // string sql = "SELECT BuGroup,UserCode,UserData FROM ACME_MAILLIST ";
            //string sql = "SELECT UserCode,UserPermit FROM ACME_FEED_SEC where MailEnable='Y' ";
            string sql = "SELECT UserCode,UserMail FROM ACME_ARES_MAIL where Active='Y' and SysCode=@SysCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SysCode", SysCode));
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
        private void MailTest(string strSubject, string SlpName, string MailAddress, string MailContent)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress("workflow@acmepoint.com", "¨t²Îµo°e");
            message.To.Add(new MailAddress(MailAddress));



            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\MailTemplates\\RMATO.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "¶iª÷¥Í¹ê·~");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>½Ð°Ñ¦Ò!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //®æ¦¡¬° Html
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                string m_File = "";

                m_File = file;
                data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                //ªþ¥ó†V®Æ
                ContentDisposition disposition = data.ContentDisposition;


                // ¥[¤J…o¥óªþ¥ó
                message.Attachments.Add(data);

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
                 
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                       
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            
            }

        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void DELETEFILE()
        {
            string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string FileName1 = lsAppDir1 + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(FileName1);
            foreach (string file in filenames)
            {


                File.Delete(file);

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void WH(CarlosAg.ExcelXmlWriter.Workbook book, DataGridView DGV, string DD)
        {



            CarlosAg.ExcelXmlWriter.Worksheet sheet = book.Worksheets.Add(DD);
            WorksheetRow headerRow = sheet.Table.Rows.Add();
            for (int i = 0; i < DGV.Columns.Count; i++)
            {
                headerRow.Cells.Add(DGV.Columns[i].HeaderText, DataType.String, "headerStyleID");
            }

            for (int i = 0; i < DGV.Rows.Count - 1; i++)
            {

                DataGridViewRow row = DGV.Rows[i];
                WorksheetRow rowS = sheet.Table.Rows.Add();

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    //if (j == 0 || j == 1)
                    //{
                    rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                    // }
                    //else
                    //{
                    //    rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                    //}
                    rowS.AutoFitHeight = true;
                    rowS.Table.DefaultColumnWidth = 100;

                }

            }
        }
    }
}
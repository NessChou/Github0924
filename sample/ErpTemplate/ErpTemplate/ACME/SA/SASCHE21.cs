using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net.Mail;
using System.Web.UI;
namespace ACME
{
    public partial class SASCHE21 : Form
    {
        string LOGIN = fmLogin.LoginID.ToString().ToUpper();
        string COMPANY = "禾中";
        string mail = "";
        public SASCHE21()
        {
            InitializeComponent();
        }

        private void sASCHE2BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sASCHE2BindingSource.EndEdit();
            this.sASCHE2TableAdapter.Update(this.sa.SASCHE2);

        }

        private void button1_Click(object sender, EventArgs e)
        {
                   
            System.Data.DataTable dt1 = Get1();

            System.Data.DataTable dt2 = sa.SASCHE2;
            if (dt1.Rows.Count > 0)
            {
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    drw2["DOCENTRY"] = drw["DOCENTRY"];
                    drw2["LINENUM"] = drw["LINENUM"];
                    drw2["CHINO"] = drw["CHINO"];
                    drw2["CARDNAME"] = drw["CARDNAME"];
                    drw2["FCARDNAME"] = drw["FCARDNAME"];
                    drw2["PO"] = drw["PO"];
                    drw2["ITEMCODE"] = drw["ITEMCODE"];
                    drw2["QTY"] = drw["QTY"];
                    drw2["WHNAME"] = drw["WHNAME"];
                    drw2["ORDDAY"] = drw["ORDDAY"];
                    drw2["WORKDAY"] = drw["WORKDAY"];
                    drw2["LEAVEDAY"] = drw["LEAVEDAY"];
                    drw2["SCHEDAY"] = drw["SCHEDAY"];
                    drw2["PAY"] = drw["PAY"];

                    drw2["SHIPDAY"] = drw["SHIPDAY"];
                    drw2["STATUS"] = drw["STATUS"];
                    drw2["MARK"] = drw["MARK"];
                    drw2["MEMO"] = drw["MEMO"];
                    drw2["TERM"] = drw["TERM"];
                    drw2["SA"] = drw["SA"];
                    drw2["SALES"] = drw["SALES"];
                    drw2["LOGIN"] = LOGIN;
                    drw2["COMPANY"] = COMPANY;

                    dt2.Rows.Add(drw2);
                }

                this.Validate();
                this.sASCHE2BindingSource.EndEdit();
                this.sASCHE2TableAdapter.Update(this.sa.SASCHE2);

                System.Data.DataTable G1 = GetVCHINO();
                if (G1.Rows.Count > 0)
                {
                    sASCHE2DataGridView.Columns["CHINO"].Visible = true;
                }
                else
                {
                    sASCHE2DataGridView.Columns["CHINO"].Visible = false;
                }

                System.Data.DataTable G2 = GetVFCARDNAME();
                if (G2.Rows.Count > 0)
                {
                    sASCHE2DataGridView.Columns["FCARDNAME"].Visible = true;
                }
                else
                {
                    sASCHE2DataGridView.Columns["FCARDNAME"].Visible = false;
                }
                System.Data.DataTable G3 = GetVPO();
                if (G3.Rows.Count > 0)
                {
                    sASCHE2DataGridView.Columns["PO"].Visible = true;
                }
                else
                {
                    sASCHE2DataGridView.Columns["PO"].Visible = false;
                }
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        public void TRUNG()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE SASCHE2 WHERE [LOGIN] =@LOGIN AND COMPANY=@COMPANY ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
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
        private System.Data.DataTable Get1()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.LINENUM ,T0.DOCENTRY ,T3.U_CHI_NO CHINO,T3.U_Beneficiary FCARDNAME,  t3.numatcard PO,T0.ITEMCODE,t0.dscription ITEMNAME,cast(T0.quantity as int) QTY,  ");
            sb.Append(" case when t3.cardname like  '%TOP GARDEN INT%' then 'TOP GARDEN' when t3.cardname like  '%CHOICE CHANNEL%' then 'CHOICE' when t3.cardname like  '%Infinite Power Group%' then 'INFINITE' when t3.cardname like  '%宇豐光電股份有限公司%' then '宇豐'  when t3.cardname like  '%達睿生%' then 'DRS' else t3.cardname end+CASE ISNULL(T3.U_BENEFICIARY,'') WHEN '' THEN '' ELSE '-'+T3.U_BENEFICIARY END CARDNAME,  ");
            sb.Append(" t7.WHSNAME WHNAME,T0.u_acme_workday+'('+CAST(T2.day AS VARCHAR)+')' WORKDAY,      Convert(varchar(8),T0.ShipDate ,112)  ORDDAY,  ");
            sb.Append(" Convert(varchar(8),T0.u_acme_shipday,112)  LEAVEDAY, Convert(varchar(8),T0.u_acme_work,112)  SCHEDAY,U_PAY PAY,U_SHIPDAY SHIPDAY  ");
            sb.Append(" ,U_SHIPSTATUS [STATUS],U_MARK MARK,U_MEMO MEMO,");
            sb.Append(" T3.u_acme_tardeTERM  TERM ,(T4.[SlpName]) SALES,(T5.[lastName]+T5.[firstName]) SA            ");
            sb.Append(" FROM acmesql02.dbo.rdr1 T0      ");
            sb.Append(" left join  acmesqlsp.dbo.WorkDay T2 on (T2.workday=T0.u_acme_workday ) ");
            sb.Append(" left join  acmesql02.dbo.ORDR T3 on (T0.DOCENTRY=T3.DOCENTRY )     ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSLP T4 ON T3.SlpCode = T4.SlpCode    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OHEM T5 ON T3.OwnerCode = T5.empID    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.owhs T7 ON T0.whscode=T7.whscode    ");
            sb.Append(" where 1=1 and t3.canceled <> 'Y' AND T3.doctype='I' AND T3.CARDCODE='1858-00' ");
            sb.Append("and  T0.DOCENTRY =  '" + textBox3.Text.ToString().Trim().Replace(" ", "") + "'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetVCHINO()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CHINO  FROM SASCHE2 WHERE  ISNULL(CHINO,'') <> '' AND [LOGIN] =@LOGIN AND COMPANY=@COMPANY  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetVFCARDNAME()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT FCARDNAME  FROM SASCHE2 WHERE  ISNULL(FCARDNAME,'') <> '' AND [LOGIN] =@LOGIN AND COMPANY=@COMPANY ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetVPO()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PO  FROM SASCHE2 WHERE  ISNULL(PO,'') <> '' AND [LOGIN] =@LOGIN AND COMPANY=@COMPANY  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void SASCHE2_Load(object sender, EventArgs e)
        {

            TRUNG();

            System.Data.DataTable T1 = GetMenu.GetWHSAALL();
            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sASCHE2BindingSource.EndEdit();
            this.sASCHE2TableAdapter.Update(this.sa.SASCHE2);

            System.Data.DataTable G2 = Get2();

            if (G2.Rows.Count > 0)
            {
                dataGridView1.DataSource = G2;


                StringBuilder ss = new StringBuilder();
                if (listBox1.SelectedItems.Count != 0)
                {


                    ArrayList al = new ArrayList();
                    for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                    {
                        string f = listBox1.SelectedItems[i].ToString().Replace("(GT)", "");
                        al.Add(f);
                    }



                    foreach (string v in al)
                    {
                        if (v == "AppleChen" || v == "ViviWeng")
                        {
                            ss.Append("" + v + "@getogether.com.hk;");
                        }
                        else
                        {
                            ss.Append("" + v + "@acmepoint.com;");
                        }
                    }
                }

                if (checkBox3.Checked)
                {

                    ArrayList al = new ArrayList();
                    for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                    {
                        string f = listBox1.Items[i].ToString().Replace("(GT)", "");
                        al.Add(f);
                    }

                    foreach (string v in al)
                    {
                        if (v == "AppleChen" || v == "ViviWeng")
                        {
                            ss.Append("" + v + "@getogether.com.hk;");
                        }
                        else
                        {
                            ss.Append("" + v + "@acmepoint.com;");
                        }
                    }
                }
                if (checkBox1.Checked)
                {
                    System.Data.DataTable SHIPSTOCCK = GetMenu.GetSAALL();
                    if (SHIPSTOCCK.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                        {
                            DataRow dd = SHIPSTOCCK.Rows[i];
                            ss.Append(dd["EMAIL"].ToString() + ";");
                        }
                    }
                }
     

                if (ss.Length > 5)
                {
                    ss.Remove(ss.Length - 1, 1);
                    mail = ss.ToString();
                    if (globals.GroupID.ToString().Trim() == "EEP")
                    {
                        mail = "lleytonchen@acmepoint.com";
                    }
                }
                else
                {
                    MessageBox.Show("請選擇收件者");
                    return;
                }


                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\SA.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);


                string Html = htmlMessageBody(dataGridView1).ToString();
                template = template.Replace("##Content##", Html);


                MailMessage message = new MailMessage();
                string[] arrurl = mail.Split(new Char[] { ';' });

                foreach (string i in arrurl)
                {

                    message.To.Add(i);

                }
                string SUB = "";
                string CARDNAME = G2.Rows[0]["客戶"].ToString();


                System.Data.DataTable G3 = Get3();
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i <= G3.Rows.Count - 1; i++)
                {

                    DataRow dd = G3.Rows[i];


                    sb.Append(dd["銷售訂單"].ToString() + "/");


                }

                sb.Remove(sb.Length - 1, 1);

                SUB = "請在禾中打單-" + CARDNAME + "-SO#" + sb.ToString();

                message.Subject = SUB;
                message.Body = template;

                //格式為 Html
                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                try
                {
                 
                        client.Send(message);
             

                    MessageBox.Show("寄信成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
        }
        private System.Data.DataTable Get2()
        {
            string V1 = "";
            string V2 = "";
            string V3 = "";
            System.Data.DataTable G1 = GetVCHINO();
            if (G1.Rows.Count > 0)
            {
                V1 = "Y";
            }

            System.Data.DataTable G2 = GetVFCARDNAME();
            if (G2.Rows.Count > 0)
            {
                V2 = "Y";
            }

            System.Data.DataTable G3 = GetVPO();
            if (G3.Rows.Count > 0)
            {
                V3 = "Y";
            }
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCENTRY 銷售訂單,");
            if (V1 == "Y")
            {
                sb.Append("CHINO 正航單號,");
            }
            sb.Append(" CARDNAME 客戶,");
            if (V2 == "Y")
            {
                sb.Append("FCARDNAME 最終客戶,");
            }
            if (V3 == "Y")
            {
                sb.Append("PO PO號碼,");
            }
            sb.Append(" ITEMCODE 料號,QTY 數量,CURRENCY 幣別,AMT 金額,WARRANTY 保固,WHNAME 倉庫,ORDDAY 訂單交期");
            sb.Append(" ,WORKDAY 工作天數,LEAVEDAY 離倉日期 ,SCHEDAY 排程日期,TERM,SALES 業務,SA 業助,PAY 付款,SHIPDAY 押出貨日 ");
            sb.Append(" ,[STATUS] 貨況,MARK 特殊嘜頭, MEMO 注意事項  FROM SASCHE2  WHERE [LOGIN] =@LOGIN AND COMPANY=@COMPANY  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();



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
            for (int i = 0; i <= dg.Rows.Count - 2; i++)
            {

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

                                strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                            }

                        }

                    
                }
                strB.AppendLine("</tr>");

            }

            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }
        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT DOCENTRY 銷售訂單 FROM SASCHE2  WHERE [LOGIN] =@LOGIN AND COMPANY=@COMPANY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        //private void fillToolStripButton_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        this.sASCHE2TableAdapter.Fill(this.sa.SASCHE2, lOGINToolStripTextBox.Text);
        //    }
        //    catch (System.Exception ex)
        //    {
        //        System.Windows.Forms.MessageBox.Show(ex.Message);
        //    }

        //}
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Net.Mail;
namespace ACME
{
    public partial class ODLN : Form
    {
        public ODLN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getprepareend();
        }

        private System.Data.DataTable Getprepareend()
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT distinct T0.DOC ID,T0.DOCENTRY SAP單號,T0.CARDNAME 客戶名稱,CASE T1.FLOWFLAG WHEN 'N' THEN '審核中' WHEN 'P' THEN '被退回' WHEN 'Z' THEN '已結案' WHEN '' THEN '未送出' ELSE FLOWFLAG END 簽核狀態");
            sb.Append("               , T2.SLPNAME 業務,CONVERT(varchar(12),T0.DOCDATE,111) 時間 FROM  acmesqleep.dbo.ACME_AUTOODLN T0");
            sb.Append("              LEFT JOIN acmesqleep.dbo.ACME_FLOW_DATA T1 ON (T0.DOC=SUBSTRING(T1.DOCENTRY,3,LEN(T1.DOCENTRY))) ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSLP T2 ON (T1.SALES=T2.MEMO COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where 1=1 ");
            if (textBox5.Text != "")
            {
                sb.Append(" AND T0.DOCENTRY in (" + textBox5.Text + ")");
            }
            else
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" AND t0.docdate between @AA and @BB ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {
                    sb.Append(" AND T0.DOCENTRY between @CC and @DD ");
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DD", textBox4.Text));

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

        private void ODLN_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();

            dataGridView1.DataSource = Getprepareend();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                for (int i = 0; i <= dataGridView1.SelectedRows.Count - 1; i++)
                {
                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[i];
                    string T1 = row.Cells["SAP單號"].Value.ToString().Trim();
                    string T2 = row.Cells["ID"].Value.ToString().Trim();
                    string T3 = row.Cells["簽核狀態"].Value.ToString().Trim();
                    string T4 = row.Cells["時間"].Value.ToString().Trim();
                    //if (T3 != "已結案")
                    //{
                    //    MessageBox.Show("未審核無法匯出");
                    //    return;

                    //}
                    if (!String.IsNullOrEmpty(T1))
                    {
                        System.Data.DataTable G2 = GetODLN1(T2);
                        string TIME = "";
                        if (G2.Rows.Count > 0)
                        {
                            TIME = G2.Rows[0]["時間"].ToString();
                        }
                        else
                        {
                            TIME = T4;
                        }

                        System.Data.DataTable G1 = GetODLN(T1, TIME);
                        string SALESS = G1.Rows[0]["銷售人員"].ToString();

                        string USER = G1.Rows[0]["USERNAME"].ToString().ToUpper();
                        string FileName = string.Empty;

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


                        FileName = lsAppDir + "\\Excel\\wh\\銷貨單.xls";
                        string SALES = G1.Rows[0]["SALES"].ToString();

                        //Excel的樣版檔

                        System.Data.DataTable G3 = Getgroup(SALES);

                        if (G3.Rows.Count > 0)
                        {
                            SALES = "業務-許心如";
                        }
                        //\\acmew08r2ap\table\SIGN\SALES
                        string B2 = "//acmew08r2ap//table//SIGN//SALES//";
                        string B3 = "//acmew08r2ap//table//SIGN//USER//";

                        string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                              DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                        try
                        {
                            ExcelReport.ExcelReportOutputODLN(G1, FileName, OutPutFile, "N", B2 + SALES + ".JPG", B3 + USER + ".JPG", B3 + "APPLECHEN.JPG");
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }

       

  

        }

        private System.Data.DataTable GetODLN(string DOCENTRY,string H1)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                                 SELECT Convert(varchar(11),T0.TAXDATE,111)  銷貨日期,Convert(varchar(11),T0.TAXDATE,111) 印表日期");
            sb.Append("                                    ,T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,T0.ADDRESS 帳單地址,T0.ADDRESS2 送貨地址");
            sb.Append("                                 ,T1.ITEMCODE 產品編號,T1.DSCRIPTION 品名規格,T1.QUANTITY 數量,T1.PRICE 單價,T1.QUANTITY*T1.PRICE 金額");
            sb.Append("                                 ,T2.zipcode 聯絡人,T2.block 連絡電話,T2.CITY 傳真號碼,SALUNITMSR 單位,T1.CURRENCY 幣別編號,T0.DOCENTRY 單號,T0.vatsum 稅額, T0.doctotal 總金額,T0.doctotal-T0.vatsum 未稅");
            sb.Append("                                 ,T0.U_ACME_DOC_RATE 匯率,T4.SLPNAME 銷售人員,T0.U_ACME_USER  製單人員,T5.LICTRADNUM 統一編號,T6.WHSNAME 倉庫名稱,CASE ISNULL(NUMATCARD,'')  WHEN '' THEN '' ELSE 'PO#'+NUMATCARD END+' '+U_ACME_PAYGUI 備註,T7.GROUPNAME SALES,Convert(varchar(11),T0.CREATEDATE,111) 製單時間,T8.HOMETEL USERNAME,T1.BASEENTRY 來源單號,@H1 AS 核准時間 FROM ODLN T0 ");
            sb.Append("                                 LEFT JOIN DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("                                 LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.shiptocode=T2.ADDRESS and T2.adrestype='S')");
            sb.Append("                                 LEFT JOIN OITM T3 on (T1.ITEMCODE=T3.ITEMCODE)");
            sb.Append("                                 LEFT JOIN OSLP T4 on (T0.SLPCODE=T4.SLPCODE)");
            sb.Append("                                 LEFT JOIN OCRD T5 on (T0.CARDCODE=T5.CARDCODE)");
            sb.Append("                                 LEFT JOIN OWHS T6 on (T1.WHSCODE=T6.WHSCODE)");
            sb.Append("  LEFT JOIN AcmeSqlEEP.dbo.GROUPS T7 ON (T4.MEMO=T7.GROUPID COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("                     LEFT JOIN OHEM T8 on (T0.U_ACME_USER=T8.LASTNAME+FIRSTNAME)");
            sb.Append("                       WHERE T0.DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@H1", H1));


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

        private System.Data.DataTable GetODLN1(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CONVERT(varchar(12), CAST(UPDATE_DATE AS DATETIME),111) 時間 from AcmeSqlEEP.dbo.SYS_TODOHIS");
            sb.Append("  where flow_desc='銷貨單流程(TFT)' AND S_STEP_ID='業務審核' AND ISNULL(form_presentation,'') <> '' AND REPLACE(substring(form_presentation,15,DataLength(form_presentation)-15),'''','')=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));



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
        private System.Data.DataTable Getgroup(string GROUPNAME)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT GROUPID FROM AcmeSqlEEP.dbo.GROUPS WHERE DESCRIPTION LIKE '%中國%' AND GROUPNAME=@GROUPNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@GROUPNAME", GROUPNAME));



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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {

                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[0];

                    string T2 = row.Cells["ID"].Value.ToString().Trim();
                    string T3 = row.Cells["簽核狀態"].Value.ToString().Trim();

                    if (T3 == "已結案")
                    {
                        MessageBox.Show("單據已審核");
                        return;

                    }


                    System.Data.DataTable T1 = GetFLOW2(T2);

                    if (T1.Rows.Count > 0)
                    {
                        DataRow drw = T1.Rows[0];
                        MailMessage message = new MailMessage();
                        // message.To.Add("LLEYTONCHEN@ACMEPOINT.COM");
                         message.To.Add(drw["MAILTO"].ToString());
                        message.Subject = "回簽提醒- "+drw["MAILHEAD"].ToString();
                        message.Body = drw["MAILTEMP"].ToString();


                        message.IsBodyHtml = true;

                        SmtpClient client = new SmtpClient();
                        client.Send(message);



                        MessageBox.Show("寄信成功");
                    }

                }
               

            }

            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public System.Data.DataTable GetFLOW2(string MAILDOC)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM ACMESQLEEP.DBO.ACME_MAIL_BACKUP WHERE LTRIM(MAILDOC)=@MAILDOC  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



            command.Parameters.Add(new SqlParameter("@MAILDOC", MAILDOC));
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
    }
}
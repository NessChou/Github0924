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
            sb.Append("    SELECT distinct T0.DOC ID,T0.DOCENTRY SAP�渹,T0.CARDNAME �Ȥ�W��,CASE T1.FLOWFLAG WHEN 'N' THEN '�f�֤�' WHEN 'P' THEN '�Q�h�^' WHEN 'Z' THEN '�w����' WHEN '' THEN '���e�X' ELSE FLOWFLAG END ñ�֪��A");
            sb.Append("               , T2.SLPNAME �~��,CONVERT(varchar(12),T0.DOCDATE,111) �ɶ� FROM  acmesqleep.dbo.ACME_AUTOODLN T0");
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
                    string T1 = row.Cells["SAP�渹"].Value.ToString().Trim();
                    string T2 = row.Cells["ID"].Value.ToString().Trim();
                    string T3 = row.Cells["ñ�֪��A"].Value.ToString().Trim();
                    string T4 = row.Cells["�ɶ�"].Value.ToString().Trim();
                    //if (T3 != "�w����")
                    //{
                    //    MessageBox.Show("���f�ֵL�k�ץX");
                    //    return;

                    //}
                    if (!String.IsNullOrEmpty(T1))
                    {
                        System.Data.DataTable G2 = GetODLN1(T2);
                        string TIME = "";
                        if (G2.Rows.Count > 0)
                        {
                            TIME = G2.Rows[0]["�ɶ�"].ToString();
                        }
                        else
                        {
                            TIME = T4;
                        }

                        System.Data.DataTable G1 = GetODLN(T1, TIME);
                        string SALESS = G1.Rows[0]["�P��H��"].ToString();

                        string USER = G1.Rows[0]["USERNAME"].ToString().ToUpper();
                        string FileName = string.Empty;

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


                        FileName = lsAppDir + "\\Excel\\wh\\�P�f��.xls";
                        string SALES = G1.Rows[0]["SALES"].ToString();

                        //Excel���˪���

                        System.Data.DataTable G3 = Getgroup(SALES);

                        if (G3.Rows.Count > 0)
                        {
                            SALES = "�~��-�\�ߦp";
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
            sb.Append("                                 SELECT Convert(varchar(11),T0.TAXDATE,111)  �P�f���,Convert(varchar(11),T0.TAXDATE,111) �L����");
            sb.Append("                                    ,T0.CARDCODE �Ȥ�s��,T0.CARDNAME �Ȥ�W��,T0.ADDRESS �b��a�},T0.ADDRESS2 �e�f�a�}");
            sb.Append("                                 ,T1.ITEMCODE ���~�s��,T1.DSCRIPTION �~�W�W��,T1.QUANTITY �ƶq,T1.PRICE ���,T1.QUANTITY*T1.PRICE ���B");
            sb.Append("                                 ,T2.zipcode �p���H,T2.block �s���q��,T2.CITY �ǯu���X,SALUNITMSR ���,T1.CURRENCY ���O�s��,T0.DOCENTRY �渹,T0.vatsum �|�B, T0.doctotal �`���B,T0.doctotal-T0.vatsum ���|");
            sb.Append("                                 ,T0.U_ACME_DOC_RATE �ײv,T4.SLPNAME �P��H��,T0.U_ACME_USER  �s��H��,T5.LICTRADNUM �Τ@�s��,T6.WHSNAME �ܮw�W��,CASE ISNULL(NUMATCARD,'')  WHEN '' THEN '' ELSE 'PO#'+NUMATCARD END+' '+U_ACME_PAYGUI �Ƶ�,T7.GROUPNAME SALES,Convert(varchar(11),T0.CREATEDATE,111) �s��ɶ�,T8.HOMETEL USERNAME,T1.BASEENTRY �ӷ��渹,@H1 AS �֭�ɶ� FROM ODLN T0 ");
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
            sb.Append(" select CONVERT(varchar(12), CAST(UPDATE_DATE AS DATETIME),111) �ɶ� from AcmeSqlEEP.dbo.SYS_TODOHIS");
            sb.Append("  where flow_desc='�P�f��y�{(TFT)' AND S_STEP_ID='�~�ȼf��' AND ISNULL(form_presentation,'') <> '' AND REPLACE(substring(form_presentation,15,DataLength(form_presentation)-15),'''','')=@DOCENTRY");

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
            sb.Append(" SELECT GROUPID FROM AcmeSqlEEP.dbo.GROUPS WHERE DESCRIPTION LIKE '%����%' AND GROUPNAME=@GROUPNAME ");

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
                    string T3 = row.Cells["ñ�֪��A"].Value.ToString().Trim();

                    if (T3 == "�w����")
                    {
                        MessageBox.Show("��ڤw�f��");
                        return;

                    }


                    System.Data.DataTable T1 = GetFLOW2(T2);

                    if (T1.Rows.Count > 0)
                    {
                        DataRow drw = T1.Rows[0];
                        MailMessage message = new MailMessage();
                        // message.To.Add("LLEYTONCHEN@ACMEPOINT.COM");
                         message.To.Add(drw["MAILTO"].ToString());
                        message.Subject = "�^ñ����- "+drw["MAILHEAD"].ToString();
                        message.Body = drw["MAILTEMP"].ToString();


                        message.IsBodyHtml = true;

                        SmtpClient client = new SmtpClient();
                        client.Send(message);



                        MessageBox.Show("�H�H���\");
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
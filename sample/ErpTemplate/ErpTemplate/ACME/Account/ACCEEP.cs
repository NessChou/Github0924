using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class ACCEEP : Form
    {
        string strEEP = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public ACCEEP()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getcc();


  
        }

        public System.Data.DataTable Getcc()
        {
            SqlConnection MyConnection = new SqlConnection(strEEP);

            StringBuilder sb = new StringBuilder();
            sb.Append("                    SELECT LISTID,申請日期,COMP 公司名稱,部門代號, 部門名稱,使用者 申請人名稱,ID EEP單號,  ");
            sb.Append("                         狀態  FROM ( SELECT LISTID,CARDNAME 受款人名稱,SERNO 受款人統編,        ");
            sb.Append("                                                                                   CASE  t4.flow_desc ");
            sb.Append("																				   WHEN '支付通知單(聿豐)' THEN '聿豐實業股份有限公司'               ");
            sb.Append("                                                                                   WHEN '支付通知單總務(宇豐)' THEN '宇豐光電股份有限公司'        ");
            sb.Append("                                                                                   WHEN '支付通知單總務(聿豐)' THEN '聿豐實業股份有限公司'        ");
            sb.Append("                                                                                   WHEN '支付通知單總務(博豐)' THEN '博豐光電股份有限公司'        ");
            sb.Append("                                                                                   WHEN '支付通知單總務(韋峰)' THEN '韋峰能源股份有限公司'        ");
            sb.Append("                                                                                     WHEN '支付通知單(禾豐)' THEN '禾豐畜牧場'         ");
            sb.Append("                                                                                   ELSE '進金生實業股份有限公司'  END COMP,       ");
            sb.Append("                                                                                   CASE WHEN t4.flow_desc  IN ( '支付通知單總務(宇豐)','支付通知單總務(聿豐)', '支付通知單總務(博豐)' ,'支付通知單總務(韋峰)','支付通知單(聿豐)' , '支付通知單(禾豐)' )  THEN ''        ");
            sb.Append("                                                                                   ELSE BUID END 部門代號,t6.PARAM_DESC 部門名稱                                                           ");
            sb.Append("                                                                                   ,T0.MEMO 申請備註,T4.日期1 申請日期,T4.時間1 申請時間            ");
            sb.Append("                                                                                   ,T0.ID,UserSign 使用者, CASE T0.FlowFlag WHEN 'Z' THEN '結案' WHEN 'P' THEN '進行中' WHEN 'N' THEN '取回' WHEN 'X' THEN '作廢'  END 狀態,       ");
            sb.Append("                                                                                   T0.BANK4 分行,T0.MM,'單號:'+T0.ID DOCID,T0.MM3,T0.MU                    ");
            sb.Append("                                                                                   FROM ACME_OITT  T0                                                              ");
            sb.Append("                                                                                   LEFT JOIN (SELECT  (substring(form_presentation,6,DataLength(form_presentation)-6)) ID, MAX(CONVERT(varchar(12), CAST(UPDATE_DATE AS DATETIME),112)) 日期1,MAX(UPDATE_TIME) 時間1,MAX(USERNAME) 簽核人1,MAX(REMARK) 備註1,MAX(flow_desc) flow_desc,MAX(LISTID) LISTID FROM dbo.SYS_TODOHIS where flow_desc in ('支付通知單(費用類)','支付通知單總務(宇豐)','支付通知單總務(聿豐)','支付通知單總務(博豐)','支付通知單總務(韋峰)','支付通知單總務(韋豐)','支付通知單(聿豐)' , '支付通知單(禾豐)' ) AND S_STEP_ID IN ('費用申請')  AND STATUS NOT IN ('NR','NF')  GROUP BY substring(form_presentation,6,DataLength(form_presentation)-6) ) T4 ON (T0.ID=REPLACE(T4.ID,'''',''))                                       ");
            sb.Append("             																		LEFT JOIN PARAMS T6 ON (T0.BUID =T6.PARAM_NO)  ");
            sb.Append("                                                								   ) AS A      ");

            sb.Append("              WHERE COMP=@公司名稱  ");
            sb.Append("   AND 申請日期 BETWEEN @申請日期1  AND @申請日期2 ");
            if (comboBox3.Text != "")
            {
                sb.Append("   AND 使用者=@使用者 ");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("   AND 狀態=@狀態 ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append("   AND ID BETWEEN @ID1  AND @ID2 ");
            }
            sb.Append(" UNION ALL ");
            sb.Append("                    SELECT LISTID,申請日期,COMP 公司名稱,部門代號, 部門名稱,使用者 申請人名稱,ID EEP單號,  ");
            sb.Append("                         狀態  FROM ( SELECT LISTID,CARDNAME 受款人名稱,SERNO 受款人統編,        ");
            sb.Append("                                                                                 COMPANY COMP,       ");
            sb.Append("                                                                                 BUID 部門代號,t6.PARAM_DESC 部門名稱                                                           ");
            sb.Append("                                                                                   ,T0.MEMO 申請備註,T4.日期1 申請日期,T4.時間1 申請時間            ");
            sb.Append("                                                                                   ,T0.ID,UserSign 使用者, CASE T0.FlowFlag WHEN 'Z' THEN '結案' WHEN 'P' THEN '進行中' WHEN 'N' THEN '取回' WHEN 'X' THEN '作廢'  END 狀態,       ");
            sb.Append("                                                                                   T0.BANK4 分行,T0.MM,'單號:'+T0.ID DOCID,T0.MM3,T0.MU                    ");
            sb.Append("                                                                                   FROM ACME_OITT  T0                                                              ");
            sb.Append("                                                              INNER JOIN (SELECT  (substring(form_presentation,6,DataLength(form_presentation)-6)) ID, MAX(CONVERT(varchar(12), CAST(UPDATE_DATE AS DATETIME),112)) 日期1,MAX(UPDATE_TIME) 時間1,MAX(USERNAME) 簽核人1,MAX(REMARK) 備註1,MAX(flow_desc) flow_desc,MAX(LISTID) LISTID FROM dbo.SYS_TODOHIS where flow_desc = ('支付通知單(管理部)') AND S_STEP_ID IN ('費用申請')  AND STATUS NOT IN ('NR','NF')  GROUP BY substring(form_presentation,6,DataLength(form_presentation)-6) ) T4 ON (T0.ID=REPLACE(T4.ID,'''',''))                                      ");
            sb.Append("             																		LEFT JOIN PARAMS T6 ON (T0.BUID =T6.PARAM_NO)  	   WHERE t4.flow_desc ='支付通知單(管理部)' AND ISNULL(COMPANY,'') <>'' ");
            sb.Append("                                                								   ) AS A   ");


            sb.Append("              WHERE COMP=@公司名稱  ");
            sb.Append("   AND 申請日期 BETWEEN @申請日期1  AND @申請日期2 ");
            if (comboBox3.Text != "")
            {
                sb.Append("   AND 使用者=@使用者 ");
            }
            if (comboBox2.Text != "")
            {
                sb.Append("   AND 狀態=@狀態 ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append("   AND ID BETWEEN @ID1  AND @ID2 ");
            }
            sb.Append("   ORDER BY ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@公司名稱", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@申請日期1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@申請日期2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@使用者", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@狀態", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@ID1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@ID2", textBox4.Text));
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

        public System.Data.DataTable Getcc2()
        {
            SqlConnection MyConnection = new SqlConnection(strEEP);

            StringBuilder sb = new StringBuilder();
            sb.Append("      					SELECT DISTINCT UserSign     FROM ACME_OITT WHERE FlowFlag='Z' ORDER BY UserSign");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@公司名稱", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@申請日期1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@申請日期2", textBox6.Text));
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
        public System.Data.DataTable GetOPTW(string ID)
        {

            SqlConnection MyConnection = new SqlConnection(strEEP);
            StringBuilder sb = new StringBuilder();
            sb.Append(" 					   SELECT * FROM ACME_ITT2 WHERE ID=@ID     ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void ACCAR_Load(object sender, EventArgs e)
        {

            textBox5.Text = GetMenu.DFirst();

            textBox6.Text = GetMenu.DLast();

            System.Data.DataTable dt4 = Getcc2();
            string USER = globals.UserID.ToUpper();
            if (USER == "FIONALAI")
            {

                comboBox1.Items.Clear();

                comboBox1.Items.Add("宇豐光電股份有限公司");
                comboBox1.Items.Add("博豐光電股份有限公司");

                comboBox1.Text = "宇豐光電股份有限公司";
                button1.Visible = true;

            }
            if (USER == "SHARONHUANG")
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("聿豐實業股份有限公司");
                comboBox1.Text = "聿豐實業股份有限公司";
                button1.Visible = true;
            }
            if (USER == "ULATSAI")
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("聿豐實業股份有限公司");
                comboBox1.Items.Add("禾豐畜牧場");
                comboBox1.Text = "聿豐實業股份有限公司";
                button1.Visible = true;
            }
            if (USER == "VICKYHSIAO")
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("進金生實業股份有限公司");
                comboBox1.Items.Add("韋峰能源股份有限公司");
                comboBox1.Text = "進金生實業股份有限公司";
                button1.Visible = true;
            }
            if (USER == "NANCYWEI" || globals.GroupID.ToString().Trim() == "EEP")
            {

                comboBox1.Items.Clear();
                comboBox1.Items.Add("進金生實業股份有限公司");
                comboBox1.Items.Add("宇豐光電股份有限公司");
                comboBox1.Items.Add("聿豐實業股份有限公司");
                comboBox1.Items.Add("博豐光電股份有限公司");
                comboBox1.Items.Add("韋峰能源股份有限公司");
                comboBox1.Items.Add("禾豐畜牧場");
                //禾豐畜牧場

                comboBox1.Text = "進金生實業股份有限公司";
                button1.Visible = true;
            }

            comboBox3.Items.Clear();

            comboBox3.Items.Add("");
            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        private System.Data.DataTable GetOrderDataAPL(string MAILDOC)
        {
            SqlConnection connection = globals.EEPConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAILTEMP  FROM ACME_MAIL_BACKUP2  WHERE  MAILDOC=@MAILDOC AND FLOWTYPE LIKE '%通知%' ");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@MAILDOC", MAILDOC));
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
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "送簽文件路徑")
                {
                    string EEP單號 = dataGridView1.CurrentRow.Cells["EEP單號"].Value.ToString();

                    System.Data.DataTable gg1 = null;

                    gg1 = GetOPTW(EEP單號);
        
                    if (gg1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= gg1.Rows.Count - 1; i++)
                        {
                            string path = gg1.Rows[i]["path"].ToString();
                            string filename = gg1.Rows[i]["filename"].ToString();


                            string aa = path;

                            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                            string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                            System.IO.File.Copy(aa, NewFileName, true);

                            System.Diagnostics.Process.Start(NewFileName);
                        }

                    }

                }

                if (dgv.Columns[e.ColumnIndex].Name == "EEP")
                {
                    string LISTID = dataGridView1.CurrentRow.Cells["LISTID"].Value.ToString();

                    System.Data.DataTable gg1 = null;

                    gg1 = GetOrderDataAPL(LISTID);

                    if (gg1.Rows.Count > 0)
                    {
                
                            string MAILTEMP = gg1.Rows[0]["MAILTEMP"].ToString();

                            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                            string NewFileName = lsAppDir + "\\MailTemplates\\LISTID.htm";

                            File.WriteAllText(NewFileName, MAILTEMP);
                            System.Diagnostics.Process.Start(NewFileName);
                        

                    }
                    else
                    {
                        MessageBox.Show("尚未結案");

                    }

                }
            }
            catch { }
        }
    }
}

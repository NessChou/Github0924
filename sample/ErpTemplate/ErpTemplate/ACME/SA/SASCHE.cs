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
    public partial class SASCHE : Form
    {
        StringBuilder sbH = new StringBuilder();
        string LOGIN = fmLogin.LoginID.ToString().ToUpper();
        string mail = "";
        int scrollPosition = 0;
        public SASCHE()
        {
            InitializeComponent();
        }

   

        private System.Data.DataTable Get1()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.LINENUM ,T0.DOCENTRY ,t3.numatcard PO,T0.ITEMCODE,t0.dscription ITEMNAME,cast(T0.quantity as int) QTY,  ");
            sb.Append(" case when t3.cardname like  '%TOP GARDEN INT%' then 'TOP GARDEN' when t3.cardname like  '%CHOICE CHANNEL%' then 'CHOICE' when t3.cardname like  '%Infinite Power Group%' then 'INFINITE' when t3.cardname like  '%宇豐光電股份有限公司%' then '宇豐'  when t3.cardname like  '%達睿生%' then 'DRS' else t3.cardname end+CASE ISNULL(T3.U_BENEFICIARY,'') WHEN '' THEN '' ELSE '-'+T3.U_BENEFICIARY END CARDNAME,  ");
            sb.Append(" cast(round(T0.PRICE,2)   as   numeric(15,2)) PRICE,t7.whscode WHCODE,t7.WHSNAME WHNAME,T0.u_acme_workday+'('+CAST(T2.day AS VARCHAR)+')' WORKDAY,  ");
            sb.Append(" Convert(varchar(8),T0.u_acme_shipday,112)  LEAVEDAY, Convert(varchar(8),T0.u_acme_work,112)  SCHEDAY,T0.U_PAY PAY,T0.U_SHIPDAY SHIPDAY  ");
            sb.Append(" ,T0.U_SHIPSTATUS [STATUS],T0.U_MARK MARK,T0.U_MEMO MEMO,T3.address2 RADDRESS,T3.[address] PADDRESS,  ");
            sb.Append(" T3.u_acme_tardeTERM  TERM, T3.u_acme_SHIPFoRM1 SHIPFROM,T3.u_acme_SHIPTO1 SHIPTO,T3.U_ACME_BYAIR SHIPWAY ");
            sb.Append(" ,(T4.[SlpName]) SALES,(T5.[lastName]+T5.[firstName]) SA,T3.U_ACME_FORWARDER FORWARDER,T8.ShippingCode  WH,T9.ShippingCode  SH                 ");
            sb.Append(" FROM acmesql02.dbo.rdr1 T0      ");
            sb.Append(" left join  acmesqlsp.dbo.WorkDay T2 on (T2.workday=T0.u_acme_workday ) ");
            sb.Append(" left join  acmesql02.dbo.ORDR T3 on (T0.DOCENTRY=T3.DOCENTRY )     ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OSLP T4 ON T3.SlpCode = T4.SlpCode    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OHEM T5 ON T3.OwnerCode = T5.empID    ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.owhs T7 ON T0.whscode=T7.whscode    ");
            sb.Append(" LEFT JOIN acmesqlsp.dbo.WH_ITEM T8 ON (CAST(T0.DocEntry AS VARCHAR) =CAST(T8.Docentry AS VARCHAR) AND T0.LineNum =T8.linenum )");
            sb.Append(" LEFT JOIN acmesqlsp.dbo.Shipping_Item  T9 ON (CAST(T0.DocEntry AS VARCHAR) =CAST(T9.Docentry AS VARCHAR) AND T0.LineNum =T9.linenum )");
            sb.Append(" where t3.canceled <> 'Y' AND T3.doctype='I'   ");
            sb.Append("and  T0.DOCENTRY =  '" + textBox3.Text.ToString().Trim().Replace(" ", "") + "'  ");
            if (textBox1.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T0.U_ACME_WORK,112)  = '" + textBox1.Text.ToString() + "'  ");
            }
            //if (textBox3.Text != "")
            //{
            //    sb.Append("and  T0.DOCENTRY =  '" + textBox3.Text.ToString().Trim().Replace(" ","") + "'  ");
            //}
            //else
            //{
            //    if (textBox1.Text != "")
            //    {
            //        sb.Append("and  Convert(varchar(8),T0.U_ACME_WORK,112)  >= '" + textBox1.Text.ToString() + "'  ");
            //    }
            //    if (textBox2.Text != "")
            //    {
            //        sb.Append("and  Convert(varchar(8),T0.U_ACME_WORK,112)  <=  '" + textBox2.Text.ToString() + "'  ");
            //    }

            //    if (comboBox1.SelectedValue.ToString() != "Please-Select")
            //    {
            //        sb.Append("and T4.SlpName =@SlpName  ");
            //    }
            //    if (comboBox2.SelectedValue.ToString() != "Please-Select")
            //    {
            //        sb.Append(" and T5.[lastName]+T5.[firstName]=@lastName ");
            //    }
            //}
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@lastName", comboBox2.SelectedValue.ToString()));
            //command.Parameters.Add(new SqlParameter("@SlpName", comboBox1.SelectedValue.ToString()));

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
        private System.Data.DataTable Get2()
        {
            string V1 = "";
            System.Data.DataTable G1 = GetVFORWARDER();
            if (G1.Rows.Count > 0)
            {
                V1 = "Y";
            }
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(WH1,'')+CAST(WH AS VARCHAR) 倉管工單,ISNULL(SH1,'')+CAST(SH AS VARCHAR) 船務工單,ISNULL(D1,'')+CAST(DOCENTRY AS VARCHAR) 銷售訂單,ISNULL(CARDNAME1,'')+CARDNAME 客戶,ISNULL(PO1,'')+ISNULL(PO,'')  PO,ISNULL(ITEMCODE1,'')+ITEMCODE 料號 ");
            sb.Append(" ,ISNULL(QTY1,'')+QTY 數量,ISNULL(PRICE1,'')+PRICE 單價 ");
            sb.Append(" ,ISNULL(WHNAME1,'')+WHNAME 倉庫名稱,ISNULL(WORKDAY1,'')+ISNULL(WORKDAY,'') 工作天數,ISNULL(LEAVEDAY1,'')+ISNULL(LEAVEDAY,'') 離倉日期 ");
            sb.Append(" ,ISNULL(SCHEDAY1,'')+ISNULL(SCHEDAY,'') 排程日期,ISNULL(PAY1,'')+ISNULL(PAY,'') 付款,ISNULL(SHIPDAY1,'')+ISNULL(SHIPDAY,'') 押出貨日 ");
            sb.Append(" ,ISNULL(STATUS1,'')+ISNULL([STATUS],'') 貨況,ISNULL(MARK1,'')+ISNULL(MARK,'') 特殊嘜頭");
            sb.Append(" ,ISNULL(MEMO1,'')+ISNULL(MEMO,'') 注意事項 ");
            sb.Append(" ,ISNULL(TERM1,'')+ISNULL(TERM,'') TERM ");
            sb.Append(" ,ISNULL(SHIPFROM1,'')+ISNULL(SHIPFROM,'') SHIPFROM,ISNULL(SHIPTO1,'')+ISNULL(SHIPTO,'') SHIPTO,ISNULL(SHIPWAY1,'')+ISNULL(SHIPWAY,'') 運輸方式");
            sb.Append("");
            if (V1 == "Y")
            {
                sb.Append(",ISNULL(FORWARDER1,'')+FORWARDER FORWARDER");
            }
            sb.Append(" ,ISNULL(SALES1,'')+SALES 業務,ISNULL(SA1,'')+SA 業助");
            sb.Append("  FROM SASCHE WHERE [LOGIN] =@LOGIN ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", fmLogin.LoginID.ToString().ToUpper()));
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

        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT DOCENTRY 銷售訂單,CARDNAME FROM SASCHE  WHERE [LOGIN] =@LOGIN");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", fmLogin.LoginID.ToString().ToUpper()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable Get4(string DTYPE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ID FROM SASCHE WHERE [LOGIN] =@LOGIN AND  " + DTYPE + "='TTRU' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", fmLogin.LoginID.ToString().ToUpper()));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static int GetWorkingDayOfWeek(DateTime dtDate)
        {

            double a = 0.0;

            switch (dtDate.DayOfWeek)
            {

                case DayOfWeek.Sunday:

                    a = 7.0;

                    break;



                case DayOfWeek.Monday:

                    a = 1.0;

                    break;



                case DayOfWeek.Tuesday:

                    a = 2.0;

                    break;



                case DayOfWeek.Wednesday:

                    a = 3.0;

                    break;



                case DayOfWeek.Thursday:

                    a = 4.0;

                    break;



                case DayOfWeek.Friday:

                    a = 5.0;

                    break;



                case DayOfWeek.Saturday:

                    a = 6.0;

                    break;

            }

            return (int)Math.Round(a);

        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];



                if (column.Name == "PO1")
                {
                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                    if (row != null)
                    {

                        string PO = Convert.ToString(row["PO1"]).Trim();
                        if (PO == "Y")
                        {
                            sASCHEDataGridView.Rows[e.RowIndex].Cells["PO"].Style.BackColor = Color.Yellow;
                        }
                        else
                        {
                            sASCHEDataGridView.Rows[e.RowIndex].Cells["PO"].Style.BackColor = Color.White;
                        }
                    }
                }




            }
        }

        private void sASCHEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sASCHEBindingSource.EndEdit();
            this.sASCHETableAdapter.Update(this.sa.SASCHE);

            MessageBox.Show("存檔成功");

        }

        private void SASCHE_Load(object sender, EventArgs e)
        {
            string LOGIN = fmLogin.LoginID.ToString().ToUpper();
            comboBox1.Text = "更正";
            TRUNG(LOGIN);
            ////UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetOslp1(), "DataValue", "DataValue");
            ////UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetOhem(), "DataValue", "DataValue");

            System.Data.DataTable O2 = GetMenu.GetSHIPOHEM(LOGIN);
                        if (O2.Rows.Count > 0)
                        {
                            label1.Visible = true;
                            textBox1.Visible = true;
                        }
            //textBox1.Text = DS.ToString("yyyyMMdd");
            //textBox2.Text = DD.ToString("yyyyMMdd");

            System.Data.DataTable T1 = GetMenu.GetWHSA();

            listBox1.Items.Clear();

            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string F1 = T1.Rows[i][0].ToString();
                listBox1.Items.Add(F1);
            }

  


        }


        private void button1_Click(object sender, EventArgs e)
        {
            //string LOGIN = fmLogin.LoginID.ToString().ToUpper();

            //TRUNG(LOGIN);
           
            System.Data.DataTable dt1 = Get1();

                System.Data.DataTable dt2 = sa.SASCHE;
                if (dt1.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        
                        drw2["DOCENTRY"] = drw["DOCENTRY"];
                        drw2["LINENUM"] = drw["LINENUM"];
                        drw2["CARDNAME"] = drw["CARDNAME"];
                        drw2["PO"] = drw["PO"];
                        drw2["ITEMCODE"] = drw["ITEMCODE"];
                        drw2["ITEMNAME"] = drw["ITEMNAME"];
                        drw2["QTY"] = drw["QTY"];
                        drw2["PRICE"] = drw["PRICE"];
                        drw2["WHCODE"] = drw["WHCODE"];
                        drw2["WHNAME"] = drw["WHNAME"];
                        drw2["WORKDAY"] = drw["WORKDAY"];
                        drw2["LEAVEDAY"] = drw["LEAVEDAY"];
                        drw2["SCHEDAY"] = drw["SCHEDAY"];
                        drw2["PAY"] = drw["PAY"];

                        drw2["SHIPDAY"] = drw["SHIPDAY"];
                        drw2["STATUS"] = drw["STATUS"];
                        drw2["MARK"] = drw["MARK"];
                        drw2["MEMO"] = drw["MEMO"];
                        drw2["RADDRESS"] = drw["RADDRESS"];
                        drw2["PADDRESS"] = drw["PADDRESS"];
                        drw2["TERM"] = drw["TERM"];
                        drw2["SHIPFROM"] = drw["SHIPFROM"];
                        drw2["SHIPTO"] = drw["SHIPTO"];
                        drw2["SHIPWAY"] = drw["SHIPWAY"];
                        drw2["FORWARDER"] = drw["FORWARDER"];
                        drw2["SA"] = drw["SA"];
                        drw2["SALES"] = drw["SALES"];
                        drw2["WH"] = drw["WH"];
                        drw2["SH"] = drw["SH"];
                        drw2["LOGIN"] = fmLogin.LoginID.ToString().ToUpper();
                        
                        
                        dt2.Rows.Add(drw2);
                    }

                    this.Validate();
                    this.sASCHEBindingSource.EndEdit();
                    this.sASCHETableAdapter.Update(this.sa.SASCHE);

                    System.Data.DataTable G1 = GetVFORWARDER();
                    if (G1.Rows.Count > 0)
                    {
                        sASCHEDataGridView.Columns["DX48"].Visible = true;
                        sASCHEDataGridView.Columns["DX49"].Visible = true;
                    }
                    else
                    {
                        sASCHEDataGridView.Columns["DX48"].Visible = false;
                        sASCHEDataGridView.Columns["DX49"].Visible = false;
                    }

                }
                else
                {
                    MessageBox.Show("沒有資料");
                }
          
        }
        private System.Data.DataTable GetVFORWARDER()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT FORWARDER  FROM SASCHE WHERE  ISNULL(FORWARDER,'') <> '' AND [LOGIN] =@LOGIN  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));
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
        public void TRUNG(string LOGIN)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE SASCHE WHERE [LOGIN] =@LOGIN ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOGIN", LOGIN));

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

                //KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                //tmpKeyValue = KeyValue;

                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                //if (string.IsNullOrEmpty(tmpKeyValue))
                //{
                //    strB.AppendLine("<td>&nbsp;</td>");
                //}
                //else
                //{
                //    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                //}


                for (int d = 0; d <= dg.Rows[i].Cells.Count-1 ; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    string DG = dgvc.Value.ToString();
                    if (DG.Length > 3 && DG.Substring(0,4)=="TTRU")
                    {
                        //f8f1cb
                        strB.AppendLine("<td style='background:#f8f1cb'>" + dgvc.Value.ToString().Replace("TTRU","") + "</td>");
 
                    }
                    else
                    {
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
                }
                strB.AppendLine("</tr>");

            }

            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }
        private void sASCHEDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (sASCHEDataGridView.Columns[e.ColumnIndex].Name == "DX1")
            {
                string PO1 = sASCHEDataGridView.Rows[e.RowIndex].Cells["DX1"].Value.ToString();
                if (PO1 == "TTRU")
                {
                    for (int i = 3; i <= 56; i++)
                    {
                        if (i != 52)
                        {
                            if (i % 2 != 0)
                            {

                                this.sASCHEDataGridView.Rows[e.RowIndex].Cells["DX" + i.ToString()].Value = "TTRU";
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 3; i <= 56; i++)
                    {
                        if (i  != 52)
                        {
                            if (i % 2 != 0)
                            {

                                this.sASCHEDataGridView.Rows[e.RowIndex].Cells["DX" + i.ToString()].Value = "";
                            }
                        }
                    }
                }
            }
   
            for (int i = 1; i <= 56; i++)
            {
                if (sASCHEDataGridView.Columns[e.ColumnIndex].Name == "DX" + i.ToString())
                {
                    string PO = sASCHEDataGridView.Rows[e.RowIndex].Cells["DX" + i.ToString()].Value.ToString();
                    if (PO == "TTRU")
                    {
                        sASCHEDataGridView.Rows[e.RowIndex].Cells["DX" + (i + 1).ToString()].Style.BackColor = Color.Yellow;
                    }
                    else
                    {
                        sASCHEDataGridView.Rows[e.RowIndex].Cells["DX" + (i + 1).ToString()].Style.BackColor = Color.White;
                    }

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sASCHEBindingSource.EndEdit();
            this.sASCHETableAdapter.Update(this.sa.SASCHE);

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
                        string f = listBox1.SelectedItems[i].ToString();
                        al.Add(listBox1.SelectedItems[i].ToString());
                    }



                    foreach (string v in al)
                    {
                        ss.Append("" + v + "@acmepoint.com;");
                    }
                }

                if (checkBox3.Checked)
                {
                    System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHSHIP();
                    if (SHIPSTOCCK.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                        {
                            DataRow dd = SHIPSTOCCK.Rows[i];
                            ss.Append(dd["EMAIL"].ToString() + ";");
                        }
                    }
                }

                if (checkBox2.Checked)
                {
                    System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHSTOCK();
                    if (SHIPSTOCCK.Rows.Count > 0)
                    {
                        for (int i = 0; i <= SHIPSTOCCK.Rows.Count - 1; i++)
                        {
                            DataRow dd = SHIPSTOCCK.Rows[i];
                            ss.Append(dd["EMAIL"].ToString() + ";");
                        }
                    }
                }

                if (checkBox1.Checked)
                {
                    System.Data.DataTable SHIPSTOCCK = GetMenu.GetWHCN();
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



                System.Data.DataTable G3 = Get3();
                string CARDNAME = G3.Rows[0]["CARDNAME"].ToString();
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i <= G3.Rows.Count - 1; i++)
                {

                    DataRow dd = G3.Rows[i];


                    sb.Append(dd["銷售訂單"].ToString() + "/");


                }

                sb.Remove(sb.Length - 1, 1);
                //string DTYPE = "";
                //if (comboBox1.Text == "更正")
                //{
                //  //  F1();
                //    DTYPE = sbH.ToString();
                //}
                message.Subject = comboBox1.Text + "排程-" + CARDNAME + "-SO#" + sb.ToString(); 
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

        private void sASCHEDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void F1()
        {

            System.Data.DataTable  F1 = Get4("PO1");
            if (F1.Rows.Count > 0)
            {
                sbH.Append("PO/");
            }
            System.Data.DataTable F2 = Get4("ITEMCODE1");
            if (F2.Rows.Count > 0)
            {
                sbH.Append("料號/");
            }
            System.Data.DataTable F3 = Get4("ITEMNAME1");
            if (F3.Rows.Count > 0)
            {
                sbH.Append("項目說明/");
            }
            System.Data.DataTable F4 = Get4("QTY1");
            if (F4.Rows.Count > 0)
            {
                sbH.Append("數量/");
            }
            System.Data.DataTable F5 = Get4("PRICE1");
            if (F5.Rows.Count > 0)
            {
                sbH.Append("單價/");
            }
            System.Data.DataTable F6 = Get4("WHCODE1");
            if (F6.Rows.Count > 0)
            {
                sbH.Append("倉庫/");
            }
            System.Data.DataTable F7 = Get4("WHNAME1");
            if (F7.Rows.Count > 0)
            {
                sbH.Append("倉庫名稱/");
            }
            System.Data.DataTable F8 = Get4("WORKDAY1");
            if (F8.Rows.Count > 0)
            {
                sbH.Append("工作天數/");
            }
            System.Data.DataTable F9 = Get4("LEAVEDAY1");
            if (F9.Rows.Count > 0)
            {
                sbH.Append("離倉日期/");
            }
            System.Data.DataTable F10 = Get4("SCHEDAY1");
            if (F10.Rows.Count > 0)
            {
                sbH.Append("排程日期/");
            }
            System.Data.DataTable F11 = Get4("PAY1");
            if (F11.Rows.Count > 0)
            {
                sbH.Append("付款/");
            }
            System.Data.DataTable F12 = Get4("SHIPDAY1");
            if (F12.Rows.Count > 0)
            {
                sbH.Append("押出貨日/");
            }
            System.Data.DataTable F13 = Get4("STATUS1");
            if (F13.Rows.Count > 0)
            {
                sbH.Append("貨況/");
            }
            System.Data.DataTable F14 = Get4("MARK1");
            if (F14.Rows.Count > 0)
            {
                sbH.Append("特殊嘜頭/");
            }
            System.Data.DataTable F15 = Get4("MEMO1");
            if (F15.Rows.Count > 0)
            {
                sbH.Append("注意事項/");
            }
            System.Data.DataTable F16 = Get4("TERM1");
            if (F16.Rows.Count > 0)
            {
                sbH.Append("TERM/");
            }
            System.Data.DataTable F17 = Get4("SHIPFROM1");
            if (F17.Rows.Count > 0)
            {
                sbH.Append("SHIPFROM/");
            }
            System.Data.DataTable F18 = Get4("SHIPTO1");
            if (F18.Rows.Count > 0)
            {
                sbH.Append("SHIPTO/");
            }
            System.Data.DataTable F19 = Get4("SHIPWAY1");
            if (F19.Rows.Count > 0)
            {
                sbH.Append("運輸方式/");
            }

            System.Data.DataTable F22 = Get4("FORWARDER1");
            if (F22.Rows.Count > 0)
            {
                sbH.Append("FORWARDER/");
            }

            System.Data.DataTable F20 = Get4("SALES1");
            if (F20.Rows.Count > 0)
            {
                sbH.Append("業務/");
            }
            System.Data.DataTable F21 = Get4("SA1");
            if (F21.Rows.Count > 0)
            {
                sbH.Append("業助/");
            }

            if (sbH.Length > 0)
            {
                sbH.Remove(sbH.Length - 1, 1);
            }
        }




    
    }
}

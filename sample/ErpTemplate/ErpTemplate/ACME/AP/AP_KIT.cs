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
    public partial class AP_KIT : Form
    {

        string FA = "acmesql98";
        string strCn98 = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        Attachment data = null;
        string STATUS = "";
        public string q;
        public string q2;
        public AP_KIT()
        {
            InitializeComponent();
        }


        private void aP_KITBindingNavigatorSaveItem_Click_4(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KITBindingSource.EndEdit();
            this.aP_KITTableAdapter.Update(this.lC.AP_KIT);

            MessageBox.Show("更新成功");

        }
        public void AddBOM4(int DOCID, string STATUS, string PO, string PODATE, string CARDNAME, string ITEMCODE, string ITEMNAME, string BOM, string QTY, string CURRENCY, string AMT, string SO, string SALES, string SA, string CARDNAME2, string PRICE, string CUSTDATE, string MEMO, string path, string filename, string CREATRDATE, string OBYTYPE, string VER)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_KITLOG(DOCID,STATUS,PO,PODATE,CARDNAME,ITEMCODE,ITEMNAME,BOM,QTY,CURRENCY,AMT,SO,SALES,SA,CARDNAME2,PRICE,CUSTDATE,MEMO,path,filename,CREATRDATE,OBYTYPE,VER) values(@DOCID,@STATUS,@PO,@PODATE,@CARDNAME,@ITEMCODE,@ITEMNAME,@BOM,@QTY,@CURRENCY,@AMT,@SO,@SALES,@SA,@CARDNAME2,@PRICE,@CUSTDATE,@MEMO,@path,@filename,@CREATRDATE,@OBYTYPE,@VER)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCID", DOCID));
            command.Parameters.Add(new SqlParameter("@STATUS", STATUS));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@PODATE", PODATE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@BOM", BOM));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CURRENCY", CURRENCY));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@SO", SO));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@SA", SA));
            command.Parameters.Add(new SqlParameter("@CARDNAME2", CARDNAME2));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@CUSTDATE", CUSTDATE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@CREATRDATE", CREATRDATE));
            command.Parameters.Add(new SqlParameter("@OBYTYPE", OBYTYPE));
            command.Parameters.Add(new SqlParameter("@VER", VER));

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
        private void AP_KIT_Load(object sender, EventArgs e)
        {
            if (globals.GroupID.ToString().Trim() != "EEP")
            {
                if (globals.DBNAME == "進金生")
                {
                    strCn98 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
                    FA = "acmesql02";
                }
            }
            this.aP_KIT8TableAdapter.Fill(this.lC.AP_KIT8);
            comboBox2.Text = "未結";
            this.aP_KIT2TableAdapter.FillBy(this.lC.AP_KIT2, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text );
            this.aP_KIT8TableAdapter.FillBy(this.lC.AP_KIT8, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            comboBox1.Text = "進貨通知";

        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT2BindingSource.EndEdit();
            this.aP_KIT2TableAdapter.Update(this.lC.AP_KIT2);
            MessageBox.Show("更新成功");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.aP_KIT2TableAdapter.FillBy(this.lC.AP_KIT2, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KIT8TableAdapter.FillBy(this.lC.AP_KIT8, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                if (aP_KITDataGridView.SelectedRows.Count == 0)
                {

                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            if (tabControl1.SelectedIndex == 1)
            {
                if (aP_KIT2DataGridView.SelectedRows.Count == 0)
                {

                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            if (tabControl1.SelectedIndex == 2)
            {
                if (aP_KIT8DataGridView.SelectedRows.Count == 0)
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

            FileName = lsAppDir + "\\Excel\\OPCH\\KIT.xls";


          

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            System.Data.DataTable G1 = null;
            

            if (tabControl1.SelectedIndex == 0)
            {
                G1 = GetEXCEL(q,"AP_KIT");
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                G1 = GetEXCEL(q, "AP_KIT2");
                FileName = lsAppDir + "\\Excel\\OPCH\\KIT2.xls";
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                G1 = GetEXCEL(q, "AP_KIT8");
            }
            ExcelReport.ExcelReportOutput(G1, FileName, OutPutFile, "N");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl1.SelectedIndex == 0)
                {
                    if (aP_KITDataGridView.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("請選擇單據");
                        return;
                    }
                }
                if (tabControl1.SelectedIndex == 1)
                {
                    if (aP_KIT2DataGridView.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("請選擇單據");
                        return;
                    }
                }
                if (tabControl1.SelectedIndex == 2)
                {
                    if (aP_KIT8DataGridView.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("請選擇單據");
                        return;
                    }
                }
                string server = "//acmesrv01//SAP_Share//TTAdvance//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download(filename);

                string a1 = "";
                DataGridViewRow row = null;
                if (tabControl1.SelectedIndex == 0)
                {
                    row = aP_KITDataGridView.SelectedRows[0];
                    a1 = row.Cells["ID"].Value.ToString();

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    row = aP_KIT2DataGridView.SelectedRows[0];
                    a1 = row.Cells["ID2"].Value.ToString();
                }
                if (tabControl1.SelectedIndex == 2)
                {
                    row = aP_KIT8DataGridView.SelectedRows[0];
                    a1 = row.Cells["ID3"].Value.ToString();
                }
                if (result == DialogResult.OK)
                {
                    MessageBox.Show("上傳檔案" + Path.GetFileName(opdf.FileName));
                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }

                    string a2 = filename;

                    string a3 = @"\\acmesrv01\SAP_Share\TTAdvance\" + filename;

                    if (tabControl1.SelectedIndex == 0)
                    {
                        aP_KITDataGridView.SelectedRows[0].Cells["filename"].Value = a2;
                        aP_KITDataGridView.SelectedRows[0].Cells["path"].Value = a3;
                    }
                    if (tabControl1.SelectedIndex == 1)
                    {
                        aP_KIT2DataGridView.SelectedRows[0].Cells["filename2"].Value = a2;
                        aP_KIT2DataGridView.SelectedRows[0].Cells["path2"].Value = a3;
                    }
                    if (tabControl1.SelectedIndex == 2)
                    {
                        aP_KIT8DataGridView.SelectedRows[0].Cells["filename3"].Value = a2;
                        aP_KIT8DataGridView.SelectedRows[0].Cells["path3"].Value = a3;
                    }
                }
                else
                {
                    result = MessageBox.Show("是否要刪除附件", "YES/NO", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        if (tabControl1.SelectedIndex == 0)
                        {
                            Updatepath("", "", a1, "A");
                        }
                        if (tabControl1.SelectedIndex == 1)
                        {
                            Updatepath("", "", a1, "B");
                        }
                        if (tabControl1.SelectedIndex == 2)
                        {
                            Updatepath("", "", a1, "C");
                        }
                        this.aP_KIT2TableAdapter.FillBy(this.lC.AP_KIT2, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
                        this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
                        this.aP_KIT8TableAdapter.FillBy(this.lC.AP_KIT8, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
                    }
                }
             

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Updatepath(string filename, string path, string ID,string ID2)
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
            if (ID2 == "C")
            {
                sb.Append(" update AP_KIT8 set filename=@filename,[path]=@path where ID=@ID");
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

        private void aP_KITDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {
                   
                    for (int j = 0; j <= 1; j++)
                    {
                        string sd = aP_KITDataGridView.CurrentRow.Cells["ID"].Value.ToString();

                        System.Data.DataTable dt1 = GetPATH(sd);
                        if (dt1.Rows.Count > 0)
                        {
                            DataRow drw = dt1.Rows[0];

                            string aa = drw["path"].ToString();
                            System.Diagnostics.Process.Start(aa);
                        }
      


                   


                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void aP_KIT2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK2")
                {
                    for (int j = 0; j <= 1; j++)
                    {


                        System.Data.DataTable dt1 = lC.AP_KIT2;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];

                        string aa = drw["path"].ToString();


                        System.Diagnostics.Process.Start(aa);


                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            string TABLE = "";
            try
            {
      
                if (tabControl1.SelectedIndex == 0)
                {
                    TABLE = "AP_KIT";
                    if (aP_KITDataGridView.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    TABLE = "AP_KIT2";
                    if (aP_KIT2DataGridView.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 2)
                {
                    TABLE = "AP_KIT8";
                    if (aP_KIT8DataGridView.SelectedRows.Count == 0)
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
            if (comboBox1.Text == "進貨通知")
            {
       
                if (tabControl1.SelectedIndex == 0)
                {
                    MAIL = "\\MailTemplates\\KIT1.htm";
              
                }
                else
                {
                    MAIL = "\\MailTemplates\\KIT4.htm";
                }
                G1 = Getbb(q, TABLE);
                SUBJECT = G1.Rows[0]["廠商"].ToString() + "進貨內湖，請查收，謝謝!!";
                SA=G1.Rows[0]["SA"].ToString();
            }
            if (comboBox1.Text == "廠商詢價")
            {
                G1 = Getbb2(q, TABLE);
                MAIL = "\\MailTemplates\\KIT2.htm";
                SUBJECT ="請幫忙提供報價單 ("+ G1.Rows[0]["廠商"].ToString() + ")";
            }
            if (comboBox1.Text == "廠商訂單")
            {
                DELETEFILE();
                G1 = Getbb3(q, TABLE);
                MAIL = "\\MailTemplates\\KIT3.htm";
                SUBJECT = "ACME PO#" + G1.Rows[0]["採購單號"].ToString() + "-" + G1.Rows[0]["廠商"].ToString();
                CARD(G1);
            }

            if (G1.Rows.Count > 0)
            {
                string DOCTYPE = "L1";
                if (comboBox1.Text == "廠商訂單")
                {
                    DOCTYPE = "L2";
                }
                string CARDNAME = G1.Rows[0]["廠商"].ToString();
                dataGridView1.DataSource = G1;
                string GG = htmlMessageBody(dataGridView1).ToString();
                string EMAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                MailTest2(SUBJECT, EMAIL,GG, CARDNAME, MAIL,SA,DOCTYPE);
                MessageBox.Show("寄信成功");
            }
        }
        private void MailTest2(string strSubject, string MailAddress, string MailContent, string CUST, string MAIL, string SA, string DOCTYPE)
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
            string GetExePath= Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
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
            template = template.Replace("##KIT4##", "請留意:");
            template = template.Replace("##KIT5##", "出貨時，請將附檔制式麥頭印出 (附採購單號.、進金生料號、購買品名、數量)，並且黏貼於外箱上助於倉庫辨識及倉庫收貨時核對，謝謝幫忙！");
            string USER = fmLogin.LoginID.ToString();
            System.Data.DataTable T1 = GETOHEM(USER);
            if (T1.Rows.Count > 0)
            {
                template = template.Replace("##EMP##", T1.Rows[0]["EMP"].ToString());
                template = template.Replace("##TEL##", "Tel : 886-2-87912868 *" + T1.Rows[0]["分機"].ToString());

                string EMAIL=T1.Rows[0]["EMAIL"].ToString();
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
            if (DOCTYPE == "L2")
            {
                string OutPutFile = lsAppDir + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {

                    string m_File = "";

                    m_File = file;
                    data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                    //附件资料
                    ContentDisposition disposition = data.ContentDisposition;


                    // 加入邮件附件
                    message.Attachments.Add(data);

                }
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
                if (DOCTYPE == "L2")
                {
                    data.Dispose();
                    message.Attachments.Dispose();
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
            sb.Append(" SELECT ");
            if (TABLE != "AP_KIT2")
            {
                sb.Append("  CARDNAME 廠商,PO 採購單號,PODATE 交期,ITEMCODE 進金生料號,ITEMNAME 購買品名");
                sb.Append(" ,QTY 採購數量,Currency,AMT 採購金額,SO 銷售訂單,Sales,SA,CARDNAME2 客戶名稱,CUSTDATE 客戶希望交期,MEMO SA備註");

            }
            else
            {
                sb.Append("  PO 採購單號,BOM 成品料號,BOMNO 生產訂單NO,");
                sb.Append("  CARDNAME 廠商,ITEMCODE 進金生料號,ITEMNAME 購買品名");
                sb.Append(" ,QTY 採購數量,Currency,AMT 採購金額,SO 銷售訂單,Sales,SA,CARDNAME2 客戶名稱,CUSTDATE 客戶希望交期,MEMO SA備註");
            }
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
        public System.Data.DataTable Getbb2(string cs, string TABLE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT CARDNAME 廠商,ITEMCODE 進金生料號,ITEMNAME 購買品名 ");
            sb.Append("               ,'' 'BOM表or線材圖面' ,QTY 採購數量 FROM " + TABLE + "  WHERE ID in ( " + cs + ") ");

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

        public System.Data.DataTable Getbb3(string cs,string TABLE)
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
        private void CARD(System.Data.DataTable OrderData)
        {

            string FileName = string.Empty;
            string FileName1 = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\OPCH\\ACME進貨麥頭.xls";
            //Excel的樣版檔
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\wh\\ACME進貨麥頭(內湖).xls";
            ExcelReport.ExcelReportOutpuwh(OrderData, ExcelTemplate, OutPutFile, "N", null,"");


        }
       
        public System.Data.DataTable GetEXCEL(string cs, string DB)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT CASE STATUS WHEN 'True' then 'Close' else '' end Status,PO 採購單號,PODATE 採購確認交期,");
            sb.Append("                CARDNAME 廠商,PODATE 交期,ITEMCODE 進金生料號,ITEMNAME 購買品名 ");
            sb.Append("               ,QTY 採購數量,Currency,AMT 採購金額,SO 銷售訂單,Sales,SA,CARDNAME2 客戶名稱,PRICE 賣價,");
            if (DB == "AP_KIT2")
            {

                sb.Append("  BOMNO 生產訂單NO,");
                
            }
            sb.Append("               CUSTDATE 客戶希望交期,MEMO SA備註 FROM  " + DB + "");
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
        public System.Data.DataTable GetTOSAP(string cs, string DB)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CARDNAME,ITEMCODE,QTY,CASE WHEN CHARINDEX('(', AMT)=0 THEN AMT ELSE ltrim(substring(AMT,0,CHARINDEX('(', AMT))) END PRICE,CURRENCY,convert(varchar,CAST(SUBSTRING(CUSTDATE,1,8) AS DATETIME), 111)  CUSTDATE,");
            if (DB == "AP_KIT")
            {
                sb.Append(" 	             MEMO + ' '+ CARDNAME2 MEMO     FROM  " + DB + "");
            }
            else
            {
                sb.Append(" 	 REPLACE(ltrim(substring(CARDNAME2,0,CHARINDEX(' ', CARDNAME2))),',','')+' for '+ITEMNAME+PRICE+Sales MEMO   FROM  " + DB + "");
            }

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

        public System.Data.DataTable GetTOSAP2(string cs)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT BOM ITEMCODE,CARDNAME,QTY PQTY,SO,CAST(CUSTDATE AS DATETIME) ENDDATE,CARDNAME2 FROM aP_KIT2");
            sb.Append("   WHERE ID in ( " + cs + ") AND BOM <>'' ");


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

        public static System.Data.DataTable GetAPKIT(string USERS)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT   * FROM AP_KITSAP WHERE USERS=@USERS ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));
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
        public System.Data.DataTable GETCARCODE(string CARDNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDCODE FROM OCRD WHERE CARDNAME LIKE  '%" + CARDNAME + "%' AND SUBSTRING(CardCode,1,1)='S' AND frozenFor='N' ");

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


        public System.Data.DataTable GETCARCODE2(string CARDNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDNAME FROM OCRD WHERE CARDNAME LIKE  '%" + CARDNAME + "%' AND frozenFor='N' ");

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
        public System.Data.DataTable GETDATE()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT TOP 1 RATE  FROM ORTT WHERE Currency ='USD' ORDER BY RATEDATE DESC ");

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
        private void aP_KITDataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KITDataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KITDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KITDataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID"].Value.ToString());

                    }
                }
      

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textPO1.Text = "";
            textCARDNAME.Text = "";
            textITEMCODE.Text = "";
            textSA.Text = "";
            textCARDNAME2.Text = "";
            comboBox2.Text = "全部";

            this.aP_KIT2TableAdapter.FillBy(this.lC.AP_KIT2, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KIT8TableAdapter.FillBy(this.lC.AP_KIT8, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
        }

        private void aP_KIT2DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT2DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT2DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT2DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID2"].Value.ToString());

                    }
                }
    

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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


      

        private void aP_KITDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["dataGridViewTextBoxColumn2"].Value = "False";
        }

        private void aP_KIT2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["dataGridViewTextBoxColumn20"].Value = "False";
        }

        private void aP_KITDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (aP_KITDataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE")
                {
                    string ITEM = aP_KITDataGridView.Rows[e.RowIndex].Cells["ITEMCODE"].Value.ToString();

                    System.Data.DataTable T1 = GetITEM(ITEM);
                    if (T1.Rows.Count > 0)
                    {

                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["ITEMNAME"].Value = T1.Rows[0][0].ToString();
                    }

                }

                if (aP_KITDataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE" || aP_KITDataGridView.Columns[e.ColumnIndex].Name == "CARDNAME")
                {
                    string CARDNAME = aP_KITDataGridView.Rows[e.RowIndex].Cells["CARDNAME"].Value.ToString();
                    string ITEM = aP_KITDataGridView.Rows[e.RowIndex].Cells["ITEMCODE"].Value.ToString();
                    System.Data.DataTable G1 = GetPRICE(ITEM, CARDNAME);
                    if (G1.Rows.Count > 0)
                    {
                        decimal AMT = Convert.ToDecimal(G1.Rows[0][0]);

                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["AMT"].Value = (AMT).ToString("G29");

                    }
                }

                if (aP_KITDataGridView.Columns[e.ColumnIndex].Name == "SO" || aP_KITDataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE")
                {
                    string SO = aP_KITDataGridView.Rows[e.RowIndex].Cells["SO"].Value.ToString();
                    string ITEM = aP_KITDataGridView.Rows[e.RowIndex].Cells["ITEMCODE"].Value.ToString();
                    System.Data.DataTable G1 = GetORDR(ITEM, SO);
                    if (G1.Rows.Count > 0)
                    {
                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["SALES"].Value = G1.Rows[0]["業務"];
                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["SA"].Value = G1.Rows[0]["業管"];
                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["CARDNAME2"].Value = G1.Rows[0]["CARDNAME"];
                        this.aP_KITDataGridView.Rows[e.RowIndex].Cells["PRICE"].Value = Convert.ToDecimal(G1.Rows[0]["PRICE"]).ToString("G29");
                    }
                }
            }
            catch { }
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
        public System.Data.DataTable GetPRICE(string ITEMCODE, string CARDNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PRICE  FROM POR1 T0 LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE ITEMCODE=@ITEMCODE AND PRICE <>'0' AND T1.CARDNAME like '%" + CARDNAME + "%' ORDER BY T1.DOCDATE DESC");

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
        public System.Data.DataTable GetPRICE2(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.PRICE,T0.Currency 幣別,T1.CARDNAME  FROM POR1 T0 LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE ITEMCODE=@ITEMCODE AND PRICE <>'0' ORDER BY T1.DOCDATE DESC");

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
        public System.Data.DataTable GetORDR(string ITEMCODE, string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT (T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,T0.CARDNAME,T1.PRICE  FROM ORDR T0");
            sb.Append(" left join RDR1 t1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode");
            sb.Append(" iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID ");
            sb.Append(" WHERE T1.ITEMCODE=@ITEMCODE AND t0.docentry=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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

        public System.Data.DataTable GetORDR2(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.DOCENTRY SO,T1.ITEMCODE,T1.DSCRIPTION ITEMNAME,T1.Currency 幣別,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,T0.CARDNAME,T1.PRICE");
       //    sb.Append(" ,T1.QUANTITY 數量,T1.PRICE 單價,'交期'+CAST(CAST(SUBSTRING(Convert(varchar(10),U_ACME_SHIPDAY,112),5,2) AS INT) AS VARCHAR)+'/'+CAST(CAST(SUBSTRING(Convert(varchar(10),U_ACME_SHIPDAY,112),7,2) AS INT) AS VARCHAR) 排程日期 FROM ORDR T0 ");
            sb.Append(" ,T1.QUANTITY 數量,T1.PRICE 單價,Convert(varchar(8),U_ACME_SHIPDAY,112)  排程日期 FROM ORDR T0 ");
            sb.Append(" LEFT JOIN RDR1 t1 on (t0.docentry=t1.docentry)   ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("               WHERE  t0.DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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
        private void aP_KIT2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE2")
                {
                    string ITEM = aP_KIT2DataGridView.Rows[e.RowIndex].Cells["ITEMCODE2"].Value.ToString();

                    System.Data.DataTable T1 = GetITEM(ITEM);
                    if (T1.Rows.Count > 0)
                    {

                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["ITEMNAME2"].Value = T1.Rows[0][0].ToString();
                    }

                }


                if (aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE2" || aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "CARDNAME22"
                       || aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "QTY2")
                {
                    string CARDNAME = aP_KIT2DataGridView.Rows[e.RowIndex].Cells["CARDNAME22"].Value.ToString();
                    string ITEM = aP_KIT2DataGridView.Rows[e.RowIndex].Cells["ITEMCODE2"].Value.ToString();
                    System.Data.DataTable G1 = GetPRICE(ITEM, CARDNAME);
                    if (G1.Rows.Count > 0)
                    {
                        decimal AMT = Convert.ToDecimal(G1.Rows[0][0]);
                        decimal QTY = Convert.ToDecimal(aP_KIT2DataGridView.Rows[e.RowIndex].Cells["QTY2"].Value);
                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["AMT2"].Value = (AMT * QTY).ToString("G29");

                    }
                }

                if (aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "SO2" || aP_KIT2DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE2")
                {
                    string SO = aP_KIT2DataGridView.Rows[e.RowIndex].Cells["SO2"].Value.ToString();
                    string ITEM = aP_KIT2DataGridView.Rows[e.RowIndex].Cells["ITEMCODE2"].Value.ToString();
                    System.Data.DataTable G1 = GetORDR(ITEM, SO);
                    if (G1.Rows.Count > 0)
                    {
                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["SALES2"].Value = G1.Rows[0]["業務"];
                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["SA2"].Value = G1.Rows[0]["業管"];
                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["CARDNAME23"].Value = G1.Rows[0]["CARDNAME"];
                        this.aP_KIT2DataGridView.Rows[e.RowIndex].Cells["PRICE2"].Value = Convert.ToDecimal(G1.Rows[0]["PRICE"]).ToString("G29");
                    }
                }
            }
            catch { }
        }

        private void aP_KIT8DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK3")
                {
                    for (int j = 0; j <= 1; j++)
                    {


                        System.Data.DataTable dt1 = lC.AP_KIT8;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];

                        string aa = drw["path"].ToString();


                        System.Diagnostics.Process.Start(aa);


                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT8BindingSource.EndEdit();
            this.aP_KIT8TableAdapter.Update(this.lC.AP_KIT8);

            MessageBox.Show("更新成功");
        }

        private void aP_KIT8DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void aP_KIT8DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE3")
                {
                    string ITEM = aP_KIT8DataGridView.Rows[e.RowIndex].Cells["ITEMCODE3"].Value.ToString();

                    System.Data.DataTable T1 = GetITEM(ITEM);
                    if (T1.Rows.Count > 0)
                    {

                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["ITEMNAME3"].Value = T1.Rows[0][0].ToString();
                    }

                }


                if (aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE3" || aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "CARDNAME3"
                       || aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "QTY3")
                {
                    string CARDNAME = aP_KIT8DataGridView.Rows[e.RowIndex].Cells["CARDNAME3"].Value.ToString();
                    string ITEM = aP_KIT8DataGridView.Rows[e.RowIndex].Cells["ITEMCODE3"].Value.ToString();
                    System.Data.DataTable G1 = GetPRICE(ITEM, CARDNAME);
                    if (G1.Rows.Count > 0)
                    {
                        decimal AMT = Convert.ToDecimal(G1.Rows[0][0]);
                        decimal QTY = Convert.ToDecimal(aP_KIT8DataGridView.Rows[e.RowIndex].Cells["QTY3"].Value);
                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["AMT3"].Value = (AMT * QTY).ToString("G29");

                    }
                }

                if (aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "SO3" || aP_KIT8DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE3")
                {
                    string SO = aP_KIT8DataGridView.Rows[e.RowIndex].Cells["SO3"].Value.ToString();
                    string ITEM = aP_KIT8DataGridView.Rows[e.RowIndex].Cells["ITEMCODE3"].Value.ToString();
                    System.Data.DataTable G1 = GetORDR(ITEM, SO);
                    if (G1.Rows.Count > 0)
                    {
                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["SALES3"].Value = G1.Rows[0]["業務"];
                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["SA3"].Value = G1.Rows[0]["業管"];
                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["CARDNAME4"].Value = G1.Rows[0]["CARDNAME"];
                        this.aP_KIT8DataGridView.Rows[e.RowIndex].Cells["PRICE3"].Value = Convert.ToDecimal(G1.Rows[0]["PRICE"]).ToString("G29");
                    }
                }
            }
            catch { }
        }

        private void textBoxITEMNAME_TextChanged(object sender, EventArgs e)
        {
            this.aP_KIT2TableAdapter.FillBy(this.lC.AP_KIT2, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
            this.aP_KIT8TableAdapter.FillBy(this.lC.AP_KIT8, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);
        }

        private void aP_KIT8DataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aP_KIT8DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = aP_KIT8DataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aP_KIT8DataGridView.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["ID3"].Value.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetORDR2(textBox1.Text);
            if (dt1.Rows.Count > 0)
            {
                System.Data.DataTable dt2 = lC.AP_KIT;
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    string ITEMCODE = drw["ITEMCODE"].ToString();
                    drw2["ITEMCODE"] = ITEMCODE;
                    drw2["ITEMNAME"] = drw["ITEMNAME"];
                    drw2["QTY"] = Convert.ToDecimal(drw["數量"]).ToString("G29"); 
                    drw2["SO"] = drw["SO"];
                    drw2["SALES"] = drw["業務"];
                    drw2["SA"] = drw["業管"];
                    drw2["CARDNAME2"] = drw["CARDNAME"];
                    drw2["PRICE"] = drw["幣別"].ToString() + Convert.ToDecimal(drw["單價"]).ToString("G29");
                    drw2["CUSTDATE"] = drw["排程日期"].ToString();
                    //CUSTDATE
                    System.Data.DataTable G1 = GetPRICE2(ITEMCODE);
                    if (G1.Rows.Count > 0)
                    {
                        drw2["CURRENCY"] = G1.Rows[0]["幣別"].ToString();
                        drw2["CARDNAME"] = G1.Rows[0]["CARDNAME"].ToString();
                        drw2["AMT"] = (Convert.ToDecimal(G1.Rows[0]["PRICE"])).ToString("G29");
                    }
                    dt2.Rows.Add(drw2);
                }
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = GetORDR2(textBox2.Text);
            if (dt1.Rows.Count > 0)
            {
                System.Data.DataTable dt2 = lC.AP_KIT2;
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    string ITEMCODE = drw["ITEMCODE"].ToString();
                    drw2["BOM"] = ITEMCODE;
                    drw2["ITEMNAME"] = drw["ITEMNAME"];
                    drw2["QTY"] = Convert.ToDecimal(drw["數量"]).ToString("G29");
                    drw2["SO"] = drw["SO"];
                    drw2["SALES"] = drw["業務"];
                    drw2["SA"] = drw["業管"];
                    drw2["CARDNAME2"] = drw["CARDNAME"];
                    drw2["PRICE"] = drw["幣別"].ToString() + Convert.ToDecimal(drw["單價"]).ToString("G29");
                    drw2["CUSTDATE"] = drw["排程日期"].ToString();
                    System.Data.DataTable G1 = GetPRICE2(ITEMCODE);
                    if (G1.Rows.Count > 0)
                    {
                        drw2["CURRENCY"] = G1.Rows[0]["幣別"].ToString();
                        drw2["CARDNAME"] = G1.Rows[0]["CARDNAME"].ToString();
                        drw2["AMT"] = (Convert.ToDecimal(G1.Rows[0]["PRICE"])).ToString("G29");
                    }
                    dt2.Rows.Add(drw2);
                }
            }
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
        private void button10_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KITDataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KITDataGridView.SelectedRows[i];

                row.Cells[1].Value = "True";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KITDataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KITDataGridView.SelectedRows[i];

                row.Cells[1].Value = "False";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT2DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT2DataGridView.SelectedRows[i];

                row.Cells[1].Value = "True";
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT2DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT2DataGridView.SelectedRows[i];

                row.Cells[1].Value = "False";
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT8DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT8DataGridView.SelectedRows[i];

                row.Cells[1].Value = "True";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = aP_KIT8DataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = aP_KIT8DataGridView.SelectedRows[i];

                row.Cells[1].Value = "False";
            }
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
        public System.Data.DataTable GetDI5()
        {
            SqlConnection connection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(DOCENTRY) DOC FROM OWOR");
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
        private void button14_Click(object sender, EventArgs e)
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
                    if (aP_KITDataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 1)
                {
                    if (aP_KIT2DataGridView.SelectedRows.Count == 0)
                    {

                        MessageBox.Show("請點選單號");
                        return;
                    }

                }
                if (tabControl1.SelectedIndex == 2)
                {
                    if (aP_KIT8DataGridView.SelectedRows.Count == 0)
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
                string ID = sb.ToString();
                q2 = sb.ToString();
                System.Data.DataTable G1 = null;


                if (tabControl1.SelectedIndex == 0)
                {
                    G1 = GetTOSAP(q2, "AP_KIT");
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    G1 = GetTOSAP(q2, "AP_KIT2");
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    G1 = GetTOSAP(q2, "AP_KIT8");
                }



                if (G1.Rows.Count > 0)
                {
                    SAPbobsCOM.Documents oPURCH = null;
                    oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                    string CARDNAME = G1.Rows[0]["CARDNAME"].ToString();
                    string CURRENCY = G1.Rows[0]["CURRENCY"].ToString();
                    System.Data.DataTable G2 = GETCARCODE(CARDNAME);
                    if (G1.Rows.Count > 0)
                    {
                        oPURCH.CardCode = G2.Rows[0][0].ToString();
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
                            string f1 = GetMenu.DLast3();
                            DateTime S1 = Convert.ToDateTime(f1);
                            oPURCH.Lines.ShipDate = S1;
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
                            string LOCAL = "";
                            if (CURRENCY == "USD")
                            {
                                LOCAL = "C";
                            }
                            else
                            {
                                LOCAL = "L";
                            }

                            UPDATEOPOR(OWTR, LOCAL);
                            for (int j = 0; j < ID.Split(',').Length; j++) 
                            {
                                UPDATEAP_KIT(OWTR, ID.Split(',')[j]);
                            }

                            this.aP_KITTableAdapter.FillBy(this.lC.AP_KIT, textPO1.Text, textCARDNAME.Text, textITEMCODE.Text, textSA.Text, textCARDNAME2.Text, STATUS, textBoxITEMNAME.Text);//更新
                        }
                    }

                }

              
            }






            else
            {
                MessageBox.Show(oCompany.GetLastErrorDescription());

            }
        }
                private void UPDATEAP_KIT(string DOCENTRY, string ID)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE AP_KIT SET PO=@PO WHERE ID = @ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            ID = ID.Substring(1, ID.Length - 2);
            command.Parameters.Add(new SqlParameter("@PO", DOCENTRY));
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
        private void UPDATEOPOR(string DOCENTRY, string CurSource)
        {

            SqlConnection connection = new SqlConnection(strCn98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OPOR SET CurSource=@CurSource WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@CurSource", CurSource));

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

        private void UPDATEOPOR(string DOCENTRY, string LINENUM, string U_MEMO, string U_Shipping_no)
        {

            SqlConnection connection = new SqlConnection(strCn98);
            //   SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET U_MEMO=@U_MEMO,U_Shipping_no=@U_Shipping_no WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@U_Shipping_no", U_Shipping_no));
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
        public void DELOPOR(string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_KITSAP WHERE USERS=@USERS ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        private void button15_Click(object sender, EventArgs e)
        
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result2 = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR(fmLogin.LoginID.ToString());
                GD5(opdf.FileName);
                System.Data.DataTable G1 = GetAPKIT(fmLogin.LoginID.ToString());



                if (G1.Rows.Count > 0)
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
                
                        SAPbobsCOM.Documents oPURCH = null;
                        oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                        string CARDCODE = G1.Rows[0]["CARDCODE"].ToString();
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
                                DateTime CUSTDATE = Convert.ToDateTime(G1.Rows[s]["CUSTDATE"]);

                                oPURCH.Lines.WarehouseCode = "TW001";
                                oPURCH.Lines.ItemCode = ITEMCODE;
                                oPURCH.Lines.Quantity = QTY;
                                oPURCH.Lines.Price = PRICE;
                                oPURCH.Lines.VatGroup = "AP5%";
                                oPURCH.Lines.CostingCode = "11111";
                                oPURCH.Lines.Currency = CURRENCY;

                                oPURCH.Lines.ShipDate = CUSTDATE;
                                
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
                                string LOCAL = "";
                                if (CURRENCY == "USD")
                                {
                                    LOCAL = "C";
                                }
                                else
                                {
                                    LOCAL = "L";
                                }

                                UPDATEOPOR(OWTR, LOCAL);

                            }
                        }

                    }
                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }

                }






             
            }
        }

        public void ADDOPOR(string CARDCODE, string ITEMCODE, decimal QTY, string CURRENCY, decimal PRICE, string MEMO, DateTime CUSTDATE, string USERS)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_KITSAP(CARDCODE,ITEMCODE,QTY,CURRENCY,PRICE,MEMO,CUSTDATE,USERS) values(@CARDCODE,@ITEMCODE,@QTY,@CURRENCY,@PRICE,@MEMO,@CUSTDATE,@USERS)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CURRENCY", CURRENCY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@CUSTDATE", CUSTDATE));
            command.Parameters.Add(new SqlParameter("@USERS", USERS));

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
        private void GD5(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);



            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string ITEMCODE;
                string CARDCODE;
                string CARDNAME;
                string CURRENCY;
                string MEMO;
                decimal QTY;
                decimal PRICE;
                DateTime CUSTDATE;
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                CARDNAME = range.Text.ToString().Trim();
                System.Data.DataTable GCARDCODE = GETCARCODE(CARDNAME);
                if (GCARDCODE.Rows.Count > 0)
                {
                    CARDCODE = GCARDCODE.Rows[0][0].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    if (ITEMCODE != "")
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                        range.Select();
                        QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                        range.Select();
                        CURRENCY = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                        range.Select();
                        PRICE = Convert.ToDecimal(range.Text.ToString().Trim());

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                        range.Select();
                        MEMO = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                        range.Select();
                        CUSTDATE = Convert.ToDateTime(range.Text.ToString().Trim());



                        try
                        {
                            if (!String.IsNullOrEmpty(ITEMCODE))
                            {



                                ADDOPOR(CARDCODE, ITEMCODE, QTY, CURRENCY, PRICE, MEMO, CUSTDATE, fmLogin.LoginID.ToString());
                            }


                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("沒有此客戶");
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
                    return;
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

        private void button16_Click(object sender, EventArgs e)
        {

            System.Data.DataTable G1 = null;

            if (tabControl1.SelectedIndex == 1)
            {
                if (aP_KIT2DataGridView.SelectedRows.Count == 0)
                {

                    MessageBox.Show("請點選單號");
                    return;
                }
                else
                {


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
                }

            }
            if (tabControl1.SelectedIndex == 1)
            {
                G1 = GetTOSAP2(q2);
            }
            if (G1.Rows.Count > 0)
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

                //oCompany.UserName = "manager";
                //oCompany.Password = "19571215";
                int result = oCompany.Connect();
                if (result == 0)
                {


                    //  ProductionOrders oProductionOrders = (ProductionOrders)B1Connections.diCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                    SAPbobsCOM.ProductionOrders oPROD = null;
                    oPROD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                    string CARDNAME = G1.Rows[0]["CARDNAME"].ToString();
                    string CARDNAME2 = G1.Rows[0]["CARDNAME2"].ToString();
                    string ITEMCODE = G1.Rows[0]["ITEMCODE"].ToString();
                    int QTY = Convert.ToInt16(G1.Rows[0]["PQTY"]);
                    int SO = Convert.ToInt32(G1.Rows[0]["SO"]);
                    DateTime ENDDATE = Convert.ToDateTime(G1.Rows[0]["ENDDATE"]);
      
                    System.Data.DataTable G2 = GETCARCODE2(CARDNAME);

                    if (G1.Rows.Count > 0)
                    {

                        oPROD.UserFields.Fields.Item("U_CARDNAME").Value = G2.Rows[0][0].ToString();
                        oPROD.ItemNo = ITEMCODE;
                        oPROD.PlannedQuantity = QTY;
                        oPROD.DueDate = ENDDATE;
                        //RLAR215AB.XXXX1
                   //     oPROD.CustomerCode = G3.Rows[0][0].ToString();
                        oPROD.ProductionOrderOriginEntry = 741;
                
                        int res = oPROD.Add();
                        if (res != 0)
                        {
                            MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                        }
                        else
                        {
                            System.Data.DataTable G4 = GetDI5();
                            string OWTR = G4.Rows[0][0].ToString();
                            MessageBox.Show("上傳成功 生產單號 : " + OWTR);

                        }
                    }

                }




                else
                {
                    MessageBox.Show(oCompany.GetLastErrorDescription());

                }
            }
        }
    }
}

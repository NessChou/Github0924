using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Web.UI;
namespace ACME
{
    public partial class SQUT2 : ACME.fmBase1
    {
        public string PublicString;
        int COPY = 0;
        public SQUT2()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            shipping_OQUTTableAdapter.Connection = MyConnection;
            shipping_OQUT1TableAdapter.Connection = MyConnection;
            shipping_OQUTDownloadTableAdapter.Connection = MyConnection;
            shipping_OQUTDownload2TableAdapter.Connection = MyConnection;

        }
        public override void AfterCopy()
        {

            if (kyes == null)
            {
                string NumberName = "SQ" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                kyes = this.shippingCodeTextBox.Text;
            }
        }
        public override void AfterCopy2()
        {
            COPY = 1;
            dOCDATETextBox.Text = "";
            eNDDATETextBox.Text = "";
            createNameTextBox.Text = fmLogin.LoginID.ToString();
            mEMOTextBox.Text = "";
        }
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
          
            iTEMCODETextBox.ReadOnly = true;
            iTEMNAMETextBox.ReadOnly = true;

            button1.Enabled = true;
            button3.Enabled = true;
            tERMTextBox.ReadOnly = true;
            button7.Enabled = true;
            sHIPWAYTextBox.ReadOnly = true;
          groupBox1.Visible = false;
        }
        public override void query()
        {
            sHIPWAYTextBox.ReadOnly = true;
            tERMTextBox.ReadOnly = true;
            groupBox1.Visible = true;
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                ship2.Shipping_OQUT.RejectChanges();
                ship2.Shipping_OQUT1.RejectChanges();
                ship2.Shipping_OQUTDownload.RejectChanges();
                ship2.Shipping_OQUTDownload2.RejectChanges();

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
        public override void AfterEdit()
        {
            shippingCodeTextBox.ReadOnly = true;
       

            iTEMCODETextBox.ReadOnly = true;
            iTEMNAMETextBox.ReadOnly = true;

            sHIPWAYTextBox.ReadOnly = true;
        }
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = shipping_OQUTBindingSource;
            MyTableName = "shipping_OQUT";
            MyIDFieldName = "ShippingCode";



            //處理複製
            MasterTable = ship2.Shipping_OQUT;

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "SQ" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;


            string username = fmLogin.LoginID.ToString();


            createNameTextBox.Text = username;
            this.shipping_OQUTBindingSource.EndEdit();
            kyes = null;


        }

        public override void FillData()
        {
            try
            {
                if (!String.IsNullOrEmpty(PublicString))
                {
                    MyID = PublicString.Trim();

                }
                
                shipping_OQUTTableAdapter.Fill(ship2.Shipping_OQUT, MyID);
                shipping_OQUT1TableAdapter.Fill(ship2.Shipping_OQUT1, MyID);
                shipping_OQUTDownloadTableAdapter.Fill(ship2.Shipping_OQUTDownload, MyID);
                shipping_OQUTDownload2TableAdapter.Fill(ship2.Shipping_OQUTDownload2, MyID);

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

                //shipping_OQUTDownloadBindingSource.MoveFirst();

                //for (int i = 0; i <= shipping_OQUTDownloadBindingSource.Count - 1; i++)
                //{
                //    DataRowView row1 = (DataRowView)shipping_OQUTDownloadBindingSource.Current;

                //    row1["seq"] = i;

                //    shipping_OQUTDownloadBindingSource.EndEdit();

                //    shipping_OQUTDownloadBindingSource.MoveNext();
                //}

                //for (int i = 0; i <= shipping_OQUTDownload2BindingSource.Count - 1; i++)
                //{
                //    DataRowView row1 = (DataRowView)shipping_OQUTDownload2BindingSource.Current;

                //    row1["seq"] = i;

                //    shipping_OQUTDownload2BindingSource.EndEdit();

                //    shipping_OQUTDownload2BindingSource.MoveNext();
                //}

                S1();
                Validate();


                shipping_OQUTTableAdapter.Connection.Open();


                shipping_OQUTBindingSource.EndEdit();
                shipping_OQUT1BindingSource.EndEdit();
                shipping_OQUTDownloadBindingSource.EndEdit();
                shipping_OQUTDownload2BindingSource.EndEdit();


                tx = shipping_OQUTTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(shipping_OQUTTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter1 = util.GetAdapter(shipping_OQUT1TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(shipping_OQUTDownloadTableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter3 = util.GetAdapter(shipping_OQUTDownload2TableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;

                

                shipping_OQUTTableAdapter.Update(ship2.Shipping_OQUT);
                ship2.Shipping_OQUT.AcceptChanges();

                shipping_OQUT1TableAdapter.Update(ship2.Shipping_OQUT1);
                ship2.Shipping_OQUT1.AcceptChanges();

                shipping_OQUTDownloadTableAdapter.Update(ship2.Shipping_OQUTDownload);
                ship2.Shipping_OQUTDownload.AcceptChanges();

                shipping_OQUTDownload2TableAdapter.Update(ship2.Shipping_OQUTDownload2);
                ship2.Shipping_OQUTDownload2.AcceptChanges();


           

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
                this.shipping_OQUTTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        private void  S1()
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            for (int i = 0; i <= shipping_OQUTDownload2DataGridView.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = shipping_OQUTDownload2DataGridView.Rows[i];
                sb.Append(row.Cells["CARDCODE"].Value.ToString() + "/");
                sb2.Append(row.Cells["CARDNAME"].Value.ToString() + "/");
            }

            if (sb2.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);
                sb2.Remove(sb2.Length - 1, 1);
                cARDCODE2TextBox.Text = sb.ToString();
                cARDANME2TextBox.Text = sb2.ToString();
            }
            else
            {
                cARDCODE2TextBox.Text = "";
                cARDANME2TextBox.Text = "";
            }
        }




        private void button1_Click(object sender, EventArgs e)
        {
            SQUTCARD frm1 = new SQUTCARD();
            //frm1.cardcode = cardCodeTextBox.Text;
            //frm1.usd = bankCodeTextBox.Text;
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                //    UtilSimple.SetLookupBinding(comboBox3, GetOCRD3(), "DataValue", "DataValue");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GETOUITEM();

            if (LookupValues != null)
            {
                iTEMCODETextBox.Text = Convert.ToString(LookupValues[0]);
                iTEMNAMETextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (sHIPWAYTextBox.Text.Trim() == "")
            {
                MessageBox.Show("請選擇運輸方式");
                return;
            }

            object[] LookupValues = null;

            string SHIP = sHIPWAYTextBox.Text.Substring(0, 1);

            LookupValues = GetMenu.GETOUITEM2(SHIP);

            if (LookupValues != null)
            {
                iTEMCODETextBox.Text = Convert.ToString(LookupValues[0]);
                iTEMNAMETextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SQUTOITM frm1 = new SQUTOITM();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("收件人地址為" + textBox2.Text + "是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {



                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\船務報價.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);


                StringBuilder sb = new StringBuilder();
                System.Data.DataTable dt = GetIV();
 
                int T1 = tabControl1.SelectedIndex;
                tabControl1.SelectedIndex = 3;
                tabControl1.SelectedIndex = T1;
                string A1="";

                template = template.Replace("##SQUT##", sQTYTextBox.Text);
             
                template = template.Replace("##SQUT2##", sITEMCODETextBox.Text.Replace(System.Environment.NewLine, "<br>"));
                MailMessage message = new MailMessage();


                message.To.Add(new MailAddress(textBox2.Text));
                //詢價：空運費-XMN-KHH_ SQ20130926002X
                message.Subject = "詢價：" + iTEMNAMETextBox.Text.Trim() + "_" + shippingCodeTextBox.Text.Trim();
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

        public System.Data.DataTable GetIV()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT QTYPE FROM shipping_OQUT1 WHERE SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
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

        private void button25_Click(object sender, EventArgs e)
        {

            try
            {
                string server = "//acmesrv01//SAP_Share//shipping/QUOTATION DOCUMENT//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.OQUTDownload(filename);

                if (dt2.Rows.Count > 0)
                {
                    MessageBox.Show("檔案名稱重複,請修改檔名");
                }
                else
                {
                    if (result == DialogResult.OK)
                    {

                        string file = opdf.FileName;
                        bool FF1 = getrma.UploadFile(file, server, false);
                        if (FF1 == false)
                        {
                            return;
                        }
                        System.Data.DataTable dt1 = ship2.Shipping_OQUTDownload;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (shipping_OQUTDownloadDataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = @"\\acmesrv01\SAP_Share\shipping\QUOTATION DOCUMENT\" + filename;
                        dt1.Rows.Add(drw);

                        shipping_OQUTDownloadBindingSource.MoveFirst();

                        for (int i = 0; i <= shipping_OQUTDownloadBindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)shipping_OQUTDownloadBindingSource.Current;

                            rowd["seq"] = i;



                            shipping_OQUTDownloadBindingSource.EndEdit();

                            shipping_OQUTDownloadBindingSource.MoveNext();
                        }

                        this.shipping_OQUTDownloadBindingSource.EndEdit();
                        this.shipping_OQUTDownloadTableAdapter.Update(ship2.Shipping_OQUTDownload);
                        ship2.Shipping_OQUTDownload.AcceptChanges();

                        MessageBox.Show("上傳成功");
                    }





                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {
            tERMTextBox.Text = comboBox5.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string ID="";

            if (shipping_OQUTDownload2DataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇上傳列");
                return;
            }

            if (shipping_OQUTDownload2DataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow row;
                StringBuilder sb = new StringBuilder();
                row = shipping_OQUTDownload2DataGridView.SelectedRows[0];
                ID = row.Cells["ID"].Value.ToString();
                
            }

            try
            {
                string server = "//acmesrv01//SAP_Share//shipping/QUOTATION DOCUMENT//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.OQUTDownload2(filename);

                if (dt2.Rows.Count > 0)
                {
                    MessageBox.Show("檔案名稱重複,請修改檔名");
                }
                else
                {
                    if (result == DialogResult.OK)
                    {

                        string file = opdf.FileName;
                        bool FF1 = getrma.UploadFile(file, server, false);
                        if (FF1 == false)
                        {
                            return;
                        }
     
                       string PATH = @"\\acmesrv01\SAP_Share\shipping\QUOTATION DOCUMENT\" + filename;

                       UpdateFILENAME(ID, filename, PATH);

                       shipping_OQUTDownload2TableAdapter.Fill(ship2.Shipping_OQUTDownload2, MyID);

                        MessageBox.Show("上傳成功");
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateFILENAME(string ID, string FILENAME, string PATH)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update shipping_OQUTDownload2 set FILENAME=@FILENAME,PATH=@PATH where ID=@ID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
            command.Parameters.Add(new SqlParameter("@PATH", PATH));
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


        private void button10_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GETUIN();

            if (LookupValues != null)
            {
                System.Data.DataTable dt1 = ship2.Shipping_OQUTDownload2;

                DataRow drw = dt1.NewRow();
                drw["ShippingCode"] = shippingCodeTextBox.Text;
                drw["seq"] = (shipping_OQUTDownload2DataGridView.Rows.Count).ToString();
                drw["CARDCODE"] = Convert.ToString(LookupValues[0]);
                drw["CARDNAME"] = Convert.ToString(LookupValues[1]);
                dt1.Rows.Add(drw);

                shipping_OQUTDownload2BindingSource.MoveFirst();

                for (int i = 0; i <= shipping_OQUTDownload2BindingSource.Count - 1; i++)
                {
                    DataRowView rowd = (DataRowView)shipping_OQUTDownload2BindingSource.Current;

                    rowd["seq"] = i;

                    shipping_OQUTDownload2BindingSource.EndEdit();

                    shipping_OQUTDownload2BindingSource.MoveNext();
                }

                this.shipping_OQUTDownload2BindingSource.EndEdit();
                this.shipping_OQUTDownload2TableAdapter.Update(ship2.Shipping_OQUTDownload2);
                ship2.Shipping_OQUTDownload2.AcceptChanges();

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetU();

            if (LookupValues != null)
            {
                System.Data.DataTable dt1 = ship2.Shipping_OQUTDownload2;

                DataRow drw = dt1.NewRow();
                drw["ShippingCode"] = shippingCodeTextBox.Text;
                drw["seq"] = (shipping_OQUTDownload2DataGridView.Rows.Count).ToString();
                drw["CARDCODE"] = Convert.ToString(LookupValues[0]);
                drw["CARDNAME"] = Convert.ToString(LookupValues[1]);
                dt1.Rows.Add(drw);

                shipping_OQUTDownload2BindingSource.MoveFirst();

                for (int i = 0; i <= shipping_OQUTDownloadBindingSource.Count - 1; i++)
                {
                    DataRowView rowd = (DataRowView)shipping_OQUTDownload2BindingSource.Current;

                    rowd["seq"] = i;



                    shipping_OQUTDownload2BindingSource.EndEdit();

                    shipping_OQUTDownload2BindingSource.MoveNext();
                }

                this.shipping_OQUTDownload2BindingSource.EndEdit();
                this.shipping_OQUTDownload2TableAdapter.Update(ship2.Shipping_OQUTDownload2);
                ship2.Shipping_OQUTDownload2.AcceptChanges();

            }
        }

        private void shipping_OQUTDownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {

                    System.Data.DataTable dt1 = ship2.Shipping_OQUTDownload;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];

                    string aa = drw["path"].ToString();


                    System.Diagnostics.Process.Start(aa);
                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;


                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void shipping_OQUTDownload2DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK2")
                {

                    System.Data.DataTable dt1 = ship2.Shipping_OQUTDownload2;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];

                    string aa = drw["path"].ToString();


                    System.Diagnostics.Process.Start(aa);
                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;


                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

      

        private void SQUT2_Load(object sender, EventArgs e)
        {
            WW();
            textBox2.Text = fmLogin.LoginID.ToString() + "@acmepoint.com";
        }

        private void comboBox5_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SQUTTERM");

            comboBox5.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox5.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox6_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SHIPWAY");

            comboBox6.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox6.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

       

        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            sHIPWAYTextBox.Text = comboBox6.Text;
        }

        private void shipping_OQUT1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (shipping_OQUT1DataGridView.Columns[e.ColumnIndex].Name == "QTYPE")
                {
                    StringBuilder sb = new StringBuilder();
                    System.Data.DataTable dt = GetIV();
                    for (int i = 0; i <= shipping_OQUT1DataGridView.Rows.Count - 2; i++)
                    {

                        DataGridViewRow row;

                        row = shipping_OQUT1DataGridView.Rows[i];
                        sb.Append(row.Cells["QTYPE"].Value.ToString() + "/");
                    }



                    sb.Remove(sb.Length - 1, 1);
                    int T1 = tabControl1.SelectedIndex;
                    tabControl1.SelectedIndex = 3;
                    tabControl1.SelectedIndex = T1;
                    sQTYTextBox.Text = "請同步報" + sb.ToString();


                }

            }
            catch { }
        }


        private void sITEMCODETextBox_Click(object sender, EventArgs e)
        {
            if (sITEMCODETextBox.Text == "")
            {
                sITEMCODETextBox.Text = "品名：" +
                          Environment.NewLine + "" +
    Environment.NewLine + "數量：" +
      Environment.NewLine + "" +
    Environment.NewLine + "貨物尺寸：" +
        Environment.NewLine + "" +
    Environment.NewLine + "貨物重量：" +
        Environment.NewLine + "" +
    Environment.NewLine + "貿易條件：" +
        Environment.NewLine + "" +
    Environment.NewLine + "派送地址：" +
        Environment.NewLine + "" +
    Environment.NewLine + "取件地址：";
            }
        }

        private void shipping_OQUTDownload2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          
        }




   

 
    }
}

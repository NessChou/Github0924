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
namespace ACME
{
    public partial class SOLARPAY : ACME.fmBase1
    {
        public SOLARPAY()
        {
            InitializeComponent();
        }
        public string PublicString2;
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sOLAR_PAYTableAdapter.Connection = MyConnection;
            sOLAR_PAYDownloadTableAdapter.Connection = MyConnection;
        }

  
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            cREATENAMETextBox.ReadOnly = true;

            button2.Enabled = true;
            button3.Enabled = true;
            dOCTYPETextBox.ReadOnly = true;
            t1TextBox.ReadOnly = true;
            t2TextBox.ReadOnly = true;
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            dOCDATETextBox.ReadOnly = false;
            cREATENAMETextBox.ReadOnly = false;
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                sOLAR.SOLAR_PAY.RejectChanges();
                sOLAR.SOLAR_PAYDownload.RejectChanges();
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
            dOCDATETextBox.ReadOnly = true;
            cREATENAMETextBox.ReadOnly = true;
            dOCTYPETextBox.ReadOnly = true;
            t1TextBox.ReadOnly = true;
            t2TextBox.ReadOnly = true;
        }
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = sOLAR_PAYBindingSource;
            MyTableName = "SOLAR_PAY";
            MyIDFieldName = "ShippingCode";

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "SP" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            string username = fmLogin.LoginID.ToString();
            cREATENAMETextBox.Text = username;
            dOCDATETextBox.Text = GetMenu.Day();
            t2TextBox.Text = "客戶";
            dOCTYPETextBox.Text = "採購請款";
            t1TextBox.Text = "T/T";
            this.sOLAR_PAYBindingSource.EndEdit();
            kyes = null;
        }
        public override void FillData()
        {
            if (!String.IsNullOrEmpty(PublicString2))
            {
                MyID = PublicString2;

            }
                sOLAR_PAYTableAdapter.Fill(sOLAR.SOLAR_PAY, MyID);
                sOLAR_PAYDownloadTableAdapter.Fill(sOLAR.SOLAR_PAYDownload, MyID);
        }

        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {
                sOLAR_PAYDownloadBindingSource.MoveFirst();

                for (int i = 0; i <= sOLAR_PAYDownloadBindingSource.Count - 1; i++)
                {
                    DataRowView row1 = (DataRowView)sOLAR_PAYDownloadBindingSource.Current;

                    row1["seq"] = i;

                    sOLAR_PAYDownloadBindingSource.EndEdit();

                    sOLAR_PAYDownloadBindingSource.MoveNext();
                }

                Validate();


                sOLAR_PAYTableAdapter.Connection.Open();


                sOLAR_PAYBindingSource.EndEdit();



                tx = sOLAR_PAYTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(sOLAR_PAYTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;


                SqlDataAdapter Adapter1 = util.GetAdapter(sOLAR_PAYDownloadTableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;


                sOLAR_PAYTableAdapter.Update(sOLAR.SOLAR_PAY);
                sOLAR.SOLAR_PAY.AcceptChanges();

                sOLAR_PAYDownloadTableAdapter.Update(sOLAR.SOLAR_PAYDownload);
                sOLAR.SOLAR_PAYDownload.AcceptChanges();




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
                this.sOLAR_PAYTableAdapter.Connection.Close();

            }
            return UpdateData;
        }



        private void button1_Click(object sender, EventArgs e)
        {
            if (oPDOCTextBox.Text == "")
            {
                MessageBox.Show("請輸入採購單號");
                return;
            }
            object[] LookupValues = GetMenu.GetOPOR(oPDOCTextBox.Text);

            if (LookupValues != null)
            {
                cARDCODETextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAMETextBox.Text = Convert.ToString(LookupValues[1]);
                oPDATETextBox.Text = Convert.ToString(LookupValues[2]);
                oPITEMTextBox.Text = Convert.ToString(LookupValues[3]);
                oPQTYTextBox.Text = Convert.ToString(LookupValues[4]);
                oPPRICETextBox.Text = Convert.ToString(LookupValues[5]);
                oPAMTTextBox.Text = Convert.ToString(LookupValues[6]);
                pRJIDTextBox.Text = Convert.ToString(LookupValues[7]);
                pRJNAMETextBox.Text = Convert.ToString(LookupValues[8]);
                oPPAYTextBox.Text = Convert.ToString(LookupValues[9]);
                sOLAR_PAYBindingSource.EndEdit();

                System.Data.DataTable H1 = GetACCOUNT(cARDCODETextBox.Text);             
                if (H1.Rows.Count > 0)
                {
                    aCCOUNTNAMETextBox.Text = H1.Rows[0]["BANKNAME"].ToString();
                    aCCOUNTTextBox.Text = H1.Rows[0]["ACCOUNT"].ToString();
                    aCCOUNTCODETextBox.Text = H1.Rows[0]["BANKCODE"].ToString();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
     

                System.Data.DataTable DT1 = DT();

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                if (dOCTYPETextBox.Text == "採購請款")
                {
                    if (globals.DBNAME == "進金生能源服務")
                    {
                        FileName = lsAppDir + "\\Excel\\ENERGY\\採購請款.xls";
                    }
                    else
                    {
                        FileName = lsAppDir + "\\Excel\\SOLAR\\採購請款.xls";
                    }

                    
                }
                if (dOCTYPETextBox.Text == "預付請款")
                {
                    if (globals.DBNAME == "進金生能源服務")
                    {
                        FileName = lsAppDir + "\\Excel\\ENERGY\\預付請款.xls";
                    }
                    else
                    {
                        FileName = lsAppDir + "\\Excel\\SOLAR\\預付請款.xls";
                    }

    
                }


                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            
        }
            
        private System.Data.DataTable DT()
        {

            string K1 = "";
            string K2 = "";
            if (t2TextBox.Text == "客戶")
            {
                K1 = "匯費負擔: ■客戶  □進金生;    幣別:   NT         匯率:";
            }
            else if (t2TextBox.Text == "進金生")
            {
                K1 = "匯費負擔: □客戶  ■進金生;    幣別:   NT         匯率:";
            }

            if (t1TextBox.Text == "票據")
            {
                K2 = "■ 票據，到期日: "+t1DATETextBox.Text.ToString().Trim()+"      □ T/T 付款日: ____       □ 依公司規定      □ 現  金";
            }
            else if (t1TextBox.Text == "T/T")
            {
                K2 = "□ 票據，到期日:       ■ T/T 付款日: " + t1DATETextBox.Text.ToString().Trim() + "       □ 依公司規定      □ 現  金";
            }
            else if (t1TextBox.Text == "依公司規定")
            {
                K2 = "□ 票據，到期日:       □ T/T 付款日: ____       ■ 依公司規定      □ 現  金";
            }
            else if (t1TextBox.Text == "現金")
            {
                K2 = "□ 票據，到期日:       □ T/T 付款日: ____       □ 依公司規定      ■ 現  金";
            }

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();



            sb.Append("    SELECT CARDCODE 廠商編號,CARDNAME 受款人名稱,CONVERT(VARCHAR(10) ,CAST(DOCDATE AS DATETIME), 111 )  填表日,OPDATE 採購日期,");
            sb.Append("             OPDOC 採購單號,OPITEM 採購項目,OPQTY 採購數量,OPPRICE 採購單價,OPAMT  採購金額,AMT 請款金額,PRJID 專案號碼,PRJNAME 專案名稱,");
            sb.Append(" @K1 AS 貸,@K2 AS 借,");
            sb.Append(" OPPAY 付款條件,MEMO 付款說明,DUEDATE 預計到貨日,''''+ACCOUNT 銀行帳號,ACCOUNTNAME 銀行名稱,''''+ACCOUNTCODE 銀行代碼,UNIT 需求單位,INVOICE 發票號碼,'支付單號: '+SHIPPINGCODE  支付單號,DOCTYPE 單據類別 FROM SOLAR_PAY ");
            sb.Append(" WHERE SHIPPINGCODE = @SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@K1", K1));
            command.Parameters.Add(new SqlParameter("@K2", K2));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void SOLARPAY_Load(object sender, EventArgs e)
        {
            WW();

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetPAY();
            ExcelReport.GridViewToExcelPotato(dataGridView1);
        }
        private System.Data.DataTable GetPAY()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT OPDATE 採購日期,CARDNAME  廠商名稱,OPITEM 採購項,OPQTY 採購數量,OPPRICE 採購單價,OPAMT 採購金額,t1DATE 付款日期,OPPAY 付款條件");
            sb.Append(" ,DOCDATE 請購日期,CREATENAME 請款人員,MEMO 付款說明");
            sb.Append("  FROM SOLAR_PAY  WHERE DOCDATE BETWEEN @DOCDATE1 AND @DOCDATE2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private System.Data.DataTable GetACCOUNT(string ACCOUNT)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BANKNAME+DFLBRANCH BANKNAME,DFLACCOUNT ACCOUNT,T0.BANKCODE FROM OCRD  T0 ");
            sb.Append(" LEFT JOIN ODSC T1 ON (T0.BANKCODE =T1.BANKCODE)  ");
            sb.Append(" where CARDCODE=@CARDCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", ACCOUNT));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }


        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SPAY1");

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void dOCTYPETextBox_TextChanged(object sender, EventArgs e)
        {
            if (dOCTYPETextBox.Text == "採購請款")
            {
                oPDOCTextBox.Visible = true;
                oPDATETextBox.Visible = true;
                iNVOICETextBox.Visible = false;
                uNITTextBox.Visible = false;

                label3.Text = "採購單號";
                label4.Text = "採購日期";
    
                groupBox6.Visible = false;
          
          
          
                button1.Visible = true;

            }
            if (dOCTYPETextBox.Text == "預付請款")
            {
                oPDOCTextBox.Visible = false;
                oPDATETextBox.Visible = false;
                iNVOICETextBox.Visible = true;
                uNITTextBox.Visible = true;

                label3.Text = "發票號碼";
                label4.Text = "需求單位";
        
                groupBox6.Visible = true;
           
                button1.Visible = false;
   
            }


        }

        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SPAY2");

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }


        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SPAY3");

            comboBox3.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            t2TextBox.Text = comboBox3.Text;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            t1TextBox.Text = comboBox2.Text;
        }

        private void button25_Click(object sender, EventArgs e)
        {

            string[] filebType = Directory.GetDirectories("//ACMEW08R2AP//SAPFILES//AttachmentsSolar2001//");
            string dd = DateTime.Now.ToString("yyyyMM");

            try
            {
                string server = "//ACMEW08R2AP//SAPFILES//AttachmentsSolar2001///";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download2(filename);

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
                        System.Data.DataTable dt1 = sOLAR.SOLAR_PAYDownload;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (sOLAR_PAYDownloadDataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = @"\\ACMEW08R2AP\SAPFILES\AttachmentsSolar2001\" + filename;
                        dt1.Rows.Add(drw);

                        sOLAR_PAYDownloadBindingSource.MoveFirst();

                        for (int i = 0; i <= sOLAR_PAYDownloadBindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)sOLAR_PAYDownloadBindingSource.Current;

                            rowd["seq"] = i;



                            sOLAR_PAYDownloadBindingSource.EndEdit();

                            sOLAR_PAYDownloadBindingSource.MoveNext();
                        }

                        this.sOLAR_PAYDownloadBindingSource.EndEdit();
                        this.sOLAR_PAYDownloadTableAdapter.Update(sOLAR.SOLAR_PAYDownload);
                        sOLAR.SOLAR_PAYDownload.AcceptChanges();

                        MessageBox.Show("上傳成功");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DELETEFILE(string aa)
        {
            string server = "//ACMEW08R2AP//SAPFILES//AttachmentsSolar2001//";
            string[] filenames = Directory.GetFiles(server);
            foreach (string file in filenames)
            {

                FileInfo filess = new FileInfo(file);
                string fd = filess.Name.ToString();
                if (fd == aa)
                {
                    File.Delete(file);
                }
            }
        }

        private void sOLAR_PAYDownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {

                    System.Data.DataTable dt1 = sOLAR.SOLAR_PAYDownload;
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

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            dOCTYPETextBox.Text = comboBox1.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            LookupValues = GetMenu.GetMenuListU2();


            if (LookupValues != null)
            {
                cARDCODETextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAMETextBox.Text = Convert.ToString(LookupValues[1]);
                aCCOUNTNAMETextBox.Text = Convert.ToString(LookupValues[2]);
                aCCOUNTTextBox.Text = Convert.ToString(LookupValues[3]);
                aCCOUNTCODETextBox.Text = Convert.ToString(LookupValues[4]);
                oPPAYTextBox.Text = Convert.ToString(LookupValues[5]);
            }
        }


    
        }
    }


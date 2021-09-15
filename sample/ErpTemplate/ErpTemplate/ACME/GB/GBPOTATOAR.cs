using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
namespace ACME
{
    public partial class GBPOTATOAR : Form
    {
        public GBPOTATOAR()
        {
            InitializeComponent();
        }

        private void gB_POTATOARBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_POTATOARBindingSource.EndEdit();
            this.gB_POTATOARTableAdapter.Update(this.accBank.GB_POTATOAR);

            MessageBox.Show("存檔成功");

        }

        private void GBPOTATOAR_Load(object sender, EventArgs e)
        {
            toolStripTextBox1.Text = DateToStr(DateTime.Today.AddMonths(-1)).Substring(0, 6);
            toolStripComboBox1.Text = "聿豐";
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            gB_POTATOARTableAdapter.FillBy(this.accBank.GB_POTATOAR, toolStripTextBox1.Text, toolStripComboBox1.Text);

            GETSUM();
        }

        

        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public System.Data.DataTable GetCHO3()
        {
           if (toolStripComboBox1.Text == "韋峰")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            //韋峰
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            if (toolStripComboBox1.Text == "忠孝")
            {
                sb.Append(" SELECT  VoucherNo TRANSID,DEPARTID BU,CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,4) DOCDATE,Summary MEMO,CAST(Amount AS INT) B,CAST(Amount/0.05 AS INT) A,CAST(Amount/0.05+Amount AS INT) C from AccVoucherSub Where CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,2)  = @DocDate1  AND SUBSTRING(VOUCHERNO,1,1) <> 'K' and SubjectID ='1281000' and Amount <> 0  AND DepartID  IN ('C2','F1')");
            }
            else if (toolStripComboBox1.Text == "韋峰")
            {
                sb.Append(" SELECT  T0.VoucherNo TRANSID,DEPARTID BU,CAST(T1.MakeDate AS VARCHAR)  DOCDATE,Summary MEMO,CAST(TempAmount AS INT) B");
                sb.Append(" ,CAST(TempAmount/0.05 AS INT) A,CAST(TempAmount/0.05+TempAmount AS INT) C");
                sb.Append("  from AccVoucherSub T0 LEFT JOIN AccVoucherMAIN T1 ON (T0.VoucherNo =T1.VoucherNo)");
                sb.Append("   Where SUBSTRING(CAST(MakeDate AS VARCHAR),1,6)  = @DocDate1   and SubjectID ='1281001' and TempAmount <> 0 ");

            }
            else
            {
                sb.Append(" SELECT  VoucherNo TRANSID,DEPARTID BU,CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,4) DOCDATE,Summary MEMO,CAST(Amount AS INT) B,CAST(Amount/0.05 AS INT) A,CAST(Amount/0.05+Amount AS INT) C from AccVoucherSub Where CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,2)  = @DocDate1  AND SUBSTRING(VOUCHERNO,1,1) <> 'K' and SubjectID ='1281000' and Amount <> 0  AND DepartID  IN ('C1','F2','AB1')");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", toolStripTextBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetINV()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND COMPANY=@COMPANY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", toolStripComboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public System.Data.DataTable GetTSUM()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '0-三聯式發票/電子計算機發票' 類別,SUM(U_IN_BSAMN) 未稅金額,SUM(U_IN_BSTAX) 稅額,SUM(U_IN_BSAMT) 含稅總額 FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='0' AND COMPANY=@COMPANY ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '1-三聯式收銀機發票/電子發票',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='1' AND COMPANY=@COMPANY");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2-有稅憑證',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='2' AND COMPANY=@COMPANY");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '3-海關代徵稅',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='3' AND COMPANY=@COMPANY");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '4-免用統一發票/收據',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='4' AND COMPANY=@COMPANY");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '5-退折',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='5'  AND COMPANY=@COMPANY");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '合計',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1) IN ('0','1','2','3','4','5')  AND COMPANY=@COMPANY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", toolStripComboBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        private void INSERT(string DelRemark, string TRANSID, string MEMO, string U_IN_BSAMN, string U_IN_BSTAX, string U_IN_BSAMT, string CARDNAME, string ACCOUNT, string U_PC_BSTY1, string U_PC_BSTY4, string SHIPDATE, string DOCDATE, string UNIT, string BU, string COMPANY)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO GB_POTATOAR(DelRemark,TRANSID,MEMO,U_IN_BSAMN,U_IN_BSTAX,U_IN_BSAMT,CARDNAME,ACCOUNT,U_PC_BSTY1,U_PC_BSTY4,SHIPDATE,DOCDATE,UNIT,BU,COMPANY) VALUES(@DelRemark,@TRANSID,@MEMO,@U_IN_BSAMN,@U_IN_BSTAX,@U_IN_BSAMT,@CARDNAME,@ACCOUNT,@U_PC_BSTY1,@U_PC_BSTY4,@SHIPDATE,@DOCDATE,@UNIT,@BU,@COMPANY)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DelRemark", DelRemark));
            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@U_IN_BSAMN", U_IN_BSAMN));
            command.Parameters.Add(new SqlParameter("@U_IN_BSTAX", U_IN_BSTAX));
            command.Parameters.Add(new SqlParameter("@U_IN_BSAMT", U_IN_BSAMT));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY1", U_PC_BSTY1));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY4", U_PC_BSTY4));
            command.Parameters.Add(new SqlParameter("@SHIPDATE", SHIPDATE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@UNIT", UNIT));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            //COMPANY
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
   
        private void TRUNCATE()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" delete   GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND COMPANY=@COMPANY ");
            // sb.Append(" delete   GB_POTATOAR  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND ISNULL(SHIPDATE,'')+ISNULL(UNIT,'')+ISNULL(U_PC_BSTY1,'')+ISNULL(DOCDATE,'')+ISNULL(U_PC_BSTY4,'') <> '' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", toolStripComboBox1.Text));


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
        private void button1_Click(object sender, EventArgs e)
        {

            string ACCOUNT = "";
            string CARDNAME = "";
            string INVOTYPE = "";
            string INVONO = "";
            string INVODATE = "";
            string UNIT = "";

       
                System.Data.DataTable J1 = GetINV();
                if (J1.Rows.Count > 0)
                {

                    DialogResult result;
                    result = MessageBox.Show("舊的資料會清除，請問是否要重新匯入資料?", "YES/NO", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        TRUNCATE();

                    }
                }
            
                System.Data.DataTable T1 = GetCHO3();
                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {
                    string TRANSID = T1.Rows[i]["TRANSID"].ToString();
                    string DOCDATE = T1.Rows[i]["DOCDATE"].ToString();
                    string BU = T1.Rows[i]["BU"].ToString();
                    string MEMO = T1.Rows[i]["MEMO"].ToString();
                    string A = T1.Rows[i]["A"].ToString();
                    string B = T1.Rows[i]["B"].ToString();
                    string C = T1.Rows[i]["C"].ToString();

                    try
                    {
                        string[] s = MEMO.Split(' ');


                         ACCOUNT = Convert.ToString(s[0]);
                         CARDNAME = Convert.ToString(s[1]);
                         INVOTYPE = util.INVOTYPE(Convert.ToString(s[2]));
                         INVONO = Convert.ToString(s[3]);
                         INVODATE = Convert.ToString(s[4]);
                         UNIT = Convert.ToString(s[5]);
                    }
                    catch { }

                    if (INVODATE.Length == 8)
                    {
                        INSERT(DOCDATE, TRANSID, MEMO, A, B, C, CARDNAME, ACCOUNT, INVOTYPE, "1-進貨及費用", INVONO, INVODATE, UNIT, BU, toolStripComboBox1.Text);
                    }
                }

                gB_POTATOARTableAdapter.FillBy(this.accBank.GB_POTATOAR, toolStripTextBox1.Text, toolStripComboBox1.Text);

                GETSUM();
            
        }
        private void GETSUM()
        {
            System.Data.DataTable F1 = GetTSUM();
            dataGridView1.DataSource = F1;
            for (int i = 1; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }
        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {
            gB_POTATOARTableAdapter.FillBy1(this.accBank.GB_POTATOAR, toolStripTextBox2.Text, toolStripComboBox1.Text);

        }

        private void button4_Click(object sender, EventArgs e)
        {
           

            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcelSHARON(gB_POTATOARDataGridView);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
        }

        private void gB_POTATOARDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gB_POTATOARDataGridView.Columns[e.ColumnIndex].Name == "U_IN_BSAMN" ||
                          gB_POTATOARDataGridView.Columns[e.ColumnIndex].Name == "U_IN_BSTAX")
                {

                    Int32 U_IN_BSAMN = 0;
                    Int32 U_IN_BSTAX = 0;

                    U_IN_BSAMN = Convert.ToInt32(this.gB_POTATOARDataGridView.Rows[e.RowIndex].Cells["U_IN_BSAMN"].Value);
                    U_IN_BSTAX = Convert.ToInt32(this.gB_POTATOARDataGridView.Rows[e.RowIndex].Cells["U_IN_BSTAX"].Value);


                    this.gB_POTATOARDataGridView.Rows[e.RowIndex].Cells["U_IN_BSAMT"].Value = (U_IN_BSAMN + U_IN_BSTAX);


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void gB_POTATOARDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
               
            }
        }

        private void gB_POTATOARDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["COMPANY"].Value = toolStripComboBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string ACCOUNT = "";
            string CARDNAME = "";
            string INVOTYPE = "";
            string INVOTYPE2 = "";
            string INVONO = "";
            string INVODATE = "";
            string UNIT = "";

            System.Data.DataTable T1 = GetCHO3();
            for (int i = 0; i <= T1.Rows.Count - 1; i++)
            {
                string TRANSID = T1.Rows[i]["TRANSID"].ToString();
                string DOCDATE = T1.Rows[i]["DOCDATE"].ToString();
                string BU = T1.Rows[i]["BU"].ToString();
                string MEMO = T1.Rows[i]["MEMO"].ToString();
                string A = T1.Rows[i]["A"].ToString();
                string B = T1.Rows[i]["B"].ToString();
                string C = T1.Rows[i]["C"].ToString();

                try
                {
                    string[] s = MEMO.Split(' ');


                    ACCOUNT = Convert.ToString(s[0]);
                    CARDNAME = Convert.ToString(s[1]);
                    INVOTYPE2 = Convert.ToString(s[2]);
                    INVOTYPE = util.INVOTYPE(Convert.ToString(s[2]));
                    INVONO = Convert.ToString(s[3]);
                    INVODATE = Convert.ToString(s[4]);
                    UNIT = Convert.ToString(s[5]);
                }
                catch { }

                if (INVOTYPE2 == "5")
                {
                    INSERT(DOCDATE, TRANSID, MEMO, A, B, C, CARDNAME, ACCOUNT, INVOTYPE, "1-進貨及費用", INVONO, INVODATE, UNIT, BU, toolStripComboBox1.Text);
                }
            }

            gB_POTATOARTableAdapter.FillBy(this.accBank.GB_POTATOAR, toolStripTextBox1.Text, toolStripComboBox1.Text);

            GETSUM();
            
        }

  


 

   
        
    }
}
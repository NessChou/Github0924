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
    public partial class ADINVOICEAP : Form
    {
        public ADINVOICEAP()
        {
            InitializeComponent();
        }

        private void aD_INVOICEAPBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aD_INVOICEAPBindingSource.EndEdit();
            this.aD_INVOICEAPTableAdapter.Update(this.accBank.AD_INVOICEAP);
            FF();
            MessageBox.Show("存檔成功");
        }
        private void FF()
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

        public System.Data.DataTable GetTSUM()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '0-三聯式發票/電子計算機發票' 類別,SUM(U_IN_BSAMN) 未稅金額,SUM(U_IN_BSTAX) 稅額,SUM(U_IN_BSAMT) 含稅總額 FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='0'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '1-三聯式收銀機發票/電子發票',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='1'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2-有稅憑證',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='2'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '3-海關代徵稅',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='3'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '4-免用統一發票/收據',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='4'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '5-退折',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1)='5'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '合計',SUM(U_IN_BSAMN) AMN,SUM(U_IN_BSTAX) TAX,SUM(U_IN_BSAMT) AMT FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA AND SUBSTRING(U_PC_BSTY1,1,1) IN ('0','1','2','3','4','5')");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));

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
        private void ADINVOICEAP_Load(object sender, EventArgs e)
        {
            toolStripTextBox1.Text = DateToStr(DateTime.Today.AddMonths(-1)).Substring(0, 6);


        }
        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            aD_INVOICEAPTableAdapter.FillBy(this.accBank.AD_INVOICEAP, toolStripTextBox1.Text);
            FF();
        
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            aD_INVOICEAPTableAdapter.FillBy1(this.accBank.AD_INVOICEAP, toolStripTextBox2.Text);
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

                    System.Data.DataTable T1 = GetCHO3();
                    for (int i = 0; i <= T1.Rows.Count - 1; i++)
                    {
                        string TRANSID = T1.Rows[i]["TRANSID"].ToString();
                        string DOCDATE = T1.Rows[i]["DOCDATE"].ToString();
                        string MEMO = T1.Rows[i]["MEMO"].ToString();
                        string A = T1.Rows[i]["A"].ToString();
                        string B = T1.Rows[i]["B"].ToString();
                        string C = T1.Rows[i]["C"].ToString();


                        string[] s = MEMO.Split(' ');

                        try
                        {

                            ACCOUNT = Convert.ToString(s[0]);
                            CARDNAME = Convert.ToString(s[1]);
                            INVOTYPE = util.INVOTYPE(Convert.ToString(s[2]));
                            INVONO = Convert.ToString(s[3]);
                            INVODATE = Convert.ToString(s[4]);
                            UNIT = Convert.ToString(s[5]);
                        }
                        catch
                        {}
                       
                        INSERT(DOCDATE, TRANSID, MEMO, A, B, C, CARDNAME, ACCOUNT,INVOTYPE,"1-進貨及費用", INVONO, INVODATE,UNIT);

                    }

                }
            }
            else
            {
                System.Data.DataTable T1 = GetCHO3();
                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {
                    string TRANSID = T1.Rows[i]["TRANSID"].ToString();
                    string DOCDATE = T1.Rows[i]["DOCDATE"].ToString();
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
                    catch
                    { }

                    INSERT(DOCDATE, TRANSID, MEMO, A, B, C, CARDNAME, ACCOUNT, INVOTYPE, "1-進貨及費用", INVONO, INVODATE, UNIT);
                }
            
            }
         

            aD_INVOICEAPTableAdapter.FillBy(this.accBank.AD_INVOICEAP, toolStripTextBox1.Text);

     
        }

       
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  VoucherNo TRANSID,CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,4) DOCDATE,Summary MEMO,CAST(SourceAmount AS INT) B,CAST(SourceAmount/0.05 AS INT) A,CAST(SourceAmount/0.05+SourceAmount AS INT) C from AccVoucherSub Where CAST(SUBSTRING(VoucherNo,1,2)+2011 AS VARCHAR)+SUBSTRING(VoucherNo,3,2)  = @DocDate1  AND SUBSTRING(VOUCHERNO,1,1) <> 'K' and SubjectID ='1281000' and SourceAmount <> 0 ");
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
        private void TRUNCATE()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" delete   AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));


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
        private void INSERT(string DelRemark, string TRANSID, string MEMO, string U_IN_BSAMN, string U_IN_BSTAX, string U_IN_BSAMT, string CARDNAME, string ACCOUNT,string U_PC_BSTY1,string U_PC_BSTY4, string SHIPDATE, string DOCDATE, string UNIT)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO AD_INVOICEAP(DelRemark,TRANSID,MEMO,U_IN_BSAMN,U_IN_BSTAX,U_IN_BSAMT,CARDNAME,ACCOUNT,U_PC_BSTY1,U_PC_BSTY4,SHIPDATE,DOCDATE,UNIT) VALUES(@DelRemark,@TRANSID,@MEMO,@U_IN_BSAMN,@U_IN_BSTAX,@U_IN_BSAMT,@CARDNAME,@ACCOUNT,@U_PC_BSTY1,@U_PC_BSTY4,@SHIPDATE,@DOCDATE,@UNIT )");
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
        public System.Data.DataTable GetINV()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM AD_INVOICEAP  WHERE SUBSTRING(DELREMARK,1,6) =@AA ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", toolStripTextBox1.Text));

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

        private void button4_Click(object sender, EventArgs e)
        {
           

            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcelSHARON(aD_INVOICEAPDataGridView);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
        }

        private void aD_INVOICEAPDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (aD_INVOICEAPDataGridView.Columns[e.ColumnIndex].Name == "U_IN_BSAMN" ||
                          aD_INVOICEAPDataGridView.Columns[e.ColumnIndex].Name == "U_IN_BSTAX")
                {

                    Int32 U_IN_BSAMN = 0;
                    Int32 U_IN_BSTAX = 0;

                    U_IN_BSAMN = Convert.ToInt32(this.aD_INVOICEAPDataGridView.Rows[e.RowIndex].Cells["U_IN_BSAMN"].Value);
                    U_IN_BSTAX = Convert.ToInt32(this.aD_INVOICEAPDataGridView.Rows[e.RowIndex].Cells["U_IN_BSTAX"].Value);


                    this.aD_INVOICEAPDataGridView.Rows[e.RowIndex].Cells["U_IN_BSAMT"].Value = (U_IN_BSAMN + U_IN_BSTAX);


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aD_INVOICEAPDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

            }
        }
    }
}

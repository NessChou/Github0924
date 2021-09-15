using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.VisualBasic.FileIO;
namespace ACME
{
    public partial class CustPay : Form
    {
        public CustPay()
        {
            InitializeComponent();
        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                if (cardCodeTextBox.Text == "")
                {
                    MessageBox.Show("請輸入資訊");
                }
                else
                {


                    //cbMainBank.SelectedIndex = 0;
          
                        ViewBatchPayment();
              
                }
                dataGridView4.Columns[0].HeaderText = "匯款日期";
                dataGridView4.Columns[1].HeaderText = "客戶編號";
                dataGridView4.Columns[2].HeaderText = "客戶名稱";
                dataGridView4.Columns[3].HeaderText = "解款行代號";
                dataGridView4.Columns[4].HeaderText = "收款人帳號";
                dataGridView4.Columns[5].HeaderText = "匯款金額";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ViewBatchPayment()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("   SELECT Convert(varchar(8),a.TrsFrDate,112) PmntDate,A.CARDCODE cardcode,A.CARDNAME cardname, ");
            sb.Append("                        B.DflBranch DflBranch,B.DflACCOUNT DflACCOUNT,PymAmount=  convert(int, A.doctotal) , ");
            sb.Append("                        LicTradNum,b.CntctPrsn CntctPrsn,A.docentry docentry FROM acmesql02.dbo.OVPM A LEFT JOIN acmesql02.dbo.OCRD B ON (A.CARDCODE=B.CARDCODE)");
            sb.Append("              where a.cardcode=@cardcode ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@cardcode", cardCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " OVPM");
            }
            finally
            {
                connection.Close();
            }
            dataGridView4.DataSource = ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();

            if (LookupValues != null)
            {
                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ACME.CustPayrpt frm = new ACME.CustPayrpt();
               frm.q = cardCodeTextBox.Text;

                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            string aa = @"\\acmesrv01\SAP_Share\LC\AccountDocument\整批付款歷史查詢.doc";
            System.Diagnostics.Process.Start(aa);
        }

      
    }
}
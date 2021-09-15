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

    public partial class GBPICKCHI : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBPICKCHI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("請確認是否要更新正航訂單數量", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                decimal AQTY = 0;
                decimal AMOUNT = 0;
                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {
                    DataGridViewRow row;
                    row = dataGridView1.Rows[i];

                    string 訂單編號 = row.Cells["訂單編號"].Value.ToString();
                    string 欄號 = row.Cells["欄號"].Value.ToString();
                    string 產品編號 = row.Cells["產品編號"].Value.ToString();
                    string 工單號碼 = row.Cells["工單號碼"].Value.ToString();
                    decimal 公斤 = Convert.ToDecimal(row.Cells["公斤"].Value.ToString());
                    System.Data.DataTable J1 = DTCHI(訂單編號, 欄號, 產品編號);

                    if (J1.Rows.Count > 0)
                    {
                        decimal PRICE = Convert.ToDecimal(J1.Rows[0][0].ToString());
                        decimal K1 = PRICE * 公斤;
                        AQTY += 公斤;
                        AMOUNT += K1;
                        UPDATEPICK(公斤, K1, 訂單編號, 欄號);
                        UPDATEPICK22(AMOUNT, AQTY, 訂單編號);
                        UPDATEPICK2(訂單編號, 欄號);
                    }

                }
           
                dataGridView1.DataSource = DT();
            }
        }

        private System.Data.DataTable DT()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT T0.SHIPPINGCODE 工單號碼,BILLNO 訂單編號,ROWNO 欄號,MAX(ITEMCODE) 產品編號,MAX(ITEMNAME) 產品名稱,SUM(CAST(PACK3 AS DECIMAL)) 公斤 FROM GB_PICK T0");
            sb.Append("  LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE ISNULL(CHI,'') <> 'Checked'");
            sb.Append(" GROUP BY T0.SHIPPINGCODE,BILLNO,ROWNO");
            sb.Append(" HAVING SUM(CAST(PACK3 AS INT)) <> 0");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


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
        private System.Data.DataTable DTCHI(string BillNO, string RowNO, string ProdID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("                select Price   FROM OrdBillSub WHERE BillNO=@BillNO AND RowNO =@ROWNO AND ProdID=ProdID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@RowNO", RowNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        private void GBPICKCHI_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = DT();
        }

        public void UPDATEPICK(decimal Quantity, decimal Amount, string BillNO, string RowNO)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("            UPDATE  OrdBillSub   SET Quantity=@Quantity,Amount=@Amount,MlAmount=@Amount,QtyRemain=@Quantity,EQuantity=@Quantity,sQuantity=@Quantity   FROM OrdBillSub WHERE BillNO=@BillNO AND RowNO =@RowNO   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
            command.Parameters.Add(new SqlParameter("@RowNO", RowNO));

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


        public void UPDATEPICK22(decimal SumAmtATax, decimal SumQty,string BillNO)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                 UPDATE OrdBillmain SET SumBTaxAmt=@SumAmtATax,SumQty=@SumQty,SumAmtATax=@SumAmtATax,LocalTotal=@SumAmtATax  WHERE BillNO=@BillNO   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SumAmtATax", SumAmtATax));
            command.Parameters.Add(new SqlParameter("@SumQty", SumQty));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

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
        public void UPDATEPICK2(string BILLNO, string ROWNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  UPDATE GB_PICK2  SET CHI='Checked' WHERE BILLNO=@BILLNO AND ROWNO=@ROWNO    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ROWNO", ROWNO));
 

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
    }
}

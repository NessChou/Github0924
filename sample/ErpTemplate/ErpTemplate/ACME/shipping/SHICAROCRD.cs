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
    public partial class SHICAROCRD : Form
    {
        public string Q;
        int scrollPosition = 0;
        public string cs;
        public SHICAROCRD(string ID)
        {
            InitializeComponent();
            Q = ID;
        }

        private void ViewBatchPayment7()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  select cardname   from ocrd WHERE SUBSTRING(CARDCODE,1,1)='U' ");
            if (textBox1.Text != "")
            {
                sb.Append(" and cardname like  '%" + textBox1.Text.ToString() + "%'  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                connection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

    

        private void APS1_Load(object sender, EventArgs e)
        {
            ViewBatchPayment7();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ViewBatchPayment7();
        }





        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "colEdit2")
                {
                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                    if (row != null)
                    {
                        cs = Convert.ToString(row["CardName"]);
                        UPDATESAP(cs, Q);
                        this.DialogResult = DialogResult.OK;
                        this.Close();


                    }


                }
            }
        }
        public void UPDATESAP(string CARTYPE, string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE shipping_CAR4 SET CARTYPE=@CARTYPE WHERE ID=@ID  ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARTYPE", CARTYPE));
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
    }
}
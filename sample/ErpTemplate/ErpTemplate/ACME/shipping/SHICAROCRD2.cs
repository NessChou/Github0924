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
    public partial class SHICAROCRD2 : Form
    {
        public string Q;
        public string Q2;
        int scrollPosition = 0;
        public string cs;
        public string csL;
        public string csW;
        public string csH;
        public SHICAROCRD2(string ID, string CARDTYPE)
        {
            InitializeComponent();
            Q = ID;
            Q2 = CARDTYPE;
        }

        private void ViewBatchPayment7(string CARTYPE)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT DISTINCT CARSIZE 車型,CARSIZEL 長,CARSIZEW 寬 ,CARSIZEH 高   FROM shipping_CAR4 FROM shipping_CAR4 WHERE CARTYPE=@CARTYPE");
        

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARTYPE", CARTYPE));

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
            ViewBatchPayment7(Q2);
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
                        cs = Convert.ToString(row["車型"]);
                        csL = Convert.ToString(row["長"]);
                        csW = Convert.ToString(row["寬"]);
                        csH = Convert.ToString(row["高"]);
                        UPDATESAP(cs, csL, csW, csH, Q);
                        this.DialogResult = DialogResult.OK;
                        this.Close();


                    }


                }
            }
        }
        public void UPDATESAP(string CARSIZE, string CARSIZEL, string CARSIZEW, string CARSIZEH, string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("UPDATE shipping_CAR4 SET CARSIZE=@CARSIZE,CARSIZEL=@CARSIZEL,CARSIZEW=@CARSIZEW,CARSIZEH=@CARSIZEH WHERE ID=@ID  ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARSIZE", CARSIZE));
            command.Parameters.Add(new SqlParameter("@CARSIZEL", CARSIZEL));
            command.Parameters.Add(new SqlParameter("@CARSIZEW", CARSIZEW));
            command.Parameters.Add(new SqlParameter("@CARSIZEH", CARSIZEH));
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
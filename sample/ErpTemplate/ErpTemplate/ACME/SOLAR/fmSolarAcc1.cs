using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ACME
{
    public partial class fmSolarAcc1 : Form
    {
        public fmSolarAcc1()
        {
            InitializeComponent();
        }

        public fmSolarAcc1(string TransID)
        {
            InitializeComponent();


           // MessageBox.Show(TransID);
            dataGridView2.AutoGenerateColumns = false;
            LoadData(TransID);
            
        }

        private void LoadData(string TransID)
        {
            this.Cursor = Cursors.AppStarting;

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("select T0.RefDate,T0.TransType,T0.ContraAct,T0.LineMemo,T0.SYSDeb,T0.SYSCred,T0.Account,T1.[AcctName] from jdt1 T0 ");
            sb.Append("Inner join  [OACT] T1  ON  T1.AcctCode = T0.Account ");
            sb.Append("where TransID=@TransID ");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            ////

            command.Parameters.Add(new SqlParameter("@TransID", TransID));
            //command.Parameters.Add(new SqlParameter("@RefDate2", RefDate2));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRC");
            }
            finally
            {
                connection.Close();
                this.Cursor = Cursors.Default;
            }
            DataTable dt = ds.Tables[0];

            dataGridView2.DataSource = dt;
          


        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == -1) return;

            DataGridView dgv = (DataGridView)sender;
            DataGridViewRow dgr = dgv.Rows[e.RowIndex];
            DataRowView row = (DataRowView)dgv.Rows[e.RowIndex].DataBoundItem;


            if (e.ColumnIndex == 7)
            {
                string s = Convert.ToString(e.Value);

                if (s == "30")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "JE";
                }
                else if (s == "15")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Delivery";
                }
                else if (s == "16")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Returns";
                }
                else if (s == "13")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/R Invoice";
                }
                else if (s == "14")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/R Credit Memo";
                }
                else if (s == "132")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Correction Invoice";
                }
                else if (s == "20")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Receipt";
                }
                else if (s == "21")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Returns";
                }
                else if (s == "18")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/P Invoice";
                }
                else if (s == "19")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "A/P Credit Memo";
                }
                else if (s == "-2")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Opening Balance";
                }
                else if (s == "58")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Stock Update";
                }
                else if (s == "59")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Receipt";
                }
                else if (s == "60")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Goods Issue";
                }
                else if (s == "67")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Inventory Transfers";
                }
                else if (s == "67")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Inventory Transfers";
                }
                else if (s == "68")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "Work Instructions";
                }
                else if (s == "-1")
                {
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "All Transactions";
                }
            }
        }
    }
}
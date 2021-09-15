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
    public partial class GBBOM : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBBOM()
        {
            InitializeComponent();
        }

        private void gB_BOMMBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_BOMMBindingSource.EndEdit();
            gB_BOMMTableAdapter.Update(pOTATO.GB_BOMM);
            this.gB_BOMDBindingSource.EndEdit();
            gB_BOMDTableAdapter.Update(pOTATO.GB_BOMD);
           
            for (int ii = 0; ii <= gB_BOMMDataGridView.Rows.Count - 2; ii++)
            {
                string CODE = gB_BOMMDataGridView.Rows[ii].Cells["CODE"].Value.ToString();
                decimal  AMT = 0;
                System.Data.DataTable L1 = GetBOMDSUM(CODE);
                if (L1.Rows.Count > 0)
                {
                    for (int ds = 0; ds <= L1.Rows.Count - 1; ds++)
                    {
                        string QTY = L1.Rows[ds]["QTY"].ToString();
                        string PRICE = L1.Rows[ds]["PTICE"].ToString();
                        if (String.IsNullOrEmpty(QTY))
                        {
                            QTY = "0";
                        }
                        if (String.IsNullOrEmpty(PRICE))
                        {
                            PRICE = "0";
                        }

                        AMT += Convert.ToInt16(QTY) * Convert.ToDecimal(PRICE);
                    }
                    UpdatePRICED(AMT, CODE);
                }

   

            }
            this.gB_BOMMTableAdapter.Fill(this.pOTATO.GB_BOMM);
            MessageBox.Show("存檔成功");

        }

        private void GBBOM_Load(object sender, EventArgs e)
        {

            this.gB_BOMDTableAdapter.Fill(this.pOTATO.GB_BOMD);
            this.gB_BOMMTableAdapter.Fill(this.pOTATO.GB_BOMM);

            gB_BOMMDataGridView.ReadOnly = true;
            gB_BOMDDataGridView.ReadOnly = true;
        }
        private void UpdatePRICED(decimal PRICED,string CODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update GB_BOMM set PRICED=@PRICED WHERE CODE=@CODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CODE", CODE));
            command.Parameters.Add(new SqlParameter("@PRICED", PRICED));
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

        private DataTable GetBOMDSUM(string FATHE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT *  FROM GB_BOMD WHERE FATHER=@FATHE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@FATHE", FATHE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            gB_BOMMDataGridView.ReadOnly = false;
            gB_BOMDDataGridView.ReadOnly = false;

            gB_BOMDDataGridView.Columns["AMT"].ReadOnly = true;
        }

        private void gB_BOMDDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (gB_BOMDDataGridView.Columns[e.ColumnIndex].Name == "QTY" ||
         gB_BOMDDataGridView.Columns[e.ColumnIndex].Name == "PTICE")
            {
                decimal QTY = 0;
                decimal PTICE = 0;

                QTY = Convert.ToInt32(this.gB_BOMDDataGridView.Rows[e.RowIndex].Cells["QTY"].Value);
                PTICE = Convert.ToDecimal(this.gB_BOMDDataGridView.Rows[e.RowIndex].Cells["PTICE"].Value);
                this.gB_BOMDDataGridView.Rows[e.RowIndex].Cells["AMT"].Value = (QTY * PTICE).ToString();

            }
            if (gB_BOMDDataGridView.Columns[e.ColumnIndex].Name == "CODE2")
            {
                string CODE = this.gB_BOMDDataGridView.Rows[e.RowIndex].Cells["CODE2"].Value.ToString();
                System.Data.DataTable G1 = GETProdID(CODE);
                if (G1.Rows.Count > 0)
                {
                    this.gB_BOMDDataGridView.Rows[e.RowIndex].Cells["CODENAME2"].Value = G1.Rows[0][0].ToString();
                }

            }
        }

        private void gB_BOMMDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (gB_BOMMDataGridView.Columns[e.ColumnIndex].Name == "CODE")
            {
                string CODE = this.gB_BOMMDataGridView.Rows[e.RowIndex].Cells["CODE"].Value.ToString();
                System.Data.DataTable G1 = GETProdID(CODE);
                if (G1.Rows.Count > 0)
                {
                    this.gB_BOMMDataGridView.Rows[e.RowIndex].Cells["CODENAME"].Value = G1.Rows[0][0].ToString();
                }

            }

        }
        public System.Data.DataTable GETProdID(string ProdID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT InvoProdName FROM comProduct WHERE ProdID =@ProdID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

    }
}

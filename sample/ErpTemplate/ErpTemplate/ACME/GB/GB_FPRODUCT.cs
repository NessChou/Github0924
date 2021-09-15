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
    public partial class GB_FPRODUCT : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_FPRODUCT()
        {
            InitializeComponent();
        }

        private void gB_FPRODUCTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_FPRODUCTBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);

        }

        private void GB_FPRODUCT_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'pOTATO.GB_FPRODUCT' 資料表。您可以視需要進行移動或移除。
            this.gB_FPRODUCTTableAdapter.Fill(this.pOTATO.GB_FPRODUCT);

        }

        private void gB_FPRODUCTDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gB_FPRODUCTDataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE")
                {
                    string T1 = gB_FPRODUCTDataGridView.Rows[e.RowIndex].Cells["ITEMCODE"].Value.ToString();
                    System.Data.DataTable G1 = GetCHO(T1);
                    if (G1.Rows.Count > 0)
                    {
                        gB_FPRODUCTDataGridView.Rows[e.RowIndex].Cells["ITEMNAME"].Value = G1.Rows[0][0].ToString();
                        gB_FPRODUCTDataGridView.Rows[e.RowIndex].Cells["UNIT"].Value = G1.Rows[0][1].ToString();
                    }

  

                }

            }
            catch
            {

            }
        }
        public System.Data.DataTable GetCHO(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ProdName,Unit   FROM comProduct where ProdID =@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
    }
}

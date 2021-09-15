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
          
    public partial class GB_FOC : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_FOC()
        {
            InitializeComponent();
        }



        private void gB_FOCBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_FOCBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.pOTATO);
            SetControlEnabled(Controls);
  
            MessageBox.Show("存檔完成");

        }
        public void SetControlEnabled(System.Windows.Forms.Control.ControlCollection originalControls)
        {
   
            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
               
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlEnabled(originalControls[i].Controls);
                }

                string f = originalControls[i].ToString();

                if (originalControls[i] is CheckBox)
                {


                    CheckBox aTextBox = (CheckBox)originalControls[i];


                    if (aTextBox.Checked)
                    {

         MessageBox.Show(aTextBox.Text);
                    }

                }
            

     

            }

        }
        private void GB_FOC_Load(object sender, EventArgs e)
        {
            this.gB_FOC3TableAdapter.Fill(this.pOTATO.GB_FOC3);
            this.gB_FOC2TableAdapter.Fill(this.pOTATO.GB_FOC2);
            this.gB_FOCTableAdapter.Fill(this.pOTATO.GB_FOC);



        }

        private void gB_FOC2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (gB_FOC2DataGridView.Columns[e.ColumnIndex].Name == "ITEMCODE")
            {
                string ITEMNAME = "";
                string ITEMCODE = gB_FOC2DataGridView.Rows[e.RowIndex].Cells["ITEMCODE"].Value.ToString();
                System.Data.DataTable KK1 = GETProdID(ITEMCODE);
                if (KK1.Rows.Count > 0)
                {
                    ITEMNAME = KK1.Rows[0][0].ToString();
                }
                this.gB_FOC2DataGridView.Rows[e.RowIndex].Cells["ITEMNAME"].Value = ITEMNAME;
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

        private void gB_FOC2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["QTY"].Value = 1;
        }

        private void gB_FOCDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["AMT"].Value = 0;
        }
    }
}

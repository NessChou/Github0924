using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class HRDIN : Form
    {
        public HRDIN()
        {
            InitializeComponent();
        }

        private void dINBENDON_USERBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.dINBENDON_USERBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.hR);

            MessageBox.Show("存檔成功");

        }

        private void HRDIN_Load(object sender, EventArgs e)
        {

            toolStripTextBox1.Text = GetMenu.Day();
            this.dINBENDON_USERTableAdapter.Fill(this.hR.DINBENDON_USER,toolStripTextBox1.Text);

        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            this.dINBENDON_USERTableAdapter.Fill(this.hR.DINBENDON_USER, toolStripTextBox1.Text);
        }

        private void dINBENDON_USERDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["日期"].Value = toolStripTextBox1.Text;
            System.Data.DataTable GF = GetshipTYPE();
            if (GF.Rows.Count > 0)
            {
                e.Row.Cells["DINNER"].Value = GF.Rows[0][0].ToString();
            }
        }
        public System.Data.DataTable GetshipTYPE()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT MENU  FROM  DNN2.DBO.DINBENDON_kido where menudate=@menudate");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@menudate", toolStripTextBox1.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
    }
}

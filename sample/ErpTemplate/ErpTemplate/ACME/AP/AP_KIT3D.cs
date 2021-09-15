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
    public partial class AP_KIT3D : Form
    {
        public string PublicString;
        public AP_KIT3D()
        {
            InitializeComponent();
        }

        private void aP_KIT9DBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_KIT9DBindingSource.EndEdit();
            this.aP_KIT9DTableAdapter.Update(this.lC.AP_KIT9D);

        }


        private void AP_KIT3D_Load(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(PublicString))
            {
                try
                {
                    mIDTextBox.Text = PublicString;
                    System.Data.DataTable GGD = GetMID();
                    if (GGD.Rows.Count == 0)
                    {


                        ADDMID();


                    }
       
                    this.aP_KIT9DTableAdapter.Fill(this.lC.AP_KIT9D, ((int)(System.Convert.ChangeType(PublicString, typeof(int)))));
                    mIDTextBox.Text = PublicString;
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }

        public System.Data.DataTable GetMID()
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MID FROM AP_KIT9D WHERE MID=@MID");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MID", mIDTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public void ADDMID()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_KIT9D(MID) values(@MID)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MID",Convert.ToInt32(mIDTextBox.Text) ));
            
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

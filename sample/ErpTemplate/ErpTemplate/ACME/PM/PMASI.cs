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
    public partial class PMASI : Form
    {
        public PMASI()
        {
            InitializeComponent();
        }

        private System.Data.DataTable GetTABLE2()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT DISTINCT BRAND,PANELTYPE, PANELMODEL FROM RMA_ASI WHERE BRAND < > 'AUO' AND PANELMODEL <> '' AND BRAND='ACME' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetTABLE3()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT DISTINCT BRAND,PANELTYPE, PANELMODEL FROM RMA_ASI WHERE BRAND < > 'AUO' AND PANELMODEL <> '' AND BRAND <> 'ACME' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMAASITYPE(string MODEL)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT  * FROM RMA_ASI_TYPE WHERE MODEL=@MODEL ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private void PMASI_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "ACME";
            System.Data.DataTable G1 = GetTABLE2();

            dataGridView1.DataSource = G1;
        }

  

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (!dgv.Focused) return;
            mODELTextBox.Text = dgv.CurrentRow.Cells[2].Value.ToString();

                System.Data.DataTable H1 = GetRMAASITYPE(mODELTextBox.Text);
                if (H1.Rows.Count > 0)
                {
                    CARDtextBox.Text = H1.Rows[0]["CARD"].ToString();
                    tYPETextBox.Text = H1.Rows[0]["TYPE"].ToString();
                    string TOUCH = H1.Rows[0]["TOUCH"].ToString();
                    string BRIGHT = H1.Rows[0]["BRIGHT"].ToString();
                    string CUT = H1.Rows[0]["CUT"].ToString();
                    if (TOUCH == "True")
                    {
                        checkBox1.Checked = true;
                    }
                    if (BRIGHT == "True")
                    {
                        checkBox2.Checked = true;
                    }
                    if (CUT == "True")
                    {
                        checkBox3.Checked = true;
                    }
                }
                else
                {
                    CARDtextBox.Text = "";
                    tYPETextBox.Text = "";
                    checkBox1.Checked = false;
                    checkBox2.Checked = false;
                    checkBox3.Checked = false;
                }
        }


        public void RMAASIIN(string MODEL, string TYPE, string TOUCH, string BRIGHT, string CUT, string CARD)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_ASI_TYPE(MODEL,TYPE,TOUCH,BRIGHT,CUT,CARD) values(@MODEL,@TYPE,@TOUCH,@BRIGHT,@CUT,@CARD)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@TOUCH", TOUCH));
            command.Parameters.Add(new SqlParameter("@BRIGHT", BRIGHT));
            command.Parameters.Add(new SqlParameter("@CUT", CUT));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));
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

        public void RMAASIUP(string MODEL, string TYPE, string TOUCH, string BRIGHT, string CUT, string CARD)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE RMA_ASI_TYPE SET [TYPE]=@TYPE,TOUCH=@TOUCH,BRIGHT=@BRIGHT,CUT=@CUT,CARD=@CARD WHERE  MODEL=@MODEL", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));
            command.Parameters.Add(new SqlParameter("@TOUCH", TOUCH));
            command.Parameters.Add(new SqlParameter("@BRIGHT", BRIGHT));
            command.Parameters.Add(new SqlParameter("@CUT", CUT));
            command.Parameters.Add(new SqlParameter("@CARD", CARD));

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

        private void button1_Click(object sender, EventArgs e)
        {
            string TOUCH="";
            string BRIGHT="";
            string CUT="";
            if (checkBox1.Checked)
            {
                TOUCH = "True";
            }
            if (checkBox2.Checked)
            {
                BRIGHT = "True";
            }
             if (checkBox3.Checked)
             {
                 CUT = "True";
             }
            System.Data.DataTable H1 = GetRMAASITYPE(mODELTextBox.Text);
            if (H1.Rows.Count > 0)
            {
                RMAASIUP(mODELTextBox.Text, tYPETextBox.Text, TOUCH, BRIGHT, CUT, CARDtextBox.Text);

            }
            else
            {
                RMAASIIN(mODELTextBox.Text, tYPETextBox.Text, TOUCH, BRIGHT, CUT, CARDtextBox.Text);
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "ACME")
            {
                System.Data.DataTable G1 = GetTABLE2();

                dataGridView1.DataSource = G1;
            }
            if (comboBox1.Text == "OTHERS")
            {
                System.Data.DataTable G1 = GetTABLE3();

                dataGridView1.DataSource = G1;
            }
        }
    }
}

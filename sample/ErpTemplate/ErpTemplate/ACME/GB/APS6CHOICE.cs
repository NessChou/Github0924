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
    public partial class APS6CHOICE : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public string MM2;

        public APS6CHOICE()
        {
            InitializeComponent();
        }

        private void ViewBatchPayment7()
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT EngName   from comCustClass where Flag =1 AND ClassID >008 AND ISNULL(EngName,'') <> ''  ");
            string GROUP = globals.GroupID.ToString().Trim();

            if (textBox1.Text != "")
            {
                sb.Append(" and EngName like  '%" + textBox1.Text.ToString() + "%'  ");
            }
      
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oitm");
            }
            finally
            {
                connection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
             

                    ArrayList al = new ArrayList();

                    for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                    {
                        al.Add(listBox1.Items[i].ToString());
                    }
                    StringBuilder sb = new StringBuilder();



                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);


                    MM2 = sb.ToString();
         
                    
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void APS1_Load(object sender, EventArgs e)
        {
            ViewBatchPayment7();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ViewBatchPayment7();
        }

        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];



                        listBox1.Items.Add(row.Cells["itemcode"].Value.ToString());

                    }

  


                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ViewBatchPayment7();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
            {
                dataGridView1.Rows[i].Selected = true;

            }

            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow row;


                for (int h = dataGridView1.SelectedRows.Count - 1; h >= 0; h--)
                {

                    row = dataGridView1.SelectedRows[h];



                    listBox1.Items.Add(row.Cells["itemcode"].Value.ToString());

                }




            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
            {
                dataGridView1.Rows[i].Selected = false;
            }
        }
    }
}
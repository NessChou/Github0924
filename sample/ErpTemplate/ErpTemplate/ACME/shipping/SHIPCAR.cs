using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class SHIPCAR : Form 
    {
        public string SHIPPINGCODE;

        public string a;
        public SHIPCAR()
        {
            InitializeComponent();
        }

        private void AP_Load(object sender, EventArgs e)
        {

      
            ViewBatchPayment2();
                                  
        }




        private void ViewBatchPayment2()
        {
            System.Data.DataTable DF1 = GetMenu.getCAR(SHIPPINGCODE);


            dataGridView1.DataSource = DF1;

        }


        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    StringBuilder sb = new StringBuilder();
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];

                        sb.Append("'" + row.Cells["FLAG1"].Value.ToString()  + "',");
                       
                    }


                    sb.Remove(sb.Length - 1, 1);
                    string q = sb.ToString();
                    a = q;
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


   
    }
}
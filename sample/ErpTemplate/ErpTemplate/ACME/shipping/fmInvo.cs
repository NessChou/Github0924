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
    public partial class fmInvo : Form
    {
        public string Invo;
        public fmInvo()
        {
            InitializeComponent();
        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt = GetAP(textBox1.Text,textBox2.Text);
                if (dt.Rows.Count > 0)
                {
                    DataRow drw = dt.Rows[0];
                    Invo = drw["SHIPPINGCODE"].ToString();
                    Close();
                }
                else

                {
                    MessageBox.Show("無資料");
                }
               
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        
  
        public static System.Data.DataTable GetAP(string INVOICENO, string INVOICENO_SEQ)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SHIPPINGCODE FROM INVOICEM WHERE INVOICENO=@INVOICENO AND INVOICENO_SEQ=@INVOICENO_SEQ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@INVOICENO_SEQ", INVOICENO_SEQ));
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
        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

      

      


      
    }
}
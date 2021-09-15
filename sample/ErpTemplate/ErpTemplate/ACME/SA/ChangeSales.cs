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
    public partial class ChangeSales : Form
    {
        int  DD = 0 ;
        public ChangeSales()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBatchPayment5();
        }
        private void ViewBatchPayment5()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SLPCODE FROM ORDR where docentry=@Docentry ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }


           // dataGridView1.DataSource = ds.Tables[0];

         //   System.Data.DataTable DF = GetOrderData8();
      
                
                DataTable fd = GetOrderData8(ds.Tables[0].Rows[0][0].ToString());

                textBox2.Text = fd.Rows[0][0].ToString();
          
        }
        private System.Data.DataTable GetOrderData8(string SLPCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  select SLPNAME  from OSLP where SLPCODE=@SLPCODE  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SLPCODE", SLPCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
              

                UPDATE2(DD);
                MessageBox.Show("更新成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
  
        }
        private void AddTRACKER_LOG(decimal U_COMMISSION, string docentry, string linenum)
        {



            SqlConnection connection = globals.shipConnection; 
            StringBuilder sb = new StringBuilder();
            sb.Append(" update INV1 set U_COMMISSION=@U_COMMISSION  where docentry=@docentry and linenum=@linenum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@U_COMMISSION", U_COMMISSION));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));


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

        private void UPDATE2(decimal SLPCODE)
        {



            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update ORDR set SLPCODE=@SLPCODE  where docentry=@docentry");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SLPCODE", SLPCODE));
            command.Parameters.Add(new SqlParameter("@Docentry", textBox1.Text));


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

        private void button3_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetLP("目的地");

            if (LookupValues != null)
            {
                textBox2.Text = Convert.ToString(LookupValues[1]);

                DD = Convert.ToInt32(LookupValues[0]);

            }
        }

        public static object[] GetLP(string aa)
        {
            string[] FieldNames = new string[] { "SLPCODE","SLPNAME" };

            string[] Captions = new string[] { "業務代碼", "業務名稱" };

            string SqlScript = " select CAST(SLPCODE AS VARCHAR) SLPCODE,SLPNAME from OSLP  ";

            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
    }
}
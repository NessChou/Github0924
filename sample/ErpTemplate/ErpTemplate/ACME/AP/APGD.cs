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
    public partial class APGD : Form
    {
        public APGD()
        {
            InitializeComponent();
        }

        private void APGD_Load(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GETOPCH();

            for (int i = 0; i <= G1.Rows.Count - 1; i++)
            {
                string CARDNAME = G1.Rows[i]["CARDNAME"].ToString();
                string INV = G1.Rows[i]["INV"].ToString();

                System.Data.DataTable G2 = GETODLN(INV);

                if (G2.Rows.Count > 0)
                {
                    for (int i2 = 0; i2 <= G2.Rows.Count - 1; i2++)
                    {
                        string DOCENTRY = G2.Rows[i2]["DOCENTRY"].ToString();
                        UpdateSQL(CARDNAME, DOCENTRY);
                    }
                }
            }
        }
        private void UpdateSQL(string U_ACME_KIND1, string docentry)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update  ODLN set U_ACME_KIND1=@U_ACME_KIND1 where docentry=@docentry  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_ACME_KIND1", U_ACME_KIND1));
            command.Parameters.Add(new SqlParameter("@docentry", docentry));

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
        private System.Data.DataTable GETODLN(string U_ACME_INV)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCENTRY,U_ACME_INV INV  FROM ODLN  WHERE U_ACME_INV like '%" + U_ACME_INV + "%' AND ISNULL(U_ACME_KIND1,'') ='' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@DocDate1", textBox10.Text));
            //command.Parameters.Add(new SqlParameter("@DocDate2", textBox11.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETOPCH()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT CARDNAME,U_ACME_INV INV FROM OPCH  WHERE U_ACME_INV <> ''  AND U_ACME_INV <>'123' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
  

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    }
}

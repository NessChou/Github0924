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
    public partial class KITINVOICE : Form
    {
        public string q1;
        public string q2;
        public string q3;
        public string q4;
        public KITINVOICE()
        {
            InitializeComponent();
        }

        private void iNVOICEDKITBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.iNVOICEDKITBindingSource.EndEdit();
            this.iNVOICEDKITTableAdapter.Update(this.ship.INVOICEDKIT);


            MessageBox.Show("存檔成功");

        }



        private void KITINVOICE_Load(object sender, EventArgs e)
        {
            T1();
            try
            {
                iNVOICEDKITTableAdapter.Connection = globals.Connection;
                this.iNVOICEDKITTableAdapter.Fill(this.ship.INVOICEDKIT, q1, q2, q3, q4);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        private void T1()
        {
            System.Data.DataTable H = GetSHIP(q1, q2, q3, q4);
            System.Data.DataTable H2 = GetSHIP2(q1, q2, q3, q4);
            if (H.Rows.Count > 0)
            {

                string DOC = H.Rows[0][0].ToString();


                if (H2.Rows.Count == 0)
                {

                    System.Data.DataTable H4 = GetSAP2(DOC, q4);
                    int VISORDER = Convert.ToInt16(H4.Rows[0][0]);
                    System.Data.DataTable H3 = GetSAP(DOC, VISORDER);
                    for (int i = 0; i <= H3.Rows.Count - 1; i++)
                    {
                        DataRow drw = H3.Rows[i];
                        int LINE = Convert.ToInt16(drw["VISORDER"]);
                        string D1 = drw["DSCRIPTION"].ToString();
                        string QTY = drw["QTY"].ToString();
                        string TREETYPE = drw["TREETYPE"].ToString();

                        if (TREETYPE != "I")
                        {
                            return;
                        }
                        try
                        {

                            int N1 = D1.IndexOf("_");
                            string D2 = D1.Substring(N1 + 1, D1.Length - N1 - 1);
                            int N2 = D2.IndexOf("_");
                            string D3 = D2.Substring(N2 + 1, D2.Length - N2 - 1);
                            int N3 = D3.IndexOf("_");
                            string D4 = D1.Substring(0, N1 + N2 + N3 + 2);
                            InsertKIT(q1, q2, q3, LINE, D4, QTY,q4);

                        }
                        catch
                        {
                            InsertKIT(q1, q2, q3, LINE, D1, QTY, q4);
                        }

                    }
                }


            }
        }
        private System.Data.DataTable GetSHIP(string SHIPPINGCODE, string INVOICENO, string INVOICENO_SEQ, string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT SOID FROM INVOICED WHERE SHIPPINGCODE=@SHIPPINGCODE AND INVOICENO=@INVOICENO AND INVOICENO_SEQ=@INVOICENO_SEQ AND ITEMCODE=@ITEMCODE AND treetype='S' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@INVOICENO_SEQ", INVOICENO_SEQ));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSHIP2(string SHIPPINGCODE, string INVOICENO, string INVOICENO_SEQ, string ITEMNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT *  FROM INVOICEDKIT WHERE SHIPPINGCODE=@SHIPPINGCODE AND INVOICENO=@INVOICENO AND INVOICENO_SEQ=@INVOICENO_SEQ  AND ITEMNAME=@ITEMNAME ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@INVOICENO_SEQ", INVOICENO_SEQ));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetSAP(string DOCENTRY, int VISORDER)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT VISORDER,DSCRIPTION,CAST(QUANTITY AS INT) QTY,TREETYPE FROM RDR1 WHERE  DOCENTRY=@DOCENTRY  AND VISORDER > @VISORDER ORDER BY VISORDER ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@VISORDER", VISORDER));
            
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetSAP2(string DOCENTRY, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT VISORDER FROM RDR1 WHERE TREETYPE='S' AND DOCENTRY=@DOCENTRY AND ITEMCODE=@ITEMCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void InsertKIT(string SHIPPINGCODE, string INVOICENO, string INVOICENO_SEQ, int LINE, string KIT, string QTY, string ITEMNAME)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO INVOICEDKIT (SHIPPINGCODE,INVOICENO,INVOICENO_SEQ,LINE,KIT,QTY,ITEMNAME) VALUES(@SHIPPINGCODE,@INVOICENO,@INVOICENO_SEQ,@LINE,@KIT,@QTY,@ITEMNAME)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@INVOICENO_SEQ", INVOICENO_SEQ));
            command.Parameters.Add(new SqlParameter("@LINE", LINE));
            command.Parameters.Add(new SqlParameter("@KIT", KIT));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
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

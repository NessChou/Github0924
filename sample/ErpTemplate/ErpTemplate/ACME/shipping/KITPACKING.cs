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
    public partial class KITPACKING : Form
    {
        decimal NET = 0;
        decimal GROSS = 0;
        public string q1;
        public string q2;
        public string q3;
        public KITPACKING()
        {
            InitializeComponent();
        }

        private void packingListDKITBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            
        }
        private void CalcTotals1C()
        {
            try
            {

             


                int i = this.packingListDKITDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    if (!String.IsNullOrEmpty(packingListDKITDataGridView.Rows[iRecs].Cells["Net1"].Value.ToString().Trim()))
                    {
                        NET += Convert.ToDecimal(packingListDKITDataGridView.Rows[iRecs].Cells["Net1"].Value.ToString().Trim());
                    }
                    if (!String.IsNullOrEmpty(packingListDKITDataGridView.Rows[iRecs].Cells["Gross1"].Value.ToString().Trim()))
                    {
                        GROSS += Convert.ToDecimal(packingListDKITDataGridView.Rows[iRecs].Cells["Gross1"].Value.ToString().Trim());
                    }



                }
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void KITPACKING_Load(object sender, EventArgs e)
        {
            T1();

            try
            {
                packingListDKITTableAdapter.Connection = globals.Connection;
                this.packingListDKITTableAdapter.Fill(this.ship.PackingListDKIT, q1, q2, q3);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void T1()
        {
            System.Data.DataTable H = GetSHIP(q1, q2, q3);
            System.Data.DataTable H2 = GetSHIP2(q1, q2, q3);
            if (H.Rows.Count > 0)
            {

                string DOC = H.Rows[0][0].ToString();


                if (H2.Rows.Count == 0)
                {
                    System.Data.DataTable H4 = GetSAP2(DOC, q3);
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

                        int SEQNO = i + 1;
                        try
                        {
                            int N1 = D1.IndexOf("_");
                            string D2 = D1.Substring(N1 + 1, D1.Length - N1 - 1);
                            int N2 = D2.IndexOf("_");
                            string D3 = D2.Substring(N2 + 1, D2.Length - N2 - 1);
                            int N3 = D3.IndexOf("_");
                            string D4 = D1.Substring(0, N1 + N2 + N3 + 2);
                            InsertKIT(q1, q2, LINE, D4, QTY, SEQNO, q3);
                        }
                        catch
                        {
                            InsertKIT(q1, q2, LINE, D1, QTY, SEQNO, q3);
                        }

                    }
                }


            }

        }
        private void InsertKIT(string SHIPPINGCODE, string PLNo, int LINE, string KIT, string QTY, int SeqNo, string ITEMNAME)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO PackingListDKIT (SHIPPINGCODE,PLNo,LINE,KIT,QTY,SeqNo,ITEMNAME) VALUES(@SHIPPINGCODE,@PLNo,@LINE,@KIT,@QTY,@SeqNo,@ITEMNAME)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@LINE", LINE));
            command.Parameters.Add(new SqlParameter("@KIT", KIT));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@SeqNo", SeqNo));
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
        private void UPKIT(string SHIPPINGCODE, string PLNo, string NET, string GROSS)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE PackingListD SET NET=@NET,GROSS=@GROSS where shippingcode=@shippingcode AND PLNo=@PLNo  AND TREETYPE='S'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@NET", NET));
            command.Parameters.Add(new SqlParameter("@GROSS", GROSS));


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
        private System.Data.DataTable GetSHIP(string SHIPPINGCODE, string PLNo, string DESCGOODS)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT SOID FROM PackingListD WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNo=@PLNo  AND DESCGOODS=@ITEM AND  treetype='S' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
            command.Parameters.Add(new SqlParameter("@ITEM", DESCGOODS));
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

        private System.Data.DataTable GetSHIP2(string SHIPPINGCODE, string PLNo, string ITEMNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT *  FROM PackingListDKIT WHERE SHIPPINGCODE=@SHIPPINGCODE AND PLNo=@PLNo AND ITEMNAME=@ITEMNAME  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PLNo", PLNo));
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

        private System.Data.DataTable GetSAP(string DOCENTRY, int  VISORDER)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT VISORDER,DSCRIPTION,CAST(QUANTITY AS INT) QTY,TREETYPE FROM RDR1 WHERE  DOCENTRY=@DOCENTRY AND  VISORDER > @VISORDER ");
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
        private System.Data.DataTable GetSAP2(string DOCENTRY, string DSCRIPTION)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT VISORDER FROM RDR1 WHERE TREETYPE='S' AND DOCENTRY=@DOCENTRY AND DSCRIPTION=@DSCRIPTION ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@DSCRIPTION", DSCRIPTION));

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
        private void button1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.packingListDKITBindingSource.EndEdit();
            this.packingListDKITTableAdapter.Update(this.ship.PackingListDKIT);
            CalcTotals1C();
            if (NET != 0 && GROSS != 0)
            {
                UPKIT(q1, q2, NET.ToString(), GROSS.ToString());
            }
            MessageBox.Show("更新成功");
        }

        private void packingListDKITDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = packingListDKITDataGridView.Rows.Count ;
            e.Row.Cells["SeqNo"].Value = iRecs.ToString();
        }
    }
}

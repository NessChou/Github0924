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
    public partial class SQUTOITM : Form
    {
        public SQUTOITM()
        {
            InitializeComponent();
        }

        private void shipping_OQUT5BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= shipping_OQUT5BindingSource.Count; i++)
            {
                DataRowView row1 = (DataRowView)shipping_OQUT5BindingSource.Current;

                row1["ITEMCODE"] = i;



                shipping_OQUT5BindingSource.EndEdit();

                shipping_OQUT5BindingSource.MoveNext();
            }

            
            this.Validate();
            this.shipping_OQUT5BindingSource.EndEdit();
            this.shipping_OQUT5TableAdapter.Update(this.ship.Shipping_OQUT5);

            MessageBox.Show("存檔成功");

        }

        private void SQUTOITM_Load(object sender, EventArgs e)
        {
            this.shipping_OQUT5TableAdapter.Fill(this.ship.Shipping_OQUT5);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddCARD();
            this.shipping_OQUT5TableAdapter.Fill(this.ship.Shipping_OQUT5);
            for (int i = 1; i <= shipping_OQUT5BindingSource.Count; i++)
            {
                DataRowView row1 = (DataRowView)shipping_OQUT5BindingSource.Current;

                row1["ITEMCODE"] = i;



                shipping_OQUT5BindingSource.EndEdit();

                shipping_OQUT5BindingSource.MoveNext();
            }

            this.Validate();
            this.shipping_OQUT5BindingSource.EndEdit();
            this.shipping_OQUT5TableAdapter.Update(this.ship.Shipping_OQUT5);


            this.shipping_OQUT5TableAdapter.Fill(this.ship.Shipping_OQUT5);

        }

        public void AddCARD()
        {
            int iRecs;

            iRecs = shipping_OQUT5DataGridView.Rows.Count;
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Shipping_OQUT5(ITEMCODE,ITEMNAME) values(@ITEMCODE,@ITEMNAME)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", iRecs.ToString()));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", textITEMNAME.Text));

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

        private void shipping_OQUT5DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            //int iRecs;

            //iRecs = shipping_OQUT5DataGridView.Rows.Count;
            //e.Row.Cells["ITEMCODE"].Value = iRecs.ToString();
        }
    }
}

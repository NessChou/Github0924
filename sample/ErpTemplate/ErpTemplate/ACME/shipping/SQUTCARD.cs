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
    public partial class SQUTCARD : Form
    {
        public SQUTCARD()
        {
            InitializeComponent();
        }

        private void shipping_OQUT4BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= shipping_OQUT4BindingSource.Count; i++)
            {
                DataRowView row1 = (DataRowView)shipping_OQUT4BindingSource.Current;

                row1["CARDCODE"] = i;



                shipping_OQUT4BindingSource.EndEdit();

                shipping_OQUT4BindingSource.MoveNext();
            }

            this.Validate();
            this.shipping_OQUT4BindingSource.EndEdit();
            this.shipping_OQUT4TableAdapter.Update(this.ship.Shipping_OQUT4);
            MessageBox.Show("存檔成功");

        }

        private void SQUTCARD_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'ship.Shipping_OQUT4' 資料表。您可以視需要進行移動或移除。
            this.shipping_OQUT4TableAdapter.Fill(this.ship.Shipping_OQUT4);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddCARD();
            this.shipping_OQUT4TableAdapter.Fill(this.ship.Shipping_OQUT4);
            for (int i = 1; i <= shipping_OQUT4BindingSource.Count; i++)
            {
                DataRowView row1 = (DataRowView)shipping_OQUT4BindingSource.Current;

                row1["CARDCODE"] = i;



                shipping_OQUT4BindingSource.EndEdit();

                shipping_OQUT4BindingSource.MoveNext();
            }
            this.Validate();
            this.shipping_OQUT4BindingSource.EndEdit();
            this.shipping_OQUT4TableAdapter.Update(this.ship.Shipping_OQUT4);

            this.shipping_OQUT4TableAdapter.Fill(this.ship.Shipping_OQUT4);
        }
        public void AddCARD()
        {
            int iRecs;

            iRecs = shipping_OQUT4DataGridView.Rows.Count;
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Shipping_OQUT4(CARDCODE,CARDNAME) values(@CARDCODE,@CARDNAME)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", iRecs.ToString()));
            command.Parameters.Add(new SqlParameter("@CARDNAME", textCARDNAME.Text));

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


        private void shipping_OQUT4DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = shipping_OQUT4DataGridView.Rows.Count;
            e.Row.Cells["CARDCODE"].Value = iRecs.ToString();
        }
    }
}

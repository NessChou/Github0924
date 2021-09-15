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
    public partial class OITT : Form
    {
        public OITT()
        {
            InitializeComponent();
        }

        private void oITTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= iTT1DataGridView.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = iTT1DataGridView.Rows[i];
                string ITEM = row.Cells["產品編號"].Value.ToString();
                string PRICE = row.Cells["Price"].Value.ToString();
                string QTY = row.Cells["Quantity"].Value.ToString();

                System.Data.DataTable GS = GETOITM(ITEM);
                if (GS.Rows.Count == 0)
                {
                    MessageBox.Show("請輸入正確產品編號");
                    return;
                }

                if (String.IsNullOrEmpty(PRICE))
                {
                    MessageBox.Show("請輸入單價");
                    return;
                }

                if (String.IsNullOrEmpty(QTY))
                {
                    MessageBox.Show("請輸入數量");
                    return;
                }
            }

            this.Validate();
            this.oITTBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.sAP);
            MessageBox.Show("更新成功");

        }

        private void OITT_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'sAP.ITT1' 資料表。您可以視需要進行移動或移除。
            this.iTT1TableAdapter.Fill(this.sAP.ITT1);
            // TODO: 這行程式碼會將資料載入 'sAP.OITT' 資料表。您可以視需要進行移動或移除。
            this.oITTTableAdapter.Fill(this.sAP.OITT);

        }

        private void iTT1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = iTT1DataGridView.Rows.Count - 1;
            e.Row.Cells["ChildNum"].Value = iRecs.ToString();
            e.Row.Cells["Quantity"].Value = 0.000000;
            e.Row.Cells["Warehouse"].Value = "TW017";
            e.Row.Cells["Price"].Value = 0.000000;
            e.Row.Cells["Currency"].Value = "";
            e.Row.Cells["PriceList"].Value = 1;
            e.Row.Cells["OrigPrice"].Value = 0.000000;
            e.Row.Cells["OrigCurr"].Value ="";
            e.Row.Cells["IssueMthd"].Value = "M";
            e.Row.Cells["Object"].Value = "66";
            e.Row.Cells["PrncpInput"].Value = "N";
            this.iTT1BindingSource.EndEdit();
        }

        private System.Data.DataTable GETOITM(string ITEMCODE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM OITM WHERE ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    }
}

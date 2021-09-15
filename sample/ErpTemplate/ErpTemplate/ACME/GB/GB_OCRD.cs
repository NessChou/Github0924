using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class GB_OCRD : Form
    {
        public GB_OCRD()
        {
            InitializeComponent();
        }

        private void gB_OCRDBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_OCRDBindingSource.EndEdit();
            this.gB_OCRDTableAdapter.Update(this.POTATO.GB_OCRD);

            this.gB_OCRD2BindingSource.EndEdit();
            this.gB_OCRD2TableAdapter.Update(this.POTATO.GB_OCRD2);

            MessageBox.Show("存檔成功");

        }

        private void GB_OCRD_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'POTATO.GB_OCRD2' 資料表。您可以視需要進行移動或移除。
            this.gB_OCRD2TableAdapter.Fill(this.POTATO.GB_OCRD2);
            this.gB_OCRDTableAdapter.Fill(this.POTATO.GB_OCRD);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_OCRDBindingSource.EndEdit();
            this.gB_OCRDTableAdapter.Update(this.POTATO.GB_OCRD);
            if (gB_OCRDDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇");
                return;
            }

            object[] LookupValues = GetMenu.GetGBOITM("大宗");

            if (LookupValues != null)
            {


                System.Data.DataTable dt2 = POTATO.GB_OCRD2;
                string ITEMCODE = Convert.ToString(LookupValues[0]);
                DataRow drw2 = dt2.NewRow();
                string da = gB_OCRDDataGridView.SelectedRows[0].Cells["ID"].Value.ToString();
                drw2["ID"] = da;
                drw2["ITEMCODE"] = Convert.ToString(LookupValues[0]);
                drw2["ITEMNAME"] = Convert.ToString(LookupValues[1]);
                drw2["PRICE"] = Convert.ToString(LookupValues[2]);
                drw2["AMOUNT"] = Convert.ToString(LookupValues[2]);
                drw2["QTY"] = Convert.ToString(LookupValues[2]);

                //PRICE
                dt2.Rows.Add(drw2);

                this.Validate();
                this.gB_OCRD2BindingSource.EndEdit();
                this.gB_OCRD2TableAdapter.Update(this.POTATO.GB_OCRD2);
            }
        }

        private void gB_OCRD2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gB_OCRD2DataGridView.Columns[e.ColumnIndex].Name == "QTY" ||
                          gB_OCRD2DataGridView.Columns[e.ColumnIndex].Name == "PRICE")
                {

                    Int32 Qty = 0;
                    Int32 PRICE = 0;

                    Qty = Convert.ToInt32(this.gB_OCRD2DataGridView.Rows[e.RowIndex].Cells["QTY"].Value);
                    PRICE = Convert.ToInt32(this.gB_OCRD2DataGridView.Rows[e.RowIndex].Cells["PRICE"].Value);


                    this.gB_OCRD2DataGridView.Rows[e.RowIndex].Cells["AMOUNT"].Value = (Qty * PRICE);


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

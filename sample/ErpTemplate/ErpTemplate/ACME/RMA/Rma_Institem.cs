using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class Rma_Institem : Form
    {
        public Rma_Institem()
        {
            InitializeComponent();
        }


        private void rma_InsuBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            //this.rma_InsuBindingSource.EndEdit();
            //this.rma_InsuTableAdapter.Update(this.rm.Rma_Insu);
            this.rma_Insu1BindingSource.EndEdit();
            this.rma_Insu1TableAdapter.Update(this.rm.Rma_Insu1);
            MessageBox.Show("存檔成功");
        }

        private void Rma_Institem_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'rm.Rma_Insu1' 資料表。您可以視需要進行移動或移除。
            this.rma_Insu1TableAdapter.Fill(this.rm.Rma_Insu1);
            // TODO: 這行程式碼會將資料載入 'rm.Rma_Insu' 資料表。您可以視需要進行移動或移除。
            this.rma_InsuTableAdapter.Fill(this.rm.Rma_Insu);

        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class WH_MAIL : Form
    {
        public WH_MAIL()
        {
            InitializeComponent();
        }

        private void wH_MAILBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.wH_MAILBindingSource.EndEdit();
            this.wH_MAILTableAdapter.Update(this.wh.WH_MAIL);

        }

        private void WH_MAIL_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'wh.WH_MAIL' 資料表。您可以視需要進行移動或移除。
            this.wH_MAILTableAdapter.Fill(this.wh.WH_MAIL);

        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class SHIP_FIREFEE : Form
    {
        public SHIP_FIREFEE()
        {
            InitializeComponent();
        }

        private void sHIP_FIREFEEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sHIP_FIREFEEBindingSource.EndEdit();
            this.sHIP_FIREFEETableAdapter.Update(this.ship.SHIP_FIREFEE);

        }

        private void SHIP_FIREFEE_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'ship.SHIP_FIREFEE' 資料表。您可以視需要進行移動或移除。
            this.sHIP_FIREFEETableAdapter.Fill(this.ship.SHIP_FIREFEE);

        }
    }
}

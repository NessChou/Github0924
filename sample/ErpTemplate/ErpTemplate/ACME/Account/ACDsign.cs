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
    public partial class ACDsign : Form
    {
        public ACDsign()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.tempBindingSource.EndEdit();
            this.tempTableAdapter.Update(this.ship.temp);

            MessageBox.Show("存檔成功");
        }




        private void ACDsign_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'ship.temp' 資料表。您可以視需要進行移動或移除。
            this.tempTableAdapter.Fill(this.ship.temp);
        
        

        }

    }
}
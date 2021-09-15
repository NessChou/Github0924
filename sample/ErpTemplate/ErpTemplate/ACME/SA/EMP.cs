using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
    public partial class EMP : Form
    {
        public EMP()
        {
            InitializeComponent();
        }

        private void dINBENDON_UsersBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.dINBENDON_UsersBindingSource.EndEdit();
            this.dINBENDON_UsersTableAdapter.Update(this.uSERS.DINBENDON_Users);

        }

        private void EMP_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'uSERS.DINBENDON_Users' 資料表。您可以視需要進行移動或移除。
            this.dINBENDON_UsersTableAdapter.Fill(this.uSERS.DINBENDON_Users);

        }
    }
}
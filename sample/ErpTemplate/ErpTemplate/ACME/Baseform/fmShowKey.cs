using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME.Baseform
{
    public partial class fmShowKey : Form
    {

        private string FKey;

        public string Key
        {
            get
            {
                return FKey;
            }
            set
            {
                FKey = value;
            }
        }
        
        public fmShowKey()
        {
            InitializeComponent();
        }

       

        private void btnOK_Click(object sender, EventArgs e)
        {

            Key = textBox1.Text;

            if (string.IsNullOrEmpty(Key))
            {
                MessageBox.Show("鍵值請輸入");
                textBox1.Focus();
                return;
            }

            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void fmShowKey_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (string.IsNullOrEmpty(Key))
            {
                MessageBox.Show("鍵值請輸入");
                textBox1.Focus();
                e.Cancel = true;
            }
        }
    }
}
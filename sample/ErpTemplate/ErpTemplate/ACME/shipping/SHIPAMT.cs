using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class SHIPAMT : Form 
    {
        public string JOBNO;
        public string b;
        public SHIPAMT()
        {
            InitializeComponent();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            b = textBox1.Text;
        }

      










   
    }
}
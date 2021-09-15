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
    public partial class SOLAPAY2 : Form
    {
        public SOLAPAY2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.sOLAR_PAY1TableAdapter.Fill(this.sOLAR.SOLAR_PAY1, textBox1.Text, textBox2.Text, comboBox1.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            if (comboBox1.Text == "採購請款")
            {
                sOLAR_PAY1DataGridView.Columns[9].Visible = false;
            }
            else if (comboBox1.Text == "預付請款")
            {
                sOLAR_PAY1DataGridView.Columns[9].Visible = true;
            }
        }

        private void SOLAPAY2_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            comboBox1.Text = "預付請款";

            try
            {
                this.sOLAR_PAY1TableAdapter.Fill(this.sOLAR.SOLAR_PAY1, textBox1.Text, textBox2.Text, comboBox1.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void sOLAR_PAY1DataGridView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (sOLAR_PAY1DataGridView.SelectedRows.Count > 0)
            {
                string da = sOLAR_PAY1DataGridView.SelectedRows[0].Cells["JOBNO"].Value.ToString();
                SOLARPAY a = new SOLARPAY();
                a.PublicString2 = da;
                a.ShowDialog();
            }
        }

        private void sOLAR_PAY1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (sOLAR_PAY1DataGridView.Columns[e.ColumnIndex].Name == "PAYCHECK")
            {
                    this.Validate();
                    this.sOLAR_PAY1BindingSource.EndEdit();
                    this.sOLAR_PAY1TableAdapter.Update(this.sOLAR.SOLAR_PAY1);
            }
        }

    }
}

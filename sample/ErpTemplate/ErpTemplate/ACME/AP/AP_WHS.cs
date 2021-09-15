 using System;
 using System.Collections.Generic;
 using System.ComponentModel;
 using System.Data;
 using System.Drawing;
 using System.Text;
 using System.Windows.Forms;
 using System.Reflection;
using System.Data.SqlClient;

namespace ACME
{
    public partial class AP_WHS : Form
    {

        private int _DocEntry = 0;

        public static string PrjCode;

        public AP_WHS()
        {
            InitializeComponent();
            TextBoxID.ReadOnly = true;
            TextBoxID.TabStop = false;

            //dpStartDate.Value = Convert.ToDateTime(DBNull.Value);
        }


        public AP_WHS(int DocEntry)
        {
            InitializeComponent();
            TextBoxID.ReadOnly = true;
            TextBoxID.TabStop = false;
            _DocEntry = DocEntry;
            btnDelete.Visible = (DocEntry > 0);

            //固定名稱
            TextBoxID.Text = _DocEntry.ToString();


            if (DocEntry > 0)
            {
                LoadData(DocEntry);
            }
            else
            {
              //  TextBoxOwner.Text = globals.UserID;

            }
        }


        private void LoadData(int DocEntry)
        {
            DataTable dt = Shipping_WHS.GetShipping_WHS(DocEntry);

            if (dt.Rows.Count > 0)
            {
                Shipping_WHS data = new Shipping_WHS();
                ReflectUtils.BindAllData(this, data, dt.Rows[0]);
            }


           
        }


        private void btnSave_Click(object sender, EventArgs e)
        {

            if (TextBoxLOCATION.Text == "")
            {
                MessageBox.Show("LOCATION不能空白");
                return;
            }
            Shipping_WHS row = new Shipping_WHS();

            ReflectUtils.GetAllData(this, row);

            if (_DocEntry == 0)
            {

                Shipping_WHS.AddShipping_WHS(row);

            }
            else
            {


                Shipping_WHS.UpdateShipping_WHS(row);

            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.None;
            this.Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("確定刪除嗎 ?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            {
                return;
            }

            Shipping_WHS row = new Shipping_WHS();

            row.ID = _DocEntry;

            Shipping_WHS.DeleteShipping_WHS(row);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }






 

    }
}


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
    public partial class UNLOCK : Form
    {

        private int _DocEntry = 0;

        public static string PrjCode;

        public UNLOCK()
        {
            InitializeComponent();
            TextBoxDocEntry.ReadOnly = true;
            TextBoxDocEntry.TabStop = false;

            //dpStartDate.Value = Convert.ToDateTime(DBNull.Value);
        }


        public UNLOCK(int DocEntry)
        {
            InitializeComponent();
            TextBoxDocEntry.ReadOnly = true;
            TextBoxDocEntry.TabStop = false;
            _DocEntry = DocEntry;
            btnDelete.Visible = (DocEntry > 0);

            //固定名稱
            TextBoxDocEntry.Text = _DocEntry.ToString();


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
            DataTable dt = ACME_CREDIT_UNLOCK.GetACME_CREDIT_UNLOCK(DocEntry);

            if (dt.Rows.Count > 0)
            {
                ACME_CREDIT_UNLOCK data = new ACME_CREDIT_UNLOCK();
                ReflectUtils.BindAllData(this, data, dt.Rows[0]);
            }


           
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            ACME_CREDIT_UNLOCK row = new ACME_CREDIT_UNLOCK();

            ReflectUtils.GetAllData(this, row);

            if (_DocEntry == 0)
            {

                row.CreateDate = DateTime.Now.ToString("yyyyMMdd");
                row.CreateTime = DateTime.Now.ToString("HHmmss");
                row.Handler = globals.UserID;

                ACME_CREDIT_UNLOCK.AddACME_CREDIT_UNLOCK(row);
            }
            else
            {


                ACME_CREDIT_UNLOCK.UpdateACME_CREDIT_UNLOCK(row);

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

            ACME_CREDIT_UNLOCK row = new ACME_CREDIT_UNLOCK();

            row.DocEntry = _DocEntry;

            ACME_CREDIT_UNLOCK.DeleteACME_CREDIT_UNLOCK(row);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void dp_ValueChanged(object sender, EventArgs e)
        {

            string sName = (sender as DateTimePicker).Name;
            sName = sName.Substring(2, sName.Length - 2);
            ((TextBox)ReflectUtils.FindControl(this, "TextBox" + sName)).Text = (sender as DateTimePicker).Value.ToString("yyyyMMdd");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           // TextBoxOwner.Text = comboBox1.Text;
        }

        private void fmACME_MIS_TASK_Load(object sender, EventArgs e)
        {
           // comboBox1.Items.Clear();

            //DataTable dt = GetOwner();

            //for (int i = 0; i <= dt.Rows.Count - 1; i++)
            //{
            //    comboBox1.Items.Add(Convert.ToString(dt.Rows[i][0]));
            //}


            //comboBox2.Items.Clear();

            //DataTable dtKind = GetKind();

            //for (int i = 0; i <= dtKind.Rows.Count - 1; i++)
            //{
            //    comboBox2.Items.Add(Convert.ToString(dtKind.Rows[i][0]));
            //}
        }


        public DataTable GetKind()
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT distinct Kind FROM ACME_MIS_TASK ";
            
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }

        public DataTable GetOwner()
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT distinct Owner FROM ACME_MIS_TASK ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MIS_TASK");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MIS_TASK"];
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
          //  TextBoxKind.Text = comboBox2.Text;
        }

        private void dpStartDate_ValueChanged(object sender, EventArgs e)
        {
            string sName = (sender as DateTimePicker).Name;
            sName = sName.Substring(2, sName.Length - 2);
            ((TextBox)ReflectUtils.FindControl(this, "TextBox" + sName)).Text = (sender as DateTimePicker).Value.ToString("yyyyMMdd");
        }

        private void dpEndDate_ValueChanged(object sender, EventArgs e)
        {
            string sName = (sender as DateTimePicker).Name;
            sName = sName.Substring(2, sName.Length - 2);
            ((TextBox)ReflectUtils.FindControl(this, "TextBox" + sName)).Text = (sender as DateTimePicker).Value.ToString("yyyyMMdd");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {
                TextBoxCardCode.Text = Convert.ToString(LookupValues[0]);
                TextBoxCardName.Text = Convert.ToString(LookupValues[1]);

            }
        }
    }
}


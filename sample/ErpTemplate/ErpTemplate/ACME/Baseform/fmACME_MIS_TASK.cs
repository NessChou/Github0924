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
    public partial class fmACME_MIS_TASK : Form
    {

        private int _ID = 0;

        public static string PrjCode;

        public fmACME_MIS_TASK()
        {
            InitializeComponent();
            TextBoxID.ReadOnly = true;
            TextBoxID.TabStop = false;

            dpStartDate.Value = Convert.ToDateTime(DBNull.Value);
        }


        public fmACME_MIS_TASK(int ID)
        {
            InitializeComponent();
            TextBoxID.ReadOnly = true;
            TextBoxID.TabStop = false;
            _ID = ID;
            btnDelete.Visible = (ID > 0);

            //固定名稱
            TextBoxID.Text = _ID.ToString();


            if (ID > 0)
            {
                LoadData(ID);
            }
            else
            {
                TextBoxOwner.Text = globals.UserID;

            }
        }


        private void LoadData(int ID)
        {
            DataTable dt = ACME_MIS_TASK.GetACME_MIS_TASK(ID);

            if (dt.Rows.Count > 0)
            {
                ACME_MIS_TASK data = new ACME_MIS_TASK();
                ReflectUtils.BindAllData(this, data, dt.Rows[0]);
            }


           
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            ACME_MIS_TASK row = new ACME_MIS_TASK();

            ReflectUtils.GetAllData(this, row);

            if (_ID == 0)
            {


                row.CreateDate = DateTime.Now.ToString("yyyyMMdd");
                row.CreateTime = DateTime.Now.ToString("HHmmss");
                row.CreateUser = globals.UserID;
                ACME_MIS_TASK.AddACME_MIS_TASK(row);
            }
            else
            {

                row.UpdateDate = DateTime.Now.ToString("yyyyMMdd");
                row.UpdateTime = DateTime.Now.ToString("HHmmss");
                row.UpdateUser = globals.UserID;

                ACME_MIS_TASK.UpdateACME_MIS_TASK(row);

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

            ACME_MIS_TASK row = new ACME_MIS_TASK();

            row.ID = _ID;

            ACME_MIS_TASK.DeleteACME_MIS_TASK(row);

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
            TextBoxOwner.Text = comboBox1.Text;
        }

        private void fmACME_MIS_TASK_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();

            DataTable dt = GetOwner();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt.Rows[i][0]));
            }


            comboBox2.Items.Clear();

            DataTable dtKind = GetKind();

            for (int i = 0; i <= dtKind.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dtKind.Rows[i][0]));
            }

            comboBox3.Items.Clear();

            DataTable dtBU = GetBU();

            for (int i = 0; i <= dtBU.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dtBU.Rows[i][0]));
            }


            comboBox4.Items.Clear();

            DataTable dtUNIT = GetUNIT();

            for (int i = 0; i <= dtUNIT.Rows.Count - 1; i++)
            {
                comboBox4.Items.Add(Convert.ToString(dtUNIT.Rows[i][0]));
            }
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
        public DataTable GetBU()
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT distinct BU FROM ACME_MIS_TASK ";

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
        public DataTable GetUNIT()
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT distinct UNIT FROM ACME_MIS_TASK ";

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
            TextBoxKind.Text = comboBox2.Text;

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxUNIT.Text = comboBox4.Text;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxBU.Text = comboBox3.Text;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxEDIT.Text = comboBox5.Text;
            TextBoxTask.Text = comboBox2.Text + " - " + comboBox5.Text + comboBox4.Text;
        }

     
    }
}


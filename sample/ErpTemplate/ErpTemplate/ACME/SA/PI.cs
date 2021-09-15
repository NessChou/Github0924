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
    public partial class PI : ACME.fmBase1
    {
        public PI()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sACUSTTableAdapter.Connection = MyConnection;
        }
        public override void SetInit()
        {
            MyBS = sACUSTBindingSource;
            MyTableName = "SACUST";
            MyIDFieldName = "id";
        }
        public override void EndEdit()
        {
            dataGridView1.DataSource = GetordOBillSub1();
        }
        public override void STOP()
        {


            if (tYPETextBox.Text == "")
            {
                MessageBox.Show("請輸入帳戶類型");
                this.SSTOPID = "1";
                tYPETextBox.Focus();
                return;
            }
        }
        public override void SetDefaultValue()
        {
            tYPETextBox.Text = "美金帳戶";
            string NumberName = "SC";
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.idTextBox.Text = NumberName + AutoNum;
            this.sACUSTBindingSource.EndEdit();
        }
        public override void FillData()
        {
            try
            {

                sACUSTTableAdapter.Fill(sa.SACUST,MyID);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable GetordOBillSub1()
        {


            //取得未交數量 > 0
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select TYPE,cardcode,memo from sacust");
   
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "sacust");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["sacust"];
        }
        public override bool UpdateData()
        {
            bool UpdateData;

            try
            {

                sACUSTTableAdapter.Connection.Open();

                Validate();

                sACUSTBindingSource.EndEdit();


                sACUSTTableAdapter.Update(sa.SACUST);

                this.MyID = this.idTextBox.Text;

                UpdateData = true;

              
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.sACUSTTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void PI_Load(object sender, EventArgs e)
        {
   

            dataGridView1.DataSource = GetordOBillSub1();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            tYPETextBox.Text = comboBox1.Text;
        }
    }
}


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
    public partial class PORT : Form
    {
        public PORT()
        {
            InitializeComponent();
        }

        private void account_Temp7BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            if (globals.DBNAME == "進金生")
            {
                if (fmLogin.LoginID.ToString().ToUpper() == "JOYCHEN" || fmLogin.LoginID.ToString().ToUpper() == "SHIRLEYJUAN" || fmLogin.LoginID.ToString().ToUpper() == "LLEYTONCHEN")
                {	
                    this.Validate();
                    this.account_Temp7BindingSource.EndEdit();
                    this.account_Temp7TableAdapter.Update(this.ship.Account_Temp7);
                    Update11();
                    this.account_Temp7TableAdapter.Fill(this.ship.Account_Temp7);
                    MessageBox.Show("儲存成功");
                }
                else
                {
                    MessageBox.Show("您沒有修改權限");
                }
            }
            else
            {
                this.Validate();
                this.account_Temp7BindingSource.EndEdit();
                this.account_Temp7TableAdapter.Update(this.ship.Account_Temp7);
                Update11();
                this.account_Temp7TableAdapter.Fill(this.ship.Account_Temp7);
                MessageBox.Show("儲存成功");
            
            }

           
        }

        private void PORT_Load(object sender, EventArgs e)
        {
            account_Temp7TableAdapter.Connection = globals.Connection;
            this.account_Temp7TableAdapter.Fill(this.ship.Account_Temp7);

        }

        private void Update11()
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update Account_Temp7 set port=upper(port) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }


        }
    }
}
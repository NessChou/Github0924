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
            if (globals.DBNAME == "�i����")
            {
                if (fmLogin.LoginID.ToString().ToUpper() == "JOYCHEN" || fmLogin.LoginID.ToString().ToUpper() == "SHIRLEYJUAN" || fmLogin.LoginID.ToString().ToUpper() == "LLEYTONCHEN")
                {	
                    this.Validate();
                    this.account_Temp7BindingSource.EndEdit();
                    this.account_Temp7TableAdapter.Update(this.ship.Account_Temp7);
                    Update11();
                    this.account_Temp7TableAdapter.Fill(this.ship.Account_Temp7);
                    MessageBox.Show("�x�s���\");
                }
                else
                {
                    MessageBox.Show("�z�S���ק��v��");
                }
            }
            else
            {
                this.Validate();
                this.account_Temp7BindingSource.EndEdit();
                this.account_Temp7TableAdapter.Update(this.ship.Account_Temp7);
                Update11();
                this.account_Temp7TableAdapter.Fill(this.ship.Account_Temp7);
                MessageBox.Show("�x�s���\");
            
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
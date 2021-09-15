using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class GBMEMBER : Form
    {

        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBMEMBER()
        {
            InitializeComponent();
        }

        private void GBMEMBER_Load(object sender, EventArgs e)
        {

            dataGridView2.DataSource = GetMEM3();

        }

        private System.Data.DataTable GetMEM3()
        {
     
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ''''+U.ID,L.ClassName 類別名稱,FullName 客戶名稱  FROM comCustomer U");
            sb.Append(" Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1 ");
            sb.Append(" WHERE U.Flag =1  AND ID NOT LIKE '%TW%'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }


    
        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
       
        }

     

   
    }
}

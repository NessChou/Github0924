using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class ACCHOROD : Form
    {

        string strCn = "";
        public ACCHOROD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox4.Text == "宇豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (comboBox4.Text == "博豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp09;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "INFINITE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "CHOICE")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (comboBox4.Text == "TOP GARDEN")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            dataGridView1.DataSource = GetCHO4();

            for (int i = 3; i <= 7; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                if (i == 7)
                {
                    col.DefaultCellStyle.Format = "#,##0.00";
                }
                else
                {
                    col.DefaultCellStyle.Format = "#,##0";
                }


            }
        }
        public System.Data.DataTable GetCHO4()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select ClassID 類別編號,ClassName 類別名稱,AccInventory 存貨科目,AccPurchased 進貨科目,ReturnPurchase 進貨退出,AccSale 銷貨收入,AccSaleCost 銷貨成本,");
            sb.Append(" ReturnSale 銷貨退回,GiftExpense 贈品費用,OtherIncome 其他收入 ");
            sb.Append("  from comProductClass ");

       
            



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        private void ACCCHOICE_Load(object sender, EventArgs e)
        {
   
            comboBox4.Text = "宇豐";

            if (globals.UserID.ToUpper() != "NANCYWEI" && globals.GroupID.ToString().Trim() != "EEP")
            {

                button1.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }


   

    }
}

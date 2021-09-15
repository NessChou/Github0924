using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
namespace ACME
{
    public partial class CART : Form
    {
        public CART()
        {
            InitializeComponent();
        }



        private void CART_Load(object sender, EventArgs e)
        {

            if (globals.DBNAME == "宇豐")
            {
                dataGridView1.DataSource = GetOrderDataAD();

                tabControl1.TabPages.Remove(tabControl1.TabPages["tabPage2"]);
                tabControl1.TabPages.Remove(tabControl1.TabPages["tabPage3"]);
                tabControl1.TabPages.Remove(tabControl1.TabPages["tabPage4"]);
                tabControl1.TabPages.Remove(tabControl1.TabPages["tabPage5"]);
                tabControl1.TabPages["tabPage1"].Text = "ADLAB";

                textBox1.Visible = false;
                label1.Visible = false;
            }
            else
            {
                dataGridView1.DataSource = GetOrderData("TFT");
                dataGridView2.DataSource = GetOrderData("OPEN");
                dataGridView3.DataSource = GetOrderData2();
                dataGridView4.DataSource = GetOrderData("TCON");
                dataGridView5.DataSource = GetOrderData("其他");
            }
     

        }
        private System.Data.DataTable GetOrderData(string A1)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT [MODEL_NO],[MODEL_Ver],TMEMO 小料號包裝共用,[CT_QTY],[CT_NW],[CT_GW],[CT_L],[CT_W],[CT_H]");
            sb.Append("              ,[PAL_QTY],[PAL_CTNS],[PAL_NW],[PAL_GW],[PAL_L],[PAL_W],[PAL_H],[CT20_QTY]");
            sb.Append("              ,[CT20_PLTS],[CT20_L],[CT20_W],[CT20_H],[CT40_QTY],[CT40_PLTS],[CT40_L],[CT40_W]");
            sb.Append("              ,[CT40_H],[CT40_HQTY],[CT40_HPLTS],[CT40_HL],[CT40_HW],[CT40_HH],");
            sb.Append("                       CASE WHEN (CASE CT_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_GW AS DECIMAL(10,2)),0) END) >");
            sb.Append("                       (CAST(ROUND((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))) THEN ");
            sb.Append("                       (CASE CT_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_GW AS DECIMAL(10,2)),0) END)");
            sb.Append("                       ELSE CAST(ROUND((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))END ");
            sb.Append("                       箱材積重,CASE WHEN (CASE PAL_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_GW AS DECIMAL(10,2)),0) END) >");
            sb.Append("                       (CAST(ROUND((CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))) THEN ");
            sb.Append("                       (CASE PAL_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_GW AS DECIMAL(10,2)),0) END)");
            sb.Append("                       ELSE CAST(ROUND((CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))END  板材積重,");
            sb.Append("                  Ceiling((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/27000)  快遞重             ");
            sb.Append("                                ,(CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END* ");
            sb.Append("                               CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END * ");
            sb.Append("                                CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/1000000 CBM,    ");
            sb.Append("             [CREATE_USER],[CREATE_DATE] ,[UPDATE_DATE],[UPDATE_USER],[memo]");
            sb.Append(" FROM [dbo].[CART] WHERE 1=1 ");

            if (A1 == "OPEN")
            {
                sb.Append("  AND DOCTYPE = 'OPEN CELL'  ");
            }
            else
            {
                sb.Append("  AND DOCTYPE = @A1 ");
            }
     
            if (textBox1.Text != "")
            {
                sb.Append("  AND MODEL_NO LIKE  @MODEL_NO ");


            }

            sb.Append(" ORDER BY   [MODEL_NO],[MODEL_Ver] ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", "%" + textBox1.Text + "%"));
            command.Parameters.Add(new SqlParameter("@A1", A1));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderDataAD()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT DOCTYPE,[MODEL_NO],[MODEL_Ver],TMEMO 小料號包裝共用,[CT_QTY],[CT_NW],[CT_GW],[CT_L],[CT_W],[CT_H]");
            sb.Append("              ,[PAL_QTY],[PAL_CTNS],[PAL_NW],[PAL_GW],[PAL_L],[PAL_W],[PAL_H],[CT20_QTY]");
            sb.Append("              ,[CT20_PLTS],[CT20_L],[CT20_W],[CT20_H],[CT40_QTY],[CT40_PLTS],[CT40_L],[CT40_W]");
            sb.Append("              ,[CT40_H],[CT40_HQTY],[CT40_HPLTS],[CT40_HL],[CT40_HW],[CT40_HH],");
            sb.Append("                       CASE WHEN (CASE CT_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_GW AS DECIMAL(10,2)),0) END) >");
            sb.Append("                       (CAST(ROUND((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))) THEN ");
            sb.Append("                       (CASE CT_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_GW AS DECIMAL(10,2)),0) END)");
            sb.Append("                       ELSE CAST(ROUND((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))END ");
            sb.Append("                       箱材積重,CASE WHEN (CASE PAL_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_GW AS DECIMAL(10,2)),0) END) >");
            sb.Append("                       (CAST(ROUND((CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))) THEN ");
            sb.Append("                       (CASE PAL_GW WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_GW AS DECIMAL(10,2)),0) END)");
            sb.Append("                       ELSE CAST(ROUND((CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/5000,2) AS DECIMAL(10,2))END  板材積重,");
            sb.Append("                  Ceiling((CASE CT_L WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_L AS DECIMAL(10,2)),0) END*");
            sb.Append("                       CASE CT_W WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_W AS DECIMAL(10,2)),0) END *");
            sb.Append("                       CASE CT_H WHEN '' THEN 0 ELSE  ISNULL(CAST(CT_H AS DECIMAL(10,2)),0) END )/27000)  快遞重             ");
            sb.Append("                                ,(CASE PAL_L WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_L AS DECIMAL(10,2)),0) END* ");
            sb.Append("                               CASE PAL_W WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_W AS DECIMAL(10,2)),0) END * ");
            sb.Append("                                CASE PAL_H WHEN '' THEN 0 ELSE  ISNULL(CAST(PAL_H AS DECIMAL(10,2)),0) END )/1000000 CBM,    ");
            sb.Append("             [CREATE_USER],[CREATE_DATE] ,[UPDATE_DATE],[UPDATE_USER],[memo]");
            sb.Append(" FROM [dbo].[CART] WHERE 1=1 ");

        

            if (textBox1.Text != "")
            {
                sb.Append("  AND MODEL_NO LIKE  @MODEL_NO ");


            }

            sb.Append(" ORDER BY   [MODEL_NO],[MODEL_Ver] ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", "%" + textBox1.Text + "%"));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [ITEMCODE] 項目號碼,[ITEMNAME] 項目說明 ,[CT_QTY],[CT_NW],[CT_GW]");
            sb.Append("       ,[UNIT] 單位,[CREATE_USER],[CREATE_DATE],[UPDATE_DATE],[memo] 備註,[UPDATE_USER]");
            sb.Append("   FROM [dbo].[CART_LED]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToExcel(dataGridView5);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetOrderData("TFT");
            dataGridView2.DataSource = GetOrderData("OPEN");
        }

     

    


    }
}
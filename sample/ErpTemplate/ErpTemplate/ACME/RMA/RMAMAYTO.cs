using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Reflection;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;

namespace ACME
{
    public partial class RMAMAYTO : Form
    {
        public RMAMAYTO()
        {
            InitializeComponent();
        }


        private System.Data.DataTable GetMAYTO(string U_RMA_NO, string U_CUSNAME_S)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_RMODEL MODEL,U_RVER VER,U_RQUINITY QTY,U_CUSNAME_S CARDNAME,U_ROUTWHAREHOUSE WH,U_RMA_NO FROM OCTR WHERE U_RMA_NO in ( " + U_RMA_NO + ") AND U_CUSNAME_S=@U_CUSNAME_S ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_CUSNAME_S", U_CUSNAME_S));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAYTO2(string U_RMA_NO, string U_ROUTWHAREHOUSE, string U_CUSNAME_S)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_RMA_NO,U_RMODEL MODEL,U_RVER VER,U_RQUINITY QTY,U_CUSNAME_S CARDNAME,U_ROUTWHAREHOUSE WH FROM OCTR WHERE U_RMA_NO in ( " + U_RMA_NO + ") AND U_ROUTWHAREHOUSE=@U_ROUTWHAREHOUSE AND U_CUSNAME_S=@U_CUSNAME_S");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ROUTWHAREHOUSE", U_ROUTWHAREHOUSE));
            command.Parameters.Add(new SqlParameter("@U_CUSNAME_S", U_CUSNAME_S));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAYTOD(string U_RMA_NO, string U_CUSNAME_S)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT U_ROUTWHAREHOUSE FROM OCTR WHERE U_RMA_NO in ( " + U_RMA_NO + ") AND U_CUSNAME_S=@U_CUSNAME_S ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_CUSNAME_S", U_CUSNAME_S));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAYTOCUST(string U_RMA_NO)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT U_CUSNAME_S CARDNAME FROM OCTR WHERE U_RMA_NO in ( " + U_RMA_NO + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;





            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button1_Click(object sender, EventArgs e)
        {

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string aa = GetMenu.SPLITDOC(textBox4.Text);

            System.Data.DataTable g3 = GetMAYTOCUST(aa);

            if (g3.Rows.Count > 0)
            {
                for (int i = 0; i <= g3.Rows.Count - 1; i++)
                {
                    string CARDNAME = g3.Rows[i][0].ToString();
                    System.Data.DataTable g1 = GetMAYTO(aa, CARDNAME);
                    System.Data.DataTable g2 = GetMAYTOD(aa, CARDNAME);
                    string OutPutFile = "";

                    if (g1.Rows.Count > 0)
                    {
                        string WH = g1.Rows[0]["WH"].ToString();

                        if (g2.Rows.Count == 1)
                        {

                            if (WH == "內湖")
                            {
                                FileName = lsAppDir + "\\Excel\\RMA\\內湖嘜頭.xls";

                            }
                            else
                            {
                                FileName = lsAppDir + "\\Excel\\RMA\\聯倉嘜頭.xls";
                            }
                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
        DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                            ExcelReport.ExcelReportOutput(g1, FileName, OutPutFile, "N");
                        }
                        else
                        {
                            FileName = lsAppDir + "\\Excel\\RMA\\內聯嘜頭.xls";
                            System.Data.DataTable H1 = GetMAYTO2(aa, "內湖", CARDNAME);
                            System.Data.DataTable H2 = GetMAYTO2(aa, "聯倉", CARDNAME);
                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
        DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                            ExcelReport.ExcelReportOutputMAYTO(H1, FileName, OutPutFile, H2);
                        }


                    }
                    else
                    {
                        MessageBox.Show("沒有資料");
                    }
                }
            }
        }
    }
}

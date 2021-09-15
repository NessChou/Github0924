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
    public partial class ProductSum : Form
    {
        public ProductSum()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTable(int EndMon)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            for (int i = 1; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString(), typeof(string));
            }

      //      dt.Columns.Add("Total", typeof(Int64));



            return dt;
        }

        private void PROD()
        {

            string YEAR = comboBox1.Text;
            int EMONTH = Convert.ToInt16(GetYEAR("Account_Temp61" + YEAR).Rows[0][0]);
            System.Data.DataTable dt = MakeTable(EMONTH);
            System.Data.DataTable dtSIZE = GetSIZE("Account_Temp61" + YEAR);
            DataRow dr;
            for (int l = 0; l <= dtSIZE.Rows.Count - 1; l++)
            {
                DataRow dz = dtSIZE.Rows[l];
                for (int i = 0; i <= 2; i++)
                {
                    string TYPE = "";
                    if (i == 0)
                    {
                        TYPE = "數量";
                    }
                    if (i == 1)
                    {
                        TYPE = "金額";
                    }
                    if (i == 2)
                    {
                        TYPE = "毛利";
                    }
                    dr = dt.NewRow();
                    string SIZE = dz["SIZE"].ToString();
                    dr["SIZE"] = SIZE;
                    dr[" "] = TYPE;
                    for (int M = 1; M <= EMONTH; M++)
                    {
                        System.Data.DataTable dh = null;
                        if (i == 0)
                        {
                            dh = GetVALUE1("Account_Temp61" + YEAR, SIZE, M);
                        }
                        if (i == 1)
                        {
                            dh = GetVALUE2("Account_Temp61" + YEAR, SIZE, M);
                        }
                        if (i == 2)
                        {
                            dh = GetVALUE3("Account_Temp61" + YEAR, SIZE, M);
                        }
                        string DHV = dh.Rows[0][0].ToString();
                        if (String.IsNullOrEmpty(DHV))
                        {
                            DHV = "0";
                        }
                        //dr[M.ToString()] = dh.Rows[0][0].ToString();
                        dr[M.ToString()] = Convert.ToDecimal(DHV).ToString("#,##0");

                    }
                    dt.Rows.Add(dr);
                }
            }
            dataGridView1.DataSource = dt;

            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn c = dataGridView1.Columns[i];
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                c.DefaultCellStyle.Format = "#,##0";

            }
        }

        System.Data.DataTable GetYEAR(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
          
            sb.Append(" select  MAX(MONTH(DDATE)) DYEAR  from " + Account_Temp6 + " T0");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetSIZE(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT T1.U_SIZE SIZE,CAST(U_SIZE AS decimal(10,2)) DSIZE  from " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE ItmsGrpCod =1032   AND ISNULL(T1.U_SIZE,'') <>'' AND   ISNUMERIC(T1.U_SIZE) =1");
            sb.Append("  ORDER BY CAST(U_SIZE AS decimal(10,2))");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetVALUE1(string Account_Temp6, string SIZE, int  MONTH)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUM(GQTY) 數量 FROM " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE ItmsGrpCod =1032 AND U_SIZE =@SIZE AND MONTH(DDATE)=@MONTH");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SIZE", SIZE));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetVALUE2(string Account_Temp6, string SIZE, int MONTH)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUM(GTOTAL) 金額 FROM " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE ItmsGrpCod =1032 AND U_SIZE =@SIZE AND MONTH(DDATE)=@MONTH");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SIZE", SIZE));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetVALUE3(string Account_Temp6, string SIZE, int MONTH)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUM(GTOTAL-GVALUE) 毛利  FROM " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OITM T1 ON (T0.ITEMCODE=T1.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE ItmsGrpCod =1032 AND U_SIZE =@SIZE AND MONTH(DDATE)=@MONTH");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SIZE", SIZE));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void ProductSum_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year2017(), "DataValue", "DataValue");
            PROD();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            PROD();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

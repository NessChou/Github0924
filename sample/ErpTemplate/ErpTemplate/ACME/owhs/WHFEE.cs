using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.IO;

namespace ACME
{
    public partial class WHFEE : Form
    {
        public WHFEE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable TempDt = MakeTable();
            System.Data.DataTable dt = GetAddress();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = TempDt.NewRow();

                dr["廠商"] = dt.Rows[i]["廠商"].ToString();
                dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                dr["產品名稱"] = dt.Rows[i]["產品名稱"].ToString();
                dr["金額"] = Convert.ToInt32(dt.Rows[i]["金額"]);
                string SHIPNO = dt.Rows[i]["SHIPNO"].ToString();
                string SHIPNO2 = dt.Rows[i]["SHIPNO2"].ToString();
                string SHIPNO3 = dt.Rows[i]["SHIPNO3"].ToString();
                dr["工單號碼"] = dt.Rows[i]["SHIPNO"].ToString();
                //工單號碼
                dr["倉庫"] = dt.Rows[i]["倉庫"].ToString();
                int G1 = SHIPNO.ToUpper().IndexOf("RMA");
                int G2 = SHIPNO.ToUpper().IndexOf("RMR");
                int G3 = SHIPNO.ToUpper().IndexOf("WH");
                int G4 = SHIPNO.ToUpper().IndexOf("SH");
                int G5 = SHIPNO.ToUpper().IndexOf("X");
                int G6 = SHIPNO.ToUpper().IndexOf("RMA#");
                int G7 = SHIPNO.ToUpper().IndexOf("RMA=");
                if (G1 != -1 && G5!= -1)
                {
                    System.Data.DataTable F1 = GetG1(SHIPNO);
                    if (F1.Rows.Count > 0)
                    {
                        dr["客戶"] = F1.Rows[0][0].ToString();
                    }
                }

                if (G6 != -1 || G7 != -1)
                {
                    System.Data.DataTable F1 = GetG1F(SHIPNO3);
                    if (F1.Rows.Count > 0)
                    {
                        dr["客戶"] = F1.Rows[0][0].ToString();
                    }
                    else
                    {
                        System.Data.DataTable F2 = GetG1F2(SHIPNO3);
                        if (F2.Rows.Count > 0)
                        {
                            dr["客戶"] = F2.Rows[0][0].ToString();
                        }
                    }
                }
                if (G2 != -1)
                {
                    System.Data.DataTable F1 = GetG2(SHIPNO);
                    if (F1.Rows.Count > 0)
                    {
                        dr["客戶"] = F1.Rows[0][0].ToString();
                    }
                }
                if (G3 != -1)
                {
                    System.Data.DataTable F1 = GetG3(SHIPNO2);
                    if (F1.Rows.Count > 0)
                    {
                        dr["客戶"] = F1.Rows[0][0].ToString();
                    }
                }
                if (G4 != -1)
                {
                    System.Data.DataTable F1 = GetG4(SHIPNO);
                    if (F1.Rows.Count > 0)
                    {
                        dr["客戶"] = F1.Rows[0][0].ToString();
                    }
                }
                TempDt.Rows.Add(dr);
            }

            dataGridView1.DataSource = TempDt;
        }

        private DataTable GetAddress()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CardName 廠商,T1.ITEMCODE 產品編號,T1.DSCRIPTION 產品名稱,CAST(T1.LINETOTAL AS INT) 金額,T1.U_Shipping_no SHIPNO,T2.WhsName  倉庫,SUBSTRING(T1.U_Shipping_no,1,14) SHIPNO2, SUBSTRING(REPLACE(REPLACE(T1.U_Shipping_no,'RMA#',''),'RMA=',''),1,8) SHIPNO3");
            sb.Append(" FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WHSCODE=T2.WHSCODE)");
            if (textBox1.Text != "")
            {
                sb.Append(" WHERE  AcctCode =@AcctCode ");
            }
            else
            {
                sb.Append(" WHERE SUBSTRING(T1.ITEMCODE,1,1)='Z'");
                if (checkBox1.Checked)
                {
                    sb.Append(" AND T1.U_Shipping_no LIKE '%RMA%'");
                }
            }
            sb.Append("  AND YEAR(T0.DOCDATE)=@YEAR");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@AcctCode", textBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private DataTable GetG1F(string RMANO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CodeName  FROM Rma_InvoiceD WHERE RMANO=@RMANO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RMANO", RMANO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetG1F2(string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_Cusname_S  FROM OCTR WHERE U_RMA_NO=@U_RMA_NO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private DataTable GetG1(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CODENAME FROM Rma_InvoiceD WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private DataTable GetG2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM Rma_mainR WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }

        private DataTable GetG3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM WH_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }


        private DataTable GetG4(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("廠商", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("金額", typeof(int));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            return dt;
        }

        private void WHFEE_Load(object sender, EventArgs e)
        {
            textBox1.Text = "62110102";
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox1.Text = "";
            }
        }
    }
}

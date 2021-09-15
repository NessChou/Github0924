using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class SHIPFEE : Form
    {
        public SHIPFEE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        //    dataGridView1.DataSource = DT();

            System.Data.DataTable dt = DT();

            System.Data.DataTable dtCost = MakeTableCombine();

            DataRow dr = null;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = dtCost.NewRow();
                string 工單 = Convert.ToString(dt.Rows[i]["工單"]);
                dr["工單"] = 工單;
                dr["客人名稱"] = Convert.ToString(dt.Rows[i]["客人名稱"]);
                dr["借項金額"] = Convert.ToDecimal(DT21(工單).Rows[0][0]);

                dr["貸項金額"] = Convert.ToDecimal(DT31(工單).Rows[0][0]) + Convert.ToDecimal(GetFEE2(工單).Rows[0][0]);
                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;
        }

        private System.Data.DataTable DT()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT T0.SHIPPINGCODE 工單,CARDNAME 客人名稱 FROM SHIPPING_MAIN T0");
            sb.Append(" LEFT JOIN SHIPPING_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE DSCRIPTION LIKE '%運費%' AND T0.quantity='已結' ");
            sb.Append(" AND SUBSTRING(T0.SHIPPINGCODE,3,4)=@YEAR");
            if (comboBox3.Text != "")
            {
                sb.Append(" AND CAST(SUBSTRING(T0.SHIPPINGCODE,7,2) AS INT)=@MONTH");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", comboBox3.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT QUANTITY 數量,ITEMPRICE 單價,ITEMAMOUNT 金額 FROM SHIPPING_ITEM WHERE DSCRIPTION LIKE '%運費%'");
            sb.Append(" AND SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT21(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(ITEMAMOUNT) 金額 FROM SHIPPING_ITEM WHERE DSCRIPTION LIKE '%運費%'");
            sb.Append(" AND SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T0.[CardName], CAST(T0.[DocNum] AS VARCHAR) DocNum, T1.[Dscription], T1.[LineTotal],T2.[SLPNAME] FROM  acmesql02.dbo.[OPOR]  T0 INNER JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry LEFT JOIN acmesql02.dbo.OSLP T2 ON T0.SLPCODE = T2.SLPCODE WHERE T0.[U_Shipping_no] =@shippingcode or T1.[U_Shipping_no]=@shippingcode ");
            sb.Append(" union all");
            sb.Append(" SELECT  '' CardName, '' DocNum, '小計' Dscription, isnull(sum(T1.LINETOTAL),0)");
            sb.Append("  aa ,'' SLPNAME");
            sb.Append(" FROM  acmesql02.dbo.[OPor]  T0 ");
            sb.Append(" left join acmesql02.dbo.[Por1]  T1 on (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T0.[U_Shipping_no] =@shippingcode or T1.[U_Shipping_no]=@shippingcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT31(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT  isnull(sum(T1.LINETOTAL),0)  AMT");
            sb.Append("               FROM  acmesql02.dbo.[OPor]  T0  ");
            sb.Append("               left join acmesql02.dbo.[Por1]  T1 on (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("               WHERE T0.[U_Shipping_no] =@shippingcode or T1.[U_Shipping_no]=@shippingcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetFEE(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CardName 供應商,SubCompany 子公司,DocDate 日期,SAP SAP單號,ITEM 費用名稱,Amount 金額,DocCur 幣別,DocCur1 匯率 FROM dbo.Shipping_Fee   where ShippingCode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }

        public static System.Data.DataTable GetFEE2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ISNULL(SUM(cast(AMOUNT AS DECIMAL(15,5))),0) AMOUNT FROM shipping_fee T0 where T0.[shippingcode]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


             
                        row = dataGridView1.SelectedRows[0];
                        string SHIPPINGCODE = row.Cells["工單"].Value.ToString();

                        dataGridView2.DataSource = DT2(SHIPPINGCODE);
                        dataGridView4.DataSource = DT3(SHIPPINGCODE);
                        dataGridView3.DataSource = GetFEE(SHIPPINGCODE);
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {


            ExcelReport.GridViewToExcel(dataGridView1);
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("工單", typeof(string));
            dt.Columns.Add("客人名稱", typeof(string));
            dt.Columns.Add("借項金額", typeof(decimal));
            dt.Columns.Add("貸項金額", typeof(decimal));
            return dt;
        }
        private void SHIPFEE_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Year(), "DataValue", "DataValue");

            comboBox2.Text = DateTime.Now.ToString("yyyy");

        }


    }
}

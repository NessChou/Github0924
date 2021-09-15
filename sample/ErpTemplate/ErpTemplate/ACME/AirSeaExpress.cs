using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ACME
{
    public partial class AirSeaExpress : Form
    {
        public AirSeaExpress()
        {
            InitializeComponent();
            ComboBoxInit();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetAirSeaExpress();
            dgvAirSeaExpress.DataSource = dt;

        }
        public System.Data.DataTable GetAirSeaExpress()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  T1.[U_Shipping_no] ,T0.[CardCode], T0.[CardName], CAST(T0.[DocNum] AS VARCHAR) DocNum, T1.[Dscription], T1.[LineTotal],T2.[SLPNAME] ");
            sb.Append("  ,T3.CloseDay 結關日,add9 報單號碼,receiveDay 運送方式,sendGoods '併櫃/CBM',kPIMEMO 拆提貨天數  ");
            sb.Append(" FROM [OPOR]  T0  ");
            sb.Append(" INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" LEFT JOIN OSLP T2 ON T0.SLPCODE = T2.SLPCODE  ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SHIPPING_MAIN T3 ON (T1.[U_Shipping_no]=T3.SHIPPINGCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" WHERE   T1.[U_Shipping_no] <> '' ");
            if (txbCardCode1.Text != "") 
            {
                sb.Append(" and (T0.CardCode like '"+ txbCardCode1.Text + "')");
            }
            if (cmbCardCode2.Text != "")
            {
                sb.Append(" and (T0.CardCode like '" + cmbCardCode2.Text + "')");
            }
            if (txbCloseDay.Text != "")
            {
                sb.Append(" and (T3.CloseDay = '" + txbCloseDay.Text + "')");
            }
            if (cmbReceiveDay.Text != "") 
            {
                sb.Append(" and (T3.receiveDay = '" + cmbReceiveDay.Text + "')");
            }
            
            sb.Append(" ORDER BY  T1.[U_Shipping_no]   ");

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
        private System.Data.DataTable GetcmbCardCode1()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  T1.[U_Shipping_no] ,T0.[CardCode], T0.[CardName], CAST(T0.[DocNum] AS VARCHAR) DocNum, T1.[Dscription], T1.[LineTotal],T2.[SLPNAME] ");
            sb.Append("  ,T3.CloseDay 結關日,add9 報單號碼,receiveDay 運送方式,sendGoods '併櫃/CBM',kPIMEMO 拆提貨天數  ");
            sb.Append(" FROM [OPOR]  T0  ");
            sb.Append(" ORDER BY  T1.[U_Shipping_no]   ");

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
        private System.Data.DataTable GetcmbCardCode2()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  T1.[U_Shipping_no] ,T0.[CardCode], T0.[CardName], CAST(T0.[DocNum] AS VARCHAR) DocNum, T1.[Dscription], T1.[LineTotal],T2.[SLPNAME] ");
            sb.Append("  ,T3.CloseDay 結關日,add9 報單號碼,receiveDay 運送方式,sendGoods '併櫃/CBM',kPIMEMO 拆提貨天數  ");
            sb.Append(" FROM [OPOR]  T0  ");
            sb.Append(" ORDER BY  T1.[U_Shipping_no]   ");

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
        private System.Data.DataTable GETcmbReceiveDay()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  T1.[U_Shipping_no] ,T0.[CardCode], T0.[CardName], CAST(T0.[DocNum] AS VARCHAR) DocNum, T1.[Dscription], T1.[LineTotal],T2.[SLPNAME] ");
            sb.Append("  ,T3.CloseDay 結關日,add9 報單號碼,receiveDay 運送方式,sendGoods '併櫃/CBM',kPIMEMO 拆提貨天數  ");
            sb.Append(" FROM [OPOR]  T0  ");
            sb.Append(" ORDER BY  T1.[U_Shipping_no]   ");

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
        private void ComboBoxInit() 
        {

        }

        private void btnGetCardCode1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

            LookupValues = GetMenu.GetMenuList();
            txbCardCode1.Text = Convert.ToString(LookupValues[0]);
            txbCardName.Text = Convert.ToString(LookupValues[1]);

        }
    }
}

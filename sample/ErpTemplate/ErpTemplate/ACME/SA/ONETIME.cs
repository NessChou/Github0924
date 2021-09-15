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
    public partial class ONETIME : Form
    {
        public ONETIME()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = GetONE1();
            dataGridView2.DataSource = GetONE2();
            dataGridView3.DataSource = GetONE3();
            dataGridView4.DataSource = GetONE4();
        
        }
        private System.Data.DataTable MakeTabe()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("PN", typeof(string));
            dt.Columns.Add("PO", typeof(string));
            dt.Columns.Add("Supplier", typeof(string));
            dt.Columns.Add("Buyer Item", typeof(string));
            dt.Columns.Add("SO", typeof(string));
            dt.Columns.Add("Customer", typeof(string));
            dt.Columns.Add("SA", typeof(string));
            dt.Columns.Add("Sa Item", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("Price", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("進貨日期", typeof(string));
            return dt;
        }

        private System.Data.DataTable GetTABLE()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE PN,CAST(ONHAND AS INT) ONHAND FROM OITM WHERE ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'  AND ONHAND > 0");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetTABLE2(string ITEMCODE, int QUANTITY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT MAX(DOCENTRY) FROM RDR1 WHERE ITEMCODE=@ITEMCODE AND CAST(QUANTITY AS INT)=@QUANTITY AND LINESTATUS='O'  GROUP BY ITEMCODE ");
   

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QUANTITY", QUANTITY));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetTABLE21(string DOCENTRY, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                 SELECT (T2.[lastName]+T2.[firstName]) SA,CARDNAME Customer,T1.DSCRIPTION ITEM  ");
            sb.Append("                             FROM ORDR T0  ");
            sb.Append("                             LEFT JOIN RDR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("                             LEFT JOIN OHEM T2 ON (T0.OwnerCode =T2.EMPID)  ");
            sb.Append("                             WHERE T0.DOCENTRY=@DOCENTRY AND T1.ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetTABLE3(string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 1 (ISNULL(T5.DOCENTRY,''))    PO,Convert(varchar(8),T0.DOCDATE,112)  DOCDATE");
            sb.Append(" FROM OPDN T0");
            sb.Append(" LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" left join POR1 t5 on (T1.baseentry=T5.docentry and  T1.baseline=T5.linenum and t5.TARGETTYPE='20'  )");
            sb.Append("  WHERE T1.ITEMCODE=@ITEMCODE  ORDER BY T0.DOCDATE DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetTABLE31(string DOCENTRY, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                        SELECT  CARDNAME Supplier,CAST(QUANTITY AS INT) QTY,CAST(Price AS DECIMAL(10,4)) Price ,U_MEMO 備註,Convert(varchar(8),T0.DOCDATE,112)  進貨日期,T1.DSCRIPTION ITEM ");
            sb.Append("                             FROM OPOR T0  ");
            sb.Append("                             LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("                             WHERE T0.DOCENTRY=@DOCENTRY AND T1.ITEMCODE=@ITEMCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetTABLE4()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT T0.ITEMCODE  FROM OITM T0 ");
            sb.Append("               LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM RDR1  WHERE LINESTATUS='O' ");
            sb.Append("               GROUP BY ITEMCODE) T1 ON (T0.ITEMCODE=T1.ITEMCODE) ");
            sb.Append("               LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM POR1  WHERE LINESTATUS='O' ");
            sb.Append("               GROUP BY ITEMCODE) T2 ON (T0.ITEMCODE=T2.ITEMCODE) ");
            sb.Append("                WHERE T0.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020' ");
            sb.Append(" and ONHAND >0");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public System.Data.DataTable GetONE1()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ITEMCODE '一次性料號',ISNULL(CAST(T2.QTY AS INT),0) '採購單數量',ISNULL(CAST(T1.QTY AS INT),0) '銷售訂單數量', ");
            sb.Append(" CAST(ONHAND AS INT) '庫存數量'  FROM OITM T0");
            sb.Append(" LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM RDR1  WHERE LINESTATUS='O'");
            sb.Append(" GROUP BY ITEMCODE) T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM POR1  WHERE LINESTATUS='O'");
            sb.Append(" GROUP BY ITEMCODE) T2 ON (T0.ITEMCODE=T2.ITEMCODE)");
            sb.Append("  WHERE T0.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'");
            if (checkBox1.Checked)
            {
                sb.Append(" AND (ISNULL(CAST(T2.QTY AS INT),0)=0 AND ISNULL(CAST(T1.QTY AS INT),0)=0 AND CAST(ONHAND AS INT)=0) ");
            }
            else
            {
                if (textBox1.Text != "")
                {
                    sb.Append("  and T0.ITEMCODE =@ITEMCODE ");
                }
            }



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", textBox1.Text));

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

        public System.Data.DataTable GetONE2()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE '一次性料號',T2.SLPNAME Buyer,CARDCODE 廠商,CARDNAME,T0.DOCENTRY 'PO NO.'");
            sb.Append(" ,T1.DSCRIPTION 品名規格,ISNULL(CAST(T1.QUANTITY AS INT),0) 數量 ");
            sb.Append(" FROM OPOR T0");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OSLP T2 ON (T0.SLPCODE=T2.SLPCODE)");
            sb.Append(" WHERE T1.LINESTATUS='O'");
            sb.Append(" AND T1.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'");

            if (checkBox1.Checked)
            {
                sb.Append(" AND  1 = 2 ");
            }
            else
            {
                if (textBox1.Text != "")
                {
                    sb.Append("  and T1.ITEMCODE =@ITEMCODE ");
                }
            }
            sb.Append(" ORDER BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", textBox1.Text));

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

        public System.Data.DataTable GetONE3()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE '一次性料號',(T2.[lastName]+T2.[firstName]) SA,CARDCODE 客戶,CARDNAME,T0.DOCENTRY 'SO NO.'");
            sb.Append(" ,T1.DSCRIPTION 品名規格,ISNULL(CAST(T1.QUANTITY AS INT),0) 數量 ");
            sb.Append(" FROM ORDR T0");
            sb.Append(" LEFT JOIN RDR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OHEM T2 ON (T0.OwnerCode =T2.EMPID)");
            sb.Append(" WHERE T1.LINESTATUS='O'");
            sb.Append(" AND T1.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'");

            if (checkBox1.Checked)
            {
                sb.Append(" AND  1 = 2 ");
            }
            else
            {
                if (textBox1.Text != "")
                {
                    sb.Append("  and T1.ITEMCODE =@ITEMCODE ");
                }
            }

            sb.Append(" ORDER BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", textBox1.Text));

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


        public System.Data.DataTable GetONE4()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ITEMCODE '一次性料號',WHSNAME 庫存倉別,T1.DSCRIPTION 品名規格,T5.DOCENTRY 採購單,");
            sb.Append(" T0.DOCDATE '進貨日期',ISNULL(CAST(T1.QUANTITY AS INT),0) 數量");
            sb.Append(" FROM OPDN T0");
            sb.Append(" LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OWHS T2 ON T2.whsCode = T1.whscode");
            sb.Append(" LEFT JOIN OITM T3 ON T1.ITEMCODE = T3.ITEMCODE");
            sb.Append(" left join POR1 t5 on (T1.baseentry=T5.docentry and  T1.baseline=T5.linenum and t5.TARGETTYPE='20'  )");
            sb.Append(" WHERE T1.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'");

            if (checkBox1.Checked)
            {
                sb.Append(" AND  T1.ITEMCODE IN (SELECT T0.ITEMCODE   FROM OITM T0");
                sb.Append(" LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM RDR1  WHERE LINESTATUS='O'");
                sb.Append(" GROUP BY ITEMCODE) T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
                sb.Append(" LEFT JOIN (SELECT SUM(QUANTITY) QTY,ITEMCODE FROM POR1  WHERE LINESTATUS='O'");
                sb.Append(" GROUP BY ITEMCODE) T2 ON (T0.ITEMCODE=T2.ITEMCODE)");
                sb.Append("  WHERE T0.ITEMCODE BETWEEN 'ACME00001.00001' AND 'ACME00021.00020'");
                sb.Append(" AND (ISNULL(CAST(T2.QTY AS INT),0)=0 AND ISNULL(CAST(T1.QTY AS INT),0)=0 AND CAST(ONHAND AS INT)=0)");
                sb.Append(" )");
            }
            else
            {
                sb.Append("  AND CAST(ONHAND AS INT) <> 0 ");
                if (textBox1.Text != "")
                {
                    sb.Append("  and T1.ITEMCODE =@ITEMCODE ");
                }
            }
            sb.Append(" ORDER BY T1.ITEMCODE,T0.DOCDATE desc");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", textBox1.Text));

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

        private void ONETIME_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetONE1();
            dataGridView2.DataSource = GetONE2();
            dataGridView3.DataSource = GetONE3();
            dataGridView4.DataSource = GetONE4();

            System.Data.DataTable dt = GetTABLE();
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTabe();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string PN = dt.Rows[i]["PN"].ToString();

                string ONHAND = dt.Rows[i]["ONHAND"].ToString();

                string SO = "";
                string SA = "";
                string Customer = "";
                string SITEM = "";
                string PO = "";
                string DOCDATE = "";
                string ITEM = "";
                string Supplier = "";
                string QTY = "";
                string Price = "";
                string 備註 = "";
               // string 進貨日期 = DOCDATE;

                System.Data.DataTable dt3 = GetTABLE3(PN);
                if (dt3.Rows.Count > 0)
                {
                    PO = dt3.Rows[0][0].ToString();
                    DOCDATE = dt3.Rows[0][1].ToString();
                }
                System.Data.DataTable dt31 = GetTABLE31(PO, PN);
                if (dt31.Rows.Count > 0)
                {
                    ITEM = dt31.Rows[0]["ITEM"].ToString();
                    Supplier = dt31.Rows[0]["Supplier"].ToString();
                    QTY = dt31.Rows[0]["QTY"].ToString();
                    Price = dt31.Rows[0]["Price"].ToString();
                    備註 = dt31.Rows[0]["備註"].ToString();


                    System.Data.DataTable dt2 = GetTABLE2(PN, Convert.ToInt16(QTY));
                    if (dt2.Rows.Count > 0)
                    {
                        SO = dt2.Rows[0][0].ToString();
                        System.Data.DataTable dt21 = GetTABLE21(SO, PN);
                        SA = dt21.Rows[0]["SA"].ToString();
                        Customer = dt21.Rows[0]["Customer"].ToString();
                        SITEM = dt21.Rows[0]["ITEM"].ToString();
                    }
                }

            

                dr["PN"] = PN;
                dr["SA"] = SA;
                dr["PO"] = PO;
                dr["SO"] = SO;
                dr["Customer"] = Customer;
                dr["Supplier"] = Supplier;
                dr["QTY"] = ONHAND;
                dr["Price"] = Price;
                dr["備註"] = 備註;
                dr["進貨日期"] = DOCDATE;
                dr["Buyer Item"] = ITEM;
                dr["Sa Item"] = SITEM;
                dtCost.Rows.Add(dr);

            }
            dataGridView5.DataSource = dtCost;
        }

        private void button2_Click(object sender, EventArgs e)
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
                ExcelReport.GridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToExcel(dataGridView5);
            }
        }


    }
}
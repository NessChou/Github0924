using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace ACME
{
    public partial class AUOPOS1 : Form
    {
        public AUOPOS1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TRUAUOGD();

            System.Data.DataTable dt = GetOrderData20();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                //收貨單單號
                string aa = dt.Rows[i]["AUINVOICE"].ToString();
                string bb = dt.Rows[i]["收貨單單號"].ToString();
                string[] sArray = aa.Split(new char[] { '/' });
                int s = 0;
                foreach (string j in sArray)
                {
                    AddAUOGD(s.ToString(), bb.ToString(), j.ToString());
                    s++;
                }
            }

            ViewBatchPayment2();
        }
        private System.Data.DataTable GetOrderData20()
        {
            String strCn = "Data Source=acmesrv13;Initial Catalog=acmesql95;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT t0.u_acme_inv 'AUINVOICE',T0.DOCENTRY 收貨單單號  FROM ODLN T0 where  t0.u_acme_inv like '%/%' and docentry > '2512' ");
            //    sb.Append(" WHERE T0.DOCENTRY='2513' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //      command.Parameters.Add(new SqlParameter("@aa", aa));

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
        private void ViewBatchPayment2()
        {
            String strCn = "Data Source=acmesrv13;Initial Catalog=acmesql95;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            SqlConnection MyConnection = new SqlConnection(strCn);

           // SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select t7.docentry ,CASE (Substring(T1.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append(" when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append(" when '4' then 'U' when '5' then 'NN' ELSE 'X'");
            sb.Append(" END 等級,Substring(T1.[ItemCode],2,8) Model,(t3.param_desc) [Size],'V.'+Substring(T1.[ItemCode],12,1) 版本");
            sb.Append(" ,(t2.numatcard) PO,(t0.u_acme_inv) INVOICE");
            sb.Append(" ,t2.cardname 客戶,cast((t1.quantity) as int) 賣進數量,cast((t2.quantity) as int) 賣出數量,");
            sb.Append(" cast((T7.PRICE) as int) 買進價格,cast((T6.PRICE) as int) 賣出價格 ");
            sb.Append(" from opdn t0");
            sb.Append(" LEFT JOIN pdn1 t1 on (t0.docentry=t1.docentry)");
            sb.Append(" INNER JOIN POR1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append(" INNER JOIN (SELECT T0.u_acme_inv u_acme_inv,T1.baseentry baseentry,");
            sb.Append(" T1.baseline baseline,T1.ITEMCODE ITEMCODE,T0.CARDNAME CARDNAME");
            sb.Append(" ,t0.numatcard numatcard,T1.QUANTITY QUANTITY FROM ODLN T0 LEFT JOIN DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)) T2");
            sb.Append(" on (t0.u_acme_inv=t2.u_acme_inv AND T1.ITEMCODE=T2.ITEMCODE)");
            sb.Append(" INNER JOIN RDR1 T6 ON (T6.docentry=T2.baseentry AND T6.linenum=T2.baseline)");
            sb.Append(" LEFT JOIN acmesqlsp.dbo.params t3 on(substring(t1.itemcode,3,3)=t3.param_no COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where substring(T1.itemcode,1,1)='T'  ");
            sb.Append(" and year(t0.docdate)='" + comboBox1.SelectedValue.ToString() + "' and month(t0.docdate)='" + comboBox2.SelectedValue.ToString() + "' AND  substring(t0.cardname,11,12) ='" + comboBox3.SelectedValue.ToString() + "'  ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT T7.DOCENTRY,CASE (Substring(T1.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append(" when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append(" when '4' then 'U' when '5' then 'NN' ELSE 'X'");
            sb.Append(" END 等級,Substring(T1.[ItemCode],2,8) Model,(t3.param_desc) [Size],'V.'+Substring(T1.[ItemCode],12,1) 版本,");
            sb.Append(" (t2.numatcard) PO,(t0.u_acme_inv) INVOICE,");
            sb.Append(" t0.cardname 客戶,cast((t1.quantity) as int) 賣進數量,cast((t2.quantity) as int) 賣出數量,");
            sb.Append(" cast((T7.PRICE) as int) 買進價格,cast((T6.PRICE) as int) 賣出價格");
            sb.Append(" FROM OPDN T0");
            sb.Append(" LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" INNER JOIN POR1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN (select T0.numatcard,T0.DOCENTRY,T1.ITEMCODE,T1.QUANTITY,T2.INVOICE,T1.baseentry,T1.baseline from ODLN T0 ");
            sb.Append("            LEFT JOIN DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("            LEFT JOIN ACMESQLSP.DBO.AP_TEMP T2 ON (T2.DOCENTRY=T1.DOCENTRY AND T2.LINENUM=T1.LINENUM)");
            sb.Append("            ) T2 ON (T2.ITEMCODE=T1.ITEMCODE AND T2.INVOICE=t0.u_acme_inv COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" LEFT JOIN RDR1 T6 ON (T6.docentry=T2.baseentry AND T6.linenum=T2.baseline)");
            sb.Append(" LEFT JOIN acmesqlsp.dbo.params t3 on(substring(t1.itemcode,3,3)=t3.param_no COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE T1.DOCENTRY>2393 and substring(T1.itemcode,1,1)='T'  AND T2.DOCENTRY IS NOT NULL");
            sb.Append(" and year(t0.docdate)='" + comboBox1.SelectedValue.ToString() + "' and month(t0.docdate)='" + comboBox2.SelectedValue.ToString() + "' AND  substring(t0.cardname,11,12) ='" + comboBox3.SelectedValue.ToString() + "'  ");
            sb.Append(" ORDER BY T7.DOCENTRY");
  
       //     sb.Append("                       group by Substring(T1.[ItemCode],2,8),t2.cardname,Substring(T1.[ItemCode],12,1),t7.docentry,Substring(T1.[ItemCode],11,1)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void AUOPOS_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetBU("shipyear"), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, Getmonth(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetBU("BU"), "DataValue", "DataValue");
        }
        System.Data.DataTable GetBU(string KIND)
        {
            SqlConnection con = globals.CommonConnection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by DataValue";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
               //con.Dispose();
            }
            return ds.Tables["RMA_PARAMS"];
        }

        System.Data.DataTable Getmonth()
        {
            SqlConnection con = globals.CommonConnection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='shipmonth' order by cast(PARAM_NO as int)";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
              //  con.Dispose();
            }
            return ds.Tables["RMA_PARAMS"];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GridViewToExcel(dataGridView1);
        }

        public void AddAUOGD(string Linenum, string Docentry, string Invoice)
        {
            String strCn = "Data Source=acmesrv13;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            SqlConnection connection = new SqlConnection(strCn);

            SqlCommand command = new SqlCommand("Insert into AP_Temp(Linenum,Docentry,Invoice) values(@Linenum,@Docentry,@Invoice)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Linenum", Linenum));
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
            command.Parameters.Add(new SqlParameter("@Invoice", Invoice));

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
        public void TRUAUOGD()
        {
            String strCn = "Data Source=acmesrv13;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

            SqlConnection connection = new SqlConnection(strCn);

            SqlCommand command = new SqlCommand("Truncate table AP_Temp ", connection);
            command.CommandType = CommandType.Text;

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
        private void GridViewToExcel(DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;

            try
            {

                int iX;

                int iY;

                for (int i = 0; i < dgv.Columns.Count; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

                    // wsheet.Font.Bold = true;

                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                        }

                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);

                        }

                    }

                }

                wapp.Visible = true;


            }

            catch (Exception ex1)
            {

                MessageBox.Show(ex1.Message);

            }

            wapp.UserControl = true;


        }
    }
}
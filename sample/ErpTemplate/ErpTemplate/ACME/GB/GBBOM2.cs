using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class GBBOM2 : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBBOM2()
        {
            InitializeComponent();
        }

        private void GBBOM2_Load(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dtCost = MakeTableCombine();
                DataRow dr = null;
                System.Data.DataTable DT1 = GetCHO3();
                string DUP = "";
                for (int i = 0; i <= DT1.Rows.Count - 1; i++)
                {
                    DataRow dd = DT1.Rows[i];
                    dr = dtCost.NewRow();
                    string ITEMCODE = dd["組合料號"].ToString();
                    dr["組合料號"] = ITEMCODE;
                    dr["品名規格"] = dd["品名規格"].ToString();
                    dr["發票品名"] = dd["發票品名"].ToString();
                    dr["組合品項"] = dd["組合品項"].ToString();

                    dr["子料號"] = dd["子料號"].ToString();
                    dr["子發票品名"] = dd["子發票品名"].ToString();
                    dr["子數量"] = Convert.ToInt32(dd["子數量"]);
                    dr["子成本"] = dd["子成本"].ToString();
                    dr["子售價"] = dd["子售價"].ToString();
                 
                        System.Data.DataTable T1 = GetCHO4(ITEMCODE);
                        if (T1.Rows.Count > 0)
                        {
                            dr["數量"] = Convert.ToInt32(T1.Rows[0]["數量"]);
                            if (DUP != ITEMCODE)
                            {
                                dr["成本"] = T1.Rows[0]["成本"].ToString();
                                dr["建議售價"] = T1.Rows[0]["建議售價"].ToString();
                                dr["毛利"] = Convert.ToDecimal(T1.Rows[0]["毛利"]);
                            }
                        }
                    
                    DUP = ITEMCODE;
                    dtCost.Rows.Add(dr);
                }
                dataGridView1.DataSource = dtCost;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("組合料號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("發票品名", typeof(string));
            dt.Columns.Add("數量", typeof(int));
            dt.Columns.Add("成本", typeof(decimal));
            dt.Columns.Add("組合品項", typeof(string));
            dt.Columns.Add("建議售價", typeof(string));
            dt.Columns.Add("毛利", typeof(decimal));
            dt.Columns.Add("子料號", typeof(string));
            dt.Columns.Add("子發票品名", typeof(string));
            dt.Columns.Add("子數量", typeof(int));
            dt.Columns.Add("子成本", typeof(decimal));
            dt.Columns.Add("子售價", typeof(string));
            return dt;
        }
        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.ProdID 組合料號,T0.ProdName 品名規格,T0.InvoProdName 發票品名, CASE WHEN K.ClassID='ACME M' THEN '豬' ");
            sb.Append(" WHEN K.ClassID='ACMECM' THEN '雞' WHEN K.ClassID='ACMEFR' THEN '運費'      ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='S' THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'       ");
            sb.Append(" WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬' WHEN SUBSTRING(K.ClassID,3,1)='G' THEN '禮盒' WHEN SUBSTRING(T1.CombSubID,1,1)='P' THEN '加工品'        ");
            sb.Append(" END 組合品項,T1.CombSubID 子料號,J.InvoProdName 子發票品名,T1.Amount 子數量,J.CAvgCost 子成本,J.SalesPriceA  子售價");
            sb.Append(" from comProduct T0");
            sb.Append("  INNER JOIN comProdCombine T1 ON (T0.ProdID=T1.ProdID)");
            sb.Append("     INNER Join comProduct J On (T1.CombSubID =J.ProdID) ");
            sb.Append("       INNER Join comProductClass K On J.ClassID=K.ClassID ");
            sb.Append("       WHERE SUBSTRING(T0.ProdID,1,1) > 'F'");
     
  
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
        public System.Data.DataTable GetCHO4(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select SUM(J.CAvgCost) 成本,MAX(T0.SalesPriceA) 建議售價,MAX(T0.SalesPriceA)-SUM(J.CAvgCost) 毛利,SUM(AMOUNT) 數量");
            sb.Append(" from comProduct T0");
            sb.Append("  INNER JOIN comProdCombine T1 ON (T0.ProdID=T1.ProdID)");
            sb.Append("     INNER Join comProduct J On (T1.CombSubID =J.ProdID) ");
            sb.Append("       WHERE T0.ProdID =@ProdID");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

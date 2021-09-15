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
    public partial class ACCFIRSTEND : Form
    {
        string strCn = "";
        public ACCFIRSTEND()
        {
            InitializeComponent();
        }
        public string d;
        private void button4_Click(object sender, EventArgs e)
        {
            APS2CHOICE frm1 = new APS2CHOICE();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = MakeTableCombine();
            if (comboBox1.Text == "聿豐")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            if (comboBox1.Text == "東門")
            {
                strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }

            DataRow dr = null;
            System.Data.DataTable DT1 = GetCHO3();

            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                DataRow dd = DT1.Rows[i];
                int 庫存量 = Convert.ToInt32(dd["庫存量"]);
                int 期初存貨 = Convert.ToInt32(dd["期初存貨"]);
                int 本期進貨 = Convert.ToInt32(dd["本期進貨"]);
                int 進貨折讓 = Convert.ToInt32(dd["進貨折讓"]);
                int 進貨退出 = Convert.ToInt32(dd["進貨退出"]);
                int 本期銷貨 = Convert.ToInt32(dd["本期銷貨"]);
                int 銷貨退回 = Convert.ToInt32(dd["銷貨退回"]);
                int 本期調整 = Convert.ToInt32(dd["本期調整"]);
                int 本期調撥 = Convert.ToInt32(dd["本期調撥"]);
                int 期末存貨 = 期初存貨 + 本期進貨 - 進貨折讓 - 進貨退出 - 本期銷貨 + 銷貨退回 + 本期調整 + 本期調撥;
                if (庫存量 != 0 || 期初存貨 != 0 || 本期進貨 != 0 || 進貨折讓 != 0 || 進貨退出 != 0 || 本期銷貨 != 0 || 銷貨退回 != 0 || 本期調整 != 0 || 本期調撥 != 0 || 期末存貨 != 0)
                {
                    dr = dtCost.NewRow();
                    dr["產品編號"] = dd["產品編號"].ToString();
                    dr["品名規格"] = dd["品名規格"].ToString();
                    dr["產品類別"] = dd["產品類別"].ToString();
                    dr["庫存量"] = 庫存量;
                    dr["期初存貨"] = 期初存貨;
                    dr["本期進貨"] = 本期進貨;
                    dr["進貨折讓"] = 進貨折讓;
                    dr["進貨退出"] = 進貨退出;
                    dr["本期銷貨"] = 本期銷貨;
                    dr["銷貨退回"] = 銷貨退回;
                    dr["本期調整"] = 本期調整;
                    dr["本期調撥"] = 本期調撥;
                    dr["期末存貨"] = 期末存貨;
                    dtCost.Rows.Add(dr);
                }

            }

            //G3
            decimal[] TotalG = new decimal[dtCost.Columns.Count - 1];

            for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
            {

                for (int j = 3; j <= 12; j++)
                {
                    TotalG[j - 1] += Convert.ToDecimal(dtCost.Rows[i][j]);

                }
            }

            DataRow rowG;

            rowG = dtCost.NewRow();

            rowG[2] = "合計";

            for (int j = 3; j <= 12; j++)
            {
                rowG[j] = TotalG[j - 1];

            }

            dtCost.Rows.Add(rowG);

            dataGridView1.DataSource = dtCost;

            for (int i = 3; i <= 12; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("庫存量", typeof(int));
            dt.Columns.Add("期初存貨", typeof(int));
            dt.Columns.Add("本期進貨", typeof(int));
            dt.Columns.Add("進貨折讓", typeof(int));
            dt.Columns.Add("進貨退出", typeof(int));
            dt.Columns.Add("本期銷貨", typeof(int));
            dt.Columns.Add("銷貨退回", typeof(int));
            dt.Columns.Add("本期調整", typeof(int));
            dt.Columns.Add("本期調撥", typeof(int));
            dt.Columns.Add("期末存貨", typeof(int));
            return dt;
        }

        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select A.ProdID 產品編號, A.ProdName 品名規格,K.ClassName 產品類別");
sb.Append(" ,CAST((select  ISNULL(SUM(CASE WHEN FLAG IN (500) THEN Quantity *-1 ELSE Quantity END),0) 庫存量 from comProdRec M where M.prodid= A.ProdID and BillDate < @DATE1 AND FLAG NOT IN (400,701)) AS INT) 庫存量");
sb.Append(" ,CAST((select  ISNULL(SUM(CASE WHEN FLAG IN (500) THEN CostForAcc *-1 ELSE CostForAcc END),0) 庫存量 from comProdRec M where M.prodid= A.ProdID and BillDate < @DATE1 AND FLAG NOT IN (400,701)) AS INT)  期初存貨");
sb.Append(" ,CAST((Select +IsNull(Sum(M.CostForAcc),0) From comProdRec  M Where M.Flag not in (318,319,320) and M.Flag Between 100 and 199 and M.HasCheck = 1 and M.YearCompressType<>1  and M.ProdID = A.ProdID  and M.WareID <> '' and M.NeedUpdate = 1  and  (M.BillDate Between @DATE1 And @DATE2) ) AS INT) As 本期進貨");
sb.Append(" ,CAST((Select +IsNull(Sum(Q.MLDist),0) From comProdRec  Q Where Q.Flag not in (318,319,320) and Q.Flag = 700 and Q.HasCheck = 1 and Q.YearCompressType<>1  and Q.ProdID = A.ProdID  and Q.WareID <> '' and Q.NeedUpdate = 1  and  (Q.BillDate Between @DATE1 And @DATE2) ) AS INT) As 進貨折讓");
sb.Append(" ,CAST((Select +IsNull(Sum(N.CostForAcc),0) From comProdRec  N Where N.Flag not in (318,319,320) and N.Flag Between 200 and 299 and N.HasCheck = 1 and N.YearCompressType<>1  and N.ProdID = A.ProdID  and N.WareID <> '' and N.NeedUpdate = 1  and  (N.BillDate Between @DATE1 And @DATE2) ) AS INT) As 進貨退出");
sb.Append(" ,CAST((Select +IsNull(Sum(I.CostForAcc),0) From comProdRec  I Where I.Flag not in (318,319,320) and I.Flag Between 500 and 599 and I.HasCheck = 1 and I.YearCompressType<>1  and I.ProdID = A.ProdID  and I.WareID <> '' and I.NeedUpdate = 1  and  (I.BillDate Between @DATE1 And @DATE2) ) AS INT) As 本期銷貨");
sb.Append(" ,CAST((Select +IsNull(Sum(J.CostForAcc),0) From comProdRec  J Where J.Flag not in (318,319,320) and J.Flag Between 600 and 699 and J.HasCheck = 1 and J.YearCompressType<>1  and J.ProdID = A.ProdID  and J.WareID <> '' and J.NeedUpdate = 1  and  (J.BillDate Between @DATE1 And @DATE2) ) AS INT) As 銷貨退回");
sb.Append(" ,CAST((Select +IsNull(Sum(S.CostForAcc),0) From comProdRec  S Where S.Flag not in (318,319,320) and S.Flag Between 300 and 399 and S.HasCheck = 1 and S.YearCompressType<>1  and S.ProdID = A.ProdID  and S.WareID <> '' and S.NeedUpdate = 1  and  (S.BillDate Between @DATE1 And @DATE2) ) AS INT) As 本期調整");
sb.Append(" ,CAST(((Select +IsNull(Sum(K.CostForAcc),0) From comProdRec  K Where K.Flag not in (318,319,320) and K.Flag Between 400 and 499 and K.HasCheck = 1 and K.YearCompressType<>1  and K.ProdID = A.ProdID  and K.WareID <> '' and K.NeedUpdate = 1  and  (K.BillDate Between @DATE1 And @DATE2) ) + (Select -IsNull(Sum(L.CostForAcc),0) From comProdRec  L Where L.Flag not in (318,319,320) and L.Flag Between 400 and 499 and L.HasCheck = 1 and L.YearCompressType<>1  and L.ProdID = A.ProdID  and L.WareID <> '' and L.NeedUpdate = 1  and  (L.BillDate Between @DATE1 And @DATE2) )) AS INT) As 本期調撥 ");
sb.Append(" From ComProduct A ");
sb.Append(" Left Join CHICOMP02.DBO.comProduct J On A.ProdID =J.ProdID            ");
sb.Append(" Left Join CHICOMP02.DBO.comProductClass K On J.ClassID =K.ClassID            ");
sb.Append("  Where ((A.ProdForm<6) or (A.ProdForm=8)) ");
sb.Append(" AND A.ClassID NOT IN ('B','C','D','D0002','D3','DVD','F','FA','FEED','H','K','L','M','MA','O','OA','OM','P')");

            if (checkBox2.Checked)
            {
                sb.Append(" and   A.ProdID in ( " + d + ") ");
            }
            else
            {
                if (textBox9.Text != "" && textBox10.Text != "")
                {

                    sb.Append("  AND A.ProdID  between @ProdID1 and @ProdID2 ");
                }
            }


            if (checkBox5.Checked)
            {
                sb.Append(" and    K.ClassName in ( " + MM + ") ");
            }
            else
            {

                if (comboBox3.Text != "")
                {

                    sb.Append("  AND K.ClassName  = @CClassName ");
                }
            }
            sb.Append("  Order By A.ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DATE1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@DATE2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@ProdID1", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@ProdID2", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@CClassName", comboBox3.Text));
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

        private void ACCFIRSTEND_Load(object sender, EventArgs e)
        {
            comboBox1.Text = "聿豐";
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.Day();
  
        }
        public string MM;
        private void button10_Click(object sender, EventArgs e)
        {
            APS5CHOICE frm1 = new APS5CHOICE();

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox5.Checked = true;
                MM = frm1.MM;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1); 
        }

    }
}

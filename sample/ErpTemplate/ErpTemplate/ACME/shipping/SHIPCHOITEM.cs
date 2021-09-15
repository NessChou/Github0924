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
    public partial class SHIPCHOITEM : Form
    {
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCHO = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public SHIPCHOITEM()
        {
            InitializeComponent();
        }

        private void SHIPCHOITEM_Load(object sender, EventArgs e)
        {

            bb();
        }

        private void bb()
        {
            try
            {

                string FileName = string.Empty;
                System.Data.DataTable dt1s = GETCHO();
                System.Data.DataTable dtCost = MakeTable();
                DataRow dr = null;
                if (dt1s.Rows.Count > 0)
                {
                    for (int ig = 0; ig <= dt1s.Rows.Count - 1; ig++)
                    {
                        DataRow drws = dt1s.Rows[ig];
                        string 公司 = drws["公司"].ToString();
                        string 產品編號 = drws["產品編號"].ToString();
                        string 品名規格 = drws["品名規格"].ToString();
                        string 發票品名 = drws["發票品名"].ToString();
                        string 船務品名 = drws["船務品名"].ToString();
                        System.Data.DataTable dt2 = GETSAP(產品編號);
                        if (dt2.Rows.Count == 0)
                        {
                            dr = dtCost.NewRow();
                            dr["公司"] = 公司;
                            dr["產品編號"] = 產品編號;
                            dr["品名規格"] = 品名規格;
                            dr["發票品名"] = 發票品名;
                            dr["船務品名"] = 船務品名;
                            dtCost.Rows.Add(dr);
                        }
                    }
                    if (dtCost.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dtCost;

                    }
                }



            }
            catch (Exception ex)
            {


            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            string CONN = "";
            for (int i = dataGridView1.Rows.Count - 1; i >= 0; i--)
            {

                row = dataGridView1.Rows[i];
                string 公司 = row.Cells["公司"].Value.ToString();
                string 產品編號 = row.Cells["產品編號"].Value.ToString();
                string 船務品名 = row.Cells["船務品名"].Value.ToString();
                
                if (公司 == "TOP GARDEN")
                {
                     CONN = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
                if (公司 == "CHOICE")
                {
                    CONN = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
                if (公司 == "INFINITE")
                {
                    CONN = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
                }
                if (!String.IsNullOrEmpty(船務品名))
                {
                    UPDATEENGNAME(CONN, 船務品名, 產品編號);
                }
            }
            MessageBox.Show("更新成功");
            bb();
        }

        public System.Data.DataTable GETCHO()
        {

            SqlConnection MyConnection = new SqlConnection(strCHO);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'TOP GARDEN' 公司,PRODID 產品編號,ProdName 品名規格,InvoProdName  發票品名,ENGNAME  船務品名 FROM CHIComp20.DBO.comProduct  WHERE  (LatestPurchDate>=20171018 OR LatestPurchDate=0) AND ISNULL(ENGNAME,'') =''");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'CHOICE' 公司,PRODID 產品編號,ProdName 品名規格,InvoProdName  發票品名,ENGNAME  船務品名 FROM CHIComp21.DBO.comProduct     WHERE (LatestPurchDate>=20171018 OR LatestPurchDate=0)  AND ISNULL(ENGNAME,'') ='' ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT 'INFINITE' 公司,PRODID 產品編號,ProdName 品名規格,InvoProdName  發票品名,ENGNAME  船務品名 FROM CHIComp22.DBO.comProduct   WHERE  (LatestPurchDate>=20171018 OR LatestPurchDate=0)  AND ISNULL(ENGNAME,'') =''");
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
        public System.Data.DataTable GETSAP(string ITEMCODE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn02);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE  FROM OITM WHERE ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public void UPDATEENGNAME(string CONN, string ENGNAME, string ProdID)
        {
            SqlConnection connection = new SqlConnection(CONN);
            SqlCommand command = new SqlCommand("UPDATE DBO.comProduct  SET ENGNAME=@ENGNAME WHERE ProdID=@ProdID ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CONN", CONN));
            command.Parameters.Add(new SqlParameter("@ENGNAME", ENGNAME));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

            //USERS
            //DDATE
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

        private System.Data.DataTable MakeTable()
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("發票品名", typeof(string));
            dt.Columns.Add("船務品名", typeof(string));

            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }



    }
}

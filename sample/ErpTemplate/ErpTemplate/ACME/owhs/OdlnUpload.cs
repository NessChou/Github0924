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
    public partial class OdlnUpload : Form
    {
        public OdlnUpload()
        {
            InitializeComponent();
        }

        private void OdlnUpload_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt = PackOP1();
            bindingSource1.DataSource = dt;
            dataTable1DataGridView.DataSource = bindingSource1.DataSource;
            dataTable1DataGridView.Columns[7].Visible = false;
            dataTable1DataGridView.Columns[8].Visible = false;
            dataTable1DataGridView.Columns[9].Visible = false;
            dataTable1DataGridView.Columns[0].Width = 70;
            dataTable1DataGridView.Columns[1].Width = 70;
            dataTable1DataGridView.Columns[2].Width = 200;
            dataTable1DataGridView.Columns[3].Width = 150;
            dataTable1DataGridView.Columns[4].Width = 150;
            dataTable1DataGridView.Columns[5].Width = 70;
            dataTable1DataGridView.Columns[6].Width = 100;
            dataTable1DataGridView.Columns[10].Width = 100;
            dataTable1DataGridView.Columns[0].HeaderText = "出貨日期";
            dataTable1DataGridView.Columns[1].HeaderText = "銷貨單號";
            dataTable1DataGridView.Columns[2].HeaderText = "客戶名稱";
            dataTable1DataGridView.Columns[3].HeaderText = "產品編號";
            dataTable1DataGridView.Columns[4].HeaderText = "品名";
            dataTable1DataGridView.Columns[5].HeaderText = "數量";
            dataTable1DataGridView.Columns[6].HeaderText = "AUO INVOICE";
            dataTable1DataGridView.Columns[10].HeaderText = "業務";
            DataGridViewLinkColumn column = new DataGridViewLinkColumn();
            column.Name = "Link";
            column.UseColumnTextForLinkValue = true;

                    column.Text = "讀取檔案";
         

    

            //マウスポインタがリンク上にあるときだけ下線をつける

            column.LinkBehavior = LinkBehavior.HoverUnderline;

            //自動的に訪問済みになるようにする

            //デフォルトでTrue

            column.TrackVisitedState = true;

            //DataGridViewに追加する

            dataTable1DataGridView.Columns.Add(column);
            //selectpic();

        }

        private void dataTable1DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;
                if (dgv.Columns[e.ColumnIndex].Name == "Link")
                {
                    System.Data.DataTable dt1 = PackOP1();
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];

                    System.Diagnostics.Process.Start(drw["path"].ToString() +"\\"+ drw["路徑"].ToString());

                    //訪問済みにする

                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;


                  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataTable1DataGridView_RowDefaultCellStyleChanged(object sender, DataGridViewRowEventArgs e)
        {
            //System.Data.DataTable dt1 = ship.DataTable1;
            //int i = e.RowIndex;
            //DataRow drw = dt1.Rows[i];

            //drw["path"].ToString() = "";
           

        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = PackOP();
            bindingSource1.DataSource = dt;
            dataTable1DataGridView.DataSource = bindingSource1.DataSource;
        }

        public System.Data.DataTable PackOP()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select Convert(varchar(8),T0.DOCDATE,112) DOCDATE,T0.DOCNUM,T0.CARDNAME,T1.ITEMCODE,T1.DSCRIPTION,CAST(T1.QUANTITY AS INT) 數量,t1.u_acme_inv INV ");
            sb.Append("            ,[Filename],t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS VARCHAR)+'.'+Fileext 路徑,T4.[SLPNAME] SLPNAME from ACMESQL02.DBO.ODLN T0");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.OCLG T2 ON (T0.DOCENTRY=T2.DOCNUM)");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append("            LEFT JOIN acmesql02.dbo.OSLP T4 ON (T0.SLPCODE = T4.SLPCODE) ");
            sb.Append("            where Fileext <>'' and substring(T3.[FILENAME],0,2)='序' ");

      
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and   Convert(varchar(8),T0.[docDATE],112) between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox4.Text != "" )
            {

                sb.Append(" and  T0.CARDCODE in  (" + textBox4.Text.ToString() + ")  ");
            }
         
            sb.Append(" ORDER BY T1.SHIPDATE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }
        public System.Data.DataTable PackOP1()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  select Convert(varchar(8),T0.DOCDATE,112) DOCDATE,T0.DOCNUM,T0.CARDNAME,T1.ITEMCODE,T1.DSCRIPTION,CAST(T1.QUANTITY AS INT) 數量,t1.u_acme_inv INV ");
            sb.Append("            ,[Filename],t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS VARCHAR)+'.'+Fileext 路徑,T4.[SLPNAME] SLPNAME from ACMESQL02.DBO.ODLN T0");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.OCLG T2 ON (T0.DOCENTRY=T2.DOCNUM)");
            sb.Append("            LEFT JOIN ACMESQL02.DBO.ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append("            LEFT JOIN acmesql02.dbo.OSLP T4 ON (T0.SLPCODE = T4.SLPCODE) ");
            sb.Append("            where Fileext <>'' and substring(T3.[FILENAME],0,2)='序' ");
            sb.Append(" ORDER BY T1.SHIPDATE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }
        private void button3_Click(object sender, EventArgs e)
        {
            APS1 frm1 = new APS1();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

                textBox4.Text = frm1.q;

            }
        }

    }
}
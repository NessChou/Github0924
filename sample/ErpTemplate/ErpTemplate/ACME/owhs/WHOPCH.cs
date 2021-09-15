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
    public partial class WHOPCH : Form
    {
        StringBuilder sbS = new StringBuilder();
        StringBuilder sbS2 = new StringBuilder();
        public WHOPCH()
        {
            InitializeComponent();
        }
        public void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }
        public System.Data.DataTable GetOPDN()
        {
            Clear(sbS);
            Clear(sbS2);
            SBS();
            SBS2();
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct t4.docentry 收貨採購單號,T4.U_SHIPPING_NO　SHIPPING工單號碼,T4.U_ACME_INV  進項發票號碼, convert(varchar,U_ACME_INVOICE, 111)　發票日期, ");
            sb.Append(" T5.ADD9 報關號碼,cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱,T6.[PATH] PATH2 ,T6.[filename] 檔案名稱2   from oclg t2    ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)    ");
            sb.Append(" inner join opdn t4 on(cast(t2.docentry as varchar)=cast(t4.docentry as varchar))        ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SHIPPING_MAIN T5 ON (T4.U_Shipping_no =T5.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.DOWNLOAD T6 ON (T5.ShippingCode =T6.shippingcode  ");
            sb.Append(" AND  (REPLACE(T6.[filename],'已塗改.PDF','')= t4.U_ACME_INV COLLATE  Chinese_Taiwan_Stroke_CI_AS OR REPLACE(T6.[filename],'已塗改.PDF','')=     T5.ADD9 COLLATE  Chinese_Taiwan_Stroke_CI_AS) ) ");
            sb.Append(" where  T2.DOCTYPE='20' AND isnull(t4.AtcEntry,'')= '' ");

            if (sbS.ToString() != "''")
            {
                sb.Append(" and   ( t4.U_ACME_INV  IN (" + sbS.ToString() + ") OR t4.U_ACME_INV LIKE '%" + sbS.ToString().Replace("'", "") + "%')   ");
            }
            if (sbS2.ToString() != "''")
            {
                sb.Append(" and  T4.U_SHIPPING_NO IN (" + sbS2.ToString() + ")  ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" select distinct t4.docentry 收貨採購單號,T4.U_SHIPPING_NO　SHIPPING工單號碼,T4.U_ACME_INV  進項發票號碼, convert(varchar,U_ACME_INVOICE, 111)　發票日期,  ");
            sb.Append(" T5.ADD9 報關號碼,cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱,T6.[PATH] PATH2 ,T6.[filename] 檔案名稱2   from ATC1 T3      ");
            sb.Append(" inner join opdn t4 on(T3.ABSENTRY=t4.AtcEntry)     ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SHIPPING_MAIN T5 ON (T4.U_Shipping_no =T5.ShippingCode COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.DOWNLOAD T6 ON (T5.ShippingCode =T6.shippingcode   ");
            sb.Append(" AND  (REPLACE(T6.[filename],'已塗改.PDF','')= t4.U_ACME_INV COLLATE  Chinese_Taiwan_Stroke_CI_AS OR REPLACE(T6.[filename],'已塗改.PDF','')=     T5.ADD9 COLLATE  Chinese_Taiwan_Stroke_CI_AS) )  ");
            sb.Append(" where  isnull(t4.AtcEntry,'')<> ''");

            if (sbS.ToString() != "''")
            {
                sb.Append(" and   ( t4.U_ACME_INV  IN (" + sbS.ToString() + ") OR t4.U_ACME_INV LIKE '%" + sbS.ToString().Replace("'","") + "%')   ");
            }
            if (sbS2.ToString() != "''")
            {
                sb.Append(" and  T4.U_SHIPPING_NO IN (" + sbS2.ToString() + ")  ");
            }
            if (sbS2.ToString() != "''")
            {
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '','','','',T5.ADD9 ,'','','',T6.[PATH] PATH2 ,T6.[filename] 檔案名稱2 FROM  ACMESQLSP.DBO.SHIPPING_MAIN T5");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.DOWNLOAD2 T6 ON (T5.ShippingCode =T6.shippingcode   ");
            sb.Append(" AND  MARK='1')");
       
                sb.Append(" WHERE   T5.SHIPPINGCODE IN (" + sbS2.ToString() + ")  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;
                //SHIPPING工單號碼
                //     string MARK = download2DataGridView.CurrentRow.Cells["MARK2"].Value.ToString();
                if (dgv.Columns[e.ColumnIndex].Name == "檔案名稱")
                {



                        string path = dataGridView1.CurrentRow.Cells["path"].Value.ToString();
                        string 路徑 = dataGridView1.CurrentRow.Cells["路徑"].Value.ToString();
                        string 檔案名稱 = dataGridView1.CurrentRow.Cells["檔案名稱"].Value.ToString();
                    
                        string aa = path + "\\" + 路徑;

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    

                }

                if (dgv.Columns[e.ColumnIndex].Name == "報關檔案")
                {




                        string PATH2 = dataGridView1.CurrentRow.Cells["PATH2"].Value.ToString();

                        string 報關檔案 = dataGridView1.CurrentRow.Cells["報關檔案"].Value.ToString();

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = 報關檔案;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(PATH2, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    

                }
                if (dgv.Columns[e.ColumnIndex].Name == "SHIPPING工單號碼")
                {
                    string MARK = dataGridView1.CurrentRow.Cells["SHIPPING工單號碼"].Value.ToString();


                    int T1 = MARK.IndexOf("SH");
                    int T2 = MARK.IndexOf("SI");

                    if (T1 != -1)
                    {
                        fmShip a = new fmShip();
                        a.PublicString = MARK;
                        a.Show();
                    }

                    if (T2 != -1)
                    {
                        APShip a = new APShip();
                        a.PublicString = MARK;
                        a.Show();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox4.Text == "")
            {
                MessageBox.Show("請輸入查詢條件");
                return;
            
            }

            System.Data.DataTable G1 = GetOPDN();
            if (G1.Rows.Count > 0)
            {
                dataGridView1.DataSource = G1;
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }
        private void SBS()
        {
            string[] arrurl = textBox1.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS.Append("'" + i + "',");
            }
            sbS.Remove(sbS.Length - 1, 1);
        }
        private void SBS2()
        {
            string[] arrurl = textBox4.Text.Split(new Char[] { ',' });

            foreach (string i in arrurl)
            {
                sbS2.Append("'" + i + "',");
            }
            sbS2.Remove(sbS2.Length - 1, 1);
        }
    }
}

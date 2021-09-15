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
    public partial class SHIPBOMDAN : Form
    {
        string SHIPNO = "";
        public SHIPBOMDAN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DELPACK();
                StringBuilder sb = new StringBuilder();


                for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
                {

                    DataGridViewRow row;

                    row = dataGridView1.Rows[i];
                    if (row.Cells["型號"].Value != null && row.Cells["數量"].Value != null)
                    {

                        string 型號 = row.Cells["型號"].Value.ToString();
                        int 數量 = Convert.ToInt32(row.Cells["數量"].Value);

                        int I1 = 型號.ToUpper().IndexOf("V.");
                        int I2 = 型號.IndexOf(".");
                        string MODEL = "";
                        string VER = "";
                        if (I2 == -1)
                        {
                            MessageBox.Show(型號 + " 請輸入正確格式");
                            return;
                        }

                        if (I1 == -1)
                        {
                            MODEL = 型號.Substring(0, I2).Trim();
                            VER = 型號.Substring(I2 + 1, 1).Trim();
                        }
                        else
                        {
                            MODEL = 型號.Substring(0, I1).Trim();
                            VER = 型號.Substring(I2 + 1, 1).Trim();
                        }
                        System.Data.DataTable G1 = GetF(MODEL, VER, 數量);
                        if (G1.Rows.Count > 0)
                        {
                            for (int i2 = 0; i2 <= G1.Rows.Count - 1; i2++)
                            {
                                string SHIPPINGCODE=G1.Rows[i2]["JOBNO"].ToString();
                                InsertBOM(SHIPPINGCODE, 型號);
                                sb.Append("'" + SHIPPINGCODE + "',");
                            }

                        }
                    }
                }
                if (sb.Length > 0)
                {
                    sb.Remove(sb.Length - 1, 1);

                    SHIPNO = sb.ToString();
               
                    DataRow dr2 = null;
                    if (!String.IsNullOrEmpty(SHIPNO))
                    {
                        System.Data.DataTable dtCost2 = MakeTable2();
                        System.Data.DataTable K1 = GetOPDN(SHIPNO);

                        if (K1.Rows.Count > 0)
                        {
                            for (int i = 0; i <= K1.Rows.Count - 1; i++)
                            {
                                dr2 = dtCost2.NewRow();
                                string SHIPPINGNO = K1.Rows[i]["JOBNO"].ToString();
                                string BOMDAN = K1.Rows[i]["BOMDAN"].ToString();
                                string PO = K1.Rows[i]["收貨採購單號"].ToString();
                                string INVOICE = K1.Rows[i]["進項發票號碼"].ToString();
                                dr2["JOBNO"] = SHIPPINGNO;
                                dr2["收貨採購單號"] = PO;
                                dr2["進項發票號碼"] = INVOICE;
                                dr2["路徑"] = K1.Rows[i]["路徑"].ToString();
                                dr2["檔案名稱"] = K1.Rows[i]["檔案名稱"].ToString();
                                System.Data.DataTable K2F = GetOPDN2(SHIPPINGNO, BOMDAN);
                                if (K2F.Rows.Count > 0)
                                {
                                    dr2["報單號碼"] = K2F.Rows[0]["報單號碼"].ToString();
                                    dr2["path2"] = K2F.Rows[0]["path"].ToString();
                                    dr2["報單下載"] = K2F.Rows[0]["檔案名稱"].ToString();
                                }
                                StringBuilder sb2 = new StringBuilder();
                                System.Data.DataTable dt = GetOPDN3(SHIPPINGNO);
                                if (dt.Rows.Count > 0)
                                {
                                    for (int i2 = 0; i2 <= dt.Rows.Count - 1; i2++)
                                    {
                                        DataRow dd = dt.Rows[i2];
                                        sb2.Append(dd["MODEL"].ToString() + "/");
                                    }

                                    sb2.Remove(sb2.Length - 1, 1);
                                    dr2["型號"] = sb2.ToString();
                                }

                           System.Data.DataTable d2 = GetOPDN4(SHIPPINGNO,PO,INVOICE);
                           if (d2.Rows.Count > 0)
                           {
                               dr2["PRINT"] = d2.Rows[0]["BCHECK"].ToString();
                               dr2["PDATE"] = d2.Rows[0]["BDATE"].ToString();
                           }
                                dtCost2.Rows.Add(dr2);
                            }
                            dataGridView3.DataSource = dtCost2;
                        }
                     
  
               
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
      
        }


        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("型號", typeof(string));
            dt.Columns.Add("收貨採購單號", typeof(string));
            dt.Columns.Add("進項發票號碼", typeof(string));
            dt.Columns.Add("路徑", typeof(string));
            dt.Columns.Add("檔案名稱", typeof(string));
            dt.Columns.Add("報單號碼", typeof(string));
            dt.Columns.Add("報單下載", typeof(string));
            dt.Columns.Add("path2", typeof(string));
            dt.Columns.Add("PRINT", typeof(string));
            dt.Columns.Add("PDATE", typeof(string));   
            return dt;
        }
        public static System.Data.DataTable GetOPDN(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select distinct  t4.u_shipping_no JOBNO,t4.docentry 收貨採購單號,T10.U_PC_BSINV 進項發票號碼,cast(t3.TRGTPATH as nvarchar(80))+'\\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,FILENAME+'.'+Fileext 檔案名稱,T11.add9 BOMDAN  from oclg t2 ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY) ");
            sb.Append(" inner join opdn t4 on(t2.docentry=t4.docentry) ");
            sb.Append(" left join PDN1 t5 on (t4.docentry=T5.docentry )");
            sb.Append(" left join PCH1 t12 on (t12.baseentry=T5.docentry and  t12.baseline=t5.linenum and t12.basetype='20'  )");
            sb.Append(" left join OPCH t10 on (T12.DOCENTRY=T10.DOCENTRY )");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SHIPPING_MAIN T11 ON (T4.u_shipping_no=T11.SHIPPINGCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where  t2.doctype='20' and isnull(t3.[FILENAME],'') <> '' and t4.u_shipping_no IN (" + SHIPPINGCODE + "  )");

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

        public static System.Data.DataTable GetOPDN2(string SHIPPINGCODE, string BOMGUN)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.SHIPPINGCODE JOBNO,ADD9 報單號碼,[PATH] ,T0.[filename] 檔案名稱 FROM Download2  T0 LEFT JOIN SHIPPING_MAIN T1 ON (T0.shippingcode =T1.ShippingCode)");
            sb.Append("  WHERE ISNULL(ADD9,'') <> ''  and T0.SHIPPINGCODE  =@SHIPPINGCODE AND T0.[filename]  like '%" + BOMGUN + "%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BOMGUN", BOMGUN));
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
        public static System.Data.DataTable GetOPDN3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MODEL FROM Shipping_BOMDAN WHERE SHIPPINGCODE =@SHIPPINGCODE AND USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public static System.Data.DataTable GetOPDN4(string SHIPPINGCODE,string PO, string INVOICE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BCHECK,BDATE  FROM Shipping_BOMDAN2 where SHIPPINGCODE=@SHIPPINGCODE  AND PO=@PO AND INVOICE=@INVOICE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
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
        private System.Data.DataTable GetF(string MODEL, string VER, int QTY)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("          SELECT  T0.SHIPPINGCODE JOBNO FROM Shipping_Main T0 ");
            sb.Append(" LEFT JOIN Shipping_Item T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                WHERE SUBSTRING(T0.CARDCODE,1,1) in ('S','U') AND BOARDCOUNTNO='進口' ");
            sb.Append("                AND ISNULL(ADD9,'') <> '' AND T2.U_TMODEL like '%" + MODEL + "%' AND T2.U_VERSION=@VER   ");
            sb.Append("  AND SUBSTRING(T0.SHIPPINGCODE,3,4)  BETWEEN @YEAR1 AND @YEAR2 ");
            sb.Append("    GROUP BY T0.SHIPPINGCODE HAVING  SUM(T1.Quantity) >= @QTY    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@YEAR1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@YEAR2", textBox2.Text));
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

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                if (dgv.Columns[e.ColumnIndex].Name == "檔案名稱")
                {



                    string 路徑 = dataGridView3.CurrentRow.Cells["路徑"].Value.ToString();
                    string 檔案名稱 = dataGridView3.CurrentRow.Cells["檔案名稱"].Value.ToString();

                    //System.Data.DataTable dt1 = GetOPDNF(JOBNOS, 檔案名稱);

                    //if (dt1.Rows.Count > 0)
                    //{
                    //    DataRow drw = dt1.Rows[0];


                    string aa = 路徑 ;


                    string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);



                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }
              //  }

                if (dgv.Columns[e.ColumnIndex].Name == "Column3")
                {


                    string kk = dataGridView3.CurrentRow.Cells["Column3"].Value.ToString();

                    string aa = dataGridView3.CurrentRow.Cells["path2"].Value.ToString();

                    string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + kk;

                    System.IO.File.Copy(aa, NewFileName, true);

                    System.Diagnostics.Process.Start(NewFileName);




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

        private void InsertBOM(string SHIPPINGCODE, string MODEL)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO Shipping_BOMDAN (SHIPPINGCODE,MODEL,USERS) VALUES(@SHIPPINGCODE,@MODEL,@USERS)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));


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


        private void InsertBOM2(string SHIPPINGCODE,string PO,string INVOICE,string BCHECK,string BDATE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO Shipping_BOMDAN2 (SHIPPINGCODE,PO,INVOICE,BCHECK,BDATE) VALUES(@SHIPPINGCODE,@PO,@INVOICE,@BCHECK,@BDATE)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@BCHECK", BCHECK));
            command.Parameters.Add(new SqlParameter("@BDATE", BDATE));
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
        public void DELPACK()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" delete Shipping_BOMDAN where users=@USERS ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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
        public void DELPACK2(string SHIPPINGCODE, string PO, string INVOICE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" DELETE Shipping_BOMDAN2 where SHIPPINGCODE=@SHIPPINGCODE AND PO=@PO AND INVOICE=@INVOICE ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PO", PO));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
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

        private void SHIPBOMDAN_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyy");
            textBox2.Text = DateTime.Now.ToString("yyyy");
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView3_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string COL = dataGridView3.Columns[e.ColumnIndex].Name.ToString();
            if (dataGridView3.Columns[e.ColumnIndex].Name.ToString() == "JOBNO") 
            {
                string ShippingCode = dataGridView3.Rows[e.RowIndex].Cells["JOBNO"].Value.ToString();
                APShip a = new APShip();
                a.PublicString = ShippingCode;
                a.Show();
            }
            
        }
    }
}

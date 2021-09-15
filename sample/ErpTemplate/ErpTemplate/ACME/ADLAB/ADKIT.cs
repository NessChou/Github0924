using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
using System.Globalization;
namespace ACME
{
    public partial class ADKIT : Form
    {
        string strCnSP = "Data Source=acmesap;Initial Catalog=acmesqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public ADKIT()
        {
            InitializeComponent();
        }


    

        private void ADKIT_Load(object sender, EventArgs e)
        {

            //\\acmesrv01\Public\ADLab 資料\資材部\客製專案規格書
            string AD = "//acmesrv01//Public//ADLab 資料//資材部//專案規格管理";
            DELAD_ASI();
                  DD2(AD);
                  FF();
        }


        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {

                string da = dataGridView1.SelectedRows[0].Cells["產品編號1"].Value.ToString();

                ADKIT2 a = new ADKIT2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }
        public System.Data.DataTable GetCHOITEM(string MTYPE)
        {
            SqlConnection MyConnection = new SqlConnection(strCn16);
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT T0.CLASSID 產品類別,T1.ClassName 類別名稱,T0.ProdID 產品編號,PRODNAME 品名規格,T0.InvoProdName 發票品名");
            sb.Append(" ,T0.ProdDesc 附註   FROM comProduct  T0 LEFT JOIN comProductClass T1  ON (T0.ClassID =T1.ClassID)");
            if (MTYPE == "Complete Display")
            {
                sb.Append(" WHERE T0.CLASSID IN ('1001','1002','1003','1004','1005','1006','1007','1008','2201','1010')");
            }
            if (MTYPE == "Panel")
            {
                sb.Append(" WHERE T0.CLASSID IN ('1201','1202','1203','1204','1211','1281','1301','1305')");
            }
            if (MTYPE == "Kit Parts")
            {
                sb.Append(" WHERE T0.CLASSID IN ('1507','1101','1102','1103','1306','1401','1402','1403','1405','1406','1407','1408','1501','1502','1503','1504','1505','1506','1601','1603','1701','1702','1801','1802','1803','1804','1805','1806','1807','1901','1902','1912','1913','1914','2001','2003','2004','1921','1915','1703')");
            }
            if (MTYPE == "Fee")
            {
                sb.Append(" WHERE T0.CLASSID IN ('A001','C001','C002','C003','D001','F001','F002','H001','I001','N001','O001','O101','R001','S001','S101','S102','T001','W0','W001')");
            }
            if (MTYPE == "SOLAR")
            {
                //2101  / 2102 / 2103 / 21S0
                sb.Append(" WHERE T0.CLASSID IN ('2101','2102','2103','21S0','2121')");
            }
            if (textBox1.Text != "")
            {
                sb.Append(" and  T0.ProdID LIKE '%" + textBox1.Text + "%'   ");
            }
            if (textBox2.Text != "")
            {
                sb.Append(" and  T0.PRODNAME LIKE '%" + textBox2.Text + "%'   ");
            }

            if (textBox3.Text != "")
            {
                sb.Append(" and  T0.InvoProdName LIKE '%" + textBox3.Text + "%'   ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MTYPE", MTYPE));
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
        public System.Data.DataTable GetASI(string MODEL)
        {
            SqlConnection MyConnection = new SqlConnection(strCnSP);
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT * FROM AD_ASI  WHERE MODEL=@MODEL");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
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

        private void button2_Click(object sender, EventArgs e)
        {

            FF();
        }

        private System.Data.DataTable MakeTable()
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("類別名稱", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("發票品名", typeof(string));
            dt.Columns.Add("附註", typeof(string));
            dt.Columns.Add("檔案下載", typeof(string));

            return dt;
        }
        public void DD2(string PATH)
        {


            string[] filebBrand = Directory.GetDirectories(PATH);
            foreach (string fileabBrand in filebBrand)
            {
                DirectoryInfo DIRINFO = new DirectoryInfo(fileabBrand);

                string DIRNAME = DIRINFO.Name.ToString();



                string[] filecSize = Directory.GetFiles(fileabBrand);
                                foreach (string fie in filecSize)
                                {
                                    int aa = fie.LastIndexOf(".");
                                    string Type;
                                    Type = fileabBrand.Replace(PATH, "");
                                    FileInfo filess = new FileInfo(fie);
                                    string dd = filess.Name.ToString();

                                    int ad = dd.LastIndexOf(".");

                                    string size = filess.Length.ToString();
                                    string FileDate = filess.CreationTime.ToString("yyyyMMdd");
                                                string PanelName = dd.Substring(0, ad).ToString();
                                                int D1 = PanelName.IndexOf("_");
                                                 string MODEL=PanelName;
                                               if (D1 != -1)
                                                {
                                                    MODEL = PanelName.Substring(0, D1);
                                                }
                                               string PP = fie.ToString().Replace(@"//acmesrv01//Public//ADLab 資料//資材部//專案規格管理", @"\\acmesrv01\Public\ADLab 資料\資材部\專案規格管理");
                                               ADDAD_ASI(DIRNAME, PanelName, PP, size, FileDate);
                                            
                               

                                }

                            
                        
                    
                
            }


        }
        public void DELAD_ASI()
        {
            SqlConnection connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE AD_ASI", connection);
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
        public void ADDAD_ASI(string MODEL, string PanelName, string Path, string FileSize, string FileDate)
        {
            SqlConnection connection = new SqlConnection(strCnSP);
            SqlCommand command = new SqlCommand("Insert into AD_ASI(MODEL,PanelName,Path,FileSize,FileDate) values(@MODEL,@PanelName,@Path,@FileSize,@FileDate)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@PanelName", PanelName));
            command.Parameters.Add(new SqlParameter("@Path", Path));
            command.Parameters.Add(new SqlParameter("@FileSize", FileSize));
            command.Parameters.Add(new SqlParameter("@FileDate", FileDate));
      
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
        private void FF()
        {

            //dt.Columns.Add("檔案下載", typeof(string));
            FF2("Complete Display", dataGridView1);
            FF2("Panel", dataGridView2);
            FF2("Kit Parts", dataGridView3);
            FF2("Fee", dataGridView4);
            FF2("SOLAR", dataGridView5);
        }
        private void FF2(string DOCTYPE,DataGridView GD)
        {

            System.Data.DataTable dt = GetCHOITEM(DOCTYPE);
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTable();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = dtCost.NewRow();
                dr["產品類別"] = dt.Rows[i]["產品類別"].ToString();
                dr["類別名稱"] = dt.Rows[i]["類別名稱"].ToString();
                dr["產品編號"] = dt.Rows[i]["產品編號"].ToString();
                string 產品編號 = dt.Rows[i]["產品編號"].ToString();
                string 品名規格 = dt.Rows[i]["品名規格"].ToString();
                string 發票品名 = dt.Rows[i]["發票品名"].ToString();
                dr["品名規格"] = 品名規格;
                dr["發票品名"] = dt.Rows[i]["發票品名"].ToString();
                dr["附註"] = dt.Rows[i]["附註"].ToString();
                int D1 = 發票品名.LastIndexOf(" ");
                     
                if (D1 != -1)
                {

                    System.Data.DataTable DD1 = GetASI(產品編號);
                    if (DD1.Rows.Count > 0)
                    {
                        dr["檔案下載"] = "下載";
                    }
                }
                dtCost.Rows.Add(dr);
            }
            GD.DataSource = dtCost;
        }
        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {

                string da = dataGridView2.SelectedRows[0].Cells["產品編號2"].Value.ToString();

                ADKIT2 a = new ADKIT2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {

                string da = dataGridView3.SelectedRows[0].Cells["產品編號3"].Value.ToString();

                ADKIT2 a = new ADKIT2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void dataGridView4_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView4.SelectedRows.Count > 0)
            {

                string da = dataGridView4.SelectedRows[0].Cells["產品編號4"].Value.ToString();

                ADKIT2 a = new ADKIT2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案下載")
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        string 產品編號 = dataGridView1.CurrentRow.Cells["產品編號1"].Value.ToString();



                        System.Data.DataTable DD1 = GetASI(產品編號);
                            if (DD1.Rows.Count > 0)
                            {
                                for (int i = 0; i <= DD1.Rows.Count - 1; i++)
                                {

                                    string aa = DD1.Rows[i]["path"].ToString();


                                    System.Diagnostics.Process.Start(aa);

                                    DataGridViewLinkCell cell =

                                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                                    cell.LinkVisited = true;
                                }

                            
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案下載2")
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        string 產品編號 = dataGridView2.CurrentRow.Cells["產品編號2"].Value.ToString();

                        System.Data.DataTable DD1 = GetASI(產品編號);
                            if (DD1.Rows.Count > 0)
                            {
                                for (int i = 0; i <= DD1.Rows.Count - 1; i++)
                                {

                                    string aa = DD1.Rows[i]["path"].ToString();


                                    System.Diagnostics.Process.Start(aa);

                                    DataGridViewLinkCell cell =

                                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                                    cell.LinkVisited = true;
                                }

                            }
                        
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案下載3")
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        string 產品編號 = dataGridView3.CurrentRow.Cells["產品編號3"].Value.ToString();


                        System.Data.DataTable DD1 = GetASI(產品編號);
                            if (DD1.Rows.Count > 0)
                            {
                                for (int i = 0; i <= DD1.Rows.Count - 1; i++)
                                {

                                    string aa = DD1.Rows[i]["path"].ToString();


                                    System.Diagnostics.Process.Start(aa);

                                    DataGridViewLinkCell cell =

                                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                                    cell.LinkVisited = true;
                                }

                            
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案下載4")
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        string 產品編號 = dataGridView4.CurrentRow.Cells["產品編號4"].Value.ToString();


                        System.Data.DataTable DD1 = GetASI(產品編號);
                            if (DD1.Rows.Count > 0)
                            {
                                for (int i = 0; i <= DD1.Rows.Count - 1; i++)
                                {

                                    string aa = DD1.Rows[i]["path"].ToString();


                                    System.Diagnostics.Process.Start(aa);

                                    DataGridViewLinkCell cell =

                                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                                    cell.LinkVisited = true;
                                }

                            
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView5_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView5.SelectedRows.Count > 0)
            {

                string da = dataGridView5.SelectedRows[0].Cells["產品編號5"].Value.ToString();

                ADKIT2 a = new ADKIT2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "檔案下載5")
                {
                    for (int j = 0; j <= 1; j++)
                    {
                        string 產品編號 = dataGridView5.CurrentRow.Cells["產品編號5"].Value.ToString();


                        System.Data.DataTable DD1 = GetASI(產品編號);
                        if (DD1.Rows.Count > 0)
                        {
                            for (int i = 0; i <= DD1.Rows.Count - 1; i++)
                            {

                                string aa = DD1.Rows[i]["path"].ToString();


                                System.Diagnostics.Process.Start(aa);

                                DataGridViewLinkCell cell =

                                    (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                                cell.LinkVisited = true;
                            }


                        }
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

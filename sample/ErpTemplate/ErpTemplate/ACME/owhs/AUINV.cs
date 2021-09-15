using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using System.IO;

namespace ACME
{
    public partial class AUINV : Form
    {
        string A1 = "";
        string A2 = "";
        System.Data.DataTable K1 = new System.Data.DataTable();
        private string ShipConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
        public AUINV()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetINV();
            for (int i = 0; i < dataGridView1.Rows.Count; i++) 
            {
                dataGridView1.Rows[i].Cells["download"].Value = "點擊下載";
            }
            
        }
        public System.Data.DataTable GetINV()
        {
            string CHECK = "";

            if (!checkBox1.Checked)
            {
                CHECK = "True";
            }
            else 
            { 
                CHECK = "False";
            }
            SqlConnection MyConnection;

            MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ShippingCode WH工單,FrgnName SI工單,SeqNo 序號,ItemCode 產品編號,BoxCheck 原廠INVOICE,A5 進貨日期,ShipDate 異常數量,");
            sb.Append(" a2 異常情況,DeCust 客戶別,A3 後續處理情形,A6 運送貨代,A7 SHIPPING備註,PiNo 對應照片號碼,case isnull(a4,'') when '' then 'False' else a4 end 結案,'點擊下載' 下載 FROM WH_Item3 ");
            sb.Append(" where ISNULL(ShipDate,'') <> '' AND (case isnull(a4,'') when '' then 'False' else a4 end)=@A4 AND SUBSTRING(ShippingCode,3,8) between @aa and @bb ");
            //sb.Append(" where SUBSTRING(ShippingCode,3,8) between @aa and @bb ");
            if (tINV.Text != "")
            {
                sb.Append(" AND BoxCheck=@BoxCheck");
            }
            if (cMODEL.SelectedValue.ToString() != "")
            {
                sb.Append(" and  ITEMCODE like '%" + cMODEL.SelectedValue.ToString() + "%'  ");
            }
            if (cVER.SelectedValue.ToString() != "")
            {
                sb.Append(" and  Substring(ITEMCODE,12,1)  = '" + cVER.SelectedValue.ToString() + "'  ");
            }
            if (tWHNO.Text != "")
            {
                sb.Append(" AND ShippingCode=@ShippingCode");
            }
            if (tSHNO.Text != "")
            {
                sb.Append(" AND FrgnName=@FrgnName");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A4", CHECK));
            command.Parameters.Add(new SqlParameter("@aa", tSDATE.Text));
            command.Parameters.Add(new SqlParameter("@bb", tEDATE.Text));
            command.Parameters.Add(new SqlParameter("@BoxCheck", tINV.Text));
            command.Parameters.Add(new SqlParameter("@ShippingCode", tWHNO.Text));
            command.Parameters.Add(new SqlParameter("@FrgnName", tSHNO.Text));
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            button1_Click(sender, e);
        }

        private void AUINV_Load(object sender, EventArgs e)
        {
            tSDATE.Text = GetMenu.DFirst();
            tEDATE.Text = GetMenu.DLast();
           
            UtilSimple.SetLookupBinding(cMODEL, GetMODEL(), "MODEL", "MODEL");
            UtilSimple.SetLookupBinding(cVER, GetVER(), "VER", "VER");


            button1_Click(sender,e);
        }


        private System.Data.DataTable GetMODEL()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("          SELECT distinct CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("                 AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append("          Substring (T1.[ItemCode],2,8) END Model");
            sb.Append("          FROM  ACMESQLSP.DBO.WH_Item3 T1 ");
            sb.Append("          left join oitm t2 on (t1.itemcode=t2.itemcode COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("          WHERE   ISNULL(ShipDate,'') <> ''  and t2.itmsgrpcod=1032");
            sb.Append("          UNION ALL SELECT '' ");
            sb.Append("          order by CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("                  SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("                 AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append("          Substring (T1.[ItemCode],2,8) END");

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


        private System.Data.DataTable GetVER()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT distinct Substring(T1.[ItemCode],12,1)  VER");
            sb.Append("              FROM  ACMESQLSP.DBO.WH_Item3 T1 ");
            sb.Append("              left join oitm t2 on (t1.itemcode=t2.itemcode COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("              WHERE   ISNULL(ShipDate,'') <> ''  and t2.itmsgrpcod=1032");
            sb.Append("              UNION ALL SELECT '' ");
            sb.Append("              ORDER BY Substring(T1.[ItemCode],12,1)");
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
        public void UPINV(string a5, string a3, string a6, string a7, string a4, string SHIPPINGCODE, string SEQNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_ITEM3 SET a5=@a5,a3=@a3,a6=@a6,a7=@a7,a4=@a4 WHERE SHIPPINGCODE=@SHIPPINGCODE AND SEQNO=@SEQNO AND A2 IS NOT NULL ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@a5", a5));
            command.Parameters.Add(new SqlParameter("@a3", a3));
            command.Parameters.Add(new SqlParameter("@a6", a6));
            command.Parameters.Add(new SqlParameter("@a7", a7));
            command.Parameters.Add(new SqlParameter("@a4", a4));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));
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
        public void UPINV2(string a2, string SHIPPINGCODE, string SEQNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_ITEM3 SET a2=@a2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND SEQNO=@SEQNO ", connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@a2", a2));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));
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
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {

                    DataGridViewRow row;

                    row = dataGridView1.Rows[i];
                    string 進貨日期 = row.Cells["A5"].Value.ToString();
                    string 後續處理情形 = row.Cells["A3"].Value.ToString();
                    string 運送貨代 = row.Cells["A6"].Value.ToString();
                    string SHIPPING備註 = row.Cells["A7"].Value.ToString();
                    string 結案 = row.Cells["結案"].Value.ToString();

                    string WH工單 = row.Cells["ShippingCode"].Value.ToString();
                    string 序號 = row.Cells["SeqNo"].Value.ToString();

                    UPINV(進貨日期, 後續處理情形, 運送貨代, SHIPPING備註, 結案, WH工單, 序號);

                }

                if (A1 != "")
                {

                    UPINV2(ta2.Text, A1, A2);
                }

                MessageBox.Show("更新成功");

                button1_Click(sender, e);


            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message);  }
        }

   

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        private void btnExcelImport_Click(object sender, EventArgs e)
        {
            try
            {

                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "請選擇上傳檔案";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    readExcel(dialog.FileName);//利用COM元件讀EXCEL
                    //ImportExcel(dialog.FileName);//利用OLEDB讀EXCEL
                }
                
                button1_Click(sender,e);
            }
            catch (Exception ex)
            {

            }


            
        }
        public void ImportExcel( string fileName) 
        {
            try
            {
                
                //定義OleDb======================================================
                //1.檔案位置
                string filepath = fileName;
                //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
                string ProviderName = "Microsoft.Jet.OLEDB.4.0;";
                //3.Excel版本，Excel 8.0 針對Excel2及以上版本，Excel5.0 針對Excel97。
                string ExtendedString = "'Excel 5.0;";
                //4.第一行是否為標題(;結尾區隔)
                string HDR = "Yes;";

                //5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取(;結尾區隔,'文字結尾)
                string IMEX = "1';";

                //=============================================================
                //連線字串
                string connectString =
                        "Data Source=" + fileName + ";" +
                        "Provider=" + ProviderName +
                        "Extended Properties=" + ExtendedString +
                        "HDR=" + HDR +
                        "IMEX=" + IMEX;
                //=============================================================
                OleDbConnection myConn = new OleDbConnection(connectString);
                myConn.Open();
                string strCom = "select * from [Sheet1$] ";
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                System.Data.DataTable table = new System.Data.DataTable("table");
                myCommand.Fill(table);
                myConn.Close();
                DataRow[] rows = table.Select("[F3] is not null");
                System.Data.DataTable dt = null;
                System.Data.DataRow dr = null;
                int line = 0;
                int Cartontmp = 0;
                

                /*
                _Application app = new Microsoft.Office.Interop.Excel.Application();
                _Workbook wk = null;
                _Worksheet sheet = null;
                Range range = null;

                app.Visible = false;*/
                
                
                dt = Maketable();

                StringBuilder sb = new StringBuilder();


                foreach (DataRow row in rows)
                {
                    if (row[0].ToString() == "收貨工單")
                    {
                        continue;
                    }
                    line++;
                    string[] num = new string[2];
                    num = row[0].ToString().Split('-');
                    dr = dt.NewRow();
                    dr["ShippingCode"] = row[0].ToString();
                    dr["FrgnName"] = row[1].ToString();
                    dr["ItemCode"] = row[2].ToString();
                    dr["BoxCheck"] = row[3].ToString();
                    dr["ShipDate"] = row[5].ToString();
                    dr["A2"] = row[6].ToString() + " " + row[7].ToString() + " " + row[8].ToString();
                    string[] date = row[4].ToString().Split('/');
                    string year = date[0];
                    string month = date[1].Length == 1 ? "0" + date[1] : date[1];
                    string day = date[2].Length == 1 ? "0" + date[2] : date[2];
                    dr["A5"] = year + month + day;
                    dr["DeCust"] = row[10].ToString();
                    dr["A3"] = "";
                    dr["A7"] = "";
                    dr["A6"] = row[10].ToString();
                    dt.Rows.Add(dr);
                }
                if (!CheckShippingCodeExistD(dr["ShippingCode"].ToString()))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string TableKey = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
                        GenerateSQL.InsertDataRow(row, TableKey);

                    }
                    MessageBox.Show(fileName + "-上傳完成");
                }
                else
                {
                    MessageBox.Show(fileName + "已存在");
                }
                tWHNO.Text = dr["ShippingCode"].ToString();
            }
            catch (Exception ex) 
            {

            }
            
            

        }
        public void readExcel(string path)
        {
            try
            {
                Excel.Application excel1;
                Excel.Workbooks wbs = null;
                Excel.Workbook wb = null;
                Excel.Sheets sheet;
                Excel.Worksheet ws = null;

                object miss = System.Reflection.Missing.Value;
                excel1 = new Excel.Application();
                excel1.UserControl = true;
                excel1.DisplayAlerts = false;
                excel1.Application.Workbooks.Open(path, miss, miss, miss, miss,
                                                 miss, miss, miss, miss,
                                                 miss, miss, miss, miss,
                                                 miss, miss);
                wbs = excel1.Workbooks;
                sheet = wbs[1].Worksheets;
                ws = (Excel.Worksheet)sheet.get_Item(1);
                int rowNum = ws.UsedRange.Cells.Rows.Count;
                int colNum = ws.UsedRange.Cells.Columns.Count;
                int seq = 0;
                DataTable dt = new DataTable();
                DataView dv;
                System.Data.DataRow dr = null;
                dt = Maketable();
                string cellStr = null;
                char ch = 'A';
                string ShippingCode = "";
                for (int i = 1; i < rowNum; i++)
                {
                    ch = 'A';
                    dr = dt.NewRow();
                    for (int j = 0; j < colNum; j++)
                    {
                        cellStr = ch.ToString() + (i+1).ToString();
                        if (ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString() == null) 
                        {
                            continue;
                        }
                        else
                        {
                            dr[j] = ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString();
                        }
                        
                        
                        ch++;
                    }
                    dt.Rows.Add(dr);
                }
                DataRow[] rows = dt.Select("FrgnName IS NOT null");
                DataTable table = new DataTable();
                table = Maketable();
                foreach (DataRow row in rows)
                {
                    
                    if (row[0].ToString() == "收貨工單" || row[0].ToString() == "")
                    {
                        continue;
                    }
                    seq++;
                    string[] num = new string[2];
                    num = row[0].ToString().Split('-');
                    dr = table.NewRow();
                    dr["ShippingCode"] = row[0].ToString();
                    dr["FrgnName"] = row[1].ToString();
                    dr["ItemCode"] = row[2].ToString();
                    dr["BoxCheck"] = row[3].ToString();
                    dr["ShipDate"] = row[5].ToString();
                    dr["A2"] = row[6].ToString() + " " + row[7].ToString() + " " + row[8].ToString();
                    string[] date = row[4].ToString().Split('/');
                    string year = date[0];
                    string month = date[1].Length == 1 ? "0" + date[1] : date[1];
                    string day = date[2].Length == 1 ? "0" + date[2] : date[2];
                    dr["A5"] = year + month + day;
                    dr["PiNo"] = row[9].ToString();
                    dr["DeCust"] = "";//客戶別
                    dr["A3"] = "";
                    dr["A7"] = "";
                    dr["A6"] = row[10].ToString();
                    dr["SeqNo"] = seq;
                    table.Rows.Add(dr);
                }
                if (!CheckShippingCodeExistD(table.Rows[0]["ShippingCode"].ToString()))
                {
                    foreach (DataRow row in table.Rows)
                    {

                        string TableKey = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
                        GenerateSQL.InsertDataRow(row, TableKey);

                    }
                    MessageBox.Show(path + "-上傳完成");
                }
                else
                {
                    var result = MessageBox.Show(path + "已存在,是否重新上傳", "重複上傳", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                    if (result == DialogResult.Yes) 
                    {
                        DELETE_WH_ITEM3(dr["ShippingCode"].ToString());

                        foreach (DataRow row in table.Rows)
                        {

                            string TableKey = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
                            GenerateSQL.InsertDataRow(row, TableKey);

                        }
                        MessageBox.Show(path + "-覆蓋上傳完成");

                    }
                }
                tWHNO.Text = dr["ShippingCode"].ToString();
            }
            catch (Exception ex) 
            {

            }
            
        }
        private System.Data.DataTable Maketable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ShippingCode", typeof(string));//0 WH工單

            dt.Columns.Add("FrgnName", typeof(string));//1 SI工單

            dt.Columns.Add("ItemCode", typeof(string));//2 產品編號

            dt.Columns.Add("BoxCheck", typeof(string));//3 InvoiceNo

            dt.Columns.Add("ShipDate", typeof(string));//4 異常數量

            dt.Columns.Add("A5", typeof(string));//5 進倉日期

            dt.Columns.Add("A2", typeof(string));//6 異常情況

            dt.Columns.Add("DeCust", typeof(string));//7

            dt.Columns.Add("A3", typeof(string));//8

            dt.Columns.Add("A7", typeof(string));//9

            dt.Columns.Add("A6", typeof(string));//10

            dt.Columns.Add("PiNo", typeof(string));//11

            dt.Columns.Add("SeqNo", typeof(int));

            dt.TableName = "WH_Item3";


            return dt;
        }
        private System.Data.DataTable Makedownloadtable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("trgtPath", typeof(string));

            dt.Columns.Add("FileName", typeof(string));

            dt.Columns.Add("Date", typeof(string));

            dt.Columns.Add("FileExt", typeof(string));

            dt.TableName = "download";


            return dt;
        }
        private void DELETE_WH_ITEM3(string WH) 
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Delete From WH_Item3 where ShippingCode like @WH and ShipDate <> '' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WH", WH));

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
        private bool CheckShippingCodeExistD(string InvoiceNo)
        {
            System.Data.DataTable dt = GetData(string.Format("select * from WH_Item3 where ShippingCode='{0}' and ShipDate <> ''", InvoiceNo));

            if (dt.Rows.Count == 0)
                return false;
            else
                return true;
        }
        public System.Data.DataTable GetData(string Sql)
        {
            SqlConnection connection = new SqlConnection(ShipConnectiongString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(Sql);
            command.CommandType = CommandType.Text;
            command.CommandText = sb.ToString();
            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_Stage");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_Stage"];
        }
        public static System.Data.DataTable GetAtc1(string shippingcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select trgtPath 路徑,FileName 檔案名稱,Date 日期,FileExt");
            sb.Append(" from  ATC1 T3 ");
            sb.Append(" left join owtr t4 on(T3.ABSENTRY) = (t4.atcentry) ");
            sb.Append(" where t4.jrnlmemo =@shippingcode ");

            

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            catch (Exception ex) 
            {

            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }

        private void dgvFile_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /*try 
            { 
                DataGridView dgv = (DataGridView)sender;


                
                int i = e.RowIndex;
                DataRow drw = dgv.Rows[i];
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string aa = drw["trgtPath"].ToString();
                string filename = drw["FileName"].ToString();
                string FileExt = drw["FileExt"].ToString();
                string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                System.IO.File.Copy(aa +"\\"+ filename +"."+ FileExt, NewFileName, true);
                System.Diagnostics.Process.Start(NewFileName);
                
                //DataGridViewLinkCell cell = (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];
                //cell.LinkVisited = true;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }*/
}

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowNum = e.RowIndex;
            string shippingcode = dataGridView1.Rows[rowNum].Cells["ShippingCode"].Value.ToString();
            DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            System.Data.DataTable dt = GetAtc1(shippingcode);
            dgvFile.DataSource = dt;
            if (cell.ColumnIndex == 1)
            {

                string s = dataGridView1.Rows[rowNum].Cells["ShippingCode"].Value.ToString();
                WH_main a = new WH_main();
                a.PublicString = s;
                a.ShowDialog();
            }
            else if (cell.ColumnIndex == 2) 
            {
                string s = dataGridView1.Rows[rowNum].Cells["FrgnName"].Value.ToString();
                APShip a = new APShip();
                a.PublicString = s;
                a.ShowDialog();
            }
            else
            {
                /*K1 = Makedownloadtable();
                K1 = GetAtc1(dataGridView1.Rows[0].Cells["ShippingCode"].Value.ToString());
                dgvFile.DataSource = K1;*/

                if (cell.Value.ToString() == "點擊下載")
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow rows = dt.Rows[i];
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string aa = rows["路徑"].ToString();
                        string filename = rows["檔案名稱"].ToString();
                        string FileExt = rows["FileExt"].ToString();
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa + "\\" + filename + "." + FileExt, NewFileName, true);
                        System.Diagnostics.Process.Start(NewFileName);
                    }
                }
            }
        }

        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 15) 
                {
                    int rowNum = e.RowIndex;
                    string shippingcode = dataGridView1.Rows[rowNum].Cells["ShippingCode"].Value.ToString();
                    System.Data.DataTable dt = GetAtc1(shippingcode);
                    dgvFile.DataSource = dt;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                            DataRow rows = dt.Rows[i];
                            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                            string aa = rows["路徑"].ToString();
                            string filename = rows["檔案名稱"].ToString();
                            string FileExt = rows["FileExt"].ToString();
                            string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                            System.IO.File.Copy(aa + "\\" + filename + "." + FileExt, NewFileName, true);
                            System.Diagnostics.Process.Start(aa + "\\" + filename + "." + FileExt, NewFileName + "." + FileExt);
                    }
                    
                }
                
            }
            catch (Exception ex) 
            {

            }
            
        }

        
    }
}
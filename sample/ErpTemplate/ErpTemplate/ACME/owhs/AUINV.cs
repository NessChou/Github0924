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
                dataGridView1.Rows[i].Cells["download"].Value = "�I���U��";
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
            sb.Append(" SELECT ShippingCode WH�u��,FrgnName SI�u��,SeqNo �Ǹ�,ItemCode ���~�s��,BoxCheck ��tINVOICE,A5 �i�f���,ShipDate ���`�ƶq,");
            sb.Append(" a2 ���`���p,DeCust �Ȥ�O,A3 ����B�z����,A6 �B�e�f�N,A7 SHIPPING�Ƶ�,PiNo �����Ӥ����X,case isnull(a4,'') when '' then 'False' else a4 end ����,'�I���U��' �U�� FROM WH_Item3 ");
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
                    string �i�f��� = row.Cells["A5"].Value.ToString();
                    string ����B�z���� = row.Cells["A3"].Value.ToString();
                    string �B�e�f�N = row.Cells["A6"].Value.ToString();
                    string SHIPPING�Ƶ� = row.Cells["A7"].Value.ToString();
                    string ���� = row.Cells["����"].Value.ToString();

                    string WH�u�� = row.Cells["ShippingCode"].Value.ToString();
                    string �Ǹ� = row.Cells["SeqNo"].Value.ToString();

                    UPINV(�i�f���, ����B�z����, �B�e�f�N, SHIPPING�Ƶ�, ����, WH�u��, �Ǹ�);

                }

                if (A1 != "")
                {

                    UPINV2(ta2.Text, A1, A2);
                }

                MessageBox.Show("��s���\");

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
                dialog.Title = "�п�ܤW���ɮ�";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    readExcel(dialog.FileName);//�Q��COM����ŪEXCEL
                    //ImportExcel(dialog.FileName);//�Q��OLEDBŪEXCEL
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
                
                //�w�qOleDb======================================================
                //1.�ɮצ�m
                string filepath = fileName;
                //2.���Ѫ̦W��  Microsoft.Jet.OLEDB.4.0�A�Ω�2003�H�e�����AMicrosoft.ACE.OLEDB.12.0 �A�Ω�2007�H�᪺�����B�z xlsx �ɮ�
                string ProviderName = "Microsoft.Jet.OLEDB.4.0;";
                //3.Excel�����AExcel 8.0 �w��Excel2�ΥH�W�����AExcel5.0 �w��Excel97�C
                string ExtendedString = "'Excel 5.0;";
                //4.�Ĥ@��O�_�����D(;�����Ϲj)
                string HDR = "Yes;";

                //5.IMEX=1 �q���X�ʵ{�ǩl�ױN�u���V�v�ƾڦC�@���奻Ū��(;�����Ϲj,'��r����)
                string IMEX = "1';";

                //=============================================================
                //�s�u�r��
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
                    if (row[0].ToString() == "���f�u��")
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
                    MessageBox.Show(fileName + "-�W�ǧ���");
                }
                else
                {
                    MessageBox.Show(fileName + "�w�s�b");
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
                    
                    if (row[0].ToString() == "���f�u��" || row[0].ToString() == "")
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
                    dr["DeCust"] = "";//�Ȥ�O
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
                    MessageBox.Show(path + "-�W�ǧ���");
                }
                else
                {
                    var result = MessageBox.Show(path + "�w�s�b,�O�_���s�W��", "���ƤW��", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                    if (result == DialogResult.Yes) 
                    {
                        DELETE_WH_ITEM3(dr["ShippingCode"].ToString());

                        foreach (DataRow row in table.Rows)
                        {

                            string TableKey = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
                            GenerateSQL.InsertDataRow(row, TableKey);

                        }
                        MessageBox.Show(path + "-�л\�W�ǧ���");

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
            dt.Columns.Add("ShippingCode", typeof(string));//0 WH�u��

            dt.Columns.Add("FrgnName", typeof(string));//1 SI�u��

            dt.Columns.Add("ItemCode", typeof(string));//2 ���~�s��

            dt.Columns.Add("BoxCheck", typeof(string));//3 InvoiceNo

            dt.Columns.Add("ShipDate", typeof(string));//4 ���`�ƶq

            dt.Columns.Add("A5", typeof(string));//5 �i�ܤ��

            dt.Columns.Add("A2", typeof(string));//6 ���`���p

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

            sb.Append(" select trgtPath ���|,FileName �ɮצW��,Date ���,FileExt");
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

                if (cell.Value.ToString() == "�I���U��")
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow rows = dt.Rows[i];
                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string aa = rows["���|"].ToString();
                        string filename = rows["�ɮצW��"].ToString();
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
                            string aa = rows["���|"].ToString();
                            string filename = rows["�ɮצW��"].ToString();
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
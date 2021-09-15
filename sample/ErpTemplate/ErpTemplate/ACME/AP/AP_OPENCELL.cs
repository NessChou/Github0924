using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using SAPbobsCOM;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class AP_OPENCELL : Form
    {
        public AP_OPENCELL()
        {
            InitializeComponent();
        }

        private void aP_OPENCELLBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_OPENCELLBindingSource.EndEdit();
            this.aP_OPENCELLTableAdapter.Update(this.lC.AP_OPENCELL);

        }

        private void AP_OPENCELL_Load(object sender, EventArgs e)
        {

            this.aP_OPENCELLTableAdapter.Fill(this.lC.AP_OPENCELL);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.aP_OPENCELLTableAdapter.FillBy(this.lC.AP_OPENCELL, textBox1.Text);
        }
        public System.Data.DataTable GETPARTNO(string ITEMCODE)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT U_PARTNO FROM OITM WHERE ITEMCODE=@ITEMCODE ");



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
        private void aP_OPENCELLDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (aP_OPENCELLDataGridView.Columns[e.ColumnIndex].Name == "KIT")
                {

                    string KIT = this.aP_OPENCELLDataGridView.Rows[e.RowIndex].Cells["KIT"].Value.ToString();
                    System.Data.DataTable G1 = GETPARTNO(KIT);
                    if (G1.Rows.Count > 0)
                    {
                        this.aP_OPENCELLDataGridView.Rows[e.RowIndex].Cells["PARTNO"].Value = G1.Rows[0][0].ToString();
                    }

                }




            }
            catch
            {

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(aP_OPENCELLDataGridView);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aP_OPENCELLBindingSource.EndEdit();
            this.aP_OPENCELLTableAdapter.Update(this.lC.AP_OPENCELL);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(aP_OPENCELLDataGridView);
        }
        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            if (btn.Text.ToString() == "OC對應TCON")
            {
                DELPENCELL2();
                System.Data.DataTable dt = GetDgvToTable(aP_OPENCELLDataGridView);

                string date = DateTime.Now.ToString("yyyy/MM");
                string FileNameTemplate = GetExePath() + "\\Excel\\wh\\OC對應TCON.xls";
                string FileName = GetExePath() + "\\Excel\\temp\\" + date + "OC對應TCON.xls";
                System.Data.DataTable table = GetDtData(dt,true);
                System.Data.DataTable dtData = GetDataSort(table, "notZero");//把一樣的相加 小計不為零
                System.Data.DataTable dtDataZero = GetDataSort(table, "Zero");//把一樣的相加 小計為零

                WriteDataTableToExcel(dtData, dtDataZero, FileNameTemplate, FileName);
            }
            else if (btn.Text.ToString() == "沒有TCON")
            {
                DELPENCELL2();
                System.Data.DataTable dt = GetDgvToTable(aP_OPENCELLDataGridView);

                string date = DateTime.Now.ToString("yyyy/MM");
                string FileNameTemplate = GetExePath() + "\\Excel\\wh\\OC沒有TCON.xls";
                string FileName = GetExePath() + "\\Excel\\temp\\" + date + "TCON缺少均價.xls";
                System.Data.DataTable table = GetDtData(dt,false);
                System.Data.DataTable DtSort = SortTable(table);

                WriteDataTableToExcel(DtSort, FileNameTemplate, FileName);
            }

            
        }
     
        private System.Data.DataTable GetDgvToTable(DataGridView dgv)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            // 列強制轉換
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循環行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private System.Data.DataTable GetDataSort(System.Data.DataTable dtData, string flag)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            DataRow rw;
            dt = MakeTable();

            foreach (DataRow row in dtData.Rows)
            {
                if (flag == "notZero" && row["小計"].ToString() != "0" && row["小計"].ToString() != "")
                {
                    //避免重複加入
                    int onhand = 0;//庫存量
                    int wait = 0;//待進貨量
                    int sum = 0;//小計
                    int Tonhand = 0;//T庫存量
                    int Twait = 0;//T待進貨量
                    int Tsum = 0;//T小計
                    string test = row["KIT"].ToString();
                    if (row["KIT"].ToString() != "" && row["KIT"].ToString() != null)
                    {
                        if (dt.Select("項目編號 = '" + row["項目編號"].ToString() + "' and KIT LIKE  '" + row["KIT"].ToString().Substring(0, 14) + "*'").Length != 0)
                        {
                            continue;
                        }
                        //DataRow[] rws = dtData.Select("項目編號 = '" + row["項目編號"].ToString() + "' and KIT like '" + row["KIT"].ToString().Substring(0,14) + "*'");


                        //var rws = dtData.AsEnumerable().Where(r => r.Field<string>("KIT").Contains(row["KIT"].ToString().Substring(0, 14)) && r.Field<string>("項目編號") == row["項目編號"].ToString());

                        var rws = from dr in dtData.AsEnumerable()
                                  where dr.Field<string>("項目編號") == row["項目編號"].ToString() && dr.Field<string>("KIT") != null && dr.Field<string>("KIT").Substring(0, 14) == row["KIT"].ToString().Substring(0, 14)
                                  select dr;

                        int rowcount = rws.Count();
                        //int rowcount = rws.Length;
                        foreach (DataRow rwss in rws)
                        {
                            onhand += Convert.ToInt32(rwss["庫存量"]);
                            wait += Convert.ToInt32(rwss["待進貨量"]);
                            sum += Convert.ToInt32(rwss["小計"]);
                            Tonhand = Convert.ToInt32(rwss["T庫存量"]);
                            Twait = Convert.ToInt32(rwss["T待進貨量"]);
                            Tsum = Convert.ToInt32(rwss["T小計"]);
                        }

                    }
                    else
                    {
                        DataRow[] rws = dtData.Select("項目編號 = '" + row["項目編號"].ToString() + "' and KIT = null");
                        onhand += Convert.ToInt32(row["庫存量"]);
                        wait += Convert.ToInt32(row["待進貨量"]);
                        sum += Convert.ToInt32(row["小計"]);
                        Tonhand = Convert.ToInt32(row["T庫存量"]);
                        Twait = Convert.ToInt32(row["T待進貨量"]);
                        Tsum = Convert.ToInt32(row["T小計"]);


                    }

                    rw = dt.NewRow();
                    rw["項目編號"] = row["項目編號"].ToString();
                    rw["BU"] = row["BU"].ToString();
                    rw["KIT"] = row["KIT"].ToString();
                    rw["PartNo"] = row["PartNo"].ToString();
                    rw["庫存量"] = onhand;
                    rw["待進貨量"] = wait;
                    rw["小計"] = sum;
                    rw["T庫存量"] = Tonhand;
                    rw["T待進貨量"] = Twait;
                    rw["T小計"] = Tsum;
                    dt.Rows.Add(rw);

                    if (onhand - Tonhand > 0)
                    {
                        string ITEMCODE = row["項目編號"].ToString();
                        //if (ITEMCODE == "O320DVN02.000")
                        //{
                        //    MessageBox.Show("as");
                        //}
                        //System.Data.DataTable K1 = GETS1(ITEMCODE);
                        //if (K1.Rows.Count > 0)
                        //{
                          //  decimal k2 = 0;
                            //for (int i = 0; i <= K1.Rows.Count - 1; i++)
                            //{
                              //  string KITEM = K1.Rows[i][0].ToString();

                                decimal k2 = Convert.ToDecimal(GETS2(ITEMCODE).Rows[0][0]);

                              //  k2 += PRICE;
                           // }
                        //    k2 = k2 / (K1.Rows.Count);

                            decimal k1 = Convert.ToDecimal(onhand - Tonhand);
                            ADDOPENCELL2(ITEMCODE, onhand - Tonhand, k2 * k1);
                        //}
                    }

                    }
                else if (flag == "Zero" && row["小計"].ToString() == "0" && row["小計"].ToString() != "")
                {
                    //避免重複加入
                    int onhand = 0;//庫存量
                    int wait = 0;//待進貨量
                    int sum = 0;//小計
                    int Tonhand = 0;//T庫存量
                    int Twait = 0;//T待進貨量
                    int Tsum = 0;//T小計
                    DataRow[] rws = dtData.Select("項目編號 = '" + row["項目編號"].ToString() + "' and KIT = '" + row["KIT"].ToString() + "'");
                    for (int i = 0; i < rws.Length; i++)
                    {
                        onhand = Convert.ToInt32(rws[i]["庫存量"]);
                        wait = Convert.ToInt32(rws[i]["待進貨量"]);
                        sum = Convert.ToInt32(rws[i]["小計"]);
                        Tonhand = Convert.ToInt32(rws[i]["T庫存量"]);
                        Twait = Convert.ToInt32(rws[i]["T待進貨量"]);
                        Tsum = Convert.ToInt32(rws[i]["T小計"]);
                    }
                    rw = dt.NewRow();
                    rw["項目編號"] = row["項目編號"].ToString();
                    rw["BU"] = row["BU"].ToString();
                    rw["KIT"] = row["KIT"].ToString();
                    rw["PartNo"] = row["PartNo"].ToString();
                    rw["庫存量"] = row["庫存量"].ToString();
                    rw["待進貨量"] = row["待進貨量"].ToString();
                    rw["小計"] = row["小計"].ToString();
                    rw["T庫存量"] = 0;
                    rw["T待進貨量"] = 0;
                    rw["T小計"] = 0;
                    dt.Rows.Add(rw);
                }
            }


            return dt;
        }
        private System.Data.DataTable SortTable(System.Data.DataTable table) 
        {
            System.Data.DataTable dt = MakeNoTconTable();
            DataRow row;
            string OcTcon = "";
            foreach (DataRow Rows in table.Rows) 
            {
                if (OcTcon.Contains(Rows["KIT"].ToString()) || OcTcon.Contains(Rows["項目編號"].ToString()) || Rows["項目編號"].ToString().Substring(0,1) == "P")
                {
                    //以KIT為判斷依據,但有時候一個項目編號會對應到兩個KIT,所以前面的項目編號若已有足夠KIT則不繼續排序,加入OcTcon以免重複排序
                    continue;
                }
                if (Rows["項目編號"].ToString() == "O320HVN05.56012" || Rows["項目編號"].ToString() == "O430QVN02.01007" || Rows["項目編號"].ToString() == "O430QVN01.00002" || Rows["項目編號"].ToString() == "O500HVN07.05002"  ) 
                {

                }
                DataRow[] KIT = table.Select("項目編號 = '" + Rows["項目編號"].ToString() + "'");//有可能一個項目編號對應兩個KIT
                DataRow[] rows = table.Select("KIT = '" + Rows["KIT"].ToString() + "'");
                int OpenCellCount = 0;
                int TconCount = Convert.ToInt32(Rows["T小計"]);
                int Price = 0;
                string ItemCode = Rows["項目編號"].ToString();
                if (KIT.Length > 1)
                {
                    //一個料號對應兩個以上KIT重記TconCount
                    TconCount = 0;
                    foreach (DataRow kit in KIT) 
                    {
                        TconCount += Convert.ToInt32(kit["T小計"]); 
                    }
                }
               
               
                foreach (DataRow rws in rows)
                {
                    OpenCellCount += Convert.ToInt32(rws["小計"]);
                    /*
                    if (Convert.ToInt32(rws["AvgPrice"]) > Price)
                    {
                        Price = Convert.ToInt32(rws["AvgPrice"]);
                        ItemCode = rws["項目編號"].ToString();//顯示的不要小計為零,用最高價的為當項項目編號
                    }*/
                    OcTcon += Rows["項目編號"].ToString() + ",";

                }
                if (OpenCellCount > TconCount) 
                {
                    foreach (DataRow rws in rows) 
                    {
                        Price = Convert.ToInt32(rws["AvgPrice"]);

                        row = dt.NewRow();
                        row["BU"] = Rows["BU"].ToString();
                        row["項目編號"] = rws["項目編號"];
                        row["數量"] = (OpenCellCount - TconCount).ToString();
                        row["價格"] = Price.ToString();
                        row["總計"] = (OpenCellCount - TconCount) * Price;
                        dt.Rows.Add(row);
                    }
                   
                    
                }
                OcTcon += Rows["KIT"].ToString() + ",";
            }

            return dt;
        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }

        public System.Data.DataTable GETS1(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            //    sb.Append(" SELECT AVG(ISNULL(STOCKVALUE,0))  STOCKVALUE FROM OITM P WHERE SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)=@ITEMCODE AND ISNULL(STOCKVALUE,0) <> 0 ");
            sb.Append(" SELECT ITEMCODE FROM OITM P");
            sb.Append(" WHERE SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)=@ITEMCODE");
            sb.Append(" AND P.ONHAND>0");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable GETS2(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            //    sb.Append(" SELECT AVG(ISNULL(STOCKVALUE,0))  STOCKVALUE FROM OITM P WHERE SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)=@ITEMCODE AND ISNULL(STOCKVALUE,0) <> 0 ");
            sb.Append("  SELECT PRICE FROM PDN1 WHERE ITEMCODE=@ITEMCODE ORDER BY DOCDATE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        private System.Data.DataTable GetDtData(System.Data.DataTable dt,bool tag)
        {
            string SapConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            SqlConnection connection = new SqlConnection(SapConnectiongString);//"server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02"
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            StringBuilder sb = new StringBuilder();
            //不同等級合併
            //sb.Append("  SELECT DISTINCT(SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)) 項目編號, cast(P.OnHand as int) as 庫存量, (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=p.itemcode ) 待進貨量,(cast(P.OnHand as int) + (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=p.itemcode )) as 小計,  ");
            //不同等級展開
            sb.Append("  SELECT  DISTINCT(P.ItemCode) 項目編號,P.U_BU AS BU, cast(P.OnHand as int) as 庫存量, ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0 WHERE T0.ItemCode = P.ItemCode AND LineStatus <> 'C') as int),0)  + CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 LEFT JOIN OWOR T0 ON(T0.DOCENTRY = T1.DOCENTRY) WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN('C', 'L') AND T1.ITEMCODE = P.ITEMCODE), 0) AS INT)  待進貨量,(cast(P.OnHand as int) + (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=p.itemcode )) as 小計,  ");
            if (tag == false)
            {
                sb.Append(" P.AvgPrice,");

            }
            sb.Append(" t3.KIT,t4.itemname PartNo,ISNULL(t4.OnHand,0) T庫存量, (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=t4.itemcode ) T待進貨量,ISNULL(t4.OnHand,0) + (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=t4.itemcode ) T小計");
            sb.Append(" FROM OITM P   left join por1 t2 on t2.itemcode = P.itemcode   ");
            sb.Append("  LEFT JOIN (SELECT MAX(KIT) KIT, OPENCELL FROM AcmeSqlSP.DBO.AP_OPENCELL WHERE ISNULL(KIT, '') <> ''GROUP BY SUBSTRING(KIT, 0, LEN(KIT)), OPENCELL) T3 ON(t2.itemcode = t3.opencell COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("  left join OITM t4 on t3.KIT = t4.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS   ");
            sb.Append("where len(P.ItemCode)=15 And ISNULL(P.U_GROUP,'') <> 'Z&R-費用類群組' AND P.FROZENFOR = 'N' AND P.CANCELED = 'N' AND  ");

            for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                /*
                //原本不同等級要合併
                if (i == 0)
                {
                    sb.Append(" (SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)) ='" + dt.Rows[i]["OPENCELL"].ToString().Substring(0, 10) + dt.Rows[i]["OPENCELL"].ToString().Substring(11, 3) + "' ");
                }
                else
                {
                    sb.Append(" OR (SUBSTRING(P.ItemCode,0,11)+SUBSTRING(P.ItemCode,12,3)) ='" + dt.Rows[i]["OPENCELL"].ToString().Substring(0, 10) + dt.Rows[i]["OPENCELL"].ToString().Substring(11, 3) + "' ");
                }
                */
                //不同等級展開
                if (i == 0)
                {
                    sb.Append(" P.ItemCode ='" + dt.Rows[i]["OPENCELL"].ToString()  + "' ");
                }
                else
                {
                    sb.Append(" OR P.ItemCode ='" + dt.Rows[i]["OPENCELL"].ToString() + "' ");
                }
            }
            sb.Append(" ORDER BY t4.itemname");

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
        private void WriteDataTableToExcel(System.Data.DataTable dt, System.Data.DataTable dtzero, string DirTemplate, string Dir)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            // Microsoft.Office.Interop.Excel.Range excelCellrange;
            object oMissing = System.Reflection.Missing.Value;


            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            excel.DisplayAlerts = false;



            //Interop params
            string ItemCode = "";
            string ItemCodeAll = "";
            foreach (DataRow itemcode in dt.Rows)
            {
                string ItemCodetmp = Convert.ToString(itemcode["項目編號"]).Split('.')[0] + "." + Convert.ToString(itemcode["項目編號"]).Split('.')[1].Substring(1, 1);//ex G170ETN01.00022 => G170ETN01.0 小數點前加小數點後第二位
                if (!ItemCode.Contains(ItemCodetmp))
                {
                    ItemCode += ItemCodetmp + ",";
                    ItemCodeAll += Convert.ToString(itemcode["項目編號"]) + ",";//完整的字串 之後做sort用
                }

            }
            ItemCode = ItemCode.Substring(0, ItemCode.Length - 1);
            int ItemCodeCount = ItemCode.Split(',').Length;

            try
            {


                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                for (int i = 0; i < 2; i++)
                {
                    // Workk sheet
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    excelSheet.Name = "OC對應TCON";
                    WriteDataTableToSheetByArray(dt, excelSheet);


                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    excelSheet.Name = "小計為零";
                    WriteDataTableToSheetByArray(dtzero, excelSheet);



                    //now save the workbook and exit Excel
                    //excelworkBook.SaveAs(saveAsLocation);
                    excelworkBook.SaveAs(Dir, XlFileFormat.xlWorkbookNormal,
                          "", "", Type.Missing, Type.Missing,
                        XlSaveAsAccessMode.xlNoChange,
                        1, false, Type.Missing, Type.Missing, Type.Missing);
                }


                SheetTemplate.Delete();
                excelworkBook.Close();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                System.Diagnostics.Process.Start(Dir);

            }
        }
        public void DELPENCELL2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE AP_OPENCELL2 WHERE USERID=@USERID", connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@USERID", fmLogin.LoginID.ToString()));


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
        private void WriteDataTableToExcel(System.Data.DataTable dt, string DirTemplate, string Dir)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            // Microsoft.Office.Interop.Excel.Range excelCellrange;
            object oMissing = System.Reflection.Missing.Value;


            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            excel.DisplayAlerts = false;



            //Interop params
            string ItemCode = "";
            string ItemCodeAll = "";

            try
            {


                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                for (int i = 0; i < 2; i++)
                {
                    // Workk sheet
                    SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                    excelSheet.Name = "OC不足TCON";
                    WriteDataTableToSheetByArray(dt, excelSheet);

                    //now save the workbook and exit Excel
                    //excelworkBook.SaveAs(saveAsLocation);
                    excelworkBook.SaveAs(Dir, XlFileFormat.xlWorkbookNormal,
                          "", "", Type.Missing, Type.Missing,
                        XlSaveAsAccessMode.xlNoChange,
                        1, false, Type.Missing, Type.Missing, Type.Missing);
                }


                SheetTemplate.Delete();
                excelworkBook.Close();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                System.Diagnostics.Process.Start(Dir);

            }
        }
 
        public void ADDOPENCELL2(string KIT, int STOCK, decimal AMT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPENCELL2(KIT,STOCK,AMT,USERID) values(@KIT,@STOCK,@AMT,@USERID)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@KIT", KIT));
            command.Parameters.Add(new SqlParameter("@STOCK", STOCK));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));

            command.Parameters.Add(new SqlParameter("@USERID", fmLogin.LoginID.ToString()));


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

            dt.Columns.Add("BU", typeof(string));

            dt.Columns.Add("項目編號", typeof(string));

            dt.Columns.Add("庫存量", typeof(int));


            dt.Columns.Add("待進貨量", typeof(int));

            dt.Columns.Add("小計", typeof(int));

            dt.Columns.Add("KIT", typeof(string));
            //TCON
            dt.Columns.Add("PartNo", typeof(string));

            dt.Columns.Add("T庫存量", typeof(int));
            //TCON
            dt.Columns.Add("T待進貨量", typeof(int));
            //TCON
            dt.Columns.Add("T小計", typeof(int));


            dt.TableName = "dt";

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }
        private System.Data.DataTable MakeNoTconTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("BU", typeof(string));

            dt.Columns.Add("項目編號", typeof(string));

            dt.Columns.Add("數量", typeof(int));


            dt.Columns.Add("價格", typeof(int));

            dt.Columns.Add("總計", typeof(int));

            dt.TableName = "dt";

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }

        private static void WriteDataTableToSheetByArray(System.Data.DataTable dataTable,
            Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range excelRange;
            int rows = dataTable.Rows.Count + 1;
            int columns = dataTable.Columns.Count;
            int rownow = 1;
            

            var data = new object[rows, columns];

            int rowcount = 0;
            for (int i = 1; i <= columns; i++)
            {
                data[rowcount, i - 1] = dataTable.Columns[i - 1].ColumnName;
            }

            rowcount += 1;
            foreach (DataRow datarow in dataTable.Rows)
            {
                for (int i = 1; i <= dataTable.Columns.Count; i++)
                {

                    // Filling the excel file 
                    data[rowcount, i - 1] = datarow[i - 1].ToString();
                    

                }
                

                rowcount += 1;
            }

            var startCell = (Range)worksheet.Cells[1, 1];
            var endCell = (Range)worksheet.Cells[rows, columns];
            var writeRange = worksheet.Range[startCell, endCell];

            //aRange.Columns.AutoFit();

            writeRange.Value2 = data;
            rowcount = 2;//第二行開始
            /*
            //KIT相同合併
            for(int i=0 ; i< dataTable.Rows.Count;i++)
            {
                try
                {
                    DataRow[] rowss = dataTable.Select("KIT = '" + dataTable.Rows[i]["KIT"].ToString() + "' and PartNo = '" + dataTable.Rows[i]["PartNo"].ToString() + "'");
                    bool combineFlag = false;
                    //處理跨行相同的情況,當前row與下一個相同料號的row,index必須相差等於一
                    for (int j = 0; j < rowss.Length - 1; j++) //最後一筆不看
                    {
                        int rowindex1 = dataTable.Rows.IndexOf(rowss[j]);//當前筆
                        int rowindex2 = dataTable.Rows.IndexOf(rowss[j + 1]);//下一筆
                        if (rowindex2 - rowindex1 == 1) //差距為1行
                        {
                            combineFlag = true;
                        }
                        else 
                        { 
                            
                            break;
                        }
                        

                    }
                    if (rowss.Length > 1 && dataTable.Rows[i]["KIT"].ToString() != "" && combineFlag == true) 
                    {
                        int j = rowss.Length - 1;
                        worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 2 + j).ToString()).MergeCells);
                    }
                    


                    rowcount += rowss.Length ;
                    i = rowss.Length > 1 ? i += rowss.Length - 1:i ;
                }
                catch (Exception ex) 
                {

                }
            }
            //項目編號相同合併
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                try
                {
                    DataRow[] rowss = dataTable.Select("項目編號 = '" + dataTable.Rows[i]["項目編號"].ToString() + "' and 庫存量 = '" + dataTable.Rows[i]["庫存量"].ToString() + "'and 待進貨量 = '" + dataTable.Rows[i]["待進貨量"].ToString() + "'" );
                    bool combineFlag = false;
                    //處理跨行相同的情況,當前row與下一個相同料號的row,index必須相差等於一
                    for (int j = 0; j < rowss.Length - 1; j++) //最後一筆不看
                    {
                        int rowindex1 = dataTable.Rows.IndexOf(rowss[j]);//當前筆
                        int rowindex2 = dataTable.Rows.IndexOf(rowss[j + 1]);//下一筆
                        if (rowindex2 - rowindex1 == 1) //差距為1行
                        {
                            combineFlag = true;
                        }
                        else
                        {

                            break;
                        }


                    }

                    if (rowss.Length > 1 && dataTable.Rows[i]["KIT"].ToString() != "" && combineFlag == true)
                    {
                        int j = rowss.Length - 1;
                        worksheet.get_Range("A" + (i + 2).ToString(), "A" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("A" + (i + 2).ToString(), "A" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 2 + j).ToString()).MergeCells);
                        worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 2 + j).ToString()).MergeCells);
                        
                    }



                    rowcount += rowss.Length;
                    i = rowss.Length > 1 ? i += rowss.Length - 1 : i;
                }
                catch (Exception ex)
                {

                }
            }*/


            //KIT相同合併
            for (int i = 0; i < dataTable.Rows.Count + 1 ; i++)
            {
                try
                {
                    int rowindex = i;

                    for (int j = 1; j < dataTable.Rows.Count; j++) 
                    {
                        string ASS= dataTable.Rows[rowindex]["項目編號"].ToString();

            //if (ASS == "P550QVN01.12002")
            //{
            //    MessageBox.Show("A");
            //}
                        int combineindex = i + j;
                        if (Convert.ToString(dataTable.Rows[rowindex]["項目編號"]) != Convert.ToString(dataTable.Rows[combineindex]["項目編號"]) && j >= 2)
                        {
                           
                            worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 2 + j).ToString()).MergeCells);
                            i = i + j - 1;
                            break;
                        }
                        else if (Convert.ToString(dataTable.Rows[rowindex]["項目編號"]) != Convert.ToString(dataTable.Rows[combineindex]["項目編號"]) && j == 1) 
                        {
                            break;
                        }
                        else if (Convert.ToString(dataTable.Rows[rowindex]["項目編號"]) == Convert.ToString(dataTable.Rows[combineindex]["項目編號"])) 
                        {
                            continue;
                        }
                       


                    }





                   
                }
                catch (Exception ex)
                {

                }
            }

            //KIT相同合併
            for (int i = 0; i < dataTable.Rows.Count + 2; i++)
            {
                try
                {
                    int rowindex = i;

                    for (int j = 1; j < dataTable.Rows.Count; j++)
                    {
                        int combineindex = i + j;
                        if (Convert.ToString(dataTable.Rows[rowindex]["KIT"]) != Convert.ToString(dataTable.Rows[combineindex]["KIT"]) && j >= 2)
                        {
                            
                            worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 1 + j).ToString()).MergeCells);
                            worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 1 + j).ToString()).MergeCells);
                            worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 1 + j).ToString()).MergeCells);
                            worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 1 + j).ToString()).MergeCells);
                            worksheet.get_Range("J" + (i + 2).ToString(), "J" + (i + 1 + j).ToString()).Merge(worksheet.get_Range("J" + (i + 2).ToString(), "J" + (i + 1 + j).ToString()).MergeCells);
                            i = i + j - 1;
                            break;
                        }
                        else if (Convert.ToString(dataTable.Rows[rowindex]["KIT"]) != Convert.ToString(dataTable.Rows[combineindex]["KIT"]) && j == 1)
                        {
                            break;
                        }
                        else if (Convert.ToString(dataTable.Rows[rowindex]["KIT"]) == Convert.ToString(dataTable.Rows[combineindex]["KIT"]))
                        {
                            continue;
                        }



                    }






                }
                catch (Exception ex)
                {

                }
            }





            writeRange.Columns.AutoFit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.aP_OPENCELLTableAdapter.FillBy(this.lC.AP_OPENCELL, textBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //if (textBox4.Text != "")
            //{
            //    MessageBox.Show("請輸入項目編號");
            //    return;
            //}
            //string PARTNO = "";
            //System.Data.DataTable G2 = GetOITM(textBox4.Text);
            //if (G2.Rows.Count == 0)
            //{
            //    MessageBox.Show("請輸入正確項目編號");
            //    return;
            //}
            //else
            //{
            //    PARTNO = G2.Rows[0][0].ToString();
            //}
            System.Data.DataTable G1 = GetOPEN(textBox2.Text,  comboBox1.Text);
            if (G1.Rows.Count > 0)
            {
                for (int i = 0; i <= G1.Rows.Count - 1; i++)
                {
                    string ITEMCODE = G1.Rows[i]["ITEMCODE"].ToString();

                    ADDOPENCELL(ITEMCODE, "", "");
                }
                MessageBox.Show("資料已匯入"+ G1.Rows.Count.ToString()+"行");
                this.aP_OPENCELLTableAdapter.FillBy(this.lC.AP_OPENCELL, textBox1.Text);
            }
            else
            {
                MessageBox.Show("無資料");
            }


            //  ADDOPENCELL
        }
        public  System.Data.DataTable GetOPEN(string U_PARTNO, string U_LOCATION)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,U_PARTNO PARTNO FROM OITM ");
            //sb.Append(" WHERE (U_TMODEL=@U_TMODEL OR 'T'+SUBSTRING(U_TMODEL,2,LEN(U_TMODEL)-1)=@U_TMODEL) AND U_VERSION=@U_VERSION AND U_LOCATION=@U_LOCATION");
            sb.Append(" WHERE U_PARTNO=@U_PARTNO  ");
            if (comboBox1.Text != "")
            {
                sb.Append("   AND U_LOCATION=@U_LOCATION");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PARTNO", U_PARTNO));

            command.Parameters.Add(new SqlParameter("@U_LOCATION", U_LOCATION));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetOITM(string ITEMCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,U_PARTNO PARTNO FROM OITM ");
            sb.Append(" WHERE ITEMCODE=@ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
           
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public void ADDOPENCELL(string OPENCELL, string KIT, string PARTNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_OPENCELL(OPENCELL,KIT,PARTNO) values(@OPENCELL,@KIT,@PARTNO)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OPENCELL", OPENCELL));
            command.Parameters.Add(new SqlParameter("@KIT", KIT));
            command.Parameters.Add(new SqlParameter("@PARTNO", PARTNO));
           

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



        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

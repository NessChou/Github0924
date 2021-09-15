using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace ACME
{
    public partial class goodsorder : Form
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        System.Data.DataTable dts = new System.Data.DataTable();
        System.Data.DataTable dtss = new System.Data.DataTable();
        System.Data.DataTable dtShow = new System.Data.DataTable();
        float Thight = 0;
        float Twidth = 0;
        float Tlength = 0;
        float Tweight = 0;
        int count = 0;
        public goodsorder()
        {

            InitializeComponent();
            init();

            cobTruckTypeinit();
            cobstartinit();
            cobendinit();
            cobTruckRegioninit();
            initDatatable();

        }
        private void init()
        {


        }
        private void cobTruckRegioninit()
        {
            cobTruckType.Items.Clear();
            cobend.Items.Clear();
            cobTruckRegion.Items.Clear();
            cobTruckRegion.Items.Add("宏高中港車車型");
            cobTruckRegion.Items.Add("宏高大陸車車型");
            cobTruckRegion.Items.Add("巨航中港車車型");
            cobTruckRegion.Items.Add("巨航大陸車車型");
            cobTruckRegion.Items.Add("英航車型");
            cobTruckRegion.Items.Add("萬泰車型");
            cobTruckRegion.Items.Add("友福車型");
            cobTruckRegion.SelectedItem = "宏高中港車車型";
            cobTruckRegion.SelectedIndex = 0;

        }
        private void cobTruckTypeinit()
        {
            cobTruckType.Items.Clear();
            cobTruckType.Text = "";
        }

        private void cobstartinit()
        {
            cobstart.Items.Clear();
            cobstart.Items.Add("");
            cobstart.Items.Add("蘇州保稅區");
            cobstart.Items.Add("廈門保稅區");
            cobstart.Items.Add("深圳保稅區");
            cobstart.Items.Add("武漢保稅區");
            cobstart.Items.Add("香港宏高");
            cobstart.Items.Add("環球櫃場");
            cobstart.Text = "";
            cobTruckType.Items.Clear();

        }
        private void cobendinit()
        {
            cobend.Items.Clear();
            cobend.Items.Add("");
            cobend.Items.Add("蘇州");
            cobend.Items.Add("廈門");
            cobend.Items.Add("深圳");
            cobend.Items.Add("香港");
            cobend.Items.Add("威海");
            cobend.Items.Add("吳江");
            cobend.Items.Add("武漢");
            cobend.Items.Add("沙頭角");
            cobend.Items.Add("聯倉");
            cobend.Items.Add("新倉");
            cobend.Text = "";
        }
        private void initDatatable()
        {
            DataGridViewColumn column;
            DataRow row;
            dt.Columns.Add("stackCheck", typeof(bool));
            dt.Columns.Add("num", typeof(string));
            dt.Columns.Add("type", typeof(string));
            dt.Columns.Add("length", typeof(float));
            dt.Columns.Add("width", typeof(float));
            dt.Columns.Add("height", typeof(float));
            dt.Columns.Add("weight", typeof(float));
            dt.Columns.Add("count", typeof(string));
            dt.Columns.Add("tag", typeof(bool));
            dt.Columns.Add("stacktag", typeof(bool));

            dts.Columns.Add("stackCheck", typeof(bool));
            dts.Columns.Add("num", typeof(string));
            dts.Columns.Add("type", typeof(string));
            dts.Columns.Add("length", typeof(string));
            dts.Columns.Add("width", typeof(string));
            dts.Columns.Add("height", typeof(string));
            dts.Columns.Add("weight", typeof(string));
            dts.Columns.Add("count", typeof(string));
            dts.Columns.Add("tag", typeof(bool));
            dts.Columns.Add("stacktag", typeof(bool));

            dtss.Columns.Add("stackCheck", typeof(bool));
            dtss.Columns.Add("num", typeof(string));
            dtss.Columns.Add("type", typeof(string));
            dtss.Columns.Add("length", typeof(string));
            dtss.Columns.Add("width", typeof(string));
            dtss.Columns.Add("height", typeof(string));
            dtss.Columns.Add("weight", typeof(string));
            dtss.Columns.Add("count", typeof(string));
            dtss.Columns.Add("tag", typeof(bool));
            dtss.Columns.Add("stacktag", typeof(bool));

            dtShow.Columns.Add("num", typeof(string));
            dtShow.Columns.Add("length", typeof(string));
            dtShow.Columns.Add("width", typeof(string));
            dtShow.Columns.Add("height", typeof(string));
            dtShow.Columns.Add("weight", typeof(string));
            dtShow.Columns.Add("GCount", typeof(string));
            dtShow.Columns.Add("tag", typeof(bool));
            dtShow.Columns.Add("stacktag", typeof(bool));
            dgvGoods.DataSource = dtShow;
            dgvGoods.Columns[2].HeaderText = "樣式";
            dgvGoods.Columns[2].Width = 60;
            dgvGoods.Columns[3].HeaderText = "板號";
            dgvGoods.Columns[3].Width = 60;
            dgvGoods.Columns[4].HeaderText = "長度";
            dgvGoods.Columns[4].Width = 60;
            dgvGoods.Columns[5].HeaderText = "寬度";
            dgvGoods.Columns[5].Width = 60;
            dgvGoods.Columns[6].HeaderText = "高度";
            dgvGoods.Columns[6].Width = 60;
            dgvGoods.Columns[7].HeaderText = "重量";
            dgvGoods.Columns[7].Width = 60;
            dgvGoods.Columns[8].HeaderText = "數量";
            dgvGoods.Columns[8].Width = 60;
            dgvGoods.Columns[10].Visible = false;


            /*
            object[] p1 = { 115, 84, 143 , 1, false};
            object[] p2 = { 115, 84, 79 , 1, false };
            object[] p3 = { 115, 84, 79 , 1, false };
            object[] p4 = { 115, 84, 79 ,1, false };
            object[] p5 = { 115, 84, 79, 1 , false };
            object[] p6 = { 80, 70, 79, 1, false };
            dt.Rows.Add(p1);
            dt.Rows.Add(p2);
            dt.Rows.Add(p3);
            dt.Rows.Add(p4);
            dt.Rows.Add(p5);
            dt.Rows.Add(p6);
            
            DataView dv = dt.DefaultView;
            dv.Sort ="length DESC";
            dt = dv.ToTable();*/



        }
        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow rowShow;
                DataRow row;
                rowShow = dtShow.NewRow();
                rowShow["length"] = "";
                rowShow["width"] = "";
                rowShow["height"] = "";
                rowShow["GCount"] = "";
                dtShow.Rows.Add(rowShow);
                dgvGoods.DataSource = dtShow;
                /*
                
                */
            }
            catch (Exception ex)
            {
                ACMELoggers.Loggers.log(ex);
            }
        }
        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "請選擇上傳檔案";
                Microsoft.Office.Interop.Excel.Application excel = null;
                Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
                Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
                //Interop params
                object oMissing = System.Reflection.Missing.Value;
                float TotalWeight = 0;
                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    /*
                    string fileName = dialog.FileName;

                    //定義OleDb======================================================
                    //1.檔案位置
                    string filepath = fileName;
                    //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
                    string ProviderName = "Microsoft.Jet.OLEDB.4.0;";
                    //3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
                    string ExtendedString = "'Excel 5.0;";
                    //4.第一行是否為標題(;結尾區隔)
                    string HDR = "No;";

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
                    myConn.Close();*/
                    Excel.Application excel1;
                    Excel.Workbooks wbs = null;
                    Excel.Workbook wb = null;
                    Excel.Sheets sheet;
                    Excel.Worksheet ws = null;
                    string path = dialog.FileName;

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
                    DataTable table = new DataTable();
                    DataView dv;
                    System.Data.DataRow rw = null;
                    table = Maketable();
                    string cellStr = null;
                    char ch = 'A';
                    string ShippingCode = "";
                    try
                    {
                        for (int i = 1; i < rowNum; i++)
                        {
                            ch = 'A';
                            rw = table.NewRow();
                            for (int j = 0; j < colNum; j++)
                            {
                                cellStr = ch.ToString() + (i + 1).ToString();
                                if (ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString() == "")
                                {
                                    continue;
                                }
                                else
                                {
                                    switch (j + 1)
                                    {
                                        case 1://麥頭
                                            rw["type"] = ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString();
                                            break;
                                        case 2://板號
                                            rw["num"] = ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString();
                                            break;
                                        case 4://長
                                            rw["length"] = float.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());
                                            break;
                                        case 5://寬
                                            rw["width"] = float.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());
                                            break;
                                        case 6://高
                                            rw["height"] = float.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());
                                            break;
                                        case 7://數量
                                            string cnt = ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString();
                                            rw["GCount"] = cnt.Contains('.') ? cnt.Split('.')[0] : int.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());

                                            break;
                                        case 3://重量
                                            rw["weight"] = float.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());
                                            TotalWeight += float.Parse(ws.UsedRange.Cells.get_Range(cellStr, miss).Text.ToString());

                                            break;
                                    }

                                }
                                ch++;
                            }
                            if (rw[1].ToString() != "")
                            {
                                table.Rows.Add(rw);
                            }

                        }
                    }
                    catch (Exception ex)
                    {

                    }


                    DataRow[] rows = table.Select("type is not null");
                    System.Data.DataTable dt = null;
                    System.Data.DataRow dr = null;
                    int line = 0;
                    string GCount = "";
                    dt = Maketable();
                    dts = MakeStacktable();
                    dtss = MakeStacktable();
                    dgvGoods.DataSource = table;

                    /*

                    foreach (DataRow row in rows)
                    {
                        if (row[2].ToString() == "板號")
                        {
                            continue;
                        }
                        line++;
                        if (row[2].ToString().Contains("-"))
                        {
                            string[] num = new string[2];
                            num = row[2].ToString().Split('-');
                            GCount = (int.Parse(num[1]) - int.Parse(num[0]) + 1).ToString();

                        }
                        else
                        {
                            GCount = "1";
                        }

                        dr = dt.NewRow();
                        dr["type"] = row[0].ToString();
                        dr["num"] = row[1].ToString();
                        dr["weight"] = row[2].ToString();
                        dr["length"] = row[3].ToString();
                        dr["width"] = row[4].ToString();
                        dr["height"] = row[5].ToString();
                        dr["GCount"] = row[6].ToString();
                        dt.Rows.Add(dr);

                        for (int i = 0; i < int.Parse(row[6].ToString()); i++)
                        {
                            TotalWeight += float.Parse(row[2].ToString());
                        }
                    }
                   */

                }
                labelweight.Text = "總重:" + TotalWeight;
            }
            catch (InvalidCastException ex)
            {

            }
            catch (Exception ex)
            {

                ACMELoggers.Loggers.log(ex);
            }
        }
        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;
            string year = DateTime.Now.ToString().Split(' ')[0].Split('/')[0];
            string month = DateTime.Now.ToString().Split(' ')[0].Split('/')[1].Length == 2 ? DateTime.Now.ToString().Split(' ')[0].Split('/')[1] : "0" + DateTime.Now.ToString().Split(' ')[0].Split('/')[1];
            string date = DateTime.Now.ToString().Split(' ')[0].Split('/')[2].Length == 2 ? DateTime.Now.ToString().Split(' ')[0].Split('/')[2] : "0" + DateTime.Now.ToString().Split(' ')[0].Split('/')[2];
            string pathFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + year + month + date + "貨物擺放";//日期桌面路徑
            // 開啟一個新的應用程式
            excelApp = new Excel.Application();

            // 讓Excel文件可見
            excelApp.Visible = true;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);

            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();



            try
            {
                // 引用第一個工作表
                wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                // 命名工作表的名稱
                wSheet.Name = "sheet1";

                // 設定工作表焦點
                wSheet.Activate();
                for (int i = 0; i < dgvSort.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvSort.Columns.Count; j++)
                    {
                        if (dgvSort.Rows[i].Cells[j].Value != null)
                        {
                            wSheet.Cells[i + 1, j + 1] = dgvSort.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
                wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
            }
            catch (Exception ex)
            {
                ACMELoggers.Loggers.log(ex);
            }
            finally
            {
                //關閉活頁簿
                wBook.Close(false, Type.Missing, Type.Missing);

                //關閉Excel
                excelApp.Quit();

                //釋放Excel資源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                wBook = null;
                wSheet = null;
                wRange = null;
                excelApp = null;
                GC.Collect();

                Console.Read();
                System.Diagnostics.Process.Start(pathFile + ".xlsx");
            }

        }
        private void btnStack_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvGoods.Rows)
            {
                if (row.Cells["type"].Value != null) row.Cells["stackCheck"].Value = true;
            }
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("是否要刪除？", "資訊", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    for (int i = dgvGoods.Rows.Count - 1; i >= 0; i--)
                    {
                        string value = dgvGoods.Rows[i].Cells["DeleteCheck"].EditedFormattedValue.ToString();
                        if (value == "True")
                        {
                            DataGridViewRow dgvr = dgvGoods.Rows[i];
                            this.dgvGoods.Rows.Remove(dgvr);
                        }
                    }
                }
                dgvGoods.Update();
            }
            catch (Exception ex)
            {
                ACMELoggers.Loggers.log(ex);
            }
        }
        private System.Data.DataTable Order1(float Tlength, float Twidth, int type) //貨車長，貨車寬，0會提示長寬不夠 1直接跳過
        {
            //左直放右橫放，一個大配一個小
            System.Data.DataTable Ordertable = new System.Data.DataTable();
            Ordertable.Columns.Add("left", typeof(string));
            Ordertable.Columns.Add("right", typeof(string));
            float Tlength_left1 = Tlength;//左邊長
            float Tlength_left2 = Tlength;//右邊長
            float twidth_left = Twidth;
            DataRow row;


            for (int i = 0; i < dts.Rows.Count; i++)
            {
                if ((bool)dts.Rows[i]["tag"] != true)
                {
                    //有堆疊的當項貨物長寬
                    float gwidth = float.Parse(dts.Rows[i]["width"].ToString());//當項貨物寬
                    float glength = float.Parse(dts.Rows[i]["length"].ToString().Split('|')[0]);//當項貨物長
                    string info = dts.Rows[i]["length"].ToString().Split('|')[1];


                    if (i % 2 == 0 && Tlength_left1 > gwidth && twidth_left > glength)
                    {
                        row = Ordertable.NewRow();
                        //基數行大的先放左邊
                        row["left"] = gwidth + "," + glength + "|" + info;
                        Tlength_left1 = Tlength_left1 - gwidth;//左邊剩下的長
                        twidth_left = Twidth - gwidth;
                        Ordertable.Rows.Add(row);
                        dts.Rows[i]["tag"] = true;
                        float swidth = 0;
                        float slength = 0;
                        for (int j = 1; j < dts.Rows.Count - i; j++)
                        {
                            swidth = float.Parse(dts.Rows[dts.Rows.Count - j]["width"].ToString());//搭配的較小貨物寬
                            slength = float.Parse(dts.Rows[dts.Rows.Count - j]["length"].ToString().Split('|')[0]);//搭配的較小貨物長
                            string sinfo = dts.Rows[dts.Rows.Count - j]["length"].ToString().Split('|')[1];
                            if (twidth_left > slength && (bool)dts.Rows[dts.Rows.Count - j]["tag"] == false)

                            {
                                //基數行右邊直放
                                Ordertable.Rows[i]["right"] = slength + "," + swidth + "|" + sinfo;
                                Tlength_left2 = Tlength_left2 - slength;//右邊剩下的長
                                dts.Rows[dts.Rows.Count - j]["tag"] = true;
                                break;
                            }
                        }




                        twidth_left = Twidth;//下一排用正常寬計算

                    }
                    else if (i % 2 == 1 && Tlength_left2 > gwidth && twidth_left > glength)
                    {
                        row = Ordertable.NewRow();
                        //偶數行大的先放右邊
                        row["right"] = gwidth + "," + glength + "|" + info;
                        Tlength_left2 = Tlength_left2 - gwidth;//右邊剩下的長
                        twidth_left = twidth_left - glength;
                        Ordertable.Rows.Add(row);
                        dts.Rows[i]["tag"] = true;
                        float swidth = 0;
                        float slength = 0;

                        for (int j = 1; j < dts.Rows.Count - i; j++)
                        {
                            swidth = float.Parse(dts.Rows[dts.Rows.Count - j]["width"].ToString());
                            slength = float.Parse(dts.Rows[dts.Rows.Count - j]["length"].ToString().Split('|')[0]);
                            string sinfo = dts.Rows[dts.Rows.Count - j]["length"].ToString().Split('|')[1];
                            if (twidth_left > swidth && (bool)dts.Rows[dts.Rows.Count - j]["tag"] == false)
                            {
                                //偶數行左邊直放
                                Ordertable.Rows[i]["left"] = slength + "," + swidth + "|" + sinfo;
                                Tlength_left1 = Tlength_left1 - slength;//左邊剩下的長
                                dts.Rows[dts.Rows.Count - j]["tag"] = true;
                                break;
                            }
                        }

                        twidth_left = Twidth;//下一排用正常寬計算
                    }
                    else
                    {
                        DataTable table = null;
                        if ((Tlength_left1 - glength < 0 || Tlength_left1 - glength < 0) && type == 0)
                        {
                            MessageBox.Show("排序法一車子長度不足");
                        }
                        if (twidth_left - gwidth < 0 && type == 0)
                        {
                            MessageBox.Show("排序法一車子寬度不足");
                        }
                        return table;
                    }
                }
            }
            return Ordertable;
        }

        private System.Data.DataTable Order2(float Tlength, float Twidth, int type)
        {
            //左右都直放，由大放到小
            System.Data.DataTable Ordertable = new System.Data.DataTable();
            DataRow row;
            float Tlength_left = Tlength;
            float Twidth_left = Twidth;
            int X = 0;
            int Y = 0;


            for (int i = 0; i < dts.Rows.Count; i++)
            {
                float gwidth = 0;//當項貨物寬
                float glength = 0;//當項貨物長
                for (int j = 0; j < dts.Rows.Count; j++)
                {
                    gwidth = float.Parse(dts.Rows[j]["width"].ToString());
                    glength = float.Parse(dts.Rows[j]["length"].ToString().Split('|')[0]);
                    string info = dts.Rows[j]["length"].ToString().Split('|')[1];
                    if ((bool)dts.Rows[j]["tag"] == false && ((Tlength_left > glength && Y == 0) || Y == 1) && Twidth_left > gwidth)
                    {
                        DataColumnCollection columns = Ordertable.Columns;
                        float swidth = float.Parse(dts.Rows[j]["width"].ToString());
                        float slength = float.Parse(dts.Rows[j]["length"].ToString().Split('|')[0]);
                        string sinfo = dts.Rows[j]["length"].ToString().Split('|')[1];
                        if (!columns.Contains(j.ToString())) Ordertable.Columns.Add(Ordertable.Columns.Count.ToString(), typeof(string));
                        dts.Rows[j]["tag"] = true;
                        if (Y == 0)
                        {
                            Tlength_left = Tlength_left - glength;
                            row = Ordertable.NewRow();
                            row[Y.ToString()] = gwidth + "," + glength + "|" + info;
                            Ordertable.Rows.Add(row);
                        }
                        else
                        {
                            Ordertable.Rows[X][Y.ToString()] = swidth + "," + slength + "|" + sinfo;
                        }
                        Twidth_left = Twidth_left - gwidth;
                        Y++;
                    }
                    else if (Y != 0 && Tlength_left < glength)
                    {
                        DataTable table = null;
                        if (type == 0) MessageBox.Show("排序法二車子長度不足");
                        return table;
                    }
                }
                Y = 0;
                X++;
                Twidth_left = Twidth;
            }

            return Ordertable;
        }
        private System.Data.DataTable Order3(float Tlength, float Twidth, int type)
        {
            //橫放大放到小
            System.Data.DataTable Ordertable = new System.Data.DataTable();
            DataRow row;
            float Tlength_left1 = Tlength;//左邊長
            float Tlength_left2 = Tlength;//右邊長
            float Twidth_left = Twidth;
            int X = 0;
            int Y = 0;
            string[,] array = new string[dts.Rows.Count, dts.Rows.Count];
            float gwidth = 0;//當項貨物寬
            float glength = 0;//當項貨物長


            for (int i = 0; i < dts.Rows.Count; i++)
            {
                for (int j = 0; j < dts.Rows.Count; j++)
                {
                    gwidth = float.Parse(dts.Rows[j]["width"].ToString());
                    glength = float.Parse(dts.Rows[j]["length"].ToString().Split('|')[0]);
                    string info = dts.Rows[j]["length"].ToString().Split('|')[1];
                    if ((bool)dts.Rows[j]["tag"] == false && Twidth_left >= glength && ((Y == 0 && Tlength_left1 >= gwidth) || (Y == 1 && Twidth_left >= glength && Tlength_left2 >= glength)))//若條件式設Twidth_left >= glength 不在左邊的有可能因為Twidth_left被扣放不進去
                    {
                        DataColumnCollection columns = Ordertable.Columns;

                        if (!columns.Contains(Y.ToString())) Ordertable.Columns.Add(Y.ToString(), typeof(string));
                        dts.Rows[j]["tag"] = true;
                        if (Y == 0)
                        {
                            Twidth_left = Twidth;
                            Tlength_left1 = Tlength_left1 - gwidth;
                            row = Ordertable.NewRow();
                            row[Y.ToString()] = glength + "," + gwidth + "|" + info;
                            Ordertable.Rows.Add(row);
                        }
                        else
                        {
                            Ordertable.Rows[X][Y.ToString()] = glength + "," + gwidth + "|" + info;
                            Tlength_left2 = Tlength_left2 - gwidth;
                        }
                        Twidth_left = Twidth_left - glength;
                        Y++;
                    }
                    else if ((bool)dts.Rows[j]["tag"] == false && ((Y == 1 && Tlength_left2 <= gwidth) || (Y == 0 && Tlength_left1 <= gwidth)))
                    {
                        DataTable table = null;
                        if (type == 0) MessageBox.Show("排序法三車子長度不足");
                        return table;
                    }
                    /*else if(glength > Twidth_left || Tlength_left < gwidth )
                    { 
                        Y = 0;
                    }*/
                }
                Twidth_left = Twidth;
                Y = 0;
                X++;
            }

            return Ordertable;
        }
        private System.Data.DataTable arrange(System.Data.DataTable OldTable)
        {

            System.Data.DataTable NewTable = new System.Data.DataTable();
            DataRow row;
            string length = "";
            string width = "";
            string lengthB = "";
            string widthB = "";
            NewTable.Columns.Add("0", typeof(string));
            NewTable.Columns.Add("1", typeof(string));
            NewTable.Columns.Add("2", typeof(string));
            NewTable.Columns.Add("3", typeof(string));
            try
            {
                for (int i = 0; i < OldTable.Columns.Count; i++)
                {
                    for (int j = 0; j < OldTable.Rows.Count; j++)
                    {
                        if (OldTable.Rows[j][i].ToString().Contains(',') && OldTable.Rows[j][i].ToString().Contains('|'))
                        {
                            width = OldTable.Rows[j][i].ToString().Split('|')[0].Split(',')[0];
                            length = OldTable.Rows[j][i].ToString().Split('|')[0].Split(',')[1];
                            string info = OldTable.Rows[j][i].ToString().Split('|')[1];
                            if (i % 2 == 0)
                            {
                                //左側貨物直接新增datarow
                                row = NewTable.NewRow();
                                row["1"] = width;
                                NewTable.Rows.Add(row);
                                row = NewTable.NewRow();
                                row["0"] = length;
                                row["1"] = info;
                                NewTable.Rows.Add(row);
                            }
                            else
                            {
                                //右側貨物直接填入table
                                if (NewTable.Rows.Count >= 2 * j + 1)
                                {
                                    NewTable.Rows[2 * j][2] = width;
                                    NewTable.Rows[2 * j + 1][3] = length;
                                    NewTable.Rows[2 * j + 1][2] = info;
                                }
                                else
                                {
                                    //有可能左邊的貨物被取去做堆疊剩下右邊的,直接填入會造成exception
                                    row = NewTable.NewRow();
                                    row["2"] = width;
                                    NewTable.Rows.Add(row);
                                    row = NewTable.NewRow();
                                    row["3"] = length;
                                    row["2"] = lengthB;
                                    NewTable.Rows.Add(row);
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ACMELoggers.Loggers.log(ex);
            }
            return NewTable;
        }

        private void btnOrder_Click(object sender, EventArgs e)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            dt.Clear();
            dts.Clear();
            System.Windows.Forms.Button btn = sender as System.Windows.Forms.Button;
            int numCount = 0;
            if (cobTruckType.Text == "")
            {
                MessageBox.Show("請選擇車子型號");
            }
            else
            {
                if (float.Parse(labelweight.Text.Split(':')[1]) + 200 > Tweight && float.Parse(labelweight.Text.Split(':')[1]) < Tweight)
                {
                    string[] items = new string[cobTruckType.Items.Count];
                    string nextWeight = "";
                    for (int i = 0; i < cobTruckType.Items.Count; i++)
                    {
                        items[i] = cobTruckType.Items[i].ToString();
                        if (i == 0) continue;
                        if (items[i - 1].ToString() == cobTruckType.Text)
                        {
                            //前一個層級的內容等於所選等級
                            nextWeight = items[i].Split('(')[0];
                            break;
                        }
                    }
                    DialogResult Result = MessageBox.Show("建議:" + nextWeight + "因為重量要多抓200KG以免卡關", "重量提醒", MessageBoxButtons.OKCancel);
                    if (Result == DialogResult.OK)
                    {
                        stacksort();//堆疊排序

                        if (Tlength != 0 && Twidth != 0)
                        {
                            switch (btn.Text)
                            {
                                case "排序法一":
                                    table = Order1(Tlength, Twidth, 0);
                                    break;
                                case "排序法二":
                                    table = Order2(Tlength, Twidth, 0);
                                    break;
                                case "排序法三":
                                    table = Order3(Tlength, Twidth, 0);
                                    break;
                                case "排序":
                                    table = Order1(Tlength, Twidth, 1);
                                    if (table == null)
                                    {

                                        stacksort();
                                        table = Order2(Tlength, Twidth, 1);
                                        if (table == null)
                                        {
                                            stacksort();
                                            table = Order3(Tlength, Twidth, 1);
                                        }
                                    }
                                    break;
                            }
                        }
                        else
                        {
                            MessageBox.Show("請選擇車型");
                        }

                        if (table != null)
                        {
                            table = arrange(table);
                            dgvSort.DataSource = table;

                        }
                        else
                        {
                            MessageBox.Show("車子太小無法排序");
                        }
                    }
                }
                else
                {
                    stacksort();

                    if (Tlength != 0 && Twidth != 0)
                    {
                        switch (btn.Text)
                        {
                            case "排序法一":
                                table = Order1(Tlength, Twidth, 0);
                                break;
                            case "排序法二":
                                table = Order2(Tlength, Twidth, 0);
                                break;
                            case "排序法三":
                                table = Order3(Tlength, Twidth, 0);
                                break;
                            case "排序":
                                table = Order1(Tlength, Twidth, 1);
                                if (table == null)
                                {

                                    stacksort();
                                    table = Order2(Tlength, Twidth, 1);
                                    if (table == null)
                                    {
                                        stacksort();
                                        table = Order3(Tlength, Twidth, 1);
                                    }
                                }
                                break;
                        }
                    }
                    else
                    {
                        MessageBox.Show("請選擇車型");
                    }

                    if (table != null)
                    {
                        table = arrange(table);
                        dgvSort.DataSource = table;

                    }
                    else
                    {
                        MessageBox.Show("車子太小無法排序");
                    }
                }

            }
        }
        private void stacksort()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            dt.Clear();
            dts.Clear();
            int numCount;
            float TotalWeight = 0;
            DataRow row;
            TotalWeight = float.Parse(labelweight.Text.Split(':')[1]);
            if (TotalWeight > Tweight)
            {
                MessageBox.Show("超出車子載重量");
                return;
            }
            for (int i = 0; i < dgvGoods.Rows.Count; i++)
            {

                if (dgvGoods.Rows[i].Cells["GCount"].Value != null)
                {
                    numCount = dgvGoods.Rows[i].Cells["num"].Value.ToString().Contains('-') ? int.Parse(dgvGoods.Rows[i].Cells["num"].Value.ToString().Split('-')[0]) : int.Parse(dgvGoods.Rows[i].Cells["num"].Value.ToString());
                    for (int j = 0; j < int.Parse(dgvGoods.Rows[i].Cells["GCount"].Value.ToString()); j++)
                    {
                        row = dt.NewRow();
                        row["num"] = numCount;
                        numCount++;
                        row["type"] = dgvGoods.Rows[i].Cells["type"].Value.ToString();
                        row["length"] = float.Parse(dgvGoods.Rows[i].Cells["length"].Value.ToString());
                        row["width"] = float.Parse(dgvGoods.Rows[i].Cells["width"].Value.ToString());
                        row["height"] = float.Parse(dgvGoods.Rows[i].Cells["height"].Value.ToString());
                        row["weight"] = float.Parse(dgvGoods.Rows[i].Cells["weight"].Value.ToString());
                        row["stackCheck"] = dgvGoods.Rows[i].Cells["stackCheck"].Value;
                        row["stacktag"] = false;
                        row["tag"] = false;


                        dt.Rows.Add(row);
                        if (int.Parse(dgvGoods.Rows[i].Cells["height"].Value.ToString()) > Thight)
                        {
                            MessageBox.Show(dgvGoods.Rows[i].Cells["num"].Value.ToString() + "物品高度超出車子高度");
                            break;
                        }


                    }
                }
            }

            DataView dv = dt.DefaultView;
            dv.Sort = "length DESC,width DESC";
            dt = dv.ToTable();


            //堆疊判斷,並把排好的放進dts
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["stackCheck"].Equals(true) && dt.Rows[i]["stacktag"].Equals(false))
                {
                    int lengthA = int.Parse(dt.Rows[i]["length"].ToString());
                    int widthA = int.Parse(dt.Rows[i]["width"].ToString());
                    int hightA = int.Parse(dt.Rows[i]["height"].ToString());

                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        if (dt.Rows[j]["length"].ToString().Contains('|'))
                        {
                            //已有堆疊
                            continue;
                        }
                        else
                        {
                            int lengthB = int.Parse(dt.Rows[j]["length"].ToString());
                            int widthB = int.Parse(dt.Rows[j]["width"].ToString());
                            int hightB = int.Parse(dt.Rows[j]["height"].ToString());
                            bool stacktag = bool.Parse(dt.Rows[j]["stacktag"].ToString());//上層還未堆疊東西

                            if (lengthA >= lengthB && widthA >= widthB && hightA + hightB < Thight && stacktag == false)
                            {
                                //有堆疊，在長寬後面加上 | 與堆疊上的長寬
                                DataRow srow;
                                srow = dts.NewRow();
                                srow["num"] = true;
                                srow["stacktag"] = true;
                                dt.Rows[i]["stacktag"] = true;
                                dt.Rows[j]["stacktag"] = true;
                                if (dt.Rows[i]["type"].ToString() == dt.Rows[j]["type"].ToString())
                                {
                                    srow["length"] = lengthA + "|" + dt.Rows[i]["type"] + " 第" + dt.Rows[i]["num"] + "+" + dt.Rows[j]["num"] + "堆";
                                }
                                else
                                {
                                    srow["length"] = lengthA + "|" + dt.Rows[i]["type"] + " 第" + dt.Rows[i]["num"] + "+" + dt.Rows[j]["type"] + "第" + dt.Rows[j]["num"] + "堆";
                                }

                                srow["width"] = widthA;
                                srow["height"] = dt.Rows[i]["height"];
                                srow["weight"] = dt.Rows[i]["weight"];
                                srow["tag"] = dt.Rows[i]["tag"];
                                dts.Rows.Add(srow);
                                break;
                            }
                        }
                    }
                    if (dt.Rows[i]["stacktag"].Equals(false))
                    {
                        //找不到可堆疊的貨物
                        DataRow srow;
                        srow = dts.NewRow();
                        srow["num"] = true;
                        srow["stacktag"] = true;
                        srow["length"] = lengthA + "|" + dt.Rows[i]["type"] + " 第" + dt.Rows[i]["num"];
                        srow["width"] = widthA;
                        srow["height"] = dt.Rows[i]["height"];
                        srow["weight"] = dt.Rows[i]["weight"];
                        srow["tag"] = dt.Rows[i]["tag"];
                        dts.Rows.Add(srow);
                    }
                }
                else if ((dt.Rows[i]["stackCheck"].Equals(false) || dt.Rows[i]["stackCheck"].ToString() == "") && dt.Rows[i]["stacktag"].Equals(false))
                {
                    //不可堆疊直接填入dts
                    DataRow srow;
                    srow = dts.NewRow();
                    srow["num"] = true;
                    srow["stacktag"] = true;
                    srow["length"] = dt.Rows[i]["length"] + "|" + dt.Rows[i]["type"] + " 第" + dt.Rows[i]["num"];
                    srow["width"] = dt.Rows[i]["width"];
                    srow["height"] = dt.Rows[i]["height"];
                    srow["weight"] = dt.Rows[i]["weight"];
                    srow["tag"] = dt.Rows[i]["tag"];
                    dts.Rows.Add(srow);
                }
            }
        }
        private void cobTruckType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] volume = new string[3];
            int leftbracket = cobTruckType.SelectedItem.ToString().IndexOf('(');//找字串中左括號位置
            int rightbrabracket = cobTruckType.SelectedItem.ToString().IndexOf(')'); ;//找字串中右括號位置
            string s = cobTruckType.SelectedItem.ToString();
            volume = s.Substring(leftbracket + 1, rightbrabracket - leftbracket - 1).Split('*');
            Thight = float.Parse(volume[2]);
            Twidth = float.Parse(volume[1]);
            Tlength = float.Parse(volume[0]);
            Tweight = float.Parse(s.Substring(rightbrabracket + 2, s.Length - rightbrabracket - 4));

        }
        private void cobTruckRegion_SelectedIndexChanged(object sender, EventArgs e)
        {
            cobTruckTypeinit();
            cobstartinit();
            cobendinit();
            switch (cobTruckRegion.Text)
            {
                case "宏高中港車車型":
                    cobTruckType.Items.Add("5T(520*210*220),3500kg");
                    cobTruckType.Items.Add("8T(750*230*250),4500kg");
                    cobTruckType.Items.Add("10T(750*230*250),6000kg");
                    cobTruckType.Items.Add("40HQ(1200*230*256),20000kg");
                    cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                    break;

                case "宏高大陸車車型":
                    break;
                case "巨航車型":
                    cobTruckTypeinit();
                    cobTruckType.Items.Add("3T(420*180*190),3000kg");
                    cobTruckType.Items.Add("5T(500*200*200),5000kg");
                    cobTruckType.Items.Add("8T(620*210*220),8000kg");
                    cobTruckType.Items.Add("10T(720*240*240),8000kg");
                    cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                    cobTruckType.Items.Add("45HQ(1350*235*2568),20000kg");
                    break;
                case "英航車型":
                    break;
                case "萬泰車型":
                    cobTruckType.Items.Add("3T(420*210*200),3000kg");
                    cobTruckType.Items.Add("5T(590*228*240),5000kg");
                    cobTruckType.Items.Add("8T(750*240*240),8000kg");
                    cobTruckType.Items.Add("10T(960*240*250),10000kg");
                    cobTruckType.Items.Add("40HQ(1200*234*268),20000kg");
                    cobTruckType.Items.Add("45HQ(1300*235*268),20000kg");
                    cobTruckType.Items.Add("53HQ(1700*265*290),20000kg");
                    break;
                case "友福車型":
                    cobTruckType.Items.Add("0.5T(200*130*130),500kg");
                    cobTruckType.Items.Add("1.5T(295*150*150),1500kg");
                    cobTruckType.Items.Add("6.8T(450*205*180),3000g");
                    cobTruckType.Items.Add("8.8T(490*210*200),3500kg");
                    cobTruckType.Items.Add("10.5T(595*215*210),5000kg");
                    cobTruckType.Items.Add("15T(690*235*220),7000kg");
                    break;
            }


            /*
                        switch (cobTruckRegion.Text)
                        {
                            case "宏高車車型":
                                cobTruckType.Items.Add("5T(520*210*220),3500kg");
                                cobTruckType.Items.Add("8T(750*230*250),4500kg");
                                cobTruckType.Items.Add("10T(750*230*250),6000kg");
                                cobTruckType.Items.Add("40HQ(1200*230*256),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                                break;

                            case "大陸車車型":
                                cobRegion.Items.Add("蘇州");
                                cobRegion.Items.Add("廈門");
                                cobRegion.Items.Add("深圳");
                                break;

                            case "巨航車型":
                                cobTruckType.Items.Add("3T(420*180*190),3000kg");
                                cobTruckType.Items.Add("5T(500*200*200),5000kg");
                                cobTruckType.Items.Add("8T(620*210*220),8000kg");
                                cobTruckType.Items.Add("10T(720*240*240),6000kg");
                                cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*235*2568),20000kg");
                                break;
                            case "英航車型":
                                cobTruckType.Items.Add("5T(520*210*220),3500kg");
                                cobTruckType.Items.Add("8T(750*230*250),4500kg");
                                cobTruckType.Items.Add("10T(750*230*250),6000kg");
                                cobTruckType.Items.Add("40HQ(1200*230*256),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                                break;
                            case "萬泰車型":
                                cobTruckType.Items.Add("3T(420*210*200),3000kg");
                                cobTruckType.Items.Add("5T(590*228*240),5000kg");
                                cobTruckType.Items.Add("8T(750*240*240),8000kg");
                                cobTruckType.Items.Add("10T(960*240*250),10000kg");
                                cobTruckType.Items.Add("40HQ(1200*234*268),20000kg");
                                cobTruckType.Items.Add("45HQ(1300*235*268),20000kg");
                                cobTruckType.Items.Add("53HQ(1700*265*290),20000kg");
                                break;
                            case "友福車型":
                                cobTruckType.Items.Add("0.5T(200*130*130),500kg");
                                cobTruckType.Items.Add("1.5T(295*150*150),1500kg");
                                cobTruckType.Items.Add("6.8T(450*205*180),3000g");
                                cobTruckType.Items.Add("8.8T(490*210*200),3500kg");
                                cobTruckType.Items.Add("10.5T(595*215*210),5000kg");
                                cobTruckType.Items.Add("15T(690*235*220),7000kg");
                                break;
                        }*/
        }
        private void cobstart_SelectedIndexChanged(object sender, EventArgs e)
        {
            cobTruckTypeinit();
            switch (cobTruckRegion.Text)
            {
                case "宏高大陸車車型":
                    switch (cobstart.Text)
                    {
                        case "蘇州保稅區":
                            cobTruckType.Items.Add("3T(620*210*210),3000kg");
                            cobTruckType.Items.Add("5T(720*235*240),5000kg");
                            cobTruckType.Items.Add("8T(960*235*240),8000kg");
                            cobTruckType.Items.Add("10T(960*235*240),8000kg");
                            cobTruckType.Items.Add("40HQ(1200*235*256),20000kg");
                            cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                            break;
                        case "廈門保稅區":
                            cobTruckType.Items.Add("5T(600*235*240),5000kg");
                            cobTruckType.Items.Add("8T(960*235*240),8000kg");
                            cobTruckType.Items.Add("10T(960*235*240),8000kg");
                            cobTruckType.Items.Add("40HQ(1200*235*256),20000kg");
                            cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                            break;
                        case "深圳保稅區":
                            cobTruckType.Items.Add("5T(520*210*210),5000kg");
                            cobTruckType.Items.Add("8T(680*235*240),8000kg");
                            cobTruckType.Items.Add("10T(680*235*240),8000kg");
                            cobTruckType.Items.Add("40HQ(1200*230*230),20000kg");
                            cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");
                            break;
                    }
                    break;
                case "巨航大陸車車型":
                    if (cobstart.Text == "深圳保稅區")
                    {
                        switch (cobend.Text)
                        {
                            case "武漢":
                                cobTruckType.Items.Add("5T(540*210*210),5000kg");
                                cobTruckType.Items.Add("8T(660*230*230),8000kg");
                                cobTruckType.Items.Add("10T(760*230*235),8000kg");
                                break;
                            case "沙頭角":
                                cobTruckType.Items.Add("3T(420*180*190),3000kg");
                                cobTruckType.Items.Add("5T(540*210*210),5000kg");
                                cobTruckType.Items.Add("8T(660*230*230),8000kg");
                                cobTruckType.Items.Add("10T(760*230*235),10000kg");
                                cobTruckType.Items.Add("40H(1200*235*238),20000kg");
                                cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*235*268),20000kg");
                                break;
                            default:

                                break;

                        }
                    }
                    break;
                case "英航車型":
                    switch (cobstart.Text)
                    {
                        case "深圳保稅區":
                            switch (cobend.Text)
                            {
                                case "":
                                    cobTruckType.Items.Add("請先選擇終點");
                                    break;
                                case "威海大宇":
                                    cobTruckType.Items.Add("200kg(115*85*100),200kg");
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;
                                case "深圳宏普欣":
                                    cobTruckType.Items.Add("請選擇別的起訖地點");
                                    break;

                            }
                            break;
                        case "蘇州保稅區":
                            switch (cobend.Text)
                            {
                                case "":
                                    cobTruckType.Items.Add("請先選擇終點");
                                    break;
                                case "威海大宇":
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;
                                case "深圳宏普欣":
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;

                            }
                            break;
                    }
                    break;
                case "宏高中港車車型":
                    if (cobstart.Text == "" || cobTruckType.Items.Count == 0)
                    {
                        cobTruckType.Items.Add("5T(520*210*220),3500kg");
                        cobTruckType.Items.Add("8T(750*230*250),4500kg");
                        cobTruckType.Items.Add("10T(750*230*250),6000kg");
                        cobTruckType.Items.Add("40HQ(1200*230*256),20000kg");
                        cobTruckType.Items.Add("45HQ(1350*240*256),20000kg");

                    }
                    break;
                case "巨航車型":
                    if (cobstart.Text == "" || cobTruckType.Items.Count == 0)
                    {
                        cobTruckType.Items.Clear();
                        cobTruckType.Items.Add("3T(420*180*190),3000kg");
                        cobTruckType.Items.Add("5T(500*200*200),5000kg");
                        cobTruckType.Items.Add("8T(620*210*220),8000kg");
                        cobTruckType.Items.Add("10T(720*240*240),8000kg");
                        cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                        cobTruckType.Items.Add("45HQ(1350*235*2568),20000kg");
                    }
                    break;
                case "有福車型":
                    if (cobstart.Text == "" || cobTruckType.Items.Count == 0)
                    {
                        cobTruckType.Items.Clear();
                        cobTruckType.Items.Add("3T(420*180*190),3000kg");
                        cobTruckType.Items.Add("5T(500*200*200),5000kg");
                        cobTruckType.Items.Add("8T(620*210*220),8000kg");
                        cobTruckType.Items.Add("10T(720*240*240),8000kg");
                        cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                        cobTruckType.Items.Add("45HQ(1350*235*2568),20000kg");
                    }
                    break;
                case "萬泰車型":
                    if (cobstart.Text == "" || cobTruckType.Items.Count == 0)
                    {
                        cobTruckType.Items.Add("3T(420*210*200),3000kg");
                        cobTruckType.Items.Add("5T(590*228*240),5000kg");
                        cobTruckType.Items.Add("8T(750*240*240),8000kg");
                        cobTruckType.Items.Add("10T(960*240*250),10000kg");
                        cobTruckType.Items.Add("40HQ(1200*234*268),20000kg");
                        cobTruckType.Items.Add("45HQ(1300*235*268),20000kg");
                        cobTruckType.Items.Add("53HQ(1700*265*290),20000kg");
                    }
                    break;

            }
            cobTruckType.Text = "";
        }
        private void cobend_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cobTruckRegion.Text)
            {
                case "英航車型":
                    switch (cobstart.Text)
                    {
                        case "":
                            cobTruckType.Items.Add("請選擇起點");
                            break;
                        case "深圳保稅區":
                            switch (cobend.Text)
                            {
                                case "":
                                    cobTruckType.Items.Add("請先選擇終點");
                                    break;
                                case "威海大宇":
                                    cobTruckType.Items.Add("200kg(115*85*100),200kg");
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;
                                case "深圳宏普欣":
                                    cobTruckType.Items.Add("請選擇別的起訖地點");
                                    break;

                            }
                            break;
                        case "蘇州保稅區":
                            switch (cobend.Text)
                            {
                                case "":
                                    cobTruckType.Items.Add("請先選擇終點");
                                    break;
                                case "威海大宇":
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;
                                case "深圳宏普欣":
                                    cobTruckType.Items.Add("20GP(591*234*234),7000kg");
                                    cobTruckType.Items.Add("9.6m(958*240*244),10000kg");
                                    cobTruckType.Items.Add("40HQ(1180*234*268),20000kg");
                                    cobTruckType.Items.Add("45HQ(1330*234*268),20000kg");
                                    cobTruckType.Items.Add("48HQ(1440*234*268),22000kg");
                                    cobTruckType.Items.Add("16m(1595*246*275),25000kg");
                                    break;

                            }
                            break;

                    }
                    break;
                case "巨航大陸車車型":
                    if (cobstart.Text == "深圳保稅區")
                    {
                        switch (cobend.Text)
                        {
                            case "武漢":
                                cobTruckType.Items.Add("5T(540*210*210),5000kg");
                                cobTruckType.Items.Add("8T(660*230*230),8000kg");
                                cobTruckType.Items.Add("10T(760*230*235),8000kg");
                                break;
                            case "沙頭角":
                                cobTruckType.Items.Add("3T(420*180*190),3000kg");
                                cobTruckType.Items.Add("5T(540*210*210),5000kg");
                                cobTruckType.Items.Add("8T(660*230*230),8000kg");
                                cobTruckType.Items.Add("10T(760*230*235),10000kg");
                                cobTruckType.Items.Add("40H(1200*235*238),20000kg");
                                cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*235*268),20000kg");
                                break;
                            default:
                                cobTruckType.Items.Add("3T(420*180*190),3000kg");
                                cobTruckType.Items.Add("5T(500*200*200),5000kg");
                                cobTruckType.Items.Add("8T(620*210*220),8000kg");
                                cobTruckType.Items.Add("10T(720*240*240),8000kg");
                                cobTruckType.Items.Add("40HQ(1200*235*268),20000kg");
                                cobTruckType.Items.Add("45HQ(1350*235*2568),20000kg");
                                break;

                        }
                    }

                    break;

                default:

                    break;
            }
        }
        private System.Data.DataTable Maketable()
        {
            //從excel匯入的datatable，將數量攤開，列數總和=總數量
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("stackCheck", typeof(bool));//是否可進行堆疊

            dt.Columns.Add("type", typeof(string));

            dt.Columns.Add("num", typeof(string));

            dt.Columns.Add("length", typeof(float));

            dt.Columns.Add("width", typeof(float));

            dt.Columns.Add("height", typeof(float));

            dt.Columns.Add("weight", typeof(float));

            dt.Columns.Add("GCount", typeof(int));

            dt.Columns.Add("stacktag", typeof(bool));//判斷是否已堆疊

            dt.TableName = "dtShow";


            return dt;
        }
        private System.Data.DataTable MakeStacktable()
        {
            //判斷堆疊後加入的datatable，由於legth,width為string無法新增欄位又需要以文字標記堆疊，故分開兩個datatable
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("stackCheck", typeof(bool));//是否可進行堆疊

            dt.Columns.Add("type", typeof(string));

            dt.Columns.Add("num", typeof(string));

            dt.Columns.Add("length", typeof(string));

            dt.Columns.Add("width", typeof(string));

            dt.Columns.Add("height", typeof(string));

            dt.Columns.Add("weight", typeof(string));

            dt.Columns.Add("GCount", typeof(string));

            dt.Columns.Add("stacktag", typeof(bool));//判斷是否已堆疊

            dt.Columns.Add("tag", typeof(bool));//判斷是否已堆疊


            return dt;
        }
        private void dgvGoods_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        class Trucktype
        {
            public string type { get; set; }
        }

        private void goodsorder_Load(object sender, EventArgs e)
        {

        }
    }
}

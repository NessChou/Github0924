using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
namespace ACME
{
    public partial class fmEscoEpay : Form
    {

        System.Data.DataTable dtCost = null;
        public fmEscoEpay()
        {
            InitializeComponent();
        }

        private void fmEscoEpay_Load(object sender, EventArgs e)
        {
            txtMonth.Text = DateTime.Now.AddMonths(-1).ToString("yyyyMM");

            ChangeDate(txtMonth.Text);

            DataTable dt = GetDefaultOrder();

            dataGridView2.DataSource = dt;

            DataTable dt1 = GetParams();

            dataGridView3.DataSource = dt1;
        }

        private void ChangeDate(string strMonth)
        {

            try
            {
                int Days = DateTime.DaysInMonth(Convert.ToInt32(strMonth.Substring(0, 4)), Convert.ToInt32(strMonth.Substring(4, 2)));
                txtStartDate.Text = strMonth + "01";
                txtEndDate.Text = strMonth + Days.ToString();


                DateTime d = new DateTime(Convert.ToInt32(strMonth.Substring(0, 4)), Convert.ToInt32(strMonth.Substring(4, 2)),1).AddMonths(1);

                Days =DateTime.DaysInMonth(d.Year,d.Month);
                txtLimit.Text = d.ToString("yyyyMM")+Days.ToString();
            }
            catch
            {
                MessageBox.Show("請輸入正確月份");
            }
        
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            if (txtMonth.Text.Length == 6)
            {
                ChangeDate(txtMonth.Text);
            }
        }

        private string FormatDateString(string sDate)
        {
            // 減 1911
            Int32 sYear = Convert.ToInt32(sDate.Substring(0, 4)) - 1911;

            //return String.Format("{0}/{1}/{2}",
            //   sYear.ToString(),
            //   sDate.Substring(4, 2),
            //   sDate.Substring(6, 2));
            return String.Format("{0}/{1}/",
               sYear.ToString(),
               sDate.Substring(4, 2));
            //return String.Format("{0}/{1}/{2}",
            //    sDate.Substring(0, 4),
            //    sDate.Substring(4, 2),
            //    sDate.Substring(6, 2));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //使用單號設定檔
            //
            if (txtMonth.Text.Length != 6)
            {
                MessageBox.Show("請輸入正確年月");
               return;
            }


            string StartDate = FormatDateString(txtStartDate.Text);
            DataTable dt = GetData(StartDate);
            f1(dt);
        }

        private void f1(System.Data.DataTable V1)
        {
            //使用單號設定檔
            //
            if (txtMonth.Text.Length != 6)
            {
                MessageBox.Show("請輸入正確年月");
                return;
            }

            DataTable dt = V1;
            System.Data.DataTable dtCost = MakeTableCombine();
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string DOC = dt.Rows[i]["訂單號碼"].ToString();
                dr["訂單號碼"] = DOC;
                dr["客戶代號"] = dt.Rows[i]["客戶代號"].ToString();
                dr["帳單金額"] = Convert.ToInt32(dt.Rows[i]["帳單金額"]);
                dr["備註"] = dt.Rows[i]["備註"].ToString();
                dr["發票號碼"] = dt.Rows[i]["發票號碼"].ToString();
                string AMT = dt.Rows[i]["總金額"].ToString();
                dr["總金額"] = Convert.ToInt32(AMT);
                if (AMT != "0")
                {
                    System.Data.DataTable G1 = GetORD1(DOC, AMT);
                    if (G1.Rows.Count > 0)
                    {
                        System.Data.DataTable G2 = GetORD2(DOC, AMT);
                        int AMT2 = Convert.ToInt32(G2.Rows[0][0]);
                        dr["已付金額"] = AMT2;
                    }
                    else
                    {
                        MessageBox.Show("單號: " + DOC + "金額輸入錯誤");
                    }
                }
                dtCost.Rows.Add(dr);
            }



            dataGridView1.DataSource = dtCost;
        }
        private DataTable GetData(string StartDate)
        {
    

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[DocEntry] 訂單號碼, T0.[CardCode] 客戶代號, T1.[PriceAfVAT] 帳單金額, T1.[Dscription] 備註,T2.MEMO2 發票號碼,ISNULL(T2.MEMO3,0) 總金額  FROM ORDR T0 ");
            sb.Append("  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("  INNER JOIN ACMESQLSP.DBO.ESCO_PAY T2 ON (T0.DOCENTRY=T2.DOCENTRY)");
            sb.Append(string.Format("WHERE T1.[Dscription] like '%{0}%'", StartDate));
            sb.Append(" ORDER BY  T0.[CardCode]");
         
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }


        private DataTable GetDataV(string c, string StartDate)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[DocEntry] 訂單號碼, T0.[CardCode] 客戶代號, T1.[PriceAfVAT] 帳單金額, T1.[Dscription] 備註,T2.MEMO2 發票號碼,ISNULL(T2.MEMO3,0) 總金額  FROM ORDR T0 ");
            sb.Append("  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append("  INNER JOIN ACMESQLSP.DBO.ESCO_PAY T2 ON (T0.DOCENTRY=T2.DOCENTRY)");
            sb.Append(" WHERE T0.[DocEntry] in ( " + c + ") ");
            sb.Append(string.Format("and T1.[Dscription] like '%{0}%'", StartDate));
            sb.Append(" ORDER BY  T0.[CardCode]");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetORD1(string DOCENTRY, string GTOTAL)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT GTOTAL FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND GTOTAL = @GTOTAL");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@GTOTAL", GTOTAL));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetORD2(string DOCENTRY, string GTOTAL)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(GTOTAL) FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND GTOTAL <> @GTOTAL");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@GTOTAL", GTOTAL));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        private DataTable GetDefaultOrder()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT [DocEntry] 訂單號碼,MEMO2 發票號碼,MEMO3 總金額 FROM ESCO_PAY  ");
            sb.Append("WHERE ENABLED ='Y' ");
            

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }


        private DataTable GetParams()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Param_Desc 說明,Param_No 值 FROM ESCO_PARAMS  ");
            sb.Append("WHERE Param_Kind ='DocKind' ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
           
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }

        private DataTable GetINV(string U_IN_BSINV)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT Convert(varchar(10),DOCDATE,111) DATE   FROM OINV WHERE U_IN_BSINV=@U_IN_BSINV ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_IN_BSINV", U_IN_BSINV));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Sales");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables[0];


            return dt;

        }
        //傳入參數
        //dataGridView
        //輸出文字檔 ,附檔名為 csv
        //使用範例  GridViewToCSV(dataGridView1, Environment.CurrentDirectory + @"\dataGridview.csv");
        private void GridViewToCSV(DataGridView dgv, string FileName)
        {

            StringBuilder sbCSV = new StringBuilder();
            int intColCount = dgv.Columns.Count;
            //int intColCount = dgv.Cells.Count;


            //表頭
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                sbCSV.Append(dgv.Columns[i].HeaderText);

                if ((i + 1) != intColCount)
                {
                    sbCSV.Append(",");
                    //tab
                    // sbCSV.Append("\t");
                }

            }
            sbCSV.Append("\n");

            foreach (DataGridViewRow dr in dgv.Rows)
            {

                //資料內容
                for (int x = 0; x < intColCount; x++)
                {

                    if (dr.Cells[x].Value != null)
                    {

                        sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",", "").Replace("\n", "").Replace("\r", ""));
                    }
                    else
                    {
                        sbCSV.Append("");
                    }


                    if ((x + 1) != intColCount)
                    {
                        sbCSV.Append(",");
                        // sbCSV.Append("\t");
                    }
                }
                sbCSV.Append("\n");
            }
            using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.Default))
            {
                sw.Write(sbCSV.ToString());
            }

            System.Diagnostics.Process.Start(FileName);

        }

        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂單號碼", typeof(string));
            dt.Columns.Add("客戶代號", typeof(string));
            dt.Columns.Add("帳單金額", typeof(int));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("總金額", typeof(int));
            dt.Columns.Add("已付金額", typeof(int));
            return dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                StringBuilder sbV = new StringBuilder();
                if (dataGridView1.SelectedRows.Count > 0)
                {
                   

                    ArrayList al = new ArrayList();

                    for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                    {
                        al.Add(listBox1.Items[i].ToString());
                    }




                    foreach (string v in al)
                    {
                        sbV.Append("'" + v + "',");
                    }

                    sbV.Remove(sbV.Length - 1, 1);
                                string StartDate = FormatDateString(txtStartDate.Text);
                                DataTable dtV = GetDataV(sbV.ToString(), StartDate);
                    f1(dtV);

                    Clear(sbV);
                    listBox1.Items.Clear();
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            DataTable dt;

            dt = dataGridView1.DataSource as DataTable;
            if (dt == null)
            {
                MessageBox.Show("請先查詢");
                return;
            }

            try
            {
               

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("沒有資料");
                    return;
                }
            }
            catch
            {
                MessageBox.Show("沒有資料");
                return;
            }




            string FileName = "";

            /// [區別碼]、[整批上傳編碼] 、[客戶代號] 、[繳款截止日] 、[帳單金額] 、[收據編號] 、[帳單起日] 、[帳單迄日] 、[建檔編號] 、[銷帳編號] 、[備註] 、[明細項目1序號] 、[明細項目1值]、[明細項目2序號] 、[明細項目2值].

//            建檔編號:固定值  0040387 固定 7碼

//1	2	3	4	5	6	7
//0	0	4	0	3	8	7

//銷帳編號固定 9碼

//建檔編號 + 銷帳編號 =16碼
//            業務單位代碼
//7: ESCO
//8: Solar	SAP 系統銷售訂單號碼(目前5碼)	收款期數 001~999
            string 業者代號 = "1000000167";
                               
            string 區別碼="D"; 
            //1000000167/909
            //類別不同,則整批上傳編碼也要變
            //商辦空調(1000000167/913/1)
            //string 整批上傳編碼 ="909"; 

            string 整批上傳編碼 = "913"; 

            DataTable dtParam = dataGridView3.DataSource as DataTable;

            整批上傳編碼 = Convert.ToString(dtParam.Rows[0]["值"]);


            string 客戶代號 ="";
            string 繳款截止日 = txtLimit.Text;
            string 帳單金額 =""; 

            string 收據編號 =""; 



            string 建檔編號 = "0040387"; 
            string 銷帳編號 =""; 
            string 備註 ="";
            string 日期 = "";
            string 發票號碼 = "";
            string 金額 ="";
            //string 明細項目1序號 =""; 
            //string 明細項目1值=""; 

            string 訂單號碼 = "";
            string 期數 = "";



            string 明細項目1序號 = "";
            string 明細項目1值 = "";

            StringBuilder sb = new StringBuilder();

            try
            {
                this.Cursor = Cursors.AppStarting;

              
                //輸出範例 1000000003_20110630_001.TXT
                FileName = GetExePath() + "\\" + 業者代號 + "_" + DateTime.Now.ToString("yyyyMMdd") + "_001.TXT";
                //依順序
                //  OutputCrossTable(dt, FileName);


                string Line = "";

                DataRow dr;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dt.Rows[i];

                    客戶代號 =Convert.ToString(dr["客戶代號"]);

                    if (客戶代號 == "7215-00")
                    {
        
                        DateTime d = new DateTime(Convert.ToInt32(txtMonth.Text.Substring(0, 4)), Convert.ToInt32(txtMonth.Text.Substring(4, 2)), 1).AddMonths(2);

                        繳款截止日 = d.ToString("yyyyMM") + "15";
                    
                    }
                    客戶代號 = 客戶代號.Substring(0, 4) + 客戶代號.Substring(5, 2);

                    帳單金額 = Convert.ToString(Convert.ToInt32(dr["帳單金額"]));
                    備註 = Convert.ToString(dr["備註"]);
                    發票號碼 = Convert.ToString(dr["發票號碼"]);
                    System.Data.DataTable t1 = GetINV(發票號碼);
                    if (t1.Rows.Count > 0)
                    {
                        日期 = t1.Rows[0][0].ToString();
                    }
                    //金額 = Convert.ToInt32(dr["總金額"]).ToString("#,##0");
                    金額 = Convert.ToInt32(dr["總金額"]).ToString();
                    訂單號碼 = Convert.ToString(dr["訂單號碼"]);

                    //LineNum 也不一定是對的
                    //42200 節能服務-空調 -103/11/01-103/11/30 -1/36

                    string[] s = 備註.Split('-');

                    string sTemp = s[4].Trim();
                    string 帳單起日 = s[2].Trim().Replace("/", "");
                    string 帳單迄日 = s[3].Trim().Replace("/", "");
                    s = sTemp.Split('/');

                    int num = Convert.ToInt16(s[0]);

                    // 3 位數
                    期數 = num.ToString("000") ; 

                    銷帳編號 = "7" + 訂單號碼 + 期數;

                    //原銷帳編號 0040387723577002 被測試資料使用 先改為 0040387723577001 

                    if (銷帳編號 == "723577002")
                    {
                        銷帳編號 = "723577001";
                    }
                    //104/09/25已預開統一發票RQ68626467金額NTD558,411

                    備註 = 帳單起日 + "-" + 帳單迄日 + " -" + sTemp;

                    if (checkBox1.Checked)
                    {
                        備註 = 備註 + "　備註:" + 日期 + "已預開　　　　統一發票" + 發票號碼 + "　　　　　金額NTD" + 金額;
                    }

                    明細項目1序號 = "1";
                    明細項目1值 = 帳單金額;

                    string DATE1 = (Convert.ToInt32(帳單起日.Substring(0, 3)) + 1911).ToString() + 帳單起日.Substring(3, 4);
                    string DATE2 = (Convert.ToInt32(帳單迄日.Substring(0, 3)) + 1911).ToString() + 帳單迄日.Substring(3, 4);

                    //o	D,24,C000000001,20110720,2000,20110720-3,20110520,20110620,11267,00000000006,備註3,1,2000

                    //沒有明細版本
                   // Line = 區別碼 +","+ 整批上傳編碼 +","+ 客戶代號 +","+ 繳款截止日 +","+ 帳單金額 +","+ 收據編號 +","+ 帳單起日 +","+ 帳單迄日 +","+ 建檔編號 +","+ 銷帳編號 +","+ 備註;
                   
                    //有一筆明細版本
                    Line = 區別碼 + "," + 整批上傳編碼 + "," + 客戶代號 + "," + 繳款截止日 + "," + 帳單金額 + "," + 收據編號 + "," + DATE1 + "," + DATE2 + "," + 建檔編號 + "," + 銷帳編號 + "," + 備註 +
                           "," + 明細項目1序號 + "," + 明細項目1值 + ",,,,,,,,,,,,,,,";
                    
                    if (checkBox1.Checked)
                    {
                        if (t1.Rows.Count > 0)
                        {
                            sb.Append(Line + "\r\n");
                        }
                    }
                    else
                    {
                        sb.Append(Line + "\r\n");
                    }
           
          
                }

                ////輸出檔
                ////string OutPutFile = GetExePath() + "\\Temp\\" +
                ////      Path.GetFileName(FileName) + "_" + DateTime.Now.ToString("yyyyMMdd");

                //string OutPutFile = GetExePath() + "\\Temp\\" +
                //     "管理報表" + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

                ////產生 Excel Report
                //DataTable dt = gvJs.DataSource as DataTable;


                


                using (StreamWriter sw = new StreamWriter(FileName, false, System.Text.Encoding.UTF8))
                {
                    sw.Write(sb.ToString());
                }

            }
            finally
            {
                this.Cursor = Cursors.Default;
               // MessageBox.Show("產生一個檔案-" + FileName);
                System.Diagnostics.Process.Start(FileName);
            }
        }
        public void Clear(StringBuilder value)
        {
            value.Length = 0;
            value.Capacity = 0;
        }
        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {

            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = dataGridView1.SelectedRows[i];
                        listBox1.Items.Add(row.Cells["訂單號碼"].Value.ToString());

                    }
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
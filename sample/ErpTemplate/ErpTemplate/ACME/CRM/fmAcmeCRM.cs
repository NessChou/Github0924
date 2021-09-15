using System;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

//
using System.Data.SqlClient;
//
using System.Collections;
//
using System.IO;

using ACME.CRM;
using ACME.CRM.CRMTableAdapters;
//http://dobon.net/vb/dotnet/control/index.html
//http://www.devolutions.net/articles/DataGridViewFAQ.htm

//工作階段與問題合一
//借出未還 ->倉管 /借出未還全部
//上週己出貨
//本週未出貨
//離倉日期

//tabPage 沒有 visible
//MyTabControl.TabPages.Remove(mTabPage)

//預測 = 月+Model+客戶 ->Brian
//預測模式 2 -> 客戶總量管理 ->EMA
//一般 ->訂單

//詢價作業 (EMA)


// 以月為單位,顯示總表 1  2  3  4
// CardName,Model ->

/* 詢價
 * CREATE TABLE [ACME_INQ] (
	[DocNnum] [varchar] (16) NOT NULL ,
	[CardCode] [varchar] (20) NULL ,
	[CardName] [varchar] (50) NULL ,
	[Model] [varchar] (20) NULL ,
	[Grade] [varchar] (5) NULL ,
	[Ver] [varchar] (5) NULL ,
	[Price] [numeric](18, 4) NULL ,
	[StartDate] [varchar] (8) NULL ,
	[UserCode] [varchar] (20) NOT NULL ,
	[Remark] [varchar] (200) NULL ,
	CONSTRAINT [PK_ACME_INQ] PRIMARY KEY  CLUSTERED 
	(
		[DocNnum],
		[UserCode]
	)  ON [PRIMARY] 
) ON [PRIMARY]
GO


 * 客戶資料增加查詢 by 客戶名 ok
 * RMA 加 己結 ok 
 * 借出還回加上種類 (本來就有..在後面...)
 * 工作週報改成工作報表 ok 
 * 詢價增加數量
 * 銷售預測 版本比較 日期
 * 工作階段的價格 -> 詢價
 * 結案狀態 ->階段+.....Confirm Cancel Closed OK
 * Lead 編號 ->連結 SAP 
 * 新需求
 * 應收帳款連結
 * Spec 連結
 * 20090309  修正工作報告查詢
 * 20090311 增加美金單價
 * 修正銷售單已結的查無資料
 * 查詢增加客戶名稱
 * 
 * 狀態控制
 * DataRowView drv  = (DataRowView)bindingSource.Current;
  textBox.Enabled = drv.IsNew;

 * */

namespace ACME
{
    public partial class fmAcmeCRM : Form
    {
        
         //全域變數 -> UserCode in ['Airy','Ema']
        private string UserPermit = "";
        private string UserFunction = "        ";
        private bool tabPage25FirstLogin = false;
        private string GlobalDate = "";
        private  string FirstLoadFlag="Y";

        private string PrevUserid = "";

        // remember the column index that was last sorted on
        private int prevColIndex = -1;

        // remember the direction the rows were last sorted on (ascending/descending)
        private ListSortDirection prevSortDirection = ListSortDirection.Ascending;

        private string ConnStr = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
        private string ConnStrSP = "server=acmesap;pwd=@rmas;uid=sapdbo;database=AcmesqlSP";
        private string ConnStr02 = "server=acmesap;pwd=@rmas;uid=sapdbo;database=Acmesql02";

        //SAP 員工編號 
        private string EmpID;
        private string EmpName;
        private string ChiNo; //正航


        private DataTable dtFCST;

        //減少開窗選取
        private DataTable dtGlobalOcrd;

        //
        //private DataTable dtOrderData;


        private DataTable PrintDataTable;

        public fmAcmeCRM()
        {


            PrevUserid = globals.UserID;

            //測試
            //globals.UserID = "JOJOHSU";
            TestUser();

            InitializeComponent();
        }

        private void fmAcmeCRM_Load(object sender, EventArgs e)
        {


            label6.Text = "使用者:" + globals.UserID;

            //  this.aCME_STAGETableAdapter.Fill(this.cRM.ACME_STAGE, globals.UserID);

            //20090324 Lleyton的規格為 依類別,讀取不同的檔案
            // Sales -> OSLP
            // Sa -> 員工檔

            string[] s = new string[3];

            s = GetEmpID(globals.UserID);

            EmpID = s[0];
            EmpName = s[1];
            ChiNo = s[2];

          //  MessageBox.Show(ChiNo);

            if (string.IsNullOrEmpty(EmpID))
            {

                MessageBox.Show("沒有對應的業務員編號,請洽 MIS");
                // Close();
                return;
            }




            dataGridView1.AutoGenerateColumns = false;

            dataGridView7.AutoGenerateColumns = false;

            //移至手動查詢
            // dataGridView1.DataSource = GetBP(EmpID, "");


            dataGridView2.AutoGenerateColumns = false;
            //  dataGridView2.DataSource = GetOrderData(EmpID);
            //DataTable dtOrderData = GetOrderData(EmpID);
            //dataGridView2.DataSource = dtOrderData;

            //outlookGrid1.AutoGenerateColumns = false;
            //outlookGrid1.BindData(dtOrderData.DataSet, "ORDR");
            //// setup the column headers
            //// HeaderText 必須手動
            //outlookGrid1.Columns[0].HeaderText = "客戶";
            //outlookGrid1.Columns["Model"].HeaderText = "Model";
            //outlookGrid1.Columns["Qty"].HeaderText = "數量";
            //outlookGrid1.Columns["OpenCreQty"].HeaderText = "未結數量";
            //outlookGrid1.Columns["DocNum"].HeaderText = "單號";
            //outlookGrid1.Columns["DocDate"].HeaderText = "日期";
            //outlookGrid1.Columns["U_ACME_SHIPDAY"].HeaderText = "離倉日期";


            //// setup the column headers
            ////outlookGrid1.Columns.Add("column1", "Id");


            //SetDefaultStyle_Int(outlookGrid1.Columns["Qty"]);
            //SetDefaultStyle_Int(outlookGrid1.Columns["OpenCreQty"]);
            //SetDefaultStyle_Numeric(outlookGrid1.Columns["Price"]);
            //SetGridSort();


            // comboBox1.SelectedIndex = 0;

            //分組

            //如果在元件中有事件就無效
            comboBox2.SelectedValueChanged -= new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            comboBox2.SelectedIndex = 0;
            comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);

            //  exBindingNavigator1.BindingSource = aCME_TASKBindingSource;


            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='STAGE'";
            UtilSimple.SetLookupBinding(comboBox3, connection, sql, aCME_STAGEBindingSource, "Step",
            "PARAM_DESC", "PARAM_DESC");

            //修正 Grade 
            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='GRADE'";
            UtilSimple.SetLookupBinding(comboBox4, connection, sql, aCME_STAGEBindingSource, "GRADE",
            "PARAM_DESC", "PARAM_DESC");


            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='VERSION'";
            UtilSimple.SetLookupBinding(comboBox5, connection, sql, aCME_STAGEBindingSource, "VER",
            "PARAM_DESC", "PARAM_DESC");

            //結案
            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='CloseFlag'";
            UtilSimple.SetLookupBinding(comboBox1, connection, sql, aCME_STAGEBindingSource, "CloseFlag",
            "PARAM_DESC", "PARAM_DESC");


            //設定未結顯示
            aCME_STAGEBindingSource.Filter = "CloseDate = '' or  CloseDate is null ";

            // aCME_QSBindingSource.Filter = "CloseDate = '' or  CloseDate is null ";


            FirstLoadFlag = "N";

            //RMA 維修單
            dataGridView3.AutoGenerateColumns = false;

            // dataGridView5.AutoGenerateColumns = false;

            SetFirstDay();

            textBox3.Text = AcmeDateTimeUtils.DateToStr(GetStartOfPriorWeek(AcmeDateTimeUtils.StrToDate(GlobalDate)));
            textBox4.Text = AcmeDateTimeUtils.DateToStr(GetEndOfPriorWeek(AcmeDateTimeUtils.StrToDate(textBox3.Text)));

            //RMA 維修單
            dataGridView5.AutoGenerateColumns = false;

            //詢價初始化
            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='GRADE'";
            SetComboBoxList(comboBox7, ConnStr, sql, "PARAM_DESC");

            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='VERSION'";
            SetComboBoxList(comboBox6, ConnStr, sql, "PARAM_DESC");


            sql = "SELECT CardCode,CardName  FROM OCRD WHERE CardType='C' and SlpCode= " + EmpID.ToString();
            SetComboBoxList(comboBox8, ConnStr02, sql, "CardName", "CardCode");

            inqTextBox.Text = DateTime.Now.ToString("yyyyMMdd");

            //預測
            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='年'";
            SetComboBoxList(comboBox9, ConnStr, sql, "PARAM_DESC");


            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='預測模式'";
            SetComboBoxList(comboBox10, ConnStr, sql, "PARAM_DESC");

            comboBox10.SelectedIndex = 0;

            //初始化銷銷預測
            dtFCST = MakeTable(12);

            dataGridView9.DataSource = dtFCST;



            for (int i = 0; i <= dataGridView9.Columns.Count - 1; i++)
            {


                if (dtFCST.Columns[i].DataType == typeof(Int32))
                {

                    SetDefaultStyle_Int(dataGridView9.Columns[i]);

                }

            }

            //rma 看最近三個月

            rmaStartDateTxt.Text = (DateTime.Now.AddDays(-90)).ToString("yyyyMMdd");
            rmaEndDateTxt.Text = DateTime.Now.ToString("yyyyMMdd");

            ordStartDateTxt.Text = (DateTime.Now.AddDays(-90)).ToString("yyyyMMdd");
            ordEndDateTxt.Text = DateTime.Now.ToString("yyyyMMdd");


            //詢價 
            dataGridView8.AutoGenerateColumns = false;

            //Hint 顯示

            toolTip1.SetToolTip(button15, "清除畫面上的查詢條件");

            toolTip1.SetToolTip(checkBox2, "勾選後,可將工作階段的價格資訊合併於此顥示 !");

            toolTip1.SetToolTip(button19, "如果有任何需求或是 Bug \r" + "請按下回報");

            toolTip1.SetToolTip(button23, "在 Grid 快按兩下 (Double Click) \r" + "也可開啟明細");

            //  toolTip1.SetToolTip(groupBox6, "Task List \r"+"工作階段的步驟處理詳細記錄");


            //設定行號

            SetGirdLineNum(aCME_STAGEDataGridView);

            //

             // 20090324 UserFunction 第二碼為 1 可以選全部客戶
            dtGlobalOcrd = GetOcrdWithLead(EmpID);
            

            //autoComplete

            comboBox8.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox8.AutoCompleteSource = AutoCompleteSource.CustomSource;


             cardNameTextBox.AutoCompleteMode = AutoCompleteMode.Suggest;
             cardNameTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            foreach (DataRow dr in dtGlobalOcrd.Rows)
            {

                comboBox8.AutoCompleteCustomSource.Add(dr["CardName"].ToString());
                cardNameTextBox.AutoCompleteCustomSource.Add(dr["CardName"].ToString());

            }

            DataTable dtItem = GetItem();

            modelTextBox.AutoCompleteMode = AutoCompleteMode.Suggest;
            modelTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource;

            foreach (DataRow dr in dtItem.Rows)
            {
                modelTextBox.AutoCompleteCustomSource.Add(dr[0].ToString());

            }

            //應收帳款
            textBox13.Text = DateTime.Now.ToString("yyyyMMdd");
        }

        public DataTable GetBP(string SlpCode, string strLike)
        {
            SqlConnection connection = new SqlConnection(ConnStr02);
            string sql = "SELECT T0.CardCode, T0.CardName, T0.Phone1, T0.Phone2, T0.Fax, T0.CntctPrsn, T0.Balance, T0.CreditLine,T0.OrdersBal  FROM OCRD T0 WHERE T0.CardType='C' ";
            
            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
            
              sql+=" and  T0.SlpCode=@SlpCode ";
            }

            if (!string.IsNullOrEmpty(strLike.Trim()))
            {
                sql += " and CardName like @CardName";
            
            }
            
           sql+= " order by  Balance desc ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            }

            if (!string.IsNullOrEmpty(strLike.Trim()))
            {

                command.Parameters.Add(new SqlParameter("@CardName", "%" + strLike + "%"));
            }

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OCRD"];
        }

        public string[] GetEmpID(string LoginId)
        {
            SqlConnection connection = new SqlConnection(ConnStrSP);
            string sql = "SELECT SalesID as EmpID,SapName,ChiNo,Kind from employee where name = @name";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@name", LoginId));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }

            string[] s = new string[3];

            try
            {
                
                s[0] = Convert.ToString(ds.Tables["OCRD"].Rows[0]["EmpID"]);
                s[1] = Convert.ToString(ds.Tables["OCRD"].Rows[0]["SapName"]);
                s[2] = Convert.ToString(ds.Tables["OCRD"].Rows[0]["ChiNo"]);
                return s;
            }
            catch
            {
                return s;
            }
        }


        //WHERE  substring(ItemCode,1,1)='T'
        //and    substring(ItemCode,2,8) ='M170EG01'
        //and    substring(ItemCode,11,1) ='0'
        //and    substring(ItemCode,12,1) ='A'

        //SELECT T0.[DocNum], Convert(varchar(8),T0.[DocDate],112) DocDate,T1.[ItemCode],T1.[Quantity],T1.[Price]  FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry 
        //WHERE T0.[DocStatus]='O' and T0.[SlpCode] =  2 and T0.[DocType]='I'

        //CASE (Substring(T1.[ItemCode],11,1)) 

        //         when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' 

        //          when '1' then 'P' when '2' then 'N' when '3' then 'V' 

        //          when '4' then 'U' when '5' then 'N' ELSE 'X'

        //          END

        //取得未結訂單
        private System.Data.DataTable GetOrderData(string SlpCode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CardName, T0.[DocNum], Convert(varchar(8),T0.[DocDate],112) DocDate,Convert(int,T1.[Quantity]) Qty,T1.[Price],  ");
            //未結量
            sb.Append(" Convert(int,T1.[OpenCreQty]) OpenCreQty, ");
            //離倉日期
            sb.Append(" Convert(Varchar(8), T1.[U_ACME_SHIPDAY],112) U_ACME_SHIPDAY , ");
            sb.Append(" substring(ItemCode,1,9) as Model, ");
            sb.Append(" CASE (Substring(T1.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append(" when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append(" when '4' then 'U' when '5' then 'N' ELSE 'X'");
            sb.Append(" END as Grade,");
            sb.Append(" substring(ItemCode,12,1) as Version,T2.SlpName ");
            sb.Append(" FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry");

            sb.Append(" INNER JOIN  OSLP T2 ON T0.SlpCode = T2.SlpCode ");

            sb.Append(" WHERE T0.[DocStatus]='O' and T0.[DocType]='I' ");

            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {

                sb.Append(" and T0.[SlpCode] = @SlpCode ");
            }
            //列狀態
            sb.Append(" and T1.[LineStatus]='O' ");

            //sb.Append(" GROUP BY Convert(varchar(6),T1.[U_ACME_SHIPDAY],112), T0.[CardName] ");
            //sb.Append(" order BY Convert(varchar(6),T1.[U_ACME_SHIPDAY],112), T0.[CardName] ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //

            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            }

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ORDR");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        //取得已結訂單
        //T1.[LineStatus] ='C'
        //WHERE T0.[DocStatus]='C' 
        private System.Data.DataTable GetOrderData(string SlpCode,string StartDate,string EndDate,string CardName)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CardName, T0.[DocNum], Convert(varchar(8),T0.[DocDate],112) DocDate,Convert(int,T1.[Quantity]) Qty,T1.[Price],  ");
            //未結量
            sb.Append(" Convert(int,T1.[OpenCreQty]) OpenCreQty, ");
            //離倉日期
            sb.Append(" Convert(Varchar(8), T1.[U_ACME_SHIPDAY],112) U_ACME_SHIPDAY , ");
            sb.Append(" substring(ItemCode,1,9) as Model, ");
            sb.Append(" CASE (Substring(T1.[ItemCode],11,1)) ");
            sb.Append(" when 'A' then 'A' when 'B' then 'B' when '0' then 'Z' ");
            sb.Append(" when '1' then 'P' when '2' then 'N' when '3' then 'V' ");
            sb.Append(" when '4' then 'U' when '5' then 'N' ELSE 'X'");
            sb.Append(" END as Grade,");
            sb.Append(" substring(ItemCode,12,1) as Version,T2.SlpName ");
            sb.Append(" FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry");
            sb.Append(" INNER JOIN  OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" WHERE T0.[DocStatus]='C' and T0.[DocType]='I' ");

            //20090324
            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                sb.Append(" and T0.[SlpCode] = @SlpCode ");
            }

            //列狀態
            sb.Append(" and T1.[LineStatus] ='C' ");

            if (!string.IsNullOrEmpty(StartDate))
            {
                sb.Append(" AND  T0.DocDate >= @StartDate");
            }

            if (!string.IsNullOrEmpty(EndDate))
            {
                sb.Append(" AND  T0.DocDate <= @EndDate");
            }


            if (!string.IsNullOrEmpty(CardName))
            {
                sb.Append(" AND  T0.CardName like  @CardName ");
            }

            //sb.Append(" GROUP BY Convert(varchar(6),T1.[U_ACME_SHIPDAY],112), T0.[CardName] ");
            //sb.Append(" order BY Convert(varchar(6),T1.[U_ACME_SHIPDAY],112), T0.[CardName] ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //

            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            }

            if (!string.IsNullOrEmpty(StartDate))
            {
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            }

            if (!string.IsNullOrEmpty(EndDate))
            {
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            }


            if (!string.IsNullOrEmpty(CardName))
            {
                command.Parameters.Add(new SqlParameter("@CardName", "%" + CardName + "%"));
            }


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ORDR");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //處理行號

            using (SolidBrush b = new SolidBrush(dataGridView2.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void dataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //處理行號

            using (SolidBrush b = new SolidBrush( (sender as DataGridView).RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void SetGirdLineNum(DataGridView dg)
        {

            dg.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dataGridView_RowPostPaint);
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {

            try
            {
                //This line of code will help you to change the apperance like size,name,style.
                Font f;
                //For background color
                Brush backBrush;
                //For forground color
                Brush foreBrush;

                //This construct will hell you to deside which tab page have current focus
                //to change the style.
                if (e.Index == this.tabControl1.SelectedIndex)
                {
                    //This line of code will help you to change the apperance like size,name,style.
                  //  f = new Font(e.Font, FontStyle.Bold | FontStyle.Bold);
                    f = new Font(e.Font, FontStyle.Bold);

                    backBrush = new System.Drawing.SolidBrush(Color.SkyBlue);
                    foreBrush = Brushes.White;
                }
                else
                {
                    f = e.Font;
                    backBrush = new SolidBrush(e.BackColor);
                    //backBrush = new System.Drawing.SolidBrush(Color.Gray);
                    foreBrush = new SolidBrush(e.ForeColor);
                }

                //To set the alignment of the caption.
                string tabName = this.tabControl1.TabPages[e.Index].Text;
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;

                //Thsi will help you to fill the interior portion of
                //selected tabpage.
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                Rectangle r = e.Bounds;
                r = new Rectangle(r.X, r.Y + 3, r.Width, r.Height - 3);
                e.Graphics.DrawString(tabName, f, foreBrush, r, sf);

                sf.Dispose();
                if (e.Index == this.tabControl1.SelectedIndex)
                {
                    f.Dispose();
                    backBrush.Dispose();
                }
                else
                {
                    backBrush.Dispose();
                    foreBrush.Dispose();
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "Error Occured", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void outlookGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 && e.ColumnIndex >= 0)
            {
                ListSortDirection direction = ListSortDirection.Ascending;
                if (e.ColumnIndex == prevColIndex) // reverse sort order
                    direction = prevSortDirection == ListSortDirection.Descending ? ListSortDirection.Ascending : ListSortDirection.Descending;

                // remember the column that was clicked and in which direction is ordered
                prevColIndex = e.ColumnIndex;
                prevSortDirection = direction;

                // set the column to be grouped
                outlookGrid1.GroupTemplate.Column = outlookGrid1.Columns[e.ColumnIndex];


                outlookGrid1.Sort(new DataRowComparer(e.ColumnIndex, direction));
            }
        }

        //依第一個欄位分組
        private void SetGridSort()
        {
            ListSortDirection direction = ListSortDirection.Ascending;
            // set the column to be grouped
            outlookGrid1.GroupTemplate.Column = outlookGrid1.Columns[0];
            outlookGrid1.Sort(new DataRowComparer(0, direction));
        }


        
        private void SetDefaultStyle_Int(DataGridViewColumn Column)
        {
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle.Format = "#,##0";
            dataGridViewCellStyle.NullValue = null;
            Column.DefaultCellStyle = dataGridViewCellStyle;
        }

        private void SetDefaultStyle_Numeric(DataGridViewColumn Column)
        {
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle.Format = "#,##0.00";
            dataGridViewCellStyle.NullValue = null;
            Column.DefaultCellStyle = dataGridViewCellStyle;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //避免第一次開窗就執行
            if (FirstLoadFlag == "Y") return;
            
            if ( Convert.ToString(comboBox2.SelectedItem) == "Model")
            {

                SetGridGroup("Model");
            }
            else if (Convert.ToString(comboBox2.SelectedItem) == "客戶")
            {
                SetGridGroup("CardName");
            }

            else if (Convert.ToString(comboBox2.SelectedItem) == "日期")
            {
                SetGridGroup("DocDate");
            }
        }

        private void SetGridGroup(string ColName)
        {

            ListSortDirection direction = ListSortDirection.Ascending;
            if (outlookGrid1.Columns[ColName].Index == prevColIndex) // reverse sort order
                direction = prevSortDirection == ListSortDirection.Descending ? ListSortDirection.Ascending : ListSortDirection.Descending;

            // remember the column that was clicked and in which direction is ordered
            prevColIndex = outlookGrid1.Columns[ColName].Index;
            prevSortDirection = direction;

            // set the column to be grouped
            outlookGrid1.GroupTemplate.Column = outlookGrid1.Columns[ColName];


            outlookGrid1.Sort(new DataRowComparer(outlookGrid1.Columns[ColName].Index, direction));

        }

        private void aCME_TASKDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch
            { 
            
            }
        }

        private void exBindingNavigator1_BeforePost(object sender, MyEventArgs args)
        {
            //if (tASK_IDTextBox.Text.Length <= 4)
            //{
            //    args._CheckOk = false;
            //    tASK_IDTextBox.Focus();
            //    MessageBox.Show("長度不足 5");
            //    return;
            //}

            args._CheckOk = true;
          //  MessageBox.Show(args.ToString());
        }

        private void exBindingNavigator1_AfterNew(object sender, EventArgs e)
        {
            //sTART_DATEDateTimePicker.Value = DateTime.Now;
            //eND_DATEDateTimePicker.Value = DateTime.Now;
            userCodeTextBox.Text = globals.UserID;
            startDateTextBox.Text = DateTime.Now.ToString("yyyyMMdd");
          //  stageNoTextBox.Text = DateTime.Now.ToString("yyyyMMddhhmmssss");

            //給值
            ((DataRowView)aCME_STAGEBindingSource.Current).Row["UserCode"] = globals.UserID;
            ((DataRowView)aCME_STAGEBindingSource.Current).Row["StageNo"] = DateTime.Now.ToString("yyyyMMddhhmmssss");

            cardNameTextBox.Focus();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            //DataView dv = new DataView(cRM.ACME_STAGE);
            //dv.RowFilter = 
            //aCME_STAGEBindingSource.DataSource = dv;
            aCME_STAGEBindingSource.Filter = "CloseDate = '' or  CloseDate is null ";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            //DataView dv = new DataView(cRM.ACME_STAGE);
            //dv.RowFilter = "CloseDate <> '' ";
            //aCME_STAGEBindingSource.DataSource = dv;

            aCME_STAGEBindingSource.Filter = "CloseDate <> '' ";

           
        }




        //取得未結的服務契約
        //有結沒結的欄位要重新規劃
        private System.Data.DataTable GetRMAData(string SlpCode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ContractID],T0.[u_pkind],T0.[U_RMA_NO],T0.[U_Cusname_S],T0.[U_RVender],  ");
            sb.Append(" T0.[U_RModel],T0.[U_RVer],t0.U_rgrade,t0.U_rquinity,T0.[U_racmetodate],T0.[U_routwharehouse],");
            sb.Append(" T0.[U_rengineer],T0.[U_rsales],T0.[U_rtoreceiving],T0.[U_AUO_RMA_NO],T0.[U_repaircenter],");
            sb.Append(" T0.[U_acme_out],T0.[U_yetqty], T0.[U_acme_recedate],T0.[U_acme_backdate],T0.[U_acme_qback] ");
            sb.Append(" ,T0.[U_acme_backdate1],T0.[U_acme_backqty1]");
            sb.Append(" FROM OCTR T0 ");
            sb.Append(" WHERE   T0.[U_PKind]  = '2' and U_rma_no <> 'null' and U_rsales = @U_rsales");
            sb.Append(" and ( (T0.[U_cusname_s]='' or T0.[U_cusname_s] is null) or (T0.[U_rmodel]='' or T0.[U_rmodel] is null)");
            sb.Append("  or (T0.[U_rver]='' or T0.[U_rver] is null) or (T0.[U_rgrade]='' or T0.[U_rgrade] is null)");
            sb.Append("  or (T0.[U_rquinity]='' or T0.[U_rquinity] is null) or ( T0.[U_racmetodate] is null) ");
            sb.Append("  or (T0.[U_rengineer]='' or T0.[U_rengineer] is null) or (T0.[U_rsales]='' or T0.[U_rsales] is null)");
            sb.Append("  or (T0.[U_rtoreceiving] is null)  or (T0.[U_auo_rma_no]='' or T0.[U_auo_rma_no] is null)   ");
            sb.Append("  or (T0.[U_repaircenter]='' or T0.[U_repaircenter] is null) or ( T0.[U_acme_out] is  null) ");
            sb.Append("  or (T0.[U_yetqty]='' or T0.[U_yetqty] is null) or ( T0.[U_acme_recedate] is null )");
            sb.Append("  or ( T0.[U_acme_qback]='' or T0.[U_acme_qback] is null) ");
            sb.Append("  or ( T0.[U_acme_backdate1] is null)  or ( T0.[U_acme_backqty1]='' or T0.[U_acme_backqty1] is null))");

            if (!string.IsNullOrEmpty(rmaStartDateTxt.Text.Trim()))
            {
                sb.Append(" AND  T0.StartDate >= @StartDate");
            }

            if (!string.IsNullOrEmpty(rmaEndDateTxt.Text.Trim()))
            {
                sb.Append(" AND  T0.StartDate <= @EndDate");
            }



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@U_rsales", SlpCode));

            if (!string.IsNullOrEmpty(rmaStartDateTxt.Text.Trim()))
            {
                command.Parameters.Add(new SqlParameter("@StartDate", rmaStartDateTxt.Text));
            }

            if (!string.IsNullOrEmpty(rmaEndDateTxt.Text.Trim()))
            {
                command.Parameters.Add(new SqlParameter("@EndDate", rmaEndDateTxt.Text));
            }



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCTR");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        //取得已結的服務契約
        //有結沒結的欄位要重新規劃
        private System.Data.DataTable GetRMAData(string SlpCode,string StartDate,string EndDate)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[ContractID],T0.[u_pkind],T0.[U_RMA_NO],T0.[U_Cusname_S],T0.[U_RVender],  ");
            sb.Append(" T0.[U_RModel],T0.[U_RVer],t0.U_rgrade,t0.U_rquinity,T0.[U_racmetodate],T0.[U_routwharehouse],");
            sb.Append(" T0.[U_rengineer],T0.[U_rsales],T0.[U_rtoreceiving],T0.[U_AUO_RMA_NO],T0.[U_repaircenter],");
            sb.Append(" T0.[U_acme_out],T0.[U_yetqty], T0.[U_acme_recedate],T0.[U_acme_backdate],T0.[U_acme_qback] ");
            sb.Append(" ,T0.[U_acme_backdate1],T0.[U_acme_backqty1]");
            sb.Append(" FROM OCTR T0 ");
            sb.Append(" WHERE   T0.[U_PKind]  = '2' and U_rma_no <> 'null' and U_rsales = @U_rsales");

            if (! string.IsNullOrEmpty(StartDate))
            {
                sb.Append(" AND  T0.StartDate >= @StartDate");
            }

            if (! string.IsNullOrEmpty(EndDate))
            {
                sb.Append(" AND  T0.StartDate <= @EndDate");
            }



            sb.Append(" and ( (T0.[U_cusname_s] <>'') or (T0.[U_rmodel]<>'' or T0.[U_rmodel] is not null)");
            sb.Append("  or (T0.[U_rver] <> '' or T0.[U_rver] is not  null) or (T0.[U_rgrade]<>'' or T0.[U_rgrade] is not null)");
            sb.Append("  or (T0.[U_rquinity]<>'' or T0.[U_rquinity] is not null) or ( T0.[U_racmetodate] is not  null) ");
            sb.Append("  or (T0.[U_rengineer]<>'' or T0.[U_rengineer] is not null) or (T0.[U_rsales]<>'' or T0.[U_rsales] is not null)");
            sb.Append("  or (T0.[U_rtoreceiving] is not null)  or (T0.[U_auo_rma_no]<>'' or T0.[U_auo_rma_no] is not null)   ");
            sb.Append("  or (T0.[U_repaircenter]<>'' or T0.[U_repaircenter] is not null) or ( T0.[U_acme_out] is not null) ");
            sb.Append("  or (T0.[U_yetqty]<>'' or T0.[U_yetqty] is not null) or ( T0.[U_acme_recedate] is not null )");
            sb.Append("  or ( T0.[U_acme_qback]<>'' or T0.[U_acme_qback] is not null) ");
            sb.Append("  or ( T0.[U_acme_backdate1] is not null)  or ( T0.[U_acme_backqty1]<>'' or T0.[U_acme_backqty1] is not null))");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@U_rsales", SlpCode));


            if (!string.IsNullOrEmpty(StartDate))
            {
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            }

            if (!string.IsNullOrEmpty(EndDate))
            {
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            }


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCTR");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        //先利用 服務 的 StartDate
        private void button1_Click(object sender, EventArgs e)
        {
            //區分己結 未結

            if (radioButton6.Checked)
            {
                dataGridView3.DataSource = GetRMAData(EmpName);
            }
                //己結
            else if (radioButton5.Checked)
            {
                dataGridView3.DataSource = GetRMAData(EmpName,rmaStartDateTxt.Text,rmaEndDateTxt.Text);
            }
        }

        private string FormatDateStr(string sDate)
        {

            return string.Format("{0}/{1}/{2}", sDate.Substring(0, 4), sDate.Substring(4, 2), sDate.Substring(6, 2));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string StartDateOfWeek = AcmeDateTimeUtils.DateToStr(GetStartOfPriorWeek(AcmeDateTimeUtils.StrToDate(GlobalDate)));
            string EndDateOfWeek = AcmeDateTimeUtils.DateToStr(GetEndOfPriorWeek(AcmeDateTimeUtils.StrToDate(StartDateOfWeek)));
            //txtStartDate.Text = FormatDateStr(StartDateOfWeek);
            //txtEndDate.Text = FormatDateStr(EndDateOfWeek);

            txtStartDate.Text = StartDateOfWeek;
            txtEndDate.Text = EndDateOfWeek;

            GlobalDate = StartDateOfWeek;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string StartDateOfWeek = AcmeDateTimeUtils.DateToStr(GetStartOfLastWeek(AcmeDateTimeUtils.StrToDate(GlobalDate)));
            string EndDateOfWeek = AcmeDateTimeUtils.DateToStr(GetEndOfLastWeek(AcmeDateTimeUtils.StrToDate(StartDateOfWeek)));
            //txtStartDate.Text = FormatDateStr(StartDateOfWeek);
            //txtEndDate.Text = FormatDateStr(EndDateOfWeek);
            txtStartDate.Text = StartDateOfWeek;
            txtEndDate.Text = EndDateOfWeek;
            GlobalDate = StartDateOfWeek;
        }

        public DateTime GetStartOfLastWeek(DateTime aDate)
        {
            DateTime dt = aDate.AddDays(7);
            return new DateTime(dt.Year, dt.Month, dt.Day, 0, 0, 0, 0);
        }

        public DateTime GetEndOfLastWeek(DateTime aDate)
        {
            DateTime dt = aDate.AddDays(6);
            return new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59, 999);
        }


        public DateTime GetStartOfPriorWeek(DateTime aDate)
        {
            DateTime dt = aDate.AddDays(-7);
            return new DateTime(dt.Year, dt.Month, dt.Day, 0, 0, 0, 0);
        }

        public DateTime GetEndOfPriorWeek(DateTime aDate)
        {
            DateTime dt = aDate.AddDays(6);
            return new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59, 999);
        }

        private void SetFirstDay()
        {
            //取得本週的第一天

            string StartDateOfWeek = AcmeDateTimeUtils.DateToStr(AcmeDateTimeUtils.GetStartDateOfWeek(DateTime.Today));
            string EndDateOfWeek = AcmeDateTimeUtils.DateToStr(AcmeDateTimeUtils.GetEndDateOfWeek(DateTime.Today));

            //txtStartDate.Text = FormatDateStr(StartDateOfWeek);
            //txtEndDate.Text = FormatDateStr(EndDateOfWeek);

            txtStartDate.Text = StartDateOfWeek;
            txtEndDate.Text = EndDateOfWeek;

            //MessageBox.Show(FormatDateStr(StartDateOfWeek));
            GlobalDate = StartDateOfWeek;
        }

        private void exBindingNavigator2_BeforePost(object sender, MyEventArgs args)
        {
            args.CheckOk = true;
        }

        //private void exBindingNavigator2_AfterNew(object sender, EventArgs e)
        //{
        //    qsNoTextBox.Text = DateTime.Now.ToString("yyyyMMddhhmmssss");
        //    userCodeTextBox1.Text = globals.UserID;

        //    //
        //    startDateTextBox1.Text = DateTime.Now.ToString("yyyyMMdd");
        //    predDateTextBox1.Text = DateTime.Now.AddDays(7).ToString("yyyyMMdd");
        //}



        public DataTable GetSecurity(string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT [UserCode],[UserFunction],[UserPermit] FROM [ACME_CRM_SEC] WHERE UserCode=@UserCode";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_CRM_SEC");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_CRM_SEC"];
        }

        //利用 ACME_CRM_SEC 設定

        private void TestUser()
        {

            DataTable dt = GetSecurity(globals.UserID);

            if (dt.Rows.Count == 0)
            {
                return;
            }

            UserFunction =Convert.ToString(dt.Rows[0]["UserFunction"]);

            //全域變數 -> UserCode in ['Airy','Ema']
            UserPermit = Convert.ToString(dt.Rows[0]["UserPermit"]);

            //第一碼代表可以使用模擬業務功能
            if (UserFunction.Substring(0, 1) == "1")
            {

                CrmSales fmCrmSales = new CrmSales();
                if (fmCrmSales.ShowDialog() == DialogResult.OK)
                {
                    globals.UserID = fmCrmSales.EmpID;
                }
            }

            //20090324
            //第二碼代表可以查詢所有客戶



            //if (globals.UserID.ToUpper() == "TERRYLEE" ||
            //    globals.UserID.ToUpper() == "LLEYTONCHEN"||
            //    globals.UserID.ToUpper() == "ANNIECHEN")
            //{

            //    CrmSales fmCrmSales = new CrmSales();
            //    if (fmCrmSales.ShowDialog() == DialogResult.OK)
            //    {
            //        globals.UserID = fmCrmSales.EmpID;
            //    }
            
            //}
        
        }

        private void fmAcmeCRM_Leave(object sender, EventArgs e)
        {
            globals.UserID = PrevUserid;
        }

        private void aCME_STAGEDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch
            {

            }
        }



        //開啟 Lookup 視窗
        private object[] GetOcrdList()
        {
            string[] FieldNames = new string[] { "CardName", "CardCode" };

            string[] Captions = new string[] { "客戶名稱", "客戶編號" };


            string SqlScript = "SELECT CardCode,CardName  FROM OCRD WHERE SlpCode=" + EmpID;


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;
            //dialog.LookUpConnection = new SqlConnection(ConnStr02); ;

            //dialog.SqlScript = SqlScript;

           // dialog.SourceDataTable = GetOcrdWithLead(EmpID);
            dialog.SourceDataTable = dtGlobalOcrd;

            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        private object[] GetOcrdListSingle()
        {
            string[] FieldNames = new string[] { "CardCode", "CardName" };

            string[] Captions = new string[] { "客戶編號", "客戶名稱" };


            string SqlScript = "SELECT CardCode,CardName  FROM OCRD WHERE SlpCode=" + EmpID;


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;
            dialog.LookUpConnection = new SqlConnection(ConnStr02); 

            dialog.SqlScript = SqlScript;

           // dialog.SourceDataTable = GetOcrdWithLead(EmpID);

            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        private DataTable GetOcrdWithLead(string UserCode)
        {
            //合計 AS 銷售金額
            DataTable dtCombine;
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            //sb.Append(" SELECT CardName,Cardcode FROM OCRD WHERE SlpCode=" + EmpID +" order by CardName");
            sb.Append(" SELECT CardName,Cardcode FROM OCRD WHERE 1=1 ");

            if ( UserFunction.Substring(1,1)=="1" )
            {
                sb.Append(" order by CardName");
            }
            else
            {
                
                sb.Append(" and SlpCode=" + EmpID + " order by CardName");
            }

          
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }

            dtCombine = ds.Tables[0];

            DataTable dtLead =GetLead(globals.UserID);

            DataRow dr;
            for (int i = 0; i <= dtLead.Rows.Count - 1; i++)
            {
                dr = dtCombine.NewRow();

                dr["CardCode"] = dtLead.Rows[i]["CardCode"];
                dr["CardName"] = dtLead.Rows[i]["CardName"];
                dtCombine.Rows.Add(dr);
            
            }




            return dtCombine;
        }

        private DataTable GetLead(string UserCode)
        {
            
            SqlConnection connection = new SqlConnection(ConnStr);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT Cardcode,CardName FROM ACME_LEAD WHERE UserCode=@UserCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_LEAD");
            }
            finally
            {
                connection.Close();
            }



            return ds.Tables[0];
        }


        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetOcrdList();

            if (LookupValues != null)
            {
                cardCodeTextBox.Text = Convert.ToString(LookupValues[1]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[0]);

            }
        }

        /// <summary>
        /// 產生工作週報        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dt = MakeTable();

            // dtOrderData - Global
           // DataTable dtOrderData = GetOrderData(EmpID);

            DataRow dr;

            //規格修正.. Order 不產生
            //for (int i = 0; i <= dtOrderData.Rows.Count - 1; i++)
            //{

                
            //    dr = dt.NewRow();
            //    dr["Step"] = "Order Process";
            //  //  dr["CardCode"] = Convert.ToString(dtOrderData.Rows[i]["CardCode"]);
            //    dr["CardName"] = Convert.ToString(dtOrderData.Rows[i]["CardName"]);
            //    dr["Model"] = Convert.ToString(dtOrderData.Rows[i]["Model"]);
            //    dr["Ver"] = Convert.ToString(dtOrderData.Rows[i]["Version"]);
            //    dr["Grade"] = Convert.ToString(dtOrderData.Rows[i]["Grade"]);
            //    dr["Price"] = Convert.ToDecimal(dtOrderData.Rows[i]["Price"]);
            //    dr["Qty"] = Convert.ToInt32(dtOrderData.Rows[i]["OpenCreQty"]);


            //    dr["StartDate"] = Convert.ToString(dtOrderData.Rows[i]["DocDate"]);

            //    try
            //    {
            //        dr["PredDate"] = Convert.ToString(dtOrderData.Rows[i]["U_ACME_SHIPDAY"]);
            //    }
            //    catch
            //    {
            //        dr["PredDate"] = "";
            //    }



            //    dt.Rows.Add(dr);
            
            //}




            DataTable dtStage = GetStageData(globals.UserID,
                txtStartDate.Text.Replace("/","")                ,
                txtEndDate.Text.Replace("/", ""));
            for (int i = 0; i <= dtStage.Rows.Count - 1; i++)
            {


                dr = dt.NewRow();
                dr["Step"] = Convert.ToString(dtStage.Rows[i]["Step"]);
                
                dr["CardName"] = Convert.ToString(dtStage.Rows[i]["CardName"]);
                dr["Model"] = Convert.ToString(dtStage.Rows[i]["Model"]);
                dr["Ver"] = Convert.ToString(dtStage.Rows[i]["Ver"]);
                dr["Grade"] = Convert.ToString(dtStage.Rows[i]["Grade"]);
                try
                {
                    dr["Price"] = Convert.ToDecimal(dtStage.Rows[i]["Price"]);
                }
                catch
                {
                    dr["Price"] = 0;
                }

                try
                {
                    dr["Qty"] = Convert.ToInt32(dtStage.Rows[i]["Qty"]);
                }
                catch
                {
                    dr["Qty"] = 0;
                }


                dr["StartDate"] = Convert.ToString(dtStage.Rows[i]["StartDate"]);

                try
                {
                    dr["PredDate"] = Convert.ToString(dtStage.Rows[i]["PredDate"]);
                }
                catch
                {
                    dr["PredDate"] = "";
                }


                dr["IssueDesc"] = Convert.ToString(dtStage.Rows[i]["IssueDesc"]);
                dr["ActionDesc"] = Convert.ToString(dtStage.Rows[i]["ActionDesc"]);
                dr["Remark"] = Convert.ToString(dtStage.Rows[i]["Remark"]);

                //TaskList
                dr["TaskList"] = GetTaskList( Convert.ToString(dtStage.Rows[i]["StageNo"]), globals.UserID);



                dt.Rows.Add(dr);

            }



            DataView dv = new DataView(dt);

            dv.Sort = "CardName Desc,Step Desc,Model Asc ";

            dataGridView5.DataSource = dv;


            PrintDataTable = dt;
        }


        //
       // GetTaskList
        private string GetTaskList(string StageNo,string UserCode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr);

            StringBuilder sb = new StringBuilder();

            //合併明細
            sb.Append(" SELECT T.* FROM ACME_Stage_D T Where StageNo=@StageNo and UserCode=@UserCode");

         
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@StageNo", StageNo));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            

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

            DataTable dt  = ds.Tables[0];

            string s="";
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                s += string.Format("日期:{0},問題描述:{1},行動方案:{2}", Convert.ToString(dt.Rows[i]["StartDate"]),
                                       Convert.ToString(dt.Rows[i]["IssueDesc"]),
                    Convert.ToString(dt.Rows[i]["ActionDesc"])) +"\n";
            
            }
            return s;

        }

        //動態產生資料結構
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("Step", typeof(string));
            dt.Columns.Add("CardCode", typeof(string));
            dt.Columns.Add("CardName", typeof(string));
            dt.Columns.Add("Model", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Ver", typeof(string));
            dt.Columns.Add("Price", typeof(decimal));
            dt.Columns.Add("Qty", typeof(Int32));
            dt.Columns.Add("IssueDesc", typeof(string));
            dt.Columns.Add("ActionDesc", typeof(string));
            dt.Columns.Add("StartDate", typeof(string));
            dt.Columns.Add("PredDate", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            // Task 明細
            dt.Columns.Add("TaskList", typeof(string));

            //dt.Columns.Add("階段", typeof(string));
            //dt.Columns.Add("客戶編號", typeof(string));
            //dt.Columns.Add("客戶名稱", typeof(string));
            //dt.Columns.Add("Model", typeof(string));
            //dt.Columns.Add("Version", typeof(string));
            //dt.Columns.Add("Grade", typeof(string));
            //dt.Columns.Add("數量", typeof(Int32));
            //dt.Columns.Add("單價", typeof(decimal));
            //dt.Columns.Add("問題描述", typeof(string));
            //dt.Columns.Add("行動方案", typeof(string));
            //dt.Columns.Add("開始日期", typeof(string));
            //dt.Columns.Add("預計完成日期", typeof(string));
            //dt.Columns.Add("備註說明", typeof(string));
            
            
            /*
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["SERIAL_NO"];
            dt.PrimaryKey = colPk;
            */

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }

        ///
//        SELECT * FROM ACME_Stage
//WHERE StartDate >'20080201'
//AND     ( CloseDate is null or CloseDate<>'' )
        /// <summary>
        /// 取得符合未結 及日期
        /// </summary>
        /// <param name="SlpCode"></param>
        /// <returns></returns>
        private System.Data.DataTable GetStageData(string SlpCode,string StartDate,string EndDate)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr);

            StringBuilder sb = new StringBuilder();

            //合併明細
            //未結
            sb.Append(" SELECT T.* FROM ACME_Stage T");
            sb.Append(" WHERE  ( (T.CloseDate is null or T.CloseDate ='') ");
            sb.Append(" and   T.UserCode = @SlpCode ) " );

            //己結

            sb.Append(" or  (T.CloseDate >= @StartDate ");
            sb.Append(" AND  T.CloseDate <= @EndDate ");
            sb.Append(" and (T.CloseDate is not null or T.CloseDate<>'') ");
            sb.Append(" and   T.UserCode = @SlpCode )");

            //單一
            //sb.Append(" SELECT T.* FROM ACME_Stage T");

            //sb.Append(" WHERE  (StartDate >= @StartDate ");
            //sb.Append(" AND   StartDate <= @EndDate ");

            //sb.Append(" and   UserCode = @SlpCode )");

            //sb.Append(" OR  (StartDate < @StartDate ");
            //sb.Append(" AND     ( CloseDate is null or CloseDate<>'' ) ");
            //sb.Append(" and   UserCode = @SlpCode )");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

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

            return ds.Tables[0];

        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // First row always displays

            if (e.RowIndex == 0)

                return;

            //設限欄位
            if (e.ColumnIndex > 0)

                return;
            
            if (IsRepeatedCellValue( e.RowIndex, e.ColumnIndex))
            {

                e.Value = string.Empty;

                e.FormattingApplied = true;

            }


        }

        private bool IsRepeatedCellValue(int rowIndex, int colIndex)
        {


            DataGridViewCell currCell = dataGridView5[colIndex,rowIndex];

            DataGridViewCell prevCell = dataGridView5[colIndex, rowIndex-1];


            if ((currCell.Value == prevCell.Value) ||

               (currCell.Value != null && prevCell.Value != null &&

               currCell.Value.ToString() == prevCell.Value.ToString()))
            {

                return true;

            }

            else
            {

                return false;

            }

        }

        private void dataGridView5_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;

            // Ignore column and row headers and first row

            if (e.RowIndex < 1 || e.ColumnIndex < 0)

                return;

            //設限欄位
            if (e.ColumnIndex > 0)
            {
                e.AdvancedBorderStyle.Top = dataGridView5.AdvancedCellBorderStyle.Top;
                return;
            }


            if (IsRepeatedCellValue(e.RowIndex, e.ColumnIndex))
            {
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
            }

            else
            {

                e.AdvancedBorderStyle.Top = dataGridView5.AdvancedCellBorderStyle.Top;

            }

        }

        private void fmAcmeCRM_Layout(object sender, LayoutEventArgs e)
        {
            //
            if (string.IsNullOrEmpty(EmpID))
            {

               
                Close();
               // return;
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView2, Environment.CurrentDirectory + @"\"+globals.UserID+DateTime.Now.ToString("yyMMddhhmmss")+".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        //傳入參數
        //dataGridView
        //輸出文字檔 ,附檔名為 csv
        //使用範例  GridViewToCSV(dataGridView1, Environment.CurrentDirectory + @"\dataGridview.csv");
        private void GridViewToCSV(DataGridView dgv, string FileName)
        {

            StringBuilder sbCSV = new StringBuilder();
            int intColCount = dgv.Columns.Count;


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

                        sbCSV.Append(dr.Cells[x].Value.ToString().Replace(",","").Replace("\n","").Replace("\r",""));
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

        private void button7_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView5, Environment.CurrentDirectory + @"\" + globals.UserID + DateTime.Now.ToString("yyMMddhhmmss") + ".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView1, Environment.CurrentDirectory + @"\" + globals.UserID + DateTime.Now.ToString("yyMMddhhmmss") + ".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView3, Environment.CurrentDirectory + @"\" + globals.UserID + DateTime.Now.ToString("yyMMddhhmmss") + ".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }


        /// <summary>
        /// 借出未還
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            DataTable dt = GetBorrowData(EmpID);
            dataGridView6.DataSource = dt;


            //合併正航
            if (!string.IsNullOrEmpty(ChiNo))
            {

                DataTable dtBorrow = GetChiBorrow(ChiNo);

                DataRow dr;
                for (int i = 0; i <= dtBorrow.Rows.Count - 1; i++)
                {

                    dr = dt.NewRow();

                    dr["單號"] = dtBorrow.Rows[i]["單號"];
                    dr["調撥日期"] = dtBorrow.Rows[i]["日期"];

                    dr["產品編號"] = dtBorrow.Rows[i]["產品編號"];
                    dr["品名"] = dtBorrow.Rows[i]["品名規格"];
                    dr["客戶編號"] = dtBorrow.Rows[i]["客戶編號"];
                    dr["客戶名稱"] = dtBorrow.Rows[i]["公司簡稱"];

                    dr["借出數量"] = dtBorrow.Rows[i]["數量"];
                    dr["已還數量"] = dtBorrow.Rows[i]["還貨數量"];
                    dr["未還數量"] = dtBorrow.Rows[i]["未還數量"];
                    dr["預設銷售人員"] = dtBorrow.Rows[i]["姓名"];
                    dr["種類"] = dtBorrow.Rows[i]["借出類別"];
                    dr["備註"] = dtBorrow.Rows[i]["備註"];

                    dt.Rows.Add(dr);                  
                
                }

            
            
            }

            //單號不轉
            for (int i = 1; i <= dataGridView6.Columns.Count - 1; i++)
            {


                if (dt.Columns[i].DataType == typeof(Int32))
                {

                    SetDefaultStyle_Int(dataGridView6.Columns[i]);

                }

            }
        }

        private System.Data.DataTable GetBorrowData(string SlpCode)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT Convert(varchar(20),T0.DocNum) 單號, Convert(varchar(8),T0.[DocDate],112) 調撥日期,T1.ItemCode 產品編號, T0.CardCode 客戶編號, T0.CardName 客戶名稱,  T1.Dscription as 品名, Convert(int, T1.Quantity) as 借出數量,");
            sb.Append(" Convert(int, (select isnull(sum(W.Quantity),0) from WTR1 W ");
            sb.Append(" INNER JOIN OWTR O ON O.DocEntry =W.DocEntry ");
            sb.Append(" WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode)) 已還數量,");
            sb.Append(" Convert(int, T1.Quantity-( select isnull(sum(W.Quantity),0) from WTR1 W ");
            sb.Append(" INNER JOIN OWTR O ON O.DocEntry =W.DocEntry ");
            sb.Append(" WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode)) as 未還數量,");
            sb.Append(" 預設銷售人員 =(SELECT T9.SlpName FROM OCRD C INNER JOIN OSLP T9 ON T9.SlpCode = C.SlpCode WHERE C.CardCode =T0.CardCode),");
            //sb.Append(" T2. [SlpName],t0.u_acme_kind1 種類,t0.comments 處理人員,T0.JRNLMEMO 備註");
            sb.Append(" t0.u_acme_kind1 種類,t0.comments 處理人員,T0.JRNLMEMO 備註");
            sb.Append(" FROM OWTR T0  ");
            sb.Append(" INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" INNER JOIN OCRD T3 ON T0.CardCode = T3.CardCode ");
            sb.Append(" WHERE T0.U_ACME_KIND ='1'");
            sb.Append(" AND T1.Quantity-( select isnull(sum(W.Quantity),0) from WTR1 W ");
            sb.Append(" INNER JOIN OWTR O ON O.DocEntry =W.DocEntry ");
            sb.Append(" WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode) >0");
            //設限業務
            sb.Append(" AND   T3.SlpCode =@SlpCode");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            //command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            //command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

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

            return ds.Tables[0];

        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView6, Environment.CurrentDirectory + @"\" + globals.UserID + DateTime.Now.ToString("yyMMddhhmmss") + ".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DataTable dt = GetOinvData(EmpID,textBox3.Text,textBox4.Text,textBox11.Text);

            dataGridView7.DataSource = dt;

            //

            //for (int i = 0; i <= dataGridView7.Columns.Count - 1; i++)
            //{ 
            
            //  if (dt.Columns[i].DataType == typeof(Int32))
            //  {

            //      SetDefaultStyle_Int(dataGridView7.Columns[i]);
              
            //  }
            
            //}

        }

        /// <summary>
        /// 20090611 增加 2008年1~2月
        /// </summary>
        /// <param name="SlpCode"></param>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
        /// <param name="CardName"></param>
        /// <returns></returns>
        private System.Data.DataTable GetOinvData(string SlpCode, string StartDate,string EndDate,string CardName)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(ConnStr02);

            StringBuilder sb = new StringBuilder();


           // sb.Append(" SELECT  採購員姓名,旗標,");
            sb.Append(" SELECT  旗標,");
            sb.Append("  日期,");
            sb.Append("  單號, 客戶名稱,");
            sb.Append("  產品編號,  品名規格,  數量,  單價 , 美金單價, ");
            sb.Append("  金額 , Convert(int, 數量 * 美金單價) AS 美金金額,採購員姓名 ");
            sb.Append(" FROM( ");
            sb.Append(" SELECT T2.[SlpName] AS 採購員姓名,'銷貨' as 旗標,");
            sb.Append(" Convert(Varchar(8),T0.DocDate,112) AS 日期,");
            sb.Append(" IsNull(T0.DocEntry,'') AS 單號, T0.CardName AS 客戶名稱,");
            sb.Append(" IsNull(T1.ItemCode,'') AS 產品編號, IsNull(T1.Dscription,'') AS 品名規格, Convert(int,IsNull(T1.Quantity,0)) AS 數量, Convert(int,IsNull(T1.Price,0)) AS 單價, ");
            sb.Append(" Convert(int,IsNull(T1.LineTotal,0)) AS 金額, t4.price as 美金單價");
            sb.Append(" FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T1.ItemCode  ");
            sb.Append(" INNER JOIN  OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            //加入美金單價
            sb.Append(" left join dln1 t3 on (t1.baseentry=T3.docentry and  t1.baseline=t3.linenum )");
            sb.Append(" left join rdr1 t4 on (t3.baseentry=T4.docentry and  t3.baseline=t4.linenum )");
            sb.Append(" WHERE T0.[DocType] ='I'");
            sb.Append("  and  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            sb.Append(" AND   T0.DocDate  >=@StartDate ");
            sb.Append(" AND   T0.DocDate  <=@EndDate ");

            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                sb.Append(" AND   T0.SlpCode =@SlpCode ");
            }

            if (!string.IsNullOrEmpty(CardName))
            {
                sb.Append(" AND   T0.CardName  like @CardName ");
            
            }

            sb.Append(" UNION ALL "); //UNION ALL 才會全部出現
            sb.Append(" SELECT T2.[SlpName] AS 採購員姓名,'銷退' as 旗標,");
            sb.Append(" Convert(Varchar(8),T0.DocDate,112) AS 日期,");
            sb.Append(" IsNull(T0.DocEntry,'') AS 單號, T0.CardName AS 客戶名稱,");
            sb.Append(" IsNull(T1.ItemCode,'') AS 產品編號, IsNull(T1.Dscription,'') AS 品名規格, Convert(int,IsNull(T1.Quantity,0)) AS 數量, Convert(int,IsNull(T1.Price,0)) AS 單價, ");
            sb.Append(" Convert(int,IsNull(T1.LineTotal,0)) AS 金額, 0 as 美金單價 ");
            sb.Append(" FROM ORIN T0  INNER JOIN RIN1 T1 ON T0.DocEntry = T1.DocEntry INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T1.ItemCode ");
            sb.Append(" INNER JOIN  OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append(" WHERE T0.[DocType] ='I'");
            sb.Append(" AND  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            //設限業務

            sb.Append(" AND   T0.DocDate  >=@StartDate");
            sb.Append(" AND   T0.DocDate  <=@EndDate");
            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                sb.Append(" AND   T0.SlpCode =@SlpCode");
            }

            if (!string.IsNullOrEmpty(CardName))
            {
                sb.Append(" AND   T0.CardName  like @CardName ");

            }

            //
            //20090611 增加 2008年1~2月
            if (AcmeDateTimeUtils.StrToDate(StartDate) <= AcmeDateTimeUtils.StrToDate("20080229"))
            {

                sb.Append(" UNION ALL "); //UNION ALL 才會全部出現
                sb.Append(" SELECT T2.[SlpName] AS 採購員姓名,'銷貨' as 旗標,");
                sb.Append(" Convert(Varchar(8),T0.DocDate,112) AS 日期,");
                sb.Append(" IsNull(T0.DocEntry,'') AS 單號, T0.CardName AS 客戶名稱,");
                sb.Append(" IsNull(T1.ItemCode,'') AS 產品編號, IsNull(T1.Dscription,'') AS 品名規格, Convert(int,IsNull(T1.Quantity,0)) AS 數量, Convert(int,IsNull(T1.Price,0)) AS 單價, ");
                sb.Append(" Convert(int,IsNull(T1.LineTotal,0)) AS 金額, t4.price as 美金單價");
                sb.Append(" FROM acmesql01..OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T1.ItemCode");
                sb.Append(" INNER JOIN  acmesql01..OSLP T2 ON T0.SlpCode = T2.SlpCode ");
                //加入美金單價
                sb.Append(" left join acmesql01..dln1 t3 on (t1.baseentry=T3.docentry and  t1.baseline=t3.linenum )");
                sb.Append(" left join acmesql01..rdr1 t4 on (t3.baseentry=T4.docentry and  t3.baseline=t4.linenum )");
                sb.Append(" WHERE T0.[DocType] ='I'");
                sb.Append(" AND ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND   T0.DocDate  >=@StartDate ");
                sb.Append(" AND   T0.DocDate  <=@EndDate ");

                if (UserFunction.Substring(1, 1) == "1")
                {

                }
                else
                {
                    sb.Append(" AND   T0.SlpCode =@SlpCode ");
                }

                if (!string.IsNullOrEmpty(CardName))
                {
                    sb.Append(" AND   T0.CardName  like @CardName ");

                }

                sb.Append(" UNION ALL "); //UNION ALL 才會全部出現
                sb.Append(" SELECT T2.[SlpName] AS 採購員姓名,'銷退' as 旗標,");
                sb.Append(" Convert(Varchar(8),T0.DocDate,112) AS 日期,");
                sb.Append(" IsNull(T0.DocEntry,'') AS 單號, T0.CardName AS 客戶名稱,");
                sb.Append(" IsNull(T1.ItemCode,'') AS 產品編號, IsNull(T1.Dscription,'') AS 品名規格, Convert(int,IsNull(T1.Quantity,0)) AS 數量, Convert(int,IsNull(T1.Price,0)) AS 單價, ");
                sb.Append(" Convert(int,IsNull(T1.LineTotal,0)) AS 金額, 0 as 美金單價 ");
                sb.Append(" FROM acmesql01..ORIN T0  INNER JOIN acmesql01..RIN1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" INNER JOIN  acmesql01..OSLP T2 ON T0.SlpCode = T2.SlpCode ");
                sb.Append(" WHERE T0.[DocType] ='I'");
                sb.Append(" AND T1.ItemCode not like 'Z%'  ");
                //設限業務

                sb.Append(" AND   T0.DocDate  >=@StartDate");
                sb.Append(" AND   T0.DocDate  <=@EndDate");
                if (UserFunction.Substring(1, 1) == "1")
                {

                }
                else
                {
                    sb.Append(" AND   T0.SlpCode =@SlpCode");
                }

                if (!string.IsNullOrEmpty(CardName))
                {
                    sb.Append(" AND   T0.CardName  like @CardName ");

                }
            
            
            
            
            
            }
            sb.Append(" ) T  order by 日期");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

            if (UserFunction.Substring(1, 1) == "1")
            {

            }
            else
            {
                command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));
            }

            if (!string.IsNullOrEmpty(CardName))
            {
                command.Parameters.Add(new SqlParameter("@CardName", "%"+CardName+"%"));

            }


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

            return ds.Tables[0];

        }

        public void  SetComboBoxList(ComboBox MyComboBox ,string MyConnStr,  string Sql,string FieldName)
        {
            SqlConnection connection = new SqlConnection(MyConnStr);
            SqlCommand command = new SqlCommand(Sql, connection);
            command.CommandType = CommandType.Text;
            
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Data");
            }
            finally
            {
                connection.Close();
            }
            //return ds.Tables["OCRD"];
            DataTable dt = ds.Tables["Data"];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                MyComboBox.Items.Add(Convert.ToString(dt.Rows[i][FieldName]));
            }

        }


        public void SetComboBoxList(ComboBox MyComboBox, string MyConnStr, string Sql, string DisplayMember,
           string ValueMember)
        {
            SqlConnection connection = new SqlConnection(MyConnStr);
            SqlCommand command = new SqlCommand(Sql, connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Data");
            }
            finally
            {
                connection.Close();
            }
            //return ds.Tables["OCRD"];
            DataTable dt = ds.Tables["Data"];

            DataRow dr;

            dr = dt.NewRow();
            dr[ValueMember] = "0";
            dr[DisplayMember] = "--Please Select--";
            dt.Rows.Add(dr);

            DataView dv = dt.DefaultView;

            dv.Sort = ValueMember + " ASC ";

            MyComboBox.DataSource = dv;
            MyComboBox.DisplayMember = DisplayMember;
            MyComboBox.ValueMember =  ValueMember;

            //for (int i = 0; i <= dt.Rows.Count - 1; i++)
            //{
            //    MyComboBox.Items.Add(Convert.ToString(dt.Rows[i][FieldName]));
            //}

        }

        private void button13_Click(object sender, EventArgs e)
        {
            ////取值
            //MessageBox.Show(Convert.ToString(comboBox8.SelectedValue));
            ////名稱 Error
            //MessageBox.Show(Convert.ToString(comboBox8.SelectedItem));
            ////空白
            //MessageBox.Show(Convert.ToString(comboBox8.SelectedText));
            ////取名稱
            //MessageBox.Show(comboBox8.Text);

            string CardCode =Convert.ToString(comboBox8.SelectedValue);
            string Model = txtModel.Text;

            DataTable dt = GetACME_INQ(CardCode,Model, globals.UserID);
            

            //增加可查詢 階段中的報價資料 
            if (checkBox2.Checked)
            {
                DataTable dtStage = GetACME_INQFromStage(CardCode, Model, globals.UserID);


                DataRow dr;
                for (int i = 0; i <= dtStage.Rows.Count - 1; i++)
                {

                    dr = dt.NewRow();
                    for (int j = 0; j <= dtStage.Columns.Count - 1; j++)
                    {

                        dr[j] = dtStage.Rows[i][j];
                    }

                    dt.Rows.Add(dr);
                
                }

            }

            bindingSource1.DataSource = dt;

            dataGridView8.DataSource = bindingSource1;

        }


        public  DataTable GetACME_INQ(string CardCode,string Model, string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT 'INQ' DocFlag,DocNnum  as DocNum,CardCode,CardName,Model,Grade,Ver,Price,StartDate,UserCode,Remark FROM ACME_INQ WHERE 1= 1  AND UserCode=@UserCode ";


            if (CardCode != "0")
            {
                sql += " AND CardCode=@CardCode ";

            }

            if (!string.IsNullOrEmpty(Model.Trim()))
            {
               // sql += " AND Model like '"+Model+"%' ";
                sql += " AND Model like @Model ";

            }


            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            

            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));


            if (CardCode != "0")
            {
             
                command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            }

            if (!string.IsNullOrEmpty(Model.Trim()))
            {

                command.Parameters.Add(new SqlParameter("@Model", Model+"%"));
            }

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_INQ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_INQ"];
        }

        public DataTable GetACME_INQFromStage(string CardCode, string Model, string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT 'Stage' DocFlag, StageNo as DocNnum,CardCode,CardName,Model,Grade,Ver,Price,StartDate,UserCode,Remark FROM ACME_STAGE WHERE 1= 1  AND UserCode=@UserCode AND Price is not null ";


            if (CardCode != "0")
            {
                sql += " AND CardCode=@CardCode ";

            }

            if (!string.IsNullOrEmpty(Model.Trim()))
            {
                // sql += " AND Model like '"+Model+"%' ";
                sql += " AND Model like @Model ";

            }


            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;




            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));


            if (CardCode != "0")
            {

                command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            }

            if (!string.IsNullOrEmpty(Model.Trim()))
            {

                command.Parameters.Add(new SqlParameter("@Model", Model + "%"));
            }

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_INQ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_INQ"];
        }



        public  void AddACME_INQ(string DocNnum, string CardCode, string CardName, string Model, string Grade, string Ver, decimal Price,Int32 Qty, string StartDate, string UserCode, string Remark)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            SqlCommand command = new SqlCommand("Insert into ACME_INQ(DocNnum,CardCode,CardName,Model,Grade,Ver,Price,Qty,StartDate,UserCode,Remark) values(@DocNnum,@CardCode,@CardName,@Model,@Grade,@Ver,@Price,@Qty,@StartDate,@UserCode,@Remark)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNnum", DocNnum));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@Grade", Grade));
            command.Parameters.Add(new SqlParameter("@Ver", Ver));
            command.Parameters.Add(new SqlParameter("@Price", Price));
            command.Parameters.Add(new SqlParameter("@Qty", Qty));
            command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            //command.Parameters.Add(new SqlParameter("@Remark", Remark));
            command.Parameters.Add(new SqlParameter("@Remark", SqlDbType.VarChar, 200));
            command.Parameters["@Remark"].Value = Remark;
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        private string GetKeyString()
        {
          return DateTime.Now.ToString("yyyyMMddhhmmssss");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string DocNnum  = GetKeyString();

            string CardCode = Convert.ToString(comboBox8.SelectedValue);

            string CardName = comboBox8.Text;
            string Model =txtModel.Text;
            string Grade =comboBox7.Text;
            string Ver  = comboBox6.Text;

            decimal Price = 0;

            string StartDate = inqTextBox.Text;
            string UserCode =globals.UserID;
            string Remark = txtRemark.Text;
            //
            Int32 Qty = 0;

            if (Remark.Length >=200)
            {
              Remark = Remark.Substring(0,200);
            }



            if (CardCode=="0")
            {
                // MessageBox.Show("Model 必須輸入");
                errorProvider1.SetError(comboBox8, "客戶 必須輸入");

                return;
            }
            else
            {
                errorProvider1.SetError(comboBox8, "");
            }

            if (string.IsNullOrEmpty(Model.Trim()))
            {
                // MessageBox.Show("Model 必須輸入");
                errorProvider1.SetError(txtModel, "Model 必須輸入");

                return;
            }
            else
            {
                errorProvider1.SetError(txtModel, "");
            }


            try
            {
                Price = Convert.ToDecimal(txtPrice.Text);
            }
            catch
            {
                errorProvider1.SetError(txtPrice, "請輸入正確單價");

                return;
            }
            errorProvider1.SetError(txtPrice, "");


            //檢查數值

            if (!string.IsNullOrEmpty(txtQty.Text))
            {
                try
                {
                    Qty = Convert.ToInt32(txtQty.Text);
                }
                catch
                {
                    errorProvider1.SetError(txtQty, "請輸入正確數量");

                    return;
                }

                errorProvider1.SetError(txtQty, "");
            }
            else
            {
                errorProvider1.SetError(txtQty, "");
            }

            try
            {

                AddACME_INQ(DocNnum, CardCode, CardName, Model, Grade, Ver, Price, Qty, StartDate, UserCode, Remark);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            if (checkBox1.Checked)
            {
                button13_Click(sender, e);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ////取值
            MessageBox.Show(Convert.ToString(comboBox8.SelectedValue));
          
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            comboBox8.SelectedIndex = 0;
            txtModel.Text = "";
        }

       
        private void exBindingNavigator2_AfterNew(object sender, EventArgs e)
        {
            //給值
            ((DataRowView)aCME_LEADBindingSource.Current).Row["UserCode"] = globals.UserID;
            ((DataRowView)aCME_LEADBindingSource.Current).Row["DocNum"] = DateTime.Now.ToString("yyyyMMddhhmmssss");
        }

        private void exBindingNavigator2_BeforePost_1(object sender, MyEventArgs args)
        {
            args._CheckOk = true;


            if (string.IsNullOrEmpty(cardCodeTextBox1.Text.Trim()))
            {
                errorProvider1.SetError(cardCodeTextBox1, "客戶編號 必須輸入");
                args._CheckOk = false;
                return;
            }
            else
            {
                errorProvider1.SetError(cardCodeTextBox1, "");
            }

            if (cardCodeTextBox1.Text.Substring(0,1) !="L")
            {
                errorProvider1.SetError(cardCodeTextBox1, "客戶編號第一碼必須為 L ");
                args._CheckOk = false;
                return;
            }
            else
            {
                errorProvider1.SetError(cardCodeTextBox1, "");
            }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetOcrdListSingle();

            if (LookupValues != null)
            {
                textBox6.Text = Convert.ToString(LookupValues[0]);
                textBox5.Text = Convert.ToString(LookupValues[1]);

            }
        }

        //銷售預測
        //固定 12 個月
        private System.Data.DataTable MakeTable(int ColCount)
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("CardCode", typeof(string));
            dt.Columns.Add("CardName", typeof(string));
            dt.Columns.Add("Model", typeof(string));
            dt.Columns.Add("Remark", typeof(string));

            string ColName = "";
            
            //01,02~12
            for (int i = 1; i <= ColCount; i++)
            {
                dt.Columns.Add("F" + i.ToString("00"), typeof(Int32));
            }

            //Model 不能重覆
            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["CardCode"];
            colPk[1] = dt.Columns["Model"];
            dt.PrimaryKey = colPk;


            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            dt.AcceptChanges();

            return dt;
        }

        private System.Data.DataTable MakeVerticalTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("CodeName", typeof(string));
            
            return dt;
        }


        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox10.SelectedIndex == 0)
            {
                dataGridView9.Columns[0].Visible = false;
                dataGridView9.Columns[1].Visible = false;
                dataGridView9.Columns[2].Visible = true;
            }
            else
            {
                dataGridView9.Columns[0].Visible = false;
                dataGridView9.Columns[1].Visible = false;
                dataGridView9.Columns[2].Visible = false;
            }
        }


        //給預設值
        private void dataGridView9_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            //Visible False 一樣可用
          // e.Row.Cells["dataGridView9CardCode"].Value = "例行工作";
           // (DataRowView)e.Row. = "123";
           // MessageBox.Show("");
            e.Row.Cells["dataGridView9CardCode"].Value = textBox6.Text;
            e.Row.Cells["dataGridView9CardName"].Value = textBox5.Text;
            if (comboBox10.SelectedIndex == 1)
            {
                e.Row.Cells["dataGridView9Model"].Value = "0";
            }

        }

        private void dataGridView9_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //DocNnum 像是專案代號

        public void AddACME_FCST(string DocNum, string CardCode, string CardName, string Model, int Qty, string FcstMonth, string UserCode, string Remark)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            SqlCommand command = new SqlCommand("Insert into ACME_FCST(DocNum,CardCode,CardName,Model,Qty,FcstMonth,UserCode,Remark) values(@DocNum,@CardCode,@CardName,@Model,@Qty,@FcstMonth,@UserCode,@Remark)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@Model", Model));
            command.Parameters.Add(new SqlParameter("@Qty", Qty));
            command.Parameters.Add(new SqlParameter("@FcstMonth", FcstMonth));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        //寫入 Log 檔 
        //DocNnum + UserCode 才是唯一號
        public void AddACME_FCST_LOG(string DocNum, string FcstMonth,string UserCode)
        {

            string DocVersion = Convert.ToString(GetMaxVersion(DocNum, FcstMonth, UserCode));

            //if (DocVersion == "0")
            //{ 
              
            //}
            
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "Insert into ACME_FCST_LOG(DocNum,CardCode,CardName,Model,Qty,FcstMonth,UserCode,Remark,DocVersion) " +
                " SELECT  '"+DocNum+"',CardCode,CardName,Model,Qty,FcstMonth,UserCode,Remark,'" +DocVersion+"' "+
                " FROM ACME_FCST  " +
                " WHERE  DocNum=@DocNum and UserCode=@UserCode " +
                " AND    FcstMonth LIKE '" + FcstMonth + "%' "; ;


            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            //command.Parameters.Add(new SqlParameter("@FcstMonth", FcstMonth));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }


        public  DataTable GetACME_FCST(string CardCode, string UserCode,string Model)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT DocNum,CardCode,CardName,Model,Qty,FcstMonth,UserCode,Remark FROM ACME_FCST WHERE  UserCode=@UserCode AND CardCode=@CardCode  ";

            sql += " AND FcstMonth LIKE '" + comboBox9.Text + "%' ";

            if (!string.IsNullOrEmpty(Model))
            {
                sql += " AND Model ='0'";
            }
            else
            {
                sql += " AND Model <>'0'";
            }

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_FCST");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_FCST"];
        }


        public DataTable GetACME_FCSTByGroup(string CardCode, string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT Distinct DocNum,CardCode,CardName,Model FROM ACME_FCST WHERE  UserCode=@UserCode ";

           // sql += " AND FcstMonth LIKE '" + comboBox9.Text + "%' ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_FCST");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_FCST"];
        }


        public Int32 GetMaxVersion(string DocNum, string FcstMonth, string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT  isnull(max(docversion),'0') from acme_fcst_log "+
                 " WHERE  DocNum=@DocNum and UserCode=@UserCode " +
                " AND    FcstMonth LIKE '" + FcstMonth + "%' "; 

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
          

            try
            {
                connection.Open();
                object A = command.ExecuteScalar();

                if (A == null)
                {
                    return 1;
                }
                else
                {
                    return Convert.ToInt32(A) + 1;
                }
            }
            finally
            {
                connection.Close();
            }
        }

        public string GeDocNum(string Model, string UserCode, string FcstMonth)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT  isnull(max(DocNum),'') from acme_fcst " +
               "WHERE UserCode=@UserCode "+
               " AND FcstMonth LIKE '" + FcstMonth + "%' ";

            if (Model == "0")
            { 
              sql+= "AND Model=@Model";
            }

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));

            if (Model == "0")
            {
                command.Parameters.Add(new SqlParameter("@Model", Model));
            }

            try
            {
                connection.Open();
                return (string)command.ExecuteScalar() ;
            }
            finally
            {
                connection.Close();
            }
        }


        public void DeleteACME_FCST(string Model, string UserCode, string FcstMonth)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "DELETE ACME_FCST WHERE UserCode=@UserCode ";

            sql += " AND FcstMonth LIKE '" + FcstMonth + "%' ";

            if (Model == "0")
            {
                sql += "AND Model=@Model";
            }

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            if (Model == "0")
            {
                command.Parameters.Add(new SqlParameter("@Model", Model));
            }
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            string CardCode = textBox6.Text;
            string UserCode =globals.UserID;


            string Model="";

            DataTable dt = null;

            if (comboBox10.SelectedIndex == 0)
            {
                dt = GetACME_FCST(CardCode, UserCode, Model);

                
            }
            else
            {
                //固定為 0
                Model = "0";
                dt = GetACME_FCST(CardCode, UserCode, Model);
            }


            //取出 DocNum 
            string DocNum = "";
            comboBoxVersion.Items.Clear();

            if (dt.Rows.Count > 0)
            {
                DocNum = Convert.ToString(dt.Rows[0]["DocNum"]);

                DataTable dtDocVersion = GetFcstDocNum(DocNum);

                //comboBoxVersion.Items.Clear();
                for (int i = 0; i <= dtDocVersion.Rows.Count - 1; i++)
                {
                    comboBoxVersion.Items.Add(Convert.ToString(dtDocVersion.Rows[i][0]));
                }
            }

            


            dtFCST.Clear();

            ConvertXTable(dt, dtFCST);

            dtFCST.AcceptChanges();

            //20090305 有 BindingSource 才可刪除
            bsAcme_fcst.DataSource = dtFCST;
            dataGridView9.DataSource = bsAcme_fcst;

           // dataGridView9.DataSource = dtFCST;

            //Conver to CrossTable
        }


//        select distinct DocVersion from acme_fcst_log
//where docnum ='20090303061805'
//order by DocVersion

        public DataTable GetFcstDocNum(string DocNum)
        {
            SqlConnection connection = new SqlConnection(ConnStr);

            string sql = "select distinct DocVersion from acme_fcst_log " +
                         "where docnum =@DocNum " +
                         "order by DocVersion";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_FCST");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_FCST"];
        }

        private void ConvertXTable(DataTable dt , DataTable dtTo)
        {

            DataRow dr;
            DataRow drFind;

            dtTo.Clear();

            string CardCode = "";
            string CardName = "";
            string Model = "";
            string Remark = "";

            Int32 Qty = 0;
            
            string FieldPos = "";


            object[] colPk = new object[2];


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                CardCode = Convert.ToString(dt.Rows[i]["CardCode"]);
                CardName = Convert.ToString(dt.Rows[i]["CardName"]);
                Model = Convert.ToString(dt.Rows[i]["Model"]);

                Remark = Convert.ToString(dt.Rows[i]["Remark"]);

                Qty = Convert.ToInt32(dt.Rows[i]["Qty"]);

                FieldPos = "F"+Convert.ToString(dt.Rows[i]["FcstMonth"]).Substring(4,2);


                colPk[0] = CardCode;
                colPk[1] = Model;

                drFind = dtTo.Rows.Find(colPk);

                if (drFind == null)
                {
                    dr = dtTo.NewRow();

                    dr["CardCode"] = CardCode;
                    dr["CardName"] = CardName;
                    dr["Remark"] = Remark;
                    dr["Model"] = Model;
                    dr[FieldPos] = Qty;
                    dtTo.Rows.Add(dr);
                }
                else
                {
                    drFind.BeginEdit();
                    drFind[FieldPos] = Qty;
                    drFind.EndEdit();

                }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //全部都先刪除

            //DeleteACME_TIMESHEET(UserID,
            //                     txtStartDate.Text.Replace("/", ""),
            //                     txtEndDate.Text.Replace("/", ""));

           // DataTable dt = dtFCST.GetChanges(DataRowState.Added);

            if (comboBox10.SelectedIndex == 1)
            {
                if (dataGridView9.Rows.Count != 1)
                {
                    MessageBox.Show("客戶總量模式時,只能輸入一筆資料");
                    return;
                }
            }

            



            if (string.IsNullOrEmpty(textBox6.Text))
            {
                errorProvider1.SetError(textBox6, "客戶 必須輸入");

                return;
            }
            else
            {
                errorProvider1.SetError(textBox6, "");

            }

            if (string.IsNullOrEmpty(textBox5.Text))
            {
                errorProvider1.SetError(textBox5, "客戶 必須輸入");

                return;
            }
            else
            {
                errorProvider1.SetError(textBox5, "");

            }

            if (string.IsNullOrEmpty( comboBox9.Text))
            {
                errorProvider1.SetError(comboBox9, "年度 必須輸入");

                return;
            }
            else
            {
                errorProvider1.SetError(comboBox9, "");

            }




            DataTable dt = dtFCST.GetChanges();

            if (dt == null)
            {
                return;

            }

            //dataGridView10.DataSource = dt;
            //return;


            //

            string s = "";



            string CardCode = textBox6.Text;
            string CardName = textBox5.Text;
            string Model = "";
            Int32  Qty = 0;
            string Remark = "";
            string DocNum ="";
            string FcstMonth = comboBox9.Text;
            string UserCode = globals.UserID;

            //客戶+ Model
            if (comboBox10.SelectedIndex == 0)
            { 

            }
            else if (comboBox10.SelectedIndex == 1)
            {
                //存鍵值使用
                Model = "0";
            }

           // string FcstMode = comboBox10.Text;

            //先作 Log...

            DocNum = GeDocNum(Model, UserCode, FcstMonth);

            if (!string.IsNullOrEmpty(DocNum))
            {
                AddACME_FCST_LOG(DocNum, FcstMonth, UserCode);

                //刪除資料 --如果是新增時....?????
                DeleteACME_FCST(Model, UserCode, FcstMonth);
            }
            else
            {
                //保留原來的 kEY 
                DocNum = GetKeyString();
            }

            if (null != dt && 0 < dt.Rows.Count)
            {
                foreach (DataRow row in dt.Rows)
                {
                    //CardCode = Convert.ToString(row["CardCode"]);

                    Model = Convert.ToString(row["Model"]);
                    Remark = Convert.ToString(row["Remark"]);


                    //前面有四個固定欄位
                    for (int j = 4; j < dt.Columns.Count; j++)
                    {
                        //row[j, DataRowVersion.Original].ToString())
                        //System.Diagnostics.Debug.WriteLine(string.Format("{0}\t", row[j, DataRowVersion.Original].ToString()));
                        try
                        {
                            s = Convert.ToString(row[j]);

                            
                        }
                        catch
                        {

                        }
                        //MessageBox.Show(s);                    }

                        if (!string.IsNullOrEmpty(s))
                        {
                            //寫入資料庫
                            try
                            {
                                Qty = Convert.ToInt32(s);
                            }
                            catch
                            {}

                           string LocalFcstMonth = FcstMonth + (j-3).ToString("00");

                           AddACME_FCST(DocNum, CardCode, CardName, Model, Qty, LocalFcstMonth, UserCode, Remark);
                        }
                        
                    }

                }

            }

            MessageBox.Show("存檔完成 !");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            CrmMis aForm = new CrmMis();
            aForm.UserId = PrevUserid;
            aForm.ShowDialog();
        }


        //客戶資料查詢
        private void button20_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetBP(EmpID, textBox7.Text);

        }

        private void button21_Click(object sender, EventArgs e)
        {
            DataTable dtOrderData=null;

            if (radioButton1.Checked)
            {
                dtOrderData = GetOrderData(EmpID);
                dataGridView2.DataSource = dtOrderData;
            }
            else if (radioButton2.Checked)
            {
                dtOrderData = GetOrderData(EmpID, ordStartDateTxt.Text, ordEndDateTxt.Text,textBox10.Text);
                dataGridView2.DataSource = dtOrderData;
            }

            //20090305

            if (dtOrderData.Rows.Count == 0)
            {
                return;
            }

            outlookGrid1.AutoGenerateColumns = false;
            outlookGrid1.BindData(dtOrderData.DataSet, "ORDR");
            // setup the column headers
            // HeaderText 必須手動
            outlookGrid1.Columns[0].HeaderText = "客戶";
            outlookGrid1.Columns["Model"].HeaderText = "Model";
            outlookGrid1.Columns["Qty"].HeaderText = "數量";
            outlookGrid1.Columns["OpenCreQty"].HeaderText = "未結數量";
            outlookGrid1.Columns["DocNum"].HeaderText = "單號";
            outlookGrid1.Columns["DocDate"].HeaderText = "日期";
            outlookGrid1.Columns["U_ACME_SHIPDAY"].HeaderText = "離倉日期";


            // setup the column headers
            //outlookGrid1.Columns.Add("column1", "Id");


            SetDefaultStyle_Int(outlookGrid1.Columns["Qty"]);
            SetDefaultStyle_Int(outlookGrid1.Columns["OpenCreQty"]);
            SetDefaultStyle_Numeric(outlookGrid1.Columns["Price"]);
            SetGridSort();

        }

        private void exBindingNavigator1_BeforeDelete(object sender, MyEventArgs args)
        {
            //結案時,不能刪
            args._CheckOk = true;

            //
            string CloseDate;

            try
            {
                 CloseDate = Convert.ToString(((DataRowView)aCME_STAGEBindingSource.Current).Row["CloseDate"]);
            }
            catch
            { 
            CloseDate="";
            }

            if (! string.IsNullOrEmpty(CloseDate))
            {
                args._CheckOk = false;

                MessageBox.Show("資料已結,無法刪除");
            }
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

    
            DataGridViewRow drv = dataGridView8.CurrentRow;

            string DocNum = Convert.ToString(drv.Cells["dg8DocNum"].Value);
            string DocFlag = Convert.ToString(drv.Cells["dg8DocFlag"].Value);

            if (DocFlag == "INQ")
            {

                DeleteACME_INQ(DocNum, globals.UserID);
                //MessageBox.Show(DocFlag + "-" + DocNum);//

                bindingSource1.RemoveCurrent();

            }
            else
            {
                MessageBox.Show("資料來源為 Stage 時,請至 工作階段維護資料 !");
            }
        }

        public  void DeleteACME_INQ(string DocNnum, string UserCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "DELETE ACME_INQ WHERE DocNnum=@DocNnum AND UserCode=@UserCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocNnum", DocNnum));
            command.Parameters.Add(new SqlParameter("@UserCode", UserCode));
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }

        private void startDateTextBox_Validating(object sender, CancelEventArgs e)
        {
            //游標進出就會檢查...
            //20090305 bug
            if (string.IsNullOrEmpty((sender as TextBox).Text))
            {
                errorProvider1.SetError(sender as TextBox, "");
                return;
            }

            if (!AcmeDateTimeUtils.IsDate((sender as TextBox).Text))
            {
                errorProvider1.SetError(sender as TextBox, "日期輸入錯誤");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(sender as TextBox, "");
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage25 && ! tabPage25FirstLogin)
            { 
            
                tabPage25FirstLogin =true;

                string sql = "SELECT DISTINCT  Product_Category  as  Product_Category  FROM AuProduct";
                SetComboBoxList(comboBox11, ConnStr, sql, "Product_Category");

                //sql = "SELECT DISTINCT  Model_No  as  Model_No  FROM AuProduct order by Model_No ";
                //SetComboBoxList(comboBox12, ConnStr, sql, "Model_No");

                comboBox11.SelectedIndex = 0;

            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            // 11

            dataGridView11.DataSource = GetAU_Product(comboBox11.Text, comboBox12.Text);
        }


        public DataTable GetAU_Product(string Product_Category, string Model_No)
        {
            SqlConnection connection = new SqlConnection(ConnStr);
            string sql = "SELECT [Screen_Type],[Model_No],[Voltage],[Physical_Size],[Display_Resolution], "+
      "[Pixel_Pitch] ,[Dot_Pitch]  ,[Viewing_Angle]  ,[CCFL_Number]  ,[Brightness]  ,[Dimensions], "+
      "[Contrast_Ratio] ,[Response_Time]      ,[Lamp_Current]     ,[Power]  ,[Temp_storage],"+
      "[Temp_operation],[Interface],[Color_Bits],[Weight],[CCFL_LifeTime] ,[RoHS] ,[Memo] FROM [AuProduct]  "+
      " where Product_Category =@Product_Category ";

            if (!string.IsNullOrEmpty(Model_No))
            {
                sql += " AND Model_No LIKE '" + comboBox12.Text + "%' ";
            }

            

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Product_Category", Product_Category));
            //command.Parameters.Add(new SqlParameter("@UserCode", UserCode));

            if (!string.IsNullOrEmpty(Model_No))
            {
                command.Parameters.Add(new SqlParameter("@Model_No", Model_No));
                
            }

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_FCST");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_FCST"];
        }

        private void button23_Click(object sender, EventArgs e)
        {
            if (dataGridView11.Rows.Count <= 0)
            {
                return;
            }

            DataGridViewRow dgRow = dataGridView11.CurrentRow;

            DataTable dtV = MakeVerticalTable();

            DataRow dr ;

            for (int i = 0; i <= dataGridView11.Columns.Count - 1; i++)
            {
                dr = dtV.NewRow();

                dr["Code"] = dataGridView11.Columns[i].DataPropertyName;
                dr["CodeName"] = dgRow.Cells[i].Value;

                dtV.Rows.Add(dr);
            }

            fmShowSpec aForm = new fmShowSpec();
            aForm.dt = dtV;
            aForm.ShowDialog();
           
        }

        private void dataGridView11_DoubleClick(object sender, EventArgs e)
        {
            button23_Click(sender, e);
        }

        private void aCME_STAGEBindingSource_PositionChanged(object sender, EventArgs e)
        {

            DataRowView dgv = (DataRowView)(aCME_STAGEBindingSource.Current);
           //MessageBox.Show(Convert.ToString(dgv["StageNo"]));

            //Bug 20090305
            if (dgv == null)
            {
                return;
            }
            string StageNo =Convert.ToString(dgv["StageNo"]);
            //自已取明細檔的資料
            aCME_STAGE_DTableAdapter.FillByKey(cRM.ACME_STAGE_D,
                                               StageNo,
                                               globals.UserID);
                                               
        }

        private void exBindingNavigator3_AfterNew(object sender, EventArgs e)
        {
            DataRowView dgv = (DataRowView)(aCME_STAGEBindingSource.Current);
            string StageNo = Convert.ToString(dgv["StageNo"]);

            //int MaxRow = aCME_STAGE_DBindingSource.Count ;
            //string StepNo = MaxRow.ToString("00");
            //這個沒用

            ((DataRowView)aCME_STAGE_DBindingSource.Current).Row["UserCode"] = globals.UserID;
            ((DataRowView)aCME_STAGE_DBindingSource.Current).Row["StageNo"] = StageNo;
            ((DataRowView)aCME_STAGE_DBindingSource.Current).Row["StepNo"] = GetMaxSeqNo(2);
            ((DataRowView)aCME_STAGE_DBindingSource.Current).Row["DocNum"] = GetKeyString();
//            ((DataRowView)aCME_STAGE_DBindingSource.Current).Row["StartDate"] = DateTime.Now.ToString("yyyyMMdd");

            //為了  Text 可以顯示
            startDateTextBox1.Text = DateTime.Now.ToString("yyyyMMdd");
        }

        private void aCME_STAGE_DDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch
            { 
            
            }
        }

        private void aCME_STAGE_DDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            //DataRowView dgv = (DataRowView)(aCME_STAGEBindingSource.Current);
            //string StageNo = Convert.ToString(dgv["StageNo"]);

            //e.Row.Cells["StepNo"].Value = GetSeqNo(2);
            //e.Row.Cells["UserCode"].Value = globals.UserID;
            //e.Row.Cells["StageNo"].Value = StageNo;

        }


        //序號
        private string GetSeqNo(int length)
        {

            int iRecs;
            iRecs = aCME_STAGE_DDataGridView.Rows.Count;

            string zeroLen = string.Empty;



            string s = "0000000000" + Convert.ToString(iRecs);

            return s.Substring(s.Length - length, length);
        }

        private string GetMaxSeqNo(int length)
        {

            int iRecs;
            iRecs = aCME_STAGE_DDataGridView.Rows.Count;

            string zeroLen = string.Empty;

            string s = "0000000000" + Convert.ToString(iRecs);


            if (iRecs <= 1)
            {
                return s.Substring(s.Length - length, length);
            }
            else
            {
                DataView dv = cRM.Tables["ACME_STAGE_D"].DefaultView;

                dv.Sort = "StepNo Desc ";

                iRecs  = Convert.ToInt32(dv[0]["StepNo"]) + 1;

                string K = "0000000000" + Convert.ToString(iRecs);

                return K.Substring(s.Length - length, length);
            }

  //           DataView   dv   =   new   DataView(DataTable1,   "ID   =   Max(ID)","",   DataViewRowState.CurrentRows);   
    
     
  //DataView   dv   =   DataTable1.DefaultView;   
  //dv.RowFilter   =   "ID   =   Max(ID)";
        }

        private void aCME_STAGEDataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            

            if (e.RowIndex < 0)
            {
                return;
            }

            //用  DataRowView 而非 DatagridViewRow
            DataGridViewRow dgr = aCME_STAGEDataGridView.Rows[e.RowIndex];

            DataRowView row = (DataRowView)aCME_STAGEDataGridView.Rows[e.RowIndex].DataBoundItem;

           // string PredDate =Convert.ToString(dgr.Cells[FindColumnByFieldName(aCME_STAGEDataGridView,"PredDate")]);
            //有資料 DataBind 的用法 
            string PredDate =Convert.ToString(row["PredDate"]);
            string CloseDate =Convert.ToString(row["CloseDate"]);

            //己結不判斷
            if (string.IsNullOrEmpty(PredDate) || !string.IsNullOrEmpty(CloseDate))
            {
                return;
            }


            //逾期
            if (AcmeDateTimeUtils.StrToDate(PredDate) < DateTime.Today)
            {
                //foreach (DataGridViewCell cell in aCME_STAGEDataGridView.Rows[e.RowIndex].Cells)
                //{
                //    cell.Style.BackColor = Color.Pink;
                //}

               // aCME_STAGEDataGridView.Rows[e.RowIndex].Cells["PredDate"].Style.BackColor = Color.Pink;
                //這樣子就不會閃 ????
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }



            if (AcmeDateTimeUtils.StrToDate(PredDate) <= DateTime.Today.AddDays(Convert.ToDouble(numericUpDown1.Value))
                && AcmeDateTimeUtils.StrToDate(PredDate) >= DateTime.Today)
            {
                dgr.DefaultCellStyle.BackColor = Color.Yellow;

                //整行會閃....
                //foreach (DataGridViewCell cell in aCME_STAGEDataGridView.Rows[e.RowIndex].Cells)
                //{
                //    cell.Style.BackColor = Color.Yellow;
                //}

               // aCME_STAGEDataGridView.Rows[e.RowIndex].Cells["PredDate"].Style.BackColor = Color.Yellow;
            }


           
            
        }

        private string FindColumnByFieldName(DataGridView dg, string FieldName)
        {
            
            string sName="";

            for (int i = 0; i <= dg.Columns.Count - 1; i++)
            {
                if (dg.Columns[i].DataPropertyName == FieldName)
                {

                    sName = dg.Columns[i].Name;
                    break;
                }
            
            }
        
            return sName;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetAcme_FcstList();

            if (LookupValues != null)
            {
                textBox6.Text = Convert.ToString(LookupValues[0]);
                textBox5.Text = Convert.ToString(LookupValues[1]);
                comboBox9.Text = Convert.ToString(LookupValues[2]);

            }
        }
        private object[] GetAcme_FcstList()
        {
            string[] FieldNames = new string[] {  "CardCode", "CardName", "FcstMonth" };

            string[] Captions = new string[] { "客戶編號", "客戶名稱",  "年度" };

            string SqlScript = "SELECT Distinct CardCode,CardName,Substring(FcstMonth,1,4) FcstMonth FROM ACME_FCST WHERE  UserCode='" + globals.UserID + "'";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;
            dialog.LookUpConnection = new SqlConnection(ConnStr);

            dialog.SqlScript = SqlScript;

            // dialog.SourceDataTable = GetOcrdWithLead(EmpID);

            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void 儲存SToolStripButton_Click(object sender, EventArgs e)
        {
            button18_Click(sender, e);
        }

        private void bsAcme_fcst_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void Show(Control control, Control area)
        {
           bool resizableTop;
           bool resizableLeft;
            
            if (control == null)
            {
                throw new ArgumentNullException("control");
            }
          //  SetOwnerItem(control);

            //resizableTop = resizableLeft = false;
            Point location = area.PointToScreen(new Point(area.Left, area.Top + area.Height));
           // Point location = area.PointToClient(new Point(area.Left, area.Top + area.Height));
            //control.Location = location;
            //Rectangle screen = Screen.FromControl(control).WorkingArea;
            //if (location.X + Size.Width > (screen.Left + screen.Width))
            //{
            //    resizableLeft = true;
            //    location.X = (screen.Left + screen.Width) - Size.Width;
            //}
            //if (location.Y + Size.Height > (screen.Top + screen.Height))
            //{
            //    resizableTop = true;
            //    location.Y -= Size.Height + area.Height;
            //}

            //location = control.PointToClient(location);
            //control.Left = location.X;
            //control.Top = location.Y;

            //Show(control, control.ClientRectangle);
        }

        //有變動就觸發
        private void cardCodeTextBox_TextChanged(object sender, EventArgs e)
        {
          //  Show(panel17, cardCodeTextBox.ClientRectangle);
            //Show(panel17, cardCodeTextBox);
        }

        private void button26_Click(object sender, EventArgs e)
        {
          //  panel17.Left = 10;

        }

        private void cardNameTextBox_Validated(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)aCME_STAGEBindingSource.Current;

            if (string.IsNullOrEmpty(cardNameTextBox.Text.Trim()))
            {
                return;
            }

            string CardCode = "";
            if (drv.IsNew || drv.IsEdit)
            {
                if (!CheckBP(EmpID, cardNameTextBox.Text, ref CardCode))
                {
                    cardCodeTextBox.Text = CardCode;
                    cardNameTextBox.Focus();
                    errorProvider1.SetError(cardNameTextBox, "客戶名稱輸入錯誤");
                    return;
                }
                else
                {
                    cardCodeTextBox.Text = CardCode;
                    errorProvider1.SetError(cardNameTextBox, "");
                }
            }


        }

        private  bool  CheckBP(string SlpCode, string CardName ,ref string CardCode)
        {
            SqlConnection connection = new SqlConnection(ConnStr02);
            string sql = "SELECT T0.CardCode FROM OCRD T0 WHERE T0.CardType='C' and  T0.SlpCode=@SlpCode ";

            sql += " and CardName like @CardName";


            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

            command.Parameters.Add(new SqlParameter("@CardName", "%" + CardName + "%"));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OCRD");
            }
            finally
            {
                connection.Close();
            }
            DataTable dt = ds.Tables["OCRD"];

            CardCode = "";

            if (dt.Rows.Count == 1)
            {
                CardCode = Convert.ToString(dt.Rows[0][0]);
                return true;

            }
            else
            {
                return false;
            }
        }


        public DataTable GetItem()
        {
            SqlConnection connection = new SqlConnection(ConnStr02);
            string sql = "SELECT distinct Substring(ItemCode,1,9)  FROM OITM WHERE Substring(ItemCode,1,1)<>'Z' order by  Substring(ItemCode,1,9) ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
//            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OITM");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OITM"];
        }

        private void button26_Click_1(object sender, EventArgs e)
        {
            //TaskbarNotifier
            //編譯不過時,請加入 TaskbarNotifier(在 CRM 目錄)
            TaskbarNotifier taskbarNotifier3 = new TaskbarNotifier();

            //自己加的 Resoucre 會存在 Ressource...
           // taskbarNotifier3.SetBackgroundBitmap(new Bitmap(GetType(),"skin3"), Color.FromArgb(255, 0, 255));
            taskbarNotifier3.SetBackgroundBitmap(pictureBox2.Image, Color.FromArgb(255, 0, 255));

         //   taskbarNotifier3.SetCloseBitmap(new Bitmap(GetType(),"close.bmp"),Color.FromArgb(255,0,255),new Point(280,57));
			taskbarNotifier3.TitleRectangle=new Rectangle(150, 57, 125, 28);
			taskbarNotifier3.ContentRectangle=new Rectangle(75, 92, 215, 55);
		//	taskbarNotifier3.TitleClick+=new EventHandler(TitleClick);
		//	taskbarNotifier3.ContentClick+=new EventHandler(ContentClick);
		//	taskbarNotifier3.CloseClick+=new EventHandler(CloseClick);
            taskbarNotifier3.NormalTitleColor = Color.White;
            //taskbarNotifier3.Left = 0;

            taskbarNotifier3.Show("Dear "+globals.UserID+",",
               string.Format("時間是{0}, 您辛苦了 !",  DateTime.Now.ToShortTimeString()),
                 500, 
                 5000,  //show content 
                 500);
        }

        private void label6_DoubleClick(object sender, EventArgs e)
        {
            button26_Click_1(sender, e);
        }


        //整點報時
        //Bug 同一分鐘啟動兩次

        private DateTime TimeStart = DateTime.Now;
        private void timer1_Tick(object sender, EventArgs e)
        {
           // string AlarmTimeHour = "09";
            //string AlarmTimeMinute = "00";

            //防呆 - MultiThread 控制不會 !!!
            TimeSpan ts = DateTime.Now - TimeStart;

            if (ts.TotalSeconds < 59 )
            {
                
                return;
            }

            //if (AlarmTimeHour == DateTime.Now.Hour.ToString("00") &&
              //  AlarmTimeMinute == DateTime.Now.Minute.ToString("00"))
            //if ( AlarmTimeMinute == DateTime.Now.Minute.ToString("00"))
            //整點報時
            if (DateTime.Now.Minute.ToString("00").Substring(1,1)=="00")
            {

                button26_Click_1(sender, e);
                TimeStart = DateTime.Now;
                //MessageBox.Show(TimeStart.ToString("hhmmss"));
            }

            
         
        }
        //應收帳款(資料來源:Account/CheckPaid)------------------------------------------------------------------------------------------------------
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private decimal sd;
        private decimal se;
        private decimal ssd;
        private decimal sse;

        private void button26_Click_2(object sender, EventArgs e)
        {
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            string 工作天數;
            string 發票號碼;
            decimal 台幣金額;
            decimal 已支付;
            string 美金單價;
            string 收款條件;
            string 發票金額;
            string 文件類型;
            string 客戶代碼;
            string 過帳日期2;
            string 業務;

            System.Data.DataTable dt = GetOrderDataAP(EmpID);

            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt1 = null;
            DataRow dr = null;
            DataRow ds = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {



                單號 = dt.Rows[i]["docentry"].ToString();
                文件類型 = dt.Rows[i]["文件類型"].ToString();
                業務 = dt.Rows[0]["業務"].ToString();
                dt1 = GetOrderDataAP1(單號, 文件類型);

                dr = dtCost.NewRow();
                總類 = dt1.Rows[0]["總類"].ToString();
                過帳日期 = dt1.Rows[0]["過帳日期"].ToString();
                工作天數 = dt1.Rows[0]["工作天數"].ToString();
                發票號碼 = dt1.Rows[0]["發票號碼"].ToString();
                台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                已支付 = Convert.ToDecimal(dt1.Rows[0]["已支付"]);
                收款條件 = dt1.Rows[0]["收款條件"].ToString();
                美金單價 = dt1.Rows[0]["美金單價"].ToString();
                發票金額 = dt1.Rows[0]["發票金額"].ToString();
                客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                過帳日期 = dt1.Rows[0]["過帳日期"].ToString();


                過帳日期2 = dt1.Rows[0]["過帳日期2"].ToString();
                dr["應收總計"] = 總類 + 單號 + 發票號碼 + 工作天數 + 發票金額;

                dr["客戶名稱"] = 客戶;

                dr["來源"] = 總類;
                dr["業務"] = 業務;

                if (總類 == "JE")
                {

                    dr["台幣金額"] = dt.Rows[i]["台幣金額"].ToString();
                }
                else
                {
                    dr["台幣金額"] = 台幣金額;
                }

                dr["收款條件"] = 收款條件;
                dr["客戶代碼"] = 客戶代碼;
                dr["過帳日期"] = 過帳日期2;

                sd = 0;
                for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                {
                    DataRow dd = dt1.Rows[j];


                    string sa = dd["數量"].ToString();
                    string sds = dd["美金單價"].ToString();
                    string sae = dd["稅率"].ToString();

                    if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                    {
                        sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);
                        se = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]) / sd;
                        dr["美金金額"] = sd.ToString("#,##0.0000");
                        dr["匯率"] = se.ToString("#,##0.0000");
                    }
                    else
                    {
                        dr["美金金額"] = 0;
                        dr["匯率"] = 0;
                    }

                    if (dt1.Rows.Count == 1)
                    {
                        dr["品名"] = dd["品名"].ToString();
                        dr["數量"] = dd["數量"].ToString();
                        if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                        {
                            decimal sr = Convert.ToDecimal(dd["美金單價"]);
                            dr["美金單價"] = sr.ToString("#,##0.0000");
                        }
                    }
                    else
                    {

                        if (j == dt1.Rows.Count - 1)
                        {
                            dr["品名"] += dd["品名"].ToString();
                            dr["數量"] += dd["數量"].ToString();
                            if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                dr["美金單價"] += sr.ToString("#,##0.0000");
                            }
                        }
                        else
                        {
                            dr["品名"] += dt1.Rows[j]["品名"].ToString() + "/";
                            dr["數量"] += dt1.Rows[j]["數量"].ToString() + "/";
                            if (!String.IsNullOrEmpty(dd["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dd["美金單價"]);
                                dr["美金單價"] += sr.ToString("#,##0.0000") + "/";
                            }
                        }
                    }
                }

                dtCost.Rows.Add(dr);


            }


            System.Data.DataTable dt2 = GetOrderDataAP2(EmpID);
            System.Data.DataTable dtt = null;
            for (int s = 0; s <= dt2.Rows.Count - 1; s++)
            {
                string 原始 = dt2.Rows[s]["原始"].ToString();
                string 原始號碼 = dt2.Rows[s]["原始號碼"].ToString();
                DataRow du = dt2.Rows[s];
                ds = dtCost.NewRow();
                ds["來源"] = du["總類"].ToString();
                ds["過帳日期"] = du["過帳日期"].ToString();
                ds["客戶代碼"] = du["客戶編號"].ToString();
                ds["客戶名稱"] = du["客戶名稱"].ToString();
                ds["應收總計"] = du["備註"].ToString();

                ds["台幣金額"] = du["台幣金額"];
                ds["收款條件"] = "";
                ds["品名"] = "";
                ds["數量"] = "";
                ds["業務"] = du["業務"].ToString();

                dtt = GetOrderDataAP3(原始號碼, 原始);
                for (int t = 0; t <= dtt.Rows.Count - 1; t++)
                {

                    DataRow dst = dtt.Rows[t];
                    string sds = dst["美金單價"].ToString();


                    if ((!String.IsNullOrEmpty(dst["美金單價"].ToString())))
                    {
                        ssd += Convert.ToDecimal(dst["美金單價"]);
                        sse = Convert.ToDecimal(du["台幣金額2"]) / ssd;
                        ds["美金金額"] = ssd.ToString("#,##0.0000");
                        ds["匯率"] = sse.ToString("#,##0.0000");
                    }
                    else
                    {
                        ds["美金金額"] = 0;
                        ds["匯率"] = 0;
                    }
                    if (dt1.Rows.Count == 1)
                    {

                        if (!String.IsNullOrEmpty(dst["美金單價"].ToString()))
                        {
                            decimal sr = Convert.ToDecimal(dst["美金單價"]);
                            ds["美金單價"] = sr.ToString("#,##0.0000");
                        }
                    }
                    else
                    {

                        if (t == dtt.Rows.Count - 1)
                        {

                            if (!String.IsNullOrEmpty(dst["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dst["美金單價"]);
                                ds["美金單價"] += sr.ToString("#,##0.0000");
                            }
                        }
                        else
                        {

                            if (!String.IsNullOrEmpty(dst["美金單價"].ToString()))
                            {
                                decimal sr = Convert.ToDecimal(dst["美金單價"]);
                                ds["美金單價"] += sr.ToString("#,##0.0000") + "/";
                            }
                        }
                    }
                }
                dtCost.Rows.Add(ds);
            }

            //bindingSource1.DataSource = dtCost;
            dtCost.DefaultView.Sort = "客戶代碼,過帳日期";
            //dataGridView10.DataSource = bindingSource1.DataSource;
            dataGridView10.DataSource = dtCost;
           // label4.Text = dtCost.Compute("Sum(台幣金額)", null).ToString();
        }


        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("來源", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("應收總計", typeof(string));
            dt.Columns.Add("美金金額", typeof(string));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("台幣金額", typeof(Int32));
            dt.Columns.Add("收款條件", typeof(string));
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            return dt;
        }
        private System.Data.DataTable GetOrderDataAP(string SlpCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'AR' 總類,t1.docentry docentry,t1.objtype 文件類型 ,0 台幣金額,t4.slpname 業務  from ojdt t0 INNER JOIN OINV T1 ON T0.TransId = T1.TransId INNER join jdt1 t2 on (t0.transid=t2.transid)  INNER join OCRD t3 on (t1.CARDCODE=t3.CARDCODE) INNER join oslp t4 on (t3.slpCODE=t4.slpCODE) where ");
            sb.Append("    Convert(varchar(8),t0.refdate,112)  between '20071231' and @DocDate2  and  T2.[Account] in ('11420101','11430101') and substring(t1.cardcode,1,1)='0' ");

            //Terry ADD
            sb.Append(" and t3.SlpCode = @SlpCode ");

            
            //if (checkBox1.Checked)
           // {
                sb.Append(" and  T2.[IntrnMatch] =0");
            //}


            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T3.[BALANCE] <> 0");
            //}
            sb.Append(" union all");
            sb.Append(" select 'AR貸項' 總類,t1.docentry docentry,t1.objtype 文件類型, 0 台幣金額,t4.slpname 業務  from ojdt t0 INNER JOIN orin T1 ON T0.TransId = T1.TransId  INNER join jdt1 t2 on (t0.transid=t2.transid)  INNER join OCRD t3 on (t1.CARDCODE=t3.CARDCODE) INNER join oslp t4 on (t3.slpCODE=t4.slpCODE) where ");
            sb.Append("   Convert(varchar(8),t0.refdate,112)  between '20071231' and @DocDate2  and  T2.[Account] in ('11420101','11430101') and substring(t1.cardcode,1,1)='0' ");

            //Terry ADD
            sb.Append(" and t3.SlpCode = @SlpCode ");
            //if (checkBox1.Checked)
           // {
                sb.Append(" and  T2.[IntrnMatch] =0");
            //}

            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T3.[BALANCE] <> 0");
            //}
            sb.Append(" UNION ALL");
            sb.Append(" select 'JE' 總類,t0.TRANSID docentry,t0.objtype 文件類型,cast(t1.credit as int)*-1 台幣金額,t4.slpname 業務 from ojdt  T0 INNER JOIN jdt1 t1 on (t0.transid=t1.transid)  inner JOIN ocrd T3 ON (T1.SHORTNAME=T3.cardcode) INNER join oslp t4 on (t3.slpCODE=t4.slpCODE) where  ");
            sb.Append("    Convert(varchar(8),t0.refdate,112)  between '20071231' and @DocDate2  and ACCOUNT='11420101' AND T0.TRANSTYPE='30' ");

            //Terry ADD
            sb.Append(" and t3.SlpCode = @SlpCode ");

            //if (checkBox1.Checked)
            //{
                sb.Append(" and  T1.[IntrnMatch] =0");
           // }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox13.Text));

            //Terry Add
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOrderDataAP2(string SlpCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" select '收款' 總類,t0.transtype 原始,t0.baseref 原始號碼,Convert(varchar(8),t0.refdate,112)  過帳日期,t0.linememo 備註,shortname 客戶編號,t1.cardname 客戶名稱,cast(debit as int) 台幣金額,cast(debit as int) 台幣金額2,t4.slpname 業務 from jdt1 t0 left join ocrd t1 on(t0.shortname=t1.cardcode) INNER join oslp t4 on (t1.slpCODE=t4.slpCODE) where Account in ('11430101','11420101') and credit = 0 ");
            sb.Append(" and TRANSTYPE='24' and substring(shortname,1,1)='0' ");
            sb.Append(" and t1.SlpCode = @SlpCode ");
            //if (checkBox1.Checked)
            //{
                sb.Append(" and  T0.[IntrnMatch] =0");
            //}
            
            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T1.[BALANCE] <> 0");
            //}


            sb.Append(" union all ");

            sb.Append(" select '收款' 總類,t0.transtype 原始,t0.baseref 原始號碼,Convert(varchar(8),t0.refdate,112)  過帳日期,t0.linememo 備註,shortname 客戶編號,t1.cardname 客戶名稱,cast(credit as int)*-1 台幣金額,cast(credit as int) 台幣金額2,t4.slpname 業務 from jdt1 t0 left join ocrd t1 on(t0.shortname=t1.cardcode) INNER join oslp t4 on (t1.slpCODE=t4.slpCODE) where Account in ('11430101','11420101') and debit = 0 ");
            sb.Append(" and TRANSTYPE='24' and substring(shortname,1,1)='0'");
            sb.Append(" and t1.SlpCode = @SlpCode ");
            //if (checkBox1.Checked)
           // {
                sb.Append(" and  T0.[IntrnMatch] =0");
           // }
            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T1.[BALANCE] <> 0");
            //}
            sb.Append(" union all ");
            sb.Append(" select '付款' 總類,t0.transtype 原始,t0.baseref 原始號碼,Convert(varchar(8),t0.refdate,112)  過帳日期,t0.linememo 備註,shortname 客戶編號,t1.cardname 客戶名稱,cast(debit as int) 台幣金額,cast(debit as int) 台幣金額2,t4.slpname 業務 from jdt1 t0 left join ocrd t1 on(t0.shortname=t1.cardcode) INNER join oslp t4 on (t1.slpCODE=t4.slpCODE) where Account in ('11430101','11420101') and credit = 0 ");
            sb.Append(" and TRANSTYPE='46' and substring(shortname,1,1)='0'");
            sb.Append(" and t1.SlpCode = @SlpCode ");
            //if (checkBox1.Checked)
           // {
                sb.Append(" and  T0.[IntrnMatch] =0");
          //  }
            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T1.[BALANCE] <> 0");
            //}
            sb.Append(" union all ");

            sb.Append(" select '付款' 總類,t0.transtype 原始,t0.baseref 原始號碼,Convert(varchar(8),t0.refdate,112)  過帳日期,t0.linememo 備註,shortname 客戶編號,t1.cardname 客戶名稱,cast(credit as int)*-1 台幣金額,cast(credit as int) 台幣金額2,t4.slpname 業務 from jdt1 t0 left join ocrd t1 on(t0.shortname=t1.cardcode) INNER join oslp t4 on (t1.slpCODE=t4.slpCODE) where Account in ('11430101','11420101') and debit = 0 ");
            sb.Append(" and TRANSTYPE='46' and substring(shortname,1,1)='0'");
            sb.Append(" and t1.SlpCode = @SlpCode ");
            //if (checkBox1.Checked)
           // {
                sb.Append(" and  T0.[IntrnMatch] =0");
           // }
           
            //if (checkBox1.Checked)
            //{
            //    sb.Append(" and  T1.[BALANCE] <> 0");
            //}

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox13.Text));
            command.Parameters.Add(new SqlParameter("@SlpCode", SlpCode));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderDataAP1(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,Convert(varchar(8),t0.docdate,112) 過帳日期2,T0.U_IN_BSINV 發票號碼,");
            sb.Append("               T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CAST(T8.PRICE AS VARCHAR)  美金單價 FROM acmesql02.dbo.OINV T0  ");
            sb.Append("              LEFT JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN acmesql02.dbo.DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN acmesql02.dbo.RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              where t1.basetype='15'");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("              union all");
            sb.Append("              SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,Convert(varchar(8),t0.docdate,112) 過帳日期2,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("              ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("              T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,CAST(T8.PRICE AS VARCHAR)  美金單價 FROM acmesql02.dbo.OINV T0  ");
            sb.Append("              LEFT JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN acmesql02.dbo.RDR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline)");
            sb.Append("              where t1.basetype='17'");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                           SELECT 'AR' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,Convert(varchar(8),t0.docdate,112) 過帳日期2,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                           T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額 ");
            sb.Append("                         ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                         T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,cast(T1.u_acme_inv AS VARCHAR)   美金單價 FROM acmesql02.dbo.OINV T0  ");
            sb.Append("                         LEFT JOIN acmesql02.dbo.INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t1.basetype =-1 and  t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append("                         SELECT 'AR貸項' 總類,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,Convert(varchar(8),t0.docdate,112) 過帳日期2,T0.U_IN_BSINV 發票號碼,");
            sb.Append("                             T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT)*-1 台幣金額 ");
            sb.Append("                           ,cast(T0.PAIDtodate as int) 已支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end 數量,");
            sb.Append("                           T0.COMMENTS 備註,t0.u_acme_pay 收款條件,t1.u_acme_workday 工作天數,t0.u_acme_paygui 發票金額,cast(T1.u_acme_inv AS VARCHAR)   美金單價 FROM acmesql02.dbo.Orin T0  ");
            sb.Append("                           LEFT JOIN acmesql02.dbo.rin1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                         where  t0.docentry=@docentry and t0.objtype=@bb");
            sb.Append("  UNION ALL");
            sb.Append("  select 'JE' 總類,CAST(T0.TRANSID AS VARCHAR) 單號,0 稅率,0 台幣單價,substring(Convert(varchar(10),t0.REFdate,111),6,6) 過帳日期");
            sb.Append(" ,Convert(varchar(8),t0.REFdate,112) 過帳日期2,'' 發票號碼,T1.SHORTNAME 客戶代碼,T2.CARDNAME 客戶名稱,");
            sb.Append("   ''  產品編號,'' 品名,CAST(t1.credit AS INT) 台幣金額");
            sb.Append(" ,0 已支付,0 數量 ,T1.LINEMEMO 備註,'' 收款條件,'' 工作天數,'' 發票金額,'' 美金單價 ");
            sb.Append("  FROM OJDT T0 inner JOIN JDT1 T1 ON (T0.TRANSID=T1.TRANSID)");
            sb.Append("   inner JOIN ocrd T2 ON (T1.SHORTNAME=T2.cardcode)");
            sb.Append(" where T0.TRANSID=@docentry and t0.objtype=@bb and credit <> 0 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderDataAP3(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               select CAST(T0.DOCnum as VARCHAR) 單號,0 稅率,cast(t1.sumapplied as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期");
            sb.Append("               ,Convert(varchar(8),t0.docdate,112) 過帳日期2,");
            sb.Append("               t1.u_usd 美金單價");
            sb.Append("               ,0 數量 ,T0.JRNLMEMO 備註");
            sb.Append("               from oRCT T0 LEFT JOIN RCT2 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append("              WHERE  t0.DOCnum=@docentry and t0.objtype=@bb ");
            sb.Append("               union all ");
            sb.Append("               select CAST(T0.DOCnum as VARCHAR) 單號,0 稅率,cast(t1.sumapplied as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期");
            sb.Append("               ,Convert(varchar(8),t0.docdate,112) 過帳日期2,");
            sb.Append("               t1.u_usd 美金單價");
            sb.Append("               ,0 數量 ,T0.JRNLMEMO 備註");
            sb.Append("               from ovpm T0 LEFT JOIN vpm2 T1 ON (T0.DOCNUM=T1.DOCNUM)");
            sb.Append("              WHERE  t0.DOCnum=@docentry and t0.objtype=@bb ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        
        //end 應收帳款(資料來源:Account/CheckPaid)------------------------------------------------------------------------------------------------------
        private void button27_Click(object sender, EventArgs e)
        {
            fmRptStage aForm = new fmRptStage();

           // aForm.dt = cRM.ACME_STAGE;
            aForm.dt = PrintDataTable;
            aForm.ShowDialog();

            //aReport.SetDataSource(aCME_STAGEBindingSource.DataSource);
            //aReport.
        }
        //整合正航借出

        public DataTable GetChiBorrow(string EmpId_ChiNo)
        {

            string ConnStrChi = "server=acmesrvchi;pwd=CHI;uid=CHI;database=SunSql21";
            SqlConnection connection = new SqlConnection(ConnStrChi);

            StringBuilder sb = new StringBuilder();
            sb.Append(" select  M.經手人員,P.姓名,M.客戶編號,C.公司簡稱,M.日期,M.單號,B.產品編號,IsNull(B.品名規格,'') AS 品名規格, ");
            sb.Append(" IsNull(B.數量,0) AS 數量,B.還貨數量,(B.數量 -B.還貨數量) 未還數量,L.名稱 as 借出類別,M.備註,'' AS 原借單號 ");
            sb.Append("  from stkBorrowMain M");
            sb.Append(" inner join comCustomer C on M.客戶編號=C.編號");
            sb.Append(" inner join comPerson P on M.經手人員=P.編號");
            sb.Append(" inner join  stkBorrowSub B on M.旗標=B.旗標 and  M.單號=B.單號");
            //left join stkBorrowClass L on M.借出類別=L.編號
            sb.Append(" left join stkBorrowClass L on M.借出類別=L.編號");
            sb.Append(" where C.旗標=1 and M.旗標=1  ");
            sb.Append(" and B.還貨數量<B.數量 ");
            sb.Append(" and M.日期 >='20070101' ");
            sb.Append(" and P.編號 >=@EmpId_ChiNo ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@EmpId_ChiNo", EmpId_ChiNo));
            //command.Parameters.Add(new SqlParameter("@UserCode", UserCode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME"];
        }

        private void button28_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.AppStarting;
            try
            {
                // GridViewToExcel(dataGridView1);
                GridViewToCSV(dataGridView7, Environment.CurrentDirectory + @"\" + globals.UserID + DateTime.Now.ToString("yyMMddhhmmss") + ".csv");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            //
            //fmCrmFcst aForm = new fmCrmFcst();

            //aForm.dt = dtGlobalOcrd;
            //aForm.ShowDialog();
        }

        private void exBindingNavigator2_AfterSave(object sender, MyFormStatusEventArgs args)
        {
            if (args.MyFormStatus == "I")
            {
                //重新取得
                dtGlobalOcrd.Clear();
                dtGlobalOcrd = GetOcrdWithLead(EmpID);
               // MessageBox.Show("");
            }
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {

            comboBox12.Items.Clear();
            comboBox12.Text = "";
            string s = Convert.ToString(comboBox11.SelectedItem);
            string sql = "SELECT DISTINCT  Model_No  as  Model_No  FROM AuProduct Where Product_Category='"+s+"' order by Model_No ";
            SetComboBoxList(comboBox12, ConnStr, sql, "Model_No");

        }




        /// <summary>
        /// 這裏可以作 圖型顯示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void aCME_STAGEDataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //{
        //    // Set the background to red for negative values in the Balance column.
        //    //if (dataGridView1.Columns[e.ColumnIndex].Name.Equals("Balance"))
        //    //{
        //    //    Int32 intValue;
        //    //    if (Int32.TryParse((String)e.Value, out intValue) &&
        //    //        (intValue < 0))
        //    //    {
        //    //        e.CellStyle.BackColor = Color.Red;
        //    //        e.CellStyle.SelectionBackColor = Color.DarkRed;
        //    //    }
        //    //}

        //    // Replace string values in the Priority column with images.

        //    //if (dataGridView1.Columns[e.ColumnIndex].Name.Equals("Priority"))
        //    //{
        //    //    // Ensure that the value is a string.
        //    //    String stringValue = e.Value as string;
        //    //    if (stringValue == null) return;

        //    //    // Set the cell ToolTip to the text value.
        //    //    DataGridViewCell cell = dataGridView1[e.ColumnIndex, e.RowIndex];
        //    //    cell.ToolTipText = stringValue;

        //    //    // Replace the string value with the image value.
        //    //    switch (stringValue)
        //    //    {
        //    //        case "high":
        //    //            e.Value = highPriImage;
        //    //            break;
        //    //        case "medium":
        //    //            e.Value = mediumPriImage;
        //    //            break;
        //    //        case "low":
        //    //            e.Value = lowPriImage;
        //    //            break;
        //    //    }
        //    //}


        //    if (e.RowIndex < 0)
        //    {
        //        return;
        //    }

        //    if (aCME_STAGEDataGridView.Columns[e.ColumnIndex].Name.Equals("ColumnFlag"))
        //    {
        //        DataGridViewRow dgr = aCME_STAGEDataGridView.Rows[e.RowIndex];

        //        DataRowView row = (DataRowView)aCME_STAGEDataGridView.Rows[e.RowIndex].DataBoundItem;

        //        // string PredDate =Convert.ToString(dgr.Cells[FindColumnByFieldName(aCME_STAGEDataGridView,"PredDate")]);
        //        //有資料 DataBind 的用法 

        //        string PredDate = Convert.ToString(row["PredDate"]);
        //        string CloseDate = Convert.ToString(row["CloseDate"]);

        //        e.Value = DbimageList.Images[13];
        //        //己結不判斷
        //        if (string.IsNullOrEmpty(PredDate) || !string.IsNullOrEmpty(CloseDate))
        //        {
        //            return;
        //        }

        //        //逾期
        //        if (AcmeDateTimeUtils.StrToDate(PredDate) < DateTime.Today)
        //        {
        //            e.Value = DbimageList.Images[12];
        //        }



        //        if (AcmeDateTimeUtils.StrToDate(PredDate) <= DateTime.Today.AddDays(Convert.ToDouble(numericUpDown1.Value))
        //            && AcmeDateTimeUtils.StrToDate(PredDate) >= DateTime.Today)
        //        {
        //            e.Value = DbimageList.Images[11];
        //        }
        //    }

        //}

        //要移動後才會給值
      //  private void aCME_STAGEDataGridView_CellValidated(object sender, DataGridViewCellEventArgs e)
      //  {

            //if (e.RowIndex < 0)
            //{
            //    return;
            //}

            //DataGridViewRow dgr = aCME_STAGEDataGridView.Rows[e.RowIndex];

            //DataRowView row = (DataRowView)aCME_STAGEDataGridView.Rows[e.RowIndex].DataBoundItem;

            //// string PredDate =Convert.ToString(dgr.Cells[FindColumnByFieldName(aCME_STAGEDataGridView,"PredDate")]);
            ////有資料 DataBind 的用法 

            //string PredDate = Convert.ToString(row["PredDate"]);
            //string CloseDate = Convert.ToString(row["CloseDate"]);

            ////己結不判斷
            //if (string.IsNullOrEmpty(PredDate) || !string.IsNullOrEmpty(CloseDate))
            //{
            //    return;
            //}

            ////逾期
            //if (AcmeDateTimeUtils.StrToDate(PredDate) < DateTime.Today)
            //{
            //    dgr.Cells["ColumnFlag"].Value = "1";
            //}



            //if (AcmeDateTimeUtils.StrToDate(PredDate) <= DateTime.Today.AddDays(Convert.ToDouble(numericUpDown1.Value))
            //    && AcmeDateTimeUtils.StrToDate(PredDate) >= DateTime.Today)
            //{
            //    dgr.Cells["ColumnFlag"].Value = "2";
            //}
       // }

        //銷售預測---------------------------------------------------------------------------------------------

        // SELECT T0.DocNum,T0.[DocDate] 調撥日期,T1.ItemCode, T0.CardCode, T0.CardName,  T1.Dscription, T1.Quantity as 借出數量,
        //(select isnull(sum(W.Quantity),0) from WTR1 W 
        //INNER JOIN OWTR O ON O.DocEntry =W.DocEntry 
        //WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode) 已還數量,
        //T1.Quantity-( select isnull(sum(W.Quantity),0) from WTR1 W 
        //INNER JOIN OWTR O ON O.DocEntry =W.DocEntry 
        //WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode) as 未還數量,
        //預設銷售人員 =(SELECT T9.SlpName FROM OCRD C INNER JOIN OSLP T9 ON T9.SlpCode = C.SlpCode WHERE C.CardCode =T0.CardCode),
        //T2. [SlpName],t0.u_acme_kind1,t0.comments,T0.JRNLMEMO
        //FROM OWTR T0  
        //INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry
        //INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode 
        //WHERE T0.U_ACME_KIND ='1'
        //AND T1.Quantity-( select isnull(sum(W.Quantity),0) from WTR1 W 
        //INNER JOIN OWTR O ON O.DocEntry =W.DocEntry 
        //WHERE O.U_ACME_KIND='2' AND W.U_BASE_DOC=T0.DocNum and w.itemcode=t1.itemcode) >0


    } //---------------------------------------------------------

    public class DataRowComparer : IComparer
    {
        ListSortDirection direction;
        int columnIndex;

        public DataRowComparer(int columnIndex, ListSortDirection direction)
        {
            this.columnIndex = columnIndex;
            this.direction = direction;
        }

        #region IComparer Members

        public int Compare(object x, object y)
        {

            DataRow obj1 = (DataRow)x;
            DataRow obj2 = (DataRow)y;
            return string.Compare(obj1[columnIndex].ToString(), obj2[columnIndex].ToString()) * (direction == ListSortDirection.Ascending ? 1 : -1);
        }
        #endregion
    }
    

}


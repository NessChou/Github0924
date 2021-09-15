using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Reflection;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
namespace ACME
{

    public partial class POTAO : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string FileName = "";
        decimal TAMT;
        decimal TQTY;
        private System.Data.DataTable TempDt;
 
        System.Data.DataTable dtCost = null;
        string GlobalMailContent = "";
        private Int32 iCount = 0;
        Attachment data = null;
        private StreamWriter sw;
        System.Data.DataTable dtGetAcmeStage = null;
        System.Data.DataTable dtGetAcmeStageG = null;
        System.Data.DataTable dtGetAcmeStageJS = null;
        public POTAO()
        {
            InitializeComponent();
        }

        private void gB_POTATOBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
   
            this.Validate();
            this.gB_POTATOBindingSource.EndEdit();
            this.gB_POTATOTableAdapter.Update(this.POTATO.GB_POTATO);
            this.gB_POTATO2BindingSource.EndEdit();
            this.gB_POTATO2TableAdapter.Update(this.POTATO.GB_POTATO2);

            this.gB_FRIENDBindingSource.EndEdit();
            this.gB_FRIENDTableAdapter.Update(this.POTATO.GB_FRIEND);
            MessageBox.Show("存檔成功");

       
        }

        private void POTAO_Load(object sender, EventArgs e)
        {
            string USER = fmLogin.LoginID.ToString().ToUpper();

            if (USER == "LLEYTONCHEN" || USER == "NANCYEWI")
            {
                panel8.Hide();
            }
            gB_POTATOTableAdapter.Connection = globals.Connection;
            gB_FRIENDTableAdapter.Connection = globals.Connection;
            gB_POTATO2TableAdapter.Connection = globals.Connection;

            toolStripTextBox1.Text = GetMenu.DFirst();
            toolStripTextBox2.Text = GetMenu.DLast();

            this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
            this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
            this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
            DELETECC();
            DELETEDD();


            toolStripComboBox2.ComboBox.DataSource = GetOslp1();
            toolStripComboBox2.ComboBox.ValueMember = "DataValue";
            toolStripComboBox2.ComboBox.DisplayMember = "DataValue";

            toolStripComboBox1.Text = "快遞單號";
            if (globals.GroupID.ToString().Trim() == "WH" || globals.GroupID.ToString().Trim() == "GB" || globals.GroupID.ToString().Trim() == "GBT" || globals.GroupID.ToString().Trim() == "EEP")
            {

            }
            else
            {
                gB_POTATOBindingNavigatorSaveItem.Visible = false;

            }


            BILLNO();
        }


        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
       
        public System.Data.DataTable dtcost(string ID)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT delcom 公司,DELADDR 地址,");
            sb.Append(" CASE DELMAN WHEN '同訂購人' THEN  ORDNAME ELSE DELMAN END+' 收' 姓名  ,");
            sb.Append(" 'TEL:'+CASE DELTEL WHEN '同訂購人' THEN  ORDTEL ELSE DELTEL END 電話,'箱數:'+CAST(QTY AS VARCHAR) 箱數 ");
            sb.Append(" FROM GB_POTATO WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["OWTR"];
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {

                if (toolStripTextBox1.Text == "" || toolStripTextBox2.Text == "")
                {
                    MessageBox.Show("請輸入日期");
                }



                    this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                    this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                    this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
                    BILLNO();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= gB_POTATODataGridView.Rows.Count - 1; i++)
            {

                DataGridViewRow row;

                row = gB_POTATODataGridView.Rows[i];
                row.ReadOnly  = true;
            }
        }


    
  


        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = download1G(toolStripTextBox1.Text, toolStripTextBox2.Text,"1");
            TOTAL2GG(dt);
 

        }

    
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("PO#", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("訂購人公司", typeof(string));
            dt.Columns.Add("訂購人EMail", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人公司", typeof(string));
            dt.Columns.Add("全雞", typeof(string));
            dt.Columns.Add("半雞", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("價錢", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交貨地點", typeof(string));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("客戶預訂日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("實際到貨日期", typeof(string));
            dt.Columns.Add("快遞單號", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
            dt.Columns.Add("付款日期", typeof(string));
            dt.Columns.Add("運送時段", typeof(string));
            dt.Columns.Add("備註", typeof(string));
             dt.Columns.Add("產品名稱", typeof(string));
             dt.Columns.Add("產品類別", typeof(string));
             dt.Columns.Add("產品代碼", typeof(string));
             dt.Columns.Add("包裝內容", typeof(string));
            return dt;
        }

        private System.Data.DataTable GetBOM(string FATHER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT QTY,INVNAME FROM GB_BOM T0  LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)  WHERE FATHER=@FATHER ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOITM(string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT INVNAME,ITEMTYPE,ITEMOI FROM GB_OITM WHERE ITEMCODE=@ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable download1G(string CreateDate,string CreateDate2,string B1)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                                          SELECT T0.ID,OrdName 訂購人,OrdTel 訂購人電話,OrdCom 訂購人公司,OrdEMail 訂購人EMail");
            sb.Append("                                          ,T1.SPERSON 收貨人,T1.STEL 收貨人電話,DelCom 收貨人公司,T1.TQTY 全雞,T1.SQTY 半雞,ISNULL(T1.TQTY,0)+ISNULL(T1.SQTY,0)  數量");
            sb.Append("                                          ,T0.PotatoWg 價錢,T0.SHIPFEE 運費,AMOUNT  總計,T1.SADDRESS 交貨地點,CreateDate 訂單日期,T1.SDATE 客戶預訂日期,");
            sb.Append(" T1.DelRemark 取貨日期,T1.Flag1 實際到貨日期,");
            sb.Append(" T1.OrdNo 快遞單號,TransMark 交易方式,PAYMAN 付款人,Flag2 付款日期,T1.STIME 運送時段,T1.MEMO 備註,T1.[NO] NO,T0.ORDERPIN,T0.PROJECT FROM dbo.GB_POTATO T0 ");
            sb.Append("                      INNER JOIN  GB_FRIEND  T1 ON (T0.ID=T1.DOCID) WHERE prodid='True'  ");
            if (B1 == "1")
            {
                sb.Append(" AND CreateDate between @CreateDate and @CreateDate2   ");
            }
            if (B1 == "2")
            {
                sb.Append(" AND T1.DelRemark between @CreateDate and @CreateDate2    ");
            }
            if (B1 == "3")
            {
                sb.Append("           AND  CreateDate > '20130301'  AND isnull(T1.OrdNo,'')  = ''  ");
            }
            sb.Append("  ORDER BY T0.ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateDate2", CreateDate2));

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }


        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                string F = opdf.FileName;

                GetExcelContentGD4(F, toolStripComboBox1.Text);
          

            }
        }
        private void GetExcelContentGD4(string ExcelFile,string T1)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2 = "";
            string id3 = "";

            int u = 0;
            int v = 0;


         
                for (int jj = 1; jj <= 30; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, jj]);
                    range.Select();
                    id = range.Text.ToString();

                    if (id.Trim().ToUpper() == "ID")
                    {

                        u = jj;
                    }

                    if (id.Trim() == T1)
                    {

                        v = jj;
                    }
                    
                }


                if (u == 0 || v == 0)
            {
                MessageBox.Show("Excel格式有誤");
                return;

            }
         

                try
                {


                    for (int j = 2; j <= iRowCnt; j++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, u]);
                        range.Select();
                        id2 = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, v]);
                        range.Select();
                        id3 = range.Text.ToString().Trim();



                        if (!String.IsNullOrEmpty(id2))
                        {

                            UPDATESAP(id3, id2, T1);

                        }


                    }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            



            //Quit
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            range = null;
            excelApp = null;
            excelBook = null;
            excelSheet = null;

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            MessageBox.Show("匯入成功");
        }

        public void UPDATESAP(string OrdNo, string ID,string TYPE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;
            if (TYPE == "快遞單號")
            {

                command = new SqlCommand("UPDATE GB_FRIEND SET OrdNo=@OrdNo  where DOCID=@ID ", connection);
            }

            if (TYPE == "付款日期")
            {

                 command = new SqlCommand("UPDATE GB_POTATO SET Flag2=@OrdNo  where ID=@ID ", connection);
            }
            if (TYPE == "實際到貨日期")
            {

                command = new SqlCommand("UPDATE GB_FRIEND SET Flag1=@OrdNo  where DOCID=@ID ", connection);
            }
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@OrdNo", OrdNo));
            command.Parameters.Add(new SqlParameter("@ID", ID));
        


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


        public void UPDATEFEE(string SFEE, string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;


            command = new SqlCommand("UPDATE GB_FRIEND SET SFEE=@SFEE  where ID=@ID ", connection);
         
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SFEE", SFEE));
            command.Parameters.Add(new SqlParameter("@ID", ID));



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
        public void BILLNO()
        {
            if (gB_POTATODataGridView.Rows.Count > 0)
            {

                for (int i = 0; i <= gB_POTATODataGridView.Rows.Count - 2; i++)
                {
                    DataGridViewRow row;

                    row = gB_POTATODataGridView.Rows[i];
                    string T1 = row.Cells["ID"].Value.ToString();
                    string T2 = row.Cells["PROJECT"].Value.ToString();

                    if (String.IsNullOrEmpty(T2))
                    {
                        System.Data.DataTable G1 = GETBILLNO(T1);
                        if (G1.Rows.Count > 0)
                        {
                            string BILLNO = G1.Rows[0][0].ToString();

                            UPDATBILLNO(BILLNO, T1);
                        }

                    }
                }

            }
        }
        public void UPDATBILLNO(string PROJECT, string ID)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = null;


            command = new SqlCommand("UPDATE GB_POTATO SET PROJECT=@PROJECT WHERE ID=@ID ", connection);

            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@ID", ID));



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

        private void gB_POTATODataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
           

            e.Row.Cells["CreateDate"].Value = GetMenu.Day();
        }




        private void button4_Click(object sender, EventArgs e)
        {

            this.Validate();
            this.gB_POTATOBindingSource.EndEdit();
            this.gB_POTATOTableAdapter.Update(this.POTATO.GB_POTATO);
            if (gB_POTATODataGridView.SelectedRows.Count == 0 )
            {
                MessageBox.Show("請選擇");
                return;
            }

            object[] LookupValues = GetMenu.GetGBOITM("內部");

            if (LookupValues != null)
            {

                int iRecs;

                iRecs = gB_POTATO2DataGridView.Rows.Count;

                System.Data.DataTable dt2 = POTATO.GB_POTATO2;
                string ITEMCODE = Convert.ToString(LookupValues[0]);
                DataRow drw2 = dt2.NewRow();
                string da = gB_POTATODataGridView.SelectedRows[0].Cells["ID"].Value.ToString();
                drw2["ID"] = da;
                drw2["LINE"] = iRecs;
                drw2["ITEMCODE"] = Convert.ToString(LookupValues[0]);
                drw2["ITEMNAME"] = Convert.ToString(LookupValues[1]);
                drw2["Qty"] = 1;
                drw2["UNIT"] =  Convert.ToString(LookupValues[3]);
                drw2["PRICE"] = Convert.ToString(LookupValues[2]);
                drw2["AMOUNT"] = Convert.ToString(LookupValues[2]);
          
                
                //PRICE
                dt2.Rows.Add(drw2);

                this.Validate();
                this.gB_POTATO2BindingSource.EndEdit();
                this.gB_POTATO2TableAdapter.Update(this.POTATO.GB_POTATO2);
            }

           
        }

        private void gB_POTATO2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = gB_POTATO2DataGridView.Rows.Count;
            e.Row.Cells["LINE"].Value = iRecs.ToString();
        }

        private void gB_POTATO2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gB_POTATO2DataGridView.Columns[e.ColumnIndex].Name == "Qty" ||
                          gB_POTATO2DataGridView.Columns[e.ColumnIndex].Name == "PRICE" )
                {

                    decimal  Qty = 0;
                    decimal PRICE = 0;
                    decimal AMT = 0;
                    Qty = Convert.ToDecimal(this.gB_POTATO2DataGridView.Rows[e.RowIndex].Cells["Qty"].Value);
                    PRICE = Convert.ToDecimal(this.gB_POTATO2DataGridView.Rows[e.RowIndex].Cells["PRICE"].Value);
                    AMT = Qty * PRICE;
                    AMT = Math.Round(AMT, 0, MidpointRounding.AwayFromZero);

                    this.gB_POTATO2DataGridView.Rows[e.RowIndex].Cells["AMOUNT"].Value = AMT;
               

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable GetTTT(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(AMOUNT),0) AMOUNT,SUM(QTY*PACK) QTY,SUM(AC*QTY) 全雞,SUM(HC*QTY) 半雞,");
            sb.Append(" AVG(CASE T0.ITEMCODE WHEN 'CB' THEN 0 ELSE T0.PRICE END) 全雞單價,AVG(CASE T0.ITEMCODE WHEN 'CA' THEN 0 ELSE T0.PRICE END) 半雞單價  FROM  dbo.GB_POTATO2 T0");
            sb.Append(" LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" WHERE T0.ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }

        public static System.Data.DataTable GetTTT1(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT *  FROM GB_POTATO2 WHERE SUBSTRING(ITEMCODE,1,1)='G' ");
            sb.Append(" AND ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
        public static System.Data.DataTable GetOPT(string ID)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;

            sb.Append("                                                   SELECT T0.ID,OrdName 訂購人,OrdTel 訂購人電話,OrdCom 訂購人公司,OrdEMail 訂購人EMail");
            sb.Append("                                                        ,T1.SPERSON 收貨人,T1.STEL 收貨人電話,DelCom 收貨人公司");
            sb.Append("                                                        ,T0.PotatoWg 價錢,T0.SHIPFEE 運費,AMOUNT  總計,T1.SADDRESS 交貨地點,CreateDate 訂單日期,T1.SDATE 客戶預訂日期,");
            sb.Append("              CASE TransMark WHEN '月結30days' THEN '薪資扣款' ELSE TransMark END 交易方式,T1.STIME 運送時段,T1.MEMO 備註  FROM dbo.GB_POTATO T0 ");
            sb.Append("                                    INNER JOIN  GB_FRIEND  T1 ON (T0.ID=T1.DOCID) ");
            sb.Append("                  where T0.ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }

        public static System.Data.DataTable download13G(string ID)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;


            sb.Append(" SELECT  T0.ID,QTY,REPLACE(T0.ITEMNAME+'('+RTRIM(ItemDesc)+')','朝貢豬_','')  KG  FROM GB_POTATO2 T0");
            sb.Append("           LEFT JOIN GB_OITM T1 ON(T0.ITEMCODE=T1.ITEMCODE) WHERE T0.ID='" + ID + "'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
       
        public static System.Data.DataTable GetDETAIL(string ID)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;

            sb.Append("                   SELECT * FROM GB_POTATO2 WHERE ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
        public static System.Data.DataTable GetFEE()
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT PARAM_NO FROM RMA_PARAMS WHERE PARAM_KIND='POFEE'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }



        private void DELETECC()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE GB_POTATO2 WHERE ID2 IN (SELECT ID2 FROM GB_POTATO2 T0 LEFT JOIN GB_POTATO T1 ON (T0.ID=T1.ID) WHERE ISNULL(T1.ID,'') = '')");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

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


        private void DELETEDD()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE GB_FRIEND WHERE DOCID IN (SELECT T0.DOCID FROM GB_FRIEND T0 LEFT JOIN GB_POTATO T1 ON (T0.DOCID=T1.ID) WHERE ISNULL(T1.ID,'') = '') ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

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


        private void MailTest2(string strSubject, string SlpName, string MailAddress, string MailContent)
        {
            MailMessage message = new MailMessage();
            string FROM = fmLogin.LoginID.ToString() + "@acmepoint.com";
            message.From = new MailAddress(FROM, "系統發送");
            message.To.Add(new MailAddress(MailAddress));



            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\MailTemplates\\POTATO.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "聿豐實業");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                string m_File = "";

                m_File = file;
                data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                //附件资料
                ContentDisposition disposition = data.ContentDisposition;


                // 加入邮件附件
                message.Attachments.Add(data);

            }

            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";

            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

            try
            {
                client.Send(message);
                data.Dispose();
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        SetMsg(String.Format("Failed to deliver message to {0}",
                            ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                        ex.ToString()));
            }

        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void SetMsg(string Msg)
        {
            label1.Text = "處理訊息:" + Msg;
            label1.Refresh();
            WriteToLog(sw, label1.Text + "\r\n");
        }
        private void WriteToLog(StreamWriter sw, string Msg)
        {
            // StreamWriter sw = new StreamWriter("file.html", true, Encoding.UTF8);//creating html file
            sw.Write(Msg);
            // sw.Close();
        }
        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  查無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }

            //create html & table
            //strB.AppendLine("<html><body><center><table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border-collapse:collapse;'>");
            strB.AppendLine("<tr class='HeaderBorder'>");
            //cteate table header
            for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
            {
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                {

                    //if (i != 0)
                    //{
                    //    strB.AppendLine("<tr class='HeaderBorder'>");
                    //    for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                    //    {
                    //        strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                    //    }
                    //    strB.AppendLine("</tr>");
                    //}

                    //處理鍵值




                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;
                }
                else
                {
                    tmpKeyValue = "";
                }


                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }



                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                }


                for (int d = 1; d <= dg.Rows[i].Cells.Count - 1; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    // HttpUtility.HtmlDecode("&nbsp;&nbsp;&nbsp;")

                    if (dgvc.ValueType == typeof(Int32))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Int32 x = Convert.ToInt32(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }

                    else if (dgvc.ValueType == typeof(Decimal) || dgvc.ValueType == typeof(Double))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Decimal x = Convert.ToDecimal(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0.00") + "</td>");
                        }


                    }
                    else
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {

                            if (dg.Columns[dgvc.ColumnIndex].HeaderText.IndexOf("日期") >= 0)
                            {
                                if (dgvc.Value.ToString() == "0")
                                {
                                    strB.AppendLine("<td>&nbsp;</td>");
                                }
                                else
                                {

                                    string sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                 dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                 dgvc.Value.ToString().Substring(6, 2);


                                    strB.AppendLine("<td>" + sDate + "</td>");
                                }
                            }
                            else
                            {
                                strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                            }
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (toolStripTextBox1.Text == "" || toolStripTextBox2.Text == "")
            {
                MessageBox.Show("請輸入日期");
            }
            this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
            this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
            this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
        }

     
        private void gB_POTATODataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
            
            if (e.RowIndex >= gB_POTATODataGridView.Rows.Count-1)
                return;
            DataGridViewRow dgr = gB_POTATODataGridView.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["ProdID"].Value.ToString() != "True")
                {

                    dgr.DefaultCellStyle.BackColor = Color.Yellow;
                }
                if (dgr.Cells["RIVAMSG"].Value.ToString() == "刷卡失敗")
                {

                    dgr.DefaultCellStyle.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 

        }

        private void gB_POTATODataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }


        private void gB_POTATODataGridView_SelectionChanged(object sender, EventArgs e)
        {
            if (gB_POTATODataGridView.SelectedRows.Count > 0)
            {
              
                string ID = gB_POTATODataGridView.SelectedRows[0].Cells["ID"].Value.ToString();

                label2.Text = "ID : " + ID;
                label3.Text = "ID : " + ID;
             
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_POTATOBindingSource.EndEdit();
            this.gB_POTATOTableAdapter.Update(this.POTATO.GB_POTATO);
            if (gB_POTATODataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇");
                return;
            }

            object[] LookupValues = GetMenu.GetGBOITM("外部");

            if (LookupValues != null)
            {

                int iRecs;

                iRecs = gB_POTATO2DataGridView.Rows.Count;

                System.Data.DataTable dt2 = POTATO.GB_POTATO2;
                string ITEMCODE = Convert.ToString(LookupValues[0]);
                DataRow drw2 = dt2.NewRow();
                string da = gB_POTATODataGridView.SelectedRows[0].Cells["ID"].Value.ToString();
                drw2["ID"] = da;
                drw2["LINE"] = iRecs;
                drw2["ITEMCODE"] = Convert.ToString(LookupValues[0]);
                drw2["ITEMNAME"] = Convert.ToString(LookupValues[1]);
                drw2["Qty"] = 1;

                drw2["PRICE"] = Convert.ToString(LookupValues[2]);
                drw2["AMOUNT"] = Convert.ToString(LookupValues[2]);
                drw2["UNIT"] = Convert.ToString(LookupValues[3]);

                //PRICE
                dt2.Rows.Add(drw2);


                this.Validate();
                this.gB_POTATO2BindingSource.EndEdit();
                this.gB_POTATO2TableAdapter.Update(this.POTATO.GB_POTATO2);
            }
        }

        private void gB_FRIENDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            
            e.Row.Cells["NO"].Value = util.GetSeqNo(2, gB_FRIENDDataGridView);
            e.Row.Cells["MEMO"].Value = "到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！";
            e.Row.Cells["SHIPCOMPANY"].Value = "大榮";
        }




        public static System.Data.DataTable download1()
        {
            System.Data.DataTable T1 = download2();
            string DATE = T1.Rows[0][0].ToString();

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                                                    SELECT T0.ID,OrdName 訂購人,''''+OrdTel 訂購人電話,OrdCom 訂購人公司,OrdEMail 訂購人EMail");
            sb.Append("                                                    ,T1.SPERSON 收貨人,''''+T1.STEL 收貨人電話,DelCom 收貨人公司,T1.TQTY 全雞,T1.SQTY 半雞,T2.KG 公斤,CASE T0.ORDNO WHEN 'True' THEN 0 else ISNULL(T1.TQTY,0)+ISNULL(T1.SQTY,0) end  數量");
            sb.Append("                                                    ,CASE T1.[NO] WHEN 1 THEN T0.PotatoWg ELSE 0 END 價錢,CASE T1.[NO] WHEN 1 THEN T0.SHIPFEE ELSE 0 END 運費,CASE T1.[NO] WHEN 1 THEN AMOUNT ELSE 0 END  總計,T1.SADDRESS 交貨地點,CreateDate 訂單日期,T1.SDATE 客戶預訂日期,T1.DelRemark 取貨日期,T1.Flag1 實際到貨日期,");
            sb.Append("                                                    T1.OrdNo 快遞單號,TransMark 交易方式,PAYMAN 付款人,Flag2 付款日期,UNIT 服務單位,T1.STIME 運送時段,T1.MEMO 備註 FROM dbo.GB_POTATO T0  ");
            sb.Append("                                INNER JOIN  GB_FRIEND  T1 ON (T0.ID=T1.DOCID)");
            sb.Append("   INNER JOIN  (SELECT  MAX(KG) KG,T0.ID  FROM GB_POTATO2 T0");
            sb.Append(" LEFT JOIN GB_OITM T1 ON(T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" GROUP BY T0.ID) T2 ON(T0.ID=T2.ID)");
            sb.Append("                                      WHERE 1=1 and isnull(prodid,'') ='True'  AND CreateDate > '20130301'  AND isnull(T1.OrdNo,'')  = '' ");
            sb.Append("          AND T1.DelRemark <=  '" + DATE + "' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public static System.Data.DataTable download12(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT  T0.ID,QTY,REPLACE(T1.ITEMNAME+'('+RTRIM(ItemDesc)+')','朝貢豬_','')  KG,TYPE  FROM GB_POTATO2 T0");
            sb.Append("           LEFT JOIN GB_OITM T1 ON(T0.ITEMCODE=T1.ITEMCODE) WHERE T0.ID='" + ID + "'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public static System.Data.DataTable download2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     select   Convert(varchar(10),DATEADD(D,-1,MAX(a.date_time)),112)     a  from     ");
            sb.Append("       (   select   top  2 *   From   acmesqlsp.dbo.Y_2004   ");
            sb.Append("           where   IsRestDay   =   0   ");
            sb.Append("           and   Convert(varchar(10),date_time,112)    >=    '" + DateTime.Now.ToString("yyyyMMdd") + "' ");
            sb.Append("           order   by   date_time    )   as a   ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public static System.Data.DataTable download13(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(CASE WHEN ITEMOI='大宗' THEN 6*QTY ELSE QTY END),0) QTY FROM GB_POTATO2 T0 LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE) WHERE T0.ID='" + ID + "' AND  SUBSTRING(T1.itemcode,1,1) <> 'M'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public static System.Data.DataTable download13GG(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(CASE WHEN ITEMOI='大宗' AND [TYPE] ='雞' THEN 6*QTY ELSE QTY END),0) QTY FROM GB_POTATO2 T0 LEFT JOIN GB_OITM T1 ON (T0.ITEMCODE=T1.ITEMCODE) WHERE T0.ID='" + ID + "'   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        

        public static System.Data.DataTable download13GGGG(string CreateDate, string CreateDate2,string AA)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                                        SELECT CASE ISNULL(KG,'') WHEN ''  THEN    T2.ITEMCODE ELSE KG END ITEMCODE, ISNULL(SUM(CASE WHEN ITEMOI='大宗' THEN 6*T2.QTY ELSE T2.QTY END),0)  數量,REPLACE(T3.ITEMNAME,'(大宗)','')  產品  FROM dbo.GB_POTATO T0 ");
            sb.Append("                    INNER JOIN  GB_FRIEND  T1 ON (T0.ID=T1.DOCID)");
            sb.Append("  INNER JOIN  GB_POTATO2  T2 ON (T0.ID=T2.ID)");
            sb.Append("  INNER JOIN  GB_OITM  T3 ON (T2.ITEMCODE=T3.ITEMCODE)");
            sb.Append("  WHERE prodid='True'   AND T0.CREATEDATE BETWEEN @CreateDate AND @CreateDate2  AND ISNULL(T3.ITEMDESC,'') <>  '' AND CASE ISNULL(KG,'') WHEN ''  THEN    T2.ITEMCODE ELSE KG END =@AA ");
            sb.Append(" GROUP BY CASE ISNULL(KG,'') WHEN ''  THEN    T2.ITEMCODE ELSE KG END,REPLACE(T3.ITEMNAME,'(大宗)','')  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateDate2", CreateDate2));
            command.Parameters.Add(new SqlParameter("@AA", AA));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("訂購人公司", typeof(string));
            dt.Columns.Add("訂購人EMail", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人公司", typeof(string));
            dt.Columns.Add("全雞", typeof(string));
            dt.Columns.Add("半雞", typeof(string));
            dt.Columns.Add("公斤", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("價錢", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交貨地點", typeof(string));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("客戶預訂日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("實際到貨日期", typeof(string));
            dt.Columns.Add("快遞單號", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
            dt.Columns.Add("付款日期", typeof(string));
            dt.Columns.Add("服務單位", typeof(string));
            dt.Columns.Add("運送時段", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }
        private void TOTAL2()
        {
            dtCost = MakeTableCombine();

            System.Data.DataTable DT1 = download1();
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string ID = DT1.Rows[i]["ID"].ToString().Trim();
                dr["ID"] = ID;
                dr["訂購人"] = DT1.Rows[i]["訂購人"].ToString().Trim();
                dr["訂購人電話"] = DT1.Rows[i]["訂購人電話"].ToString().Trim();
                dr["訂購人公司"] = DT1.Rows[i]["訂購人公司"].ToString().Trim();
                dr["訂購人EMail"] = DT1.Rows[i]["訂購人EMail"].ToString().Trim();
                dr["收貨人"] = DT1.Rows[i]["收貨人"].ToString().Trim();
                dr["收貨人電話"] = DT1.Rows[i]["收貨人電話"].ToString().Trim();
                dr["收貨人公司"] = DT1.Rows[i]["收貨人公司"].ToString().Trim();

                dr["全雞"] = DT1.Rows[i]["全雞"].ToString().Trim();
                dr["半雞"] = DT1.Rows[i]["半雞"].ToString().Trim();
                dr["數量"] = DT1.Rows[i]["數量"].ToString().Trim();
                dr["價錢"] = DT1.Rows[i]["價錢"].ToString().Trim();
                dr["運費"] = DT1.Rows[i]["運費"].ToString().Trim();
                dr["總計"] = DT1.Rows[i]["總計"].ToString().Trim();
                dr["交貨地點"] = DT1.Rows[i]["交貨地點"].ToString().Trim();
                dr["訂單日期"] = DT1.Rows[i]["訂單日期"].ToString().Trim();
                dr["客戶預訂日期"] = DT1.Rows[i]["客戶預訂日期"].ToString().Trim();
                dr["取貨日期"] = DT1.Rows[i]["取貨日期"].ToString().Trim();
                dr["實際到貨日期"] = DT1.Rows[i]["實際到貨日期"].ToString().Trim();
                dr["快遞單號"] = DT1.Rows[i]["快遞單號"].ToString().Trim();
                dr["交易方式"] = DT1.Rows[i]["交易方式"].ToString().Trim();
                dr["付款人"] = DT1.Rows[i]["付款人"].ToString().Trim();
                dr["付款日期"] = DT1.Rows[i]["付款日期"].ToString().Trim();
                dr["服務單位"] = DT1.Rows[i]["服務單位"].ToString().Trim();
                dr["運送時段"] = DT1.Rows[i]["運送時段"].ToString().Trim();
                dr["備註"] = DT1.Rows[i]["備註"].ToString().Trim();
                StringBuilder sb = new StringBuilder();

                System.Data.DataTable DT = download12(ID);
                if (DT.Rows.Count > 0)
                {
                    if (DT.Rows.Count == 1)
                    {
                        dr["公斤"] = DT1.Rows[i]["公斤"].ToString().Trim();
                    }
                    else
                    {
                        for (int S = 0; S <= DT.Rows.Count - 1; S++)
                        {
                            DataRow dd = DT.Rows[S];
                            string QTY = dd["QTY"].ToString();
                            string KG = dd["KG"].ToString();
                            sb.Append(KG + "*" + QTY + "+");
                        }

                        sb.Remove(sb.Length - 1, 1);
                        dr["公斤"] = sb.ToString();
     
                    }
                }


                dtCost.Rows.Add(dr);
            }

        }
        private void button6_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.gB_POTATOBindingSource.EndEdit();
            this.gB_POTATOTableAdapter.Update(this.POTATO.GB_POTATO);
            if (gB_POTATODataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("請選擇");
                return;
            }

            object[] LookupValues = GetMenu.GetGBOITM("大宗");

            if (LookupValues != null)
            {

                int iRecs;

                iRecs = gB_POTATO2DataGridView.Rows.Count;

                System.Data.DataTable dt2 = POTATO.GB_POTATO2;
                DataRow drw2 = dt2.NewRow();
                string da = gB_POTATODataGridView.SelectedRows[0].Cells["ID"].Value.ToString();
                drw2["ID"] = da;
                drw2["LINE"] = iRecs;
                drw2["ITEMCODE"] = Convert.ToString(LookupValues[0]);
                drw2["ITEMNAME"] = Convert.ToString(LookupValues[1]);
                drw2["QTY"] = Convert.ToString(LookupValues[4]);
                drw2["PRICE"] = Convert.ToString(LookupValues[2]);
                drw2["AMOUNT"] = Convert.ToString(LookupValues[2]);
                drw2["UNIT"] = Convert.ToString(LookupValues[3]);
     
                //PRICE
                dt2.Rows.Add(drw2);

                this.Validate();
                this.gB_POTATO2BindingSource.EndEdit();
                this.gB_POTATO2TableAdapter.Update(this.POTATO.GB_POTATO2);
            }

        }

        private void gB_FRIENDDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (gB_FRIENDDataGridView.Columns[e.ColumnIndex].Name == "SFEE2")
            {

                string SFEE2 = Convert.ToString(this.gB_FRIENDDataGridView.Rows[e.RowIndex].Cells["SFEE2"].Value);
                string ID2 = Convert.ToString(this.gB_FRIENDDataGridView.Rows[e.RowIndex].Cells["ID2"].Value);
                if (SFEE2 != "True")
                {
                    SFEE2 = "False";
                }

                UPDATEFEE(SFEE2, ID2);
            }
        }
        private System.Data.DataTable MakeTableCombineG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("訂購人公司", typeof(string));
            dt.Columns.Add("訂購人EMail", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人公司", typeof(string));
            dt.Columns.Add("訂單明細", typeof(string));
            dt.Columns.Add("價錢", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交貨地點", typeof(string));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("客戶預訂日期", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("運送時段", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombineGG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("訂購人公司", typeof(string));
            dt.Columns.Add("訂購人EMail", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人公司", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("訂單明細", typeof(string));
            dt.Columns.Add("價錢", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交貨地點", typeof(string));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("客戶預訂日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("實際到貨日期", typeof(string));
            dt.Columns.Add("快遞單號", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
            dt.Columns.Add("付款日期", typeof(string));
            dt.Columns.Add("運送時段", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("外部訂單編號", typeof(string));
            dt.Columns.Add("正航單號", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableCombineJS()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("庫存產品", typeof(string));
            dt.Columns.Add("數量", typeof(string));
  
 
            return dt;
        }
        private void TOTAL2GG(System.Data.DataTable dt)
        {
            dtGetAcmeStageG = MakeTableCombineGG();

            System.Data.DataTable DT1 = dt;
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtGetAcmeStageG.NewRow();
                string ID = DT1.Rows[i]["ID"].ToString().Trim();
                dr["ID"] = ID;
                dr["訂購人"] = DT1.Rows[i]["訂購人"].ToString().Trim();
                dr["訂購人電話"] = DT1.Rows[i]["訂購人電話"].ToString().Trim();
                dr["訂購人公司"] = DT1.Rows[i]["訂購人公司"].ToString().Trim();
                dr["訂購人EMail"] = DT1.Rows[i]["訂購人EMail"].ToString().Trim();
                dr["收貨人"] = DT1.Rows[i]["收貨人"].ToString().Trim();
                dr["收貨人電話"] = DT1.Rows[i]["收貨人電話"].ToString().Trim();
                dr["收貨人公司"] = DT1.Rows[i]["收貨人公司"].ToString().Trim();
                System.Data.DataTable DT3 = download13(ID);
                dr["數量"] = DT3.Rows[0]["QTY"].ToString().Trim();
                dr["價錢"] = DT1.Rows[i]["價錢"].ToString().Trim();
                dr["運費"] = DT1.Rows[i]["運費"].ToString().Trim();
                dr["總計"] = DT1.Rows[i]["總計"].ToString().Trim();
                dr["交貨地點"] = DT1.Rows[i]["交貨地點"].ToString().Trim();
                dr["訂單日期"] = DT1.Rows[i]["訂單日期"].ToString().Trim();
                dr["客戶預訂日期"] = DT1.Rows[i]["客戶預訂日期"].ToString().Trim();
                dr["取貨日期"] = DT1.Rows[i]["取貨日期"].ToString().Trim();
                dr["實際到貨日期"] = DT1.Rows[i]["實際到貨日期"].ToString().Trim();
                dr["快遞單號"] = DT1.Rows[i]["快遞單號"].ToString().Trim();
                dr["交易方式"] = DT1.Rows[i]["交易方式"].ToString().Trim();
                dr["付款人"] = DT1.Rows[i]["付款人"].ToString().Trim();
                dr["付款日期"] = DT1.Rows[i]["付款日期"].ToString().Trim();
                dr["運送時段"] = DT1.Rows[i]["運送時段"].ToString().Trim();
                dr["備註"] = DT1.Rows[i]["備註"].ToString().Trim();
                dr["外部訂單編號"] = DT1.Rows[i]["ORDERPIN"].ToString().Trim();
                dr["正航單號"] = DT1.Rows[i]["PROJECT"].ToString().Trim();
                
                StringBuilder sb = new StringBuilder();

                System.Data.DataTable DT = download12(ID);
                if (DT.Rows.Count > 0)
                {

                    for (int S = 0; S <= DT.Rows.Count - 1; S++)
                    {
                        int G1 = S + 1;
                        DataRow dd = DT.Rows[S];
                        string QTY = dd["QTY"].ToString();
                        string KG = dd["KG"].ToString();
                        string TYPE = dd["TYPE"].ToString();

                        if (TYPE.Trim() == "雞")
                        {
                            sb.Append(KG + "\n");
                        }
                        else
                        {

                            sb.Append(G1.ToString() + "." + KG + "*" + QTY + "\n");
                        }

                        dr["類型"] = dd["TYPE"].ToString();
                    }

                    sb.Remove(sb.Length - 1, 1);
                    dr["訂單明細"] = sb.ToString();

                }


                dtGetAcmeStageG.Rows.Add(dr);
            }
            if (dtGetAcmeStageG.Rows.Count > 0)
            {
                dataGridView1.DataSource = dtGetAcmeStageG;

                ExcelReport.GridViewToExcelPotato(dataGridView1);
            }
        }


        public static System.Data.DataTable GetOslp1()
        {

            SqlConnection con = globals.Connection;
            string sql = "SELECT distinct rtrim(isnull(TransMark,'')) DataValue FROM dbo.GB_POTATO  ORDER BY rtrim(isnull(TransMark,''))";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["oslp"];
        }

        private void toolStripComboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {

            this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
        
        }

       

   
        private void GetExcelProduct2(string ExcelFile,string ID)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string AMT;
                string QTY;
                DataRow dr;
                DataRow drFind;
                int S = 0;
       
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 21]);
                    range.Select();
                    AMT = range.Text.ToString().Trim().Replace(",", "");
           
                 
                    if (SERIAL_NO == ID)
                    {
                        TQTY += Convert.ToDecimal(QTY);
                        TAMT += Convert.ToDecimal(AMT);
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }
        private void GetExcelProduct2S(string ExcelFile, string ID)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string AMT;
                string QTY;
                DataRow dr;
                DataRow drFind;
                int S = 0;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    AMT = range.Text.ToString().Trim().Replace(",", "");


                    if (SERIAL_NO == ID)
                    {
                        TQTY += Convert.ToDecimal(QTY);
                        TAMT += Convert.ToDecimal(AMT);
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }

        private void GetExcelProduct2SSALES(string ExcelFile, string ID)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string AMT;
                string QTY;
                DataRow dr;
                DataRow drFind;
                int S = 0;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    AMT = range.Text.ToString().Trim().Replace(",", "");


                    if (SERIAL_NO == ID)
                    {
                        TQTY += Convert.ToDecimal(QTY);
                        TAMT += Convert.ToDecimal(AMT);
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }
        public System.Data.DataTable GETProdID(string ProdID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT InvoProdName FROM comProduct WHERE ProdID =@ProdID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public System.Data.DataTable GETPERSONID(string PersonID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PersonName FROM comPerson where PersonID=@PersonID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PersonID", PersonID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public System.Data.DataTable GETBILLNO(string CustBillNo)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BillNO   FROM ordBillMain  WHERE CustBillNo=@CustBillNo ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustBillNo", CustBillNo));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        private void button7_Click(object sender, EventArgs e)
        {
            TUANGO("1");

            TAMT = 0;
            TQTY = 0;
            TempDt = MakeTableS();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                GetExcelProduct(FileName);

                for (int ii = 0; ii <= TempDt.Rows.Count - 1; ii++)
                {
                    string ID = TempDt.Rows[ii][0].ToString();
                    string AMT = TempDt.Rows[ii][1].ToString();
                    string QTY = TempDt.Rows[ii][2].ToString();
                    GetExcelProduct3(FileName, ID, AMT, QTY);
                }
                this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
            }
        }
        private void TUANGO(string F1)
        {
            TAMT = 0;
            TQTY = 0;
            TempDt = MakeTableS();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                if (F1 == "1")
                {
                    GetExcelProduct(FileName);
                }

                for (int ii = 0; ii <= TempDt.Rows.Count - 1; ii++)
                {
                    string ID = TempDt.Rows[ii][0].ToString();
                    string AMT = TempDt.Rows[ii][1].ToString();
                    string QTY = TempDt.Rows[ii][2].ToString();
                    GetExcelProduct3(FileName, ID, AMT, QTY);
                }
                this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
            }
        }
        private System.Data.DataTable MakeTableS()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("Mo", typeof(string));
            dt.Columns.Add("AMT", typeof(int));
            dt.Columns.Add("QTY", typeof(int));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Mo"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private void GetExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();
                    range.Select();


                    //如果找不到時才新增
                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {
                        drFind = TempDt.Rows.Find(SERIAL_NO);

                        if (drFind == null)
                        {
                            dr = TempDt.NewRow();
                            TAMT = 0;
                            TQTY = 0;
                            dr["Mo"] = SERIAL_NO;
                            GetExcelProduct2(FileName, SERIAL_NO);
                            dr["AMT"] = TAMT;
                            dr["QTY"] = TQTY;
                            TempDt.Rows.Add(dr);
                        }
                    }

                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }

        private void GetExcelProductS(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();
                    range.Select();


                    //如果找不到時才新增
                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {
                        drFind = TempDt.Rows.Find(SERIAL_NO);

                        if (drFind == null)
                        {
                            dr = TempDt.NewRow();
                            TAMT = 0;
                            TQTY = 0;
                            dr["Mo"] = SERIAL_NO;
                            GetExcelProduct2S(FileName, SERIAL_NO);
                            dr["AMT"] = TAMT;
                            dr["QTY"] = TQTY;
                            TempDt.Rows.Add(dr);
                        }
                    }

                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }


        private void GetExcelProductSSALES(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;

                DataRow dr;

                DataRow drFind;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();
                    range.Select();


                    //如果找不到時才新增
                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {
                        drFind = TempDt.Rows.Find(SERIAL_NO);

                        if (drFind == null)
                        {
                            dr = TempDt.NewRow();
                            TAMT = 0;
                            TQTY = 0;
                            dr["Mo"] = SERIAL_NO;
                            GetExcelProduct2SSALES(FileName, SERIAL_NO);
                            dr["AMT"] = TAMT;
                            dr["QTY"] = TQTY;
                            TempDt.Rows.Add(dr);
                        }
                    }

                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }
        private void GetExcelProduct3(string ExcelFile, string ID, string AMT, string QTY)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string billing_name;
                string payment_method;
                string shipping_name;
                string shipping_tel;
                string shipping_address;
                string shipping_date;
                string ITEMCODE;
                string ITEMNAME = "";
                string sTime = "";
                string MEMO = "";
                string RivaCoupon = "";
                string RivaCoupon2 = "";
                string RivaCoupon3 = "";
                string SHIPFEE = "";
                decimal  QTYS = 0;
                int AMTS = 0;

                int S = 0;
                Int32 AutoNo = 0;
                Double Price = 0;
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    billing_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim().Replace("'", "");
                    System.Data.DataTable KK1 = GETProdID(ITEMCODE);
                    if (KK1.Rows.Count > 0)
                    {
                        ITEMNAME = KK1.Rows[0][0].ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    QTYS = Convert.ToDecimal(range.Text.ToString().Trim().Replace(",", ""));

                    //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    //range.Select();
                    //UNIT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    SHIPFEE = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 21]);
                    range.Select();
                    AMTS = Convert.ToInt16(range.Text.ToString().Trim().Replace(",", ""));

                    Price = Convert.ToDouble( Convert.ToDouble(AMTS) / Convert.ToDouble(QTYS));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    shipping_date = range.Text.ToString().Trim().Replace("/", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 27]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    shipping_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    shipping_address = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    shipping_tel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 23]);
                    range.Select();
                    RivaCoupon = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 24]);
                    range.Select();
                    RivaCoupon2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 36]);
                    range.Select();
                    payment_method = range.Text.ToString().Trim();

                    int G1 = MEMO.IndexOf("上午");
                    int G2 = MEMO.IndexOf("下午");
                    int G3 = MEMO.IndexOf("晚上");
                    if (G1 != -1)
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    else if  (G2 != -1)
                    {
                        sTime = "下午(12~17小時)";
                    }
                    else if (G3 != -1)
                    {
                        sTime = "晚上(17~20小時)";
                    }
                    else
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    if (!String.IsNullOrEmpty(RivaCoupon))
                    {

                        RivaCoupon3 = RivaCoupon + "-" + RivaCoupon2;
                    }
                    //int H1 = payment_method.IndexOf("信用卡");
                    //if (G1 != -1)
                    //{
                    //    payment_method = "信用卡付款";
                    //}
                    payment_method = "半月結15天";
                    string CRDATE = DateTime.Now.ToString("yyyyMMdd");
                    string CRETIME = DateTime.Now.ToString("HHmmss");
                    double RivaDiscount = 0;
                    Int32 DetailAmt = 0;

                    try
                    {
                        DetailAmt = Convert.ToInt32(AMT);
                    }
                    catch
                    {
                    }

                    if (SERIAL_NO == ID)
                    {
                        if (S == 0)
                        {
                             AMT = (Convert.ToInt32(AMT) + Convert.ToInt32(SHIPFEE)).ToString();
                            AutoNo = AddGB_POTATOIEMAIN("九易宇軒", "", "", AMT, billing_name, CRDATE, CRETIME, billing_name, CRDATE, CRETIME, DetailAmt.ToString(),SHIPFEE, SERIAL_NO, payment_method, "", "", shipping_name, shipping_tel, shipping_address,
                                 "1", "", AMT, Convert.ToInt16(QTY), RivaCoupon3, RivaDiscount, MEMO, "APP", "", "", "", SERIAL_NO,"");

                            AddFRIEND(AutoNo, 1, shipping_name, shipping_tel, "大榮", shipping_address, Convert.ToInt16(QTY), 0, shipping_date, sTime, "到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！", "");
                        }
                        AddGB_POTATOIEITEM(AutoNo, S, ITEMCODE, ITEMNAME, Price, QTYS, AMTS, "", "箱", ""  );
                        S++;
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }
        public void AddFRIEND(int DOCID, int NO, string SPERSON, string STEL, string COMPANY, string SADDRESS, int TQTY, int SQTY, string SDATE, string STIME, string MEMO, string ORDERPIN)
        {


            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_FRIEND(DOCID,NO,SPERSON,STEL,COMPANY,SADDRESS,TQTY,SQTY,SDATE,STIME,MEMO,ORDERPIN,SHIPCOMPANY) values(@DOCID,@NO,@SPERSON,@STEL,@COMPANY,@SADDRESS,@TQTY,@SQTY,@SDATE,@STIME,@MEMO,@ORDERPIN,'大榮')", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCID", DOCID));
            command.Parameters.Add(new SqlParameter("@NO", NO));
            command.Parameters.Add(new SqlParameter("@SPERSON", SPERSON));
            command.Parameters.Add(new SqlParameter("@STEL", STEL));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            command.Parameters.Add(new SqlParameter("@SADDRESS", SADDRESS));
            command.Parameters.Add(new SqlParameter("@TQTY", TQTY));
            command.Parameters.Add(new SqlParameter("@SQTY", SQTY));
            command.Parameters.Add(new SqlParameter("@SDATE", SDATE));
            command.Parameters.Add(new SqlParameter("@STIME", STIME));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@ORDERPIN", ORDERPIN));



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
        public Int32 AddGB_POTATOIEMAIN(string OrdName, string OrdTel, string OrdCom, string Amount, string CreateUser, string CreateDate, string CreateTime, string UpdateUser, string UpdateDate, string UpdateTime, string PotatoWg, string SHIPFEE, string ORDERPIN, string TRANSMARK, string UNIT, string OrdEMail, string DelMan, string DelTel, string DelAddr, string RivaMode, string RivaMsg, string RivaTotal, Int32 Qty, string RivaCoupon, double RivaDiscount, string RivaNote, string CUSTTYPE, string PROJECT, string SALES, string SALESNAME, string CUSTNO, string PotatoKind)
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_POTATO(OrdName,OrdTel,OrdCom,Amount,CreateUser,CreateDate,CreateTime,UpdateUser,UpdateDate,UpdateTime,PotatoWg,SHIPFEE,ORDERPIN,TRANSMARK,UNIT,OrdEMail,DelMan,DelTel,DelAddr,RivaMode,RivaMsg,RivaTotal,Qty,RivaCoupon,RivaDiscount,RivaNote,CUSTTYPE,PROJECT,SALES,SALESNAME,CUSTNO,PotatoKind) values (@OrdName,@OrdTel,@OrdCom,@Amount,@CreateUser,@CreateDate,@CreateTime,@UpdateUser,@UpdateDate,@UpdateTime,@PotatoWg,@SHIPFEE,@ORDERPIN,@TRANSMARK,@UNIT,@OrdEMail,@DelMan,@DelTel,@DelAddr,@RivaMode,@RivaMsg,@RivaTotal,@Qty,@RivaCoupon,@RivaDiscount,@RivaNote,@CUSTTYPE,@PROJECT,@SALES,@SALESNAME,@CUSTNO,@PotatoKind);SELECT @@IDENTITY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OrdName", OrdName));
            command.Parameters.Add(new SqlParameter("@OrdTel", OrdTel));
            command.Parameters.Add(new SqlParameter("@OrdCom", OrdCom));

            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@CreateUser", CreateUser));
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateTime", CreateTime));
            command.Parameters.Add(new SqlParameter("@UpdateUser", UpdateUser));
            command.Parameters.Add(new SqlParameter("@UpdateDate", UpdateDate));
            command.Parameters.Add(new SqlParameter("@UpdateTime", UpdateTime));
            command.Parameters.Add(new SqlParameter("@PotatoWg", PotatoWg));
            command.Parameters.Add(new SqlParameter("@SHIPFEE", SHIPFEE));
            command.Parameters.Add(new SqlParameter("@ORDERPIN", ORDERPIN));
            command.Parameters.Add(new SqlParameter("@TRANSMARK", TRANSMARK));
            command.Parameters.Add(new SqlParameter("@UNIT", UNIT));

            command.Parameters.Add(new SqlParameter("@OrdEMail", OrdEMail));
            command.Parameters.Add(new SqlParameter("@DelMan", DelMan));
            command.Parameters.Add(new SqlParameter("@DelTel", DelTel));
            command.Parameters.Add(new SqlParameter("@DelAddr", DelAddr));


            //20140903 RivaMode
            command.Parameters.Add(new SqlParameter("@RivaMode", RivaMode));

            //20140912
            //,RivaMsg,RivaTotal
            command.Parameters.Add(new SqlParameter("@RivaMsg", RivaMsg));
            command.Parameters.Add(new SqlParameter("@RivaTotal", RivaTotal));

            command.Parameters.Add(new SqlParameter("@Qty", Qty));

            command.Parameters.Add(new SqlParameter("@RivaCoupon", RivaCoupon));
            command.Parameters.Add(new SqlParameter("@RivaDiscount", RivaDiscount));

            command.Parameters.Add(new SqlParameter("@RivaNote", RivaNote));
            command.Parameters.Add(new SqlParameter("@CUSTTYPE", CUSTTYPE));
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@SALESNAME", SALESNAME));
            command.Parameters.Add(new SqlParameter("@CUSTNO", CUSTNO));
            command.Parameters.Add(new SqlParameter("@PotatoKind", PotatoKind));
            //PotatoKind
            Int32 AutoNo = 0;
            try
            {
                connection.Open();

                AutoNo = Convert.ToInt32(command.ExecuteScalar());
            }
            finally
            {
                connection.Close();
            }
            return AutoNo;

        }

        public void AddGB_POTATOIEITEM(int ID, int LINE, string ITEMCODE, string ITEMNAME, double PRICE,Decimal Qty, int AMOUNT, string ORDERPIN, string PackUnit, string ItemRemark)
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_POTATO2(ID,LINE,ITEMCODE,ITEMNAME,PRICE,Qty,AMOUNT,ORDERPIN,Unit,ItemRemark) values(@ID,@LINE,@ITEMCODE,@ITEMNAME,@PRICE,@Qty,@AMOUNT,@ORDERPIN,@Unit,@ItemRemark)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@LINE", LINE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@Qty", Qty));
            command.Parameters.Add(new SqlParameter("@AMOUNT", AMOUNT));
            command.Parameters.Add(new SqlParameter("@ORDERPIN", ORDERPIN));
            command.Parameters.Add(new SqlParameter("@Unit", PackUnit));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    ////MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            TAMT = 0;
            TQTY = 0;
            TempDt = MakeTableS();

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog2.FileName;
                GetExcelProductS(FileName);

                for (int ii = 0; ii <= TempDt.Rows.Count - 1; ii++)
                {
                    string ID = TempDt.Rows[ii][0].ToString();
                    string AMT = TempDt.Rows[ii][1].ToString();
                    string QTY = TempDt.Rows[ii][2].ToString();
                    GetExcelProduct3S(FileName, ID, AMT, QTY);
                }
                this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
            }
        }
        private void GetExcelProduct3S(string ExcelFile, string ID, string AMT, string QTY)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string billing_name;
                string payment_method;
                string shipping_name;
                string ord_tel;
                string shipping_tel;
                string shipping_address;
                string shipping_date;
                string ITEMCODE;
                string ITEMNAME = "";
                string delivery_time;
                string sTime = "";
                string UNIT;
                string SOURCE;
                decimal QTYS = 0;
                int AMTS = 0;
                string ItemRemark;
                int S = 0;
                string VAT = "";
                Int32 AutoNo = 0;
                Double Price = 0;
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    SOURCE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    billing_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    ord_tel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();
                    System.Data.DataTable KK1 = GETProdID(ITEMCODE);
                    if (KK1.Rows.Count > 0)
                    {
                        ITEMNAME = KK1.Rows[0][0].ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    string G1 = range.Text.ToString().Trim().Replace(",", "");
                    if (G1 == "")
                    {
                        return;
                    }
                    QTYS = Convert.ToDecimal(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    UNIT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    Price = Convert.ToDouble(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    AMTS = Convert.ToInt32(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    shipping_date = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    delivery_time = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    shipping_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    shipping_address = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    shipping_tel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    payment_method = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 19]);
                    range.Select();
                    VAT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 20]);
                    range.Select();
                    ItemRemark = range.Text.ToString().Trim();

                    if (delivery_time == "中午前")
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    else if (delivery_time == "下午")
                    {
                        sTime = "下午(12~17小時)";
                    }
                    else if (delivery_time == "晚上")
                    {
                        sTime = "晚上(17~20小時)";
                    }
                    else
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    string CRDATE = DateTime.Now.ToString("yyyyMMdd");
                    string CRETIME = DateTime.Now.ToString("HHmmss");
                    double RivaDiscount = 0;
                    Int32 DetailAmt = 0;

                    try
                    {
                        DetailAmt = Convert.ToInt32(AMT);
                    }
                    catch
                    {
                    }

                    if (SERIAL_NO == ID)
                    {
                        if (S == 0)
                        {

                            AutoNo = AddGB_POTATOIEMAIN(billing_name,ord_tel, "", AMT, billing_name, CRDATE, CRETIME, billing_name, CRDATE, CRETIME, DetailAmt.ToString(), "0", "", payment_method, VAT, "", shipping_name, shipping_tel, shipping_address,
                                 "1", "", AMT, Convert.ToInt16(QTY), "", RivaDiscount, "", "團購", "", "", "",SERIAL_NO,SOURCE);

                            AddFRIEND(AutoNo, 1, shipping_name, shipping_tel, "大榮", shipping_address, Convert.ToInt16(QTY), 0, shipping_date, sTime, "到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！", "");
                        }
                        AddGB_POTATOIEITEM(AutoNo, S, ITEMCODE, ITEMNAME, Price, QTYS, AMTS, "", UNIT, ItemRemark);
                        S++;
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }
        private void GetExcelProduct3SSALES(string ExcelFile, string ID, string AMT, string QTY)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;
            //Microsoft.Office.Interop.Excel.Range FixedRange = null;




            try
            {
                string SERIAL_NO;
                string billing_name;
                string payment_method;
                string shipping_name;
                string ord_tel;
                string shipping_tel;
                string shipping_address;
                string shipping_date;
                string ITEMCODE;
                string ITEMNAME = "";
                string delivery_time;
                string PROJECT;
                string sTime = "";
                string UNIT;
                string SALES;
                string SALESNAME = "";
                string CUSTNO;
                decimal QTYS = 0;
                int AMTS = 0;
                string ItemRemark;
                int S = 0;
                Int32 AutoNo = 0;
                Double Price = 0;
                string RIVANOTE = "";
                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim().ToUpper();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    billing_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    ord_tel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();
                    System.Data.DataTable KK1 = GETProdID(ITEMCODE);
                    if (KK1.Rows.Count > 0)
                    {
                        ITEMNAME = KK1.Rows[0][0].ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    string G1 = range.Text.ToString().Trim().Replace(",", "");
                    if (G1 == "")
                    {
                        return;
                    }
                    QTYS = Convert.ToInt16(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    UNIT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    Price = Convert.ToInt32(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    AMTS = Convert.ToInt32(range.Text.ToString().Trim().Replace(",", "").Replace(".00", ""));

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    shipping_date = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    delivery_time = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    shipping_name = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    shipping_address = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    shipping_tel = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    payment_method = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    PROJECT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    SALES = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 19]);
                    range.Select();
                    CUSTNO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 20]);
                    range.Select();
                    RIVANOTE = range.Text.ToString().Trim();

                    System.Data.DataTable KK2 = GETPERSONID(SALES);
                    if (KK2.Rows.Count > 0)
                    {
                        SALESNAME = KK2.Rows[0][0].ToString();
                    }



                    ItemRemark = "#" + CUSTNO + "#" + SERIAL_NO + "#" + AMT;
                    if (delivery_time == "中午前")
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    else if (delivery_time == "下午")
                    {
                        sTime = "下午(12~17小時)";
                    }
                    else if (delivery_time == "晚上")
                    {
                        sTime = "晚上(17~20小時)";
                    }
                    else
                    {
                        sTime = "中午前(9~12小時)";
                    }
                    string CRDATE = DateTime.Now.ToString("yyyyMMdd");
                    string CRETIME = DateTime.Now.ToString("HHmmss");
                    double RivaDiscount = 0;
                    Int32 DetailAmt = 0;

                    try
                    {
                        DetailAmt = Convert.ToInt32(AMT);
                    }
                    catch
                    {
                    }

                    if (SERIAL_NO == ID)
                    {
                        if (S == 0)
                        {

                            AutoNo = AddGB_POTATOIEMAIN(billing_name, ord_tel, "", AMT, billing_name, CRDATE, CRETIME, billing_name, CRDATE, CRETIME, DetailAmt.ToString(), "0", "", payment_method, "", "", shipping_name, shipping_tel, shipping_address,
                                 "1", "", AMT, Convert.ToInt16(QTY), "", RivaDiscount, RIVANOTE, "業務", "", SALES, SALESNAME, SERIAL_NO,"");

                            AddFRIEND(AutoNo, 1, shipping_name, shipping_tel, "大榮", shipping_address, Convert.ToInt16(QTY), 0, shipping_date, sTime, "到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！", "");
                        }
                        AddGB_POTATOIEITEM(AutoNo, S, ITEMCODE, ITEMNAME, Price, QTYS, AMTS, "", UNIT, ItemRemark);
                        S++;
                    }
                }

            }
            finally
            {


                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            TAMT = 0;
            TQTY = 0;
            TempDt = MakeTableS();

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog2.FileName;
                GetExcelProductSSALES(FileName);

                for (int ii = 0; ii <= TempDt.Rows.Count - 1; ii++)
                {
                    string ID = TempDt.Rows[ii][0].ToString();
                    string AMT = TempDt.Rows[ii][1].ToString();
                    string QTY = TempDt.Rows[ii][2].ToString();
                    GetExcelProduct3SSALES(FileName, ID, AMT, QTY);
                }
                this.gB_POTATOTableAdapter.Fill(this.POTATO.GB_POTATO, toolStripTextBox1.Text, toolStripTextBox2.Text, toolStripComboBox2.Text, crToolStripTextBox1.Text);
                this.gB_POTATO2TableAdapter.Fill(this.POTATO.GB_POTATO2);
                this.gB_FRIENDTableAdapter.Fill(this.POTATO.GB_FRIEND);
            }
        }
        string HW = "";
        private void button36_Click(object sender, EventArgs e)
        {
            if (HW == "")
            {
                panel8.Show();
                HW = "1";
            }
            else
            {
                panel8.Hide();
                HW = "";
            }
        }



 

    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using CarlosAg.ExcelXmlWriter;

namespace ACME
{
    public partial class POS : Form
    {
        public static string ConnectiongString = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        private string FF;
        string strCn = "Data Source=ACMESRVCHIPOS\\SQLEXPRESS;Initial Catalog=POS0001;Persist Security Info=True;User ID=SYSDBA;Password=masterkey";
        string CHI02 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
      
        string CHI98 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP98;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        System.Data.DataTable TempDt;
        public POS()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EXEC();

      
        }

        private void  EXEC()
        {
            System.Data.DataTable G1=GETPOS(1);
            if (G1.Rows.Count > 0)
            {
                DTPOS1(1);
                DTPOS1(2);
                DTPSUM();

                System.Data.DataTable G2 = GETPOS2();
                string g = G2.Rows[0][1].ToString();
                string gk = G2.Rows[0][2].ToString();
                string gk2 = G2.Rows[0][0].ToString();
                decimal sh = Convert.ToDecimal(g);
                decimal shk = Convert.ToDecimal(gk);
                decimal shk2 = Convert.ToDecimal(gk2);
                label1.Text = "數量:" + sh.ToString("#,##0");
                label2.Text = "金額:" + shk.ToString("#,##0");
                label3.Text = "訂單數:" + shk2.ToString("#,##0");

                DTCOST();
            }

            System.Data.DataTable G5 = GETPOSINV();
            dataGridView5.DataSource = G5;
        }
   
        private void DTCOST()
        {
            System.Data.DataTable dtCost = MakeTableCombine();
          System.Data.DataTable dtD = GETDATE();
          DataRow dr = null;
          System.Data.DataTable dt4 = GETPOS6();
          if (dt4.Rows.Count > 0)
          {
              for (int i3 = 0; i3 <= dt4.Rows.Count - 1; i3++)
              {
                  DataRow dd2 = dt4.Rows[i3];
                  dr = dtCost.NewRow();
                  dr["日期"] = "";
                  dr["星期"] = "總計";
                  dr["付款類型"] = dd2["付款類型"].ToString();
                  dr["銷貨金額"] = dd2["銷貨金額"].ToString();
                  dr["銷退金額"] = dd2["銷退金額"].ToString();
                  dr["作廢金額"] = dd2["作廢金額"].ToString();
                  dr["實收金額"] = dd2["實收金額"].ToString();
                  dr["百分比"] = dd2["百分比"].ToString();
                  dr["訂單數"] = dd2["訂單數"].ToString();
                  dr["平均訂單金額"] = dd2["平均訂單金額"].ToString();
                  //平均訂單金額
                  dtCost.Rows.Add(dr);
              }
          }

          System.Data.DataTable dt5 = GETPOS7();
          if (dt5.Rows.Count > 0)
          {
              DataRow dd4 = dt5.Rows[0];
              dr = dtCost.NewRow();
              dr["日期"] = "";
              dr["星期"] = "總計";
              dr["付款類型"] = dd4["付款類型"].ToString();
              dr["銷貨金額"] = dd4["銷貨金額"].ToString();
              dr["銷退金額"] = dd4["銷退金額"].ToString();
              dr["作廢金額"] = dd4["作廢金額"].ToString();
              dr["實收金額"] = dd4["實收金額"].ToString();
              dr["百分比"] = dd4["百分比"].ToString();
              dr["訂單數"] = dd4["訂單數"].ToString();
              dr["平均訂單金額"] = dd4["平均訂單金額"].ToString();
              dtCost.Rows.Add(dr);
          }
            for (int i2 = 0; i2 <= dtD.Rows.Count - 1; i2++)
            {
                DataRow dd = dtD.Rows[i2];
                string 日期 = dd["日期"].ToString();
                string 日期2 = dd["日期2"].ToString();
                string 星期 = dd["星期"].ToString();
                System.Data.DataTable dt2 = GETPOS4(日期2);
                if (dt2.Rows.Count > 0)
                {
                    for (int i3 = 0; i3 <= dt2.Rows.Count - 1; i3++)
                    {
                        DataRow dd2 = dt2.Rows[i3];
                        dr = dtCost.NewRow();
                        dr["日期"] = dd["日期"].ToString();
                        dr["星期"] = dd["星期"].ToString();
                        dr["付款類型"] = dd2["付款類型"].ToString();
                        dr["銷貨金額"] = dd2["銷貨金額"].ToString();
                        dr["銷退金額"] = dd2["銷退金額"].ToString();
                        dr["作廢金額"] = dd2["作廢金額"].ToString();
                        dr["實收金額"] = dd2["實收金額"].ToString();
                        dr["百分比"] = dd2["百分比"].ToString();
                        dr["訂單數"] = dd2["訂單數"].ToString();
                        dr["平均訂單金額"] = dd2["平均訂單金額"].ToString();
                        dtCost.Rows.Add(dr);
                    }
                }

                System.Data.DataTable dt3 = GETPOS5(日期2);
                if (dt3.Rows.Count > 0)
                {
                    DataRow dd3 = dt3.Rows[0];
                    string AMT = dd3["銷貨金額"].ToString();
                    if (AMT != "0")
                    {
                        dr = dtCost.NewRow();
                        dr["日期"] = dd["日期"].ToString();
                        dr["星期"] = dd["星期"].ToString();
                        dr["付款類型"] = dd3["付款類型"].ToString();


                        dr["銷貨金額"] = dd3["銷貨金額"].ToString();
                        dr["銷退金額"] = dd3["銷退金額"].ToString();
                        dr["作廢金額"] = dd3["作廢金額"].ToString();
                        dr["實收金額"] = dd3["實收金額"].ToString();
                        dr["百分比"] = dd3["百分比"].ToString();
                        dr["訂單數"] = dd3["訂單數"].ToString();
                        dr["平均訂單金額"] = dd3["平均訂單金額"].ToString();
                        dtCost.Rows.Add(dr);
                    }
                }
            }

            if (comboBox1.Text == "刷卡")
            {
                dtCost.DefaultView.RowFilter = " 付款類型='刷卡' ";
            }
            if (comboBox1.Text == "現金")
            {
                dtCost.DefaultView.RowFilter = " 付款類型='現金' ";
            }
            if (comboBox1.Text == "合計")
            {
                dtCost.DefaultView.RowFilter = " 付款類型='合計' ";
            }
            dataGridView3.DataSource = dtCost;

            for (int i = 3; i <= 9; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.DefaultCellStyle.Format = "#,##0";

            }

        }
        private void DTPSUM()
        {

            System.Data.DataTable dtCost = MakeTableCombinePOSUM();

            DataRow dr = null;
            System.Data.DataTable dt4 = GETPOS3();
            if (dt4.Rows.Count > 0)
            {
                for (int i3 = 0; i3 <= dt4.Rows.Count - 1; i3++)
                {
                    DataRow dd2 = dt4.Rows[i3];
                    dr = dtCost.NewRow();

                    string PRODID = dd2["產品編號"].ToString();
                    dr["產品編號"] = PRODID;
                    System.Data.DataTable G1 = GETPRODCLASS(PRODID);
                    if (G1.Rows.Count > 0)
                    {
                        dr["產品類別"] = G1.Rows[0][0].ToString();
                    }
                    dr["產品名稱"] = dd2["產品名稱"].ToString();
                    dr["數量"] = dd2["數量"].ToString();

                    dr["金額"] = dd2["金額"].ToString();

                    dtCost.Rows.Add(dr);
                }
                //G3
                decimal[] TotalG = new decimal[dtCost.Columns.Count - 1];

                for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
                {

                    for (int j = 3; j <= 4; j++)
                    {
                        TotalG[j - 1] += Convert.ToDecimal(dtCost.Rows[i][j]);

                    }
                }

                DataRow rowG;

                rowG = dtCost.NewRow();

                rowG[2] = "合計";

                for (int j = 3; j <= 4; j++)
                {
                    rowG[j] = TotalG[j - 1];

                }

                dtCost.Rows.Add(rowG);

    
            }


            dataGridView2.DataSource = dtCost;




            for (int i = 3; i <= 4; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

        }
        private void DTPOS1(int DTYPE)
        {
           
            System.Data.DataTable dtCost = MakeTableCombinePOS1();

            DataRow dr = null;
            System.Data.DataTable dt4 = GETPOS(DTYPE);
            if (dt4.Rows.Count > 0)
            {
                int QTY = 0;
                for (int i3 = 0; i3 <= dt4.Rows.Count - 1; i3++)
                {
                    DataRow dd2 = dt4.Rows[i3];
                    dr = dtCost.NewRow();
                    dr["日期"] = dd2["日期"].ToString();
                    dr["時間"] = dd2["時間"].ToString();
                    dr["單號"] = dd2["單號"].ToString();
                    dr["客戶編號"] = dd2["客戶編號"].ToString();
                    dr["客戶名稱"] = dd2["客戶名稱"].ToString();
                    dr["統編"] = dd2["統編"].ToString();
                    string PRODID = dd2["產品編號"].ToString();
                    dr["產品編號"] = PRODID;
                    System.Data.DataTable G1 = GETPRODCLASS(PRODID);
                    if (G1.Rows.Count > 0)
                    {
                        dr["產品類別"] = G1.Rows[0][0].ToString();
                    }
                    dr["產品名稱"] = dd2["產品名稱"].ToString();
                    QTY += Convert.ToInt16(dd2["數量"]);
                    dr["數量"] = dd2["數量"].ToString();
                    dr["單價"] = dd2["單價"].ToString();
                    dr["金額"] = dd2["金額"].ToString();
                    dr["機台"] = dd2["機台"].ToString();
                    dr["發票號碼"] = dd2["INVOICE"].ToString();
                    dr["信用卡卡號"] = dd2["VISA_NO"].ToString();
                    dr["電子發票"] = dd2["電子發票"].ToString();
                    dr["付款方式"] = dd2["付款方式"].ToString();
                    dtCost.Rows.Add(dr);
                }

                System.Data.DataTable dt4SUM = GETPOSSUM(DTYPE);
                dr = dtCost.NewRow();
                dr["產品名稱"] = "合計";
                dr["數量"] = QTY.ToString();
                dr["金額"] = Convert.ToDecimal(dt4SUM.Rows[0][0]);
                dtCost.Rows.Add(dr);

                if (DTYPE == 1)
                {
                    dataGridView1.DataSource = dtCost;
                }

                if (DTYPE == 2)
                {
                    dataGridView4.DataSource = dtCost;
                }


                if (DTYPE == 1)
                {
                    for (int i = 9; i <= 11; i++)
                    {
                        DataGridViewColumn col = dataGridView1.Columns[i];
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        col.DefaultCellStyle.Format = "#,##0";
                    }
                }

                if (DTYPE == 2)
                {
                    for (int i = 9; i <= 11; i++)
                    {
                        DataGridViewColumn col = dataGridView4.Columns[i];
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        col.DefaultCellStyle.Format = "#,##0";
                    }
                }
            }

        }
        public System.Data.DataTable GETDATE()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SET LANGUAGE N'Simplified Chinese'  SELECT Convert(varchar(10),DATE_TIME,111) 日期,Convert(varchar(10),DATE_TIME,112) 日期2,DATENAME(Weekday, DATE_TIME) 星期 FROM Y_2004 WHERE  Convert(varchar(10),DATE_TIME,112) BETWEEN @BillDate1 and @BillDate2 AND Convert(varchar(10),DATE_TIME,112)>'20190116' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GETADJPRICE(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(CHI98);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CAST(ROUND(SUM(TotalCost )/SUM(Quantity),0) AS INT) PRICE  FROM comWareAmount     WHERE PRODID=@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        public System.Data.DataTable GETPRODCLASS(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(CHI02);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.ClassName  FROM comProduct T0 ");
            sb.Append(" LEFT JOIN comProductClass T1 ON (T0.ClassID =T1.ClassID)");
            sb.Append(" WHERE T0.ProdID =@ProdID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        public System.Data.DataTable GETCUST1()
        {

            SqlConnection MyConnection = new SqlConnection(CHI02);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PRODID,PRODDESC  FROM DBO.comProduct  WHERE  ProdDesc  IN ('0','1')  ");
 
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GETPRODNAME(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(CHI02);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT PRODNAME FROM comProduct  WHERE ProdID =@ProdID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
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
        public System.Data.DataTable GETADJID()
        {

            SqlConnection MyConnection = new SqlConnection(CHI98);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT ISNULL(MAX(ModAdjNO),0)+1 ID FROM stkModAdjMain  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        public System.Data.DataTable GETADJCLASS(string ClassID)
        {

            SqlConnection MyConnection = new SqlConnection(CHI98);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT AdjustStyle  FROM stkAdjustClass WHERE ClassID =@ClassID  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ClassID", ClassID));
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
        public System.Data.DataTable GETCHI(string NAME)
        {

            SqlConnection MyConnection = new SqlConnection(CHI02);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT ID FROM comCustomer WHERE FULLNAME LIKE  '%" + NAME + "%'  ");
       
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GETCHI2(string TEL)
        {

            SqlConnection MyConnection = new SqlConnection(CHI02);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT ID FROM comCustomer WHERE  REPLACE(Telephone1,'-','')  LIKE '%" + TEL + "%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable GETPOSINV()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
    
            sb.Append(" SELECT  CONVERT(varchar(10),INV_DATE,111)  日期, ");
            sb.Append(" isnull(sum(1),0) 開立, ");
            sb.Append(" isnull(sum(case when DEL_CHK='作廢'  then 1 end),0) 作廢 ");
            sb.Append(" FROM INVO9");
            sb.Append("  WHERE CONVERT(varchar(8),INV_DATE,112)  between @BillDate1 and @BillDate2 ");
            sb.Append(" AND INV_HOW <>'不使用'");
            sb.Append(" group by INV_DATE ");
            sb.Append(" order by INV_DATE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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
        public System.Data.DataTable GETPOS(int FLAG)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                 select CONVERT(varchar(10),TRAN_DATE,111) 日期,time1 時間,T0.BO_NO 單號,T0.CUST_NO 客戶編號,T1.CNAME 客戶名稱,T2.CMP_ID 統編,T0.ITEM_NO 產品編號,T0.[DESC] 產品名稱,QTY 數量,PRICE1 單價,ROUND(AMT,0) 金額,T0.WORK_NO 機台,T0.INVOICE,T2.VISA_NO,T2.INV_HOW 電子發票,CASE WHEN I_AMT <> 0 THEN '現金'ELSE CASE WHEN  ISNULL(VISA_NO,'') = '' THEN    ");
            sb.Append("                            CASE WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER2_NAME))) ,'') = 'LINE PAY' then 'LINE PAY'   ");
            sb.Append("                            WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER3_NAME))) ,'') = '貨到付款' then '貨到付款' ELSE    ");
            sb.Append("                            '現金' END ELSE '刷卡' END END 付款方式         from tran9 T0   ");
            sb.Append("              LEFT JOIN CUSTOMER T1 ON (T0.CUST_NO =T1.CODE)  ");
            sb.Append("              LEFT JOIN INVO9 T2 ON (T0.BO_NO=T2.ACR_NO) ");

            sb.Append("  WHERE CONVERT(varchar(8),TRAN_DATE,112)  between @BillDate1 and @BillDate2 ");
            if (FLAG == 1)
            {
                sb.Append(" AND ISNULL(T0.DEL_CHK,'') = ''");
            }
            if (FLAG == 2)
            {
                sb.Append(" AND ISNULL(T0.DEL_CHK,'') <>''");
            }
              sb.Append(" AND  T0.BO_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')");  

            sb.Append(" ORDER BY CONVERT(varchar(8),TRAN_DATE,112),time1 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GETPOSSUM(int FLAG)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select ISNULL(SUM(ROUND(INV_AMT1,0)),0) 金額    from INVO9   ");
            sb.Append(" WHERE CONVERT(varchar(8),INV_DATE,112)  between @BillDate1 and @BillDate2");
            sb.Append(" AND  ACR_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')");
            if (FLAG == 1)
            {
                sb.Append(" AND ISNULL(DEL_CHK,'') = ''");
            }
            if (FLAG == 2)
            {
                sb.Append(" AND ISNULL(DEL_CHK,'') <>''");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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
        public System.Data.DataTable GETPOS2()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  count(distinct BO_NO) 訂單數,ISNULL(SUM(QTY),0) QTY,SUM(ROUND(AMT,0)) AMT  from tran9");
            sb.Append("  WHERE CONVERT(varchar(8),TRAN_DATE,112)  between @BillDate1 and @BillDate2 AND ISNULL(DEL_CHK,'') ='' ");
            sb.Append("  AND BO_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')");  
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GETPOS3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select T0.ITEM_NO 產品編號,T0.[DESC] 產品名稱,SUM(QTY) 數量,SUM(ROUND(AMT,0)) 金額     from tran9 T0   ");
            sb.Append(" LEFT JOIN INVO9 T2 ON (T0.BO_NO=T2.ACR_NO)  ");
            sb.Append("  WHERE CONVERT(varchar(8),TRAN_DATE,112)  between @BillDate1 and @BillDate2 ");
            sb.Append(" AND ISNULL(T0.DEL_CHK,'') = '' ");
            sb.Append(" AND BO_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')   ");
            sb.Append(" GROUP BY T0.ITEM_NO,T0.[DESC]  ");
            sb.Append(" ORDER BY SUM(QTY) DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GETPOS4(string TRAN_DATE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CASE WHEN  ISNULL(VISA_NO,'') = '' THEN  ");
            sb.Append(" CASE WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER2_NAME))) ,'') = 'LINE PAY' then 'LINE PAY' ");
            sb.Append(" WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER3_NAME))) ,'') = '貨到付款' then '貨到付款' ELSE  ");
            sb.Append(" '現金' END ELSE '刷卡' END 付款類型    ");
            sb.Append(" ,ISNULL(SUM(ROUND(INV_AMT1,0)),0)  銷貨金額,   ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '銷退' THEN ROUND(INV_AMT1,0) ELSE 0  END) 銷退金額,  ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '作廢' THEN ROUND(INV_AMT1,0) ELSE 0  END) 作廢金額,  ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) 實收金額,  ");
            sb.Append(" CAST(CAST(((SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) /  ");
            sb.Append(" (SELECT SUM(ROUND(INV_AMT1,0)) FROM  INVO9  WHERE CONVERT(varchar(8),INV_DATE,112) =@TRAN_DATE  AND ISNULL(DEL_CHK,'') = '' ))*100) AS decimal(10,2)) AS VARCHAR)+'%' 百分比,  ");
            sb.Append(" count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) 訂單數,  ");
            sb.Append(" CASE count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END)  WHEN 0 THEN 0 ELSE CAST(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END)/count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) AS INT) END 平均訂單金額    ");
            sb.Append(" FROM INVO9 WHERE CONVERT(varchar(8),INV_DATE,112) =@TRAN_DATE  AND ACR_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002') ");
            sb.Append(" GROUP BY CASE WHEN  ISNULL(VISA_NO,'') = '' THEN  ");
            sb.Append(" CASE WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER2_NAME))) ,'') = 'LINE PAY'  then 'LINE PAY' ");
            sb.Append(" WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER3_NAME))) ,'') = '貨到付款' then '貨到付款' ");
            sb.Append(" ELSE'現金' END ELSE '刷卡' END ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRAN_DATE", TRAN_DATE));
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

        public System.Data.DataTable GETPOS5(string TRAN_DATE)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT '合計' 付款類型 ");
            sb.Append(" ,ISNULL(SUM(ROUND(INV_AMT1,0)),0)  銷貨金額,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '銷退' THEN ROUND(INV_AMT1,0) ELSE 0  END),0) 銷退金額, ");
            sb.Append(" ISNULL(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '作廢' THEN ROUND(INV_AMT1,0) ELSE 0  END),0) 作廢金額, ");
            sb.Append(" ISNULL(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END),0) 實收金額, ");
            sb.Append(" CAST(CAST(((SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) / ");
            sb.Append(" (SELECT SUM(ROUND(INV_AMT1,0)) FROM  INVO9   WHERE CONVERT(varchar(8),INV_DATE,112) = @TRAN_DATE AND ISNULL(DEL_CHK,'') = '' ))*100) AS decimal(10,2)) AS VARCHAR)+'%' 百分比, ");
            sb.Append(" count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) 訂單數, ");
            sb.Append(" CAST(ISNULL(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END),0)/count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) AS INT) 平均訂單金額 ");
            sb.Append(" FROM INVO9 WHERE CONVERT(varchar(8),INV_DATE,112) = @TRAN_DATE ");
            sb.Append(" AND ACR_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002') 	HAVING ISNULL(SUM(ROUND(INV_AMT1,0)),0) <>0   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRAN_DATE", TRAN_DATE));
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
        public System.Data.DataTable GETPOS6()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CASE WHEN I_AMT <> 0 THEN '現金' ELSE CASE WHEN  ISNULL(T2.VISA_NO,'') = '' THEN   ");
            sb.Append("              CASE WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER2_NAME))) ,'') = 'LINE PAY' then 'LINE PAY' WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER3_NAME))) ,'') = '貨到付款' then '貨到付款'  ELSE   ");
            sb.Append("              '現金' END ELSE '刷卡' END END 付款類型     ");
            sb.Append("              ,ISNULL(SUM(ROUND(INV_AMT1,0)),0)    銷貨金額,      ");
            sb.Append("              SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '銷退' THEN ROUND(INV_AMT1,0) ELSE 0  END) 銷退金額,     ");
            sb.Append("              SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '作廢' THEN ROUND(T2.INV_AMT1,0) ELSE 0  END) 作廢金額,     ");
            sb.Append("              SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) 實收金額,     ");
            sb.Append("              CASE   SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END)  WHEN 0 THEN '0.00%' ELSE   ");
            sb.Append("              CAST(CAST(((SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) /     ");
            sb.Append("              (SELECT SUM(ROUND(INV_AMT1,0)) FROM  INVO9    WHERE CONVERT(varchar(8),INV_DATE,112)   between @BillDate1 and @BillDate2   AND ISNULL(DEL_CHK,'') = '' ))*100) AS decimal(10,2)) AS VARCHAR)+'%' END 百分比,     ");
            sb.Append("              count(distinct CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ACR_NO END) 訂單數 ,    ");
            sb.Append("              CASE   SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END)  WHEN 0 THEN 0 ELSE CAST(SUM(CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) / count(distinct CASE WHEN ISNULL(T2.DEL_CHK,'') = '' THEN ACR_NO END)  AS INT) END 平均訂單金額    ");
            sb.Append("              FROM  INVO9 T2 WHERE CONVERT(varchar(8),INV_DATE,112)   between @BillDate1 and @BillDate2 ");
            sb.Append("              AND T2.ACR_NO  NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')       ");
            sb.Append("              GROUP BY CASE WHEN I_AMT <> 0 THEN '現金' ELSE CASE WHEN  ISNULL(T2.VISA_NO,'') = '' THEN   ");
            sb.Append("              CASE WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER2_NAME))) ,'') = 'LINE PAY'  then 'LINE PAY' WHEN  ISNULL(UPPER(RTRIM(LTRIM(USER3_NAME))) ,'') = '貨到付款' then '貨到付款'  ELSE   ");
            sb.Append("              '現金' END ELSE '刷卡' END  END");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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

        public System.Data.DataTable GETPOS7()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '合計' 付款類型  ");
            sb.Append(" ,ISNULL(SUM(ROUND(INV_AMT1,0)),0)  銷貨金額,   ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '銷退' THEN ROUND(INV_AMT1,0) ELSE 0  END) 銷退金額,  ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '作廢' THEN ROUND(INV_AMT1,0) ELSE 0  END) 作廢金額,  ");
            sb.Append(" SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) 實收金額,  ");
            sb.Append(" CAST(CAST(((SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END) /  ");
            sb.Append(" (SELECT SUM(ROUND(INV_AMT1,0)) FROM  INVO9   WHERE CONVERT(varchar(8),INV_DATE,112)  between @BillDate1 and @BillDate2   AND ISNULL(DEL_CHK,'') = '' ))*100) AS decimal(10,2)) AS VARCHAR)+'%' 百分比,  ");
            sb.Append(" count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) 訂單數 , ");
            sb.Append(" CAST(SUM(CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ROUND(INV_AMT1,0) ELSE 0  END)/count(distinct CASE WHEN ISNULL(DEL_CHK,'') = '' THEN ACR_NO END) AS INT) 平均訂單金額 ");
            sb.Append(" FROM  INVO9 T2 WHERE CONVERT(varchar(8),INV_DATE,112) between @BillDate1 and @BillDate2");
            sb.Append(" AND ACR_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')   ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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
        public System.Data.DataTable GETPOSR()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CONVERT(varchar(8),TRAN_DATE,112) 日期,time1 時間,T0.BO_NO 單號,T0.CUST_NO 客戶編號,CASE WHEN T0.CUST_NO <> '9001' THEN '東門會員' ELSE T1.CNAME END 客戶名稱,T3.NAME 產品類別,T0.ITEM_NO 產品編號,T0.[DESC] 產品名稱,QTY 數量,PRICE1 單價,ROUND(AMT,0) 金額,T0.WORK_NO 機台,T0.INVOICE,CASE T2.VISA_NO WHEN '' THEN '現金' ELSE '刷卡' END 付款類型,T2.INV_HOW 電子發票, T0.DEL_CHK MEMO     from tran9 T0  ");
            sb.Append(" LEFT JOIN CUSTOMER T1 ON (T0.CUST_NO =T1.CODE)  ");
            sb.Append(" LEFT JOIN INVO9 T2 ON (T0.BO_NO=T2.ACR_NO) ");
            sb.Append(" LEFT JOIN PRODCLAS T3 ON (T3.CODE =T0.TYPE) ");
            sb.Append("  WHERE CONVERT(varchar(8),TRAN_DATE,112)  between @BillDate1 and @BillDate2 ");
            sb.Append(" AND ISNULL(T0.DEL_CHK,'') = ''");
            sb.Append(" AND T0.BO_NO NOT IN ('1080116C01001','1080116C01002','1080116C01003','1080116C02001','1080116C02002')");  
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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
        private void POS_Load(object sender, EventArgs e)
        {
            //if (fmLogin.LoginID.ToString().ToUpper() == "LLEYTONCHEN" || fmLogin.LoginID.ToString().ToUpper() == "TIFFANYFANG" || fmLogin.LoginID.ToString().ToUpper() == "SANDYLO" || fmLogin.LoginID.ToString().ToUpper() == "MAXHUNG")
            //{
                button2.Visible = true;
                button5.Visible = true;
                //button6.Visible = true;
                button7.Visible = true;
        //    }
            textBox5.Text = GetMenu.Day();
            textBox6.Text = GetMenu.Day();

            EXEC();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            execuse();
        } 
        private System.Data.DataTable MakeTableCombinePOSUM()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("金額", typeof(decimal));

            return dt;
        }
        private System.Data.DataTable MakeTableCombinePOS1()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("時間", typeof(string));
            dt.Columns.Add("單號", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("統編", typeof(string));
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("單價", typeof(decimal));
            dt.Columns.Add("金額", typeof(decimal));
            dt.Columns.Add("機台", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("信用卡卡號", typeof(string));
            dt.Columns.Add("電子發票", typeof(string));
            dt.Columns.Add("付款方式", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("星期", typeof(string));
            dt.Columns.Add("付款類型", typeof(string));
            dt.Columns.Add("銷貨金額", typeof(decimal));
            dt.Columns.Add("銷退金額", typeof(decimal));
            dt.Columns.Add("作廢金額", typeof(decimal));
            dt.Columns.Add("實收金額", typeof(decimal));
            dt.Columns.Add("百分比", typeof(string));
            dt.Columns.Add("訂單數", typeof(decimal));
            dt.Columns.Add("平均訂單金額", typeof(decimal));
                  return dt;
        }
        private void execuse()
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\POS.xlsx";


                System.Data.DataTable OrderData = GETPOSR();


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportPOS(OrderData, ExcelTemplate, OutPutFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                FF = openFileDialog1.FileName;

                WriteExcelAP2(FF);
                MessageBox.Show("匯入成功");
            }


        }
        private void WriteExcelAP(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {


                string TEL;
                string CARDNAME;
  
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    TEL = range.Text.ToString().Trim().Replace("-", "");

                    System.Data.DataTable H1 = GETCHI(CARDNAME);
                    if (!String.IsNullOrEmpty(CARDNAME))
                    {
                        if (H1.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            range.Value2 = H1.Rows[0][0].ToString();
                        }
                    }

                    if (!String.IsNullOrEmpty(TEL))
                    {
                        System.Data.DataTable H2 = GETCHI2(TEL);
                        if (H2.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                            range.Select();
                            range.Value2 = H2.Rows[0][0].ToString();
                        }
                    }
                 //   AddAP(CARDCODE);
                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FF) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FF);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelAP2(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string ID;
                string TEL;
                string CARDNAME;
                string EMAIL;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    ID = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    TEL = range.Text.ToString().Trim().Replace("-", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    EMAIL = range.Text.ToString().Trim();
                    string SHORT = "";
                    if (!String.IsNullOrEmpty(ID))
                    {
                        if (CARDNAME.Length > 5)
                        {
                            SHORT = CARDNAME.Substring(0, 6);
                        }
                        else
                        {
                            SHORT = CARDNAME;
                            if (CARDNAME == "")
                            {
                                CARDNAME = "東門臨時會員";

                                SHORT = "東門臨時會員";
                            }
                        }
                        UPDATECUST(CARDNAME, SHORT, TEL, EMAIL, ID);
                    }

                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FF) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FF);


                //try
                //{
                //    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //}
                //catch
                //{
                //}
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelALA(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string ITEMCODE;
                string ITEMNAME;
                string QTY;
                string PRICE;
                string REMARK;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    ITEMNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace("-", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    PRICE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    REMARK = range.Text.ToString().Trim();

                    string SHORT = "";
                    if (!String.IsNullOrEmpty(ITEMCODE))
                    {

                        AddOrdBillSub(20210406, iRecord + 1, ITEMCODE, ITEMNAME, Convert.ToInt32(QTY), Convert.ToDouble(PRICE), Convert.ToDouble(QTY) * Convert.ToDouble(PRICE), 0.00, 0, 0, 2, "2021040601", iRecord + 1, "20210406", "", REMARK, "");
                    }
                   
                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FF) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FF);


                //try
                //{
                //    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //}
                //catch
                //{
                //}
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        public void AddOrdBillSub(int BillDate, int SerNO, string ProdID, string ProdName, int Quantity, Double Price, Double Amount, double TaxRate, int TaxAmt, int Discount,
    int Flag, string BillNO, int RowNO, string PreInDate, string ItemRemark, string Detail, string con)
        {
            SqlConnection connection = null;
            connection = new SqlConnection(ConnectiongString);
            string sql = "Insert Into OrdBillSub (BillDate,SerNO,ProdID,ProdName,Quantity,Price,Amount,TaxRate,TaxAmt,Discount,Flag,BillNO,RowNO,sQuantity,sPrice,QtyRemain,PreInDate,ItemRemark,Detail) " +
            "values (@BillDate,@SerNO,@ProdID,@ProdName,@Quantity,@Price,@Amount,@TaxRate,@TaxAmt,@Discount,@Flag,@BillNO,@RowNO,@sQuantity,@sPrice,@QtyRemain,@PreInDate,@ItemRemark,@Detail)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@BillDate", BillDate));
            command.Parameters.Add(new SqlParameter("@SerNO", SerNO));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@ProdName", ProdName));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Price", Price));

            command.Parameters.Add(new SqlParameter("@sQuantity", Quantity));
            command.Parameters.Add(new SqlParameter("@sPrice", Price));

            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@TaxRate", TaxRate));
            command.Parameters.Add(new SqlParameter("@TaxAmt", TaxAmt));
            command.Parameters.Add(new SqlParameter("@Discount", Discount));
            command.Parameters.Add(new SqlParameter("@RowNO", RowNO));


            command.Parameters.Add(new SqlParameter("@Flag", Flag));
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

            //未出數量
            command.Parameters.Add(new SqlParameter("@QtyRemain", Quantity));

            //PreInDate
            command.Parameters.Add(new SqlParameter("@PreInDate", PreInDate));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@Detail", Detail));





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

        private void WriteADJ(string ExcelFile,string DOCDATE)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string ADJTYPE;
                string BU;
                string ADJDATE;
                string PRODID;
                string PRODNAME = "";
                string WH;
                string QTY;
                string memo;
                int ROW = 0;
                    for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                    {
                 
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                        range.Select();
                        ADJTYPE = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                        range.Select();
                        BU = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                        range.Select();
                        ADJDATE = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                        range.Select();
                        PRODID = range.Text.ToString().Trim();
                        System.Data.DataTable G1 = GETPRODNAME(PRODID);
                        if (G1.Rows.Count > 0)
                        {
                            PRODNAME = G1.Rows[0][0].ToString();
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                        range.Select();
                        WH = range.Text.ToString().Trim();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                        range.Select();
                        QTY = range.Text.ToString().Trim().Replace("-", "");

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                        range.Select();
                        memo = range.Text.ToString().Trim();
                        if (DOCDATE == ADJDATE)
                        {
                            ROW++;
                            string ID = GETADJID().Rows[0][0].ToString();
                            double QT = Convert.ToDouble(QTY);
                            double PRICE = Convert.ToDouble(GETADJPRICE(PRODID).Rows[0][0]);
                            int AMT = Convert.ToInt32(QT * PRICE);
                            int CLASS = Convert.ToInt16(GETADJCLASS(ADJTYPE).Rows[0][0]);
                            if (ROW == 1)
                            {
                                ADDADJ(ID, ADJDATE, ADJTYPE, CLASS, "");
                            }
                            ADDADJ2(ID, ROW, PRODID, PRODNAME, WH, Convert.ToInt16(QTY), PRICE, AMT, memo);
                        }
                    
              


                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FF) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FF);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void FAN(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string DOCDATE;
                string QTY;
                string TYPE;
                string ITEMNAME;
                int SEQ = 0;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord,  7]);
                    range.Select();
                    ITEMNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(DOCDATE))
                    {
                        DOCDATE = DOCDATE.Substring(0, 4) + "/" +
                                     DOCDATE.Substring(4, 2) + "/" +
                                     DOCDATE.Substring(6, 2);
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 22]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(".0", "").Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 23]);
                    range.Select();
                    TYPE = range.Text.ToString().Trim();
                    int FTYPE = TYPE.IndexOf("可用庫存");
                    int FTYPE2 = ITEMNAME.IndexOf("保冷袋");
           
                    int Q1= ITEMNAME.IndexOf("餛飩");
                    int Q2 = ITEMNAME.IndexOf("雞精");
                    int Q3 = ITEMNAME.IndexOf("水餃");
                    int Q4 = ITEMNAME.IndexOf("雞湯");
                    int Q5 = ITEMNAME.IndexOf("貢丸");
                    int Q6 = ITEMNAME.IndexOf("雞");
                    int Q7 = ITEMNAME.IndexOf("內臟");
                    int Q8 = ITEMNAME.IndexOf("豬");
                    int Q9 = ITEMNAME.IndexOf("烏魚子");
                    int Q10 = ITEMNAME.IndexOf("滷味");
                    int Q11 = ITEMNAME.IndexOf("肉酥");
                    
                    
                    SEQ = 0;
                    if (Q1 != -1 || Q2 != -1 || Q3 != -1 || Q4 != -1 || Q5 != -1 || Q9 != -1 || Q10 != -1 || Q11 != -1)
                    {
                        SEQ = 4;
                    }
                    else if (Q6 != -1)
                    {
                        SEQ = 3;
                    }
                    else if (Q7 != -1)
                    {
                        SEQ = 2;
                    }
                    else if (Q8 != -1)
                    {
                        SEQ = 1;
                    }
                    int n;
                    if (FTYPE != -1 && FTYPE2==-1)
                    { 
                         if (int.TryParse(QTY, out n))
                         {
                             ADDFAN(ITEMNAME, Convert.ToDateTime(DOCDATE), Convert.ToInt32(QTY), SEQ); 
                         
                         }
                                                

                    }
     

                }




            }
            finally
            {

                try
                {
 
                }
                catch
                {
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("DOCDATE", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["DOCDATE"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private void GetExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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

                    //range = ((Microsoft.Office.Interop.Excel.Range)FixedRange.Cells[iField, iRecord]);
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();
                    range.Select();


                    //如果找不到時才新增
                    drFind = TempDt.Rows.Find(SERIAL_NO);

                    if (drFind == null)
                    {

                        dr = TempDt.NewRow();

                        dr["DOCDATE"] = SERIAL_NO;

                        TempDt.Rows.Add(dr);
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
        private void WritePRICE(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string PROD;
                string IPRICE;
                string DPRICE;
                string CLASS;
                int TAX;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {




                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    PROD = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    IPRICE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    DPRICE = range.Text.ToString().Trim();

                    System.Data.DataTable G1 = GETPRODCLASS(PROD);
                    if (G1.Rows.Count > 0)
                    {
                        CLASS = G1.Rows[0][0].ToString().Substring(0, 3);

                        if (CLASS == "外購品" || CLASS == "加工品")
                        {
                            TAX = 1;
                        }
                        else
                        {
                            TAX = 0;
                        }
                        if (!String.IsNullOrEmpty(PROD))
                        {
                            UPDATEP(PROD, TAX, Convert.ToDecimal(DPRICE), Convert.ToDecimal(IPRICE));
                        }
                    }
                  

                }




            }
            finally
            {

                string NewFileName = Path.GetDirectoryName(FF) + "\\" +
           DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FF);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

   
        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView3.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView3.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["付款類型"].Value.ToString() == "合計")
                {
             
                    dgr.DefaultCellStyle.BackColor = Color.Yellow;
                }
     
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 

        }
        private void UPDATECUST(string FullName, string ShortName, string Telephone1, string Email, string ID)
        {

            SqlConnection connection = new SqlConnection(CHI02);

            StringBuilder sb = new StringBuilder();

            sb.Append("UPDATE  CHICOMP03.DBO.comCustomer SET FullName =@FullName,ShortName=@ShortName,Telephone1=@Telephone1,Email=@Email  WHERE ID=@ID AND FLAG=1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@FullName", FullName));
            command.Parameters.Add(new SqlParameter("@ShortName", ShortName));
            command.Parameters.Add(new SqlParameter("@Telephone1", Telephone1));
            command.Parameters.Add(new SqlParameter("@Email", Email));
            command.Parameters.Add(new SqlParameter("@ID", ID));
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
        private void UPPOSCUST(string POSDISFLAG, string CODE)
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();

            sb.Append("UPDATE  PRODUCT SET POSDISFLAG=@POSDISFLAG WHERE CODE=@CODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@POSDISFLAG", POSDISFLAG));
            command.Parameters.Add(new SqlParameter("@CODE", CODE));

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
        private void ADDFAN(string ITEMNAME, DateTime DOCDATE, int QTY, int SEQ)
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_FAN(ITEMNAME,DOCDATE,QTY,SEQ) values(@ITEMNAME,@DOCDATE,@QTY,@SEQ)", connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@SEQ", SEQ));

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
        private void TRUNFAN()
        {

            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("TRUNCATE TABLE GB_FAN", connection);
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
        private void ADDADJ(string ModAdjNO, string ModAdjName, string AdjustType, int AdjustStyle, string Remark)
        {

            SqlConnection connection = new SqlConnection(CHI98);
            SqlCommand command = new SqlCommand("Insert into stkModAdjMain(ModAdjNO,ModAdjName,AdjustType,AdjustStyle,Remark,MergeOutState) values(@ModAdjNO,@ModAdjName,@AdjustType,@AdjustStyle,@Remark,0)", connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ModAdjNO", ModAdjNO));
            command.Parameters.Add(new SqlParameter("@ModAdjName", ModAdjName));
            command.Parameters.Add(new SqlParameter("@AdjustType", AdjustType));
            command.Parameters.Add(new SqlParameter("@AdjustStyle", AdjustStyle));
            command.Parameters.Add(new SqlParameter("@Remark", Remark));
   
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
        private void ADDADJ2(string ModAdjNO, int SerNo, string ProdID, string ProdName, string WareHouseID, int Quantity, double Price, int Amount, string ItemRemark)
        {

            SqlConnection connection = new SqlConnection(CHI98);
            SqlCommand command = new SqlCommand("Insert into stkModAdjSub(ModAdjNO,SerNo,ProdID,ProdName,WareHouseID,Quantity,Price,Amount,ItemRemark,RowNO,EQuantity,EUnitID,EUnitRelation) values(@ModAdjNO,@SerNo,@ProdID,@ProdName,@WareHouseID,@Quantity,@Price,@Amount,@ItemRemark,@RowNO,@EQuantity,@EUnitID,0)", connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ModAdjNO", ModAdjNO));
            command.Parameters.Add(new SqlParameter("@SerNo", SerNo));
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            command.Parameters.Add(new SqlParameter("@ProdName", ProdName));
            command.Parameters.Add(new SqlParameter("@WareHouseID", WareHouseID));
            command.Parameters.Add(new SqlParameter("@Quantity", Quantity));
            command.Parameters.Add(new SqlParameter("@Price", Price));
            command.Parameters.Add(new SqlParameter("@Amount", Amount));
            command.Parameters.Add(new SqlParameter("@ItemRemark", ItemRemark));
            command.Parameters.Add(new SqlParameter("@RowNO", SerNo));
            command.Parameters.Add(new SqlParameter("@EQuantity", Quantity));
            command.Parameters.Add(new SqlParameter("@EUnitID", ""));

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
        private void UPDATEP(string PRODID, int PriceOfTax, decimal SalesPriceE, decimal StdPrice)
        {

            SqlConnection connection = new SqlConnection(CHI02);

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  CHICOMP02.DBO.comProduct SET PriceOfTax =@PriceOfTax,SalesPriceE =@SalesPriceE WHERE PRODID=@PRODID");

            sb.Append(" UPDATE  CHICOMP03.DBO.comProduct SET PriceOfTax =@PriceOfTax,StdPrice =@StdPrice WHERE PRODID=@PRODID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@PRODID", PRODID));
            command.Parameters.Add(new SqlParameter("@PriceOfTax", PriceOfTax));
            command.Parameters.Add(new SqlParameter("@SalesPriceE", SalesPriceE));
            command.Parameters.Add(new SqlParameter("@StdPrice", StdPrice));

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

        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex  == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }

            if (tabControl1.SelectedIndex ==3)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }

            if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                FF = openFileDialog1.FileName;

                WritePRICE(FF);
                MessageBox.Show("匯入成功");
            }
        }



        private void button7_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                TempDt = MakeTable();
                FF = openFileDialog1.FileName;
                GetExcelProduct(FF);
                for (int i = 0; i <= TempDt.Rows.Count - 1; i++)
                {
                    string DOCDATE = TempDt.Rows[i][0].ToString();
                    WriteADJ(FF, DOCDATE);
                }
                MessageBox.Show("匯入成功");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                TRUNFAN();
                FF = openFileDialog1.FileName;

                FAN(FF);
                FANS();
                
                //string dd = "庫存分析" + GetMenu.Day() + ".xls";
                string dd = "庫存分析.xls";
                string OutPutFile = "//acmesrv01//ACMEGB_Public//02_生管暨會員經營部//01_生管//29. 逢泰庫存表//" + dd;
                WH2(dataGridView6, OutPutFile);
                MessageBox.Show("匯入成功");
            }

        }
        private System.Data.DataTable MakeTableFAN()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("品名", typeof(string));

            System.Data.DataTable G1 = GETFAN();
            if (G1.Rows.Count > 0)
            {
                for (int i = 0; i <= G1.Rows.Count-1; i++)
                {
                    dt.Columns.Add(G1.Rows[i][1].ToString(), typeof(string));
                }

            }



            dt.Columns.Add("總計", typeof(string));



            return dt;
        }

        public System.Data.DataTable GETFAN()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CAST(max(month(DOCDATE)) AS VARCHAR)+'月' 月,SUBSTRING(convert(varchar, DOCDATE, 112),1,6) DOCDATE FROM GB_FAN");
            sb.Append(" GROUP BY  SUBSTRING(convert(varchar, DOCDATE, 112),1,6)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        public System.Data.DataTable GETFAN2(string DOCDATE, string ITEMNAME)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY) 數量 FROM GB_FAN WHERE SUBSTRING(convert(varchar, DOCDATE, 112),1,6)=@DOCDATE AND ITEMNAME=@ITEMNAME GROUP BY ITEMNAME");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
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
        public System.Data.DataTable GETFAN3( string ITEMNAME)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY) 數量 FROM GB_FAN WHERE  ITEMNAME=@ITEMNAME GROUP BY ITEMNAME");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
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
        public System.Data.DataTable GETFAN3()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT ITEMNAME,SEQ FROM GB_FAN ORDER BY SEQ ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        private void FANS()
        {
            System.Data.DataTable dtCost = MakeTableFAN();

            System.Data.DataTable dtD2 = GETFAN3();
            DataRow dr = null;
            if (dtD2.Rows.Count > 0)
            {
                for (int i3 = 0; i3 <= dtD2.Rows.Count - 1; i3++)
                {

                    DataRow dd2 = dtD2.Rows[i3];
                    dr = dtCost.NewRow();
                    string ITEMNAME = dd2["ITEMNAME"].ToString();
                    dr["品名"] = ITEMNAME;
                    System.Data.DataTable dtD = GETFAN();
                    for (int i = 0; i <= dtD.Rows.Count - 1; i++)
                    {
                        string DOCDATE = dtD.Rows[i]["DOCDATE"].ToString();
 
                        System.Data.DataTable dtD3 = GETFAN2(DOCDATE,ITEMNAME);
                        if (dtD3.Rows.Count > 0)
                        {
                            dr[DOCDATE] = dtD3.Rows[0]["數量"].ToString();
                        }

                    }
                    System.Data.DataTable dtD4 = GETFAN3(ITEMNAME);
                    if (dtD4.Rows.Count > 0)
                    {
                        dr["總計"] = dtD4.Rows[0]["數量"].ToString();
                    }
                    dtCost.Rows.Add(dr);
                }
            }

            dataGridView6.DataSource = dtCost;

        }


        private void WH2(DataGridView DGV, string OutPutFile)
        {
            CarlosAg.ExcelXmlWriter.Workbook book = new CarlosAg.ExcelXmlWriter.Workbook();
            WorksheetStyle headerStyle = book.Styles.Add("headerStyleID");
            headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            headerStyle.Alignment.WrapText = true;
            headerStyle.Interior.Color = "#284775";
            headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;
            headerStyle.Font.Color = "white";
            headerStyle.Font.Bold = true;

            WorksheetStyle defaultStyle = book.Styles.Add("workbookStyleID");
            defaultStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            defaultStyle.Alignment.WrapText = true;
            defaultStyle.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");
            defaultStyle.Font.Size = 10;

          

            CarlosAg.ExcelXmlWriter.Worksheet sheet = book.Worksheets.Add("庫存分析");
            WorksheetRow headerRow = sheet.Table.Rows.Add();
            for (int i = 0; i <= DGV.Columns.Count-1; i++)
            {
                string HEAD = "";
                if (i == 0 || i == DGV.Columns.Count-1)
                {
                    HEAD = DGV.Columns[i].HeaderText;
                }
                else
                {
                    int F2 = Convert.ToInt16(DGV.Columns[i].HeaderText.Substring(4, 2));

                    HEAD = F2.ToString() + "月";
                }
                headerRow.Cells.Add(HEAD, DataType.String, "headerStyleID");
            }

            for (int i = 0; i < DGV.Rows.Count-1; i++)
            {

                DataGridViewRow row = DGV.Rows[i];
                WorksheetRow rowS = sheet.Table.Rows.Add();

                for (int j = 0; j < row.Cells.Count; j++)
                {
                   
                        DataGridViewCell cell = row.Cells[j];
                        rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                        rowS.Table.DefaultColumnWidth = 80;
                        //if (j == 0)
                        //{
                        //    rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                        //    rowS.Table.DefaultColumnWidth = 80;
                        //}
                        //else
                        //{
                        //    rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID");
                        //    rowS.Table.DefaultColumnWidth = 80;
                        //}


                        rowS.AutoFitHeight = true;

                    
                }

            }
            book.Save(OutPutFile);

        }

        private void button6_Click(object sender, EventArgs e)
        {
                      System.Data.DataTable dt4 = GETCUST1();
                      if (dt4.Rows.Count > 0)
                      {
                          for (int i3 = 0; i3 <= dt4.Rows.Count - 1; i3++)
                          {
                              DataRow dd2 = dt4.Rows[i3];

                              string PRODID = dd2["PRODID"].ToString();
                              string PRODDESC = dd2["PRODDESC"].ToString();
                              UPPOSCUST(PRODDESC, PRODID);

                          }

                          MessageBox.Show("更新成功");
                      }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                FF = openFileDialog1.FileName;

                WriteExcelALA(FF);
                MessageBox.Show("匯入成功");
            }
        }

    }
}

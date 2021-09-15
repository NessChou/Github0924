using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
namespace ACME
{
    public partial class fmAcmeFCHI : Form
    {

        decimal StockCostS = 0;
        public fmAcmeFCHI()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            StockCostS = 0;
            string DocDate = textBoxDocDate.Text;

            System.Data.DataTable dt = null;
            System.Data.DataTable dt2 = null;

            dt = GetStockListToAge(DocDate, "Qty");

            dt2 = GetStockListToAge(DocDate, "Amt");
            

            dataGridView1.DataSource = dt;

            dataGridView2.DataSource = dt2;


            for (int i = 5; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0.0";
            }

     

            //�[�J�@���X�p
            decimal[] Total = new decimal[dt.Columns.Count - 1];

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                for (int j = 5; j <= dt.Columns.Count - 1; j++)
                {
                    Total[j - 1] += Convert.ToDecimal(dt.Rows[i][j]);

                }
            }

            DataRow row;

            row = dt.NewRow();

            row[2] = "�X�p";
            row[3] = ":";
            for (int j = 5; j <= dt.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dt.Rows.Add(row);




            //�[�J�@���X�p
            Int32[] Total2 = new Int32[dt2.Columns.Count - 1];

            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {


                for (int j = 5; j <= dt2.Columns.Count - 1; j++)
                {
                    Total2[j - 1] += Convert.ToInt32(dt2.Rows[i][j]);

                }
            }


            decimal h1 = 0;
            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {


                h1 += Convert.ToDecimal(dt2.Rows[i][5]);


            }

            DataRow row2;

            row2 = dt2.NewRow();

            row2[2] = "�X�p";
            row2[3] = ":";
            for (int j = 5; j <= dt2.Columns.Count - 1; j++)
            {
                if (j == 5)
                {
                    row2[5] = h1;

                }
                else if (j == 13)
                {
                    row2[j] = StockCostS;
                }
                else
                {
                    row2[j] = Total2[j - 1];
                }

            }
            dt2.Rows.Add(row2);


            for (int i = 5; i <= dataGridView2.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];
                if (i == 5)
                {


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0.0";
                }
        
                else
                {

                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0";
                }
            }

            System.Data.DataTable dtCost2 = MakeTable_StockF();
            string �s��;
            string �s��2;
            string ���~�s��;
            DataRow dr22 = null;
            for (int l = 0; l <= dt.Rows.Count - 1; l++)
            {
                DataRow drFind;

                DataRow dz = dt.Rows[l];
                �s�� = dz["�s��"].ToString();
                �s��2 = dz["�s��2"].ToString();
                ���~�s�� = dz["���~�s��"].ToString();
                drFind = dtCost2.Rows.Find(���~�s��);

                if (drFind == null)
                {
                    dr22 = dtCost2.NewRow();
            

                    dr22["�s��"] = �s��;
                    dr22["�s��2"] = �s��2;
                    dr22["���~�s��"] = ���~�s��;

                    dr22["0-30"] = dt.Compute("Sum([0-30])", "���~�s��='" + ���~�s�� + "'");
                    dr22["31-60"] = dt.Compute("Sum([31-60])", "���~�s��='" + ���~�s�� + "'");
                    dr22["61-90"] = dt.Compute("Sum([61-90])", "���~�s��='" + ���~�s�� + "'");
                    dr22["91-120"] = dt.Compute("Sum([91-120])", "���~�s��='" + ���~�s�� + "'");
                    dr22["121-180"] = dt.Compute("Sum([121-180])", "���~�s��='" + ���~�s�� + "'");
                    dr22["181-360"] = dt.Compute("Sum([181-360])", "���~�s��='" + ���~�s�� + "'");
                    dr22["360�H�W"] = dt.Compute("Sum([360�H�W])", "���~�s��='" + ���~�s�� + "'");
                    dr22["�p�p"] = dt.Compute("Sum([�p�p])", "���~�s��='" + ���~�s�� + "'");
                    dtCost2.Rows.Add(dr22);
                }
            }

            dataGridView3.DataSource = dtCost2;

            for (int i = 3; i <= dataGridView3.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0.0";
            }

            System.Data.DataTable dtCost3 = MakeTable_Stock2F();
            string F�s��;
            string F�s��2;
            string F���~�s��;
            DataRow Fdr22 = null;
            for (int l = 0; l <= dt2.Rows.Count - 1; l++)
            {
                DataRow drFind;

                DataRow dz = dt2.Rows[l];
                F�s�� = dz["�s��"].ToString();
                F�s��2 = dz["�s��2"].ToString();
                F���~�s�� = dz["���~�s��"].ToString();
            
                drFind = dtCost3.Rows.Find(F���~�s��);

                if (drFind == null)
                {
                    Fdr22 = dtCost3.NewRow();


                    Fdr22["�s��"] = F�s��;
                    Fdr22["�s��2"] = F�s��2;
                    Fdr22["���~�s��"] = F���~�s��;
                    Fdr22["�ƶq"] = dt2.Compute("Sum([�ƶq])", "���~�s��='" + F���~�s�� + "'"); ;
                    Fdr22["0-30"] = dt2.Compute("Sum([0-30])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["31-60"] = dt2.Compute("Sum([31-60])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["61-90"] = dt2.Compute("Sum([61-90])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["91-120"] = dt2.Compute("Sum([91-120])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["121-180"] = dt2.Compute("Sum([121-180])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["181-360"] = dt2.Compute("Sum([181-360])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["360�H�W"] = dt2.Compute("Sum([360�H�W])", "���~�s��='" + F���~�s�� + "'");
                    Fdr22["�p�p"] = dt2.Compute("Sum([�p�p])", "���~�s��='" + F���~�s�� + "'");
                    dtCost3.Rows.Add(Fdr22);
                }
            }

            dataGridView4.DataSource = dtCost3;

            for (int i = 3; i <= dataGridView4.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView4.Columns[i];
                if (i == 3)
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0.0";
                }
                else
                {
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    col.DefaultCellStyle.Format = "#,##0";
                }
            }

        }
        private System.Data.DataTable GetStockListToAge( string DocDate, string Mode)
        {




            System.Data.DataTable dt = GetStockList(DocDate);
       
   
            System.Data.DataTable dtStock = null;

            if (Mode == "Qty")
            {
                dtStock = MakeTable_Stock2();
            }
            else
            {
                dtStock = MakeTable_Stock();
            }

            DataRow row;
            DataRow rowFind;
            string ItemCode;
            string WH;
            Int32 StockDays;
            decimal  Qty;

            object[] objKeys = new object[2];
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                ItemCode = Convert.ToString(dt.Rows[i]["���~�s��"]);
          
                StockDays = Convert.ToInt32(dt.Rows[i]["�w�s�Ѽ�"]);

                WH = Convert.ToString(dt.Rows[i]["�ܮw"]);
                Qty = Convert.ToDecimal(dt.Rows[i]["�w�s�q"]);
            
                objKeys[0] = Convert.ToString(dt.Rows[i]["���~�s��"]);
                objKeys[1] = Convert.ToString(dt.Rows[i]["�ܮw"]);

                row = dtStock.Rows.Find(objKeys);


                if (row != null)
                {
                    row.BeginEdit();

                    GoEdit(row, StockDays, Qty, "U");

                    row.EndEdit();
                }
                else
                {

                    row = dtStock.NewRow();
                    row["�s��"] = Convert.ToString(dt.Rows[i]["�s��"]);
                    string ITEM = Convert.ToString(dt.Rows[i]["�s��2"]);
                    row["�s��2"] = ITEM.Replace("-", "");
                    row["���~�s��"] = Convert.ToString(dt.Rows[i]["���~�s��"]);
                    row["�ܮw"] = Convert.ToString(dt.Rows[i]["�ܮw"]);
                    row["�ܮw�W��"] = Convert.ToString(dt.Rows[i]["�ܮw�W��"]);
                    GoEdit(row, StockDays, Qty, "A");


                    dtStock.Rows.Add(row);
                }



            }
         

            //�p�G�O���B -> �ƶq * ���ئ���

            if (Mode == "Qty")
            {
                return dtStock;
            }

            decimal StockCost = 0;
  
            for (int i = 0; i <= dtStock.Rows.Count - 1; i++)
            {
                ItemCode = Convert.ToString(dtStock.Rows[i]["���~�s��"]);
                //���o Cost
           
 
                StockCost = GetItemCost(ItemCode);

                row = dtStock.Rows[i];

                row.BeginEdit();
                row["�ƶq"] = Convert.ToDecimal(row["0-30"]) +
                Convert.ToDecimal(row["31-60"]) +
                Convert.ToDecimal(row["61-90"]) +
                Convert.ToDecimal(row["91-120"]) +
                Convert.ToDecimal(row["121-180"]) +
                Convert.ToDecimal(row["181-360"]) +
                Convert.ToDecimal(row["360�H�W"]);
         
                row["0-30"] = Convert.ToDecimal(row["0-30"]) * StockCost;
                row["31-60"] = Convert.ToDecimal(row["31-60"]) * StockCost;
                row["61-90"] = Convert.ToDecimal(row["61-90"]) * StockCost;
                row["91-120"] = Convert.ToDecimal(row["91-120"]) * StockCost;
                row["121-180"] = Convert.ToDecimal(row["121-180"]) * StockCost;
                row["181-360"] = Convert.ToDecimal(row["181-360"]) * StockCost;
                row["360�H�W"] = Convert.ToDecimal(row["360�H�W"]) * StockCost;

                row["�p�p"] = Convert.ToDecimal(row["0-30"]) +
                                Convert.ToDecimal(row["31-60"]) +
                                Convert.ToDecimal(row["61-90"]) +
                                Convert.ToDecimal(row["91-120"]) +
                                Convert.ToDecimal(row["121-180"]) +
                                Convert.ToDecimal(row["181-360"]) +
                                Convert.ToDecimal(row["360�H�W"]);

                StockCostS += Convert.ToDecimal(row["�p�p"]);
                row.EndEdit();


            }

            return dtStock;

        }
        private decimal GetItemCost(string ItemCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT cast((cast(((select sum(b.TransValue)  from oinm b where b.itemcode = t0.[ItemCode]  and  Convert(varchar(8),B.docdate,112) <=@DOCDATE and InvntAct is not null and InvntAct <>'')/case (SUM(T0.[InQty])-SUM(T0.[OutQty])) when 0 then 1 else (SUM(T0.[InQty])-SUM(T0.[OutQty])) end) as decimal(23,10))) as varchar) StockCost");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  (T0.[docdate] >='2007.12.31' AND  Convert(varchar(8),T0.docdate,112) <=@DOCDATE) ");
            sb.Append("  and  ISNULL(T1.U_GROUP,'') <> 'Z&R-�O�����s��'   AND T0.ITEMCODE=@ItemCode ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode", ItemCode));
            command.Parameters.Add(new SqlParameter("@DOCDATE", textBoxDocDate.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            if (!Convert.IsDBNull(dt.Rows[0][0]))
            {
                return Convert.ToDecimal(dt.Rows[0][0]);
            }
            else
            {
                return 0;
            }


        }
        private void GoEdit(DataRow row, Int32 StockDays, decimal Qty, string Mode)
        {

            if (Mode == "A")
            {
                row["0-30"] = 0;
                row["31-60"] = 0;
                row["61-90"] = 0;
                row["91-120"] = 0;
                row["121-180"] = 0;
                row["181-360"] = 0;
                row["360�H�W"] = 0;

            }

            if (StockDays >= 0 && StockDays <= 30)
            {
                if (Mode == "A")
                {
                    row["0-30"] = Qty;


                }
                else
                {
                    row["0-30"] = Qty + Convert.ToDecimal(row["0-30"]);
                }
            }
            else if (StockDays >= 31 && StockDays <= 60)
            {
                if (Mode == "A")
                {
                    row["31-60"] = Qty;


                }
                else
                {
                    row["31-60"] = Qty + Convert.ToDecimal(row["31-60"]);
                }
            }
            else if (StockDays >= 61 && StockDays <= 90)
            {
                if (Mode == "A")
                {
                    row["61-90"] = Qty;


                }
                else
                {
                    row["61-90"] = Qty + Convert.ToDecimal(row["61-90"]);
                }
            }
            else if (StockDays >= 91 && StockDays <= 120)
            {
                if (Mode == "A")
                {
                    row["91-120"] = Qty;


                }
                else
                {
                    row["91-120"] = Qty + Convert.ToDecimal(row["91-120"]);
                }
            }
            else if (StockDays >= 121 && StockDays <= 180)
            {
                if (Mode == "A")
                {
                    row["121-180"] = Qty;


                }
                else
                {
                    row["121-180"] = Qty + Convert.ToDecimal(row["121-180"]);
                }
            }
            else if (StockDays >= 181 && StockDays <= 360)
            {
                if (Mode == "A")
                {
                    row["181-360"] = Qty;


                }
                else
                {
                    row["181-360"] = Qty + Convert.ToDecimal(row["181-360"]);
                }
            }
            else
            {
                if (Mode == "A")
                {
                    row["360�H�W"] = Qty;


                }
                else
                {
                    row["360�H�W"] = Qty + Convert.ToDecimal(row["360�H�W"]);
                }

            }








            if (Mode == "A")
            {
                row["�p�p"] = Qty;
            }
            else
            {

                row["�p�p"] = Qty + Convert.ToDecimal(row["�p�p"]);
            }

        }
        private System.Data.DataTable MakeTable_Stock()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("�s��", typeof(string));
            dt.Columns.Add("�s��2", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("�ܮw", typeof(string));
            dt.Columns.Add("�ܮw�W��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(decimal));
  
            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360�H�W", typeof(decimal));

            dt.Columns.Add("�p�p", typeof(decimal));

            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["���~�s��"];
            colPk[1] = dt.Columns["�ܮw"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable MakeTable_Stock2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("�s��", typeof(string));
            dt.Columns.Add("�s��2", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("�ܮw", typeof(string));
            dt.Columns.Add("�ܮw�W��", typeof(string));
            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360�H�W", typeof(decimal));

            dt.Columns.Add("�p�p", typeof(decimal));

            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["���~�s��"];
            colPk[1] = dt.Columns["�ܮw"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable GetStockList( string DocDate)
        {

            System.Data.DataTable dt = GetItemHisList(DocDate);


            System.Data.DataTable dtStock = MakeTable();


            System.Data.DataTable dtTmp;

            DataRow row;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                
                string ItemCode = Convert.ToString(dt.Rows[i]["ItemCode"]);
         
                string WH = Convert.ToString(dt.Rows[i]["�ܮw"]);
                dtTmp = GetFIFO_Stock(ItemCode, DocDate, WH);

                for (int j = 0; j <= dtTmp.Rows.Count - 1; j++)
                {
                    row = dtStock.NewRow();
                    row["���~�s��"] = Convert.ToString(dtTmp.Rows[j]["ItemCode"]);
                    row["���"] = Convert.ToString(dtTmp.Rows[j]["DocDate"]);
                    row["�s��"] = Convert.ToString(dt.Rows[i]["�s��"]);
                    row["�s��2"] = Convert.ToString(dt.Rows[i]["�s��2"]);
                    row["�ܮw"] = Convert.ToString(dt.Rows[i]["�ܮw"]);
                    row["�ܮw�W��"] = Convert.ToString(dt.Rows[i]["�ܮw�W��"]);
                    row["�w�s�q"] = Convert.ToDecimal(dtTmp.Rows[j]["�w�s�q"]);
    

                    row["�w�s�Ѽ�"] = CountDays(StrToDate(DocDate), StrToDate(Convert.ToString(dtTmp.Rows[j]["DocDate"])), false);
                    dtStock.Rows.Add(row);

                }
            }

            return dtStock;

        }

        int CountDays(DateTime dateFrom, DateTime dateTo, bool including)
        {
            return ((System.TimeSpan)(dateTo - dateFrom)).Days * (-1) + (including ? 1 : 0);
        }
        private System.Data.DataTable GetItemHisList( string DocDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT MAX(substring(T3.itmsgrpNAM,4,15)) �s��,max(substring(t2.u_group,5,20)) �s��2,T0.warehouse as �ܮw,W.WhsName �ܮw�W��, ");
            sb.Append("               T0.[ItemCode] ItemCode, SUM(T0.[InQty])-SUM(T0.[OutQty]) Qty  ");
            sb.Append("               FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T2   ");
            sb.Append("               ON  T2.[ItemCode] = T0.ItemCode    ");
            sb.Append("               LEFT JOIN OWHS W on (T0.warehouse=W.whscode)  ");
            sb.Append("               INNER  JOIN [dbo].[OITB] T3  ON  T2.itmsgrpcod = T3.itmsgrpcod    ");
            sb.Append("               WHERE  T2.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE )  ");
            sb.Append("                and  ISNULL(T2.U_GROUP,'') <> 'Z&R-�O�����s��'   AND  T0.[ItemCode] NOT LIKE '%-C%'    ");
//            sb.Append("                and  ISNULL(T2.U_GROUP,'') <> 'Z&R-�O�����s��'   AND  T0.[ItemCode] NOT LIKE '%-C%'   AND T0.ITEMCODE='4EPMO.LINX.0003'  ");
            sb.Append(" AND T0.ITEMCODE not in (SELECT T0.[ItemCode] ");
            sb.Append("  FROM  [dbo].[OINM] T0  INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append(" WHERE  T1.[InvntItem] = 'Y' and  ( Convert(varchar(8),t0.docdate,112)  between '20071231' AND  @DOCDATE ) ");
         //   sb.Append("  and  ISNULL(T1.U_GROUP,'') <> 'Z&R-�O�����s��'   AND  T0.[ItemCode] NOT LIKE '%-C%' AND T0.ITEMCODE='4EPMO.LINX.0003' ");
            sb.Append("  and  ISNULL(T1.U_GROUP,'') <> 'Z&R-�O�����s��'   AND  T0.[ItemCode] NOT LIKE '%-C%' ");
            sb.Append(" GROUP BY T0.[ItemCode]  ");
            sb.Append(" Having SUM(T0.[InQty])-SUM(T0.[OutQty]) = 0)");
            sb.Append(" GROUP BY T0.warehouse,W.WhsName,T0.[ItemCode]");
            sb.Append(" Having (SUM(T0.[InQty])-SUM(T0.[OutQty]) <> 0) order by T0.[ItemCode]");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
 
            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));
            command.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }

       
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }
        private System.Data.DataTable GetFIFO_Stock(string ItemCode1, string DocDate, string WH)
        {
          
            System.Data.DataTable dt = GetFIFO(ItemCode1, DocDate,WH);


            DataRow row = null;

            decimal CalQty = 0;
            Int32 CalValue = 0;
   

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                row = dt.Rows[i];
                row.BeginEdit();

                CalQty = CalQty + Convert.ToDecimal(row["Qty"]);
                row["�֭p�ƶq"] = CalQty;


                CalValue = CalValue + Convert.ToInt32(row["TransValue"]);
                row["�֭p��"] = CalValue;

                row.EndEdit();

            }

            //�ϱ��^�h,�ѤU���w�s�q,���ӬO���X�Ӥ��

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {


                row = dt.Rows[i];

                if (Convert.ToDecimal(row["InQty"]) == 0)
                {
                    continue;
                }

                row.BeginEdit();

                row["�֭p�ƶq"] = 0;
                row["�֭p��"] = 0;


                if (CalQty - Convert.ToDecimal(row["InQty"]) <= 0)
                {
                    row["�w�s�q"] = CalQty;
                    row.EndEdit();



                    break;
                }
                else
                {
                    row["�w�s�q"] = Convert.ToDecimal(row["InQty"]);
                }

                CalQty = CalQty - Convert.ToDecimal(row["InQty"]);

                row.EndEdit();
            }



            DataView dv = dt.DefaultView;
            string fd = row["�w�s�q"].ToString();
            //dv.RowFilter = "�w�s�q > 0";

            return dv.ToTable();

        }

        private System.Data.DataTable GetFIFO_StockM(string ItemCode1, string DocDate)
        {

            System.Data.DataTable dt = GetFIFOM(ItemCode1, DocDate);


            DataRow row = null;

            decimal CalQty = 0;
            Int32 CalValue = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                row = dt.Rows[i];
                row.BeginEdit();

                CalQty = CalQty + Convert.ToDecimal(row["Qty"]);
                row["�֭p�ƶq"] = CalQty;


                CalValue = CalValue + Convert.ToInt32(row["TransValue"]);
                row["�֭p��"] = CalValue;

                row.EndEdit();

            }

            //�ϱ��^�h,�ѤU���w�s�q,���ӬO���X�Ӥ��

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {

                string InQty = row["InQty"].ToString();
                row = dt.Rows[i];

                if (Convert.ToDecimal(row["InQty"]) == 0)
                {
                    continue;
                }

                row.BeginEdit();

                row["�֭p�ƶq"] = 0;
                row["�֭p��"] = 0;


                if (CalQty - Convert.ToDecimal(row["InQty"]) <= 0)
                {
                    row["�w�s�q"] = CalQty;
                    row.EndEdit();



                    break;
                }
                else
                {
                    row["�w�s�q"] = Convert.ToDecimal(row["InQty"]);
                }

                CalQty = CalQty - Convert.ToDecimal(row["InQty"]);

                row.EndEdit();
            }


            // return dt;

            DataView dv = dt.DefaultView;
            string fd = row["Qty"].ToString();
            // dv.RowFilter = "�w�s�q > 0";

            return dv.ToTable();

        }
        private System.Data.DataTable GetFIFO(string ItemCode1, string DocDate, string WAREHOUSE)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Convert(Varchar(8),T0.[DocDate],112) DocDate,T0.[ItemCode], T1.[ItemName],SUM(T0.[InQty]) InQty, SUM(T0.[OutQty]) OutQty,SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue, ");

            //�p�����
            sb.Append(" 0.0 as �֭p�ƶq, 0 as �֭p��, 0.0 �w�s�q ");

            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("WHERE  T0.[DocDate] <= @DocDate AND T0.WAREHOUSE=@WAREHOUSE ");
            sb.Append("AND  T0.[ItemCode] = @ItemCode1  ");
            sb.Append("AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
          //  sb.Append("     and  T0.TransType <>67");
            sb.Append("GROUP BY T0.[DocDate],T0.[ItemCode], T1.[ItemName] ");
            sb.Append("ORDER BY  T0.[DocDate]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode1", ItemCode1));
            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));
            command.Parameters.Add(new SqlParameter("@WAREHOUSE", WAREHOUSE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;


        }
        private System.Data.DataTable GetFIFOM(string ItemCode1, string DocDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT Convert(Varchar(8),T0.[DocDate],112) DocDate,T0.[ItemCode], T1.[ItemName],SUM(T0.[InQty]) InQty, SUM(T0.[OutQty]) OutQty,SUM(T0.[InQty] - T0.[OutQty]) Qty,SUM(T0.[TransValue]) TransValue, ");
            sb.Append(" 0.0 as �֭p�ƶq, 0 as �֭p��, 0.0 �w�s�q ");
            sb.Append("FROM  [dbo].[OINM] T0  ");
            sb.Append("INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode  ");
            sb.Append("WHERE  T0.[DocDate] <= @DocDate ");
            sb.Append("AND  T0.[ItemCode] = @ItemCode1  ");
            sb.Append("AND  (T0.[TransValue] <> 0  OR  T0.[InQty] <> 0  OR  T0.[OutQty] <> 0  OR  T0.[TransType] = 162 )  ");
            sb.Append("     and  T0.TransType <>67");
            sb.Append("GROUP BY T0.[DocDate],T0.[ItemCode], T1.[ItemName] ");
            sb.Append("ORDER BY  T0.[DocDate]");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ItemCode1", ItemCode1));
            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;


        }

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //�Ĥ@�өT�w���
            //
            dt.Columns.Add("�s��", typeof(string));
            dt.Columns.Add("�s��2", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("�ܮw", typeof(string));
            dt.Columns.Add("�ܮw�W��", typeof(string));
            dt.Columns.Add("�w�s�q", typeof(decimal));
            dt.Columns.Add("�w�s�Ѽ�", typeof(Int32));

            DataColumn[] colPk = new DataColumn[3];
            colPk[0] = dt.Columns["���~�s��"];
            colPk[1] = dt.Columns["���"];
            colPk[2] = dt.Columns["�ܮw"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable MakeTable_StockF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("�s��", typeof(string));
            dt.Columns.Add("�s��2", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360�H�W", typeof(decimal));

            dt.Columns.Add("�p�p", typeof(decimal));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["���~�s��"];
            dt.PrimaryKey = colPk;


            return dt;
        }

        private System.Data.DataTable MakeTable_Stock2F()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("�s��", typeof(string));
            dt.Columns.Add("�s��2", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(decimal));
            dt.Columns.Add("0-30", typeof(decimal));
            dt.Columns.Add("31-60", typeof(decimal));
            dt.Columns.Add("61-90", typeof(decimal));
            dt.Columns.Add("91-120", typeof(decimal));
            dt.Columns.Add("121-180", typeof(decimal));
            dt.Columns.Add("181-360", typeof(decimal));
            dt.Columns.Add("360�H�W", typeof(decimal));

            dt.Columns.Add("�p�p", typeof(decimal));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["���~�s��"];
            dt.PrimaryKey = colPk;


            return dt;
        }
        private System.Data.DataTable MakeTableM()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

   
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("�w�s�q", typeof(int));
            dt.Columns.Add("�w�s�Ѽ�", typeof(Int32));



            return dt;
        }
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView1.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void fmAcmeFIFO_Load(object sender, EventArgs e)
        {
            if (globals.GroupID.ToString().Trim() == "ACCS")
            {
                button8.Visible = false;
                button9.Visible = false;
                return;
            }


            textBoxDocDate.Text = GetMenu.Day();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView4);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView2.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView2.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView3.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView3.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

        private void dataGridView4_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = dataGridView4.Rows[e.RowIndex];

            if (e.RowIndex == dataGridView4.Rows.Count - 1)
            {
                dgr.DefaultCellStyle.BackColor = Color.Pink;
            }
        }

    

    }
}
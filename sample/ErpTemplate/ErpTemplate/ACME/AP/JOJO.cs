using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Linq;
namespace ACME
{
    public partial class JOJO : Form
    {
      
        public JOJO()
        {
            InitializeComponent();
        }
        public string cs;
        private void button1_Click_1(object sender, EventArgs e)
        {
            

            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
              
                cs = frm1.q;
          
                string ed = "";

                if (!String.IsNullOrEmpty(cs))
                {
                   System.Data.DataTable dt2 = Getbb(cs);
                    dataGridView1.DataSource = dt2;



                    System.Data.DataTable dt3 = Getcc(cs, ed);
                    dataGridView2.DataSource = dt3;

                    decimal[] Total = new decimal[dt3.Columns.Count - 1];

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {
                        string S = dt3.Rows[i][3].ToString();

                        Total[2] += Convert.ToDecimal(dt3.Rows[i][4]);

                    }


                    DataRow row;

                    row = dt3.NewRow();

                    row[3] = "合計";

                    row[4] = Total[2];


                    dt3.Rows.Add(row);

                    System.Data.DataTable dt4 = Getdd(cs);
                    dataGridView3.DataSource = dt4;


                    decimal[] Total2 = new decimal[dt4.Columns.Count - 1];

                    for (int i = 0; i <= dt4.Rows.Count - 1; i++)
                    {


                        Total2[2] += Convert.ToDecimal(dt4.Rows[i][7]);

                    }


                    DataRow row2;

                    row2 = dt4.NewRow();

                    row2[6] = "合計";

                    row2[7] = Total2[2];


                    dt4.Rows.Add(row2);

                    System.Data.DataTable dt5 = Getee(cs);
                    dataGridView4.DataSource = dt5;
                }
            }
        }

 
        public System.Data.DataTable Getbb(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT itemcode 項目料號, CAST(T0.[OnHand] AS INT) 庫存,T0.[AvgPrice] 平均成本 FROM OITM T0 ");
            sb.Append(" where  T0.[ItemCode] in ( " + cs + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
         

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
        public System.Data.DataTable Getcc(string cs,string es)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '1' SEQ,t1.itemcode 項目料號, Convert(varchar(8),T0.docdate,112) 過帳日期,T0.cardname 客戶,cast(T1.quantity as int) 數量,T5.currency 幣別,cast(t5.price as varchar) 單價,T0.U_beneficiary 最終客戶 FROM oinv T0 ");
            sb.Append(" INNER JOIN inv1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
            sb.Append(" left join odln t9 on (t4.docentry=T9.docentry )");
            sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
            sb.Append("  WHERE ISNULL(cast(t5.price as varchar),'') <> ''");
            if (cs != "")
            {
                sb.Append(" AND  T1.[ItemCode] in ( " + cs + ")  ");
            }
            if (es != "")
            {

                sb.Append(" AND  T0.[cardcode] in ( " + es + ") and CAST(t1.quantity AS INT) <> 0  ");
            }
            sb.Append(" union all");
            sb.Append(" SELECT '1' SEQ,t1.itemcode 項目料號, Convert(varchar(8),T0.docdate,112) 過帳日期,T0.cardname 客戶,cast(T1.quantity as int) 數量,T5.currency 幣別,cast(t5.price as varchar) 單價,T0.U_beneficiary 最終客戶 FROM oinv T0 ");
            sb.Append(" INNER JOIN inv1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" left join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
            sb.Append("  WHERE ISNULL(cast(t5.price as varchar),'') <> ''  and t0.updinvnt='c'   ");
            if (cs != "")
            {
                sb.Append(" AND  T1.[ItemCode] in ( " + cs + ")  ");
            }
            if (es != "")
            {

                sb.Append(" AND  T0.[cardcode] in ( " + es + ") and CAST(t1.quantity AS INT) <> 0  ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" union all");
                sb.Append(" SELECT '2' SEQ,t1.itemcode 項目料號, Convert(varchar(8),T0.docdate,112) 過帳日期,T0.cardname 客戶,cast(T1.quantity as int) 數量,T5.currency 幣別,cast(t5.price as varchar) 單價,T0.U_beneficiary 最終客戶 FROM oinv T0 ");
                sb.Append(" INNER JOIN inv1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
                sb.Append(" left join odln t9 on (t4.docentry=T9.docentry )");
                sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
                sb.Append("  WHERE ISNULL(cast(t5.price as varchar),'') <> ''");
                if (cs != "")
                {
                    sb.Append(" AND  (T1.[ItemCode] in(SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in ( " + cs + ") ))");
                }
                sb.Append(" union all");
                sb.Append(" SELECT '2' SEQ,t1.itemcode 項目料號, Convert(varchar(8),T0.docdate,112) 過帳日期,T0.cardname 客戶,cast(T1.quantity as int) 數量,T5.currency 幣別,cast(t5.price as varchar) 單價,T0.U_beneficiary 最終客戶 FROM oinv T0 ");
                sb.Append(" INNER JOIN inv1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" left join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )");
                sb.Append("  WHERE ISNULL(cast(t5.price as varchar),'') <> ''  and t0.updinvnt='c'   ");
                if (cs != "")
                {
                    sb.Append(" AND  (T1.[ItemCode] in(SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in ( " + cs + ") ))");
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;

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
        public System.Data.DataTable Getdd(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select '1' SEQ, t1.itemcode 項目料號, Convert(varchar(8),t0.docdate,112) 採購單日期,Convert(varchar(8),t6.docdate,112) 收貨採購單日期,CAST(t5.DOCENTRY AS VARCHAR) AP單號,t0.cardname 客戶,cast(t1.price as varchar) 單價,cast(t1.quantity as int) 數量 ,T1.U_MEMO 備註 from opor t0 ");
            sb.Append(" inner join por1 t1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" inner join pdn1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
            sb.Append(" inner join opdn t6 on (t4.docentry=t6.docentry)  ");
            sb.Append(" inner join pch1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='20'  ) ");
            sb.Append(" where  T1.[ItemCode]  in ( " + cs + ")    ");
            if (checkBox1.Checked)
            {
                sb.Append(" UNION ALL");
                sb.Append(" select '2',t1.itemcode 項目料號, Convert(varchar(8),t0.docdate,112) 採購單日期,Convert(varchar(8),t6.docdate,112) 收貨採購單日期,CAST(t5.DOCENTRY AS VARCHAR) AP單號,t0.cardname 客戶,cast(t1.price as varchar) 單價,cast(t1.quantity as int) 數量 ,T1.U_MEMO 備註 from opor t0 ");
                sb.Append(" inner join por1 t1 on (t0.docentry=t1.docentry)  ");
                sb.Append(" inner join pdn1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
                sb.Append(" inner join opdn t6 on (t4.docentry=t6.docentry)  ");
                sb.Append(" inner join pch1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='20'  ) ");
                sb.Append(" where  (T1.[ItemCode] in(SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in ( " + cs + ") ))");
            }
                sb.Append(" ORDER BY SEQ,t1.itemcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable Getee(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct Convert(varchar(10),t0.docdate,112) 日期,t1.itemcode 原始料號,t2.itemcode 調整料號,cast(t1.quantity as int) 數量 from oige t0");
            sb.Append(" inner join ige1 t1 on (t0.docentry=t1.docentry)");
            sb.Append(" inner join (select t0.docdate,t1.itemcode,t1.quantity,t1.u_base_doc from oign t0 left join ign1 t1 on (t0.docentry=t1.docentry) where t0.u_acme_kind1 like '%料號調整%' ) t2 on (t0.docdate=t2.docdate and t1.quantity=t2.quantity and t1.docentry=t2.u_base_doc)");
            sb.Append(" where t0.u_acme_kind1 like '%料號調整%'  and   T1.[ItemCode] in ( " + cs + ") and t1.itemcode <> t2.itemcode ");
            sb.Append(" union all");
            sb.Append(" select '','數量加總','',sum(cast(t1.quantity as int)) 數量 from oige t0");
            sb.Append(" inner join ige1 t1 on (t0.docentry=t1.docentry)");
            sb.Append(" inner join (select t0.docdate,t1.itemcode,t1.quantity,t1.u_base_doc from oign t0 left join ign1 t1 on (t0.docentry=t1.docentry) where t0.u_acme_kind1 like '%料號調整%' ) t2 on (t0.docdate=t2.docdate and t1.quantity=t2.quantity and t1.docentry=t2.u_base_doc )");
            sb.Append(" where t0.u_acme_kind1 like '%料號調整%'  and   T1.[ItemCode] in ( " + cs + ") and t1.itemcode <> t2.itemcode ");
            sb.Append("    order by t1.itemcode,Convert(varchar(10),t0.docdate,112)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        public System.Data.DataTable Getee1()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                           SELECT Convert(varchar(8),T4.[actdelDate],112) 過帳日期,T0.[docentry] 採購單號 ,t4.docentry 交貨單號, ");
            sb.Append("                            CASE SUBSTRING(ITMSGRPNAM,4,5) WHEN 'Pleas' THEN '' ELSE SUBSTRING(ITMSGRPNAM,4,5) END 部門,(T0.[CardName]) 客戶名稱, ");
            sb.Append("           T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本, ");
            sb.Append("                            (Substring(T1.[ItemCode],15,1)) 地區,cast(t4.quantity as int) 數量,cast(T1.[Price] as numeric(16,2)) 單價,T1.CURRENCY 幣別,T9.U_ACME_INV AUOINVOICE,T1.U_MEMO 備註 ");
            sb.Append("                            FROM OPOR T0  ");
            sb.Append("                            INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("                            left JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("                            left JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("                            left join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
            sb.Append("                            left join OPDN T9 on (t4.DOCENTRY=T9.DOCENTRY ) ");
            sb.Append("                            left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  ) ");
            sb.Append("                            left JOIN Owhs T7 ON T7.whsCode = T1.whscode  ");
            sb.Append("               left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               left JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod ");
            sb.Append("                            WHERE   T1.[LINESTATUS] ='C' and T1.trgetentry <>'' and isnull(T8.basetype,'') ='' AND ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T4.[actdelDate],112) between @DocDate1 and @DocDate2 ");
            }
            if (textBox4.Text != "")
            {
                sb.Append(" and  T0.[CardCode] ='" + textBox4.Text.ToString() + "' ");
            }
            sb.Append(" order by (t0.docentry) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        public System.Data.DataTable GetALL()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_BU 項目群組,CAST(sum(T0.[StockValue]) AS INT) 存貨金額   FROM OITM T0 ");
            sb.Append(" INNER  JOIN [dbo].[OITB] T2  ON  T0.itmsgrpcod = T2.itmsgrpcod ");
            sb.Append(" where T0.[OnHand]>0 and t0.itemcode not in (select itemcode from oitm where invntitem='N' AND substring(itemcode,1,1) IN ('R','Z'))");
            sb.Append(" And substring(t0.itemcode,1,2) <> 'ZR'");
            sb.Append(" And substring(t0.itemcode,1,2) <> 'ZA'");
            sb.Append(" And substring(t0.itemcode,1,2) <> 'ZB'");
            sb.Append(" group by U_BU ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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

        public System.Data.DataTable Ged1()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT Convert(varchar(8),t0.docdate,112) 製單日期,(T0.[CardName]) 客戶名稱,   ");
            sb.Append("                   T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本,(Substring(T1.[ItemCode],15,1)) 地區,(t1.opencreqty) 未結數量,cast(cast(T1.[Price] as numeric(16,2)) as varchar) 單價,T1.CURRENCY 幣別 ");
            sb.Append("               ,Convert(varchar(8),t1.ShipDate,112) 訂單交期,Convert(varchar(8),t1.u_acme_shipday,112)  離倉日期,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,t0.u_beneficiary 最終客戶,T0.DOCENTRY 單號,t1.u_acme_dscription SA備註,CASE SUBSTRING(ITMSGRPNAM,4,5) WHEN 'Pleas' THEN '' ELSE SUBSTRING(ITMSGRPNAM,4,5) END  部門  ");
            sb.Append("               FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("               iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("               iNNER JOIN OWHS T4 ON T4.whsCode = T1.whscode ");
            sb.Append("               left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               left JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod ");
            sb.Append("               WHERE    T1.[LINESTATUS] ='O' AND  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            if (textBox10.Text != "")
            {
                sb.Append(" and  T0.[cardcode] ='" + textBox10.Text.ToString() + "' ");
            }
            if (textBox12.Text != "" && textBox11.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[docdate],112) between @DocDate1 and @DocDate2 ");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and Substring (T1.[ItemCode],2,8) ='" + comboBox1.SelectedValue.ToString() + "'  ");
            }
            sb.Append(" order by (t0.docdate) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox12.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox11.Text));
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
        public System.Data.DataTable Ged2()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),t0.docdate,112) 製單日期,(T0.[CardName]) 客戶名稱,   ");
            sb.Append(" T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本,(Substring(T1.[ItemCode],15,1)) 地區,(t1.opencreqty) 未結數量,cast(cast(T1.[Price] as numeric(16,2)) as varchar) 單價,T1.CURRENCY 幣別  ");
            sb.Append(" ,Convert(varchar(8),t1.ShipDate,112) 訂單交期,(T2.[SlpName]) 業務,(T3.[lastName]+T3.[firstName]) 業管,T0.DOCENTRY 單號,CASE SUBSTRING(ITMSGRPNAM,4,5) WHEN 'Pleas' THEN '' ELSE SUBSTRING(ITMSGRPNAM,4,5) END  部門 ,T1.U_MEMO 備註");
            sb.Append(" FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode  ");
            sb.Append(" iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID   ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
            sb.Append(" left JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod  ");
            sb.Append(" WHERE    T1.[LINESTATUS] ='O' AND ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' ");
            if (textBox14.Text != "")
            {
                sb.Append(" and  T0.[cardcode] = '" + textBox14.Text.ToString() + "' ");
            }
            if (textBox16.Text != "" && textBox15.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[docdate],112) between @DocDate1 and @DocDate2 ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and Substring (T1.[ItemCode],2,8) ='" + comboBox2.SelectedValue.ToString() + "'  ");
            }
            sb.Append(" order by (t0.docdate) desc");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox16.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox15.Text));
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

        private void button2_Click_1(object sender, EventArgs e)
        {
            APS1 frm1 = new APS1();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                cs = frm1.q;
                string ed = "";
                if (!String.IsNullOrEmpty(cs))
                {
                    

                    System.Data.DataTable dt3 = Getcc(ed, cs);
                    dataGridView2.DataSource = dt3;

                    decimal[] Total = new decimal[dt3.Columns.Count - 1];

                    for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                    {


                        Total[2] += Convert.ToDecimal(dt3.Rows[i][4]);

                    }
                    

                    DataRow row;

                    row = dt3.NewRow();

                    row[3] = "合計";
           
                        row[4] = Total[2];

                    
                    dt3.Rows.Add(row);
                 
                }
            }
            tabControl1.SelectedIndex = 3;
        }

      

       

        private void button5_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToCSV2(dataGridView1, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToCSV2(dataGridView3, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToCSV2(dataGridView4, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                ExcelReport.GridViewToCSV2(dataGridView2, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                ExcelReport.GridViewToCSV2(dataGridView5, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                ExcelReport.GridViewToCSV2(dataGridView6, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                ExcelReport.GridViewToCSV2(dataGridView7, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 7)
            {
                ExcelReport.GridViewToCSV(dataGridView8, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
            else if (tabControl1.SelectedIndex == 8)
            {
                ExcelReport.GridViewToCSV2(dataGridView9, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            }
            else if (tabControl1.SelectedIndex == 9)
            {
                ExcelReport.GridViewToCSV2(dataGridView10, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
           
            }
            else if (tabControl1.SelectedIndex == 10)
            {
                ExcelReport.GridViewToCSV2(dataGridView11, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            }
            else if (tabControl1.SelectedIndex == 11)
            {
                ExcelReport.GridViewToCSV2(dataGridView12, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            }
        }

      
        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();

            if (LookupValues != null)
            {
              textBox4.Text = Convert.ToString(LookupValues[0]);
              textBox3.Text = Convert.ToString(LookupValues[1]);

            }
        
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            dataGridView5.DataSource = Getee1();
        }

        private void JOJO_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();
            textBox12.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox11.Text = GetMenu.Day();
            textBox16.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox15.Text = GetMenu.Day();
            textBox17.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox18.Text = GetMenu.Day();
            textBox20.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox21.Text = GetMenu.Day();
            UtilSimple.SetLookupBinding(comboBox1, GetOslp(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetOslp2(), "DataValue", "DataValue");
            label20.Text = "";
        }
        System.Data.DataTable GetOslp()
        {

            SqlConnection con = globals.shipConnection;
            string sql = " SELECT DISTINCT Substring (T1.[ItemCode],2,8) DataText,Substring (T1.[ItemCode],2,8) DataValue FROM RDR1 T1 iNNER JOIN OWHS T4 ON T4.whsCode = T1.whscode WHERE T1.[LINESTATUS] ='O' UNION ALL SELECT '0', 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM'  order by DataText ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ousr");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ousr"];
        }
        System.Data.DataTable GetOslp2()
        {

            SqlConnection con = globals.shipConnection;
            string sql = " SELECT DISTINCT Substring (T1.[ItemCode],2,8) DataText,Substring (T1.[ItemCode],2,8) DataValue FROM por1 T1  INNER  JOIN [dbo].[OITM] T11  ON  T1.[ItemCode] = T11.ItemCode WHERE T1.[LINESTATUS] ='O'  AND ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  UNION ALL SELECT '0', 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM'  order by DataText ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ousr");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ousr"];
        }
        private void button6_Click(object sender, EventArgs e)
        {

            System.Data.DataTable t1 = Gen_201004();
            dataGridView6.DataSource = t1;
        }



        private void button8_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListC();

            if (LookupValues != null)
            {
                textBox7.Text = Convert.ToString(LookupValues[0]);
                textBox8.Text = Convert.ToString(LookupValues[1]);

            }
        }

  

   
        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView7.DataSource = Ged1();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = Ged2();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListC();

            if (LookupValues != null)
            {
                textBox10.Text = Convert.ToString(LookupValues[0]);
                textBox9.Text = Convert.ToString(LookupValues[1]);

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuListS();

            if (LookupValues != null)
            {
                textBox14.Text = Convert.ToString(LookupValues[0]);
                textBox13.Text = Convert.ToString(LookupValues[1]);

            }
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


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
        private void button12_Click(object sender, EventArgs e)
        {

            DELPENCELL2();
          //  System.Data.DataTable dtData = GetDataSort(table, "notZero");//把一樣的相加 小計不為零
            dataGridView9.DataSource = GetALL();

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
                        System.Data.DataTable K1 = GETS1(ITEMCODE);
                        if (K1.Rows.Count > 0)
                        {
                            decimal k2 = 0;
                            for (int i = 0; i <= K1.Rows.Count - 1; i++)
                            {
                                string KITEM = K1.Rows[i][0].ToString();

                                decimal PRICE = Convert.ToDecimal(GETS2(KITEM).Rows[0][0]);

                                k2 += PRICE;
                            }
                            k2 = k2 / (K1.Rows.Count);

                            decimal k1 = Convert.ToDecimal(onhand - Tonhand);
                            ADDOPENCELL2(ITEMCODE, onhand - Tonhand, k2 * k1);
                        }


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

        private System.Data.DataTable Gen_201004()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;



            StringBuilder sb = new StringBuilder();

            //彙總
            sb.Append("               SELECT Convert(varchar(8),T0.[DocDate],112)  過帳日期, T0.CardName 客戶名稱,SUBSTRING(ITMSGRPNAM,4,5) 部門, ");
            sb.Append("                     T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本,");
            sb.Append("               (Substring(T2.[ItemCode],15,1)) 地區,(T2.Quantity) 數量, ");
            sb.Append("               cast(cast(T5.[Price] as numeric(16,2)) as varchar) 訂單單價,T2.[Price] 台幣單價, ");
            sb.Append("               CAST(T2.LineTotal AS INT) 台幣金額, ");
            sb.Append("               CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT) 台幣毛利, ");
            sb.Append("               T3.SLPNAME 業務,'AR'+CAST(T0.DOCENTRY AS VARCHAR) 單據 FROM OINV T0  ");
            sb.Append("               INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode  ");
            sb.Append("               left join dln1 t4 on (t2.baseentry=T4.docentry and  t2.baseline=t4.linenum  and t2.basetype='15') ");
            sb.Append("               left join odln t9 on (t4.docentry=T9.docentry ) ");
            sb.Append("               left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append("               INNER JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               INNER JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod  ");
            sb.Append("               WHERE T0.[DocType] ='I'  ");
            sb.Append("               AND  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.U_IN_BSTYC <> '1'  ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112) between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T0.[CardCode] ='" + textBox7.Text.ToString() + "' ");
            }
            sb.Append(" UNION ALL");
            sb.Append("              SELECT Convert(varchar(8),T0.[DocDate],112)  過帳日期, T0.CardName 客戶名稱,SUBSTRING(ITMSGRPNAM,4,5) 部門, ");
            sb.Append("              T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本, ");
            sb.Append("               (Substring(T2.[ItemCode],15,1)) 地區,(T2.Quantity)*-1 數量, ");
            sb.Append("               T2.U_ACME_INV 訂單單價,T2.[Price] 台幣單價, ");
            sb.Append("               CAST(T2.LineTotal AS INT)*-1 台幣金額, ");
            sb.Append("               (CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT))*-1 台幣毛利, ");
            sb.Append("               T3.SLPNAME 業務,'貸項'+CAST(T0.DOCENTRY AS VARCHAR) 單據 FROM ORIN T0  ");
            sb.Append("               INNER JOIN RIN1 T2 ON T0.DocEntry = T2.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode  ");
            sb.Append("               INNER JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               INNER JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod  ");
            sb.Append("               WHERE T0.[DocType] ='I'   ");
            sb.Append("               and  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");


            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112) between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T0.[CardCode] ='" + textBox7.Text.ToString() + "' ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox5.Text));

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox6.Text));





            SqlDataAdapter da = new SqlDataAdapter(command);



            DataSet ds = new DataSet();

            try
            {

                connection.Open();

                da.Fill(ds, "OINV");

            }

            finally
            {

                connection.Close();

            }

            return ds.Tables[0];



        }

        private System.Data.DataTable Gen_201005()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            string FS = textBox19.Text;

            StringBuilder sb = new StringBuilder();

            //彙總
            sb.Append("             SELECT  T2.[ItemCode] 項目料號,T0.CardName 客戶名稱,Convert(varchar(8),T0.[DocDate],112)  過帳日期, ");
            sb.Append("        (T2.Quantity) 數量,T2.[Price] 單價,");
            sb.Append("              CAST(T2.LineTotal AS INT) 銷售金額,CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) 成本,");
            sb.Append("              CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT) 毛利,");
            sb.Append("         CAST(CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT)/T2.Quantity AS INT)    每片毛利, T3.SLPNAME 業務,");
            sb.Append("              'AR'+CAST(T0.DOCENTRY AS VARCHAR) 單據 FROM OINV T0 ");
            sb.Append("              INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry ");
            sb.Append(" INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode   INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T2.ItemCode ");
            sb.Append("              WHERE T0.[DocType] ='I' ");
            sb.Append("              and  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.U_IN_BSTYC <> '1' ");


            if (textBox17.Text != "" && textBox18.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112)  between @DocDate1 and @DocDate2 ");
            }

            sb.Append(" and  T2.[ItemCode]  in ( " + FS + ")");

            sb.Append("  ORDER BY T2.ItemCode,Convert(varchar(8),T0.[DocDate],112)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox17.Text));

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox18.Text));





            SqlDataAdapter da = new SqlDataAdapter(command);



            DataSet ds = new DataSet();

            try
            {

                connection.Open();

                da.Fill(ds, "OINV");

            }

            finally
            {

                connection.Close();

            }

            return ds.Tables[0];



        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox19.Text == "")
            {
                MessageBox.Show("請輸入項目號碼");
                return;
            }

            System.Data.DataTable t1 = Gen_201005();
            if (t1.Rows.Count > 0)
            {
                dataGridView10.DataSource = t1;

                string g = t1.Compute("AVG(每片毛利)", null).ToString();


                decimal sh = Convert.ToDecimal(g);

                label20.Text = "平均每片毛利 " + sh.ToString("#,##0");

            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
  
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {

                cs = frm1.q;

            

                if (!String.IsNullOrEmpty(cs))
                {
                    textBox19.Text = cs;
                }
            }
        }


        private void button16_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetTABLE();
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTabe();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string 銷售單號 = dt.Rows[i]["銷售單號"].ToString();
                string 客戶名稱 = dt.Rows[i]["客戶名稱"].ToString();
                string 銷售單價 =  dt.Rows[i]["銷售單價"].ToString();
                string 毛利 = dt.Rows[i]["毛利"].ToString();
                string 生產訂單 = dt.Rows[i]["生產訂單"].ToString();
                string 廠商 = dt.Rows[i]["廠商"].ToString();
                string 項目號碼 = dt.Rows[i]["項目號碼"].ToString();
                string 項目說明 = dt.Rows[i]["項目說明"].ToString();
                string 數量 = dt.Rows[i]["數量"].ToString();
                string f2 = dt.Rows[i]["單價"].ToString();
                if (f2 == "")
                {
                    f2 = "0";
                }
                int 單價 = Convert.ToInt16(f2);
                string 出貨日期 = dt.Rows[i]["出貨日期"].ToString();

                dr["銷售訂單"] = 銷售單號;
                dr["客戶名稱"] = 客戶名稱;
                dr["銷售單價"] = 銷售單價;
                dr["毛利"] = 毛利;
                dr["生產訂單"] = 生產訂單;
                dr["廠商"] = 廠商;
                dr["項目號碼"] = 項目號碼;
                dr["項目說明"] = 項目說明;
                dr["數量"] = 數量;
                dr["單價"] = 單價;
                dr["出貨日期"] = 出貨日期;
                System.Data.DataTable dt3 = GetTABLE3(生產訂單);
                if (dt3.Rows.Count > 0)
                {
                    dr["生產日期"] = dt3.Rows[0][0].ToString();
                }
                dtCost.Rows.Add(dr);

                System.Data.DataTable dt2 = GetTABLE2(生產訂單);
                if (dt2.Rows.Count > 0)
                {
                    for (int j = 0; j <= dt2.Rows.Count - 1; j++)
                    {
                        dr = dtCost.NewRow();
                        dr["生產日期"] = "";
                        dr["出貨日期"] = "";
                        dr["銷售訂單"] = "";
                        dr["客戶名稱"] = "";
                        dr["銷售單價"] = "";
                        dr["毛利"] = "";
                        dr["生產訂單"] = "";
                        dr["廠商"] = "";
                        dr["項目號碼"] = dt2.Rows[j]["子件編號"].ToString();
                        dr["項目說明"] = dt2.Rows[j]["產品名稱"].ToString();
                        dr["數量"] = "";
                        string f1 = dt2.Rows[j]["單價"].ToString();
                        if (f1 == "")
                        {
                            f1 = "0";
                        }
                        dr["單價"] = Convert.ToInt16(f1);

                        dtCost.Rows.Add(dr);
                    }
                }


            }


            dataGridView11.DataSource = dtCost;

       
        }

        private System.Data.DataTable MakeTabe()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("生產日期", typeof(string));
            dt.Columns.Add("出貨日期", typeof(string));
            dt.Columns.Add("銷售訂單", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("銷售單價", typeof(string));
            dt.Columns.Add("毛利", typeof(string));
            dt.Columns.Add("生產訂單", typeof(string));
            dt.Columns.Add("廠商", typeof(string));
            dt.Columns.Add("項目號碼", typeof(string));
            dt.Columns.Add("項目說明", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(int));
            return dt;
        }
        private System.Data.DataTable MakeTabeCHECKPAID()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("銷售單號", typeof(string));
            dt.Columns.Add("AR單號", typeof(string));
            dt.Columns.Add("型號", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("美金單價", typeof(string));
            dt.Columns.Add("美金總額", typeof(string));
            dt.Columns.Add("台幣總額", typeof(string));
            dt.Columns.Add("過帳日期", typeof(string));
            dt.Columns.Add("到帳日期", typeof(string));
            dt.Columns.Add("入帳日期", typeof(string));
            dt.Columns.Add("逾期天數", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("付款方法", typeof(string));
            return dt;
        }
        private System.Data.DataTable GetTABLE()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                         SELECT T0.[OriginNum]  銷售單號, T2.[CardName] 客戶名稱,");
            sb.Append("                                              'USD'+CAST(CAST(T5.PRICE AS DECIMAL(10,2)) AS VARCHAR) 銷售單價");
            sb.Append("                                            ,T5.DOCTOTAL 毛利, T0.DOCENTRY 生產訂單,T0.U_CARDNAME 廠商,");
            sb.Append("                                              T0.[ItemCode] 項目號碼, T1.[ItemName] 項目說明, ");
            sb.Append("                                             CAST(T0.PLANNEDQTY AS INT ) 數量,");
            sb.Append("            單價 = (SELECT abs(Convert(int,Sum(T7.[TransValue]))) FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum  AND  T7.[AppObjType] = 'C')/  CAST(T0.PLANNEDQTY AS INT ) ");
            sb.Append(" ,Convert(varchar(8),T5.SHIPDATE,112) 出貨日期");
            sb.Append("                                              FROM OWOR T0 ");
            sb.Append("                                              INNER JOIN OITM T1 ON T0.ItemCode= T1.ItemCode");
            sb.Append("                                              Left join ORDR T2 on T2.DocEntry=T0.[OriginNum]");
            sb.Append("               LEFT JOIN (SELECT CAST(SUM(T2.LineTotal) - SUM(Round(T2.StockPrice*T2.Quantity,0)) AS INT) DOCTOTAL,T1.DOCENTRY,T1.[ItemCode],AVG(T1.PRICE) PRICE,MAX(T6.DOCDATE)  SHIPDATE FROM RDR1 T1");
            sb.Append("                                 LEFT JOIN DLN1 T2 ON (T2.baseentry=T1.docentry and  T2.baseline=T1.linenum  and T1.targettype='15')");
            sb.Append("                             LEFT JOIN ODLN T6 ON (T2.DOCENTRY=T6.DOCENTRY)       ");
            sb.Append("     LEFT JOIN INV1 T3 ON (T3.baseentry=T1.docentry and  T3.baseline=T1.linenum  and T2.targettype='13')");
            sb.Append("                GROUP BY T1.DOCENTRY,T1.[ItemCode] ) T5 ON (T2.DOCENTRY=T5.DOCENTRY AND  T0.[ItemCode]=T5.[ItemCode])");
            sb.Append("                                      Where 1=1  AND T1.ITMSGRPCOD='1032' AND T0.STATUS <> 'C' ");
            sb.Append(" and  Convert(varchar(8),T0.[PostDate],112) between @DocDate1 and @DocDate2 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox20.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2 ", textBox21.Text));

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
        private System.Data.DataTable GetTABLECHECKPAID()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("       SELECT (T0.[CardName]) 客戶,T5.[docentry] 銷售單號 ,t0.docentry AR單號,T11.U_TMODEL 型號,CAST(t1.quantity AS INT) 數量,");
            sb.Append(" T5.CURRENCY+cast(cast(T5.[Price] as numeric(16,2)) as varchar) 美金單價,cast(T5.[GtotalFC] as varchar) 美金總額,");
            sb.Append("   T1.[Gtotal] 台幣總額,Convert(varchar(10),T0.[docDate],112)  過帳日期, Convert(varchar(8),dbo.fun_CreditDate(T8.u_acme_pay,T0.CardCode,T0.DocDate),112)  到帳日期             ");
            sb.Append("             ,t9.u_acme_pay 付款條件,T5.LINENUM    FROM OINV T0 ");
            sb.Append("               INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("               left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15') ");
            sb.Append("               left join odln t9 on (t4.docentry=T9.docentry ) ");
            sb.Append("               left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append("               left join ordr t8 on (t8.docentry=T5.docentry  ) ");
            sb.Append("               left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               left JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod ");
            sb.Append("               WHERE   T0.CARDNAME LIKE  '%" + textBox22.Text + "%' and Convert(varchar(10),T0.[docDate],111)  > '20121231'   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox20.Text));
        

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
        private System.Data.DataTable GetTABLE2(string DocNum)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT W1.ItemCode 子件編號, T1.[ItemName] 產品名稱,");
            sb.Append("  單價 = CAST((SELECT abs(Convert(int,Sum(T7.[TransValue])))   FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum AND  T7.[AppObjLine] = W1.LineNum AND T7.[ItemCode] = W1.ItemCode AND  T7.[AppObjType] = 'C'  )/W1.[PlannedQty]   AS INT)");
            sb.Append("  FROM OWOR T0 ");
            sb.Append("  INNER JOIN WOR1 W1 ON W1.DocEntry=T0.DocNum");
            sb.Append("  Left JOIN OITM T1 ON T1.ItemCode= W1.ItemCode");
            sb.Append("  Where T0.DocNum=@DocNum");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocNum", DocNum));

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

        private System.Data.DataTable GetTABLE3(string BASEREF)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T0.DOCDATE,112)  FROM OIGN T0 ");
            sb.Append(" LEFT JOIN IGN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T1.BASEREF=@BASEREF");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BASEREF", BASEREF));

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
        private System.Data.DataTable GetTABLE3CHECKPAID(string DOCENTRY, string LINENUM)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                select PAYCHECK 付款方法,TTDATE 入帳日期 from satt1 t0 left join satt t1 on (t0.ttcode=t1.ttcode) ");
            sb.Append("                                 left join satt2 t2 on (t2.ttcode=t0.ttcode AND T2.ID=T0.SEQNO) ");
            sb.Append("           WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        DateTime T1;
        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox22.Text=="")
            {

                MessageBox.Show("請輸入客戶簡稱");
                return;
            }

            System.Data.DataTable dt = GetTABLECHECKPAID();
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTabeCHECKPAID();
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                string DOC = dt.Rows[i]["銷售單號"].ToString();
                string LINENUM = dt.Rows[i]["LINENUM"].ToString();
                dr = dtCost.NewRow();
                dr["客戶"] = dt.Rows[i]["客戶"].ToString();
                dr["銷售單號"] = dt.Rows[i]["銷售單號"].ToString();
                dr["AR單號"] = dt.Rows[i]["AR單號"].ToString();
                dr["型號"] = dt.Rows[i]["型號"].ToString();
                dr["數量"] = dt.Rows[i]["數量"].ToString();
                string S1 = dt.Rows[i]["到帳日期"].ToString();
                string DATE = S1.Substring(0, 4) + "/" + S1.Substring(4, 2) + "/" + S1.Substring(6, 2);
                dr["美金單價"] = dt.Rows[i]["美金單價"].ToString();
                dr["美金總額"] = dt.Rows[i]["美金總額"].ToString();
                dr["台幣總額"] = dt.Rows[i]["台幣總額"].ToString();
                dr["過帳日期"] = dt.Rows[i]["過帳日期"].ToString();
                dr["到帳日期"] = S1;
                dr["付款條件"] = dt.Rows[i]["付款條件"].ToString();
               
                if (!String.IsNullOrEmpty(DATE))
                {
                     T1 = Convert.ToDateTime(DATE);
                }

                System.Data.DataTable dt3 = GetTABLE3CHECKPAID(DOC, LINENUM);
                if (dt3.Rows.Count > 0)
                {
                    string S2 = dt3.Rows[0]["入帳日期"].ToString();
                    dr["付款方法"] = dt3.Rows[0]["付款方法"].ToString();
                    dr["入帳日期"] = S2;
                    //逾期天數
                    string DATE2 = S2.Substring(0, 4) + "/" + S2.Substring(4, 2) + "/" + S2.Substring(6, 2);
                    if (!String.IsNullOrEmpty(S2))
                    {
                        DateTime T2 = Convert.ToDateTime(DATE2);

                        if (!String.IsNullOrEmpty(S1))
                        {
                            TimeSpan ts = T2.Subtract(T1);
                            dr["逾期天數"] = ts.Days.ToString();
                        }
                    }
                }
              
             
            
                //      dt.Columns.Add("逾期天數", typeof(string));
                dtCost.Rows.Add(dr);
            }
            dataGridView12.DataSource = dtCost;
        }

        private void dataGridView9_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView9.SelectedRows.Count > 0)
            {

                string da = dataGridView9.SelectedRows[0].Cells["項目群組"].Value.ToString();

                JOJO2 a = new JOJO2();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

   
    }
}
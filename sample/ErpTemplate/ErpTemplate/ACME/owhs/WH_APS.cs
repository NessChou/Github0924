using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace ACME
{

    public partial class WH_APS : Form
    {
        public string s;
        public WH_APS()
        {
            InitializeComponent();
        }
        System.Data.DataTable aa;
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = PackOP();
            dataGridView2.DataSource = PackOP2();
            dataGridView3.DataSource = PackOP3();
            dataGridView4.DataSource = PackOP4();
         
        }

        public System.Data.DataTable PackOP()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T0.[TAXDate],112)  文件日期,T0.DOCENTRY 採購單號,T0.CARDNAME 廠商名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ");
            sb.Append(" ,T1.U_MEMO  備註,'1' SEQ FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode    left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  WHERE T1.[LINESTATUS] ='O'   and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" union all");
                sb.Append(" SELECT Convert(varchar(8),T0.[TAXDate],112)  文件日期,T0.DOCENTRY 採購單號,T0.CARDNAME 廠商名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ");
                sb.Append(" ,T1.U_MEMO  備註,'2' SEQ FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode    left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  WHERE T1.[LINESTATUS] ='O'   and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");
                sb.Append("               SELECT DISTINCT T1.ITEMCODE  COLLATE Chinese_PRC_CI_AS 產品編號 FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("               INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode    left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  WHERE T1.[LINESTATUS] ='O'   and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {

                    sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                }
                if (checkBox1.Checked)
                {
                    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                }
                else
                {
                    if (textBox5.Text != "" && textBox6.Text != "")
                    {
                        sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                    }
                }
                if (checkBox2.Checked)
                {
                    sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                    }
                }
                sb.Append(" ) )) ");
            }
            sb.Append("  ORDER BY SEQ,文件日期 DESC ");
         
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }


        public System.Data.DataTable PackOP2()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T0.CARDNAME 廠商名稱, ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
            sb.Append(" ELSE   T1.ITEMCODE END 產品編號, ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm'");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm'");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm'");
            sb.Append(" ELSE  T1.[Dscription]  END 品名規格, ");
            sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 , ");
            sb.Append(" T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'1' SEQ  ");
            sb.Append(" FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)   ");
            sb.Append(" INNER join opdn t44 on (t4.docentry=t44.docentry)   ");
            sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode    ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
            sb.Append(" left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)   ");
            sb.Append(" left join OPCH t55 on (t5.docentry=t55.docentry)   ");
            sb.Append(" left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )   ");
            sb.Append(" left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )   ");
            sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND (T1.[Quantity]-ISNULL(T8.QUANTITY,0) <> 0)  AND (T1.[Quantity]-ISNULL(T9.QUANTITY,0) <> 0)     ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T44.[DOCDate],112)  between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ELSE   T1.ITEMCODE END in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ELSE   T1.ITEMCODE END between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append(" union all");
                sb.Append(" SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T0.CARDNAME 廠商名稱,");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
                sb.Append(" ELSE   T1.ITEMCODE END 產品編號, ");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm'");
                sb.Append(" ELSE  T1.[Dscription]  END 品名規格, ");
                sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,");
                sb.Append(" T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'2' SEQ ");
                sb.Append(" FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append(" INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)  ");
                sb.Append(" INNER join opdn t44 on (t4.docentry=t44.docentry)  ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode   ");
                sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
                sb.Append(" left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)  ");
                sb.Append(" left join OPCH t55 on (t5.docentry=t55.docentry)  ");
                sb.Append(" left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )  ");
                sb.Append(" left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )  ");
                sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND (T1.[Quantity]-ISNULL(T8.QUANTITY,0) <> 0)  AND (T1.[Quantity]-ISNULL(T9.QUANTITY,0) <> 0)    ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");

                sb.Append(" SELECT DISTINCT   CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6'  ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6'  ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6'  ");
                sb.Append(" ELSE   T1.ITEMCODE END COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號");
                sb.Append(" FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry    ");
                sb.Append(" INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)    ");
                sb.Append(" INNER join opdn t44 on (t4.docentry=t44.docentry)    ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode     ");
                sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
                sb.Append(" left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)    ");
                sb.Append(" left join OPCH t55 on (t5.docentry=t55.docentry)    ");
                sb.Append(" left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )    ");
                sb.Append(" left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )   ");
                sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND (T1.[Quantity]-ISNULL(T8.QUANTITY,0) <> 0)  AND (T1.[Quantity]-ISNULL(T9.QUANTITY,0) <> 0)");

                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {

                    sb.Append(" and  Convert(varchar(8),T44.[DOCDate],112)  between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                }

                if (checkBox1.Checked)
                {
                    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                }
                else
                {
                    if (textBox5.Text != "" && textBox6.Text != "")
                    {
                        sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                    }
                }
                if (checkBox2.Checked)
                {
                    sb.Append(" and CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
                    sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
                    sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
                    sb.Append(" ELSE   T1.ITEMCODE END  ");
                    sb.Append("  in ( " + d + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
                        sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
                        sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
                        sb.Append(" ELSE   T1.ITEMCODE END  ");
                        sb.Append(" between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                    }
                }

                sb.Append(" ) )) ");
            }
            sb.Append("  ORDER BY SEQ,進貨日期 DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
     
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }

        public System.Data.DataTable PackOP3()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                                   SELECT Convert(varchar(8),T0.[DOCDate],112)    出貨日期,T44.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別");
            sb.Append(" ,t8.u_beneficiary 最終客戶,T8.U_ACME_MEMO   備註,'1' SEQ ");
            sb.Append("                    FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append("               INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )             ");
            sb.Append("  left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
            sb.Append("               INNER join OINV t44 on (t4.docentry=t44.docentry)  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
            sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }

            sb.Append(" UNION ALL");
            sb.Append("                                                 SELECT Convert(varchar(8),T0.[DOCDate],112)   出貨日期,T0.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別 ");
            sb.Append("               ,t8.u_beneficiary 最終客戶,T8.U_ACME_PAYGUI   發票金額,'1' SEQ  ");
            sb.Append("                                  FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append("                             INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')  ");
            sb.Append("               left join ordr t8 on (t8.docentry=T5.docentry  )              ");
            sb.Append("  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
            sb.Append("                             WHERE              ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }
            if (checkBox3.Checked)
            {

                //SELECT ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.OPCH_SER  WHERE [ItemCode] i
                sb.Append(" UNION ALL");
                sb.Append("                                   SELECT Convert(varchar(8),T0.[DOCDate],112)    出貨日期,T44.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別");
                sb.Append(" ,t8.u_beneficiary 最終客戶,T8.U_ACME_MEMO   備註,'2' SEQ ");
                sb.Append("                    FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry   ");
                sb.Append("               INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15') ");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )             ");
                sb.Append("  left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum ) ");
                sb.Append("               INNER join OINV t44 on (t4.docentry=t44.docentry)  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
                sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");

                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");

                sb.Append(" SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號 ");
                sb.Append(" FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry    ");
                sb.Append(" INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15')  ");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )              ");
                sb.Append(" left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum )  ");
                sb.Append(" INNER join OINV t44 on (t4.docentry=t44.docentry)  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
                sb.Append(" WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'     ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {
                    sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                }
                if (checkBox1.Checked)
                {
                    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                }
                else
                {
                    if (textBox5.Text != "" && textBox6.Text != "")
                    {
                        sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                    }
                }
                if (checkBox2.Checked)
                {
                    sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                    }
                }
                sb.Append(" ) )) ");
                sb.Append(" UNION ALL");
                sb.Append("                                                 SELECT Convert(varchar(8),T0.[DOCDate],112)   出貨日期,T0.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別 ");
                sb.Append("               ,t8.u_beneficiary 最終客戶,T8.U_ACME_PAYGUI   發票金額,'2' SEQ  ");
                sb.Append("                                  FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry    ");
                sb.Append("                             INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')  ");
                sb.Append("               left join ordr t8 on (t8.docentry=T5.docentry  )              ");
                sb.Append("  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
                sb.Append("                             WHERE              ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");
                sb.Append(" SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號   ");
                sb.Append(" FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry     ");
                sb.Append(" INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')   ");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )               ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE     ");
                sb.Append(" WHERE              ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   ");

                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {
                    sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                }
                if (checkBox1.Checked)
                {
                    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                }
                else
                {
                    if (textBox5.Text != "" && textBox6.Text != "")
                    {
                        sb.Append(" and  T2.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                    }
                }
                if (checkBox2.Checked)
                {
                    sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                    }
                }

                sb.Append(" ) )) ");
            }
            sb.Append("  ORDER BY SEQ,出貨日期 DESC ");
     
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }
        public System.Data.DataTable PackOP4()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                          SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[SHIPDate],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別  ");
            sb.Append("  ,T0.u_beneficiary 最終客戶,T0.U_ACME_MEMO   備註,'1' SEQ                  FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032 ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T0.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }
            sb.Append("    UNION ALL ");

            sb.Append("                          SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[SHIPDate],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別  ");
            sb.Append("  ,T0.u_beneficiary 最終客戶,T0.U_ACME_MEMO   備註,'1' SEQ                    FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("    INNER JOIN (SELECT T1.DOCENTRY,QUANTITY,BASEREF,BASELINE FROM INV1 T0 LEFT JOIN OINV T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE BASETYPE=17 AND TRGETENTRY='' AND updinvnt ='c' AND CAST(T0.DOCENTRY AS VARCHAR) NOT IN (SELECT ISNULL(U_ACME_ARAP,'') FROM ORIN where doctype <> 's' AND ISNULL(U_ACME_ARAP,'') <>''   )) T5 ON(T5.BASEREF = T1.DOCENTRY AND T5.BASELINE=T1.LINENUM) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {
                sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
            }
            else
            {
                if (textBox5.Text != "" && textBox6.Text != "")
                {
                    sb.Append(" and  T0.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                }
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
            }
            else
            {
                if (textBox7.Text != "" && textBox8.Text != "")
                {
                    sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                }
            }
            if (checkBox3.Checked)
            {
                sb.Append("    UNION ALL ");
                sb.Append("                          SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[SHIPDate],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別  ");
                sb.Append("  ,T0.u_beneficiary 最終客戶,T0.U_ACME_MEMO   備註,'2' SEQ                  FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
                sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032 ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");

                sb.Append("                          SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號 ");
                sb.Append("      FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("  left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
                sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032 ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                }
                if (textBox3.Text != "" && textBox4.Text != "")
                {
                    sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                }
                if (checkBox1.Checked)
                {
                    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                }
                else
                {
                    if (textBox5.Text != "" && textBox6.Text != "")
                    {
                        sb.Append(" and  T0.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                    }
                }
                if (checkBox2.Checked)
                {
                    sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
                }
                else
                {
                    if (textBox7.Text != "" && textBox8.Text != "")
                    {
                        sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                    }
                }

                sb.Append(" ) )) ");
             //  sb.Append("    UNION ALL ");

                //
                //sb.Append("                          SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[SHIPDate],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別  ");
                //sb.Append("  ,T0.u_beneficiary 最終客戶,T0.U_ACME_MEMO   備註,'2' SEQ                    FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                //sb.Append("    INNER JOIN (SELECT T1.DOCENTRY,QUANTITY,BASEREF,BASELINE FROM INV1 T0 LEFT JOIN OINV T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE BASETYPE=17 AND TRGETENTRY='' AND updinvnt ='c' AND CAST(T0.DOCENTRY AS VARCHAR) NOT IN (SELECT ISNULL(U_ACME_ARAP,'') FROM ORIN where doctype <> 's' AND ISNULL(U_ACME_ARAP,'') <>''  )) T5 ON(T5.BASEREF = T1.DOCENTRY AND T5.BASELINE=T1.LINENUM) ");
                //sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");
                //sb.Append("                          SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號 ");
                //sb.Append("    FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry  ");
                //sb.Append("    INNER JOIN (SELECT T1.DOCENTRY,QUANTITY,BASEREF,BASELINE FROM INV1 T0 LEFT JOIN OINV T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE BASETYPE=17 AND TRGETENTRY='' AND updinvnt ='c' AND CAST(T0.DOCENTRY AS VARCHAR) NOT IN (SELECT ISNULL(U_ACME_ARAP,'') FROM ORIN where doctype <> 's' AND ISNULL(U_ACME_ARAP,'') <>''  )) T5 ON(T5.BASEREF = T1.DOCENTRY AND T5.BASELINE=T1.LINENUM) ");
                //if (textBox1.Text != "" && textBox2.Text != "")
                //{
                //    sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
                //}
                //if (textBox3.Text != "" && textBox4.Text != "")
                //{
                //    sb.Append(" and  Convert(varchar(8),T0.[TAXDate],112) between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
                //}
                //if (checkBox1.Checked)
                //{
                //    sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
                //}
                //else
                //{
                //    if (textBox5.Text != "" && textBox6.Text != "")
                //    {
                //        sb.Append(" and  T0.[docentry] between '" + textBox9.Text.ToString() + "' and '" + textBox10.Text.ToString() + "' ");
                //    }
                //}
                //if (checkBox2.Checked)
                //{
                //    sb.Append(" and  T1.[ItemCode] in ( " + d + ") ");
                //}
                //else
                //{
                //    if (textBox7.Text != "" && textBox8.Text != "")
                //    {
                //        sb.Append(" and  T1.[ItemCode] between '" + textBox7.Text.ToString() + "' and '" + textBox8.Text.ToString() + "' ");
                //    }
                //}

                //sb.Append(" ) )) ");
            }
            sb.Append("  ORDER BY SEQ,過帳日期 DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPOR"];
        }

        private void APS_Load(object sender, EventArgs e)
        {
         
        }
        public string c;
        private void button3_Click(object sender, EventArgs e)
        {
                 APS1 frm1 = new APS1();
                 if (frm1.ShowDialog() == DialogResult.OK)
                 {
                     checkBox1.Checked = true;
                     c = frm1.q;

                 }
       

        }
          public string d;
        private void button4_Click(object sender, EventArgs e)
        {
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;
               
            }
       
        }

        private void textBox7_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetOitm();

            if (LookupValues != null)
            {
                textBox7.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox8_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.Get0itm();

            if (LookupValues != null)
            {
                textBox8.Text = Convert.ToString(LookupValues[0]);
       
            }
        }
        public string e,f;
        private void textBox5_DoubleClick(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetOcrd();

            if (LookupValues != null)
            {
                textBox5.Text = Convert.ToString(LookupValues[1]);
                textBox9.Text = Convert.ToString(LookupValues[0]);

            }
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetOcrd();

            if (LookupValues != null)
            {
                textBox6.Text = Convert.ToString(LookupValues[1]);
                textBox10.Text = Convert.ToString(LookupValues[0]);

            }
        }



        private void button5_Click(object sender, EventArgs e)
        {
            //string aa;
            //if (textBox3.Text != "")
            //{
            //    aa = "日期別已進狀況表";
            //}
            //else if (textBox5.Text != "" || checkBox1.Checked)
            //{
            //    aa = "廠商別已進狀況表";
            //}
            //else
            //{
            //    aa = "產品別已進狀況表";
            //}

            
            //ACME.ApsCry3 frm4 = new ACME.ApsCry3();
            //frm4.dt = PackOP2();
            //frm4.s = aa;
            //frm4.Show();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("yyyy");
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Text = DateTime.Now.ToString("yyyy");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //string aa;
            //if (textBox3.Text != "")
            //{
            //    aa = "日期別已出狀況表";
            //}
            //else if (textBox5.Text != "" || checkBox1.Checked)
            //{
            //    aa = "廠商別已出狀況表";
            //}
            //else
            //{
            //    aa = "產品別已出況表";
            //}


            //ACME.ApsCry3 frm4 = new ACME.ApsCry3();
            //frm4.dt = PackOP3();
            //frm4.s = aa;
            //frm4.Show();
        }



        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
             GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
            GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
              GridViewToExcel(dataGridView3);
            }
        }
        public static void GridViewToExcel(DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();
            wapp.DisplayAlerts = false;
            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;

            try
            {

                for (int i = 0; i < dgv.Columns.Count-1; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j < row.Cells.Count-1; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                        }

                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);

                        }

                    }

                }

                wapp.Visible = true;


            }

            catch (Exception ex1)
            {

                MessageBox.Show(ex1.Message);

            }

            wapp.UserControl = true;
            const int MAX_EVALUATE_LINES = 10;
        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                CalcTotals1();
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                CalcTotals2();
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                CalcTotals4();
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                CalcTotals3();
            }
        }

        private void CalcTotals1()
        {

            Int32 iTotal = 0;
       

            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["QTY1"].Value);

            }
            textBox11.Text = iTotal.ToString("#,##0");

        }
        private void CalcTotals2()
        {

            Int32 iTotal = 0;


            int i = this.dataGridView2.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView2.SelectedRows[iRecs].Cells["QTY2"].Value);

            }
            textBox11.Text = iTotal.ToString("#,##0");

        }
        private void CalcTotals3()
        {

            Int32 iTotal = 0;


            int i = this.dataGridView3.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView3.SelectedRows[iRecs].Cells["QTY3"].Value);

            }

            textBox11.Text = iTotal.ToString("#,##0");

        }
        private void CalcTotals4()
        {

            Int32 iTotal = 0;


            int i = this.dataGridView4.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView4.SelectedRows[iRecs].Cells["QTY5"].Value);

            }

            textBox11.Text = iTotal.ToString("#,##0");

        }



 
     
    }
}
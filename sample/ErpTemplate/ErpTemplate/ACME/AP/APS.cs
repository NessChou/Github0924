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

    public partial class APS : Form
    {
        string strCn = "";
        public string s;
        public APS()
        {
            InitializeComponent();
        }
        System.Data.DataTable aa;
        private void button1_Click(object sender, EventArgs e)
        {
           
            if (comboBox1.Text =="進金生")
            {
                strCn = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            }

            if (comboBox1.Text == "達睿生")
            {
                strCn = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            }
            if (comboBox1.Text == "測試98")
            {
                strCn = "Data Source=acmesap;Initial Catalog=acmesql98;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            }
            dataGridView1.DataSource = PackOP();
            dataGridView7.DataSource = PackOPQ();
            if (comboBox1.Text == "達睿生")
            {
                dataGridView2.DataSource = PackOP2DRS();
            }
            else
            {
                dataGridView2.DataSource = PackOP2();
            }
            dataGridView3.DataSource = PackOP3();
            dataGridView4.DataSource = PackOP4();
            dataGridView6.DataSource = PackA();
        }
        public System.Data.DataTable PackOP()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
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
        public System.Data.DataTable PackOPQ()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T0.[TAXDate],112)  文件日期,T0.DOCENTRY 採購報價單號,T0.CARDNAME 廠商名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[OpenCreQty] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ");
            sb.Append(" ,T1.U_MEMO  備註,'1' SEQ FROM OPQT T0  INNER JOIN PQT1 T1 ON T0.DocEntry = T1.DocEntry ");
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
                sb.Append(" ,T1.U_MEMO  備註,'2' SEQ FROM OPQT T0  INNER JOIN PQT1 T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode    left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  WHERE T1.[LINESTATUS] ='O'   and ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");
                sb.Append("               SELECT DISTINCT T1.ITEMCODE  COLLATE Chinese_PRC_CI_AS 產品編號 FROM OPQT T0  INNER JOIN PQT1 T1 ON T0.DocEntry = T1.DocEntry  ");
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
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T0.CARDNAME 廠商名稱, ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
            sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
            sb.Append(" ELSE   T1.ITEMCODE END 產品編號, ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm'");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm'");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm'");
            sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz' ");
            sb.Append(" ELSE  T1.[Dscription]  END 品名規格, ");
            sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 , ");
            sb.Append(" T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'1' SEQ,T5.LINENUM  LINENUM2 ");
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
            sb.Append(" union all");
            sb.Append(" SELECT Convert(varchar(8),T0.[DOCDate],112) 進貨日期,T0.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T2.U_CARDNAME  廠商名稱,  ");
            sb.Append(" T1.ITEMCODE 產品編號,T1.[Dscription]  品名規格,  ");
            sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,  ");
            sb.Append(" '' 原廠進貨匯率,'' 發票號碼,''AP單號,T1.U_MEMO 備註,'1' SEQ,''   ");
            sb.Append(" FROM OIGN T0  INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    LEFT JOIN OWOR T2 ON (T1.BASEENTRY=T2.DOCENTRY)  ");
            sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  AND T1.BaseType =202");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DOCDate],112)  between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
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
                sb.Append(" SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T0.CARDNAME 廠商名稱,");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
                sb.Append(" ELSE   T1.ITEMCODE END 產品編號, ");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz' ");
                sb.Append(" ELSE  T1.[Dscription]  END 品名規格, ");
                sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,");
                sb.Append(" T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'2' SEQ,T5.LINENUM  LINENUM2  ");
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
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
                    sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
                        sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
        public System.Data.DataTable PackOP2DRS()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T0.CARDNAME 廠商名稱,  ");
            sb.Append("              CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6'  ");
            sb.Append("              WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6'  ");
            sb.Append("              WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6'  ");
            sb.Append("              WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P'  ");
            sb.Append("              ELSE   T1.ITEMCODE END 產品編號,  ");
            sb.Append("              CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm' ");
            sb.Append("              WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm' ");
            sb.Append("              WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm' ");
            sb.Append("              WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz'  ");
            sb.Append("              ELSE  T1.[Dscription]  END 品名規格,  ");
            sb.Append("              T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,  ");
            sb.Append("              T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'1' SEQ,T5.LINENUM  LINENUM2  ");
            sb.Append("              FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append("              left join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)    ");
            sb.Append("              left join opdn t44 on (t4.docentry=t44.docentry)    ");
            sb.Append("              left JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode     ");
            sb.Append("              left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
            sb.Append("              left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)    ");
            sb.Append("              left join OPCH t55 on (t5.docentry=t55.docentry)    ");
            sb.Append("              left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )    ");
            sb.Append("              left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )    ");
            sb.Append("              WHERE T0.DOCSTATUS='C'");

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
            sb.Append(" union all");
            sb.Append(" SELECT Convert(varchar(8),T0.[DOCDate],112) 進貨日期,T0.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T2.U_CARDNAME  廠商名稱,  ");
            sb.Append(" T1.ITEMCODE 產品編號,T1.[Dscription]  品名規格,  ");
            sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,  ");
            sb.Append(" '' 原廠進貨匯率,'' 發票號碼,''AP單號,T1.U_MEMO 備註,'1' SEQ,''   ");
            sb.Append(" FROM OIGN T0  INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    LEFT JOIN OWOR T2 ON (T1.BASEENTRY=T2.DOCENTRY)  ");
            sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  AND T1.BaseType =202");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  T0.[Docnum] between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "" && textBox4.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DOCDate],112)  between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "' ");
            }

            if (checkBox1.Checked)
            {
                sb.Append(" and  T0.[CardCode] in ( " + c + ") ");
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
                sb.Append(" SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T1.LINENUM,T0.CARDNAME 廠商名稱,");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6' ");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6' ");
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
                sb.Append(" ELSE   T1.ITEMCODE END 產品編號, ");
                sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm'");
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz' ");
                sb.Append(" ELSE  T1.[Dscription]  END 品名規格, ");
                sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T11.AVGPRICE 平均成本 ,");
                sb.Append(" T44.U_ACME_RATE1 原廠進貨匯率,T55.U_PC_BSINV 發票號碼,t5.docentry AP單號,T1.U_MEMO 備註,'2' SEQ,T5.LINENUM  LINENUM2  ");
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
                sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
                    sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
                        sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P' ");
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
        public System.Data.DataTable PackA()
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT Convert(varchar(8),T44.[DOCDate],112) 進貨日期,t4.DOCENTRY 收採單號,T12.WHSNAME 倉庫,T44.U_ACME_INV  INV,");
            sb.Append("CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6'  ");
            sb.Append("WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6'  ");
            sb.Append("WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6'  ");
            sb.Append("WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P'  ");
            sb.Append("ELSE   T1.ITEMCODE END 產品編號,  ");
            sb.Append("CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm' ");
            sb.Append("WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm' ");
            sb.Append("WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm' ");
            sb.Append("WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz'  ");
            sb.Append("ELSE  T1.[Dscription]  END 品名規格,CAST(T1.[Quantity] AS INT) 數量 ");
            sb.Append("FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append("INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)    ");
            sb.Append("INNER join opdn t44 on (t4.docentry=t44.docentry)    ");
            sb.Append("INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode     ");
            sb.Append("left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
            sb.Append("left join PCH1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum AND T5.BASETYPE=20)    ");
            sb.Append("left join OPCH t55 on (t5.docentry=t55.docentry)    ");
            sb.Append("left join RPD1 t8 on (t8.baseentry=T4.docentry and  t8.baseline=t4.linenum and t8.basetype='20'  )    ");
            sb.Append("left join RPC1 t9 on (t9.baseentry=T5.docentry and  t9.baseline=t5.linenum and t9.basetype='18'  )    ");
            sb.Append("LEFT JOIN OWHS T12 ON (T4.WHSCODE=T12.WHSCODE)");
            sb.Append("WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND (T1.[Quantity]-ISNULL(T8.QUANTITY,0) <> 0)  AND (T1.[Quantity]-ISNULL(T9.QUANTITY,0) <> 0)      ");
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


            sb.Append("  ORDER BY 進貨日期 DESC ");
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
        public System.Data.DataTable PackOP2M(string d,string DOCTYPE)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  Convert(varchar(8),T44.[DOCDate],112)  進貨日期,T44.U_ACME_INV  INV,T0.DOCENTRY 採購單號,T0.CARDNAME 廠商名稱,  ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'O270HVR01.110A6'  ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6'  ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6'  ");
            sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'O320HVN05.0000P'  ");
            sb.Append(" ELSE   T1.ITEMCODE END 產品編號,  ");
            sb.Append(" CASE WHEN T1.DocEntry =38309 AND T1.LineNum=1 THEN 'OPEN CELL_AU M270HVR01.10A-P (91.27M18.10A) POL-1mm' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=15 THEN 'O270HVR01.110A6- OPEN CELL _AU M270HVR01 V.10A-P POL-1mm' ");
            sb.Append(" WHEN T1.DocEntry =38657 AND T1.LineNum=7 THEN 'O270HVR01.510A6- OPEN CELL _AU M270HVR01 V.10A-N POL-1mm' ");
            sb.Append(" WHEN T1.DocEntry =41222 AND T1.LineNum=0 THEN 'OPEN CELL_AU P320HVN05.000-Z (91.32P15.000)_144Hz / 165Hz'  ");
            sb.Append(" ELSE  T1.[Dscription]  END 品名規格,  ");
            sb.Append(" T1.[Quantity] 數量,T1.PRICE 單價,");
            if (DOCTYPE == "1")
            {
                sb.Append(" (T1.QUANTITY*T1.PRICE) 小計");
            }
            if (DOCTYPE == "2")
            {
                sb.Append(" (T1.QUANTITY*T1.PRICE*1.05) 小計");
            }
            sb.Append(" FROM OPOR T0  INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" INNER join PDN1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BASETYPE=22)    ");
            sb.Append(" INNER join opdn t44 on (t4.docentry=t44.docentry)    ");
            sb.Append(" where  cast(T1.docentry as varchar)+' '+cast(T1.LINENUM as varchar) IN ( " + d + ") ");
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
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("                                   SELECT Convert(varchar(8),T0.[DOCDate],112)    出貨日期,T44.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別,T5.U_FINALPRICE  最終售價");
            sb.Append(" ,T8.U_ACME_PAYGUI 美金備註,t8.u_beneficiary 最終客戶,T8.U_ACME_MEMO   備註,'1' SEQ,(T3.[lastName]+T3.[firstName]) 業管  ");
            sb.Append("                    FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("               INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )             ");
            sb.Append("  left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum AND T4.BaseType =15) ");
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
            sb.Append("                                                 SELECT Convert(varchar(8),T0.[DOCDate],112)   出貨日期,T0.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別,T5.U_FINALPRICE  最終售價 ");
            sb.Append("               ,T8.U_ACME_PAYGUI 美金備註,t8.u_beneficiary 最終客戶,T8.U_ACME_PAYGUI 備註,'1' SEQ,(T3.[lastName]+T3.[firstName]) 業管");
            sb.Append("                                  FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("                             INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum   and T1.BaseType  ='17')  ");
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

                sb.Append(" UNION ALL");
                sb.Append("                                   SELECT Convert(varchar(8),T0.[DOCDate],112)    出貨日期,T44.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別,T5.U_FINALPRICE  最終售價");
                sb.Append(" ,T8.U_ACME_PAYGUI 美金備註,t8.u_beneficiary 最終客戶,T8.U_ACME_MEMO 備註,'2' SEQ,(T3.[lastName]+T3.[firstName]) 業管 ");
                sb.Append("                    FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry   ");
                sb.Append("               INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15') ");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )             ");
                sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
                sb.Append("  left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BaseType =15) ");
                sb.Append("               INNER join OINV t44 on (t4.docentry=t44.docentry)  INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
                sb.Append("               WHERE            ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'  ");

                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in (");

                sb.Append(" SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號 ");
                sb.Append(" FROM ODLN T0  INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry    ");
                sb.Append(" INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='15')  ");
                sb.Append(" left join ordr t8 on (t8.docentry=T5.docentry  )              ");
                sb.Append(" left join INV1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  AND T4.BaseType =15)  ");
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
                sb.Append("                                                 SELECT Convert(varchar(8),T0.[DOCDate],112)   出貨日期,T0.U_IN_BSINV  INV,Convert(varchar(8),T0.[TAXDate],112)  文件日期,T5.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T5.PRICE 單價,T5.[Currency] 幣別,T5.U_FINALPRICE  最終售價 ");
                sb.Append("              ,T8.U_ACME_PAYGUI 美金備註,t8.u_beneficiary 最終客戶,T8.U_ACME_PAYGUI 備註,'2' SEQ,(T3.[lastName]+T3.[firstName]) 業管  ");
                sb.Append("                                  FROM OINV T0  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry    ");
                sb.Append("                             INNER join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13')  ");
                sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
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
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[U_ACME_SHIPDAY],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號,  ");
            sb.Append(" T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,T1.[Currency] 幣別,T2.WhsName  倉庫,(T3.[lastName]+T3.[firstName]) 業管 ,T0.u_beneficiary 最終客戶,T1.U_FINALPRICE  最終售價    ");
            sb.Append(" ,T0.U_ACME_PAYGUI 美金備註,T0.U_ACME_MEMO   備註,'1' SEQ,T1.U_SHIPSTATUS 貨況 ,T0.U_ACME_IIS IIS                     ");
            sb.Append(" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE    ");
            sb.Append(" LEFT JOIN OWHS T2 ON (T1.WhsCode =T2.WhsCode) ");
            sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append(" WHERE ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032   ");

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
            //sb.Append("    UNION ALL ");
            //sb.Append(" SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[U_ACME_SHIPDAY],112)  訂單日期, ");
            //sb.Append(" T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量, ");
            //sb.Append(" T1.PRICE 單價,T1.[Currency] 幣別,T2.WhsName  倉庫,T1.U_FINALPRICE  最終售價    ");
            //sb.Append(" ,T0.U_ACME_PAYGUI 美金備註,T0.u_beneficiary 最終客戶,T0.U_ACME_MEMO   備註,'1' SEQ,(T3.[lastName]+T3.[firstName]) 業管                  ");
            //sb.Append(" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry    ");
            //sb.Append(" LEFT JOIN OWHS T2 ON (T1.WhsCode =T2.WhsCode) ");
            //sb.Append(" INNER JOIN (SELECT T1.DOCENTRY,QUANTITY,BASEREF,BASELINE FROM INV1 T0  ");
            //sb.Append(" LEFT JOIN OINV T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            //sb.Append(" WHERE BASETYPE=17 AND TRGETENTRY='' AND updinvnt ='c'  ");
            //sb.Append(" AND CAST(T0.DOCENTRY AS VARCHAR) NOT IN (SELECT ISNULL(U_ACME_ARAP,'') FROM ORIN where doctype <> 's' AND ISNULL(U_ACME_ARAP,'') <>''   )) T5 ON(T5.BASEREF = T1.DOCENTRY AND T5.BASELINE=T1.LINENUM)   ");
            //sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
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
            if (checkBox3.Checked)
            {
                sb.Append("    UNION ALL ");
                sb.Append(" SELECT Convert(varchar(8),T0.[DOCDate],112)  過帳日期,Convert(varchar(8),T1.[U_ACME_SHIPDAY],112)  訂單日期,T1.DOCENTRY 銷售單號,T0.CARDNAME 客戶名稱,T1.ITEMCODE 產品編號, T1.[Dscription] 品名規格, T1.[Quantity] 數量,T1.PRICE 單價,");
                sb.Append(" T1.[Currency] 幣別 ,T2.WhsName  倉庫,(T3.[lastName]+T3.[firstName]) 業管,T0.u_beneficiary 最終客戶,T1.U_FINALPRICE  最終售價  ");
                sb.Append(" ,T0.U_ACME_PAYGUI 美金備註,T0.U_ACME_MEMO   備註,'2' SEQ,T1.U_SHIPSTATUS 貨況 ,T0.U_ACME_IIS IIS                 ");
                sb.Append(" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry   ");
                sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
                sb.Append(" LEFT JOIN OWHS T2 ON (T1.WhsCode =T2.WhsCode)");
                sb.Append(" LEFT JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
                sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032  ");
                sb.Append(" AND  (T1.[ItemCode] COLLATE  Chinese_Taiwan_Stroke_CI_AS  in (SELECT KIT COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.AP_OPENCELL where opencell  in ( ");
                sb.Append(" SELECT DISTINCT T1.ITEMCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS 產品編號  ");
                sb.Append(" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry   ");
                sb.Append(" left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE   ");
                sb.Append(" WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   AND T1.LINESTATUS='O' AND T11.ITMSGRPCOD=1032  ");
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
            comboBox1.Text = "進金生";
            checkBox4.Checked = false;
            string USER = fmLogin.LoginID.ToString().ToUpper();

            if (USER == "CLOUDIAWU" || USER == "MAGGIEWENG")
            {
                button8.Visible = true;
            }
            else
            {
                button8.Visible = false;
            }
            if (USER != "APPLECHEN" && USER != "LLEYTONCHEN")
            {
                tabControl1.TabPages.Remove(tabControl1.TabPages["APPLE查詢"]);
            }
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
            else if(tabControl1.SelectedIndex == 1)
            {
                GridViewToExcel(dataGridView7);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
             GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 3)
            {
            GridViewToExcel(dataGridView4);
            }
            else if (tabControl1.SelectedIndex == 4)
            {
              GridViewToExcel(dataGridView3);
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                ExcelReport.GridViewToExcel(dataGridView6);
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
                CalcTotals2();
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                CalcTotals4();
            }
            else if (tabControl1.SelectedIndex == 4)
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

        private void button6_Click_1(object sender, EventArgs e)
        {
            GetMenu.InsertEXEXPORT("07採購線上庫存", "", fmLogin.LoginID.ToString(), DateTime.Now.ToString("yyyyMMddHHmmss"), "");

            MessageBox.Show("訊息已送出");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    StringBuilder sb = new StringBuilder();
                    for (int i = dataGridView2.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView2.SelectedRows[i];

                        sb.Append("'" + row.Cells["採購單號"].Value.ToString() + " " + row.Cells["LINENUM"].Value.ToString() + "',");

                    }


                    sb.Remove(sb.Length - 1, 1);
                    string q = sb.ToString();
                    string DOCTYPE="1";
                    if(checkBox4.Checked)
                    {
                        DOCTYPE = "2";
                    }
                    System.Data.DataTable GG1 = PackOP2M(q, DOCTYPE);
                    if (GG1.Rows.Count > 0)
                    {
                        decimal[] Total = new decimal[GG1.Columns.Count - 1];

                        for (int i = 0; i <= GG1.Rows.Count - 1; i++)
                        {

                     
                                Total[7] += Convert.ToDecimal(GG1.Rows[i][8]);

                            
                        }

                        DataRow row2;

                        row2 = GG1.NewRow();

                            row2[8] = Total[7];
                            string HEAD = "請查收附檔水單 US$" + Total[7].ToString("#,##0.00");
                        
                        GG1.Rows.Add(row2);

                        dataGridView5.DataSource = GG1;
                        string GG = htmlMessageBody(dataGridView5).ToString();
                        string EMAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                        MailTest2(HEAD, EMAIL, GG, Total[7].ToString("#,##0.00"));
                        MessageBox.Show("寄信成功");
                    
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

        
            //string TABLE = "";
            //try
            //{

            //    if (tabControl1.SelectedIndex == 0)
            //    {
            //        TABLE = "AP_KIT";
            //        if (dataGridView2.SelectedRows.Count == 0)
            //        {
            //            MessageBox.Show("請點選單號");
            //            return;
            //        }

            //    }
            
            //    ArrayList al = new ArrayList();

            //    for (int i = 0; i <= listBox1.Items.Count - 1; i++)
            //    {
            //        al.Add(listBox1.Items[i].ToString());
            //    }


            //    StringBuilder sb = new StringBuilder();



            //    foreach (string v in al)
            //    {
            //        sb.Append("'" + v + "',");
            //    }

            //    sb.Remove(sb.Length - 1, 1);


            //    q = sb.ToString();


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}


            //string MAIL = "";
            //string SUBJECT = "";
            //string SA = "";
            //System.Data.DataTable G1 = null;
            //if (comboBox1.Text == "進貨通知")
            //{

            //    if (tabControl1.SelectedIndex == 0)
            //    {
            //        MAIL = "\\MailTemplates\\KIT1.htm";

            //    }
            //    else
            //    {
            //        MAIL = "\\MailTemplates\\KIT4.htm";
            //    }
            //    G1 = Getbb(q, TABLE);
            //    SUBJECT = G1.Rows[0]["廠商"].ToString() + "進貨內湖，請查收，謝謝!!";
            //    SA = G1.Rows[0]["SA"].ToString();
            //}
  

            //if (G1.Rows.Count > 0)
            //{
            //    string DOCTYPE = "L1";
            //    if (comboBox1.Text == "廠商訂單")
            //    {
            //        DOCTYPE = "L2";
            //    }
            //    string CARDNAME = G1.Rows[0]["廠商"].ToString();
            //    dataGridView1.DataSource = G1;
            //    string GG = htmlMessageBody(dataGridView1).ToString();
            //    string EMAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
            //    MailTest2(SUBJECT, EMAIL, GG, CARDNAME, MAIL, SA, DOCTYPE);
            //    MessageBox.Show("寄信成功");
            //}
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

                //if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                //{



                //    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                //    tmpKeyValue = KeyValue;
                //}
                //else
                //{
                //    tmpKeyValue = "";
                //}
                KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                tmpKeyValue = KeyValue;

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
        private void MailTest2(string strSubject, string MailAddress, string MailContent, string CUST2)
        {
            MailMessage message = new MailMessage();
            string FROM = fmLogin.LoginID.ToString() + "@acmepoint.com";
            message.From = new MailAddress(FROM, "系統發送");

            message.To.Add(new MailAddress(MailAddress));

            string template;
            StreamReader objReader;
            string GetExePath = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //"\\MailTemplates\\KIT1.htm"
            objReader = new StreamReader(GetExePath + "\\MailTemplates\\AP3.htm");

            template = objReader.ReadToEnd();
            objReader.Close();

            template = template.Replace("##Content##", MailContent);
            template = template.Replace("##CUST1##", "請查收附檔水單 US$");
            template = template.Replace("##CUST2##", CUST2);
   
            string USER = fmLogin.LoginID.ToString();

            message.Subject = strSubject;
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
     

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
        
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {

                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }

                }
            }
            catch (Exception ex)
            {

            }

        }

        private void dataGridView2_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {

                if (dataGridView2.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;

                    StringBuilder sb = new StringBuilder();

                    for (int i = dataGridView2.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = dataGridView2.SelectedRows[i];
                        sb.Append("'" + row.Cells["採購單號"].Value.ToString() + " " + row.Cells["LINENUM"].Value.ToString() + "',");

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int i = this.dataGridView2.Rows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {

                string DOCENTRY = dataGridView2.Rows[iRecs].Cells["採購單號"].Value.ToString();
                string DOCENTRY2 = dataGridView2.Rows[iRecs].Cells["AP單號"].Value.ToString().Trim();
                string U_MEMO = dataGridView2.Rows[iRecs].Cells["備註2"].Value.ToString();
                string 品名規格 = dataGridView2.Rows[iRecs].Cells["品名規格"].Value.ToString();
                string LINENUM = dataGridView2.Rows[iRecs].Cells["LINENUM"].Value.ToString();
                string LINENUM2 = dataGridView2.Rows[iRecs].Cells["LINENUM2"].Value.ToString();
                
                if (DOCENTRY2 != "0")
                {
                    UPLC(DOCENTRY, LINENUM, DOCENTRY2, LINENUM2, U_MEMO, 品名規格);
                }

            }

            MessageBox.Show("備註已更新");
        }

        private void UPLC(string DOCENTRY, string LINENUM, string DOCENTRY2, string LINENUM2, string U_MEMO, string DSCRIPTION)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE POR1 SET U_MEMO=@U_MEMO,DSCRIPTION=@DSCRIPTION WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM");
            sb.Append(" UPDATE PDN1 SET U_MEMO=@U_MEMO,DSCRIPTION=@DSCRIPTION  WHERE BASEENTRY=@DOCENTRY AND BASELINE=@LINENUM ");
            sb.Append(" UPDATE PCH1 SET U_MEMO=@U_MEMO,DSCRIPTION=@DSCRIPTION  WHERE DOCENTRY=@DOCENTRY2 AND LINENUM=@LINENUM2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@DOCENTRY2", DOCENTRY2));
            command.Parameters.Add(new SqlParameter("@LINENUM2", LINENUM2));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@DSCRIPTION", DSCRIPTION));
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

 
    }
}
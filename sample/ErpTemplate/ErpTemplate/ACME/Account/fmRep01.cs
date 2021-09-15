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

//20110104 加入格式化
//20110104 加入銷售訂單金額

namespace ACME
{
    public partial class fmRep01 : Form
    {
        public fmRep01()
        {
            InitializeComponent();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string DocNum1 = txtDocNo1.Text;
            string DocNum2 = txtDocNo2.Text;


            string DocDate1 = txtDocDate1.Text;
            string DocDate2 = txtDocDate2.Text;

            string CLOSE1 = texCLOSE1.Text;
            string CLOSE2 = texCLOSE2.Text;

            System.Data.DataTable dtData = GetOworData(DocNum1, DocNum2, DocDate1, DocDate2, CLOSE1, CLOSE2);
            gvData.DataSource = dtData;



            System.Data.DataTable dtDetail = GetOworDetail(DocNum1, DocNum2, DocDate1, DocDate2, CLOSE1, CLOSE2);
            gvDataDetail.DataSource = dtDetail;


            for (int i = 0; i <= gvData.Columns.Count - 1; i++)
            {


                if (dtData.Columns[i].ColumnName == "銷售單號" || dtData.Columns[i].ColumnName == "生產訂單")
                {
                    continue;
                }

                if (dtData.Columns[i].DataType == typeof(Int32))
                {

                    SetDefaultStyle_Int(gvData.Columns[i]);

                }
                else   if (dtData.Columns[i].DataType == typeof(Decimal))
                {
                    SetDefaultStyle_0(gvData.Columns[i]);
                }

            }


            for (int i = 0; i <= gvDataDetail.Columns.Count - 1; i++)
            {


                if (dtDetail.Columns[i].ColumnName == "銷售單號" || dtDetail.Columns[i].ColumnName == "生產訂單")
                {
                    continue;
                }

                if (dtDetail.Columns[i].DataType == typeof(Int32))
                {

                    SetDefaultStyle_Int(gvDataDetail.Columns[i]);

                }
                else if (dtDetail.Columns[i].DataType == typeof(Decimal))
                {
                    SetDefaultStyle_0(gvDataDetail.Columns[i]);
                }

            }




        }
        public System.Data.DataTable GetOworData(string DocNum1, string DocNum2, string DocDate1, string DocDate2, string CLOSE1, string CLOSE2)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            sb.Append("                                         SELECT  T0.[OriginNum]  銷售單號,''''+T0.[CardCode] 客戶編號, T55.[CardName] 客戶名稱, ");
            sb.Append("                                                     T0.[u_projectcode] 專案代號, T56.PRJNAME 專案名稱, ");
            sb.Append("                                                     T0.[PostDate] 生產訂單日期, T0.[DueDate] 到期日期,T0.CLOSEDATE 完成日期,DATEDIFF(D,T0.[PostDate],T0.CLOSEDATE) 完成天數,T1.[lastpurdat] 最近購買日期, T3.U_Name Owner,T0.[DocNum] 生產訂單, ");
            sb.Append("                                                     狀態 =Case  when T0.[Status]='P' then '計劃'  when T0.[Status]='R' then '已核發' when T0.[Status]='L' then '結案' end,  ");
            sb.Append("                                                     T0.[ItemCode] 產品編號, T1.[ItemName] 產品名稱, T0.[PlannedQty] 計劃數量,T0.[CmpltQty] 完成數量,  ");
            sb.Append("                                                     已發貨成本 = (SELECT abs(Convert(int,Sum(T7.[TransValue]))) FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum  AND  T7.[AppObjType] = 'C'),  ");
            sb.Append("                                                     實際產品成本 = (SELECT abs(Convert(int,Sum(T7.[TransValue]))) FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum  AND  T7.[AppObjType] = 'P'),  ");
            sb.Append("                                                     T2.DocTotal-T2.VatSum-ISNULL(TOTAL,0)-ISNULL(DISCTOTAL,0) 銷售訂單金額 ");
            sb.Append("                                                     FROM OWOR T0  ");
            sb.Append("                                                     INNER JOIN OITM T1 ON T0.ItemCode= T1.ItemCode ");
            sb.Append("                                                     Left join ORDR T2 on T2.DocEntry=T0.[OriginNum] ");
            sb.Append("                                                     Left join OUSR T3 on T3.Userid=T0.[UserSign] ");
            sb.Append("                              Left join OCRD T55 on T0.CARDCODE=T55.CARDCODE ");
            sb.Append("                         Left join OPRJ T56 on T0.[u_projectcode]=T56.PRJCODE ");
            sb.Append("                                        LEFT JOIN (SELECT SUM(T4.DocTotal) TOTAL,T1.DOCENTRY FROM RDR1 T1 ");
            sb.Append("                                        INNER JOIN INV1 T2 ON (T2.baseentry=T1.docentry and  T2.baseline=T1.linenum  and T1.targettype='13') ");
            sb.Append("                                        INNER JOIN OINV T5 ON (T2.DOCENTRY=T5.DOCENTRY)                                        ");
            sb.Append(" INNER JOIN RIN1 T3 ON (T3.baseentry=T2.docentry and  T3.baseline=T2.linenum  and T2.targettype='14') ");
            sb.Append("                                        INNER JOIN ORIN T4 ON (T3.DOCENTRY=T4.DOCENTRY) ");
            sb.Append("                                      WHERE T5.UPDINVNT='C'  GROUP BY T1.DOCENTRY) T4 ON (T2.DOCENTRY=T4.DOCENTRY) ");
            sb.Append("                         LEFT JOIN (SELECT MAX(T3.DISCSUM) DISCTOTAL,T1.DOCENTRY FROM RDR1 T1 ");
            sb.Append("                                        LEFT JOIN INV1 T2 ON (T2.baseentry=T1.docentry and  T2.baseline=T1.linenum  and T1.targettype='13') ");
            sb.Append("                                        LEFT JOIN OINV T3 ON (T2.DOCENTRY=T3.DOCENTRY) ");
            sb.Append("                                        GROUP BY T1.DOCENTRY) T5 ON (T2.DOCENTRY=T5.DOCENTRY) ");
            sb.Append("                                             Where 1=1   ");

        
            if (!string.IsNullOrEmpty(txtPrj1.Text))
            {
                sb.Append(" AND T0.[u_projectcode] >=@u_projectcode1 ");
                command.Parameters.Add(new SqlParameter("@u_projectcode1", txtPrj1.Text));
            }


            if (!string.IsNullOrEmpty(txtPrj2.Text))
            {
                sb.Append(" AND T0.[u_projectcode] <=@u_projectcode2 ");
                command.Parameters.Add(new SqlParameter("@u_projectcode2", txtPrj2.Text));
            }



            if (!string.IsNullOrEmpty(txtDocNo1.Text))
            {
                sb.Append(" AND T0.[DocNum] >=@DocNum1 ");
                command.Parameters.Add(new SqlParameter("@DocNum1", DocNum1));
            }

            if (!string.IsNullOrEmpty(txtDocNo2.Text))
            {
                sb.Append(" AND T0.[DocNum] <=@DocNum2 ");
                command.Parameters.Add(new SqlParameter("@DocNum2", DocNum2));
            }

            if (!string.IsNullOrEmpty(txtDocDate1.Text))
            {
                sb.Append(" AND T0.[PostDate] >=@DocDate1 ");
                command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            }

            if (!string.IsNullOrEmpty(txtDocDate2.Text))
            {
                sb.Append(" AND T0.[PostDate] <=@DocDate2 ");
                command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            }


            if (!string.IsNullOrEmpty(texCLOSE1.Text))
            {
                sb.Append(" AND T0.CLOSEDATE >=@CLOSE1 ");
                command.Parameters.Add(new SqlParameter("@CLOSE1", CLOSE1));
            }

            if (!string.IsNullOrEmpty(texCLOSE2.Text))
            {
                sb.Append(" AND T0.CLOSEDATE  <=@CLOSE2 ");
                command.Parameters.Add(new SqlParameter("@CLOSE2", CLOSE2));
            }
            //未結
            if (radioButton1.Checked)
            {
                sb.Append(" and StatuS in ('P','R') ");
            }

            //已結
            if (radioButton2.Checked)
            {
                sb.Append(" and StatuS in ('L') ");
            }


            //不含取消
            if (radioButton3.Checked)
            {
                sb.Append(" and StatuS in ('P','R','L') ");
            }

            if (comboBox1.Text == "TFT")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '1' ");
            }
            if (comboBox1.Text == "ESCO")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '4' ");
            }
            if (comboBox1.Text == "SOLAR")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '3' ");
            }

            command.CommandText = sb.ToString();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWOR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        public System.Data.DataTable GetOworDetail(string DocNum1, string DocNum2, string DocDate1, string DocDate2, string CLOSE1, string CLOSE2)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandType = CommandType.Text;

            sb.Append("SELECT T0.[OriginNum]  銷售單號,''''+T2.[CardCode] 客戶編號, T55.[CardName] 客戶名稱,");
            sb.Append(" T0.[u_projectcode] 專案代號, T56.PRJNAME  專案名稱,");
            sb.Append(" T0.[PostDate] 生產訂單日期, T0.[DueDate] 到期日期, T3.U_Name Owner,T0.[DocNum] 生產訂單,");
            sb.Append(" 狀態 =Case  when T0.[Status]='P' then '計劃'  when T0.[Status]='R' then '已核發' when T0.[Status]='L' then '結案' end, ");
            sb.Append(" T0.[ItemCode] 母件編號,''''+W1.ItemCode 子件編號, T1.[ItemName] 產品名稱,W1.[BaseQty] 基礎數量, W1.[PlannedQty] 計劃數量,W1.[IssuedQty] 發貨數量, ");
            sb.Append(" 已發貨成本 = (SELECT abs(Convert(int,Sum(T7.[TransValue])))   FROM  [dbo].[OINM] T7 WHERE T7.[ApplObj] = 202  AND  T7.[AppObjAbs] = T0.DocNum AND  T7.[AppObjLine] = W1.LineNum AND T7.[ItemCode] = W1.ItemCode AND  T7.[AppObjType] = 'C'  ), ");
            sb.Append(" T1.InvntryUom 計量單位,u_acme_work 排程日期,u_shipday 押出貨日,T1.U_GROUP 群組,W1.U_MEMO 備註");
            sb.Append(" FROM OWOR T0 ");
            sb.Append(" INNER JOIN WOR1 W1 ON W1.DocEntry=T0.DocNum");
            sb.Append(" Left JOIN OITM T1 ON T1.ItemCode= W1.ItemCode");
            sb.Append(" Left join ORDR T2 on T2.DocEntry=T0.[OriginNum]");
            sb.Append(" Left join OUSR T3 on T3.Userid=T0.[UserSign]");
            sb.Append("                Left join OCRD T55 on T2.CARDCODE=T55.CARDCODE");
            sb.Append("           Left join OPRJ T56 on T0.[u_projectcode]=T56.PRJCODE");
            sb.Append(" Where 1=1 ");
            if (globals.GroupID.ToString().Trim() == "ACCS" || globals.GroupID.ToString().Trim() == "SOLAR")
            {
                sb.Append("          AND T1.ITMSGRPCOD=102 ");
            }
            if (!string.IsNullOrEmpty(txtPrj1.Text))
            {
                sb.Append(" AND T0.[u_projectcode] >=@u_projectcode1 ");
                command.Parameters.Add(new SqlParameter("@u_projectcode1", txtPrj1.Text));
            }


            if (!string.IsNullOrEmpty(txtPrj2.Text))
            {
                sb.Append(" AND T0.[u_projectcode] <=@u_projectcode2 ");
                command.Parameters.Add(new SqlParameter("@u_projectcode2", txtPrj2.Text));
            }


            if (!string.IsNullOrEmpty(txtDocNo1.Text))
            {
                sb.Append(" AND T0.[DocNum] >=@DocNum1 ");
                command.Parameters.Add(new SqlParameter("@DocNum1", DocNum1));
            }

            if (!string.IsNullOrEmpty(txtDocNo2.Text))
            {
                sb.Append(" AND T0.[DocNum] <=@DocNum2 ");
                command.Parameters.Add(new SqlParameter("@DocNum2", DocNum2));
            }

            if (!string.IsNullOrEmpty(txtDocDate1.Text))
            {
                sb.Append(" AND T0.[PostDate] >=@DocDate1 ");
                command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            }

            if (!string.IsNullOrEmpty(txtDocDate2.Text))
            {
                sb.Append(" AND T0.[PostDate] <=@DocDate2 ");
                command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            }

            if (!string.IsNullOrEmpty(texCLOSE1.Text))
            {
                sb.Append(" AND T0.CLOSEDATE >=@CLOSE1 ");
                command.Parameters.Add(new SqlParameter("@CLOSE1", CLOSE1));
            }

            if (!string.IsNullOrEmpty(texCLOSE2.Text))
            {
                sb.Append(" AND T0.CLOSEDATE  <=@CLOSE2 ");
                command.Parameters.Add(new SqlParameter("@CLOSE2", CLOSE2));
            }

            //未結
            if (radioButton1.Checked)
            {
                sb.Append(" and StatuS in ('P','R') ");
            }

            //已結
            if (radioButton2.Checked)
            {
                sb.Append(" and StatuS in ('L') ");
            }


            //不含取消
            if (radioButton3.Checked)
            {
                sb.Append(" and StatuS in ('P','R','L') ");
            }

            if (comboBox1.Text == "TFT")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '1' ");
            }
            if (comboBox1.Text == "ESCO")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '4' ");
            }
            if (comboBox1.Text == "SOLAR")
            {
                sb.Append(" AND substring(T0.u_projectcode,1,1) = '3' ");
            }
            sb.Append(" ORDER BY  T0.[DocNum],T1.U_GROUP,W1.ItemCode ");
            command.CommandText = sb.ToString();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OWOR");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
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


        private void SetDefaultStyle_0(DataGridViewColumn Column)
        {
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle.Format = "#,##0";
            dataGridViewCellStyle.NullValue = null;
            Column.DefaultCellStyle = dataGridViewCellStyle;
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



        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gvData);
          
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gvDataDetail);
    
        }

        private void fmRep01_Load(object sender, EventArgs e)
        {
     
                UtilSimple.SetLookupBinding(comboBox1, GetMenu.SolarBU(), "DataText", "DataText");
            
        }

       
    
    }
}
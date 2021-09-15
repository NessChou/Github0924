using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Diagnostics;


namespace ACME
{
    public partial class SOLARPIV : Form
    {
        public SOLARPIV()
        {
            InitializeComponent();
        }

        public static System.Data.DataTable GetOrder()
        {
            SqlConnection connection = globals.shipConnection;
            //string sql = "SELECT '' as 項目,substring(T0.[CardName],1,2) 經銷商,Convert(varchar,datepart(mm,T0.[DocDate]))+'月' 月,datepart(dd,T0.[DocDate]) 日 , T0.[U_Beneficiary] 業主,";
            //sql += "substring(T1.ItemCode,4,3) KW,'AC' AC, substring(T1.ItemCode,2,8) 型號, substring(T1.ItemCode,4,3)+'W' W, T1.[Quantity],'' 備註, Convert(Varchar,T1.[ShipDate],111) FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry ";
            //離倉日期
            string sql = "SELECT '' as 項目,substring(T0.[CardName],1,2) 經銷商,Convert(varchar,datepart(mm,T1.[ShipDate]))+'月' 月,datepart(dd,T1.[ShipDate]) 日 , T0.[U_Beneficiary] 業主,";
            sql += "substring(T1.ItemCode,4,3) KW,'AC' AC, substring(T1.ItemCode,2,8) 型號, substring(T1.ItemCode,4,3)+'W' W, T1.[Quantity],T0.Comments 備註, Convert(Varchar,T0.[DocDate],111) as 下單日,T0.[U_ACME_ByAir] 出貨狀況,T0.[U_ACME_PAY] 付款狀況 FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry ";
            //Convert(Varchar,T1.[ShipDate],111) //Dscription 型號
            sql += " WHERE (Substring(T0.CardCode,1,4)='8801' or Substring(T0.CardCode,1,4)='8700' ) and len(T1.ItemCode) =15 and (substring(T1.ItemCode,2,2)='PM' or substring(T1.ItemCode,1,3)='311') and Convert(Varchar(8),T0.DocDate,112) >= @DocDate ";
            sql += " order by T0.[DocDate]";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            //string CardCode = "8801-00";
            // string CardCode = "8700";
            string DocDate = "20130101";
            //   command.Parameters.Add(new SqlParameter("@CardCode", CardCode));
            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));


            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TABLE");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TABLE"];
        }

        private void button1_Click(object sender, EventArgs e)
        {


            if (IsRunning())
            {
                return;
            }




            System.Data.DataTable dt = GetOrder();
            dataGridView1.DataSource = dt;

            //return

            string FileName = GetExePath() + "\\EXCEL\\SOLAR\\" + "PV_Data.xls";
            //依順序
            //  OutputCrossTable(dt, FileName);


            //輸出檔
            string OutPutFile = GetExePath() + "\\EXCEL\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report

            //Pivot 固定一個
            ExcelReport.ExcelReportOutput(dt, FileName, OutPutFile, "pivot");
        }
        private bool IsRunning()
        {

            bool b = false;

            Process[] pProcess;
            pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

            if (pProcess.Length > 0)
            {
                MessageBox.Show("請先將 Excel 關閉");
                b = true;
            }

            return b;
        }

        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }

    }
}

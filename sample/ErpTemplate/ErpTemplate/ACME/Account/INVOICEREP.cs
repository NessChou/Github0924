using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class INVOICEREP : Form
    {
        public INVOICEREP()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTableMONTH()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("TYPE", typeof(string));
            dt.Columns.Add("1月", typeof(int));
            dt.Columns.Add("2月", typeof(int));
            dt.Columns.Add("3月", typeof(int));
            dt.Columns.Add("4月", typeof(int));
            dt.Columns.Add("5月", typeof(int));
            dt.Columns.Add("6月", typeof(int));
            dt.Columns.Add("7月", typeof(int));
            dt.Columns.Add("8月", typeof(int));
            dt.Columns.Add("9月", typeof(int));
            dt.Columns.Add("10月", typeof(int));
            dt.Columns.Add("11月", typeof(int));
            dt.Columns.Add("12月", typeof(int));
            dt.Columns.Add("合計", typeof(int));
            
            return dt;

        }

        System.Data.DataTable GetTemp(string YEAR, int MONTH, string TYPE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT　CASE U_PC_BSTYC WHEN 1 THEN '作廢' ELSE TAX END 課稅類別,COUNT(CASE U_PC_BSTYC WHEN 1 THEN '作廢' ELSE TAX END) 張數   FROM (    ");
            sb.Append(" SELECT  T0.[U_PC_BSINV], T0.[U_PC_BSTYC],  ");
            sb.Append(" CASE  T0.[U_PC_BSTY2] WHEN 0 THEN '應稅' WHEN 1 THEN '零稅率' ");
            sb.Append(" WHEN 2 THEN '免稅' WHEN 3 THEN '不計稅' END TAX,T0.U_PC_BSAPP DDATE ");
            sb.Append(" FROM OPCH T0  ");
            sb.Append(" WHERE   YEAR(T0.U_PC_BSAPP)=@YEAR  ");
            sb.Append(" AND MONTH(T0.U_PC_BSAPP)=@MONTH  ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT  T0.[U_BSINV], T2.[U_RP_BSTYC], ");
            sb.Append(" CASE  T0.[U_BSTY2] WHEN 0 THEN '應稅' WHEN 1 THEN '零稅率' ");
            sb.Append(" WHEN 2 THEN '免稅' WHEN 3 THEN '不計稅' END,T2.U_RP_BSAPP DDATE ");
            sb.Append(" FROM [@CADMEN_PMD1] T0  ");
            sb.Append(" left join [@CADMEN_PMD]  T1 on T0.DOCENTRY=T1.DOCENTRY  ");
            sb.Append(" left join [ORPC]  T2 on T1.U_BSREN=T2.DOCENTRY  ");
            sb.Append(" WHERE   YEAR(T2.U_RP_BSAPP)=@YEAR  ");
            sb.Append(" AND MONTH(T2.U_RP_BSAPP)=@MONTH  ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT  T0.[U_PC_BSINV],  T0.[U_PC_BSTYC],  ");
            sb.Append(" CASE  T0.[U_PC_BSTY2] WHEN 0 THEN '應稅' WHEN 1 THEN '零稅率' ");
            sb.Append(" WHEN 2 THEN '免稅' WHEN 3 THEN '不計稅' END ,T0.U_PC_BSAPP DDATE ");
            sb.Append(" FROM [@CADMEN_FMD1] T0  ");
            sb.Append(" left join [@CADMEN_FMD]  T1 on T0.DOCENTRY=T1.DOCENTRY  ");
            sb.Append(" WHERE  YEAR(T0.[U_PC_BSAPP])=@YEAR   ");
            sb.Append(" AND MONTH(T0.[U_PC_BSAPP])=@MONTH  ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT U_IN_BSINV U_PC_BSINV,U_IN_BSTYC U_PC_BSTYC, ");
            sb.Append(" CASE  T0.U_IN_BSTY2 WHEN 0 THEN '應稅' WHEN 1 THEN '零稅率' ");
            sb.Append(" WHEN 2 THEN '免稅' WHEN 3 THEN '不計稅' END,T0.U_IN_BSAPP DDATE  FROM OINV T0   ");
            sb.Append(" WHERE YEAR(T0.U_IN_BSAPP)=@YEAR   ");
            sb.Append(" AND MONTH(T0.U_IN_BSAPP)=@MONTH   ");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT  T0.[U_BSINV],T2.[U_RI_BSTYC],  ");
            sb.Append(" CASE  T0.[U_BSTY2] WHEN 0 THEN '應稅' WHEN 1 THEN '零稅率' ");
            sb.Append(" WHEN 2 THEN '免稅' WHEN 3 THEN '不計稅' END,T2.U_RI_BSAPP DDATE   ");
            sb.Append(" FROM [@CADMEN_CMD1] T0  ");
            sb.Append(" left join [@CADMEN_CMD]  T1 on T0.DOCENTRY=T1.DOCENTRY  ");
            sb.Append(" left join [ORIN]  T2 on T1.U_BSREN=T2.DOCENTRY  ");
            sb.Append(" WHERE YEAR(T2.U_RI_BSAPP)=@YEAR  ");
            sb.Append(" AND MONTH(T2.U_RI_BSAPP)=@MONTH   ");
            sb.Append(" ) aS A  WHERE U_PC_BSINV <> '__________' AND  CASE U_PC_BSTYC WHEN 1 THEN '作廢' ELSE TAX END =@TYPE  ");
            sb.Append(" GROUP BY CASE U_PC_BSTYC WHEN 1 THEN '作廢' ELSE TAX END ");
    
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@TYPE", TYPE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
    
    
        private void button1_Click(object sender, EventArgs e)
        {

            EXEC();


    

        }
        private void EXEC()
        {
            System.Data.DataTable dtemp5 = GetTemp61();
            System.Data.DataTable dtCostDD = MakeTableMONTH();
            DataRow drtemp5 = null;
            string YEAR = comboBox1.Text;


            for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
            {

                string TYPE = dtemp5.Rows[i]["PARAM_NO"].ToString();

         
             
                    drtemp5 = dtCostDD.NewRow();

                    drtemp5["TYPE"] = TYPE;
                    int G = 0;
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = null;
                        dh = GetTemp(YEAR, y, TYPE);

                        if (dh.Rows.Count > 0)
                        {
                            drtemp5[y + "月"] = dh.Rows[0]["張數"];
                            G += Convert.ToInt16(dh.Rows[0]["張數"]);
                        }
                        else
                        {
                            drtemp5[y + "月"] = 0;
                        }

                    }
                    drtemp5["合計"] = G.ToString();
                    dtCostDD.Rows.Add(drtemp5);
                
            
            }
            dataGridView1.DataSource = dtCostDD;

            //加入一筆合計
            decimal[] Total = new decimal[dtCostDD.Columns.Count - 1];

            for (int i = 0; i <= dtCostDD.Rows.Count - 1; i++)
            {

                for (int j = 1; j <= dtCostDD.Columns.Count - 1; j++)
                {
                    Total[j - 1] += Convert.ToInt16(dtCostDD.Rows[i][j]);

                }
            }

            DataRow row;

            row = dtCostDD.NewRow();

            row[0] = "合計";
            for (int j = 1; j <= dtCostDD.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dtCostDD.Rows.Add(row);

            for (int i = 1; i <= dataGridView1.Columns.Count - 1; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void LCREPORT_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
   
        }
        System.Data.DataTable GetTemp61()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                       SELECT PARAM_NO FROM PARAMS WHERE PARAM_KIND='INVOICEREP'  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1); 
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //if (e.RowIndex >= dataGridView1.Rows.Count)
            //    return;
            //DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["BANK"].Value.ToString() == "小計")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Yellow;
            //        dgr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //    }
            //    else if (dgr.Cells["BANK"].Value.ToString() == "加總")
            //    {

            //        dgr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //        dgr.DefaultCellStyle.BackColor = Color.YellowGreen;
            //    }
          
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //} 
        }
    }

}

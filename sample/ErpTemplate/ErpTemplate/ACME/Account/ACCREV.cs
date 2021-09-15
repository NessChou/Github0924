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
    public partial class ACCREV : Form
    {
        System.Data.DataTable dtCost = null;
        public ACCREV()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
    
            System.Data.DataTable V1 = GetItemmDD("Account_Temp61"+comboBox1.Text);
            string YEAR = comboBox1.Text;
            string MONTH = V1.Rows[0][0].ToString();
            int M = Convert.ToInt16(MONTH);
            DataRow dr = null;
            dtCost = MakeTable(YEAR, M);
     
     
            dr = dtCost.NewRow();
            dr["月份/資料"] = "月份/資料";

            for (int i = 1; i <= M; i++)
            {
                for (int i2 = 1; i2 <= 5; i2++)
                {
                    dr[YEAR + "-" + i + "-" + i2] = YEAR + "-" + i;
                }
            }

            for (int i2 = 1; i2 <= 5; i2++)
            {
                dr[YEAR + "年度合計-" + i2] = YEAR + "年度合計";
            }
            dtCost.Rows.Add(dr);

            dr = dtCost.NewRow();
            dr["月份/資料"] = "產品編號";

            for (int i = 1; i <= M; i++)
            {
                for (int i2 = 1; i2 <= 5; i2++)
                {
                    if (i2 == 1)
                    {
                        dr[YEAR + "-" + i + "-" + i2] = "數量";
                    }
                    if (i2 == 2)
                    {
                        dr[YEAR + "-" + i + "-" + i2] = "總收入";
                    }
                    if (i2 == 3)
                    {
                        dr[YEAR + "-" + i + "-" + i2] = "總成本";
                    }
                    if (i2 == 4)
                    {
                        dr[YEAR + "-" + i + "-" + i2] = "總毛利";
                    }
                    if (i2 == 5)
                    {
                        dr[YEAR + "-" + i + "-" + i2] = "毛利率";
                    }
                }
            }

            for (int i2 = 1; i2 <= 5; i2++)
            {
                if (i2 == 1)
                {
                    dr[YEAR + "年度合計-" + i2] = "數量";
                }
                if (i2 == 2)
                {
                    dr[YEAR + "年度合計-" + i2] = "總收入";
                }
                if (i2 == 3)
                {
                    dr[YEAR + "年度合計-" + i2] = "總成本";
                }
                if (i2 == 4)
                {
                    dr[YEAR + "年度合計-" + i2] = "總毛利";
                }
                if (i2 == 5)
                {
                    dr[YEAR + "年度合計-" + i2] = "毛利率";
                }

       
            }
            dtCost.Rows.Add(dr);

            System.Data.DataTable dt = GetREV1("Account_Temp61" + comboBox1.Text, 0, "", "", "");
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                string ITEMCODE = dt.Rows[j]["ITEMCODE"].ToString();
                //if (ITEMCODE == "G070VVN01.02002")
                //{

                //    MessageBox.Show("A");

                //}
                dr = dtCost.NewRow();
                dr["月份/資料"] = ITEMCODE;
                for (int i = 1; i <= M; i++)
                {
                    for (int i2 = 1; i2 <= 5; i2++)
                    {

                        System.Data.DataTable dt2 = GetREV1("Account_Temp61" + comboBox1.Text, i, ITEMCODE, "M", "T");
                        if (dt2.Rows.Count > 0)
                        {
                            string GQTY = dt2.Rows[0][1].ToString();
                            string GTOTAL = dt2.Rows[0][2].ToString();
                            string GVALUE = dt2.Rows[0][3].ToString();
                            string GREV = dt2.Rows[0][4].ToString();
                            string REV = dt2.Rows[0][5].ToString();
                            if (i2 == 1)
                            {
                                dr[YEAR + "-" + i + "-" + i2] = GQTY;
                            }
                            if (i2 == 2)
                            {
                                dr[YEAR + "-" + i + "-" + i2] = GTOTAL;
                            }
                            if (i2 == 3)
                            {
                                dr[YEAR + "-" + i + "-" + i2] = GVALUE;
                            }
                            if (i2 == 4)
                            {
                                dr[YEAR + "-" + i + "-" + i2] = GREV;
                            }
                            if (i2 == 5)
                            {
                                dr[YEAR + "-" + i + "-" + i2] = REV;
                            }
                        }
                    }
                }

                for (int i2 = 1; i2 <= 5; i2++)
                {

                    System.Data.DataTable dt2 = GetREV1("Account_Temp61" + comboBox1.Text, 0, ITEMCODE, "", "T");
                    string GQTY = dt2.Rows[0][1].ToString();
                    string GTOTAL = dt2.Rows[0][2].ToString();
                    string GVALUE = dt2.Rows[0][3].ToString();
                    string GREV = dt2.Rows[0][4].ToString();
                    string REV = dt2.Rows[0][5].ToString();

                    if (i2 == 1)
                    {
                        dr[YEAR + "年度合計-" + i2] = GQTY;
                    }
                    if (i2 == 2)
                    {
                        dr[YEAR + "年度合計-" + i2] = GTOTAL;
                    }
                    if (i2 == 3)
                    {
                        dr[YEAR + "年度合計-" + i2] = GVALUE;
                    }
                    if (i2 == 4)
                    {
                        dr[YEAR + "年度合計-" + i2] = GREV;
                    }
                    if (i2 == 5)
                    {
                        dr[YEAR + "年度合計-" + i2] = REV;
                    }


                }

                dtCost.Rows.Add(dr);
            }

            dataGridView1.DataSource = dtCost;
        }

        private System.Data.DataTable MakeTable(string YEAR,int M)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("月份/資料");
            for (int i = 1; i <= M; i++)
            {
                for (int i2 = 1; i2 <= 5; i2++)
                {
                    dt.Columns.Add(YEAR + "-" + i + "-" + i2, typeof(string));
                }
            }

            for (int i2 = 1; i2 <= 5; i2++)
            {
                dt.Columns.Add(YEAR + "年度合計-" + i2, typeof(string));
            }

            return dt;
        }

        System.Data.DataTable GetItemmDD(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(MONTH(DDATE)) M from " + Account_Temp6 + " ");


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
        System.Data.DataTable GetREV1(string Account_Temp6, int  MONTH, string ITEMCODE, string DOCTYPE, string DOCTYPE2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,SUM(GQTY) QTY,SUM(GTotal) GTOTAL,SUM(GVALUE) GVALUE,(SUM(GTotal) - SUM(GVALUE)) GREV,CASE WHEN SUM(GTOTAL) <> 0 THEN CAST(CAST(((SUM(GTotal) - SUM(GVALUE))/SUM(GTotal))*100 AS decimal(16,2)) AS VARCHAR)+'%' END  REV from " + Account_Temp6 + "  WHERE ITEMCODE <>'0' ");
            if (DOCTYPE == "M")
            {
                sb.Append(" AND  MONTH(DDATE)=@MONTH");
            }
            if (DOCTYPE2 == "T")
            {
                sb.Append(" AND  ITEMCODE=@ITEMCODE");
            }
            sb.Append(" GROUP BY ITEMCODE ORDER BY ITEMCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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

        private void ACCREV_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year2017(), "DataValue", "DataValue");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcelNOTHEAD(dataGridView1);
        }

    }
}

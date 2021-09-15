using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class ACCRATE : Form
    {
        public ACCRATE()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DA();
        }

        System.Data.DataTable GeRATE()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  Convert(varchar(8),RateDate,112) 日期,RATE 央行匯率,HBUY 海關三旬買進匯率,HFAR 合庫匯率");
            sb.Append(" FROM ACMESQL02.DBO.ORTT T0");
            sb.Append(" LEFT JOIN WH_HAIGUAN T1 ON (YEAR(T0.RateDate)=T1.HYEAR");
            sb.Append(" AND MONTH(RateDate)=T1.HMON");
            sb.Append(" AND CASE WHEN DAY(RateDate) BETWEEN 1 AND 10 THEN '1-10' ");
            sb.Append(" WHEN DAY(RateDate) BETWEEN 11 AND 20 THEN '11-20' ");
            sb.Append(" WHEN DAY(RateDate) BETWEEN 21 AND 31 THEN '21-31' ");
            sb.Append(" END=T1.HDAY");
            sb.Append(" )");
            sb.Append("  WHERE CURRENCY='USD' AND YEAR(RateDate) BETWEEN @DAY1 AND @DAY2");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DAY1", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DAY2", comboBox2.Text));
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
        System.Data.DataTable GeRATE2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT HDAY 日期,HPRICE 匯率,DC WD FROM WH_HAIGUAN2 WHERE SUBSTRING(HDAY,1,4) BETWEEN @DAY1 AND @DAY2 ORDER BY HDAY ");
     

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DAY1", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DAY2", comboBox2.Text));
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
        System.Data.DataTable GeRATE3()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT HDAY 日期,BPRICE  '即期匯率-美金買入' ,SPRICE  '即期匯率-美金賣出',CAST(ROUND((CAST(BPRICE AS DECIMAL(10,3))+CAST(SPRICE AS DECIMAL(10,3)))/2,3) AS decimal(10,3)) 即期美金平均匯率 FROM WH_HAIGUAN3 T0");
            sb.Append(" INNER JOIN (  SELECT   MAX(  Convert(varchar(10),DATE_TIME,112))  TDATE FROM    acmesqlsp.dbo.Y_2004  WHERE  (IsRestDay   =   0 OR WD = 'Y') ");
            sb.Append("  GROUP BY YEAR(DATE_TIME),MONTH(DATE_TIME) ) T1 ON (T0.HDAY =T1.TDATE)");
            sb.Append("  WHERE SUBSTRING(HDAY,1,4) BETWEEN @DAY1 AND @DAY2");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DAY1", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DAY2", comboBox2.Text));
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

        private void ACCRATE_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year2017(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Year2017(), "DataValue", "DataValue");

            DA();
        }

        private void DA()
        {
            dataGridView1.DataSource = GeRATE();
            dataGridView2.DataSource = GeRATE2();
            dataGridView3.DataSource = GeRATE3();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }

        }

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= dataGridView2.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView2.Rows[e.RowIndex];
            try
            {
                if (!String.IsNullOrEmpty(dgr.Cells["WD"].Value.ToString()))
                {



                    dgr.DefaultCellStyle.BackColor = Color.Pink;

                }

            }
            catch (Exception ex)
            {
             
            }
        }
    }
}

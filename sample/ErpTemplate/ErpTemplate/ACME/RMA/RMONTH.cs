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
    public partial class RMONTH : Form
    {
        public RMONTH()
        {
            InitializeComponent();
        }
        private System.Data.DataTable MakeTableMONTH()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("ZPN", typeof(string));
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

        System.Data.DataTable GetTemp(string MONTH, string MODEL, string VER, string ZPN)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  ISNULL(COUNT(*),0) QTY");
            sb.Append(" from octr t0");
            sb.Append(" left join ctr1 t1 on (t0.contractid=t1.contractid)");
            sb.Append(" where  t0.[U_RMA_NO] <> '' and  t0.u_pkind <>'6'");
            sb.Append(" and  '20'+T0.U_RMAYEAR=@YEAR AND SUBSTRING(U_S_SEQ,13,1) =@ZPN ");
            sb.Append(" AND [U_rmodel] = @MODEL ");
            sb.Append("  AND U_RVer=@VER and (substring(t0.U_RMA_NO,4,2))=@MONTH ");
            sb.Append(" AND ISNULL([U_rmodel],'') <> ''");
            sb.Append(" GROUP BY [U_rmodel] ,U_RVer ,SUBSTRING(U_S_SEQ,13,1) ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@ZPN", ZPN));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));


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

                string MODEL = dtemp5.Rows[i]["MODEL"].ToString();
                string VER = dtemp5.Rows[i]["VER"].ToString();
                string ZPN = dtemp5.Rows[i]["ZPN"].ToString();

                    drtemp5 = dtCostDD.NewRow();

                    drtemp5["MODEL"] = MODEL;
                    drtemp5["VER"] = VER;
                    drtemp5["ZPN"] = ZPN;
                    int G = 0;
                    for (int y = 1; y <= 12; y++)
                    {
                        string MONTH = y.ToString();
                        if (y < 10)
                        {
                            MONTH = "0" + MONTH;
                        }

                        System.Data.DataTable dh = null;
                        dh = GetTemp(MONTH, MODEL, VER, ZPN);

                        if (dh.Rows.Count > 0)
                        {
                            drtemp5[y + "月"] = dh.Rows[0]["QTY"];
                            G += Convert.ToInt16(dh.Rows[0]["QTY"]);
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

                for (int j =3; j <= dtCostDD.Columns.Count - 1; j++)
                {
                    Total[j - 1] += Convert.ToInt16(dtCostDD.Rows[i][j]);

                }
            }

            DataRow row;

            row = dtCostDD.NewRow();

            row[0] = "合計";
            for (int j = 3; j <= dtCostDD.Columns.Count - 1; j++)
            {
                row[j] = Total[j - 1];

            }
            dtCostDD.Rows.Add(row);

            for (int i = 3; i <= dataGridView1.Columns.Count - 1; i++)
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
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT [U_rmodel] MODEL,U_RVer VER,SUBSTRING(U_S_SEQ,13,1) ZPN");
            sb.Append(" from octr t0");
            sb.Append(" left join ctr1 t1 on (t0.contractid=t1.contractid)");
            sb.Append(" where  t0.[U_RMA_NO] <> '' and  t0.u_pkind <>'6'");
            sb.Append(" and  '20'+T0.U_RMAYEAR=@YEAR AND SUBSTRING(U_S_SEQ,13,1) IN ('Z','P','N') ");
            if (checkBox1.Checked)
            {
                sb.Append(" AND [U_rmodel]  LIKE '%OPEN%'");
            }
            else
            {
                sb.Append(" AND [U_rmodel] NOT LIKE '%OPEN%'");
            }
            sb.Append(" AND ISNULL([U_rmodel],'') <> ''");
            sb.Append(" ORDER BY [U_rmodel],U_RVer,SUBSTRING(U_S_SEQ,13,1) ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));

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

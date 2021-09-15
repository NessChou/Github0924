using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
namespace ACME
{
    public partial class PLATE : Form
    {
        public PLATE()
        {
            InitializeComponent();
        }

        private void wH_PLATEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.wH_PLATEBindingSource.EndEdit();
            this.wH_PLATETableAdapter.Update(this.wh.WH_PLATE);

        }



        private void PLATE_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Month(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetMenu.GetBUPLATE(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, GetMenu.GetWHPLATE(comboBox3.Text), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox5, GetMenu.GetWHPLATE(comboBox6.Text), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox6, GetMenu.GetBUPLATE(), "DataValue", "DataValue");


            comboBox1.Text = DateTime.Now.ToString("yyyy");
            comboBox2.Text = Convert.ToString(Convert.ToInt16(DateTime.Now.ToString("MM")));

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            
        }

        private void wH_PLATEDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
       
     

            e.Row.Cells["DOCYEAR"].Value = comboBox1.Text;
            e.Row.Cells["DOCMONTH"].Value = comboBox2.Text;
            e.Row.Cells["BU"].Value = comboBox3.Text;
            e.Row.Cells["WHSCODE"].Value = comboBox4.Text;
            
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Data.DataTable J1 = oclg3();

            if (J1.Rows.Count == 0)
            {
                System.Data.DataTable J2 = oclg2();
                int K1 = Convert.ToInt16(J2.Rows[0][0].ToString());

                for (int h = 1; h <= K1; h++)
                {


                    AddPLATE((h).ToString());
                }


            
            }
           
            
            try
            {
                this.wH_PLATETableAdapter.Fill(this.wh.WH_PLATE, new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox1.Text, typeof(int))))), new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox2.Text, typeof(int))))), comboBox4.Text, comboBox3.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        public void AddPLATE(string DOCDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO [AcmeSqlSP].[dbo].[WH_PLATE]");
            sb.Append("            (DOCYEAR,DOCMONTH,DOCDATE,WHSCODE,BU)");
            sb.Append("      VALUES(@DOCYEAR,@DOCMONTH,@DOCDATE,@WHSCODE,@BU)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCYEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCMONTH", comboBox2.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@WHSCODE", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@BU", comboBox3.Text));
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
        private void button2_Click(object sender, EventArgs e)
        {

            this.Validate();
            this.wH_PLATEBindingSource.EndEdit();
            this.wH_PLATETableAdapter.Update(this.wh.WH_PLATE);

            try
            {
                DataGridViewRow row;
                for (int h = 0; h <= wH_PLATEDataGridView.Rows.Count - 1; h++)
                {

                    row = wH_PLATEDataGridView.Rows[h];



                    string DOCYEAR = row.Cells["DOCYEAR"].Value.ToString();
                    string DOCMONTH = row.Cells["DOCMONTH"].Value.ToString();
                    string DOCDATE = row.Cells["DOCDATE"].Value.ToString();
                    string WHSCODE = row.Cells["WHSCODE"].Value.ToString();
                    string BU = row.Cells["BU"].Value.ToString();
                    System.Data.DataTable T1 = GetItem3(WHSCODE, BU, DOCYEAR, DOCMONTH, DOCDATE);
                    DataRow dow = T1.Rows[0];
                    string TOTALQTY = dow["TOTALQTY"].ToString();
                    string T2 = dow["QTY"].ToString();
                    string T3 = dow["ID"].ToString();
                    string DOCDATE1 = dow["DOCDATE1"].ToString();
                    if (DOCDATE1 != "0")
                    {
                        Updatecho(T2, T3);
                    }
                }
            }
            catch { }

            MessageBox.Show("存檔成功");

            try
            {
                this.wH_PLATETableAdapter.Fill(this.wh.WH_PLATE, new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox1.Text, typeof(int))))), new System.Nullable<int>(((int)(System.Convert.ChangeType(comboBox2.Text, typeof(int))))), comboBox4.Text, comboBox3.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

  


        private System.Data.DataTable GetItem3(string WHSCODE, string BU, string DOCYEAR, string DOCMONTH, string DOCDATE)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ID, T0.DOCYEAR, T0.DOCMONTH, T0.DOCDATE, T0.WHSCODE, T0.BU, T0.INQTY, T0.OUTQTY,isnull(T1.DOCDATE,'') DOCDATE1,ISNULL(T1.TOTALQTY,0) TOTALQTY,ISNULL(T1.TOTALQTY,0)+T0.INQTY-T0.OUTQTY QTY");
            sb.Append(" FROM dbo.WH_PLATE  T0");
            sb.Append(" LEFT JOIN WH_PLATE T1 ON (DATEADD(day,-1,cast(cast(T0.DOCYEAR as varchar)+(CASE WHEN T0.DOCMONTH < 10 THEN '0'+CAST(T0.DOCMONTH AS VARCHAR) ELSE CAST(T0.DOCMONTH AS VARCHAR) END)+(CASE WHEN T0.DOCDATE < 10 THEN '0'+CAST(T0.DOCDATE AS VARCHAR) ELSE CAST(T0.DOCDATE AS VARCHAR) END) as datetime)) =cast(cast(T1.DOCYEAR as varchar)+(CASE WHEN T1.DOCMONTH < 10 THEN '0'+CAST(T1.DOCMONTH AS VARCHAR) ELSE CAST(T1.DOCMONTH AS VARCHAR) END)+(CASE WHEN T1.DOCDATE < 10 THEN '0'+CAST(T1.DOCDATE AS VARCHAR) ELSE CAST(T1.DOCDATE AS VARCHAR) END) as datetime) AND T0.WHSCODE=T1.WHSCODE AND T0.BU=T1.BU )");
            sb.Append(" where T0.WHSCODE=@WHSCODE AND T0.BU=@BU AND T0.DOCYEAR=@DOCYEAR AND T0.DOCMONTH=@DOCMONTH AND T0.DOCDATE=@DOCDATE ORDER BY T0.DOCDATE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));
            command.Parameters.Add(new SqlParameter("@BU", BU));
            command.Parameters.Add(new SqlParameter("@DOCYEAR", DOCYEAR));
            command.Parameters.Add(new SqlParameter("@DOCMONTH", DOCMONTH));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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

        private System.Data.DataTable GetItem4(string ID)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.ID, T0.DOCYEAR, T0.DOCMONTH, T0.DOCDATE, T0.WHSCODE, T0.BU, T0.INQTY, T0.OUTQTY,T1.DOCDATE,ISNULL(T1.TOTALQTY,0) TOTALQTY,ISNULL(T1.TOTALQTY,0)+T0.INQTY-T0.OUTQTY QTY");
            sb.Append(" FROM dbo.WH_PLATE  T0");
            sb.Append(" LEFT JOIN WH_PLATE T1 ON (DATEADD(day,-1,cast(cast(T0.DOCYEAR as varchar)+(CASE WHEN T0.DOCMONTH < 10 THEN '0'+CAST(T0.DOCMONTH AS VARCHAR) ELSE CAST(T0.DOCMONTH AS VARCHAR) END)+(CASE WHEN T0.DOCDATE < 10 THEN '0'+CAST(T0.DOCDATE AS VARCHAR) ELSE CAST(T0.DOCDATE AS VARCHAR) END) as datetime)) =cast(cast(T1.DOCYEAR as varchar)+(CASE WHEN T1.DOCMONTH < 10 THEN '0'+CAST(T1.DOCMONTH AS VARCHAR) ELSE CAST(T1.DOCMONTH AS VARCHAR) END)+(CASE WHEN T1.DOCDATE < 10 THEN '0'+CAST(T1.DOCDATE AS VARCHAR) ELSE CAST(T1.DOCDATE AS VARCHAR) END) as datetime) AND T0.WHSCODE=T1.WHSCODE AND T0.BU=T1.BU )");
            sb.Append(" where T0.ID=@ID ORDER BY T0.DOCDATE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));


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
        private void Updatecho(string TOTALQTY, string ID)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE WH_PLATE SET TOTALQTY=@TOTALQTY WHERE ID=@ID   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@TOTALQTY", TOTALQTY));
            command.Parameters.Add(new SqlParameter("@ID", ID));
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

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        private System.Data.DataTable oclg1()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append(" select cast(DOCYEAR as varchar)+(CASE WHEN DOCMONTH < 10 THEN '0'+CAST(DOCMONTH AS VARCHAR) ELSE CAST(DOCMONTH AS VARCHAR) END)+(CASE WHEN DOCDATE < 10 THEN '0'+CAST(DOCDATE AS VARCHAR) ELSE CAST(DOCDATE AS VARCHAR) END) 日期");
            sb.Append(" ,WHSCODE 倉庫,INQTY 進貨,OUTQTY 出貨,TOTALQTY 當天板數");
            sb.Append("   from dbo.WH_PLATE WHERE  BU=@BU AND cast(DOCYEAR as varchar)+(CASE WHEN DOCMONTH < 10 THEN '0'+CAST(DOCMONTH AS VARCHAR) ELSE CAST(DOCMONTH AS VARCHAR) END)+(CASE WHEN DOCDATE < 10 THEN '0'+CAST(DOCDATE AS VARCHAR) ELSE CAST(DOCDATE AS VARCHAR) END)  BETWEEN @AA AND @BB AND WHSCODE=@WHSCODE  ");
            sb.Append("   AND  ISNULL(INQTY,'')+ISNULL(OUTQTY,'')+ISNULL(TOTALQTY,'') <>'' ");
            sb.Append(" ORDER BY  cast(DOCYEAR as varchar)+(CASE WHEN DOCMONTH < 10 THEN '0'+CAST(DOCMONTH AS VARCHAR) ELSE CAST(DOCMONTH AS VARCHAR) END)+(CASE WHEN DOCDATE < 10 THEN '0'+CAST(DOCDATE AS VARCHAR) ELSE CAST(DOCDATE AS VARCHAR) END) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@BU", comboBox6.Text));
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@WHSCODE", comboBox5.Text));

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
        private System.Data.DataTable oclg2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append(" select count(*) 筆數 from Y_2004 where year(date_time)=@YEAR and month(date_time)=@MON ");
          
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@YEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@MON", comboBox2.Text));

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


        private System.Data.DataTable oclg3()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM WH_PLATE T0 where T0.WHSCODE=@WHSCODE AND T0.BU=@BU AND T0.DOCYEAR=@DOCYEAR AND T0.DOCMONTH=@DOCMONTH ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHSCODE", comboBox4.Text ));
            command.Parameters.Add(new SqlParameter("@BU", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@DOCYEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCMONTH", comboBox2.Text));
     

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
        private System.Data.DataTable oclg4()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);


            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCMONTH 月,DOCDATE 日,INQTY 進貨,OUTQTY 出貨,TOTALQTY 當天板數 FROM WH_PLATE T0 where T0.WHSCODE=@WHSCODE AND T0.BU=@BU AND T0.DOCYEAR=@DOCYEAR AND T0.DOCMONTH=@DOCMONTH ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@WHSCODE", comboBox4.Text));
            command.Parameters.Add(new SqlParameter("@BU", comboBox3.Text));
            command.Parameters.Add(new SqlParameter("@DOCYEAR", comboBox1.Text));
            command.Parameters.Add(new SqlParameter("@DOCMONTH", comboBox2.Text));


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
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox4, GetMenu.GetWHPLATE(comboBox3.Text), "DataValue", "DataValue");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\wh\\PLATE2.xls";


                System.Data.DataTable OrderData = oclg4();


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox5, GetMenu.GetWHPLATE(comboBox6.Text), "DataValue", "DataValue");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Data.DataTable OrderData = oclg1();

            dataGridView1.DataSource = OrderData;
        }

 
    }
}
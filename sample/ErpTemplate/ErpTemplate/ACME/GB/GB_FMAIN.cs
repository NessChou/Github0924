using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class GB_FMAIN : Form
    {

        string DAY = "";
        string SWEEK = "";
        string S1 = "";
        string E1 = "";
        string E2 = "";
        public GB_FMAIN()
        {
            InitializeComponent();
        }

        private System.Data.DataTable MakeTableWeek()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("WK-1", typeof(string));
            dt.Columns.Add("WK-2", typeof(string));
            dt.Columns.Add("WK-3", typeof(string));
            dt.Columns.Add("WK-4", typeof(string));
            dt.Columns.Add("WK-5", typeof(string));
            dt.Columns.Add("WK-6", typeof(string));
            dt.Columns.Add("WK-7", typeof(string));
            dt.Columns.Add("WK-8", typeof(string));
            dt.Columns.Add("WK-9", typeof(string));
            dt.Columns.Add("WK-10", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTableWeek2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("SEQ", typeof(string));
            dt.Columns.Add("WK", typeof(string));
  

            return dt;
        }
        private System.Data.DataTable MakeTableWeek3()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("RGN", typeof(string));
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("LINE", typeof(string));
            dt.Columns.Add("DOCDATE", typeof(string));
            dt.Columns.Add("DOCDATE1", typeof(string));
            dt.Columns.Add("ITEMCODE", typeof(string));
            dt.Columns.Add("ITEMNAME", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("NONE", typeof(string));
   
            return dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            G1("外站Roots", dataGridView1, "0");

        }
        private void G0(string FTYPE, DataGridView G,string FF)
        {
            System.Data.DataTable dtWeek = MakeTableWeek();
            System.Data.DataTable dt = GetT1();
            DataRow dr = null;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtWeek.NewRow();
                string ITEMCODE = dt.Rows[i]["ITEMCODE"].ToString();
                dr["料號"] = ITEMCODE;
                dr["品名規格"] = dt.Rows[i]["ITEMNAME"].ToString();
                for (int F = 1; F <= 10; F++)
                {
                    G4(ITEMCODE, F, FTYPE, FF);
                    dr["WK-" + F.ToString()] = DAY;
                }

                dtWeek.Rows.Add(dr);
            }

            G.DataSource = dtWeek;

            for (int F = 1; F <= 10; F++)
            {
                string D1 = "";

                if (FF == "1")
                {
                    string DATE = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
                    DateTime DD = Convert.ToDateTime(DATE);

                    if (F == 1)
                    {
                        D1 = DD.ToString("yyyyMMdd");
                    }
                    else
                    {
                        D1 = DD.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                    }
                }
                else
                {
                    if (F == 1)
                    {
                        D1 = DateTime.Now.ToString("yyyyMMdd");
                    }
                    else
                    {
                        D1 = DateTime.Now.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                    }

                }
           
                
                
                System.Data.DataTable T1 = GetT2(D1);
                if (T1.Rows.Count > 0)
                {
                    S1 = T1.Rows[0][0].ToString().Substring(4, 4);
                    E1 = T1.Rows[0][1].ToString().Substring(4, 4);
                    E2 = T1.Rows[0][2].ToString();
                    G.Columns[F + 1].HeaderText = S1 + "~" + E1 + " W" + E2.ToString();
                }
            }
        }

        private void GW()
        {
            System.Data.DataTable dtWeek = MakeTableWeek2();
            DataRow dr = null;
            for (int F = 1; F <= 10; F++)
            {

                string D1 = "";

                string DATE = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
                DateTime DD = Convert.ToDateTime(DATE);

                if (F == 1)
                {
                    D1 = DD.ToString("yyyyMMdd");
                }
                else
                {
                    D1 = DD.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                }

    

                System.Data.DataTable T1 = GetT2(D1);
                if (T1.Rows.Count > 0)
                {
                    dr = dtWeek.NewRow();
                    E2 = T1.Rows[0][2].ToString();
                    dr["SEQ"] = F.ToString();
                    dr["WK"] = E2;
                    dtWeek.Rows.Add(dr);
                }

            }
            UtilSimple.SetLookupBinding(comboBox1, dtWeek, "WK", "SEQ");
        }
        private void G1(string FTYPE,DataGridView G,string FF)
        {
            for (int i = 0; i <= G.Rows.Count - 2; i++)
            {
                DataGridViewRow row;

                row = G.Rows[i];
                string ITEMCODE = row.Cells["料號"].Value.ToString().Trim();
                string ITEMNAME = row.Cells["品名規格"].Value.ToString().Trim();

                for (int F = 1; F <= 10; F++)
                {
                    string D1 = "";

                    if (FF == "1")
                    {
                        string DATE = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
                        DateTime DD = Convert.ToDateTime(DATE);
                        if (F == 1)
                        {
                            D1 = DD.ToString("yyyyMMdd");
                        }
                        else
                        {
                            D1 = DD.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                        }
                    }
                    else
                    {
                        if (F == 1)
                        {
                            D1 = DateTime.Now.ToString("yyyyMMdd");
                        }
                        else
                        {
                            D1 = DateTime.Now.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                        }
                    }
           

                    System.Data.DataTable T1 = GetT2(D1);
                    if (T1.Rows.Count > 0)
                    {
                        S1 = T1.Rows[0][0].ToString();
                        string WK1 = row.Cells[1 + F].Value.ToString().Trim();
                        DELMAIN(ITEMCODE, S1, FTYPE);
                        AddFMAIN(ITEMCODE, S1, WK1, FTYPE);
                    }
                }
            }
            MessageBox.Show("更新完成");
        }
        private void GB_FMAIN_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();

      

            G0("外站Roots",dataGridView1,"");
            G0("棉花田", dataGridView2, "");
            G0("安永", dataGridView3, "");
            G0("電話傳真", dataGridView4, "");
            G0("員購", dataGridView5, "");
            G0("短效品", dataGridView6, "");
            G0("官網", dataGridView7, "");
            G0("批發", dataGridView8, "");
            G0("大宗樣品", dataGridView9, "");
            G0("其他銷貨", dataGridView10, "");
            G0("預計進貨", dataGridView11, "1");

            System.Data.DataTable H1 = RETAB();
            for (int i = 0; i <= H1.Rows.Count - 1; i++)
            {
                string CELLNAME = H1.Rows[i]["CELLNAME"].ToString().Replace("+", "");

                System.Data.DataTable H2 = RETAB2(globals.UserID.ToString().Trim(), CELLNAME);
                if (H2.Rows.Count > 0)
                {
                    tabControl1.TabPages.Remove(tabControl1.TabPages[CELLNAME]);
                }

            }

            GW();
        }
        private System.Data.DataTable GetT1()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,ITEMNAME FROM GB_FPRODUCT WHERE ISNULL(ENABLE,'') <> 'TRUE'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetT2(string D1)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SWEEK,EWEEK,WNUM  FROM GB_FWEEK WHERE @D1 BETWEEN SWEEK AND EWEEK ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@D1", D1));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetT3(string ITEMCODE, string STARTDAY, string FTYPE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT QTY FROM GB_FMAIN WHERE ITEMCODE=@ITEMCODE AND STARTDAY=@STARTDAY AND FTYPE=@FTYPE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@STARTDAY", STARTDAY));
            command.Parameters.Add(new SqlParameter("@FTYPE", FTYPE));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void G4(string ITEMCODE,int F,string FTYPE,string FF)
        {
            string D1 = "";

            if (FF == "1")
            {
                string DATE = textBox1.Text.Substring(0, 4) + "/" + textBox1.Text.Substring(4, 2) + "/" + textBox1.Text.Substring(6, 2);
                DateTime DD = Convert.ToDateTime(DATE);

                if (F == 1)
                {
                    D1 = DD.ToString("yyyyMMdd");
                }
                else
                {
                    D1 = DD.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                }
            }
            else
            {
                if (F == 1)
                {
                    D1 = DateTime.Now.ToString("yyyyMMdd");
                }
                else
                {
                    D1 = DateTime.Now.AddDays((F - 1) * 7).ToString("yyyyMMdd");
                }

            }

            System.Data.DataTable T1 = GetT2(D1);
            if (T1.Rows.Count > 0)
            {
                 SWEEK = T1.Rows[0][0].ToString();

                 System.Data.DataTable T2 = GetT3(ITEMCODE, SWEEK, FTYPE);
                if (T2.Rows.Count > 0)
                {
                    DAY = T2.Rows[0][0].ToString();
                }
                else
                {
                    DAY = "";
                }
            }
        }

        public void AddFMAIN(string ITEMCODE, string STARTDAY, string QTY, string FTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_FMAIN(ITEMCODE,STARTDAY,QTY,FTYPE) values(@ITEMCODE,@STARTDAY,@QTY,@FTYPE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@STARTDAY", STARTDAY));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@FTYPE", FTYPE));

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


        public void DELMAIN(string ITEMCODE, string STARTDAY, string FTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE GB_FMAIN WHERE ITEMCODE=@ITEMCODE AND STARTDAY=@STARTDAY AND FTYPE=@FTYPE", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@STARTDAY", STARTDAY));
            command.Parameters.Add(new SqlParameter("@FTYPE", FTYPE));

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
        public static System.Data.DataTable RETAB2(string USERID, string CELLNAME)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT CELLNAME FROM GB_FRIGHT WHERE  CELLNAME NOT IN (SELECT CELLNAME  FROM GB_FRIGHT WHERE USERID=@USERID ) AND CELLNAME=@CELLNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERID", USERID));
            command.Parameters.Add(new SqlParameter("@CELLNAME", CELLNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "inv1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["inv1"];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            G1("棉花田", dataGridView2, "0");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            G1("安永", dataGridView3, "0");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            G1("電話傳真", dataGridView4, "0");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            G1("員購", dataGridView5, "0");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            G1("短效品", dataGridView6, "0");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            G1("官網", dataGridView7, "0");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            G1("批發", dataGridView8, "0");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            G1("大宗樣品", dataGridView9, "0");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            G1("其他銷貨", dataGridView10, "0");
        }

        public static System.Data.DataTable RETAB()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT DISTINCT CELLNAME FROM GB_FRIGHT ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "inv1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["inv1"];
        }

        private void button11_Click(object sender, EventArgs e)
        {
            G0("預計進貨", dataGridView11, "1");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            G1("預計進貨", dataGridView11, "1");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView3);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView4);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView5);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView6);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView7);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView9);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView10);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView11);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtWeek = MakeTableWeek3();
            DataRow dr2 = null;
            string F = comboBox1.SelectedValue.ToString();
            int i = this.dataGridView11.Rows.Count - 2;
            int G = 0;
            string NumberName = "GBF" + DateTime.Now.ToString("yyyyMMdd");
            string AutoNum = util.GetAutoNumber(globals.Connection, NumberName);
            if (dataGridView11.Rows.Count > 1)
            {
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string QTY = dataGridView11.Rows[iRecs].Cells["WK-" + F].Value.ToString();
                    if (QTY != "0" && QTY != "")
                    {
                        G++;

                        string ITEMCODE = dataGridView11.Rows[iRecs].Cells["料號"].Value.ToString();
                        string ITEMNAME = dataGridView11.Rows[iRecs].Cells["品名規格"].Value.ToString();
                        string DOCDATE = GetMenu.Day();
                        string DOCDATE1 = DateTime.Now.AddDays(1).ToString("yyyyMMdd");
                        dr2 = dtWeek.NewRow();
                        dr2["RGN"] = "RGN";
                        dr2["ID"] = DOCDATE + AutoNum;
                        dr2["LINE"] = G.ToString();
                        dr2["DOCDATE"] = DOCDATE;
                        dr2["DOCDATE1"] = DOCDATE1;
                        dr2["ITEMCODE"] = ITEMCODE;
                        dr2["ITEMNAME"] = ITEMNAME;
                        dr2["QTY"] = QTY;
                        dr2["NONE"] = "";
                        dtWeek.Rows.Add(dr2);
                    }

                }
            }

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\GW\\入倉明細.xls";

            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report
            ExcelReport.ExcelReportOutput(dtWeek, ExcelTemplate, OutPutFile, "N");
        }

  
    }
}

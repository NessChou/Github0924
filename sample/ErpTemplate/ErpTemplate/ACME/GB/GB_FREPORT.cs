using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
namespace ACME
{
    public partial class GB_FREPORT : Form
    {
        string II = "";
        string DAY = "";
        int QTY = 0;
        string SWEEK = "";
        string SWNUM = "";
        int FF1 = 0;
        int FF2 = 0;
        int FF3 = 0;
        int FF4 = 0;
        int FF5 = 0;
        int FF6 = 0;
        int FG1 = 0;
        int FG2 = 0;
        int FG3 = 0;
        int FG4 = 0;
        int FG5 = 0;
        int FG6 = 0;
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_FREPORT()
        {
            InitializeComponent();
        }

        private void GB_FREPORT_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();


            G0(dataGridView11);

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
        private void G0(DataGridView G)
        {
            try
            {
                System.Data.DataTable dtWeek = MakeTableWeek();
                System.Data.DataTable dt = GetT1();
                DataRow dr = null;

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtWeek.NewRow();
                    string ITEMCODE = dt.Rows[i]["ITEMCODE"].ToString();
                    //if (ITEMCODE == "MCK010001")
                    //{
                    //    MessageBox.Show("A");
                    //}
                    II = ITEMCODE;
             
                    dr["料號"] = ITEMCODE;
                    dr["品名規格"] = dt.Rows[i]["ITEMNAME"].ToString();
                    System.Data.DataTable G1 = GetSTOCK(ITEMCODE);
                    int STOCK = 0;
                    if (G1.Rows.Count > 0)
                    {
                        STOCK = Convert.ToInt16(G1.Rows[0][0]);
                        dr["正航庫存"] = G1.Rows[0][0].ToString();
                    }

                    for (int F = 1; F <= 6; F++)
                    {
                        int DAY1 = 0;
                        G4(ITEMCODE, F, "2");
                        DAY1 = QTY;
                        dr["WK-" + F.ToString() + "     預訂進貨"] = DAY;
                        if (F == 1)
                        {
                            FG1 = DAY1;
                        }
                        if (F == 2)
                        {
                            FG2 = DAY1;
                        }
                        if (F == 3)
                        {
                            FG3 = DAY1;
                        }
                        if (F == 4)
                        {
                            FG4 = DAY1;
                        }
                        if (F == 5)
                        {
                            FG5 = DAY1;
                        }
                        if (F == 6)
                        {
                            FG6 = DAY1;
                        }
                    }
                    ArrayList al = new ArrayList();
                    for (int F = 1; F <= 6; F++)
                    {
                        G4(ITEMCODE, F, "1");
                        dr["WK-" + F.ToString() + "     預估銷量"] = DAY;

                        if (!String.IsNullOrEmpty(DAY))
                        {
                            al.Add(DAY + ",");
                        }
                        
                    }

                    for (int F = 1; F <= 6; F++)
                    {
           
                        int DAY2 = 0;
                 
                 
                    
                        G4(ITEMCODE, F, "1");
                        DAY2 = QTY;


                        if (F == 1)
                        {
                            FF1 = DAY2;
                        }
                        if (F == 2)
                        {
                            FF2 = DAY2;
                        }
                        if (F == 3)
                        {
                            FF3 = DAY2;
                        }
                        if (F == 4)
                        {
                            FF4 = DAY2;
                        }
                        if (F == 5)
                        {
                            FF5 = DAY2;
                        }
                        if (F == 6)
                        {
                            FF6 = DAY2;
                        }
                    }
                    for (int F = 1; F <= 6; F++)
                    {
                        if (F == 1)
                        {
                            dr["WK-1     預估庫存"] = STOCK - FF1 + FG1;
                        }
                        if (F == 2)
                        {
                            dr["WK-2     預估庫存"] = STOCK - FF1 - FF2 + FG1 + FG2 ;
                        }
                        if (F == 3)
                        {
                            dr["WK-3     預估庫存"] = STOCK - FF1 - FF2 - FF3 + FG1 + FG2 + FG3;
                        }
                        if (F == 4)
                        {
                            dr["WK-4     預估庫存"] = STOCK - FF1 - FF2 - FF3 - FF4 + FG1 + FG2 + FG3 + FG4;
                        }
                        if (F == 5)
                        {
                            dr["WK-5     預估庫存"] = STOCK - FF1 - FF2 - FF3 - FF4 - FF5 + FG1 + FG2 + FG3 + FG4;
                        }
                        if (F == 6)
                        {
                            dr["WK-6     預估庫存"] = STOCK - FF1 - FF2 - FF3 - FF4 - FF5 - FF6 + FG1 + FG4;
                        }
                    }
                    for (int F = 1; F <= 6; F++)
                    {

                        if (al.Count == 6)
                        {
                            int M1 = Convert.ToInt32(dr["WK-" + F.ToString() + "     預估庫存"]);
                            int M2 = 0;
                            if (F == 1)
                            {
                                M2 = FF2 + FF3 + FF4 + FF5 + FF6;
                                dr["WK-1     庫存可賣"] = ((M1 - M2) / GetMedian(al) + 5).ToString("#,##0");
                            }
                            if (F == 2)
                            {
                                M2 = FF3 + FF4 + FF5 + FF6;
                                dr["WK-2     庫存可賣"] = ((M1 - M2) / GetMedian(al) + 4).ToString("#,##0");
                            }
                            if (F == 3)
                            {
                                M2 = FF4 + FF5 + FF6;
                                dr["WK-3     庫存可賣"] = ((M1 - M2) / GetMedian(al) + 3).ToString("#,##0");
                            }
                            if (F == 4)
                            {
                                M2 = FF5 + FF6;
                                dr["WK-4     庫存可賣"] = ((M1 - M2) / GetMedian(al) + 2).ToString("#,##0");
                            }
                            if (F == 5)
                            {
                                M2 = FF6;
                                dr["WK-5     庫存可賣"] = ((M1 - M2) / GetMedian(al) + 1).ToString("#,##0");
                            }
                            if (F == 6)
                            {
                                dr["WK-6     庫存可賣"] = ((M1) / GetMedian(al)).ToString("#,##0");
                            }

                        }
                    }

                    dtWeek.Rows.Add(dr);
                }

                G.DataSource = dtWeek;

                System.Data.DataTable T1 = GetT2(textBox1.Text);
                if (T1.Rows.Count > 0)
                {
                    int H1 = Convert.ToInt32(T1.Rows[0][2]);
                    for (int F = 0; F <= 5; F++)
                    {
                        G.Columns[3 + F].HeaderText = "WK-" + (H1 + F).ToString() + "     預訂進貨";
                    }
                    for (int F = 0; F <= 5; F++)
                    {
                        G.Columns[9 + F].HeaderText = "WK-" + (H1 + F).ToString() + "     預估銷量";
                    }
                    for (int F = 0; F <= 5; F++)
                    {
                        G.Columns[15 + F].HeaderText = "WK-" + (H1 + F).ToString() + "     預估庫存";
                    }
                    for (int F = 0; F <= 5; F++)
                    {
                        G.Columns[21 + F].HeaderText = "WK-" + (H1 + F).ToString() + "     庫存可賣";
                    }
                }
         
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + II);
            }


        }



        private double GetMedian(ArrayList aryData)
        {

            double _value = 0;



            if (aryData.Count % 2 == 0)//數量為偶數
            {

                int _index = aryData.Count / 2;

                double valLeft = double.Parse(aryData[_index - 1].ToString());

                double valRight = double.Parse(aryData[_index].ToString());

                _value = (valLeft + valRight) / 2;

            }

            else//數量為奇數
            {

                int _index = (aryData.Count + 1) / 2;

                _value = double.Parse(aryData[_index - 1].ToString());

            }



            return _value;

        }

 

        private System.Data.DataTable MakeTableWeek()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("料號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("正航庫存", typeof(string));
            dt.Columns.Add("WK-1     預訂進貨", typeof(string));
            dt.Columns.Add("WK-2     預訂進貨", typeof(string));
            dt.Columns.Add("WK-3     預訂進貨", typeof(string));
            dt.Columns.Add("WK-4     預訂進貨", typeof(string));
            dt.Columns.Add("WK-5     預訂進貨", typeof(string));
            dt.Columns.Add("WK-6     預訂進貨", typeof(string));
            dt.Columns.Add("WK-1     預估銷量", typeof(string));
            dt.Columns.Add("WK-2     預估銷量", typeof(string));
            dt.Columns.Add("WK-3     預估銷量", typeof(string));
            dt.Columns.Add("WK-4     預估銷量", typeof(string));
            dt.Columns.Add("WK-5     預估銷量", typeof(string));
            dt.Columns.Add("WK-6     預估銷量", typeof(string));
            dt.Columns.Add("WK-1     預估庫存", typeof(string));
            dt.Columns.Add("WK-2     預估庫存", typeof(string));
            dt.Columns.Add("WK-3     預估庫存", typeof(string));
            dt.Columns.Add("WK-4     預估庫存", typeof(string));
            dt.Columns.Add("WK-5     預估庫存", typeof(string));
            dt.Columns.Add("WK-6     預估庫存", typeof(string));
            dt.Columns.Add("WK-1     庫存可賣", typeof(string));
            dt.Columns.Add("WK-2     庫存可賣", typeof(string));
            dt.Columns.Add("WK-3     庫存可賣", typeof(string));
            dt.Columns.Add("WK-4     庫存可賣", typeof(string));
            dt.Columns.Add("WK-5     庫存可賣", typeof(string));
            dt.Columns.Add("WK-6     庫存可賣", typeof(string));
            return dt;
        }
        private void G4(string ITEMCODE, int F, string FTYPE)
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
                SWEEK = T1.Rows[0][0].ToString();
                SWNUM = T1.Rows[0][2].ToString();
                System.Data.DataTable T2 = GetT3(ITEMCODE, SWEEK, FTYPE);
                if (T2.Rows.Count > 0)
                {
                    DAY = T2.Rows[0][0].ToString();
                    QTY = Convert.ToInt16(T2.Rows[0][0]);
                    if (DAY == "0")
                    {
                        DAY = "";
                    }
                }
         
            }
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
            if (FTYPE == "1")
            {
                sb.Append(" SELECT ISNULL(SUM(CAST(QTY AS INT)),0) QTY FROM GB_FMAIN WHERE ITEMCODE=@ITEMCODE AND STARTDAY=@STARTDAY AND FTYPE <>  '預計進貨' ");
            }
            if (FTYPE == "2")
            {
                sb.Append(" SELECT ISNULL(SUM(CAST(QTY AS INT)),0) QTY FROM GB_FMAIN WHERE ITEMCODE=@ITEMCODE AND STARTDAY=@STARTDAY AND FTYPE =  '預計進貨' ");
            }
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
        public System.Data.DataTable GetSTOCK(string ProdID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("     SELECT ISNULL(CAST(SUM(Quantity) AS INT),0) QTY  FROM comWareAmount where ProdID=@ProdID AND WareID IN ('A08','A14')  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        private void button11_Click(object sender, EventArgs e)
        {
            G0(dataGridView11);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView11);
        }


    
    }
}

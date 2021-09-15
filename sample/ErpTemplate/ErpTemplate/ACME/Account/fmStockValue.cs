using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Data.SqlClient;

//Excel
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class fmStockValue : Form
    {

        public fmStockValue()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt = MakeTable();

                DataRow dr;


                DateTime StartDate = StrToDate(textBox1.Text);
                DateTime EndDate = StrToDate(textBox2.Text);
                while (StartDate <= EndDate)
                {

                    dr = dt.NewRow();
                    dr["���"] = DateToStr(StartDate);
                    dr["��ؾl�B"] = GetAccAmount(DateToStr(StartDate));
                    dr["�s�f����"] = GetStockAmount(DateToStr(StartDate));
                    dr["�t��"] = Convert.ToInt32(dr["��ؾl�B"]) - Convert.ToInt32(dr["�s�f����"]);

                    dt.Rows.Add(dr);

                    StartDate = StartDate.AddDays(1);
                }

                dataGridView1.AutoGenerateColumns = false;

                dataGridView1.DataSource = dt;

                System.Data.DataTable df = GetCountry(textBox1.Text, textBox2.Text);
                dataGridView2.DataSource = df;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
                
        }

        private void fmStockValue_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        //����B�z--------------------------------------------------------------------------------------------
        private DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }

        //���X�W�@�Ӥ�
        private string GetPriorMonth()
        {

            return DateToStr(DateTime.Today.AddMonths(-1)).Substring(0, 6);
            
        }

        //���X�W�@�Ӥ�
        private string GetPriorMonthDate(string date)
        {

            int year = Convert.ToInt32(date.Substring(0, 4));
            int month = Convert.ToInt32(date.Substring(4, 2));

            //���o���Ѽ�
            int days = DateTime.DaysInMonth(year, month);

            return date.Substring(0, 6) + days;

        }



        //����B�z--------------------------------------------------------------------------------------------
        //�ʺA���͸�Ƶ��c
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("��ؾl�B", typeof(Int32));
            dt.Columns.Add("�s�f����", typeof(Int32));
            dt.Columns.Add("�t��", typeof(Int32));
            
            /*
            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["SERIAL_NO"];
            dt.PrimaryKey = colPk;
            */

            //�g�J���
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "�q��i��";
            //dt.Rows.Add(dr);
            return dt;
        }

        //���o��ؾl�B
        private Int32 GetAccAmount(string RefDate)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SUM(T0.[Debit])- SUM(T0.[Credit]) �l�B ");
            sb.Append("FROM  [dbo].[JDT1] T0 inner join OACT T1 on T0.Account = T1.AcctCode  ");
            sb.Append("WHERE T0.[RefDate] <= @RefDate   ");
            sb.Append("AND  T0.[Account]  like '12000%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@RefDate", RefDate));

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

            return  Convert.ToInt32(ds.Tables[0].Rows[0]["�l�B"]);

        }


        //���o���v�w�s�l�B
        private Int32 GetStockAmount(string RefDate)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT sum(T0.TransValue) �w�s���B ");
            sb.Append("FROM  [dbo].[OINM] T0   left JOIN OITM T11 ON T0.ITEMCODE = T11.ITEMCODE     ");
            sb.Append("WHERE  T0.[docdate] <= @RefDate   ");
            sb.Append("     and ISNULL(T11.U_GROUP,'') <> 'Z&R-�O�����s��' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@RefDate", RefDate));

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

            return Convert.ToInt32(ds.Tables[0].Rows[0]["�w�s���B"]);

        }


        private void button3_Click(object sender, EventArgs e)
        {
            ACME.CheckDetail frm = new ACME.CheckDetail();
            frm.ShowDialog();
        }
        private System.Data.DataTable GetCountry(string STARTDATE, string ENDDATE)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT Convert(varchar(8),(T0.docdate),112) ���,CASE TRANSTYPE WHEN 59 THEN '���f��' ");
            sb.Append("              WHEN 60 THEN '�o�f��' WHEN 67 THEN '�w�s�ռ�' END ���,(T3.WHSNAME) �ܮw,BASE_REF �渹,");
            sb.Append("              CAST(CASE WHEN ISNULL((T1.TRANSID),'') ='' THEN  (T2.TRANSID) ELSE (T1.TRANSID) END AS NVARCHAR) �ǲ����X");
            sb.Append("              ,T0.ITEMCODE �Ƹ�,cast((TransValue) as int)*-1 ���B,CAST(OUTQTY-INQTY AS INT) �ƶq");
            sb.Append("              FROM OINM T0");
            sb.Append("              LEFT JOIN OIGN T1 ON (T0.BASE_REF=T1.DOCENTRY AND T0.TRANSTYPE=59)");
            sb.Append("              LEFT JOIN OIGE T2 ON (T0.BASE_REF=T2.DOCENTRY AND T0.TRANSTYPE=60)");
            sb.Append("              LEFT JOIN OWHS T3 ON(T0.WAREHOUSE=T3.WHSCODE)   left JOIN OITM T11 ON T0.ITEMCODE = T11.ITEMCODE   ");
            sb.Append("              WHERE  ISNULL(T11.U_GROUP,'') <> 'Z&R-�O�����s��'  AND T0.[docdate] BETWEEN @STARTDATE AND @ENDDATE");
            sb.Append("              AND WAREHOUSE IN ('CC001','CC002') AND TRANSTYPE IN (59,60) ");
            sb.Append("              UNION ALL ");
            sb.Append("              SELECT '' ���,'' ���,'' �ܮw,'' �渹,'','�[�`',cast(SUM(TransValue) as int)*-1 ���B,SUM(OUTQTY)-SUM(INQTY) �ƶq");
            sb.Append("              FROM OINM T0     left JOIN OITM T11 ON T0.ITEMCODE = T11.ITEMCODE ");
            sb.Append("              WHERE   ISNULL(T11.U_GROUP,'') <> 'Z&R-�O�����s��' AND T0.[docdate] BETWEEN @STARTDATE AND @ENDDATE");
            sb.Append("              AND WAREHOUSE IN ('CC001','CC002') AND TRANSTYPE IN (59,60) ORDER BY T0.ITEMCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@STARTDATE", STARTDATE));
            command.Parameters.Add(new SqlParameter("@ENDDATE", ENDDATE));
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1); 
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2); 
            }
        }

 
      
    }
}
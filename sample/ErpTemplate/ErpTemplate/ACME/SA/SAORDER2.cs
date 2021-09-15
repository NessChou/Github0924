using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace ACME
{
    public partial class SAORDER2 : Form
    {
        string FA = "acmesql02";
       
        public SAORDER2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
   
            System.Data.DataTable G2 = GetORDR();
            dataGridView1.DataSource = G2;
            dataGridView1.Columns[0].Visible  = false;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[4].ReadOnly = true;
            dataGridView1.Columns[5].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;
        }

        public  System.Data.DataTable GetORDR()
        {
            SqlConnection MyConnection = globals.shipConnection;
            string OHEM = fmLogin.LoginID.ToString().ToUpper();
            StringBuilder sb = new StringBuilder();


            sb.Append("SELECT T0.DOCENTRY SO,T1.LINENUM LINENUM,T0.CARDNAME 客戶,Convert(varchar(10),T0.CREATEDATE,102) 建立日期,T0.DOCENTRY 訂單單號, ");
            sb.Append("ISNULL(Convert(varchar(10),T2.[U_ACME_SHIPDAY] ,102),Convert(varchar(10),T1.[U_ACME_SHIPDAY] ,102))  原始離倉日期,Convert(varchar(10),T1.[U_ACME_SHIPDAY] ,102)   目前離倉日期, ");
            sb.Append("DATEDIFF(D,ISNULL(Convert(varchar(10),T2.[U_ACME_SHIPDAY] ,102),Convert(varchar(10),T1.[U_ACME_SHIPDAY] ,102)),T1.[U_ACME_SHIPDAY])-30 逾期天數,T1.DSCRIPTION  Model,(CAST(t1.opencreqty AS INT)) 未結數量, ");
            sb.Append("CAST(T3.ONHAND AS INT) 庫存數量,T1.CURRENCY+' '+CAST(CAST(T1.PRICE AS DECIMAL(16,2)) AS VARCHAR) 單價,T5.REMARK 原因說明 FROM ORDR T0  ");
            sb.Append("INNER JOIN RDR1 T1 ON (T0.DocEntry = T1.DocEntry)       ");
            sb.Append("LEFT JOIN ADO1 T2 ON (T1.DocEntry =T2.DocEntry AND T1.LineNum = T2.LineNum AND T2.LogInstanc=1 AND T2.ObjType =17 ) ");
            sb.Append("INNER  JOIN [dbo].[OITM] T3  ON  T1.[ItemCode] = T3.ItemCode    ");
            sb.Append("LEFT JOIN ACMESQLSP.DBO.SA_ORDER2 T5 ON (T1.DOCENTRY=T5.DOCENTRY AND T1.LINENUM =T5.LINENUM)   ");
            sb.Append("left JOIN OSLP T6 ON T0.SlpCode = T6.SlpCode   ");
            sb.Append("left JOIN OHEM T7 ON T0.OwnerCode = T7.empID   ");
            sb.Append("WHERE  T1.LineStatus ='O'  ");

            if (textBox1.Text.ToString() != "")
            {
                sb.Append("AND  (DATEDIFF(D,T2.[U_ACME_SHIPDAY],T1.[U_ACME_SHIPDAY])-30-" + textBox1.Text.ToString() + ") > 0  ");
            }
            //

            if (artextBox12.Text != "")
            {
                sb.Append(" and  T0.[cardname] like '%" + artextBox12.Text.ToString() + "%'");
            }
            if (comboBox1.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and T6.SlpName =@SlpName  ");
            }
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append(" and T7.[lastName]+T7.[firstName]=@lastName ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
           
            command.Parameters.Add(new SqlParameter("@lastName", comboBox2.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@SlpName", comboBox1.SelectedValue.ToString()));
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
        private System.Data.DataTable GetSO(int DOCENTRY, int LINENUM)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LINENUM FROM SA_ORDER2 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSO2(int DOCENTRY, int LINENUM,string DOCDATE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(10),dbo.fun_SHIPDATE(U_ACME_WORKDAY,@DOCDATE),112)   FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {

                DataGridViewRow row;

                row = dataGridView1.Rows[i];
                int SO = Convert.ToInt32(row.Cells["SO"].Value);
                int LINENUM = Convert.ToInt16(row.Cells["LINENUM"].Value);
  
                string 原因說明 = row.Cells["原因說明"].Value.ToString();



                System.Data.DataTable GO = GetSO(SO, LINENUM);

                if (GO.Rows.Count > 0)
                {

                    UPDATEORDER(原因說明, SO, LINENUM);
                }
                else
                {

                    AddORDER(原因說明, SO, LINENUM);
                }

            }

     
            MessageBox.Show("存檔成功");


        }


        public void AddORDER(string REMARK, int DOCENTRY, int LINENUM)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into SA_ORDER2(DOCENTRY,LINENUM,REMARK) values(@DOCENTRY,@LINENUM,@REMARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));

            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));

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

        private void UPDATEORDER(string REMARK, int DOCENTRY, int LINENUM)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE SA_ORDER2 SET REMARK=@REMARK WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));


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


        private void SAORDER_Load(object sender, EventArgs e)
        {

            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetOslp1(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetOhem(), "DataValue", "DataValue");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcelES(dataGridView1);
        }



    }
}

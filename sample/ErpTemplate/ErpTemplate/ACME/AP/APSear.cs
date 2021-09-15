using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ACME
{
    public partial class APSear : Form
    {
        private DataGridViewCellStyle defaultCellStyle;
        private DataGridViewCellStyle groupCellStyle;
        public APSear()
        {
            InitializeComponent();
        }



        private void ViewBatchPayment()
        {
      
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  select T0.docnum 單號,MAX(ISNULL(T0.LCNO,0)) LCNO, MAX(ISNULL(T0.lcamt,0)) 已沖金額,MAX(ISNULL(T0.lctotal,0)) 未沖金額");
            sb.Append("                from APLC T0 LEFT JOIN PLC1 T1 ON (T0.DocNum=T1.DocNum)");
            sb.Append(" WHERE isnull(T1.STATUS,0) <> 'True' ");
            sb.Append(" GROUP BY T0.docnum ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void APSear_Load(object sender, EventArgs e)
        {
            //ViewBatchPayment();
            this.defaultCellStyle = new DataGridViewCellStyle();

            this.groupCellStyle = new DataGridViewCellStyle();

            this.groupCellStyle.ForeColor = Color.White;

            this.groupCellStyle.BackColor = Color.YellowGreen;

            this.groupCellStyle.SelectionBackColor = Color.DarkBlue;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            System.Data.DataTable dt;
          
            if (radioButton1.Checked)
            {
              dt = GetMail(textBox3.Text,textBox4.Text);
            }
            else

            {
                dt = GetMail1(textBox3.Text, textBox4.Text);
            }
            string Key = string.Empty;
            string  Total = "0";
            


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                if (i > 0 && Key != Convert.ToString(dt.Rows[i]["LCNO"]))
                {
                    dataGridView1.Rows.Add(new object[14] {
                                "",
                               "",
                "",
                "",
                                   "",
                               "",
                "",
                "",
                                   "未沖金額",
                               Total,
                "",
                                "", 
                        "",
                                ""  });  
                }

                dataGridView1.Rows.Add(new object[14] {
                     Convert.ToString(dt.Rows[i]["銀行別"]),
                                Convert.ToString(dt.Rows[i]["LCNO"]),
                                Convert.ToString(dt.Rows[i]["幣別"]),
                    Convert.ToString(dt.Rows[i]["金額2"]),
                                Convert.ToString(dt.Rows[i]["單號"]), 
                               Convert.ToString(dt.Rows[i]["品名"]),
                     Convert.ToString(dt.Rows[i]["數量"]),
                     Convert.ToString(dt.Rows[i]["單價"]),
                     Convert.ToString(dt.Rows[i]["稅額"]),
                                Convert.ToString(dt.Rows[i]["金額"]), 
                  Convert.ToString(dt.Rows[i]["INVOICE"]), 
                  Convert.ToString(dt.Rows[i]["出貨時間"]), 
                            Convert.ToString(dt.Rows[i]["押匯時間"]), 
                  Convert.ToString(dt.Rows[i]["寄送時間"]) });
                Total = Convert.ToString(dt.Rows[i]["LcTotal"]);
                Key = Convert.ToString(dt.Rows[i]["LCNO"]);
            }


            dataGridView1.Rows.Add(new object[14] {
                                "",
                               "",
                "",
                "",
                                   "",
                               "",
                "",
                "",
                                   "未沖金額",
                               Total,
                "",
                "",
                                "", 
                                ""  });

        }
        public System.Data.DataTable GetMail(string DocDate1, string DocDate2)
        {
            SqlConnection MyConnection = globals.Connection;
            
            StringBuilder sb = new StringBuilder();
            sb.Append("              select t0.docnum,t0.bankName 銀行別,t0.lcNo LCNO,");
            sb.Append("              ISNULL(bankCode,'') 幣別,isnull(T0.LcTotal,0) LcTotal,T1.DonNo 單號,T1.Itemcode 品名,ISNULL(T1.Qty,0) 數量,ISNULL(T1.Price,0) 單價,ISNULL(T1.Tax,0) 稅額,isnull(T1.AMT,0) 金額,");
            sb.Append("                            isnull(cast((case when  (row_number()  over (partition by T0.LCAMT order by t0.docnum))=1 ");
            sb.Append("                            then T0.LCAMT else null end) as varchar),'') as 金額2, ");
            sb.Append("                                              ISNULL(T1.CargoDate,'') 出貨時間,ISNULL(T1.SendDate,'') 寄送時間 ,ISNULL(T1.CargoDate2,'') 押匯時間 ,ISNULL(T1.CardName,'') 公司,ISNULL(T1.InvoceNo,'') INVOICE");
            sb.Append("                                              from acmesqlsp.dbo.APLC T0 LEFT JOIN acmesqlsp.dbo.PLC1 T1 ON (T0.DocNum=T1.DocNum)");
            sb.Append("                                WHERE  t0.lcclose <> 'Checked'");
            if (comboBox1.Text != "")
            {
                if (comboBox1.Text == "押匯時間")
                {
                    sb.Append(" and  T1.[CargoDate2] between @DocDate1 and @DocDate2 ");
                }
                else if (comboBox1.Text == "寄送時間")
                {
                    sb.Append(" and  T1.[SendDate] between @DocDate1 and @DocDate2 ");
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
           
            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        public System.Data.DataTable GetMail1(string DocDate1, string DocDate2)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("              select t0.docnum,t0.bankName 銀行別,t0.lcNo LCNO,");
            sb.Append("              ISNULL(bankCode,'') 幣別,isnull(T0.LcTotal,0) LcTotal,T1.DonNo 單號,T1.Itemcode 品名,ISNULL(T1.Qty,0) 數量,ISNULL(T1.Price,0) 單價,ISNULL(T1.Tax,0) 稅額,isnull(T1.AMT,0) 金額,");
            sb.Append("                            isnull(cast((case when  (row_number()  over (partition by T0.LCAMT order by t0.docnum))=1 ");
            sb.Append("                            then T0.LCAMT else null end) as varchar),'') as 金額2, ");
            sb.Append("                                              ISNULL(T1.CargoDate,'') 出貨時間,ISNULL(T1.SendDate,'') 寄送時間 ,ISNULL(T1.CargoDate2,'') 押匯時間 ,ISNULL(T1.CardName,'') 公司,ISNULL(T1.InvoceNo,'') INVOICE");
            sb.Append("                                              from acmesqlsp.dbo.APLC T0 LEFT JOIN acmesqlsp.dbo.PLC1 T1 ON (T0.DocNum=T1.DocNum)");
            sb.Append("                                WHERE t0.lcclose = 'Checked' ");
            if (comboBox1.Text != "")
            {
                if (comboBox1.Text == "押匯時間")
                {
                    sb.Append(" and  T1.[CargoDate2] between @DocDate1 and @DocDate2 ");
                }
                else if (comboBox1.Text == "寄送時間")
                {
                    sb.Append(" and  T1.[SendDate] between @DocDate1 and @DocDate2 ");
                }
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            DataGridViewRow dr = new DataGridViewRow();
            //dr.Cells[0].Value = "total";
            if (e.ColumnIndex == 0 && e.RowIndex >= 0 && e.RowIndex != dgv.NewRowIndex)
            {
                if (dgv[8, e.RowIndex].Value.Equals("未沖金額"))
                {
                  //  dgv.Rows[e.RowIndex].DefaultCellStyle = this.groupCellStyle;
                    dgv.Rows[e.RowIndex].DefaultCellStyle = this.groupCellStyle;
             
                }
            }
        }
      
        

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void dataGridView1_MouseCaptureChanged(object sender, EventArgs e)
        {
          
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CalcTotals2();
        }
        private void CalcTotals2()
        {


            Int32 iTotal = 0;
            decimal iVatSum = 0;


            int i = this.dataGridView1.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                iTotal += Convert.ToInt32(dataGridView1.SelectedRows[iRecs].Cells["數量"].Value);
                iVatSum += Convert.ToDecimal(dataGridView1.SelectedRows[iRecs].Cells["金額"].Value);

            }

            textBox1.Text = iTotal.ToString("0");
            textBox2.Text = iVatSum.ToString();



        }

       
    }
}
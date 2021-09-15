using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class APOWHS : Form 
    {
        public string ITEMCODE;

        public string a,c;
        public APOWHS()
        {
            InitializeComponent();
        }

        private void AP_Load(object sender, EventArgs e)
        {
            ViewBatchPayment(ITEMCODE);
            
              
        }


        private void ViewBatchPayment(string ITEMCODE)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append("       SELECT * FROM (     SELECT  T0.U_ACME_INV INVOICE,Convert(varchar(8),T0.U_ACME_INVOICE,112) INVOICE日期,SUM(CAST(T1.QUANTITY AS INT))-SUM(ISNULL(T2.QUANTITY,0))  數量 FROM OPCH T0  ");
            sb.Append("                             LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("                             LEFT JOIN (SELECT ITEMCODE,INVOICE,SUM(CAST(CAST(QUANTITY AS DECIMAL(14,0)) AS INT)) QUANTITY FROM AcmeSqlSP.DBO.WH_ITEM   ");
            sb.Append("                            WHERE ISNULL(INVOICE,'') <> ''  ");
            sb.Append("                            GROUP BY ITEMCODE,INVOICE)  ");
            sb.Append("                             T2  ON (T1.ITEMCODE=T2.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.U_ACME_INV=T2.INVOICE COLLATE Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("   WHERE T1.ITEMCODE=@ITEMCODE ");
            sb.Append("                            GROUP BY T0.U_ACME_INV,T0.U_ACME_INVOICE  ");
            sb.Append("                            HAVING SUM(CAST(T1.QUANTITY AS INT))-SUM(ISNULL(T2.QUANTITY,0))  >0  ");
            sb.Append("             			   UNION ALL ");
            sb.Append("             			    SELECT  T0.U_ACME_INV INVOICE,Convert(varchar(8),T0.DOCDATE,112) INVOICE日期,SUM(CAST(T1.QUANTITY AS INT))-SUM(ISNULL(T2.QUANTITY,0))  數量 FROM OIGE T0  ");
            sb.Append("                             LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("                             LEFT JOIN (SELECT ITEMCODE,INVOICE,SUM(CAST(CAST(QUANTITY AS DECIMAL(14,0)) AS INT)) QUANTITY FROM AcmeSqlSP.DBO.WH_ITEM   ");
            sb.Append("                            WHERE ISNULL(INVOICE,'') <> ''  ");
            sb.Append("                            GROUP BY ITEMCODE,INVOICE)  ");
            sb.Append("                             T2  ON (T1.ITEMCODE=T2.ITEMCODE COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.U_ACME_INV=T2.INVOICE COLLATE Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("   WHERE T1.ITEMCODE=@ITEMCODE ");
            sb.Append("             			   GROUP BY T0.U_ACME_INV,T0.DOCDATE  ");
            sb.Append("                            HAVING SUM(CAST(T1.QUANTITY AS INT))-SUM(ISNULL(T2.QUANTITY,0))  >0  )AS A ");
            sb.Append("							WHERE INVOICE IN (");
            sb.Append("							SELECT INVOICE COLLATE  Chinese_Taiwan_Stroke_CI_AS FROM ACMESQLSP.DBO.WH_WEBSTOCK WHERE ITEMCODE=@ITEMCODE)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, " POR1");
            }
            finally
            {
                connection.Close();
            }


            dataGridView1.DataSource = ds.Tables[0];

        }
 



        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;
                    StringBuilder sb = new StringBuilder();
                    row = dataGridView1.SelectedRows[0];
           
                    //for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    //{

                    //    row = dataGridView1.SelectedRows[i];

                    //    sb.Append("'" + row.Cells["INVOICE"].Value.ToString() + " " + row.Cells["LINENUM"].Value.ToString() + "',");
                    //}
                 
             


                    //sb.Remove(sb.Length - 1, 1);

                    ////linenum
                    //string q = sb.ToString();

                    a = row.Cells["INVOICE"].Value.ToString() + "/" + row.Cells["INVOICE日期"].Value.ToString();
           
                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    


    


   
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class ACCSTOCK : Form
    {
        public ACCSTOCK()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Data.DataTable t1 = Gen_201004();
            dataGridView1.DataSource = t1;
        }
        private System.Data.DataTable Gen_201004()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;



            StringBuilder sb = new StringBuilder();

            //彙總
            sb.Append("               SELECT Convert(varchar(8),T0.[DocDate],112)  過帳日期, T0.CardName 客戶名稱,SUBSTRING(ITMSGRPNAM,4,5) 部門, ");
            sb.Append("               T2.[ItemCode] 料號,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本,");
            sb.Append("                  T11.U_LOCATION 地區,(T2.Quantity) 數量, ");
            sb.Append("               cast(cast(T5.[Price] as numeric(16,2)) as varchar) 訂單單價,T2.[Price] 台幣單價, ");
            sb.Append("               CAST(T2.LineTotal AS INT) 台幣金額, ");
            sb.Append("               CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT) 台幣毛利, ");
            sb.Append("               T3.SLPNAME 業務,'AR'+CAST(T0.DOCENTRY AS VARCHAR) 單據 FROM OINV T0  ");
            sb.Append("               INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode  ");
            sb.Append("               left join dln1 t4 on (t2.baseentry=T4.docentry and  t2.baseline=t4.linenum  and t2.basetype='15') ");
            sb.Append("               left join odln t9 on (t4.docentry=T9.docentry ) ");
            sb.Append("               left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15') ");
            sb.Append("               INNER JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               INNER JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod  ");
            sb.Append("               WHERE T0.[DocType] ='I'  ");
            sb.Append("               AND  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組' AND T0.U_IN_BSTYC <> '1'  ");

            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112) between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T0.[CardCode] ='" + textBox7.Text.ToString() + "' ");
            }
            sb.Append(" UNION ALL");
            sb.Append("              SELECT Convert(varchar(8),T0.[DocDate],112)  過帳日期, T0.CardName 客戶名稱,SUBSTRING(ITMSGRPNAM,4,5) 部門, ");
            sb.Append("              T2.[ItemCode] 料號,T11.U_TMODEL Model,T11.U_GRADE 等級, ''''+U_VERSION 版本, ");
            sb.Append("               T11.U_LOCATION 地區,(T2.Quantity)*-1 數量, ");
            sb.Append("               T2.U_ACME_INV 訂單單價,T2.[Price] 台幣單價, ");
            sb.Append("               CAST(T2.LineTotal AS INT)*-1 台幣金額, ");
            sb.Append("               (CAST((T2.LineTotal) - (Round(T2.StockPrice*T2.Quantity,0)) AS INT))*-1 台幣毛利, ");
            sb.Append("               T3.SLPNAME 業務,'貸項'+CAST(T0.DOCENTRY AS VARCHAR) 單據 FROM ORIN T0  ");
            sb.Append("               INNER JOIN RIN1 T2 ON T0.DocEntry = T2.DocEntry  ");
            sb.Append("               INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode  ");
            sb.Append("               INNER JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("               INNER JOIN OITB T12 ON T11.itmsgrpcod = T12.itmsgrpcod  ");
            sb.Append("               WHERE T0.[DocType] ='I'   ");
            sb.Append("               and  ISNULL(T11.U_GROUP,'') <> 'Z&R-費用類群組'   ");


            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112) between @DocDate1 and @DocDate2 ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  T0.[CardCode] ='" + textBox7.Text.ToString() + "' ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox5.Text));

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox6.Text));





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

        private void ACCSTOCK_Load(object sender, EventArgs e)
        {

            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }
    }
}

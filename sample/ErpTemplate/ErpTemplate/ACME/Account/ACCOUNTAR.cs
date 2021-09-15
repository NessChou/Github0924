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
    public partial class ACCOUNTAR : Form
    {
        public ACCOUNTAR()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Gen_201005();
        }

        private System.Data.DataTable Gen_201005()
        {

            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                                           SELECT  T2.ITEMCODE 項目料號,T2.DSCRIPTION 項目說明,T0.CARDCODE 客戶編號,case when T0.cardname like  '%TOP GARDEN INT%' then 'TOP GARDEN' when T0.cardname like  '%CHOICE CHANNEL%' then 'CHOICE' when T0.cardname like  '%Infinite Power Group%' then 'INFINITE' when T0.cardname like  '%宇豐光電股份有限公司%' then '宇豐' when T0.cardname like  '%達睿生%' then 'DRS' else t0.cardname end+CASE ISNULL(T0.U_BENEFICIARY,'') WHEN '' THEN '' ELSE '-'+T0.U_BENEFICIARY END 客戶名稱,Convert(varchar(8),T0.[DocDate],112)  過帳日期,   ");
            sb.Append("                                                             (CAST(T2.Quantity AS INT)) 數量,T2.[Price] 單價,  ");
            sb.Append("                                                      CAST(T2.LineTotal AS INT) 銷售金額,CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END 成本,  ");
            sb.Append("                                                      CAST((T2.LineTotal) - (CASE WHEN  T0.UpdInvnt='C'  THEN CAST(Round(T2.GROSSBUYPR*T2.Quantity,0) AS INT) ELSE CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT) END) AS INT) 毛利, T3.SLPNAME 業務,T33.[lastName]+T33.[firstName] 業管,  ");
            sb.Append("                                                     CAST(T0.DOCENTRY AS VARCHAR) AR單號,T5.DOCENTRY 銷售單號 FROM OINV T0   ");
            sb.Append("                                                      INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry   ");
            sb.Append("                               left join dln1 t4 on (t2.baseentry=T4.docentry and  t2.baseline=t4.linenum  and t2.basetype='15')  ");
            sb.Append("                              left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')  ");
            sb.Append("                                         INNER JOIN OSLP T3 ON T0.SlpCode = T3.SlpCode  left  JOIN OHEM T33 ON T0.OwnerCode = T33.empID  ");
            sb.Append("                                   left JOIN OITM T11 ON T2.ITEMCODE = T11.ITEMCODE  ");
            sb.Append("              WHERE 1=1 ");

            if (textBox1.Text != "" && textBox2.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T0.[DocDate],112)  between @DocDate1 and @DocDate2 ");
            }

            if (textBox3.Text != "")
            {
                sb.Append(" and  T0.DOCENTRY > @AR ");
            }

            sb.Append("  ORDER BY T0.DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@AR", textBox3.Text));




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

        private void ACCOUNTAR_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
            textBox2.Text = GetMenu.Day();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class ACCAP : Form
    {
        public ACCAP()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getcc();


            for (int i = 8; i <= 13; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                if (i == 8)
                {
                    col.DefaultCellStyle.Format = "#,##0";
                }
                else if (i == 10)
                {
                    col.DefaultCellStyle.Format = "#,##0.00";
                }
                else
                {
                    col.DefaultCellStyle.Format = "#,##0.0000";
                }


            }
        }

        public System.Data.DataTable Getcc()
        {
            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AP' 類別,Convert(varchar(8),T0.docdate,112) 過帳日期,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,T5.DOCENTRY 採購單號碼,T0.DOCENTRY AP號碼 ");
            sb.Append(" ,t1.itemcode 項目號碼,T1.DSCRIPTION 項目說明,cast(T1.quantity as int) 數量,t5.price '美金單價(未稅)', ");
            sb.Append(" CASE T5.TotalFrgn WHEN 0 THEN 0 ELSE CAST(ROUND((T1.[LineTotal]/T5.TotalFrgn),2) AS DECIMAL(18,2)) END AR匯率,T1.[LineTotal]  '總計(未稅)〈本幣〉',T1.vatsum  '稅額(VAT)' ,T1.[LineTotal]+T1.vatsum '總計(含稅)〈本幣〉'  ");
            sb.Append(" ,T1.AcctCode 總帳科目,T1.WHSCODE 倉庫,T10.WHSNAME 倉庫名稱 FROM OPCH T0   ");
            sb.Append(" LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline) ");
            sb.Append(" LEFT JOIN POR1 T5 ON (T5.docentry=T7.baseentry AND T5.linenum=T7.baseline) ");
            sb.Append(" LEFT JOIN OWHS T10 ON (T1.WHSCODE=T10.WHSCODE)  ");
            sb.Append(" WHERE ISNULL(cast(t5.price as varchar),'') <> ''  ");
            sb.Append(" AND  Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2  ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT  'AP貸',Convert(varchar(8),T0.docdate,112) 過帳日期,T0.CARDCODE 客戶編號,T0.CARDNAME 客戶名稱,'' 訂單號碼,T0.DOCENTRY AR號碼  ");
            sb.Append(" ,t1.itemcode 項目號碼,T1.DSCRIPTION 項目說明,cast(T1.quantity as int) 數量,REPLACE(T1.U_ACME_INV,',','') '美金單價(未稅)',  ");
            sb.Append(" CASE  WHEN t1.U_ACME_INV='0' THEN '0' WHEN  ISNUMERIC(t1.U_ACME_INV)=0 THEN '0' ELSE CAST(ROUND((T1.[LineTotal]/(CAST(REPLACE(T1.U_ACME_INV,',','') AS DECIMAL(10,2))*(CASE Quantity WHEN 0 THEN 1 ELSE Quantity END))),2) AS DECIMAL(18,2)) END AR匯率,  ");
            sb.Append(" T1.[LineTotal]*-1  '總計(未稅)〈本幣〉',T1.vatsum*-1  '稅額(VAT)' ,(T1.[LineTotal]+T1.vatsum)*-1 '總計(含稅)〈本幣〉'   ");
            sb.Append(" ,T1.AcctCode 總帳科目,T1.WHSCODE 倉庫,T10.WHSNAME 倉庫名稱 FROM ORPC T0    ");
            sb.Append(" INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" LEFT JOIN OWHS T10 ON (T1.WHSCODE=T10.WHSCODE)   ");
            sb.Append(" WHERE    Convert(varchar(8),t0.docdate,112)  between @DocDate1 and @DocDate2   ORDER BY 類別,T0.DOCENTRY   ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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

        private void ACCAR_Load(object sender, EventArgs e)
        {
            textBox4.Text = GetMenu.DFirst();

            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

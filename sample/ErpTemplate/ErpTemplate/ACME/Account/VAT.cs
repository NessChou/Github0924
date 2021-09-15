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
    public partial class VAT : Form
    {
        public VAT()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetItem();
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        private System.Data.DataTable GetItem()
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END [TYPE],");
            sb.Append(" P1.[CardCode] 廠商代號,P1.[CardName] 廠商名稱,");
            sb.Append(" CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.[U_PC_BSNOT]  ELSE T1.[U_PC_BSNOT] END  as 統一編號,");
            sb.Append(" CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSINV ELSE T1.U_PC_BSINV END as 發票號碼,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int)  未稅金額,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int)   稅額,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)  含稅總額 FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND Substring(Convert(varchar(8), T1.[U_PC_BSAPP],112),1,8)   BETWEEN @A1 AND @A2   ) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID AND  Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,8)  BETWEEN @A1 AND @A2  )");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");





            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox2.Text));


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
        private System.Data.DataTable GetItem2()
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END [TYPE],P1.[CardCode] 廠商代號,P1.[CardName] 廠商名稱,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  未稅金額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  稅額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) 含稅總額 FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND  Substring(Convert(varchar(8), T1.[U_PC_BSAPP],112),1,4)  =@A1  ) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID  AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1 )");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END ");
            sb.Append(" ORDER BY P1.[CardCode]");




            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox3.Text));
        


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

        private System.Data.DataTable GetItem3()
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  月份, [TYPE] ,廠商代號,廠商名稱,未稅金額,稅額,含稅總額 FROM (SELECT '1' LINE, CASE WHEN ISNULL(P.[U_PC_BSAPP],'') = '' THEN Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),5,2) ELSE Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),5,2) END 月份,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END [TYPE],P1.[CardCode] 廠商代號,P1.[CardName] 廠商名稱,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  未稅金額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  稅額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) 含稅總額 FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),1,4) =@A1  ) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1  )");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END,CASE WHEN ISNULL(P.[U_PC_BSAPP],'') = '' THEN Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),5,2) ELSE Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),5,2) END");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2' LINE,'TOTAL' 月份,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END [TYPE],P1.[CardCode] 廠商代號,P1.[CardName] 廠商名稱,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  未稅金額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  稅額,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) 含稅總額 FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry  AND Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),1,4)  =@A1) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID  AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1)");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '進貨'");
            sb.Append(" WHEN 1  THEN '費用'");
            sb.Append(" WHEN 2  THEN '固定資產' END END ) AS A");
            sb.Append(" ORDER BY A.廠商代號,[TYPE],月份");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@A1", textBox4.Text));



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
        private void VAT_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            textBox3.Text = DateTime.Now.ToString("yyyy");
            textBox4.Text = DateTime.Now.ToString("yyyy");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetItem2();
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetItem3();
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}
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
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END [TYPE],");
            sb.Append(" P1.[CardCode] �t�ӥN��,P1.[CardName] �t�ӦW��,");
            sb.Append(" CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.[U_PC_BSNOT]  ELSE T1.[U_PC_BSNOT] END  as �Τ@�s��,");
            sb.Append(" CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSINV ELSE T1.U_PC_BSINV END as �o�����X,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int)  ���|���B,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int)   �|�B,");
            sb.Append(" cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)  �t�|�`�B FROM OJDT  J");
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
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END [TYPE],P1.[CardCode] �t�ӥN��,P1.[CardName] �t�ӦW��,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  ���|���B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  �|�B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) �t�|�`�B FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND  Substring(Convert(varchar(8), T1.[U_PC_BSAPP],112),1,4)  =@A1  ) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID  AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1 )");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END ");
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
            sb.Append(" SELECT  ���, [TYPE] ,�t�ӥN��,�t�ӦW��,���|���B,�|�B,�t�|�`�B FROM (SELECT '1' LINE, CASE WHEN ISNULL(P.[U_PC_BSAPP],'') = '' THEN Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),5,2) ELSE Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),5,2) END ���,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END [TYPE],P1.[CardCode] �t�ӥN��,P1.[CardName] �t�ӦW��,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  ���|���B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  �|�B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) �t�|�`�B FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry AND Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),1,4) =@A1  ) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1  )");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END,CASE WHEN ISNULL(P.[U_PC_BSAPP],'') = '' THEN Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),5,2) ELSE Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),5,2) END");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2' LINE,'TOTAL' ���,CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END [TYPE],P1.[CardCode] �t�ӥN��,P1.[CardName] �t�ӦW��,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMN ELSE T1.U_PC_BSAMN END,0) as int))  ���|���B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE T1.U_PC_BSTAX END,0) as int))  �|�B,");
            sb.Append(" SUM(cast(round(CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSAMT ELSE T1.U_PC_BSAMT END,0) as int)) �t�|�`�B FROM OJDT  J");
            sb.Append(" LEFT JOIN [dbo].[@CADMEN_FMD] T0 ON  J.TransID=T0.[U_BSREN]");
            sb.Append(" LEFT JOIN  [dbo].[@CADMEN_FMD1]  T1 ON (T0.DocEntry = T1.DocEntry  AND Substring(Convert(varchar(8),T1.[U_PC_BSAPP],112),1,4)  =@A1) ");
            sb.Append(" LEFT JOIN OPCH P ON (J.TRANSID=P.TRANSID  AND Substring(Convert(varchar(8),P.[U_PC_BSAPP],112),1,4)  =@A1)");
            sb.Append(" LEFT JOIN OPCH P1 ON (J.TRANSID=P1.TRANSID)");
            sb.Append(" WHERE  CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN P.U_PC_BSTAX ELSE  T1.U_PC_BSTAX END <> 0");
            sb.Append(" GROUP BY P1.[CardCode],P1.[CardName],CASE ISNULL(T1.U_PC_BSTAX,0) WHEN 0 THEN CASE P.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣'");
            sb.Append("  END ELSE CASE T1.U_PC_BSTY4 ");
            sb.Append(" WHEN 0  THEN '�i�f'");
            sb.Append(" WHEN 1  THEN '�O��'");
            sb.Append(" WHEN 2  THEN '�T�w�겣' END END ) AS A");
            sb.Append(" ORDER BY A.�t�ӥN��,[TYPE],���");


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
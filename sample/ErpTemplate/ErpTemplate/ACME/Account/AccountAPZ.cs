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
    public partial class AccountAPZ : Form
    {
        public AccountAPZ()
        {
            InitializeComponent();
        }

        private void AccountCard_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast(); 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private System.Data.DataTable GetOrderData6()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT 'AP' 單據總類,T0.[DocEntry] 單號, Convert(varchar(8),T1.[docdate],112)  過帳日期,T0.[ItemCode] 產品編號, T0.[Dscription] 產品名稱,CASE ISNULL(T2.DOCENTRY,'') WHEN '' THEN  T0.[AcctCode] ELSE T2.[AcctCode] END 科目代號,t1.cardcode 客戶編號,t1.cardname 客戶名稱, T0.[Price] 單價,T0.[LineTotal] 未稅總計,t0.vatsum 稅額,T0.[LineTotal]+t0.vatsum 總計,(T3.[lastName]+T3.[firstName]) 所有人 FROM pch1 T0");
            sb.Append("              left join opch t1 on t0.docentry=t1.docentry");
            sb.Append("           left join PDN1 t2 on t0.baseentry=T2.docentry and  t0.baseline=t2.linenum  and  t0.basetype='20' ");
            sb.Append("              left JOIN OHEM T3 ON T1.OwnerCode = T3.empID ");
            sb.Append("              where substring(t0.itemcode,1,1) = 'Z' and t0.itemcode not in ('ZB0840000.00001','ZBAR00000.00001') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T1.[docdate],112) between @DocDate1 and @DocDate2 ");
            }
            sb.Append(" union all ");
            sb.Append(" SELECT 'AP貸項' 單據總類,T0.[DocEntry] 單號, Convert(varchar(8),T1.[docdate],112)  過帳日期,T0.[ItemCode] 產品編號, T0.[Dscription] 產品名稱, T0.[AcctCode] 科目代號,t1.cardcode 客戶編號,t1.cardname 客戶名稱, T0.[Price]*-1 單價,T0.[LineTotal]*-1 未稅總計,t0.vatsum*-1 稅額,(T0.[LineTotal]+t0.vatsum)*-1 總計,(T3.[lastName]+T3.[firstName]) 所有人 FROM rpc1 T0");
            sb.Append(" left join orpc t1 on t0.docentry=t1.docentry");
            sb.Append(" left JOIN OHEM T3 ON T1.OwnerCode = T3.empID ");
            sb.Append(" where substring(t0.itemcode,1,1) = 'Z' and t0.itemcode not in ('ZB0840000.00001','ZBAR00000.00001') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {

                sb.Append(" and  Convert(varchar(8),T1.[docdate],112) between @DocDate1 and @DocDate2 ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Pdn1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }



        private void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetOrderData6();
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1.DataSource;

            label3.Text = dt.Compute("Sum(未稅總計)", null).ToString();
        }
    }
}
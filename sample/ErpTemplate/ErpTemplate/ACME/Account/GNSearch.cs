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
    public partial class GNSearch : Form
    {
        public GNSearch()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
    
            ViewBatchPayment2();
        }

        private void ViewBatchPayment2()
        {
            SqlConnection MyConnection;

                MyConnection = globals.shipConnection;

                StringBuilder sb = new StringBuilder();
                sb.Append(" SELECT  t0.transid 傳票單號,Convert(varchar(10),t0.refdate,112) 過帳日期 ");
                sb.Append(" ,t0.memo 種類,t0.baseref 原始號碼,case t0.transtype when '60' then cast(t3.price as int)*-1 else  cast(t4.price as int) end 收發貨金額,");
                sb.Append(" case t0.transtype when '60' then cast(T1.credit as int)*-1 else cast(T1.credit as int) end 傳票金額,");
                sb.Append(" isnull(cast(t2.CREDIT as int),0) 調整金額,");
                sb.Append(" case t0.transtype when '60' then cast(T3.price as int)-cast(T1.credit as int)+cast(isnull(t2.CREDIT,0) as int) else cast(T4.price as int)-cast(T1.credit as int)+cast(isnull(t2.CREDIT,0) as int) end 金額差異 ");
                sb.Append(" FROM ojdt T0 ");
                sb.Append(" left join (select transid,sum(credit) credit from jdt1 t0 group by transid) t1 on (t0.transid=t1.transid)");
                sb.Append(" left join jdt1 t2 on (t0.memo=SUBSTRING(T2.linememo,6,2) AND T0.BASEREF=SUBSTRING(T2.linememo,0,4) AND T2.LINE_ID='1') ");
                sb.Append(" left join (select BASE_REF docentry,sum(CAST(TRANSVALUE AS INT))*-1 price from oinm t0 where  t0.transtype in ('60') AND TRANSVALUE <> 0 group by BASE_REF ");
                sb.Append("   ) t3 on (t3.docentry=T0.BASEREF )");
                sb.Append(" left join (select BASE_REF docentry,sum(CAST(TRANSVALUE AS INT)) price from oinm t0 where  t0.transtype in ('59') AND TRANSVALUE <> 0 group by BASE_REF ");
                sb.Append("   ) t4 on (t4.docentry=T0.BASEREF )");
                sb.Append(" where t0.transtype in ('59','60') ");
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append(" and  Convert(varchar(8),t0.refdate,112) between @DocDate1 and @DocDate2 ");
                }
                sb.Append(" order by t0.transid ");
                SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
                command.CommandType = CommandType.Text;
                command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
                command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
                SqlDataAdapter da = new SqlDataAdapter(command);
              
                DataSet ds = new DataSet();
                try
                {
                    MyConnection.Open();
                    da.Fill(ds, "OPOR");
                }
                finally
                {
                    MyConnection.Close();
                }


                bindingSource1.DataSource = ds.Tables[0];
                dataGridView1.DataSource = bindingSource1;
            



        }


        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }



        private void GNSearch_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox2.Text = DateTime.Now.ToString("yyyyMMdd");
        }

    }
}
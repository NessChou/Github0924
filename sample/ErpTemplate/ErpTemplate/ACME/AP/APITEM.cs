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
    public partial class APITEM : Form
    {
        public APITEM()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GETOITM();

            dataGridView1.DataSource = G1;
        }

        private void APITEM_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            System.Data.DataTable G1 = GETOITM();

            dataGridView1.DataSource = G1;
        }
        private System.Data.DataTable GETOITM()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE 項目號碼,ITEMNAME 項目說明,U_PARTNO PARTNO,Convert(varchar(8),CreateDate,112)   新增日期   FROM OITM　 WHERE ");
            sb.Append(" U_PARTNO  NOT IN (SELECT MODEL COLLATE  Chinese_Taiwan_Stroke_CI_AS　FROM ACMESQLEEP.DBO.ACME_OITM WHERE DOCENTRY IN(SELECT　substring(form_presentation,11,DataLength(form_presentation)-11) FROM  ACMESQLEEP.DBO.SYS_TODOHIS 　WHERE FLOW_DESC='料號申請流程'　AND S_STEP_ID ='料號申請'))");
            sb.Append(" AND ITEMCODE  NOT IN (SELECT MODEL COLLATE  Chinese_Taiwan_Stroke_CI_AS　FROM ACMESQLEEP.DBO.ACME_OITM WHERE DOCENTRY IN(SELECT　substring(form_presentation,11,DataLength(form_presentation)-11) FROM  ACMESQLEEP.DBO.SYS_TODOHIS 　WHERE FLOW_DESC='料號申請流程'　AND S_STEP_ID ='料號申請'))");
            sb.Append(" AND ItmsGrpCod=1032　");
            sb.Append(" AND Convert(varchar(8),CreateDate,112)  between  @DocDate1 and @DocDate2  ORDER BY CreateDate");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

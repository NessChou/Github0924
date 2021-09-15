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
    public partial class CHIVOUCHER : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
  
        public CHIVOUCHER()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dtCost = MakeTableCombine();
            DataRow dr = null;
            System.Data.DataTable DT1 = null;
            if (comboBox1.Text == "韋峰")
            {
                DT1 = GetCHO3WE();
            }
            else
            {
                DT1 = GetCHO3();
            }

            string DuplicateKey = "";
            decimal BAL = 0;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                DataRow dd = DT1.Rows[i];
                string SUBJECT = dd["科目"].ToString();
                string SUBNAME = dd["科目名稱"].ToString();
                string COMPANY = dd["公司"].ToString();
                System.Data.DataTable G1 = GetCHO4(SUBJECT, COMPANY);
                if (DuplicateKey != SUBJECT + COMPANY)
                {


                    dr = dtCost.NewRow();
                    if (G1.Rows.Count > 0)
                    {
                        BAL = Convert.ToDecimal(G1.Rows[0][0]);
                    }
                    else
                    {
                        BAL = 0;
                    }
                    dr["公司"] = COMPANY;
                    dr["科目"] = SUBJECT;
                    dr["科目名稱"] = SUBNAME;
                    dr["傳票號碼"] = "期初餘額";
                    dr["餘額"] = BAL;
                    dtCost.Rows.Add(dr);
                }

                DuplicateKey = SUBJECT + COMPANY;
                dr = dtCost.NewRow();
                dr["公司"] = COMPANY;

                dr["科目"] = SUBJECT;
                dr["科目名稱"] = dd["科目名稱"].ToString();
                dr["傳票日期"] = dd["傳票日期"].ToString();
                dr["傳票號碼"] = dd["傳票號碼"].ToString();
                dr["部門"] = dd["部門"].ToString();
                dr["部門名稱"] = dd["部門名稱"].ToString();
                dr["明細"] = dd["明細"].ToString();
                decimal CD = Convert.ToDecimal(dd["C/D"].ToString());
                dr["C/D"] = dd["C/D"].ToString();
                dr["借"] = dd["借"].ToString();
                dr["貸"] = dd["貸"].ToString();
                BAL = BAL + CD;
                dr["餘額"] = BAL ;
                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;

            if (textBox7.Text != "")
            {
                dtCost.DefaultView.RowFilter = " 明細 LIKE   '%" + textBox7.Text.ToString() + "%' ";
            }

            for (int i = 8; i <= 11; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("公司", typeof(string));
            dt.Columns.Add("科目", typeof(string));
            dt.Columns.Add("科目名稱", typeof(string));
            dt.Columns.Add("傳票日期", typeof(string));
            dt.Columns.Add("傳票號碼", typeof(string));
            dt.Columns.Add("部門", typeof(string));
            dt.Columns.Add("部門名稱", typeof(string));
            dt.Columns.Add("明細", typeof(string));
            dt.Columns.Add("C/D", typeof(decimal));
            dt.Columns.Add("借", typeof(decimal));
            dt.Columns.Add("貸", typeof(decimal));
            dt.Columns.Add("餘額", typeof(decimal));
            return dt;
        }

        public System.Data.DataTable GetCHO3()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select '聿豐' 公司,A.SubjectID 科目,C.SubjectName 科目名稱,Convert(varchar(10),CAST(CAST(MakeDate AS VARCHAR) AS DATETIME),111) 傳票日期, ''''+A.VoucherNo 傳票號碼,A.DepartID 部門,B.DepartName 部門名稱,A.SUMMARY 明細, ");
            sb.Append(" CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) -ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0)  AS INT) 'C/D', ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) AS INT) AS INT) 借, ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0) AS INT) AS INT) 貸 ");
            sb.Append(" From CHICOMP02.DBO.AccVoucherSub A   ");
            sb.Append(" Left Join CHICOMP02.DBO.comDepartment B On B.DepartID=A.DepartID   ");
            sb.Append(" Left Join CHICOMP02.DBO.ComSubject C On C.SubjectID=A.SubjectID   ");
            sb.Append(" Left Join CHICOMP02.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo    ");
            sb.Append(" WHERE  MakeDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  B.DepartID in ( " + c + ") ");
            }
            else
            {
                sb.Append("  And B.DepartID between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  A.SubjectID in ( " + d + ") ");
            }
            else
            {
                sb.Append("  And A.SubjectID between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'");
            }
            if (comboBox1.Text != "")
            {
                if (comboBox1.Text != "聿豐")
                {

                    sb.Append(" and  1 =2 ");
                }
            }
            sb.Append(" UNION ALL ");
            sb.Append(" Select  '東門' 公司,A.SubjectID 科目,C.SubjectName 科目名稱,Convert(varchar(10),CAST(CAST(MakeDate AS VARCHAR) AS DATETIME),111) 傳票日期, ''''+A.VoucherNo 傳票號碼,A.DepartID 部門,B.DepartName,A.SUMMARY 明細, ");
            sb.Append(" CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) -ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0)  AS INT) 'C/D', ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) AS INT) AS INT) 借, ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0) AS INT) AS INT) 貸 ");
            sb.Append(" From CHICOMP03.DBO.AccVoucherSub A   ");
            sb.Append(" Left Join CHICOMP03.DBO.comDepartment B On B.DepartID=A.DepartID   ");
            sb.Append(" Left Join CHICOMP03.DBO.ComSubject C On C.SubjectID=A.SubjectID   ");
            sb.Append(" Left Join CHICOMP03.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo    ");
            sb.Append(" WHERE  MakeDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
            if (comboBox1.Text != "")
            {
                if (comboBox1.Text != "東門")
                {

                    sb.Append(" and  1 = 2 ");
                }
            }

            if (checkBox1.Checked)
            {
                sb.Append(" and  B.DepartID in ( " + c + ") ");
            }
            else
            {
                sb.Append("  And B.DepartID between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  A.SubjectID in ( " + d + ") ");
            }
            else
            {
                sb.Append("  And A.SubjectID between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'");
            }
            sb.Append(" ORDER BY 公司,A.SubjectID,傳票號碼  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public System.Data.DataTable GetCHO3WE()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select '韋峰' 公司,A.SubjectID 科目,C.SubjectName 科目名稱,Convert(varchar(10),CAST(CAST(MakeDate AS VARCHAR) AS DATETIME),111) 傳票日期, ''''+A.VoucherNo 傳票號碼,A.DepartID 部門,B.DepartName 部門名稱,A.SUMMARY 明細, ");
            sb.Append(" CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) -ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0)  AS INT) 'C/D', ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =1 THEN AMOUNT END,0) AS INT) AS INT) 借, ");
            sb.Append(" CAST(CAST(ISNULL(CASE WHEN DebitCredit =0 THEN AMOUNT END,0) AS INT) AS INT) 貸 ");
            sb.Append(" From CHICOMP17.DBO.AccVoucherSub A   ");
            sb.Append(" Left Join CHICOMP17.DBO.comDepartment B On B.DepartID=A.DepartID   ");
            sb.Append(" Left Join CHICOMP17.DBO.ComSubject C On C.SubjectID=A.SubjectID   ");
            sb.Append(" Left Join CHICOMP17.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo    ");
            sb.Append(" WHERE  MakeDate between '" + textBox5.Text.ToString() + "' and '" + textBox6.Text.ToString() + "' ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  B.DepartID in ( " + c + ") ");
            }
            else
            {
                sb.Append("  And B.DepartID between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            if (checkBox2.Checked)
            {
                sb.Append(" and  A.SubjectID in ( " + d + ") ");
            }
            else
            {
                sb.Append("  And A.SubjectID between '" + textBox3.Text.ToString() + "' and '" + textBox4.Text.ToString() + "'");
            }
           
            sb.Append(" ORDER BY A.SubjectID,傳票號碼  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }


        public System.Data.DataTable GetCHO4(string SubjectID,string COMPANY)
        {
            string ST = "";
            if (COMPANY == "聿豐")
            {
                ST = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (COMPANY == "東門")
            {
                ST = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (COMPANY == "韋峰")
            {
                ST = "Data Source=10.10.1.40;Initial Catalog=CHICOMP17;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection MyConnection = new SqlConnection(ST);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select IsNull(sum(B.Amount*(2*B.DebitCredit-1)),0) as Amount From AccVoucherMain A");
            sb.Append("  Join AccVoucherSub B on (A.VoucherNO = B.VoucherNO) ");
            sb.Append("  Where B.SubjectID =@SubjectID and A.IsTransfer <> 2  ");
            if (checkBox1.Checked)
            {
                sb.Append(" and  B.DepartID in ( " + c + ") ");
            }
            else
            {
                sb.Append("  And B.DepartID between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'");
            }
            sb.Append("  And A.MakeDate<'" + textBox5.Text.ToString() + "' Group By B.SubjectID order by B.SubjectID");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SubjectID", SubjectID));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        private void CHIVOUCHER_Load(object sender, EventArgs e)
        {
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.Day();
        }



        private void button3_Click_1(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        public string c;
        public string d;
        private void button8_Click(object sender, EventArgs e)
        {
            CHIVOUCHER2 frm1 = new CHIVOUCHER2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox1.Checked = true;
                c = frm1.q;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CHIVOUCHER1 frm1 = new CHIVOUCHER1();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                checkBox2.Checked = true;
                d = frm1.q;

            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {

            if (e.RowIndex >= dataGridView1.Rows.Count)
                return;
            DataGridViewRow dgr = dataGridView1.Rows[e.RowIndex];
            try
            {
                string VOU = dgr.Cells["傳票號碼"].Value.ToString();
                if (!String.IsNullOrEmpty(VOU))
                {

                    if (VOU == "期初餘額")
                    {
                        dgr.DefaultCellStyle.BackColor = Color.Pink;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
   
    }
}

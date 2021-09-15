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
    public partial class Statistics : Form
    {
        public Statistics()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = null;
            if (comboBox1.SelectedValue.ToString() == "1")
            {
                 dt = GetSAPRevenue1();


            }
            else if (comboBox1.SelectedValue.ToString() == "2")
            {
                dt = GetSAPRevenue2();
            }
            else if (comboBox1.SelectedValue.ToString() == "3")
            {

                dt = GetSAPRevenue3();
            }
            else if (comboBox1.SelectedValue.ToString() == "4")
            {
                dt = GetSAPRevenue4();
            }
            else if (comboBox1.SelectedValue.ToString() == "5")
            {
                dt = GetSAPRevenue5();
            }
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1.DataSource;
        }

        private void Statistics_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst(); 
            textBox2.Text = GetMenu.DLast(); 
            UtilSimple.SetLookupBinding(comboBox1, GetBU("StockPapare"), "DataText", "DataValue");
        }
        System.Data.DataTable GetBU(string KIND)
        {
            SqlConnection con = globals.CommonConnection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by DataValue";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application wapp;

            Microsoft.Office.Interop.Excel.Worksheet wsheet;

            Microsoft.Office.Interop.Excel.Workbook wbook;

            wapp = new Microsoft.Office.Interop.Excel.Application();

            wapp.Visible = false;

            wbook = wapp.Workbooks.Add(true);

            wsheet = (Worksheet)wbook.ActiveSheet;
            try
            {

                int iX;
                int iY;



                for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
                {


                    wsheet.Cells[1, i + 1] = this.dataGridView1.Columns[i].HeaderText;



                }



                for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                {



                    DataGridViewRow row = this.dataGridView1.Rows[i];



                    for (int j = 0; j < row.Cells.Count; j++)
                    {



                        DataGridViewCell cell = row.Cells[j];



                        try
                        {



                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : "'" + cell.Value.ToString();



                        }



                        catch (Exception ex)
                        {



                            MessageBox.Show(ex.Message);



                        }

                    }

                }

                wapp.Visible = true;


            }

            catch (Exception ex1)
            {



                MessageBox.Show(ex1.Message);



            }

            wapp.UserControl = true;
           

        }
        private System.Data.DataTable GetSAPRevenue1()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'SAP' 總類,COUNT(*) 筆數,");
            sb.Append(" substring(T0.employee,0,4) 製單人員 FROM Stock_Prepare T0");
            sb.Append(" where (T0.employee like '%碧珠%' or T0.employee like '%筱倩%'");
            sb.Append("  or T0.employee like '%淑芳%'  or T0.employee like '%思怡%'");
            sb.Append("  or T0.employee like '%雅卉%')");
            sb.Append("  and Convert(varchar(8),t0.docdate,112) between @aa AND @bb ");
            sb.Append(" GROUP BY substring(T0.employee,0,4)");





            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenue2()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT COUNT(*) 筆數,T0.U_ACME_USER 製單人 FROM ODLN T0");
            //sb.Append(" where (T0.U_ACME_USER like '%sunny%' or T0.U_ACME_USER like '%demi%'");
            //sb.Append("  or T0.U_ACME_USER like '%daly%'  or T0.U_ACME_USER like '%joy%'");
            //sb.Append("  or T0.U_ACME_USER like '%maggie%')");
            //sb.Append("  and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            //sb.Append(" GROUP BY T0.U_ACME_USER");
            sb.Append(" SELECT S.總類,SUM(S.筆數) 筆數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 筆數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ODLN T0");
            sb.Append("           where ((T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 筆數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ORDN T0");
            sb.Append("              where ((T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (5,6)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (5,6)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員 )");
            sb.Append(" S GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenue3()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            //sb.Append(" SELECT COUNT(*) 筆數,case u_acme_kind when 1 then ");
            //sb.Append(" '借出' when '2' then '借出還回' when '3' then '調撥' end 總類");
            //sb.Append(" ,T0.U_ACME_USER 製單人 FROM Owtr T0");
            //sb.Append(" where (T0.U_ACME_USER like '%sunny%' or T0.U_ACME_USER like '%demi%'");
            //sb.Append("  or T0.U_ACME_USER like '%daly%'  or T0.U_ACME_USER like '%joy%'");
            //sb.Append("  or T0.U_ACME_USER like '%maggie%')");
            //sb.Append("  and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            //sb.Append(" GROUP BY T0.U_ACME_USER,u_acme_kind");
            sb.Append(" SELECT S.總類,SUM(S.筆數) 筆數,S.製單人員 FROM   (      ");
            sb.Append("     SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 筆數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM Opdn T0");
            sb.Append("           where ((T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 筆數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM ORPD T0");
            sb.Append("              where ((T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT CAST('SAP' AS NVARCHAR) 總類,COUNT(*) 筆數,substring(T0.U_ACME_USER,0,4) COLLATE Chinese_Taiwan_Stroke_CI_AS 製單人員 FROM OPOR T0");
            sb.Append("              where ((T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("              and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4),u_acme_kind");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (1,2)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.ORDOBillMain");
            sb.Append(" WHERE 旗標 IN (3)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.ORDOBillMain");
            sb.Append(" WHERE 旗標 IN (3)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (1,2)");
            sb.Append("              and 日期 between @aa AND @bb ");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append(" GROUP BY 製單人員 )");
            sb.Append(" S GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenue4()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT S.總類,SUM(筆數) 筆數,S.製單人員 FROM    ( SELECT 'SAP' 總類,COUNT(*) 筆數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM OWTR T0");
            sb.Append("                           where ((T0.ref2 like '%碧珠%' or T0.ref2 like '%筱倩%'");
            sb.Append("                            or T0.ref2 like '%淑芳%'  or T0.ref2 like '%思怡%'");
            sb.Append("                            or T0.ref2 like '%雅卉%') OR (T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("                            or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("                            or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END  ");
            sb.Append("            UNION ALL       ");
            sb.Append("             SELECT 'SAP' 總類,COUNT(*) 筆數");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人員 FROM oigN T0");
            sb.Append("                           where ((T0.ref2 like '%碧珠%' or T0.ref2 like '%筱倩%'");
            sb.Append("                            or T0.ref2 like '%淑芳%'  or T0.ref2 like '%思怡%'");
            sb.Append("                            or T0.ref2 like '%雅卉%') OR (T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("                            or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("                            or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("                            and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END,u_acme_kind");
            sb.Append("                           union all");
            sb.Append("                       SELECT 'SAP' 總類, COUNT(*)");
            sb.Append("                           ,CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4)  END 製單人 FROM oigE T0");
            sb.Append("                                                  where ((T0.ref2 like '%碧珠%' or T0.ref2 like '%筱倩%'");
            sb.Append("                            or T0.ref2 like '%淑芳%'  or T0.ref2 like '%思怡%'");
            sb.Append("                            or T0.ref2 like '%雅卉%') OR (T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("                            or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("                            or T0.U_ACME_USER like '%雅卉%'))");
            sb.Append("                           and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("                           GROUP BY CASE ISNULL(T0.ref2,0) WHEN '0' THEN substring(T0.U_ACME_USER,0,4) WHEN NULL THEN substring(T0.U_ACME_USER,0,4) ELSE  substring(T0.ref2,0,4) END,u_acme_kind");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (3,4)");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                            and 日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.stkBillMain");
            sb.Append(" WHERE 旗標 IN (3,4)");
            sb.Append(" AND   (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                            and 日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.stkBORROWMain");
            sb.Append(" WHERE  (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                            and 日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL20.DBO.stkRETURNMain");
            sb.Append(" WHERE  (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                           AND    日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.stkBORROWMain");
            sb.Append(" WHERE  (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                      AND     日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT CAST('正航' AS NVARCHAR) 總類,COUNT(*) 筆數,製單人員 COLLATE Chinese_Taiwan_Stroke_CI_AS FROM OPENDATASOURCE('SQLOLEDB',");
            sb.Append("      'Data Source=acmesrvCHI;User ID=CHI;Password=CHI'");
            sb.Append("       ).SUNSQL21.DBO.stkRETURNMain");
            sb.Append(" WHERE  (製單人員  like '%碧珠%' or 製單人員  like '%筱倩%'");
            sb.Append("               or 製單人員  like '%淑芳%'  or 製單人員  like '%思怡%'");
            sb.Append("               or 製單人員  like '%雅卉%')");
            sb.Append("                    AND      日期 between @aa AND @bb ");
            sb.Append(" GROUP BY 製單人員 ) S");
            sb.Append(" GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSAPRevenue5()
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT S.總類,SUM(S.筆數) 筆數,S.製單人員 FROM (  SELECT 'SAP' 總類,COUNT(*) 筆數");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM Oinv T0");
            sb.Append("              where (T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%')");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4)");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT 'SAP' 總類,COUNT(*) 筆數 ");
            sb.Append("              ,substring(T0.U_ACME_USER,0,4) 製單人員 FROM ORIN T0");
            sb.Append("              where (T0.U_ACME_USER like '%碧珠%' or T0.U_ACME_USER like '%筱倩%'");
            sb.Append("               or T0.U_ACME_USER like '%淑芳%'  or T0.U_ACME_USER like '%思怡%'");
            sb.Append("               or T0.U_ACME_USER like '%雅卉%')");
            sb.Append("               and Convert(varchar(8),t0.taxdate,112) between @aa AND @bb ");
            sb.Append("              GROUP BY substring(T0.U_ACME_USER,0,4) ) S");
            sb.Append(" GROUP BY S.總類,S.製單人員");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }

    
    }
}
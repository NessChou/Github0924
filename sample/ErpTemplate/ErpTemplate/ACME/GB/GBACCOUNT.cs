using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using ACME.Service;
namespace ACME
{
    public partial class GBACCOUNT : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        private System.Data.DataTable dtDept, dtDept2, dtDept3;
        public GBACCOUNT()
        {
            InitializeComponent();
        }

        private void GBACCOUNT_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox2.Text = GetMenu.Day();

            ArrayList al2 = new ArrayList();
            StringBuilder sb2 = new StringBuilder();

            dtDept = GetDept();
            DataRow dr2;
            for (int i = 0; i <= dtDept.Rows.Count - 1; i++)
            {
                dr2 = dtDept.Rows[i];

                string BU = Convert.ToString(dr2["PARAM_NO"]);

                listBox1.Items.Add(BU);
            }

            comboBox1.Text ="試算表";
        }

        private void execuse()
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\GBACCOUNT.xls";

                ArrayList al = new ArrayList();
                ArrayList al2 = new ArrayList();
                StringBuilder sb = new StringBuilder();
                string ff = "";
                if (listBox3.SelectedItems.Count > 0)
                {
                    for (int i = 0; i <= listBox3.SelectedItems.Count - 1; i++)
                    {
                        al.Add(listBox3.SelectedItems[i].ToString());

                        string fd = listBox3.SelectedItems[i].ToString();
                        ff = listBox3.SelectedItems[0].ToString();
                    }

                    foreach (string v in al)
                    {
                        sb.Append("'" + v + "',");
                    }

                    sb.Remove(sb.Length - 1, 1);


                }

                StringBuilder sb2 = new StringBuilder();
                StringBuilder sb3 = new StringBuilder();
                StringBuilder sb4 = new StringBuilder();
                if (listBox1.SelectedItems.Count > 0)
                {

                    for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                    {
                        al2.Add(listBox1.SelectedItems[i].ToString());

                    }

                    foreach (string v in al2)
                    {
                        sb2.Append("'" + v + "',");
                    }
                    sb2.Remove(sb2.Length - 1, 1);
                }

                string J1 = sb2.ToString();
                if (J1.IndexOf("TFT") != -1)
                {
                    System.Data.DataTable k1 = GetBU("TFT");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                if (J1.IndexOf("禾豐牧場") != -1)
                {
                    System.Data.DataTable k1 = GetBU("禾豐牧場");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                if (J1.IndexOf("利豐漁場") != -1)
                {
                    System.Data.DataTable k1 = GetBU("利豐漁場");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                if (J1.IndexOf("董事室") != -1)
                {
                    System.Data.DataTable k1 = GetBU("董事室");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                if (J1.IndexOf("品牌行銷") != -1)
                {
                    System.Data.DataTable k1 = GetBU("品牌行銷");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }
                if (J1.IndexOf("生物科技") != -1)
                {
                    System.Data.DataTable k1 = GetBU("生物科技");
                    for (int i = 0; i <= k1.Rows.Count - 1; i++)
                    {

                        DataRow dd = k1.Rows[i];

                        string F = dd["DEPT"].ToString();
                        sb3.Append("'" + F + "',");

                    }

                }

                System.Data.DataTable OrderData = null;
                string SBF = "";
                if (listBox3.SelectedItems.Count > 0)
                {
                    SBF = sb.ToString();
   
                }
                else
                {
                    if (sb3.Length > 0)
                    {
                        sb3.Remove(sb3.Length - 1, 1);
                    }
                    SBF = sb3.ToString();
                }

                if (listBox2.Text == "ALL" || listBox2.Text == "")
                {
                    OrderData = GetSAPCostByProductALL(SBF);
                }
                else
                {
                    OrderData = GetSAPCostByProduct(SBF);
                }
                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                if (OrderData.Rows.Count > 0)
                {
                    WebService1 proxy = new WebService1();
         //           proxy.Url = "http://wf.acmepoint.net/rma/WebService/AcmeService.asmx";
      //    proxy.
                    ExcelReport.ExcelReportOutput2(OrderData, ExcelTemplate, OutPutFile, "N");
                }
                else
                {
                    MessageBox.Show("沒有資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GetDept()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PARAM_NO FROM PARAMS WHERE PARAM_KIND='GBCHO' ORDER BY PARAM_NO ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRC");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables["OPRC"];



        }

        private void button1_Click(object sender, EventArgs e)
        {
            execuse();
        }

        private System.Data.DataTable GetSAPCostByProduct(string ACC)
        {
            string strCnF = "";

            if (listBox2.Text  == "創田")
            {
                strCnF = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (listBox2.Text == "聿豐")
            {
                strCnF = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (listBox2.Text == "東門")
            {
                strCnF = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            if (listBox2.Text == "禾豐")
            {
                strCnF = "Data Source=10.10.1.40;Initial Catalog=CHICOMP06;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            }
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 年月,科目,科目名稱,DepartID,COMPANY,LINE,SUM(AMT) AMT FROM (");

                sb.Append(" Select SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6) 年月,SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8) 日,A.SubjectID 科目,C.SubjectName 科目名稱,A.DepartID,B.DEPARTNAME, SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
                sb.Append(" ,CASE WHEN A.DepartID IN ('D10','D1','D11','S3','S03') THEN 'A 禾豐牧場' ");
                sb.Append(" WHEN A.DepartID IN ('S01','S02','S04','S05','S1','S2','S4') THEN 'B 利豐漁場' ");
                sb.Append(" WHEN A.DepartID IN ('AB1','D99') THEN 'C 董事室' ");
                sb.Append(" WHEN A.DepartID IN ('C1','C2','F2') THEN 'D 品牌行銷' ");
                sb.Append(" WHEN A.DepartID IN ('F1') THEN 'E 生物科技' WHEN  A.DepartID IN ('A10') THEN 'F TFT' END COMPANY");
                sb.Append(" ,CASE A.DepartID WHEN 'D10' THEN '禾豐牧場-豬' WHEN 'D1' THEN '禾豐牧場-豬' WHEN 'D11' THEN '禾豐牧場-雞'");
                sb.Append(" WHEN 'S01' THEN '利豐-崗山' WHEN 'S02' THEN '利豐-車城' WHEN 'S03' THEN '禾豐-麟洛' WHEN 'S04' THEN '利豐-枋山' WHEN 'S05' THEN '利豐-新街'");
                sb.Append(" WHEN 'S1' THEN '利豐-崗山' WHEN 'S2' THEN '利豐-車城' WHEN 'S3' THEN '禾豐-麟洛' WHEN 'S4' THEN '利豐-枋山' ");
                sb.Append(" WHEN 'AB1' THEN '董事室' WHEN 'D99' THEN '管報調整'  WHEN 'C1' THEN '品牌-行銷' WHEN 'C2' THEN '品牌-行銷-東門店' WHEN 'F2' THEN '品牌-行銷' WHEN 'F1' THEN '生物科技' WHEN 'A10' THEN 'TFT'  END LINE");
                sb.Append(" From DBO.AccVoucherSub A ");
                sb.Append(" Left Join DBO.comDepartment B On B.DepartID=A.DepartID ");
                sb.Append(" Left Join DBO.ComSubject C On C.SubjectID=A.SubjectID");
                sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
                sb.Append(" WHERE  1=1 ");
                if (!String.IsNullOrEmpty(ACC))
                {
                    sb.Append(" AND  A.DepartID IN (" + ACC + " )");
                }
                if(comboBox1.Text =="損益表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 4 AND 9 ");
                }

                if (comboBox1.Text == "資產負債表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 1 AND 3 ");
                }

                sb.Append(" GROUP BY  SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6),SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8), A.DepartID,B.DEPARTNAME,A.SubjectID,C.SubjectName");
          
            sb.Append("  ) AS A  ");
            sb.Append(" WHERE AMT <> 0 AND  日 between @DocDate1 and @DocDate2 ");
            sb.Append(" GROUP BY 年月,DepartID,科目,科目名稱,LINE,COMPANY ORDER BY COMPANY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            System.Data.DataTable dt = ds.Tables[0];

            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["Model"];
            colPk[1] = dt.Columns["月份"];
            dt.PrimaryKey = colPk;


            return dt;


        }

        private System.Data.DataTable GetSAPCostByProductALL(string ACC)
        {
            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 年月,科目,科目名稱,DepartID,COMPANY,LINE,SUM(AMT) AMT FROM (");

   
                sb.Append(" Select SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6) 年月,SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8) 日,A.SubjectID 科目,C.SubjectName 科目名稱,A.DepartID,B.DEPARTNAME, SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
                sb.Append(" ,CASE WHEN A.DepartID IN ('D10','D1','D11','S3','S03') THEN 'A 禾豐牧場' ");
                sb.Append(" WHEN A.DepartID IN ('S01','S02','S04','S05','S1','S2','S4') THEN 'B 利豐漁場' ");
                sb.Append(" WHEN A.DepartID IN ('AB1','D99') THEN 'C 董事室' ");
                sb.Append(" WHEN A.DepartID IN ('C1','C2','F2') THEN 'D 品牌行銷' ");
                sb.Append(" WHEN A.DepartID IN ('F1') THEN 'E 生物科技' WHEN  A.DepartID IN ('A10') THEN 'F TFT' END COMPANY");
                sb.Append(" ,CASE A.DepartID WHEN 'D10' THEN '禾豐牧場-豬' WHEN 'D1' THEN '禾豐牧場-豬' WHEN 'D11' THEN '禾豐牧場-雞'");
                sb.Append(" WHEN 'S01' THEN '利豐-崗山' WHEN 'S02' THEN '利豐-車城' WHEN 'S03' THEN '禾豐-麟洛' WHEN 'S04' THEN '利豐-枋山' WHEN 'S05' THEN '利豐-新街'");
                sb.Append(" WHEN 'S1' THEN '利豐-崗山' WHEN 'S2' THEN '利豐-車城' WHEN 'S3' THEN '禾豐-麟洛' WHEN 'S4' THEN '利豐-枋山' ");
                sb.Append(" WHEN 'AB1' THEN '董事室' WHEN 'D99' THEN '管報調整'  WHEN 'C1' THEN '品牌-行銷' WHEN 'C2' THEN '品牌-行銷-東門店' WHEN 'F2' THEN '品牌-行銷' WHEN 'F1' THEN '生物科技' WHEN 'A10' THEN 'TFT'  END LINE");
                sb.Append(" From CHIComp21.DBO.AccVoucherSub A ");
                sb.Append(" Left Join CHIComp21.DBO.comDepartment B On B.DepartID=A.DepartID ");
                sb.Append(" Left Join CHIComp21.DBO.ComSubject C On C.SubjectID=A.SubjectID");
                sb.Append(" Left Join CHIComp21.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
                sb.Append(" WHERE  1=1 ");
                if (!String.IsNullOrEmpty(ACC))
                {
                    sb.Append(" AND  A.DepartID IN (" + ACC + " )");
                }
                if (comboBox1.Text == "損益表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 4 AND 9 ");
                }

                if (comboBox1.Text == "資產負債表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 1 AND 3 ");
                }

                sb.Append(" GROUP BY  SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6),SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8), A.DepartID,B.DEPARTNAME,A.SubjectID,C.SubjectName");
            
            sb.Append(" UNION ALL");
           
                sb.Append(" Select SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6) 年月,SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8) 日,A.SubjectID 科目,C.SubjectName 科目名稱,A.DepartID,B.DEPARTNAME, SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
                sb.Append(" ,CASE WHEN A.DepartID IN ('D10','D1','D11','S3','S03') THEN 'A 禾豐牧場' ");
                sb.Append(" WHEN A.DepartID IN ('S01','S02','S04','S05','S1','S2','S4') THEN 'B 利豐漁場' ");
                sb.Append(" WHEN A.DepartID IN ('AB1','D99') THEN 'C 董事室' ");
                sb.Append(" WHEN A.DepartID IN ('C1','C2','F2') THEN 'D 品牌行銷' ");
                sb.Append(" WHEN A.DepartID IN ('F1') THEN 'E 生物科技' WHEN  A.DepartID IN ('A10') THEN 'F TFT' END COMPANY");
                sb.Append(" ,CASE A.DepartID WHEN 'D10' THEN '禾豐牧場-豬' WHEN 'D1' THEN '禾豐牧場-豬' WHEN 'D11' THEN '禾豐牧場-雞'");
                sb.Append(" WHEN 'S01' THEN '利豐-崗山' WHEN 'S02' THEN '利豐-車城' WHEN 'S03' THEN '禾豐-麟洛' WHEN 'S04' THEN '利豐-枋山' WHEN 'S05' THEN '利豐-新街'");
                sb.Append(" WHEN 'S1' THEN '利豐-崗山' WHEN 'S2' THEN '利豐-車城' WHEN 'S3' THEN '禾豐-麟洛' WHEN 'S4' THEN '利豐-枋山' ");
                sb.Append(" WHEN 'AB1' THEN '董事室'  WHEN 'D99' THEN '管報調整' WHEN 'C1' THEN '品牌-行銷' WHEN 'C2' THEN '品牌-行銷-東門店' WHEN 'F2' THEN '品牌-行銷' WHEN 'F1' THEN '生物科技' WHEN 'A10' THEN 'TFT'  END LINE");
                sb.Append(" From CHIComp02.DBO.AccVoucherSub A ");
                sb.Append(" Left Join CHIComp02.DBO.comDepartment B On B.DepartID=A.DepartID ");
                sb.Append(" Left Join CHIComp02.DBO.ComSubject C On C.SubjectID=A.SubjectID");
                sb.Append(" Left Join CHIComp02.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
                sb.Append(" WHERE  1=1 ");
                if (!String.IsNullOrEmpty(ACC))
                {
                    sb.Append(" AND  A.DepartID IN (" + ACC + " )");
                }
                if (comboBox1.Text == "損益表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 4 AND 9 ");
                }

                if (comboBox1.Text == "資產負債表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 1 AND 3 ");
                }
                sb.Append(" GROUP BY  SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6),SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8), A.DepartID,B.DEPARTNAME,A.SubjectID,C.SubjectName  ");


                sb.Append(" UNION ALL");

                sb.Append(" Select SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6) 年月,SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8) 日,A.SubjectID 科目,C.SubjectName 科目名稱,A.DepartID,B.DEPARTNAME, SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
                sb.Append(" ,CASE WHEN A.DepartID IN ('D10','D1','D11','S3','S03') THEN 'A 禾豐牧場' ");
                sb.Append(" WHEN A.DepartID IN ('S01','S02','S04','S05','S1','S2','S4') THEN 'B 利豐漁場' ");
                sb.Append(" WHEN A.DepartID IN ('AB1','D99') THEN 'C 董事室' ");
                sb.Append(" WHEN A.DepartID IN ('C1','C2','F2') THEN 'D 品牌行銷' ");
                sb.Append(" WHEN A.DepartID IN ('F1') THEN 'E 生物科技' WHEN  A.DepartID IN ('A10') THEN 'F TFT' END COMPANY");
                sb.Append(" ,CASE A.DepartID WHEN 'D10' THEN '禾豐牧場-豬' WHEN 'D1' THEN '禾豐牧場-豬' WHEN 'D11' THEN '禾豐牧場-雞'");
                sb.Append(" WHEN 'S01' THEN '利豐-崗山' WHEN 'S02' THEN '利豐-車城' WHEN 'S03' THEN '禾豐-麟洛' WHEN 'S04' THEN '利豐-枋山' WHEN 'S05' THEN '利豐-新街'");
                sb.Append(" WHEN 'S1' THEN '利豐-崗山' WHEN 'S2' THEN '利豐-車城' WHEN 'S3' THEN '禾豐-麟洛' WHEN 'S4' THEN '利豐-枋山' ");
                sb.Append(" WHEN 'AB1' THEN '董事室'  WHEN 'D99' THEN '管報調整' WHEN 'C1' THEN '品牌-行銷' WHEN 'C2' THEN '品牌-行銷-東門店' WHEN 'F2' THEN '品牌-行銷' WHEN 'F1' THEN '生物科技' WHEN 'A10' THEN 'TFT'  END LINE");
                sb.Append(" From CHIComp03.DBO.AccVoucherSub A ");
                sb.Append(" Left Join CHIComp03.DBO.comDepartment B On B.DepartID=A.DepartID ");
                sb.Append(" Left Join CHIComp03.DBO.ComSubject C On C.SubjectID=A.SubjectID");
                sb.Append(" Left Join CHIComp03.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
                sb.Append(" WHERE  1=1 ");
                if (!String.IsNullOrEmpty(ACC))
                {
                    sb.Append(" AND  A.DepartID IN (" + ACC + " )");
                }
                if (comboBox1.Text == "損益表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 4 AND 9 ");
                }

                if (comboBox1.Text == "資產負債表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 1 AND 3 ");
                }
                sb.Append(" GROUP BY  SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6),SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8), A.DepartID,B.DEPARTNAME,A.SubjectID,C.SubjectName  ");

                sb.Append(" UNION ALL");

                sb.Append(" Select SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6) 年月,SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8) 日,A.SubjectID 科目,C.SubjectName 科目名稱,A.DepartID,B.DEPARTNAME, SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
                sb.Append(" ,CASE WHEN A.DepartID IN ('D10','D1','D11','S3','S03') THEN 'A 禾豐牧場' ");
                sb.Append(" WHEN A.DepartID IN ('S01','S02','S04','S05','S1','S2','S4') THEN 'B 利豐漁場' ");
                sb.Append(" WHEN A.DepartID IN ('AB1','D99') THEN 'C 董事室' ");
                sb.Append(" WHEN A.DepartID IN ('C1','C2','F2') THEN 'D 品牌行銷' ");
                sb.Append(" WHEN A.DepartID IN ('F1') THEN 'E 生物科技' WHEN  A.DepartID IN ('A10') THEN 'F TFT' END COMPANY");
                sb.Append(" ,CASE A.DepartID WHEN 'D10' THEN '禾豐牧場-豬' WHEN 'D1' THEN '禾豐牧場-豬' WHEN 'D11' THEN '禾豐牧場-雞'");
                sb.Append(" WHEN 'S01' THEN '利豐-崗山' WHEN 'S02' THEN '利豐-車城' WHEN 'S03' THEN '禾豐-麟洛' WHEN 'S04' THEN '利豐-枋山' WHEN 'S05' THEN '利豐-新街'");
                sb.Append(" WHEN 'S1' THEN '利豐-崗山' WHEN 'S2' THEN '利豐-車城' WHEN 'S3' THEN '禾豐-麟洛' WHEN 'S4' THEN '利豐-枋山' ");
                sb.Append(" WHEN 'AB1' THEN '董事室'  WHEN 'D99' THEN '管報調整' WHEN 'C1' THEN '品牌-行銷' WHEN 'C2' THEN '品牌-行銷-東門店' WHEN 'F2' THEN '品牌-行銷' WHEN 'F1' THEN '生物科技' WHEN 'A10' THEN 'TFT'  END LINE");
                sb.Append(" From CHIComp06.DBO.AccVoucherSub A ");
                sb.Append(" Left Join CHIComp06.DBO.comDepartment B On B.DepartID=A.DepartID ");
                sb.Append(" Left Join CHIComp06.DBO.ComSubject C On C.SubjectID=A.SubjectID");
                sb.Append(" Left Join CHIComp06.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
                sb.Append(" WHERE  1=1 ");
                if (!String.IsNullOrEmpty(ACC))
                {
                    sb.Append(" AND  A.DepartID IN (" + ACC + " )");
                }
                if (comboBox1.Text == "損益表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 4 AND 9 ");
                }

                if (comboBox1.Text == "資產負債表")
                {
                    sb.Append(" AND SUBSTRING(A.SubjectID,1,1) BETWEEN 1 AND 3 ");
                }
                sb.Append(" GROUP BY  SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,6),SUBSTRING(cast(T0.MakeDate AS VARCHAR),1,8), A.DepartID,B.DEPARTNAME,A.SubjectID,C.SubjectName  ");

            
            sb.Append("  ) AS A  ");
            sb.Append(" WHERE AMT <> 0 AND  日 between @DocDate1 and @DocDate2 ");
            sb.Append(" GROUP BY 年月,DepartID,科目,科目名稱,LINE,COMPANY ORDER BY COMPANY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            System.Data.DataTable dt = ds.Tables[0];

            DataColumn[] colPk = new DataColumn[2];
            colPk[0] = dt.Columns["Model"];
            colPk[1] = dt.Columns["月份"];
            dt.PrimaryKey = colPk;


            return dt;


        }
        private System.Data.DataTable GetAcme_Acc_Model()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM ACME_ACC_MODEL");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            System.Data.DataTable dt = ds.Tables[0];

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["Model"];
            dt.PrimaryKey = colPk;


            return dt;


        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();

            StringBuilder sb2 = new StringBuilder();

            ArrayList al2 = new ArrayList();
            if (listBox1.SelectedItems.Count > 0)
            {

                for (int i = 0; i <= listBox1.SelectedItems.Count - 1; i++)
                {
                    al2.Add(listBox1.SelectedItems[i].ToString());

                }

                foreach (string v in al2)
                {
                    sb2.Append("'" + v + "',");
                }
                sb2.Remove(sb2.Length - 1, 1);
            }

            string J1 = sb2.ToString();
            if (J1.IndexOf("TFT") != -1)
            {
                System.Data.DataTable k1 = GetBU("TFT");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }


            }
            if (J1.IndexOf("禾豐牧場") != -1)
            {
                System.Data.DataTable k1 = GetBU("禾豐牧場");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }

            }
            if (J1.IndexOf("利豐漁場") != -1)
            {
                System.Data.DataTable k1 = GetBU("利豐漁場");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }

            }
            if (J1.IndexOf("董事室") != -1)
            {
                System.Data.DataTable k1 = GetBU("董事室");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }



            }
            if (J1.IndexOf("品牌行銷") != -1)
            {
                System.Data.DataTable k1 = GetBU("品牌行銷");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }
            }
            if (J1.IndexOf("生物科技") != -1)
            {
                System.Data.DataTable k1 = GetBU("生物科技");
                DataRow dr;
                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    dr = k1.Rows[i];

                    string PrcCode = Convert.ToString(dr["DEPT"]);

                    listBox3.Items.Add(PrcCode);
                }
            }
        }
        private System.Data.DataTable GetBU(string MEMO)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PARAM_NO DEPT  FROM PARAMS WHERE PARAM_KIND='GBCHO2' AND MEMO=@MEMO ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRC");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables["OPRC"];



        }
    }
}
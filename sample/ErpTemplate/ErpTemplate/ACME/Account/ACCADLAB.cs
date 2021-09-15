using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class ACCADLAB : Form
    {
        string str16 = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        System.Data.DataTable dtAD = null;
        int QQ1 = 0;
        int QQ2 = 0;
        int QQ3 = 0;
        int QQ4 = 0;
        int QQ5 = 0;
        int QQ6 = 0;

        System.Data.DataTable dtADD = null;
        int QQ1D = 0;
        int QQ2D = 0;
        int QQ3D = 0;
        int QQ4D = 0;
        int QQ5D = 0;

        public ACCADLAB()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             QQ1 = 0;
             QQ2 = 0;
             QQ3 = 0;
             QQ4 = 0;
             QQ5 = 0;
             QQ6 = 0;
            dtAD = MakeTableCombine();
            System.Data.DataTable F1 = GetCHO3();
            for (int i = 0; i <= F1.Rows.Count-1; i++)
            {
                string S1 = F1.Rows[i][0].ToString();

                    Eun22(S1);
                
    //            Eun22("EU");
            }

            DataRow dr = null;
            dr = dtAD.NewRow();

            dr["專案名稱"] =  "總計";
            dr["數量"] = QQ1;
            dr["收入"] = QQ2;

            dr["成本"] = QQ3;
            dr["毛利"] = QQ4;
            if (QQ4 == 0 || QQ2 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(QQ4) / Convert.ToDecimal(QQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["毛利率"] = G;
            }
            dr["費用匯差"] = QQ5;
            dr["稅前淨利"] = QQ6;

            if (QQ2 == 0 || QQ6 == 0)
            {
                dr["稅前淨利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(QQ6) / Convert.ToDecimal(QQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["稅前淨利率"] = G;
            }
            dtAD.Rows.Add(dr);


            QQ1D = 0;
            QQ2D = 0;
            QQ3D = 0;
            QQ4D = 0;
            QQ5D = 0;

            dtADD = MakeTableCombine2();
            for (int i = 0; i <= F1.Rows.Count - 1; i++)
            {
                string S1 = F1.Rows[i][0].ToString();
                System.Data.DataTable F2 = GetCHO4(S1);
                for (int i2 = 0; i2 <= F2.Rows.Count - 1; i2++)
                {
                    string S2 = F2.Rows[i2][0].ToString();
                    Eun22D(S2);
                }
            }

            DataRow drD = null;
            drD = dtADD.NewRow();

            drD["專案名稱"] = "總計";
            drD["數量"] = QQ1D;
            drD["收入"] = QQ2D;

            drD["成本"] = QQ3D;
            drD["毛利"] = QQ4D;
            if (QQ4D == 0 || QQ2D == 0)
            {
                drD["毛利率"] = "0.00%";
            }
            else
            {
                string GD = Math.Round((Convert.ToDecimal(QQ4D) / Convert.ToDecimal(QQ2D)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                drD["毛利率"] = GD;
            }
            drD["費用匯差"] = QQ5D;
            drD["稅前淨利"] = QQ4D - QQ5D;

            if (QQ2D == 0 || (QQ4D - QQ5D) == 0)
            {
                drD["稅前淨利率"] = "0.00%";
            }
            else
            {
                string GD = Math.Round((Convert.ToDecimal(QQ4D - QQ5D) / Convert.ToDecimal(QQ2D)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                drD["稅前淨利率"] = GD;
            }
            dtADD.Rows.Add(drD);


            dataGridView1.DataSource = dtAD;

            for (int i = 5; i <= 12; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }



            dataGridView2.DataSource = dtADD;

            for (int i = 10; i <= 17; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";


            }

        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("區域", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("專案編號", typeof(string));
            dt.Columns.Add("專案名稱", typeof(string));
            dt.Columns.Add("數量", typeof(int));
            dt.Columns.Add("收入", typeof(int));
            dt.Columns.Add("成本", typeof(int));
            dt.Columns.Add("毛利", typeof(int));
            dt.Columns.Add("毛利率", typeof(string));
            dt.Columns.Add("費用匯差", typeof(int));
            dt.Columns.Add("稅前淨利", typeof(int));
            dt.Columns.Add("稅前淨利率", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂購單號", typeof(string));
            dt.Columns.Add("區域", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("產品類別", typeof(string));
            dt.Columns.Add("專案編號", typeof(string));
            dt.Columns.Add("專案名稱", typeof(string));
            dt.Columns.Add("科目名稱", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("銷貨單號碼", typeof(string));
            dt.Columns.Add("數量", typeof(int));
            dt.Columns.Add("收入", typeof(int));
            dt.Columns.Add("成本", typeof(int));
            dt.Columns.Add("毛利", typeof(int));
            dt.Columns.Add("毛利率", typeof(string));
            dt.Columns.Add("費用匯差", typeof(int));
            dt.Columns.Add("稅前淨利", typeof(int));
            dt.Columns.Add("稅前淨利率", typeof(string));
            return dt;
        }
        private void Eun22(string LOC)
        {


            DataRow dr = null;

            System.Data.DataTable H1 = GetCHO(LOC);
            int Q1 = 0;
            int Q2 = 0;
            int Q3 = 0;
            int Q4 = 0;
            int Q5 = 0;
            int Q6 = 0;
            for (int i = 0; i <= H1.Rows.Count - 1; i++)
            {
                dr = dtAD.NewRow();
                dr["區域"] = H1.Rows[i]["區域"].ToString();
                dr["客戶編號"] = H1.Rows[i]["客戶編號"].ToString();
                dr["客戶名稱"] = H1.Rows[i]["客戶名稱"].ToString();
                dr["專案編號"] = H1.Rows[i]["專案編號"].ToString();

                dr["專案名稱"] = H1.Rows[i]["專案名稱"].ToString();
           
                dr["數量"] = Convert.ToInt32(H1.Rows[i]["數量"]);
                dr["收入"] = Convert.ToInt32(H1.Rows[i]["收入"]);

                dr["成本"] = Convert.ToInt32(H1.Rows[i]["成本"]);
                dr["毛利"] = Convert.ToInt32(H1.Rows[i]["毛利"]);
       
                dr["費用匯差"] = Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                dr["稅前淨利"] = Convert.ToInt32(H1.Rows[i]["稅前淨利"]);


                int QQQ2 = Convert.ToInt32(H1.Rows[i]["收入"]);
                int QQQ4 = Convert.ToInt32(H1.Rows[i]["毛利"]);
                int QQQ6 = Convert.ToInt32(H1.Rows[i]["稅前淨利"]);
                Q1 += Convert.ToInt32(H1.Rows[i]["數量"]);
                Q2 += Convert.ToInt32(H1.Rows[i]["收入"]);
                Q3 += Convert.ToInt32(H1.Rows[i]["成本"]);
                Q4 += Convert.ToInt32(H1.Rows[i]["毛利"]);
                Q5 += Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                Q6 += Convert.ToInt32(H1.Rows[i]["稅前淨利"]);
                if (QQQ4 == 0 || QQQ2 == 0)
                {
                    dr["毛利率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(QQQ4) / Convert.ToDecimal(QQQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }

                if (QQQ2 == 0 || QQQ6 == 0)
                {
                    dr["稅前淨利率"] = "0.00%";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(QQQ6) / Convert.ToDecimal(QQQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["稅前淨利率"] = G;
                }
                QQ1 += Convert.ToInt32(H1.Rows[i]["數量"]);
                QQ2 += Convert.ToInt32(H1.Rows[i]["收入"]);
                QQ3 += Convert.ToInt32(H1.Rows[i]["成本"]);
                QQ4 += Convert.ToInt32(H1.Rows[i]["毛利"]);
                QQ5 += Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                QQ6 += Convert.ToInt32(H1.Rows[i]["稅前淨利"]);
                dtAD.Rows.Add(dr);
            }


            dr = dtAD.NewRow();

            dr["專案名稱"] = "小計";
            dr["數量"] = Q1;
            dr["收入"] = Q2;

            dr["成本"] = Q3;
            dr["毛利"] = Q4;
            if (Q4 == 0 || Q2 == 0)
            {
                dr["毛利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["毛利率"] = G;
            }
            dr["費用匯差"] = Q5;
            dr["稅前淨利"] = Q6;

            if (Q2 == 0 || Q6 == 0)
            {
                dr["稅前淨利率"] = "0.00%";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(Q6) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["稅前淨利率"] = G;
            }
            dtAD.Rows.Add(dr);


        }
        private void Eun22D(string PRJ)
        {


            DataRow dr = null;

            System.Data.DataTable H1 = GetCHO2(PRJ);
            int Q1 = 0;
            int Q2 = 0;
            int Q3 = 0;
            int Q4 = 0;
            int Q5 = 0;
            for (int i = 0; i <= H1.Rows.Count - 1; i++)
            {
                dr = dtADD.NewRow();

       
                dr["訂購單號"] = H1.Rows[i]["訂購單號"].ToString();
                dr["產品類別"] = H1.Rows[i]["產品類別"].ToString();
                dr["科目名稱"] = H1.Rows[i]["科目名稱"].ToString();
                dr["區域"] = H1.Rows[i]["區域"].ToString();
                dr["客戶編號"] = H1.Rows[i]["客戶編號"].ToString();
                dr["客戶名稱"] = H1.Rows[i]["客戶名稱"].ToString();
                dr["專案編號"] = H1.Rows[i]["專案編號"].ToString();

                dr["專案名稱"] = H1.Rows[i]["專案名稱"].ToString();
                dr["日期"] = H1.Rows[i]["日期"].ToString();
                dr["銷貨單號碼"] = H1.Rows[i]["銷貨單號碼"].ToString();
                dr["數量"] = Convert.ToInt32(H1.Rows[i]["數量"]);
                dr["收入"] = Convert.ToInt32(H1.Rows[i]["收入"]);

                dr["成本"] = Convert.ToInt32(H1.Rows[i]["成本"]);
                dr["毛利"] = Convert.ToInt32(H1.Rows[i]["毛利"]);

                dr["費用匯差"] = Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                dr["稅前淨利"] = Convert.ToInt32(H1.Rows[i]["稅前淨利"]);


                int QQQ2 = Convert.ToInt32(H1.Rows[i]["收入"]);
                int QQQ4 = Convert.ToInt32(H1.Rows[i]["毛利"]);
                int QQQ6 = Convert.ToInt32(H1.Rows[i]["稅前淨利"]);
                Q1 += Convert.ToInt32(H1.Rows[i]["數量"]);
                Q2 += Convert.ToInt32(H1.Rows[i]["收入"]);
                Q3 += Convert.ToInt32(H1.Rows[i]["成本"]);
                Q4 += Convert.ToInt32(H1.Rows[i]["毛利"]);
                Q5 += Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                if (QQQ4 == 0 || QQQ2 == 0)
                {
                    dr["毛利率"] = "";
                }
                else
                {
                    string G = Math.Round((Convert.ToDecimal(QQQ4) / Convert.ToDecimal(QQQ2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                    dr["毛利率"] = G;
                }

        
                QQ1D += Convert.ToInt32(H1.Rows[i]["數量"]);
                QQ2D += Convert.ToInt32(H1.Rows[i]["收入"]);
                QQ3D += Convert.ToInt32(H1.Rows[i]["成本"]);
                QQ4D += Convert.ToInt32(H1.Rows[i]["毛利"]);
                QQ5D += Convert.ToInt32(H1.Rows[i]["費用匯差"]);
                dtADD.Rows.Add(dr);
            }


            dr = dtADD.NewRow();

            dr["專案名稱"] = H1.Rows[0]["專案名稱"].ToString() + "小計";
            dr["數量"] = Q1;
            dr["收入"] = Q2;

            dr["成本"] = Q3;
            dr["毛利"] = Q4;
            if (Q4 == 0 || Q2 == 0)
            {
                dr["毛利率"] = "";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(Q4) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["毛利率"] = G;
            }
            dr["費用匯差"] = Q5;
            dr["稅前淨利"] = Q4 - Q5;

            if (Q2 == 0 || (Q4 - Q5) == 0)
            {
                dr["稅前淨利率"] = "";
            }
            else
            {
                string G = Math.Round((Convert.ToDecimal(Q4 - Q5) / Convert.ToDecimal(Q2)) * 100, 2, MidpointRounding.AwayFromZero).ToString() + "%";
                dr["稅前淨利率"] = G;
            }
            dtADD.Rows.Add(dr);


        }
        public System.Data.DataTable GetCHO(string LOC)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT   SUBSTRING(MAX(P.ProjectName),1,2) 區域,MAX(T1.客戶編號) 客戶編號,MAX(T1.客戶名稱) 客戶名稱,A.ProjectID 專案編號,MAX(P.ProjectName) 專案名稱,  ");
            sb.Append(" ISNULL(MAX(數量),0) 數量,ISNULL(MAX(收入),0) 收入,ISNULL(MAX(成本),0) 成本,ISNULL(MAX(毛利),0) 毛利,  ");
            sb.Append(" ISNULL(SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)),0) 費用匯差,  ");
            sb.Append(" ISNULL(ISNULL(MAX(毛利),0)-(SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0))),0) 稅前淨利  ");
            sb.Append(" FROM  DBO.AccVoucherSub A  ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID   ");
            sb.Append(" LEFT JOIN (Select  MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN 0 ELSE A.Quantity*-1 END),0) 數量,      ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,      ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,  ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利  ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A    ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag         ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''  ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" GROUP BY T.ProjectID  ");
            sb.Append(" ) T1 ON (A.ProjectID =T1.專案編號) ");
            sb.Append(" WHERE  T0.MakeDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7) ");
            sb.Append(" AND SUBSTRING(P.ProjectName,1,2)=@LOC   ");
            sb.Append(" GROUP BY A.ProjectID ");
            sb.Append(" UNION ALL");
            sb.Append(" Select  SUBSTRING(MAX(P.ProjectName),1,2) 區域,MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號,MAX(P.ProjectName ) 專案名稱,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,    ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,    ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利,");
            sb.Append(" 0 費用匯差,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 稅前淨利");
            sb.Append(" From CHIComp16.DBO.ComProdRec A  ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag        ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1        ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  T.ProjectID=P.ProjectID AND U.Flag =1        ");
            sb.Append(" LEFT JOIN (SELECT A.ProjectID,SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT");
            sb.Append(" FROM  DBO.AccVoucherSub A");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo");
            sb.Append(" WHERE  T0.MakeDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7) GROUP BY  A.ProjectID) T0 ON (T.ProjectID=T0.ProjectID)                  ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> '' AND SUBSTRING(P.ProjectName,1,2)=@LOC    ");
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND T.PROJECTID NOT IN (");
            sb.Append(" SELECT   DISTINCT A.ProjectID 專案編號");
            sb.Append(" FROM  DBO.AccVoucherSub A ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo ");
            sb.Append(" WHERE  T0.MakeDate    between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7) ");
            sb.Append(" GROUP BY A.ProjectID )");
            sb.Append(" GROUP BY T.ProjectID ORDER BY 毛利 DESC,SUBSTRING(MAX(P.ProjectName) ,1,2) DESC,客戶名稱  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOC", LOC));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public System.Data.DataTable GetCHO2(string ProjectID)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  ''''+T1.BILLNO 訂購單號, SUBSTRING(MAX(P.ProjectName),1,2) 區域,MAX(T1.客戶編號) 客戶編號,MAX(T1.客戶名稱) 客戶名稱,MAX(CNAME) 產品類別,  ");
            sb.Append(" A.ProjectID 專案編號,MAX(P.ProjectName) 專案名稱, ");
            sb.Append(" CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),CAST(CAST(MAX(T1.BillDate) AS VARCHAR) AS DATETIME),20) - 1911) + '/' +");
            sb.Append(" SUBSTRING(CONVERT(VARCHAR(10),	  CAST(CAST(MAX(T1.BillDate) AS VARCHAR) AS DATETIME),20),6,2) + '/' +");
            sb.Append(" SUBSTRING(CONVERT(VARCHAR(10),	  CAST(CAST(MAX(T1.BillDate) AS VARCHAR) AS DATETIME),20),9,2) 日期,''''+MAX(T1.BILLNO2) 銷貨單號碼,");
            sb.Append(" '' 科目名稱,  ISNULL(MAX(數量),0) 數量,ISNULL(MAX(收入),0) 收入,  ");
            sb.Append(" ISNULL(MAX(成本),0) 成本,ISNULL(MAX(毛利),0) 毛利,0 費用匯差,0 稅前淨利   ");
            sb.Append(" FROM  DBO.AccVoucherSub A     ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo     ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID   ");
            sb.Append(" INNER JOIN (Select   MAX(A.FromNO)  BILLNO, MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN 0 ELSE A.Quantity*-1 END),0) 數量,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利     ");
            sb.Append(" ,MAX(K.ClassName ) CNAME,A.ProdID,A.BILLNO BILLNO2,A.BillDate    From CHIComp16.DBO.ComProdRec A    ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag            ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1            ");
            sb.Append(" Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID              ");
            sb.Append(" Left Join CHIComp16.DBO.comProductClass K On J.ClassID =K.ClassID              ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''   AND A.BillDate   between   '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'     ");
            sb.Append(" GROUP BY T.ProjectID,A.FromNO,A.ProdID,A.BILLNO,A.BillDate       ) T1 ON (A.ProjectID =T1.專案編號)    ");
            sb.Append(" WHERE  T0.MakeDate   between   '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   AND ISNULL(A.ProjectID,'')=@ProjectID  AND SUBSTRING(A.SUBJECTID,1,1) IN (5,6,7)    ");
            sb.Append(" GROUP BY A.ProjectID,T1.BILLNO,T1.ProdID,T1.BILLNO2  ");
            sb.Append(" UNION ALL   ");
            sb.Append("              SELECT DISTINCT  '', SUBSTRING(MAX(P.ProjectName),1,2) 區域,MAX(CASE WHEN ISNULL(A.ProjectID,'')='A1-1907004' THEN 'TW216' ELSE T1.客戶編號 END) 客戶編號,MAX(CASE WHEN ISNULL(A.ProjectID,'')='A1-1907004' THEN '昇恒昌股份有限公司' ELSE T1.客戶名稱 END) 客戶名稱,'' 產品類別,     ");
            sb.Append("              A.ProjectID 專案編號,MAX(P.ProjectName) 專案名稱,'','',(C.SubjectName) 科目名稱, '' 數量,'',   ");
            sb.Append("              '' 成本,'' 毛利,ISNULL(SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)),0) 費用匯差, ''    ");
            sb.Append("              FROM  DBO.AccVoucherSub A      ");
            sb.Append("              Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo      ");
            sb.Append("              Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID    ");
            sb.Append("              Left Join CHIComp16.DBO.ComSubject C On C.SubjectID=A.SubjectID   ");
            sb.Append("              LEFT JOIN (Select  MAX(A.FromNO)  BILLNO, MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利      ");
            sb.Append("              ,MAX(K.ClassName ) CNAME  From CHIComp16.DBO.ComProdRec A     ");
            sb.Append("              Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag             ");
            sb.Append("              Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1             ");
            sb.Append("              Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID               ");
            sb.Append("              Left Join CHIComp16.DBO.comProductClass K On J.ClassID =K.ClassID               ");
            sb.Append("              Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''   AND A.BillDate       between   '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'    ");
            sb.Append("              GROUP BY  T.ProjectID     ) T1 ON (A.ProjectID =T1.專案編號)     ");
            sb.Append("              WHERE  T0.MakeDate  between   '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'     ");
            sb.Append("			      AND ISNULL(A.ProjectID,'')=@ProjectID  AND SUBSTRING(A.SUBJECTID,1,1) IN (5,6,7)     ");
            sb.Append("				   GROUP BY A.ProjectID,C.SubjectName  ");
            sb.Append(" UNION ALL   ");
            sb.Append(" Select  ''''+A.FromNO  , SUBSTRING(MAX(P.ProjectName),1,2) 區域,MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,MAX(K.ClassName)  產品類別,(T.ProjectID) 專案編號,MAX(P.ProjectName ) 專案名稱");
            sb.Append(" , CONVERT(VARCHAR(3),CONVERT(VARCHAR(4),CAST(CAST(MAX(A.BillDate) AS VARCHAR) AS DATETIME),20) - 1911) + '/' +");
            sb.Append(" SUBSTRING(CONVERT(VARCHAR(10),	  CAST(CAST(MAX(A.BillDate) AS VARCHAR) AS DATETIME),20),6,2) + '/' +");
            sb.Append(" SUBSTRING(CONVERT(VARCHAR(10),	  CAST(CAST(MAX(A.BillDate) AS VARCHAR) AS DATETIME),20),9,2) 日期,''''+MAX(A.BILLNO) 銷貨單號碼,");
            sb.Append(" MAX(T0.SubjectName) 科目名稱,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,     ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,     ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利, 0 費用匯差,0 稅前淨利   ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A    ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag          ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  T.ProjectID=P.ProjectID AND U.Flag =1      ");
            sb.Append(" Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID              ");
            sb.Append(" Left Join CHIComp16.DBO.comProductClass K On J.ClassID =K.ClassID             ");
            sb.Append(" LEFT JOIN (SELECT A.ProjectID,MAX(C.SubjectName) SubjectName,SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT FROM  CHIComp16.DBO.AccVoucherSub A   ");
            sb.Append(" Left Join CHIComp16.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo   ");
            sb.Append(" Left Join CHIComp16.DBO.ComSubject C On C.SubjectID=A.SubjectID  ");
            sb.Append(" WHERE  T0.MakeDate   between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'      AND SUBSTRING(A.SUBJECTID,1,1) IN (5,6,7) GROUP BY  A.ProjectID) T0 ON (T.ProjectID=T0.ProjectID)                 ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') =@ProjectID   ");
            sb.Append(" AND A.BillDate   between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'        AND T.PROJECTID NOT IN ( SELECT   DISTINCT A.ProjectID 專案編號   ");
            sb.Append(" FROM  DBO.AccVoucherSub A    ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  WHERE  T0.MakeDate    between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'        AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7)  GROUP BY A.ProjectID )   ");
            sb.Append(" GROUP BY T.ProjectID,A.FromNO,A.ProdID ORDER BY SUBSTRING(MAX(P.ProjectName),1,2) DESC,客戶名稱,專案編號   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));
            //=@ProjectID
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public System.Data.DataTable GetCHO3()
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 專案編號,SUM(毛利) 毛利 FROM (      SELECT   SUBSTRING(P.ProjectName,1,2) 專案編號,ISNULL(MAX(毛利),0) 毛利");
            sb.Append(" FROM  CHIComp16.DBO.AccVoucherSub A   ");
            sb.Append(" Left Join CHIComp16.DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo   ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID    ");
            sb.Append(" LEFT JOIN (Select  MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,       ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,       ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,   ");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利   ");
            sb.Append(" From CHIComp16.DBO.ComProdRec A     ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag            ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1              ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''   ");
            sb.Append(" AND A.BillDate   between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" GROUP BY T.ProjectID   ");
            sb.Append(" ) T1 ON (A.ProjectID =T1.專案編號)  ");
            sb.Append(" WHERE  T0.MakeDate   between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7)  ");
            sb.Append(" GROUP BY   SUBSTRING(P.ProjectName,1,2)");
            sb.Append(" UNION ALL ");
            sb.Append(" Select  SUBSTRING(P.ProjectName,1,2) 專案編號,");
            sb.Append(" ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利");
            sb.Append(" From CHIComp16.DBO.ComProdRec A   ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag     ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1         ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  T.ProjectID=P.ProjectID AND U.Flag =1         ");
            sb.Append(" LEFT JOIN (SELECT A.ProjectID,SUM(ISNULL((CASE DebitCredit WHEN 1 THEN AMOUNT END),0)) -SUM(ISNULL((CASE DebitCredit WHEN 0 THEN AMOUNT END),0)) AMT ");
            sb.Append(" FROM  DBO.AccVoucherSub A ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo ");
            sb.Append(" WHERE  T0.MakeDate   between   '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7) GROUP BY  A.ProjectID) T0 ON (T.ProjectID=T0.ProjectID)                   ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> '' ");
            sb.Append(" AND A.BillDate   between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND T.PROJECTID NOT IN ( ");
            sb.Append(" SELECT   DISTINCT A.ProjectID 專案編號 ");
            sb.Append(" FROM  DBO.AccVoucherSub A  ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  ");
            sb.Append(" WHERE  T0.MakeDate    between  '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "' ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7))");
            sb.Append(" GROUP BY   SUBSTRING(P.ProjectName,1,2) ) AS A");
            sb.Append(" GROUP BY 專案編號 ORDER BY 毛利 DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public System.Data.DataTable GetCHO4(string LOC)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 專案編號,SUM(毛利) 毛利  FROM (");
            sb.Append(" SELECT   A.ProjectID 專案編號,ISNULL(MAX(毛利),0) 毛利");
            sb.Append(" FROM  DBO.AccVoucherSub A    ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo    ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID  ");
            sb.Append(" Left Join CHIComp16.DBO.ComSubject C On C.SubjectID=A.SubjectID ");
            sb.Append(" INNER JOIN (Select  A.FromNO  BILLNO, MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利    ");
            sb.Append(" ,MAX(K.ClassName ) CNAME,A.ProdID  From CHIComp16.DBO.ComProdRec A   ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag         ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1           ");
            sb.Append(" Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID             ");
            sb.Append(" Left Join CHIComp16.DBO.comProductClass K On J.ClassID =K.ClassID             ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''   AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'    ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" GROUP BY T.ProjectID,A.FromNO,A.ProdID    ) T1 ON (A.ProjectID =T1.專案編號)  ");
            sb.Append(" WHERE  T0.MakeDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'    AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(A.SUBJECTID,1,1) IN (5,6,7)   ");
            sb.Append(" AND SUBSTRING(P.ProjectName,1,2)=@LOC     GROUP BY A.ProjectID");
            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT DISTINCT");
            sb.Append(" A.ProjectID 專案編號,'' 毛利");
            sb.Append(" FROM  DBO.AccVoucherSub A    ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo    ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  A.ProjectID=P.ProjectID  ");
            sb.Append(" Left Join CHIComp16.DBO.ComSubject C On C.SubjectID=A.SubjectID ");
            sb.Append(" LEFT JOIN (Select  A.FromNO  BILLNO, MAX(T.CustID) 客戶編號,MAX(U.FullName) 客戶名稱,(T.ProjectID) 專案編號, ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.Quantity  WHEN A.Flag=701 THEN A.Quantity*-1 ELSE A.Quantity*-1 END),0) 數量,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END),0) 收入,       ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 成本,   ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利    ");
            sb.Append(" ,MAX(K.ClassName ) CNAME,A.ProdID  From CHIComp16.DBO.ComProdRec A   ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag       ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1           ");
            sb.Append(" Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID             ");
            sb.Append(" Left Join CHIComp16.DBO.comProductClass K On J.ClassID =K.ClassID             ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> ''   AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'   ");
            sb.Append(" GROUP BY T.ProjectID,A.FromNO,A.ProdID    ) T1 ON (A.ProjectID =T1.專案編號)   ");
            sb.Append(" WHERE  T0.MakeDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(A.SUBJECTID,1,1) IN (5,6,7)   ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND A.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND SUBSTRING(P.ProjectName,1,2)=@LOC  GROUP BY A.ProjectID,T1.BILLNO,T1.ProdID,C.SubjectName");
            sb.Append(" UNION ALL  ");
            sb.Append(" Select (T.ProjectID) 專案編號,ISNULL(SUM(CASE WHEN A.Flag=500 THEN A.MLAmount WHEN A.Flag=701 THEN A.MLDIST*-1 ELSE A.MLAmount*-1 END-CASE WHEN A.Flag=500 THEN A.CostForAcc ELSE  A.CostForAcc*-1 END),0) 毛利");
            sb.Append(" From CHIComp16.DBO.ComProdRec A   ");
            sb.Append(" Left join CHIComp16.DBO.comBillAccounts T ON A.BillNO=T.FundBillNo  AND CASE A.Flag WHEN 701 THEN 698 ELSE A.Flag END=T.Flag        ");
            sb.Append(" Left join CHIComp16.DBO.comCustomer U On  U.ID=T.CustID AND U.Flag =1        ");
            sb.Append(" Left join CHIComp16.DBO.comProject P On  T.ProjectID=P.ProjectID AND U.Flag =1     ");
            sb.Append(" Left Join CHIComp16.DBO.comProduct J On A.ProdID =J.ProdID             ");
            sb.Append(" Where A.Flag IN (500,600,701)  AND ISNULL( T.ProjectID,'') <> '' AND SUBSTRING(P.ProjectName,1,2)=@LOC      ");
            if (textBox3.Text != "")
            {
                sb.Append(" AND T.ProjectID= '" + textBox3.Text.ToString() + "'    ");
            }
            sb.Append(" AND A.BillDate   between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'  AND T.PROJECTID NOT IN ( SELECT   DISTINCT A.ProjectID 專案編號  ");
            sb.Append(" FROM  DBO.AccVoucherSub A   ");
            sb.Append(" Left Join DBO.accVoucherMain T0 On T0.VoucherNo=A.VoucherNo  WHERE  T0.MakeDate    between '" + textBox1.Text.ToString() + "' and '" + textBox2.Text.ToString() + "'    AND ISNULL(A.ProjectID,'') <>'' AND SUBSTRING(SUBJECTID,1,1) IN (5,6,7)  GROUP BY A.ProjectID )  ");
            sb.Append(" GROUP BY T.ProjectID) A GROUP BY 專案編號  ORDER BY 毛利 DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LOC", LOC));
            command.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }
        private void ACCADLAB_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }
    }
}

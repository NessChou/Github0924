using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;


namespace ACME
{
    
    /// <summary>    
    /// 專案管理 http://twpug.net/modules/smartsection/item.php?itemid=41
    /// 
    /// </summary>
    public partial class fmAcmeSolar : Form
    {

        private string LoginUser;
        private string globalPrjCode;
        private string globalPrjName;

        int scrollPosition = 0;



        public fmAcmeSolar()
        {
            InitializeComponent();

        
            
            gvORDR.AutoGenerateColumns = false;
            gvRdr1.AutoGenerateColumns = false;


            gvOPOR.AutoGenerateColumns = false;
            gvPOR1.AutoGenerateColumns = false;


            gvOwor.AutoGenerateColumns = false;
            gvWor1.AutoGenerateColumns = false;


            gvExpense.AutoGenerateColumns = false;
            gvRevenue.AutoGenerateColumns = false;


            gvAsset.AutoGenerateColumns = false;
            gvAsset2.AutoGenerateColumns = false;
                  }

        public DataTable GetProject()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT distinct O.PrjCode,O.PrjName,U_MEMO MEMO,U_MEMO2 MEMO2  FROM OPRJ O    WHERE  Substring(O.PrjCode,1,1)='4' AND ISNULL(U_MEMO,'') NOT　IN ('專案碼刪除','撤案') ORDER BY PrjCode   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }


        public DataTable GetProjectExist(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;
            //string sql = "SELECT PrjCode,PrjName,0 as PrjPercent FROM OPRJ WHERE  PrjCode=@PrjCode";

            string sql = "SELECT O.PrjCode,O.PrjName,isnull(s.PrjPercent,0) PrjPercent FROM OPRJ O inner join acmesqlsp..acme_task_solar s on  O.PrjCode=s.PrjCode COLLATE  Chinese_Taiwan_Stroke_CI_AS  WHERE  O.PrjCode=@PrjCode";
            //string sql = "SELECT PrjCode,PrjName FROM OPRJ ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        public DataTable GetAllProject()
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT * from acme_task_solar  ";
            //string sql = "SELECT PrjCode,PrjName FROM OPRJ ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        private void fmAcmeSolar_Load(object sender, EventArgs e)
        {
            string PRJ = "";
            try
            {
                textBox9.Text = DateTime.Now.ToString("yyyyMM");
                textBox12.Text = DateTime.Now.ToString("yyyyMMdd");

                gvProject.AutoGenerateColumns = false;

                DataTable dt = GetProject();
                gvProject.DataSource = dt;

                BindProject(Convert.ToString(dt.Rows[0]["PrjCode"]), Convert.ToString(dt.Rows[0]["PrjName"]));

                System.Data.DataTable G1 = GetORDR_Project(globalPrjCode);
                gvORDR.DataSource = G1;
                if (G1.Rows.Count > 0)
                {
                    txtCardName.Text = G1.Rows[0]["CardName"].ToString();
                    txtSlpName.Text = G1.Rows[0]["SLPNAME"].ToString();
                    textBox5.Text = G1.Rows[G1.Rows.Count - 1]["DocDueDate"].ToString();
                    textBox6.Text = G1.Rows[0]["MEMO"].ToString();
                }
                else
                {
                    txtCardName.Text = "";
                    txtSlpName.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                }
                //採購單
                gvOPOR.DataSource = GetOPOR_Project(globalPrjCode);


                DataTable dtOwor = GetOwor_Project(globalPrjCode);
                gvOwor.DataSource = dtOwor;


                DataTable dtExpense = GetExpense(globalPrjCode);
                gvExpense.DataSource = dtExpense;

                DataTable dtRevenue = GetRevenue(globalPrjCode);
                gvRevenue.DataSource = dtRevenue;

                DataTable dtExpense1 = GetExpense1(globalPrjCode);
                decimal T1 = Convert.ToDecimal(dtExpense1.Rows[0][0]);
                label16.Text = "金額加總 : " + T1.ToString("#,##0");
                if (T1.ToString("#,##0") == "0")
                {
                    label16.Text = "";
                }

                decimal T2 = Convert.ToDecimal(GetORDRCOUNT2(globalPrjCode, 1).Rows[0][0]);
                label15.Text = "金額加總 : " + T2.ToString("#,##0");
                textBox2.Text = (Convert.ToDouble(T2)).ToString("#,##0");

                if (T2.ToString("#,##0") == "0")
                {
                    label15.Text = "";
                }
                //Depreciation 折舊
                DataTable dtAsset = GetAsset(globalPrjCode);
                gvAsset.DataSource = dtAsset;
                DataTable dtAsset1 = GetAsset1(globalPrjCode);
               // double T3 = (Convert.ToDouble(dtAsset1.Rows[0][0]) + Convert.ToDouble(T1)) * 1.05;
                double T3 = (Convert.ToDouble(dtAsset1.Rows[0][0]) + Convert.ToDouble(T1));
                textBox3.Text = T3.ToString("#,##0");
                textBox8.Text = GetORDRCOUNT(globalPrjCode, 1).Rows[0][0].ToString();
                textBox7.Text = GetORDRCOUNT(globalPrjCode, 0).Rows[0][0].ToString();

                System.Data.DataTable ESCO2 = GetESCO_PROJECT2(globalPrjCode);
                if (ESCO2.Rows.Count > 0)
                {
                    double M1 = Convert.ToDouble(ESCO2.Rows[0][0].ToString());
                    textBox10.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox10.Text = "";
                }

                //Depreciation 累計折舊
                double M2 = Convert.ToDouble(GetORDRCOUNT(globalPrjCode, 1).Rows[0][0]);
                if (M2 != 0)
                {
                    double M1 = ((Convert.ToDouble(T2)) / M2) * Convert.ToDouble(GetORDRCOUNT(globalPrjCode, 0).Rows[0][0]);
                    textBox4.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox4.Text = "";
                }


                System.Data.DataTable ESCO1 = GetESCO_PROJECT(globalPrjCode);
                if (ESCO1.Rows.Count > 0)
                {
                    double M1 = Convert.ToDouble(ESCO1.Rows[0][0].ToString());
                    textBox1.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox1.Text = "";
                }
                System.Data.DataTable ESCO3 = GetESCO_PROJECT3(globalPrjCode);
                if (ESCO3.Rows.Count > 0)
                {
                    string M1 = ESCO3.Rows[0][0].ToString();
                    currentStage.Text = M1.ToString();
                }
                else
                {
                    currentStage.Text = "";
                }

                DataTable dtAsset2 = GetAsset2(globalPrjCode);
                gvAsset2.DataSource = dtAsset2;


                DataRow dr = null;
                System.Data.DataTable dtCost = MakeTable();
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();

                    string PRJCODE = dt.Rows[i]["PrjCode"].ToString();
                    PRJ = PRJCODE;
                   
                    dr["代碼"] = PRJCODE;
                    dr["名稱"] = dt.Rows[i]["PrjName"].ToString();
              
                    System.Data.DataTable G1S = GetORDR_Project(PRJCODE);
                    if (G1S.Rows.Count > 0)
                    {
                        dr["負責業務"] = G1S.Rows[0]["SLPNAME"].ToString();
                        dr["合約生效日"] = G1S.Rows[0]["MEMO"].ToString();
                        dr["合約終止日"] = G1S.Rows[G1S.Rows.Count-1]["DocDueDate"].ToString();
                    

                    }


                    System.Data.DataTable ESCO1S = GetESCO_PROJECT(PRJCODE);
                    if (ESCO1S.Rows.Count > 0)
                    {
                        dr["預估分享"] = ESCO1S.Rows[0]["SHARE"].ToString();
                    }
                    else
                    {
                        dr["預估分享"] = 0;
                    }

                    //實際已收入
                    decimal T2S = Convert.ToDecimal(GetORDRCOUNT2(PRJCODE, 1).Rows[0][0]);
                    dr["已收帳款"] = T2S;
                    double M2S = Convert.ToDouble(GetORDRCOUNT(PRJCODE, 1).Rows[0][0]);
                    if (M2S != 0)
                    {
                        double M1S = ((Convert.ToDouble(T2S)) / M2S) * Convert.ToDouble(GetORDRCOUNT(PRJCODE, 0).Rows[0][0]);
                        dr["調整後預估分享"] = M1S;
                    }
                    else
                    {
                        dr["調整後預估分享"] = 0;
                    }
                    System.Data.DataTable P1 = GetESCO_PROJECT2(PRJCODE);
                    if (P1.Rows.Count > 0)
                    {
                        double M3S = Convert.ToDouble(P1.Rows[0][0]);
                        dr["服務總金額"] = M3S;
                    }
                    else
                    {
                        dr["服務總金額"] = 0;
                    }
                    DataTable dtExpense1S = GetExpense1(PRJCODE);
                    decimal T1S = Convert.ToDecimal(dtExpense1S.Rows[0][0]);
                    DataTable dtAsset1S = GetAsset1(PRJCODE);
                    //double T3S = (Convert.ToDouble(dtAsset1S.Rows[0][0]) + Convert.ToDouble(T1S)) * 1.05;
                    double T3S = (Convert.ToDouble(dtAsset1S.Rows[0][0]) + Convert.ToDouble(T1S));
                    dr["成本"] = T3S;
                    dr["己收期數"] = GetORDRCOUNT(PRJCODE, 1).Rows[0][0].ToString();
                    dr["應收期數"] = GetORDRCOUNT(PRJCODE, 0).Rows[0][0].ToString();
                    System.Data.DataTable P3 = GetESCO_PROJECT3(PRJCODE);
                    if (P3.Rows.Count > 0)
                    {
                        dr["專案階段"] = P3.Rows[0][0].ToString();
                    }
                    dtCost.Rows.Add(dr);
                }

             
                decimal[] Total = new decimal[7];

                for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
                {

                    for (int j = 3; j <= 7; j++)
                    {
                        Total[j - 1] += Convert.ToDecimal(dtCost.Rows[i][j]);

                    }
                }

                DataRow row;

                row = dtCost.NewRow();

                row[2] = "合計";
                for (int j = 3; j <= 7; j++)
                {
                    row[j] = Total[j - 1];

                }
                dtCost.Rows.Add(row);



                dataGridView1.DataSource = dtCost;
                for (int i = 3; i <= 9; i++)
                {
                    DataGridViewColumn col = dataGridView1.Columns[i];


                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    col.DefaultCellStyle.Format = "#,##0";
                }


                UtilSimple.SetLookupBinding(cbYear, Year(), "DataValue", "DataValue");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + PRJ);
            }
        }

        private void BindProject(string PrjCode, string PrjName)
        {

            textBoxPrjCode.Text = PrjCode;
            textBoxPrjName.Text = PrjName;

            globalPrjCode = PrjCode;
            globalPrjName = PrjName;

        }

        private void gvProject_SelectionChanged(object sender, EventArgs e)
        {
           
            //避免觸發
            if (!gvProject.Focused) return;

            //try
            //{
                //MessageBox.Show(gvProject.CurrentRow.Cells[0].Value.ToString());
                BindProject(gvProject.CurrentRow.Cells[0].Value.ToString(), gvProject.CurrentRow.Cells[1].Value.ToString());

                DataTable dt = GetORDR_Project(globalPrjCode);
                gvORDR.DataSource = dt;

                if (dt.Rows.Count == 0)
                {
                    gvRdr1.DataSource = null;
                    txtCardName.Text = "";
                    txtSlpName.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                }
                else
                {
                    txtCardName.Text = dt.Rows[0]["CardName"].ToString();
                    txtSlpName.Text = dt.Rows[0]["SLPNAME"].ToString();
                    textBox5.Text = dt.Rows[dt.Rows.Count-1]["DocDueDate"].ToString();
                    textBox6.Text = dt.Rows[0]["MEMO"].ToString();
                }


                //採購單
                DataTable dtOPOR = GetOPOR_Project(globalPrjCode);
                gvOPOR.DataSource = dtOPOR;

                if (dtOPOR.Rows.Count == 0)
                {
                    //處理明細檔
                    gvPOR1.DataSource = null;
                }
                


                DataTable dtOwor = GetOwor_Project(globalPrjCode);
                gvOwor.DataSource = dtOwor;

                if (dtOwor.Rows.Count == 0)
                {
                    gvWor1.DataSource = null;
                }

                DataTable dtExpense = GetExpense(globalPrjCode);
                gvExpense.DataSource = dtExpense;

                DataTable dtRevenue = GetRevenue(globalPrjCode);
                gvRevenue.DataSource = dtRevenue;

                //Depreciation 折舊
                DataTable dtAsset = GetAsset(globalPrjCode);
                gvAsset.DataSource = dtAsset;


                //Depreciation 累計折舊
                DataTable dtAsset2 = GetAsset2(globalPrjCode);
                gvAsset2.DataSource = dtAsset2;

                DataTable dtExpense1 = GetExpense1(globalPrjCode);
                decimal T1 = Convert.ToDecimal(dtExpense1.Rows[0][0]);
                label16.Text = "金額加總 : " + T1.ToString("#,##0");
                if (T1.ToString("#,##0") == "0")
                {
                    label16.Text = "";
                }


                decimal T2 = Convert.ToDecimal(GetORDRCOUNT2(globalPrjCode, 1).Rows[0][0]);
                label15.Text = "金額加總 : " + T2.ToString("#,##0");
                //textBox2.Text = (Convert.ToDouble(T2) * 1.05).ToString("#,##0");
                textBox2.Text = (Convert.ToDouble(T2)).ToString("#,##0");
                if (T2.ToString("#,##0") == "0")
                {
                    label15.Text = "";
                }


                DataTable dtAsset1 = GetAsset1(globalPrjCode);
              //  double T3 = (Convert.ToDouble(dtAsset1.Rows[0][0]) + Convert.ToDouble(T1)) * 1.05;
                double T3 = (Convert.ToDouble(dtAsset1.Rows[0][0]) + Convert.ToDouble(T1));
                textBox3.Text = T3.ToString("#,##0");
                textBox8.Text = GetORDRCOUNT(globalPrjCode, 1).Rows[0][0].ToString();
                textBox7.Text = GetORDRCOUNT(globalPrjCode, 0).Rows[0][0].ToString();

                System.Data.DataTable ESCO2 = GetESCO_PROJECT2(globalPrjCode);
                if (ESCO2.Rows.Count > 0)
                {
                    double M1 = Convert.ToDouble(ESCO2.Rows[0][0].ToString());
                    textBox10.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox10.Text = "";
                }

                //Depreciation 累計折舊
                double M2 = Convert.ToDouble(GetORDRCOUNT(globalPrjCode, 1).Rows[0][0]);
                if (M2 != 0)
                {
                    double M1 = ((Convert.ToDouble(T2)) / M2) * Convert.ToDouble(GetORDRCOUNT(globalPrjCode, 0).Rows[0][0]);
                    textBox4.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox4.Text = "";
                }



                System.Data.DataTable ESCO1 = GetESCO_PROJECT(globalPrjCode);
                if (ESCO1.Rows.Count > 0)
                {
                    double M1 = Convert.ToDouble(ESCO1.Rows[0][0].ToString());
                    textBox1.Text = M1.ToString("#,##0");
                }
                else
                {
                    textBox1.Text = "";
                }

                System.Data.DataTable ESCO3 = GetESCO_PROJECT3(globalPrjCode);
                if (ESCO3.Rows.Count > 0)
                {
                    string M1 = ESCO3.Rows[0][0].ToString();
                    currentStage.Text = M1.ToString();
                }
                else
                {
                    currentStage.Text = "";
                }
            //}
            //catch
            //{

            //}
            //
           
        }

    
        public DataTable GetESCO_PROJECT(string PROJECT)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT SHARE  FROM ESCO_PROJECT  WHERE PROJECT=@PROJECT";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetESCO_PROJECT2(string PROJECT)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT ISNULL(SHARE,0)  FROM ESCO_PROJECT2  WHERE PROJECT=@PROJECT AND SHARE <> 0 ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetESCO_PROJECT3(string PROJECT)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT STAGE  FROM ESCO_PROJECT3  WHERE PROJECT=@PROJECT AND STAGE <> '' ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetORDR_Project(string Project)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT DocEntry,CardName,DocDate,CARDCODE,Convert(varchar(8),DocDueDate,112)  DocDueDate,U_ACME_MEMO1 MEMO,CASE DocStatus WHEN 'C' THEN '已結' WHEN 'O' THEN '未結' END  DocStatus,T1.SLPNAME,T0.U_ACME_Warranty AWS編號 FROM ORDR T0 LEFT JOIN OSLP T1 ON (T0.SLPCODE=T1.SLPCODE)  WHERE    Project=@Project AND T0.DOCENTRY NOT IN ('20569','27161')  ORDER BY T0.DOCENTRY DESC";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetORDRCOUNT(string Project,int T)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(COUNT(*),'') COUN  FROM ORDR T0 ");
            sb.Append(" LEFT JOIN RDR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" WHERE  T0.Project=@Project AND T0.DOCENTRY NOT IN ('20569','27161')  AND CANCELED = 'N' AND DSCRIPTION NOT LIKE '%工程款項明%' AND ISNULL(T1.U_BASE_DOC,'') <> 'N'");
            if (T == 1)
            {
                sb.Append(" AND (T1.LINESTATUS = 'C' OR ISNULL(T1.U_BASE_DOC,'')='Y')   ");
            
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetORDRCOUNT2(string Project, int T)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(SUM(LINETOTAL),0) LINETOTAL FROM ORDR T0 ");
            sb.Append(" LEFT JOIN RDR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append(" WHERE  T0.Project=@Project AND T0.DOCENTRY NOT IN ('20569','27161')  AND CANCELED = 'N' AND DSCRIPTION NOT LIKE '%工程款項明%'");
            if (T == 1)
            {
                sb.Append(" AND (T1.LINESTATUS = 'C' OR T1.U_BASE_DOC='Y')  ");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }
        public DataTable GetOPOR_Project(string Project)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT DISTINCT T0.DocEntry,T0.CardName,T0.DocDate,T0.DocDueDate,CASE T0.DocStatus WHEN 'C' THEN '已結' WHEN 'O' THEN '未結' END  DocStatus  FROM OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE   T1.Project=@Project";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        public DataTable GetPOR1(string DocEntry,string Project)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT * FROM POR1 WHERE  DocEntry=@DocEntry AND Project=@Project";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        private void gvORDR_SelectionChanged(object sender, EventArgs e)
        {
            string DocEntry = gvORDR.CurrentRow.Cells[0].Value.ToString();
            gvRdr1.DataSource = GetORDR1(DocEntry);
        }


        public DataTable GetORDR1(string DocEntry)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM RDR1 T0");
            sb.Append(" WHERE  T0.DocEntry=@DocEntry AND  DSCRIPTION NOT LIKE '%工程款項明%' ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }



        private void gvOwor_SelectionChanged(object sender, EventArgs e)
        {
            string DocEntry = gvOwor.CurrentRow.Cells[0].Value.ToString();
            gvWor1.DataSource = GetWor1(DocEntry);


        }

        public DataTable GetOwor(string DocEntry)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT * FROM Owor WHERE  DocEntry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        public DataTable GetWor1(string DocEntry)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT W.*,I.ItemName FROM Wor1 W  inner join Oitm I on I.ItemCode=W.ItemCode WHERE  W.DocEntry=@DocEntry ";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }

        public DataTable GetOwor_Project(string Project)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT * FROM Owor WHERE  u_projectcode=@Project";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Project", Project));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }








        

      
        public DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }


        public string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }



        /* 說明:判斷輸入的日期是否正確 
         * 傳入值:        
         * 回傳值:
         * 格  式: yyyyMMdd 
         * 範  例:StrToDate(sDate) 
         */
        public bool IsDateString(string sDate)
        {

            try
            {
                UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
                UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
                UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));
                return true;
            }
            catch
            {
                return false;
            }


        }




        public int GetACME_TASKS_Count(string ProjectID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT Count(*) FROM ACME_TASKS  WHERE  ProjectID=@ProjectID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));
            
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }



        public int GetACME_TASK_Solar_Count(string PrjCode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT Count(*) FROM ACME_TASK_Solar  WHERE  PrjCode=@PrjCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                connection.Close();
            }
        }

        /// <summary>
        /// 取得太陽能專案預設工作樣版
        /// </summary>
        /// <param name="TpID"></param>
        /// <returns></returns>
        public DataTable GetACME_TASK_TL(string TpID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT * FROM ACME_TASK_TL  WHERE  TpID=@TpID Order by SortID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TpID", TpID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }


        public DataTable GetACME_TASKS(string ProjectID)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT * FROM ACME_TASKS  WHERE  ProjectID=@ProjectID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProjectID", ProjectID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPRJ");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0];
        }



        private void GenerateTree(DataTable dt, string rootID, DataTable dbTree, int level)
        {

            string strExp;

            if (rootID == "")
            {
                //strExp = "[ParentID] is null or [ParentID]=''";
                strExp = "[ParentID] is null";
            }
            else
            {
                strExp = "[ParentID] = " + rootID;
            }

            DataRow[] childRows = dt.Select(strExp);

            foreach (DataRow dr in childRows)
            {
                string rowID = dr["TaskID"].ToString();
                string ParentID = dr["ParentID"].ToString();
                string Title = Convert.ToString(dr["Title"]);

                DataRow drTree;
                drTree = dbTree.NewRow();
                drTree["TaskID"] = rowID;
                drTree["Title"] = blankStr(level) + Title;
                drTree["Level"] = level.ToString();

                dbTree.Rows.Add(drTree);


                GenerateTree(dt, rowID, dbTree, level + 1);
            }

        }

        private string blankStr(int count)
        {

            string s = "";
            for (int i = 1; i < count; i++)
            {
                s += "  ";
            }
            return s;
        }



     
        private bool CheckDate(TextBox t)
        {

            if (t.Text.Trim()=="")
                return true;

            try
            {
                StrToDate(t.Text);
                return true;
            }
            catch
            {
                MessageBox.Show("日期輸入錯誤");
                Control c = t.Parent;

                if (c is TabPage)
                {

                    //tabControl3.SelectedTab = (c as TabPage);
             //     tabControl3.SelectedTab = (TabPage)c;
                   // tabControl3.SelectedIndex = 2;
                
                }

                t.Focus();
                t.SelectAll();
                return false;
            }
        
        }



    
      

        private void button11_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = openFileDialog1.FileName;

                FileInfo tmp = new System.IO.FileInfo(FileName);
                tmp.CopyTo(@"\\acmesrv01\Public\Users\TerryLee\" + tmp.Name, true);

            }
        }

      


        public DataTable GetACME_TASK_File(string PrjCode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT 'N' as flag,ID,PrjCode,FileName FROM ACME_TASK_File WHERE PrjCode=@PrjCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }


        public DataTable GetExpense(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT T0.[RefDate], T1.[Account], T2.[AcctName],T1.[Debit]- T1.[Credit] [Debit], T1.[LineMemo] ");
            sb.Append("FROM OJDT T0 ");
            sb.Append("INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("INNER JOIN OACT T2 ON T1.Account = T2.AcctCode ");
            sb.Append("WHERE T1.[Project] =@PrjCode   AND T1.[Account] NOT IN ('12500108','52200101')  AND SUBSTRING(T1.ACCOUNT,1,1) in (5,6) ");
            sb.Append(" order by T0.RefDate ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        public DataTable GetExpense1(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(T1.[Debit]- T1.[Credit]),0) Debit  FROM OJDT T0");
            sb.Append( " INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" WHERE T1.[Project] =@PrjCode and T0.TransType=30 AND T1.[Account] NOT IN ('12500108','52200101')  AND SUBSTRING(T1.ACCOUNT,1,1) in (5,6) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        public DataTable GetRevenue(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[RefDate], T1.[Account], T2.[AcctName],T1.[Credit]-Debit [Credit], T1.[LineMemo]  ");
            sb.Append(" FROM OJDT T0  ");
            sb.Append(" INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append(" INNER JOIN OACT T2 ON T1.Account = T2.AcctCode  ");
            sb.Append(" WHERE T1.[Project] =@PrjCode");
      //     sb.Append(" AND   T1.[Account] IN ('42100101','22610103')  ");
            sb.Append(" AND   T1.[Account] = ('42100101')  ");
            sb.Append(" order by T0.RefDate  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        public DataTable GetRevenueN(string PrjCode, string YEAR, int MONTH)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("        SELECT SUM(T1) T1 FROM (       SELECT ISNULL(SUM(T1.[Credit]-T1.[Debit]),0) T1 FROM OJDT T0  ");
            sb.Append("               INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               WHERE T1.[Project] =@PrjCode  ");
            sb.Append("               AND   T1.[Account] = ('42100101') AND MONTH(T0.[RefDate])=@MONTH AND YEAR(T0.[RefDate])=@YEAR AND MEMO NOT LIKE '%作廢%'  ");
            sb.Append("                         UNION ALL   ");
            sb.Append("                             SELECT ISNULL(SUM(T1.[Credit]),0) T1 FROM OJDT T0   ");
            sb.Append("                             INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId   ");
            sb.Append("                             WHERE T1.[Project] =@PrjCode AND T0.TRANSID IN (259122,307116) ");
            sb.Append(" AND MONTH(T0.[RefDate])=@MONTH AND YEAR(T0.[RefDate])=@YEAR  ) AS A");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }

        public DataTable GetRevenueN2(string PrjCode, int YEAR)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("        SELECT SUM(T1) T1 FROM (       SELECT ISNULL(SUM(T1.[Credit]-T1.[Debit]),0) T1 FROM OJDT T0  ");
            sb.Append("               INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               WHERE T1.[Project] =@PrjCode  ");
            sb.Append("               AND   T1.[Account] = ('42100101') AND YEAR(T0.[RefDate])=@YEAR AND CONVERT(VARCHAR(8),T0.[RefDate] ,112)<=@DATE AND MEMO NOT LIKE '%作廢%'  ");
            sb.Append("                         UNION ALL   ");
            sb.Append("                             SELECT ISNULL(SUM(T1.[Credit]),0) T1 FROM OJDT T0   ");
            sb.Append("                             INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId   ");
            sb.Append("                             WHERE T1.[Project] =@PrjCode AND CONVERT(VARCHAR(8),T0.[RefDate] ,112)<=@DATE AND T0.TRANSID IN (259122,307116) ");
            sb.Append("  AND YEAR(T0.[RefDate])=@YEAR  ) AS A");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@DATE", textBox12.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        public DataTable GetAsset(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT T0.[RefDate], T1.[Account], T2.[AcctName], T1.[Debit], T1.[Credit], T1.[LineMemo] ");
            sb.Append("FROM OJDT T0 ");
            sb.Append("INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("INNER JOIN OACT T2 ON T1.Account = T2.AcctCode ");
            sb.Append("WHERE T1.[Project] =@PrjCode ");
            sb.Append("AND   T1.[Account] ='15410101' ");

            sb.Append(" order by T0.RefDate ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }

        public DataTable GetAsset1(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(T1.[Debit]),0) FROM OJDT T0  ");
            sb.Append(" INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" WHERE T1.[Project] =@PrjCode ");
            sb.Append(" AND   T1.[Account] ='15410101' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
        public DataTable GetAsset2(string PrjCode)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            //sb.Append("          SELECT '' RefDate,'' [Account],'' [AcctName],AMOUNT Debit,MEMO LineMemo FROM ACMESQLSP.DBO.ESCO_DISC WHERE ACCOUNT=@PrjCode");
            //sb.Append(" UNION ALL");
            sb.Append("              SELECT   CONVERT(VARCHAR(8),T0.[RefDate] ,112) RefDate, T1.[Account], T2.[AcctName], T1.[Debit], T1.[LineMemo]  ");
            sb.Append("              FROM OJDT T0  ");
            sb.Append("              INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("              INNER JOIN OACT T2 ON T1.Account = T2.AcctCode  ");
            sb.Append("              WHERE T1.[Project] =@PrjCode  ");
            sb.Append("              AND   T1.[Account] ='52200101'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }
      
        private void gvProject_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

       
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //focus 移走後,就沒有作用 
            //Graphics g = tabControl1.CreateGraphics();
            //Rectangle rect = new Rectangle(tabControl1.SelectedIndex * tabControl1.ItemSize.Width + 2, 2, tabControl1.ItemSize.Width - 2, tabControl1.ItemSize.Height - 2);
            //g.FillRectangle(Brushes.LightBlue, rect); g.DrawString(tabControl1.SelectedTab.Text, new Font(tabControl1.SelectedTab.Font, FontStyle.Bold), Brushes.Black, rect);

        }



       



        private void button15_Click_1(object sender, EventArgs e)
        {
            //if (MessageBox.Show("是否離開程式 ?", "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //{
                Close();

            //}
            
        }




        private void fmAcmeSolar_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("是否離開程式 ?", "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }

        }




        private void gvOPOR_SelectionChanged(object sender, EventArgs e)
        {
            string PRJ = gvProject.CurrentRow.Cells[0].Value.ToString();
            string DocEntry = gvOPOR.CurrentRow.Cells[0].Value.ToString();
            gvPOR1.DataSource = GetPOR1(DocEntry, PRJ);
        }



        public void SetControlEnabled(System.Windows.Forms.Control.ControlCollection originalControls, bool EnabledFlag)
        {
            // Control aControl;
            bool anti;
            if (EnabledFlag == true)
            {
                anti = false;
            }
            else
            {
                anti = true;
            }
            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlEnabled(originalControls[i].Controls, EnabledFlag);
                }

                if (originalControls[i] is TextBox)
                {


                    TextBox aTextBox = (TextBox)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;


                    aTextBox.ReadOnly = anti;
                    //修改 Enabled 的顏色
                    if (EnabledFlag)
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                    }
                    else
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                        //  MessageBox.Show("");

                    }
                    // aTextBox.ReadOnly = ! Enabled;
                }


                if (originalControls[i] is CheckBox)
                {


                    CheckBox aTextBox = (CheckBox)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;


                    aTextBox.Enabled = EnabledFlag;

                }
                if (originalControls[i] is Button)
                {


                    Button aTextBox = (Button)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;


                    aTextBox.Enabled = EnabledFlag;
                    //修改 Enabled 的顏色

                    // aTextBox.ReadOnly = ! Enabled;
                }
                if (originalControls[i] is ComboBox)
                {


                    ComboBox aTextBox = (ComboBox)originalControls[i];
                    //DropDownList 才會顏色變對
                    aTextBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    aTextBox.Enabled = EnabledFlag;
                    //DropDownList 才會顏色變對
                    //   aTextBox.r = anti;
                    //修改 Enabled 的顏色
                    if (EnabledFlag)
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                    }
                    else
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                        //  MessageBox.Show("");

                    }
                    // aTextBox.ReadOnly = ! Enabled;
                }


                //if (originalControls[i] is DataGridView)
                //{


                //    DataGridView DataGridView = (DataGridView)originalControls[i];

                //}

                if (originalControls[i] is DateTimePicker)
                {


                    DateTimePicker a = (DateTimePicker)originalControls[i];

                    a.Enabled = EnabledFlag;

                }
            }
        }



      
        private void tabPage12_Paint(object sender, PaintEventArgs e)
        {
            //RepaintControls(sender as TabPage);
        }

        System.Drawing.Color PageStartColor = Color.White;
        System.Drawing.Color PageEndColor = Color.CadetBlue;

        private void RepaintControls(TabControl tc)
        {
            foreach (TabPage ctl in tc.TabPages)
            {
                System.Drawing.Drawing2D.LinearGradientBrush gradBrush;
                gradBrush = new System.Drawing.Drawing2D.LinearGradientBrush(new Point(0, 0),
                new Point(ctl.Width, ctl.Height), PageStartColor, PageEndColor);

                Bitmap bmp = new Bitmap(ctl.Width, ctl.Height);

                Graphics g = Graphics.FromImage(bmp);
                g.FillRectangle(gradBrush, new Rectangle(0, 0, ctl.Width, ctl.Height));
                ctl.BackgroundImage = bmp;
                ctl.BackgroundImageLayout = ImageLayout.Stretch;
            }

        }

        private void tabPage12_Resize(object sender, EventArgs e)
        {
            //RepaintControls(sender as TabPage);
        }

        private void tabControl3_Resize(object sender, EventArgs e)
        {
            RepaintControls(sender as TabControl);
        }

        //this.s_PID.Text = this.dataGridViewL[1, this.dataGridViewL.CurrentCell.RowIndex].Value.ToString();


        public void ChangeLabel(System.Windows.Forms.Control.ControlCollection originalControls)
        {

            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    ChangeLabel(originalControls[i].Controls);
                }

                if (originalControls[i] is Label)
                {


                    Label a = (Label)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;
                    a.BackColor = Color.Transparent;
                }



            }
        }

    

        private void tabControl4_Resize(object sender, EventArgs e)
        {
            RepaintControls(sender as TabControl);
        }

   

    
     
        /// <summary>
        /// 取得執行檔位置 
        /// </summary>
        /// <returns></returns>
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }

        private void gvRevenue_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private void gvRdr1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 6);
            }
        }

        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("代碼", typeof(string));
            dt.Columns.Add("名稱", typeof(string));
            dt.Columns.Add("負責業務", typeof(string));
            dt.Columns.Add("預估分享", typeof(double));
            dt.Columns.Add("已收帳款", typeof(double));
            dt.Columns.Add("調整後預估分享", typeof(double));
            dt.Columns.Add("服務總金額", typeof(double));
            dt.Columns.Add("成本", typeof(double));
            dt.Columns.Add("己收期數", typeof(string));
            dt.Columns.Add("應收期數", typeof(string));
            dt.Columns.Add("合約生效日", typeof(string));
            dt.Columns.Add("合約終止日", typeof(string));
            dt.Columns.Add("專案階段", typeof(string));
            //專案階段
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
             int n;
             if (!int.TryParse(textBox1.Text, out n))
             {
                 MessageBox.Show("請輸入數字");
                 return;
             }

             DELETEESCO(textBoxPrjCode.Text);

             AddAUOGD61(textBoxPrjCode.Text, Convert.ToInt32(textBox1.Text));

             MessageBox.Show("更新完成");
        }
        public void AddAUOGD61(string PROJECT, int  SHARE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_PROJECT(PROJECT,SHARE) values(@PROJECT,@SHARE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@SHARE", SHARE));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        public void AddAUOGD612(string PROJECT, int SHARE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_PROJECT2(PROJECT,SHARE) values(@PROJECT,@SHARE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@SHARE", SHARE));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void AddAUOGD613(string PROJECT, string STAGE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_PROJECT3(PROJECT,STAGE) values(@PROJECT,@STAGE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));
            command.Parameters.Add(new SqlParameter("@STAGE", STAGE));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void DELETEESCO(string PROJECT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ESCO_PROJECT WHERE PROJECT=@PROJECT", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void DELETEESCO2(string PROJECT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ESCO_PROJECT2 WHERE PROJECT=@PROJECT", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void DELETEESCO3(string PROJECT)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ESCO_PROJECT3 WHERE PROJECT=@PROJECT", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PROJECT", PROJECT));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ESCO\\MONTH.xls";
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = "";

            OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-每月結算收款" + ".xls";

            ExcelReport.ExcelReportOutput(GetMONORDER(), ExcelTemplate, OutPutFile, "N");
        }
        public DataTable GetMONORDER()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("               SELECT  RANK() OVER (ORDER BY T4.DOCENTRY DESC) AS 項目,CAST(YEAR(T0.DOCDATE)-1911 AS VARCHAR)+'年'+CAST(MONTH(T0.DOCDATE) AS VARCHAR)+'月請款 ('+CAST(MONTH(DATEADD(MONTH,-1,T0.DOCDATE)) AS VARCHAR)+'月份計價)' MM,T0.PROJECT 專案號碼 ,T2.PRJNAME 專案名稱,''''+CARDCODE 客戶代碼,T0.CARDNAME 客戶名稱, ");
            sb.Append("               T4.DOCENTRY 銷售訂單,T1.LINETOTAL 本期金額,T3.LINETOTAL 累計金額 ");
            sb.Append("                FROM OINV T0  ");
            sb.Append("               LEFT JOIN INV1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("               LEFT JOIN OPRJ T2 ON (T0.PROJECT=T2.PRJCODE)  ");
            sb.Append("               left join RDR1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='17') ");
            sb.Append("               LEFT JOIN (SELECT T1.DOCENTRY,SUM(LINETOTAL) LINETOTAL FROM RDR1 T0 LEFT JOIN ORDR T1 ON (T0.DOCENTRY=T1.DOCENTRY) ");
            sb.Append("               WHERE    T0.LINESTATUS='C' AND T0.DSCRIPTION NOT LIKE '%工程款項明%'    GROUP BY T1.DOCENTRY) ");
            sb.Append("               T3 ON (T4.DOCENTRY=T3.DOCENTRY) ");
            sb.Append("               WHERE   T0.PROJECT  IN (SELECT O.PrjCode FROM OPRJ O    ");
            sb.Append("               WHERE  Substring(O.PrjCode,1,1)='4' )  AND  CONVERT(VARCHAR(6),T0.DOCDATE ,112) ='" + textBox9.Text + "'  AND T1.DSCRIPTION NOT LIKE '%工程款項明%'   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
           
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_TASK_File");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_TASK_File"];
        }

        private void button4_Click(object sender, EventArgs e)
        {

            int n;
            if (!int.TryParse(textBox10.Text, out n))
            {
                MessageBox.Show("請輸入數字");
                return;
            }

            DELETEESCO2(textBoxPrjCode.Text);

            AddAUOGD612(textBoxPrjCode.Text, Convert.ToInt32(textBox10.Text));

            MessageBox.Show("更新完成");
        }
        private System.Data.DataTable MakeTableN(int EndMon)
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("代碼", typeof(string));
            dt.Columns.Add("名稱", typeof(string));
            dt.Columns.Add("負責業務", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("AWS編號", typeof(string));
            for (int i = 1; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }
            dt.Columns.Add("Total", typeof(Int64));


            dt.Columns.Add("續約備註", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableN2(int EndMon)
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("代碼", typeof(string));
            dt.Columns.Add("名稱", typeof(string));
            dt.Columns.Add("負責業務", typeof(string));
            dt.Columns.Add("客戶代碼", typeof(string));

            for (int i = 2013; i <= EndMon; i++)
            {
                dt.Columns.Add(i.ToString("00"), typeof(Int64));
            }
            dt.Columns.Add("Total", typeof(Int64));


            dt.Columns.Add("續約備註", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            button5f();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

        private void button5f()
        {
            dataGridView2.Columns.Clear();
            System.Data.DataTable dtCost = null;
            int iMon2 = 12;
            if (cbYear.Text == DateTime.Now.ToString("yyyy"))
            {
                iMon2 = DateTime.Now.Month;
            }
            DataTable dt = GetProject();
            DataRow dr = null;
            dtCost = MakeTableN(iMon2);
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                string PRJCODE = dt.Rows[i]["PrjCode"].ToString();

                dr["代碼"] = PRJCODE;
                dr["名稱"] = dt.Rows[i]["PrjName"].ToString();
                dr["備註"] = dt.Rows[i]["MEMO"].ToString();
                dr["續約備註"] = dt.Rows[i]["MEMO2"].ToString();
                //AWS編號 
                System.Data.DataTable G1S = GetORDR_Project(PRJCODE);
                if (G1S.Rows.Count > 0)
                {
                    dr["客戶代碼"] = G1S.Rows[0]["CARDCODE"].ToString();
                    dr["負責業務"] = G1S.Rows[0]["SLPNAME"].ToString();
                    dr["AWS編號"] = G1S.Rows[0]["AWS編號"].ToString();
                }
                Int64 Total = 0;

                for (int j = 1; j <= iMon2; j++)
                {
                    Int64 h1 = Convert.ToInt32(GetRevenueN(PRJCODE, cbYear.Text, j).Rows[0][0]);
                    //if (PRJCODE == "41609012")
                    //{
                    //    if (cbYear.Text == "2019")
                    //    {
                    //        if (j == 8)
                    //        {
                    //            h1 =  80866;
                    //        }
                    //    }
                    //}
                    dr[j.ToString("00")] = h1;
                    Total += Convert.ToInt64(h1);
                }
                dr["Total"] = Total;
                dtCost.Rows.Add(dr);


            }
            int MM = iMon2 + 5;
            decimal[] Totalf = new decimal[MM];

            for (int s = 0; s <= dtCost.Rows.Count - 1; s++)
            {

                for (int j = 5; j <= MM; j++)
                {
                    Totalf[j - 1] += Convert.ToDecimal(dtCost.Rows[s][j]);

                }
            }

            DataRow row;

            row = dtCost.NewRow();

            row[3] = "合計";
            for (int js = 4; js <= MM; js++)
            {
                row[js] = Totalf[js - 1];

            }
            dtCost.Rows.Add(row);
            dataGridView2.DataSource = dtCost;

            for (int i = 5; i <= MM; i++)
            {
                DataGridViewColumn col = dataGridView2.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void button6f()
        {
            dataGridView2.Columns.Clear();
            System.Data.DataTable dtCost = null;
            int iMon2 = Convert.ToInt32(textBox12.Text.Substring(0, 4));
            DataTable dt = GetProject();
            DataRow dr = null;
            dtCost = MakeTableN2(iMon2);
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                string PRJCODE = dt.Rows[i]["PrjCode"].ToString();

                dr["代碼"] = PRJCODE;
                dr["名稱"] = dt.Rows[i]["PrjName"].ToString();
                dr["備註"] = dt.Rows[i]["MEMO"].ToString();
                dr["續約備註"] = dt.Rows[i]["MEMO2"].ToString();
                //續約備註
                System.Data.DataTable G1S = GetORDR_Project(PRJCODE);
                if (G1S.Rows.Count > 0)
                {
                    dr["客戶代碼"] = G1S.Rows[0]["CARDCODE"].ToString();
                    dr["負責業務"] = G1S.Rows[0]["SLPNAME"].ToString();
                }
                Int64 Total = 0;

                for (int j = 2013; j <= iMon2; j++)
                {
                    Int64 h1 = Convert.ToInt32(GetRevenueN2(PRJCODE, j).Rows[0][0]);
                    dr[j.ToString("00")] = h1;
                    Total += Convert.ToInt64(h1);
                }
                dr["Total"] = Total;
                dtCost.Rows.Add(dr);


            }
            int MM = iMon2 + 4 - 2013 + 1;
            decimal[] Totalf = new decimal[MM];

            for (int s = 0; s <= dtCost.Rows.Count - 1; s++)
            {

                for (int j = 4; j <= MM; j++)
                {
                    Totalf[j - 1] += Convert.ToDecimal(dtCost.Rows[s][j]);

                }
            }

            DataRow row;

            row = dtCost.NewRow();

            row[3] = "合計";
            for (int js = 4; js <= MM; js++)
            {
                row[js] = Totalf[js - 1];

            }
            dtCost.Rows.Add(row);
            dataGridView3.DataSource = dtCost;

            for (int i = 4; i <= MM; i++)
            {
                DataGridViewColumn col = dataGridView3.Columns[i];


                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {


            DELETEESCO3(textBoxPrjCode.Text);

            AddAUOGD613(textBoxPrjCode.Text, currentStage.Text);

            MessageBox.Show("更新完成");
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }

            if (e.RowIndex >= dataGridView1.Rows.Count - 1)
                return;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                WriteExcelAP2(openFileDialog1.FileName);
                button5f();
            }
        }

        private void WriteExcelAP2(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string PRJCODE;
                string MEMO;
                string MEMO2;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    PRJCODE = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    MEMO2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();

                    if (PRJCODE != "代碼")
                    {

                        if (!String.IsNullOrEmpty(PRJCODE) && !String.IsNullOrEmpty(MEMO))
                        {
                            UPDATEMEMO(MEMO, PRJCODE);
                        }

                        if (!String.IsNullOrEmpty(PRJCODE) && !String.IsNullOrEmpty(MEMO2))
                        {
                            UPDATEMEMO2(MEMO2, PRJCODE);
                        }
                    }
                }




            }
            finally
            {



   
                //Quit
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();



            }



        }

        public void UPDATEMEMO(string U_MEMO, string PrjCode)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OPRJ SET U_MEMO=@U_MEMO WHERE PrjCode=@PrjCode", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void UPDATEMEMO2(string U_MEMO2, string PrjCode)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OPRJ SET U_MEMO2=@U_MEMO2 WHERE PrjCode=@PrjCode", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_MEMO2", U_MEMO2));
            command.Parameters.Add(new SqlParameter("@PrjCode", PrjCode));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        private static System.Data.DataTable Year()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM acmesqlsp.dbo.RMA_PARAMS where param_kind='shipyear' and PARAM_NO >2012 order by DataValue desc";

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
        private void button11_Click_1(object sender, EventArgs e)
        {
            int L1 = textBox12.Text.Length;
            if (L1 != 8)
            {
                return;
            }
             int n;
             if (int.TryParse(textBox12.Text, out n))
             {
                 button6f();
             }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView3);
        }

  

      
    }//
}//


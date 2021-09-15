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
namespace ACME
{
    public partial class RMASAPSEARCH : Form
    {
        public RMASAPSEARCH()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox7.Text == "")
            {
                MessageBox.Show("RMA管制表請輸入年份");
                return;
            }
            if (textBox1.Text == "")
            {
                MessageBox.Show("RMA未結案明細請輸入年份");
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("出貨Vender超過2周未還回請輸入收貨Vender日");
                return;
            }
            if (textBox3.Text == "")
            {
                MessageBox.Show("出貨Vender超過7天未收到Receiving notice請輸入出貨Vender日");
                return;
            }
            if (textBox4.Text == "")
            {
                MessageBox.Show("客戶RMA#OPEN 超過2週未寄回請輸入ACME退運通知日");
                return;
            }
            if (textBox5.Text == "")
            {
                MessageBox.Show("收到Vender還貨超過3天未還回客戶請輸入Vender還貨日");
                return;
            }
            if (textBox6.Text == "")
            {
                MessageBox.Show("收到客戶RMA Panel超過7天未出貨Vender請輸入ACME收貨日");
                return;
            }
            FF2();
            dataGridView2.DataSource = GetRMA1(textBox1.Text);
            dataGridView3.DataSource = GetRMA2(textBox2.Text);
            dataGridView4.DataSource = GetRMA3(textBox3.Text);
            dataGridView5.DataSource = GetRMA4();
            dataGridView6.DataSource = GetRMA5(textBox4.Text);
            dataGridView7.DataSource = GetRMA6(textBox5.Text);
            dataGridView8.DataSource = GetRMA7(textBox6.Text);


           
        }
        private System.Data.DataTable MakeTable()
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("契約號碼", typeof(string));
            dt.Columns.Add("單據種類", typeof(string));
            dt.Columns.Add("RMAStatus", typeof(string));
            dt.Columns.Add("RMANO", typeof(string));
            dt.Columns.Add("客戶簡稱", typeof(string));
            dt.Columns.Add("Vender", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("Model", typeof(string));
            dt.Columns.Add("Ver", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Qty", typeof(string));
            dt.Columns.Add("ACME通知退運日", typeof(string));
            dt.Columns.Add("退運倉別", typeof(string));
            dt.Columns.Add("ACME收貨日", typeof(string));
            dt.Columns.Add("VenderRMANo", typeof(string));
            dt.Columns.Add("RepairCenter", typeof(string));
            dt.Columns.Add("出貨Vender日", typeof(string));
            dt.Columns.Add("Vender未還數量", typeof(string));
            dt.Columns.Add("Vender收貨日", typeof(string));
            dt.Columns.Add("ACME還貨數量", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("Engineer", typeof(string));
            dt.Columns.Add("Sales", typeof(string));
      
            return dt;
        }

        private void FF2()
        {

            System.Data.DataTable dt = GetRMA0(textBox7.Text);
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTable();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                dr = dtCost.NewRow();


                dr["契約號碼"] = dt.Rows[i]["契約號碼"].ToString();
                dr["單據種類"] = dt.Rows[i]["單據種類"].ToString();
                dr["RMAStatus"] = dt.Rows[i]["RMAStatus"].ToString();
                dr["RMANO"] = dt.Rows[i]["RMANO"].ToString();
                dr["客戶簡稱"] = dt.Rows[i]["客戶簡稱"].ToString();
                dr["Vender"] = dt.Rows[i]["Vender"].ToString();
                string MODEL = dt.Rows[i]["Model"].ToString();
                string MODEL2 = dt.Rows[i]["MODEL2"].ToString();
                System.Data.DataTable GM1 = GetRMA0S(MODEL2);
                if (GM1.Rows.Count > 0)
                {
                    dr["BU"] = GM1.Rows[0][0].ToString();
                }
                dr["Model"] = MODEL;
                dr["Ver"] = dt.Rows[i]["Ver"].ToString();
                dr["Grade"] = dt.Rows[i]["Grade"].ToString();
                dr["Qty"] = dt.Rows[i]["Qty"].ToString();
                dr["ACME通知退運日"] = dt.Rows[i]["ACME通知退運日"].ToString();
                dr["退運倉別"] = dt.Rows[i]["退運倉別"].ToString();
                dr["ACME收貨日"] = dt.Rows[i]["ACME收貨日"].ToString();
                dr["VenderRMANo"] = dt.Rows[i]["VenderRMANo"].ToString();
                dr["RepairCenter"] = dt.Rows[i]["RepairCenter"].ToString();
                dr["出貨Vender日"] = dt.Rows[i]["出貨Vender日"].ToString();
                dr["Vender未還數量"] = dt.Rows[i]["Vender未還數量"].ToString();
                dr["Vender收貨日"] = dt.Rows[i]["Vender收貨日"].ToString();
                dr["ACME還貨數量"] = dt.Rows[i]["ACME還貨數量"].ToString();
                dr["備註"] = dt.Rows[i]["備註"].ToString();
                dr["Engineer"] = dt.Rows[i]["Engineer"].ToString();
                dr["Sales"] = dt.Rows[i]["Sales"].ToString();
                dtCost.Rows.Add(dr);
            }
            dataGridView1.DataSource = dtCost;
        }
        private System.Data.DataTable GetRMA0(string U_RMAYEAR)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T0.[ContractID] 契約號碼,DESCR 單據種類,T0.U_FMA RMAStatus,T0.[U_RMA_NO] RMANO,  ");
            sb.Append(" T0.[U_Cusname_S] 客戶簡稱,T0.[U_RVender] Vender,T0.[U_RModel] Model,T0.[U_RVer] Ver,T0.[U_RGrade] Grade ");
            sb.Append(" ,T0.[U_Rquinity] Qty,T0.[U_Racmetodate] ACME通知退運日,T0.[U_Routwharehouse] 退運倉別 ");
            sb.Append(" ,T0.[U_RtoReceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_RepairCenter] RepairCenter,  ");
            sb.Append(" T0.[U_ACME_Out] 出貨Vender日,T0.[U_yetqty] Vender未還數量,T0.[U_ACME_recedate] Vender收貨日 ");
            sb.Append(" ,T0.[U_ACME_BackDate]  Vender還貨日,T0.U_ACME_QBACK Vender已還數量,T0.[U_ACME_BackDate1] ACME還貨日 ");
            sb.Append(" ,T0.[U_ACME_BackQty1]  ACME還貨數量,cast(T0.[Remarks2] as varchar(250)) 備註,T0.[U_REngineer] Engineer,T0.[U_RSales] Sales");
            sb.Append(" ,substring(T0.[U_RModel],0,CASE CHARINDEX('_', T0.[U_RModel]) WHEN 0 THEN 100 ELSE CHARINDEX('_', T0.[U_RModel]) END)  MODEL2 FROM ACMESQL02.DBO.OCTR T0  ");
            sb.Append(" LEFT JOIN UFD1 T1 ON (T0.[u_pkind]=T1.FLDVALUE AND TABLEID='OCTR' AND FIELDID=21) ");
            sb.Append(" WHERE  T0.[U_RMA_NO] <> ''  AND substring(T0.[U_RMA_NO],1,1) = 'A'   ");
            sb.Append("  AND T0.[U_RMAYEAR]=@U_RMAYEAR and");
            sb.Append(" (( T0.[U_RMA_NO] = @U_RMA_NO)  or  (@U_RMA_NO=''))  ");
            sb.Append(" and  ((T0.[U_Cusname_S] = @U_Cusname_S)  or  (@U_Cusname_S=''))  ");
            sb.Append(" and  ((T0.[U_RVender]  = @U_RVender)  or (@U_RVender=''))  ");
            sb.Append(" and  ((T0.[U_RModel] = @U_RModel)  or (@U_RModel=''))  ");
            sb.Append(" and  ((T0.[U_RVer] = @U_RVer)  or (@U_RVer='')) ");
            sb.Append(" and  ((T0.[U_RGrade] = @U_RGrade)  or (@U_RGrade='')) ");
            sb.Append(" and  ((T0.[U_Rquinity] =@Rquinity)  or (@Rquinity='') ) ");
            sb.Append(" and  ((T0.[U_Racmetodate] = @U_Racmetodate)  or (@U_Racmetodate='')) ");
            sb.Append(" and  ((T0.[U_Routwharehouse] = @U_Routwharehouse)  or (@U_Routwharehouse='')) ");
            sb.Append(" and  ((T0.[U_RtoReceiving] =@U_RtoReceiving)  or (@U_RtoReceiving='')) ");
            sb.Append(" and  ((T0.[U_AUO_RMA_NO] = @U_AUO_RMA_NO ) or (@U_AUO_RMA_NO=''))  ");
            sb.Append(" and  ((T0.[U_RepairCenter] = @U_RepairCenter)  or (@U_RepairCenter='')) ");
            sb.Append(" and  ((T0.[U_ACME_Out] = @U_ACME_Out)  or (@U_ACME_Out='') ) ");
            sb.Append(" and  ((T0.[U_yetqty] =@U_yetqty)  or (@U_yetqty=''))");
            sb.Append(" and  ((T0.[U_ACME_recedate] = @U_ACME_recedate)  or (@U_ACME_recedate=''))  ");
            sb.Append(" and  ((T0.[U_ACME_BackDate] = @U_ACME_BackDate)  or  (@U_ACME_BackDate='') ) ");
            sb.Append(" and  ((T0.[U_ACME_BackDate1] = @U_ACME_BackDate1)  or (@U_ACME_BackDate1='') ) ");
            sb.Append(" and  ((T0.[U_ACME_BackQty1] = @U_ACME_BackQty1)  or (@U_ACME_BackQty1='')) ");
            sb.Append(" and  ((T0.[U_REngineer] = @U_REngineer)  or (@U_REngineer=''))  ");
            sb.Append(" and  ((T0.[U_RSales] = @U_RSales)  or  (@U_RSales='')) ");
            sb.Append(" AND (( T0.[Remarks2] LIKE  '%" + textBoxRemarks2.Text.ToString() + "%')  or  ('" + textBoxRemarks2.Text.ToString() + "'='')) and T0.[U_PKind] <> '1'");
            sb.Append(" UNION ALL ");
            sb.Append("  SELECT  T0.[ContractID] 契約號碼,DESCR 單據種類,T0.U_FMA RMAStatus,T0.[U_RMA_NO] RMANO, ");
            sb.Append(" T0.[U_Cusname_S] 客戶簡稱,T0.[U_RVender] Vender,T0.[U_RModel] Model,T0.[U_RVer] Ver,T0.[U_RGrade] Grade");
            sb.Append(" ,T0.[U_Rquinity] Qty,T0.[U_Racmetodate] ACME通知退運日,T0.[U_Routwharehouse] 退運倉別");
            sb.Append(" ,T0.[U_RtoReceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_RepairCenter] RepairCenter, ");
            sb.Append(" T0.[U_ACME_Out] 出貨Vender日,T0.[U_yetqty] Vender未還數量,T0.[U_ACME_recedate] Vender收貨日");
            sb.Append(" ,T0.[U_ACME_BackDate]  Vender還貨日,T0.U_ACME_QBACK Vender已還數量,T0.[U_ACME_BackDate1] ACME還貨日");
            sb.Append(" ,T0.[U_ACME_BackQty1]  ACME還貨數量,cast(T0.[Remarks2] as varchar(250)) 備註,T0.[U_REngineer] Engineer,T0.[U_RSales] Sales,substring(T0.[U_RModel],0,CASE CHARINDEX('_', T0.[U_RModel]) WHEN 0 THEN 100 ELSE CHARINDEX('_', T0.[U_RModel]) END)  MODEL2 FROM ACMESQL05.DBO.OCTR T0 ");
            sb.Append(" LEFT JOIN UFD1 T1 ON (T0.[u_pkind]=T1.FLDVALUE AND TABLEID='OCTR' AND FIELDID=21)");
            sb.Append("  WHERE  T0.[U_RMA_NO] <> ''  ");
            sb.Append("  AND T0.[U_RMAYEAR]=@U_RMAYEAR and");
            sb.Append(" (( T0.[U_RMA_NO] = @U_RMA_NO)  or  (@U_RMA_NO=''))  ");
            sb.Append(" and  ((T0.[U_Cusname_S] = @U_Cusname_S)  or  (@U_Cusname_S=''))  ");
            sb.Append(" and  ((T0.[U_RVender]  = @U_RVender)  or (@U_RVender=''))  ");
            sb.Append(" and  ((T0.[U_RModel] = @U_RModel)  or (@U_RModel=''))  ");
            sb.Append(" and  ((T0.[U_RVer] = @U_RVer)  or (@U_RVer='')) ");
            sb.Append(" and  ((T0.[U_RGrade] = @U_RGrade)  or (@U_RGrade='')) ");
            sb.Append(" and  ((T0.[U_Rquinity] =@Rquinity)  or (@Rquinity='') ) ");
            sb.Append(" and  ((T0.[U_Racmetodate] = @U_Racmetodate)  or (@U_Racmetodate='')) ");
            sb.Append(" and  ((T0.[U_Routwharehouse] = @U_Routwharehouse)  or (@U_Routwharehouse='')) ");
            sb.Append(" and  ((T0.[U_RtoReceiving] =@U_RtoReceiving)  or (@U_RtoReceiving='')) ");
            sb.Append(" and  ((T0.[U_AUO_RMA_NO] = @U_AUO_RMA_NO ) or (@U_AUO_RMA_NO=''))  ");
            sb.Append(" and  ((T0.[U_RepairCenter] = @U_RepairCenter)  or (@U_RepairCenter='')) ");
            sb.Append(" and  ((T0.[U_ACME_Out] = @U_ACME_Out)  or (@U_ACME_Out='') ) ");
            sb.Append(" and  ((T0.[U_yetqty] =@U_yetqty)  or (@U_yetqty=''))");
            sb.Append(" and  ((T0.[U_ACME_recedate] = @U_ACME_recedate)  or (@U_ACME_recedate=''))  ");
            sb.Append(" and  ((T0.[U_ACME_BackDate] = @U_ACME_BackDate)  or  (@U_ACME_BackDate='') ) ");
            sb.Append(" and  ((T0.[U_ACME_BackDate1] = @U_ACME_BackDate1)  or (@U_ACME_BackDate1='') ) ");
            sb.Append(" and  ((T0.[U_ACME_BackQty1] = @U_ACME_BackQty1)  or (@U_ACME_BackQty1='')) ");
            sb.Append(" and  ((T0.[U_REngineer] = @U_REngineer)  or (@U_REngineer=''))  ");
            sb.Append(" and  ((T0.[U_RSales] = @U_RSales)  or  (@U_RSales='')) ");
            sb.Append(" AND (( T0.[Remarks2] LIKE  '%" + textBoxRemarks2.Text.ToString() + "%')  or  ('" + textBoxRemarks2.Text.ToString() + "'='')) and T0.[U_PKind] <> '1'");
            sb.Append(" ORDER BY U_RMA_NO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMAYEAR", U_RMAYEAR));
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", textBoxRMANO.Text));
            command.Parameters.Add(new SqlParameter("@U_Cusname_S", textBoxCUSTNAME.Text));
            command.Parameters.Add(new SqlParameter("@U_RVender", textBoxVender.Text));
            command.Parameters.Add(new SqlParameter("@U_RModel", textBoxModel.Text));
            command.Parameters.Add(new SqlParameter("@U_RVer", textBoxVer.Text));
            command.Parameters.Add(new SqlParameter("@U_RGrade", textBoxGrade.Text));
            command.Parameters.Add(new SqlParameter("@Rquinity", textBoxQty.Text));
            command.Parameters.Add(new SqlParameter("@U_Racmetodate", textBoxU_Racmetodate.Text));
            command.Parameters.Add(new SqlParameter("@U_Routwharehouse", textBoxU_Routwharehouse.Text));
            command.Parameters.Add(new SqlParameter("@U_RtoReceiving", textU_RtoReceiving.Text));
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", textAUO_RMA_NO.Text));
            command.Parameters.Add(new SqlParameter("@U_RepairCenter", textBoxRepairCenter.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_Out", textBoxU_ACME_Out.Text));
            command.Parameters.Add(new SqlParameter("@U_yetqty", textBoxU_yetqty.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_recedate", textBoxU_ACME_recedate.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate", textBoxU_ACME_BackDate.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate1", textBoxU_ACME_BackDate1.Text));
            command.Parameters.Add(new SqlParameter("@U_ACME_BackQty1", textBoxU_ACME_BackQty1.Text));
            command.Parameters.Add(new SqlParameter("@U_REngineer", textBoxEngineer.Text));
            command.Parameters.Add(new SqlParameter("@U_RSales", textBoxU_RSales.Text));
            //textBoxEngineer
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetRMA0S(string U_TMODEL)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  TOP 1 U_BU BU FROM OITM WHERE U_TMODEL =@U_TMODEL AND  ISNULL(U_BU,'') <> ''");
    
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_TMODEL", U_TMODEL));

            //textBoxEngineer
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA1(string U_RMAYEAR)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();



            sb.Append("               SELECT T0.[ContractID] 契約號碼,DESCR 單據種類,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RVender] Vender,T0.[U_RModel] Model,T0.[U_RVer] Ver,t0.U_rgrade Grade,t0.U_rquinity Qty,T0.[U_racmetodate] ACME通知退運日,T0.[U_routwharehouse] 退運倉別,T0.[U_rengineer] Engineer,T0.[U_rsales] Sales,T0.[U_rtoreceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_repaircenter] RepairCenter,T0.[U_acme_out] 出貨Vender日,T0.[U_yetqty] Vender未還數量, T0.[U_acme_recedate] Vender收貨日,T0.[U_acme_backdate] Vender還貨日,T0.[U_acme_qback] Vender已還數量,T0.[U_acme_backdate1] ACME還貨日,T0.[U_acme_backqty1] ACME還貨數量,remarks2 備註 ");
            sb.Append("               ,T0.U_RMAYEAR 年份 FROM ACMESQL02.DBO.OCTR T0 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.UFD1 T1 ON (T0.[u_pkind]=T1.FLDVALUE AND TABLEID='OCTR' AND FIELDID=21)");
            sb.Append("                WHERE    U_rma_no <> 'null' AND substring(T0.[U_RMA_NO],1,1) = 'A' and T0.[u_pkind] <> '1' ");
            sb.Append("               and  ((T0.[U_cusname_s]='' or T0.[U_cusname_s] is null) or (T0.[U_rmodel]='' or T0.[U_rmodel] is null) or (T0.[U_rver]='' or T0.[U_rver] is null) or (T0.[U_rgrade]='' or T0.[U_rgrade] is null) or (T0.[U_rquinity]='' or T0.[U_rquinity] is null) or ( T0.[U_racmetodate] is null) or (T0.[U_rengineer]='' or T0.[U_rengineer] is null) or (T0.[U_rsales]='' or T0.[U_rsales] is null)  ");
            sb.Append("                or (T0.[U_rtoreceiving] is null)  or (T0.[U_auo_rma_no]='' or T0.[U_auo_rma_no] is null)   or (T0.[U_repaircenter]='' or T0.[U_repaircenter] is null)     or ( T0.[U_acme_out] is  null) or (T0.[U_yetqty]='' or T0.[U_yetqty] is null) or ( T0.[U_acme_recedate] is null ) ");
            sb.Append("                or ( ISNULL(T0.[U_acme_qback],'') = '' ) or ( T0.[U_acme_backdate1] is null)  or ( T0.[U_acme_backqty1]='' or T0.[U_acme_backqty1] is null))  ");
            sb.Append("                AND T0.U_RMAYEAR=@U_RMAYEAR  ");
            sb.Append(" UNION ALL ");
            sb.Append("               SELECT T0.[ContractID] 契約號碼,DESCR 單據種類,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RVender] Vender,T0.[U_RModel] Model,T0.[U_RVer] Ver,t0.U_rgrade Grade,t0.U_rquinity Qty,T0.[U_racmetodate] ACME通知退運日,T0.[U_routwharehouse] 退運倉別,T0.[U_rengineer] Engineer,T0.[U_rsales] Sales,T0.[U_rtoreceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_repaircenter] RepairCenter,T0.[U_acme_out] 出貨Vender日,T0.[U_yetqty] Vender未還數量, T0.[U_acme_recedate] Vender收貨日,T0.[U_acme_backdate] Vender還貨日,T0.[U_acme_qback] Vender已還數量,T0.[U_acme_backdate1] ACME還貨日,T0.[U_acme_backqty1] ACME還貨數量,remarks2 備註 ");
            sb.Append("               ,T0.U_RMAYEAR 年份 FROM ACMESQL05.DBO.OCTR T0 ");
            sb.Append(" LEFT JOIN ACMESQL05.DBO.UFD1 T1 ON (T0.[u_pkind]=T1.FLDVALUE AND TABLEID='OCTR' AND FIELDID=21)");
            sb.Append("                WHERE    U_rma_no <> 'null'  ");
            sb.Append("               and  ((T0.[U_cusname_s]='' or T0.[U_cusname_s] is null) or (T0.[U_rmodel]='' or T0.[U_rmodel] is null) or (T0.[U_rver]='' or T0.[U_rver] is null) or (T0.[U_rgrade]='' or T0.[U_rgrade] is null) or (T0.[U_rquinity]='' or T0.[U_rquinity] is null) or ( T0.[U_racmetodate] is null) or (T0.[U_rengineer]='' or T0.[U_rengineer] is null) or (T0.[U_rsales]='' or T0.[U_rsales] is null)  ");
            sb.Append("                or (T0.[U_rtoreceiving] is null)  or (T0.[U_auo_rma_no]='' or T0.[U_auo_rma_no] is null)   or (T0.[U_repaircenter]='' or T0.[U_repaircenter] is null)     or ( T0.[U_acme_out] is  null) or (T0.[U_yetqty]='' or T0.[U_yetqty] is null) or ( T0.[U_acme_recedate] is null ) ");
            sb.Append("                or ( ISNULL(T0.[U_acme_qback],'') = '' ) or ( T0.[U_acme_backdate1] is null)  or ( T0.[U_acme_backqty1]='' or T0.[U_acme_backqty1] is null))  ");
            sb.Append("                AND T0.U_RMAYEAR=@U_RMAYEAR ORDER BY U_RMA_NO ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMAYEAR", U_RMAYEAR));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA2(string recedate)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model,T0.[U_RVer] Ver,T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty, ");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo, T0.[U_ACME_recedate] Vender收貨日, T0.[U_ACME_BackDate] Vender還貨日, T0.[U_ACME_QBACK] Vender已還數量,");
            sb.Append(" ISNULL(CASE WHEN U_YETQTY ='/' THEN 0 WHEN U_YETQTY LIKE '%+%' THEN 0 WHEN U_YETQTY LIKE '%.%' THEN 0 ELSE U_YETQTY END ,0)-CAST(ISNULL(CASE WHEN U_ACME_QBACK = '/' THEN 0 WHEN U_ACME_QBACK LIKE '%+%' THEN 0 ELSE U_ACME_QBACK END ,0) AS INT) Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註   FROM ACMESQL02.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append("  WHERE T0.[U_ACME_BackDate] is null ");
            sb.Append(" and  T0.[U_ACME_recedate] < cast(@recedate as datetime)-14 ");
            sb.Append("  AND substring(T0.[U_RMA_NO],1,1) = 'A' and T0.[TermDate]  is null and T0.[U_PKind] <> '1'");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT  distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model,T0.[U_RVer] Ver,T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty, ");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo, T0.[U_ACME_recedate] Vender收貨日, T0.[U_ACME_BackDate] Vender還貨日, T0.[U_ACME_QBACK] Vender已還數量,");
            sb.Append(" ISNULL(CASE WHEN U_YETQTY ='/' THEN 0 WHEN U_YETQTY LIKE '%+%' THEN 0 WHEN U_YETQTY LIKE '%.%' THEN 0 ELSE U_YETQTY END ,0)-CAST(ISNULL(CASE WHEN U_ACME_QBACK = '/' THEN 0 WHEN U_ACME_QBACK LIKE '%+%' THEN 0 ELSE U_ACME_QBACK END ,0) AS INT) Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註   FROM ACMESQL05.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append("  WHERE T0.[U_ACME_BackDate] is null ");
            sb.Append(" and  T0.[U_ACME_recedate] < cast(@recedate as datetime)-14 ");
            sb.Append("  and T0.[TermDate]  is null ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@recedate", recedate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA3(string Out)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model,T0.[U_RVer] Ver,");
            sb.Append(" T0.[U_RGrade] Grade,T0.[U_Rquinity] Qty,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_ACME_Out] 出貨Vender日,T0.[U_ACME_recedate] Vender收貨日");
            sb.Append(" ,T0.[U_yetqty] Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註");
            sb.Append("  FROM ACMESQL02.DBO.OCTR T0 INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID");
            sb.Append("  WHERE T0.[U_ACME_recedate]  IS NULL ");
            sb.Append(" AND   T0.[U_ACME_Out] < cast(@Out as datetime) -7 AND substring(T0.[U_RMA_NO],1,1) = 'A' ");
            sb.Append(" AND T0.[U_ACME_Out] <> '1900.01.01' and T0.[TermDate]  is null AND T0.[U_PKind] <> '1'");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model,T0.[U_RVer] Ver,");
            sb.Append(" T0.[U_RGrade] Grade,T0.[U_Rquinity] Qty,''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_ACME_Out] 出貨Vender日,T0.[U_ACME_recedate] Vender收貨日");
            sb.Append(" ,T0.[U_yetqty] Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註");
            sb.Append("  FROM ACMESQL05.DBO.OCTR T0 INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID");
            sb.Append("  WHERE T0.[U_ACME_recedate]  IS NULL ");
            sb.Append(" AND   T0.[U_ACME_Out] < cast(@Out as datetime) -7 ");
            sb.Append(" AND T0.[U_ACME_Out] <> '1900.01.01' and T0.[TermDate]  is null ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Out", Out));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA4()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱,");
            sb.Append(" T0.[U_RModel] Model, T0.[U_RVer] Ver,T0.[U_RGrade] Grade,T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日, ");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_ACME_BackDate] Vender還貨日,T0.[U_ACME_BackDate1], ");
            sb.Append(" T0.[U_ACME_BackQty1] ACME還貨日,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL02.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE  T0.[TermDate]  is null and T0.[U_PKind] <> '1'");
            sb.Append(" and (cast(isnull(");
            sb.Append(" case when T0.[U_Rquinity] like '%/%' then '0'");
            sb.Append(" when T0.[U_Rquinity] like '%+%' then '0' ");
            sb.Append(" when T0.[U_Rquinity] like '%.%' then '0' ");
            sb.Append(" when T0.[U_Rquinity] like '%#%' then '0' ");
            sb.Append(" else T0.[U_Rquinity] end,0) as int)  - cast(isnull(case ");
            sb.Append(" when T0.[U_ACME_BackQty1]  like '%/%' then 0 ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%+%' then '0' ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%.%' then '0' ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%#%' then '0' ");
            sb.Append(" else T0.[U_ACME_BackQty1] end ,0) as int) <> 0)");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%/%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%+%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%.%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%#%'");
            sb.Append(" and T0.[U_Rquinity] not like '%/%'");
            sb.Append(" and T0.[U_Rquinity] not like '%+%'");
            sb.Append(" and T0.[U_Rquinity] not like '%.%'");
            sb.Append(" and T0.[U_Rquinity] not like '%#%'");
            sb.Append("  and T0.[U_Rquinity] <> '/'");
            sb.Append(" and T0.[U_Cusname_S] <> 'AUO_2nd RMA' and T0.[U_ACME_BackDate1] <> '1900-01-01 00:00:00.000'");
            sb.Append("  AND substring(T0.[U_RMA_NO],1,1) = 'A'");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱,");
            sb.Append(" T0.[U_RModel] Model, T0.[U_RVer] Ver,T0.[U_RGrade] Grade,T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日, ");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.[U_ACME_BackDate] Vender還貨日,T0.[U_ACME_BackDate1], ");
            sb.Append(" T0.[U_ACME_BackQty1] ACME還貨日,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL05.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE  T0.[TermDate]  is null ");
            sb.Append(" and (cast(isnull(");
            sb.Append(" case when T0.[U_Rquinity] like '%/%' then '0'");
            sb.Append(" when T0.[U_Rquinity] like '%+%' then '0' ");
            sb.Append(" when T0.[U_Rquinity] like '%.%' then '0' ");
            sb.Append(" when T0.[U_Rquinity] like '%#%' then '0' ");
            sb.Append(" else T0.[U_Rquinity] end,0) as int)  - cast(isnull(case ");
            sb.Append(" when T0.[U_ACME_BackQty1]  like '%/%' then 0 ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%+%' then '0' ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%.%' then '0' ");
            sb.Append(" when T0.[U_ACME_BackQty1] like '%#%' then '0' ");
            sb.Append(" else T0.[U_ACME_BackQty1] end ,0) as int) <> 0)");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%/%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%+%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%.%'");
            sb.Append(" and T0.[U_ACME_BackQty1] not like '%#%'");
            sb.Append(" and T0.[U_Rquinity] not like '%/%'");
            sb.Append(" and T0.[U_Rquinity] not like '%+%'");
            sb.Append(" and T0.[U_Rquinity] not like '%.%'");
            sb.Append(" and T0.[U_Rquinity] not like '%#%'");
            sb.Append("  and T0.[U_Rquinity] <> '/'");
            sb.Append(" and T0.[U_Cusname_S] <> 'AUO_2nd RMA' and T0.[U_ACME_BackDate1] <> '1900-01-01 00:00:00.000'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
  

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA5(string Racmetodate)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼, T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱,");
            sb.Append(" T0.[U_RModel] Model, T0.[U_RVer] Ver, T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_Racmetodate] ACME通知退運日, T0.[U_RtoReceiving] ACME收貨日, ");
            sb.Append(" cast(T0.[Remarks2] as varchar(250)) 備註, T0.[U_REngineer] Engineer FROM ACMESQL02.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE T0.[U_RtoReceiving]  is null and  ");
            sb.Append("  T0.[U_Racmetodate]  < cast(@Racmetodate as datetime)-14  AND substring(T0.[U_RMA_NO],1,1) = 'A' ");
            sb.Append(" and T0.[TermDate]  is null and T0.[U_PKind] <> '1' ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼, T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱,");
            sb.Append(" T0.[U_RModel] Model, T0.[U_RVer] Ver, T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_Racmetodate] ACME通知退運日, T0.[U_RtoReceiving] ACME收貨日, ");
            sb.Append(" cast(T0.[Remarks2] as varchar(250)) 備註, T0.[U_REngineer] Engineer FROM ACMESQL05.DBO.OCTR T0  ");
            sb.Append(" INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE T0.[U_RtoReceiving]  is null and  ");
            sb.Append("  T0.[U_Racmetodate]  < cast(@Racmetodate as datetime)-14   ");
            sb.Append(" and T0.[TermDate]  is null  order by T0.U_RMA_NO");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Racmetodate", Racmetodate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA6(string BackDate)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼, T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱, ");
            sb.Append(" T0.[U_RModel] Model,T0.[U_RVer] Ver, T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日,");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.U_YETQTY Vender未還數量, T0.[U_ACME_BackDate] Vender還貨日");
            sb.Append(" ,t0.u_acme_qback Vender已還數量,T0.[U_ACME_BackDate1] ACME還貨日,T0.[U_ACME_BackQty1] ACME還貨數量,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL02.DBO.OCTR T0 ");
            sb.Append("  INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE T0.[U_ACME_BackDate1]  IS NULL AND  T0.[U_ACME_BackDate]  < cast(@BackDate as datetime)- 3 ");
            sb.Append("  AND substring(T0.[U_RMA_NO],1,1) = 'A'  and T0.[TermDate]  is null ");
            sb.Append(" and T0.[U_PKind] <> '1'");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼, T0.[U_RMA_NO] RMANo, T0.[U_Cusname_S] 客戶簡稱, ");
            sb.Append(" T0.[U_RModel] Model,T0.[U_RVer] Ver, T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日,");
            sb.Append(" ''''+T0.[U_AUO_RMA_NO] VenderRMANo,T0.U_YETQTY Vender未還數量, T0.[U_ACME_BackDate] Vender還貨日");
            sb.Append(" ,t0.u_acme_qback Vender已還數量,T0.[U_ACME_BackDate1] ACME還貨日,T0.[U_ACME_BackQty1] ACME還貨數量,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL05.DBO.OCTR T0 ");
            sb.Append("  INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID ");
            sb.Append(" WHERE T0.[U_ACME_BackDate1]  IS NULL AND  T0.[U_ACME_BackDate]  < cast(@BackDate as datetime)- 3 ");
            sb.Append("   and T0.[TermDate]  is null ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BackDate", BackDate));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA7(string RtoReceiving)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model, T0.[U_RVer] Ver");
            sb.Append(" ,T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,");
            sb.Append(" T0.[U_ACME_Out] 出貨Vender日,T0.[U_yetqty] Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL02.DBO.OCTR T0");
            sb.Append("   INNER JOIN ACMESQL02.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID");
            sb.Append("  WHERE T0.[U_ACME_Out] IS NULL AND  T0.[U_RtoReceiving] < cast(@RtoReceiving as datetime)- 7  ");
            sb.Append(" AND substring(T0.[U_RMA_NO],1,1) = 'A' and T0.[TermDate]  is null");
            sb.Append("  and T0.[U_PKind] <>  '1' and T0.[U_RtoReceiving] <> '1900.01.01'");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT distinct T1.[ContractID] 契約號碼,T0.[U_RMA_NO] RMANo,T0.[U_Cusname_S] 客戶簡稱,T0.[U_RModel] Model, T0.[U_RVer] Ver");
            sb.Append(" ,T0.[U_RGrade] Grade, T0.[U_Rquinity] Qty,T0.[U_RtoReceiving] ACME收貨日,''''+T0.[U_AUO_RMA_NO] VenderRMANo,");
            sb.Append(" T0.[U_ACME_Out] 出貨Vender日,T0.[U_yetqty] Vender未還數量,cast(T0.[Remarks2] as varchar(250)) 備註 FROM ACMESQL05.DBO.OCTR T0");
            sb.Append("   INNER JOIN ACMESQL05.DBO.CTR1 T1 ON T0.ContractID = T1.ContractID");
            sb.Append("  WHERE T0.[U_ACME_Out] IS NULL AND  T0.[U_RtoReceiving] < cast(@RtoReceiving as datetime)- 7  ");
            sb.Append("  and T0.[TermDate]  is null");
            sb.Append("   and T0.[U_RtoReceiving] <> '1900.01.01'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RtoReceiving", RtoReceiving));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetRMA9()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT t0.cardname 客戶名稱,t2.slpname 業務,T1.HOMETEL SA");
            sb.Append(" ,T5.ENGINEER 工程師 FROM OCRD T0");
            sb.Append("  LEFT JOIN OSLP T2 ON (T0.SlpCode = T2.SlpCode)");
            sb.Append("  LEFT JOIN OCRG T3 ON (T0.GROUPCODE = T3.GROUPCODE)");
            sb.Append(" LEFT JOIN (SELECT CARDCODE,MAX(OWNERCODE) OWNERCODE FROM ORDR  where OWNERCODE in (select empid from ohem where isnull(termdate,'') =  '' )");
            sb.Append("   GROUP BY CARDCODE) T4 ON (T0.CARDCODE=T4.CARDCODE)");
            sb.Append(" LEFT JOIN OHEM T1 ON (T4.OWNERCODE=T1.EMPID)");
            sb.Append(" LEFT JOIN (SELECT MAX(CONTRACTID) ID,CSTMRCODE CARDCODE,U_RENGINEER ENGINEER FROM OCTR  ");
            sb.Append(" WHERE ISNULL(U_RENGINEER,'') <> '' AND YEAR(STARTDATE) > 2014");
            sb.Append(" GROUP BY CSTMRCODE,U_RENGINEER) T5 ON (T0.CARDCODE=T5.CARDCODE)");
            sb.Append(" where SUBSTRING(T3.GROUPNAME,4,15) ='TFT' AND T0.CARDCODE NOT IN ('0003-01','0003-52')");
            if (comboBox1.Text != "")
            {
                sb.Append(" AND T5.ENGINEER=@U_RENGINEER");
            }
            sb.Append("  ORDER BY t0.cardname");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RENGINEER", comboBox1.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetRMA9COM()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT ISNULL(U_RENGINEER,'') ENGINEER FROM OCTR   WHERE  ISNULL(U_RENGINEER,'') <> '/' AND YEAR(STARTDATE) > 2014 ORDER BY ISNULL(U_RENGINEER,'') ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void RMASAPSEARCH_Load(object sender, EventArgs e)
        {
            string DATETIME = GetMenu.Day();
            textBox1.Text = DateTime.Now.ToString("yyyy").Substring(2, 2);
            textBox2.Text = DATETIME;
            textBox3.Text = DATETIME;
            textBox4.Text = DATETIME;
            textBox5.Text = DATETIME;
            textBox6.Text = DATETIME;
            textBox7.Text = DateTime.Now.ToString("yyyy").Substring(2, 2);

            System.Data.DataTable dt4 = GetRMA9COM();
            comboBox1.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + "SAP查詢管理員.xls";

            GridViewToExcelSZRMA(dataGridView8, OutPutFile);
            string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
            ExcelReport.MailTest("SAP查詢管理員", fmLogin.LoginID.ToString(), MAIL, "", "");
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        public void GridViewToExcelSZRMA(DataGridView dgv, string OutPutFile)
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

                for (int i = 0; i < dgv.Columns.Count; i++)
                {

                    wsheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

                }

                for (int i = 0; i < dgv.Rows.Count; i++)
                {

                    DataGridViewRow row = dgv.Rows[i];

                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                        DataGridViewCell cell = row.Cells[j];

                        try
                        {

                            wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                        }

                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);

                        }

                    }

                }
                string SHNAME = "";
                for (int S = 1; S <= 7; S++)
                {
                    DataGridView dgv2 = null;
                    if (S == 1)
                    {
                        dgv2 = dataGridView7;
                        SHNAME = "收到Vender還貨超過3天未還回客戶";
                    }
                    if (S == 2)
                    {
                        dgv2 = dataGridView6;
                        SHNAME = "客戶RMA#OPEN 超過2週未寄回";
                    }
                    if (S == 3)
                    {
                        dgv2 = dataGridView5;
                        SHNAME = "還貨客人數量與核發數量不符";
                    }
                    if (S == 4)
                    {
                        dgv2 = dataGridView4;
                        SHNAME = "出貨Vender超過7天未收到Receiving notice";
                    }
                    if (S == 5)
                    {
                        dgv2 = dataGridView3;
                        SHNAME = "出貨Vender超過2周未還回";
                    }
                    if (S == 6)
                    {
                        dgv2 = dataGridView2;
                        SHNAME = "RMA未結案明細";
                    }
                    if (S == 7)
                    {
                        dgv2 = dataGridView1;
                        SHNAME = "RMA管制表";
                    }
                    wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wsheet.Name = SHNAME;
                    for (int i = 0; i < dgv2.Columns.Count; i++)
                    {

                        wsheet.Cells[1, i + 1] = dgv2.Columns[i].HeaderText;

                    }

                    for (int i = 0; i < dgv2.Rows.Count; i++)
                    {

                        DataGridViewRow row = dgv2.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {

                            DataGridViewCell cell = row.Cells[j];

                            try
                            {

                                wsheet.Cells[i + 2, j + 1] = (cell.Value == null) ? "" : cell.Value.ToString();

                            }

                            catch (Exception ex)
                            {

                                MessageBox.Show(ex.Message);

                            }

                        }

                    }
                }

       


            }

            catch (Exception ex1)
            {

                MessageBox.Show(ex1.Message);

            }
            object oMissing = System.Reflection.Missing.Value;
          //  wapp.UserControl = true;
            try
            {
                wsheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message);
            }

            //增加一個 Close
            wbook.Close(oMissing, oMissing, oMissing);
            //Quit
            wapp.Quit();


            System.Runtime.InteropServices.Marshal.ReleaseComObject(wsheet);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wapp);



 

            System.GC.Collect();
            //可以將 Excel.exe 清除
            System.GC.WaitForPendingFinalizers();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView9.DataSource = GetRMA9();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView9);
        }
 
    }
}

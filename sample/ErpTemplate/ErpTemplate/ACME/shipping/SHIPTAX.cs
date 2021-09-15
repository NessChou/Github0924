using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Web.UI;
using System.IO;
using System.Diagnostics;
namespace ACME
{
    public partial class SHIPTAX : Form
    {
        System.Data.DataTable dtCost = null;
        System.Data.DataTable dtCost2 = null;
        DataRow dr = null;
        DataRow dr2 = null;
        public SHIPTAX()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetDATA();
            System.Data.DataTable dtRMA = GetDATARMA();
            dtCost = MakeTable();
            T1(dt,"SHIP");
            T1(dtRMA, "RMA");
            dataGridView1.DataSource = dtCost;

            System.Data.DataTable dt2 = GetDATA2();
            System.Data.DataTable dt2RMA = GetDATA2RMA();
            dtCost2 = MakeTable2();
            T2(dt2, "SHIP");
            T2(dt2RMA, "RMA");
            dataGridView2.DataSource = dtCost2;
            if (dtCost2.Rows.Count == 0)
            {
                MessageBox.Show("未請款沒有資料");
            }
        }
        private void T1(System.Data.DataTable dt, string DOCTYPE)
        {
         
      
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string SHIPNO = dt.Rows[i]["工單號碼"].ToString();
                string OPNO = dt.Rows[i]["採購單號"].ToString();
                string CHNO = "";
                string TRANSID = "";
                dr["工單號碼"] = SHIPNO;
                dr["採購單號"] = OPNO;
                dr["廠商編號"] = dt.Rows[i]["廠商編號"].ToString();
                dr["廠商名稱"] = dt.Rows[i]["廠商名稱"].ToString();
                dr["請款"] = dt.Rows[i]["請款"].ToString();
                System.Data.DataTable G1 = null;
                if (DOCTYPE == "SHIP")
                {
                    G1 = GetINVO(SHIPNO);
                }
                if (DOCTYPE == "RMA")
                {
                    G1 = GetINVORMA(SHIPNO);
                }
                System.Data.DataTable G2 = GetINVO2(OPNO);

                dr["INV金額"] = G1.Rows[0][0].ToString();
                if (G2.Rows.Count > 0)
                {

                    CHNO = G2.Rows[0][0].ToString();
                    TRANSID = G2.Rows[0][2].ToString();
                    dr["已付款"] = G2.Rows[0][1].ToString();
                }
                dr["AP號碼"] = CHNO;
                System.Data.DataTable G3 = GetINVO3(TRANSID);
                if (G3.Rows.Count > 0)
                {
                    dr["稅單號"] = G3.Rows[0][0].ToString();
                }
                System.Data.DataTable G4 = GetINVO4(CHNO);
                string PAYDATE = "";
                if (G4.Rows.Count > 0)
                {
                    PAYDATE = G4.Rows[0][0].ToString();
                }
                dr["安排付款日期"] = PAYDATE;
                if (!String.IsNullOrEmpty(PAYDATE))
                {
                    DateTime F1 = Convert.ToDateTime(PAYDATE);
                    DateTime F2 = DateTime.Now;
                    int result = DateTime.Compare(F2, F1);
                    if (result == -1)
                    {
                        dr["已付款"] = "未付款";
                    }
                }
                dtCost.Rows.Add(dr);
            }
        }


        private void T2(System.Data.DataTable dt2,string DOCTYPE)
        {
          
            for (int i = 0; i <= dt2.Rows.Count - 1; i++)
            {
                dr2 = dtCost2.NewRow();
                string SHIPNO = dt2.Rows[i]["SHIPPINGCODE"].ToString();
            
                System.Data.DataTable G1 = null;
                if (DOCTYPE == "SHIP")
                {
                    G1 = GetOPOR(SHIPNO);
                }
                if (DOCTYPE == "RMA")
                {
                    G1 = GetOPORRMA(SHIPNO);
                }

                if (G1.Rows.Count == 0)
                {
                    System.Data.DataTable G2 = null;
                    if (DOCTYPE == "SHIP")
                    {
                        G2 = GetINVO(SHIPNO);
                    }
                    if (DOCTYPE == "RMA")
                    {
                        G2 = GetINVORMA(SHIPNO);
                    }

                    System.Data.DataTable dtDate = GetShipping_Main(SHIPNO);


                    dr2["工單號碼"] = SHIPNO;
                    dr2["報單號碼"] = dt2.Rows[i]["報單號碼"].ToString();
                    dr2["預計抵達日期"] = dtDate.Rows[0]["arriveDay"].ToString();
                    dr2["INV金額"] = G2.Rows[0][0].ToString();
                    dtCost2.Rows.Add(dr2);
                }
            }
        }
        private void TaxCount() 
        {
            foreach (DataRow row in dtCost2.Rows) 
            {

            }
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("AP號碼", typeof(string));
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("INV金額", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("稅單號", typeof(string));
            dt.Columns.Add("請款", typeof(string));
            dt.Columns.Add("已付款", typeof(string));
            dt.Columns.Add("安排付款日期", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("報單號碼", typeof(string));
            dt.Columns.Add("預計抵達日期", typeof(string));
            dt.Columns.Add("INV金額", typeof(decimal));
            dt.Columns.Add("關稅百分比", typeof(string));
            dt.Columns.Add("關稅", typeof(decimal));
            dt.Columns.Add("推貿費", typeof(decimal));
            dt.Columns.Add("營業稅", typeof(decimal));
            dt.Columns.Add("總額", typeof(int));



            return dt;
        }
        private System.Data.DataTable GetDATA()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT  T0.U_SHIPPING_NO 工單號碼,T0.DOCENTRY 採購單號,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,CAST(T0.DOCTOTAL AS INT) 請款,T1.ARRIVEDAY FROM OPOR T0  ");
            sb.Append("               LEFT JOIN ACMESQLSP.DBO.SHIPPING_MAIN T1 ON (T0.U_SHIPPING_NO=T1.SHIPPINGCODE COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("               WHERE T0.CARDCODE IN ('U0121','U0133','U0134','U0135','U0221')  AND T1.ARRIVEDAY between @t1 and @t2");
            if (tAXCHECKCheckBox.Checked)
            {
                sb.Append("    and isnull(tAXCHECK,'')='Checked'           ");
            
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDATARMA()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                          SELECT  DISTINCT T1.U_SHIPPING_NO 工單號碼,T0.DOCENTRY 採購單號,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,CAST(T0.DOCTOTAL AS INT) 請款 FROM OPOR T0 ");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("                             LEFT JOIN ACMESQLSP.DBO.RMA_MAIN T2 ON (T1.U_SHIPPING_NO=T2.SHIPPINGCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                             WHERE T0.CARDCODE IN ('U0121','U0133','U0134','U0135','U0221')  AND T2.ArriveDay between @t1 and @t2");
            if (tAXCHECKCheckBox.Checked)
            {
                sb.Append("    and isnull(boatName,'')='Checked'           ");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDATA2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.SHIPPINGCODE,add9 報單號碼 FROM SHIPPING_MAIN T0");
            sb.Append(" WHERE   T0.ARRIVEDAY between @t1 and @t2");
            sb.Append("    and isnull(tAXCHECK,'')='Checked'            ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDATA2RMA()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.SHIPPINGCODE,add1 報單號碼 FROM RMA_MAIN T0");
            sb.Append(" WHERE   T0.ARRIVEDAY between @t1 and @t2");
            sb.Append("    and isnull(boatName,'')='Checked'           ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@t1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@t2", textBox2.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetINVO(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CAST(SUM(AMOUNT) AS INT) AMT  FROM INVOICED T0 ");
            sb.Append("  where T0.[shippingcode]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetINVORMA(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                SELECT CAST(SUM(AMOUNT) AS INT) AMT  FROM RMA_INVOICED T0  ");
            sb.Append("                where T0.[shippingcode]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public System.Data.DataTable GetShipping_Main(string shippingcode) 
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                SELECT * From shipping_main  ");
            sb.Append("                where shippingcode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " shipping_main ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" shipping_main "];
        }
        public static System.Data.DataTable GetOPOR(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  t0.docentry");
            sb.Append(" FROM  acmesql02.dbo.[OPOR]  T0 ");
            sb.Append(" INNER JOIN acmesql02.dbo.POR1 T1 ON T0.DocEntry = T1.DocEntry ");
            sb.Append(" WHERE T0.[U_Shipping_no] =@shippingcode or T1.[U_Shipping_no]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetOPORRMA(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T0.DOCENTRY FROM  acmesql02.dbo.PCH1 T0");
            sb.Append(" WHERE  T0.DOCENTRY  not in (select ISNULL(baseref,0) from  acmesql02.dbo.rPC1) ");
            sb.Append(" and T0.[U_Shipping_no]=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetINVO2(string BASEENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.DOCENTRY,CASE DOCTOTAL-PAIDTODATE WHEN 0 THEN '已付款' ELSE '未付款' END PAY,T0.TRANSID    FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T1.BASETYPE=22 AND T1.BASEENTRY=@BASEENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BASEENTRY", BASEENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetINVO3(string U_PC_BSCUS)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_PC_BSCUS FROM [@CADMEN_FMD] T0");
            sb.Append(" LEFT JOIN [@CADMEN_FMD1]  T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T0.U_BSREN=@U_PC_BSCUS");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_PC_BSCUS", U_PC_BSCUS));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetINVO4(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT Convert(varchar(10),T1.DOCDATE,111) DATE FROM VPM2  T0 ");
            sb.Append("               inner join OVPM t1 on (t0.docnum=t1.docnum)  ");
            sb.Append("               WHERE  t1.canceled <> 'Y' AND T0.DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public static System.Data.DataTable GetHAIGUAN(string year ,string month ,string date)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM WH_HAIGUAN ");
            sb.Append("WHERE HYEAR = @year and HMON = @month and (HDAY2 <= @date and HDAY3>= @date)");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@year", year));
            command.Parameters.Add(new SqlParameter("@month", month));
            command.Parameters.Add(new SqlParameter("@date", date));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " WH_HAIGUAN ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" WH_HAIGUAN "];
        }
        public static System.Data.DataTable GetHAIGUAN()
        {
            //抓最後一筆
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT top 1 * FROM WH_HAIGUAN ");
            sb.Append(" ORDER BY ID DESC ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " WH_HAIGUAN ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" WH_HAIGUAN "];
        }

        private void SHIPTAX_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
        }

        private void dataGridView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {

                string da = dataGridView2.SelectedRows[0].Cells["工單號碼"].Value.ToString();

                fmShip a = new fmShip();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private void btnTaxCount_Click(object sender, EventArgs e)
        {
            bool persent = false;
            foreach (DataGridViewRow row in dataGridView2.Rows) 
            {
                if (row.Cells["工單號碼"].Value == null) 
                {
                    continue;
                } 
                decimal InvAmount = Convert.ToDecimal(row.Cells["INV金額"].Value);
                decimal tax;//關稅趴數
                decimal tariff;//關稅金額
                decimal Promotion;//推貿費
                decimal Turnover;//營業稅
                decimal AmountAll;//總額
                try
                {
                    if (row.Cells["關稅百分比"].Value.ToString() != null && row.Cells["關稅百分比"].Value.ToString() != "")
                    {
                        tax = (Convert.ToDecimal(row.Cells["關稅百分比"].Value) / 100);

                        tariff = InvAmount * tax;
                        row.Cells["關稅"].Value = tariff;

                        Promotion = InvAmount * (1 + tax) * Convert.ToDecimal(0.0004);//推貿費固定4/10000
                        row.Cells["推貿費"].Value = Promotion;

                        Turnover = InvAmount * (1 + tax) * Convert.ToDecimal(0.05);//營業額固定0.05
                        row.Cells["營業稅"].Value = Turnover;

                        AmountAll = tariff + Promotion + Turnover;

                        string year = Convert.ToString(row.Cells["預計抵達日期"].Value).Substring(0, 4);
                        string month = Convert.ToString(row.Cells["預計抵達日期"].Value).Substring(4, 2);
                        string Day = Convert.ToString(row.Cells["預計抵達日期"].Value).Substring(6, 2);

                        System.Data.DataTable dt = GetHAIGUAN(year, month, Day);
                        decimal rate;
                        if (dt.Rows.Count > 0)
                        {
                            rate = Convert.ToDecimal(dt.Rows[0]["HSELL"]);
                        }
                        else
                        {
                            dt = GetHAIGUAN();
                            rate = Convert.ToDecimal(dt.Rows[0]["HSELL"]);
                        }
                        int result;
                        result = Convert.ToInt32(Math.Ceiling(AmountAll * rate));//無條件進位
                        row.Cells["總額"].Value = result;
                        persent = true;

                    }
                }
                catch (Exception ex) 
                {

                }
            }
            if (persent == false)
            {
                MessageBox.Show("請填入關稅百分比");
            }
        }

        private void btnEmail_Click(object sender, EventArgs e)
        {
            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\MailTemplates\\Shiptax.html";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border:3px #F5F5DC groove;'>");

            int year = 0;
            int month = 0;
            int day = 0;


            sb.AppendLine("<tr>");
            for (int j = 0; j < dataGridView2.Columns.Count; j++)
            {
                //欄位名稱
                sb.AppendLine("<td bgcolor=\"#726E6D\"><span style=\"color: white;\">" + dataGridView2.Columns[j].HeaderText + "</span></td>");
            }
            sb.AppendLine("</tr>");

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (i % 2 == 0)
                {
                    sb.AppendLine("<tr bgcolor=\"#C0C0C0\">");
                   
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        sb.AppendLine("<td>" + dataGridView2.Rows[i].Cells[j].FormattedValue + "</td>");
                        if (dataGridView2.Rows[i].Cells[j].Value == "" || dataGridView2.Rows[i].Cells[j].Value == null) continue;
                        if (j == 2) 
                        {
                            //預計抵達日期
                            year = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(0, 4));
                            month = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(4, 2));
                            day = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(6, 2));
                        }
                    }
                    sb.AppendLine("</tr>");
                    
                }
                else
                {
                    sb.AppendLine("<tr bgcolor=\"#E5E4E2\">");
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        //只到成本差異原因,所以18
                        sb.AppendLine("<td>" + dataGridView2.Rows[i].Cells[j].FormattedValue + "</td>");
                        if (dataGridView2.Rows[i].Cells[j].Value == "" || dataGridView2.Rows[i].Cells[j].Value == null) continue;
                        if (j == 2)
                        {
                            //預計抵達日期
                            year = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(0, 4));
                            month = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(4, 2));
                            day = Convert.ToInt32(Convert.ToString(dataGridView2.Rows[i].Cells[j].Value).Substring(6, 2));
                        }
                    }
                    sb.AppendLine("</tr>");
                }

            }

            DateTime date = new DateTime(year, month, day);
            date = date.AddDays(14);//要加14天
            sb.AppendLine("</table>");

            template = template.Replace("##Date##", date.Month + "月" + date.Day + "日");

            template = template.Replace("##Template##", sb.ToString());

            string SlpName = globals.UserID;
            string MailToAddress =  globals.UserID + "@acmepoint.com";
            //string MailToAddress = globals.UserID + "@acmepoint.com" + ";" + "vickyhsiao@acmepoint.com";
            string strSubject = "進口稅金預估通知 -" + date.Month + "/" + date.Day;



            string MailFromAddress = "workflow@acmepoint.com";




            MailToPD(strSubject, MailFromAddress, MailToAddress, template);


        }
        private void MailToPD(string strSubject, string MailFromAddress, string MailToAddress, string MailContent)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress(MailFromAddress, "系統發送");
            string[] MailToAdd = MailToAddress.Split(';');
            foreach (string add in MailToAdd)
            {
                message.To.Add(new MailAddress(add));
            }


            string myMailEncoding = "utf-8";
            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = MailContent;
            //格式為 Html
            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";

            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            try
            {
                client.Send(message);
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        //SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);

                    }
                    else
                    {
                        //SetMsg(String.Format("Failed to deliver message to {0}",
                        // ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                // ex.ToString()));
            }

        }

    }
}

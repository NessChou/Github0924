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
    public partial class GB_INVOCHECK : Form
    {
        string invoice = "";
        string CONN = "";
        string strCn2 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn3 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP03;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_INVOCHECK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //發票兌獎
           
            System.Data.DataTable J1 = GetTRADE2();
            string SDATE = J1.Rows[0]["SDATE"].ToString();
            string EDATE = J1.Rows[0]["EDATE"].ToString();
            System.Data.DataTable dtGetAcmeStageG = MakeTableCombine();
            System.Data.DataTable DT = DD1(SDATE, EDATE);
            DataRow dr = null;
            if (DT.Rows.Count > 0)
            {
                for (int i = 0; i <= DT.Rows.Count - 1; i++)
                {


                    string CARDCODE = DT.Rows[i]["客戶編號"].ToString().Trim();
                    System.Data.DataTable F5 = GetCUST(CARDCODE);
                    if (F5.Rows.Count > 0)
                    {
                        dr = dtGetAcmeStageG.NewRow();
                        string RESULT = "";
                        string T1 = DT.Rows[i]["發票號碼"].ToString().Trim();
                        dr["發票號碼"] = T1;
                        dr["銷貨單號"] = DT.Rows[i]["銷貨單號"].ToString().Trim();
                        dr["客戶編號"] = CARDCODE;
                        dr["客戶名稱"] = DT.Rows[i]["客戶名稱"].ToString().Trim();
                        dr["金額"] = DT.Rows[i]["金額"].ToString().Trim();
                        dr["發票日期"] = DT.Rows[i]["發票日期"].ToString().Trim();
                        dr["發票月份"] = DT.Rows[i]["發票月份"].ToString().Trim();
                        string InvNo = T1.Substring(2, 8);
                        System.Data.DataTable F1 = GetCHECK("特別獎");
                        System.Data.DataTable F2 = GetCHECK("特獎");
                        System.Data.DataTable F3 = GetCHECK("頭獎");
                        System.Data.DataTable F4 = GetCHECK("增開六獎");

                        for (int i2 = 0; i2 <= F1.Rows.Count - 1; i2++)
                        {
                            string C1 = F1.Rows[i2][0].ToString().Trim();

                            if (InvNo == C1)
                            {
                                RESULT = "特別獎";
                            }
                        }
                        for (int i2 = 0; i2 <= F2.Rows.Count - 1; i2++)
                        {
                            string C1 = F2.Rows[i2][0].ToString().Trim();

                            if (InvNo == C1)
                            {
                                RESULT = "特獎";
                            }
                        }
                        for (int i2 = 0; i2 <= F3.Rows.Count - 1; i2++)
                        {
                            string Prize_First = F3.Rows[i2][0].ToString().Trim();

                            if (InvNo == Prize_First)
                            {
                                RESULT = "頭獎";
                            }
                            else if (InvNo.Substring(1, 7) == Prize_First.Substring(1, 7))
                            {
                                RESULT = "二獎";
                            }
                            else if (InvNo.Substring(3, 5) == Prize_First.Substring(3, 5))
                            {
                                RESULT = "三獎";
                            }
                            else if (InvNo.Substring(3, 5) == Prize_First.Substring(3, 5))
                            {
                                RESULT = "四獎";
                            }
                            else if (InvNo.Substring(4, 4) == Prize_First.Substring(4, 4))
                            {
                                RESULT = "五獎";
                            }
                            else if (InvNo.Substring(5, 3) == Prize_First.Substring(5, 3))
                            {
                                RESULT = "六獎";
                            }
                        }

                        for (int i2 = 0; i2 <= F4.Rows.Count - 1; i2++)
                        {
                            string Prize_Add = F4.Rows[i2][0].ToString().Trim();

                            if (InvNo.Substring(5, 3) == Prize_Add)
                            {
                                RESULT = "六獎";

                            }
                        }

                        dr["發票兌獎"] = RESULT;
                        dtGetAcmeStageG.Rows.Add(dr);
                    }
                }

                if (checkBox1.Checked)
                {
                    dtGetAcmeStageG.DefaultView.RowFilter = " ISNULL(發票兌獎,'') <> '' ";
                }
            }
            else
            {
                dtGetAcmeStageG.DefaultView.RowFilter = " ISNULL(發票兌獎,'') = 'ss' ";
            }
            dataGridView1.DataSource = dtGetAcmeStageG;
            dataGridView2.DataSource = GetTRADE3();
        }

        public System.Data.DataTable DD1(string SDATE,string EDATE)
        {
            if (comboBox2.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox2.Text == "東門")
            {
                CONN = strCn3;
            }
            SqlConnection MyConnection = new SqlConnection(CONN);
            StringBuilder sb = new StringBuilder();
            //'94761178' 
            sb.Append(" SELECT InvoiceNO  發票號碼,SrcBillNO 銷貨單號,CustomerID 客戶編號,CompanyName 客戶名稱,CAST(AMOUNT AS INT) 金額,InvoiceDate 發票日期,ApplyMonth 發票月份   FROM comInvoice WHERE Flag =2 AND ISNULL(InvoiceNO,'') <> '' AND ApplyMonth > 201312 and IsCancel <> 1 ");
            sb.Append("   AND  ApplyMonth BETWEEN @SDATE AND @EDATE  ");
   
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SDATE", SDATE));
            command.Parameters.Add(new SqlParameter("@EDATE", EDATE));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private System.Data.DataTable GetCUST(string ID)
        {
            if (comboBox2.Text == "聿豐")
            {
                CONN = strCn2;
            }
            if (comboBox2.Text == "東門")
            {
                CONN = strCn3;
            }

            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();
            sb.Append("   select ID from comCustomer where Flag =1 AND ID=@ID AND ISNULL(taxno,'') =''  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));

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
        private void GB_INVOCHECK_Load(object sender, EventArgs e)
        {
            comboBox2.Text = "聿豐";
            UtilSimple.SetLookupBinding(comboBox1, GetTRADE(), "DataValue", "DataValue");
        }
        public System.Data.DataTable GetTRADE()
        {

            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT YMON DataValue FROM GB_INVNO ORDER BY YMON DESC";

            SqlDataAdapter da = new SqlDataAdapter(sql, MyConnection);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["oslp"];
        }
        public System.Data.DataTable GetTRADE2()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SDATE,EDATE FROM GB_INVNO WHERE YMON=@YMON ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@YMON", comboBox1.Text));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        public System.Data.DataTable GetTRADE3()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PNAME 中獎分類,PMONEY 中獎金額,PNUM 發票號碼 FROM GB_INVNO WHERE YMON=@YMON ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@YMON", comboBox1.Text));
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OWTR");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OWTR"];
        }
        private System.Data.DataTable MakeTableCombine()
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("銷貨單號", typeof(string));
            dt.Columns.Add("客戶編號", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("金額", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("發票月份", typeof(string));
            dt.Columns.Add("發票兌獎", typeof(string));
            return dt;
        }

        private System.Data.DataTable GetCHECK(string PNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   select PNUM from GB_INVNO where YMON=@YMON AND PNAME=@PNAME");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YMON", comboBox1.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@PNAME", PNAME));

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

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GetInvoiceFromWeb("http://invoice.etax.nat.gov.tw/");
        }

        private void GetInvoiceFromWeb(string addr)
        {
            try
            {
                string F1 = DateTime.Now.AddMonths(-2).ToString("yyyyMM");
                string F2 = DateTime.Now.AddMonths(-1).ToString("MM");
                string F4 = DateTime.Now.AddMonths(-1).ToString("yyyyMM");
                string F3 = F1 + "-" + F2;
                System.Data.DataTable K1 = GetOrderData22(F1, "頭獎");

                if (K1.Rows.Count == 0)
                {


                    this.Cursor = Cursors.WaitCursor;
                    System.Net.WebClient wc = new System.Net.WebClient();
                    byte[] ws;

                    ws = wc.DownloadData(addr); //下載網頁
                    invoice = Encoding.UTF8.GetString(ws, 0, ws.Length);//把網頁存到invoice內.

                    //int mon = invoice.IndexOf("月統一發票中獎號碼單", 0);
                    int isno = 0;
                    int Seek = 0;

                    //lbl_month.Text = invoice.Substring(mon - 8, 17).Trim();

                    double dubleVal = 0;
                    int Special = invoice.IndexOf("特別獎", 0);
                    for (Seek = Special; Seek < invoice.Length; Seek++)
                    {
                        if (double.TryParse(invoice.Substring(Seek, 1), System.Globalization.NumberStyles.Integer, null, out dubleVal))
                        {
                            isno++;
                            if (isno == 8)
                            {
                                break;
                            }
                        }
                        else
                        {
                            isno = 0;
                        }
                    }
                    string df = invoice.Substring(Seek - 7, 8);
                    AddINVON(F3, "特別獎", 10000000, df, F1, F4);

                    int Special2 = invoice.IndexOf("特獎", 0);
                    for (Seek = Special2; Seek < invoice.Length; Seek++)
                    {
                        if (double.TryParse(invoice.Substring(Seek, 1), System.Globalization.NumberStyles.Integer, null, out dubleVal))
                        {
                            isno++;
                            if (isno == 8)
                            {
                                break;
                            }
                        }
                        else
                        {
                            isno = 0;
                        }
                    }
                    string df2 = invoice.Substring(Seek - 7, 8);
                    AddINVON(F3, "特獎", 2000000, df2, F1, F4);

                    int BigPrize = invoice.IndexOf("頭獎", 0);
                    string[] BPno = new string[3];
                    int G = 0;
                    for (int i = 0; i < 3; i++)
                    {

                        if (G != 0)
                        {
                            BigPrize = G;
                        }

                        for (Seek = BigPrize; Seek < invoice.Length; Seek++)
                        {
                            if (double.TryParse(invoice.Substring(Seek, 1), System.Globalization.NumberStyles.Integer, null, out dubleVal))
                            {
                                isno++;
                                if (isno == 8)
                                {
                                    G = Seek;
                                    break;
                                }
                            }
                            else
                            {
                                isno = 0;
                            }
                        }
                        BPno[i] = invoice.Substring(Seek - 7, 8);
                        AddINVON(F3, "頭獎", 200000, BPno[i].ToString(), F1, F4);

                    }


                    int BigPrize2 = invoice.IndexOf("增開六獎", 0);
                    int BigPrize3 = invoice.IndexOf("同期統一發票收執聯末3位數號碼與增開六獎號碼相同者各得獎金", 0);
                    int H1 = 0;
                    for (Seek = BigPrize2; Seek < BigPrize3; Seek++)
                    {
                        if (double.TryParse(invoice.Substring(Seek, 1), System.Globalization.NumberStyles.Integer, null, out dubleVal))
                        {
                            isno++;
                            if (isno == 3)
                            {
                                H1 += 1;
                            }

                        }
                        else
                        {
                            isno = 0;
                        }

                    }
                    string[] BPno2 = new string[H1];
                    int G2 = 0;
                    for (int i = 0; i < H1; i++)
                    {
                        if (G2 != 0)
                        {
                            BigPrize2 = G2;
                        }

                        for (Seek = BigPrize2; Seek < invoice.Length; Seek++)
                        {
                            if (double.TryParse(invoice.Substring(Seek, 1), System.Globalization.NumberStyles.Integer, null, out dubleVal))
                            {
                                isno++;
                                if (isno == 3)
                                {
                                    G2 = Seek;
                                    break;
                                }
                            }
                            else
                            {
                                isno = 0;
                            }
                        }
                        BPno2[i] = invoice.Substring(Seek - 2, 3);
                        string ss = BPno2[i].ToString();
                        AddINVON(F3, "增開六獎", 200, BPno2[i].ToString(), F1, F4);

                    }

                    UtilSimple.SetLookupBinding(comboBox1, GetTRADE(), "DataValue", "DataValue");
                }
                else
                {
                    MessageBox.Show("已取過發票號碼");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "程式即將關閉", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private System.Data.DataTable GetOrderData22(string SDATE, string PNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   select PNUM from GB_INVNO where SDATE=@SDATE AND PNAME=@PNAME");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SDATE", SDATE));
            command.Parameters.Add(new SqlParameter("@PNAME", PNAME));

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
        public void AddINVON(string YMON, string PNAME, int PMONEY, string PNUM, string SDATE, string EDATE)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into GB_INVNO(YMON,PNAME,PMONEY,PNUM,SDATE,EDATE) values(@YMON,@PNAME,@PMONEY,@PNUM,@SDATE,@EDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YMON", YMON));
            command.Parameters.Add(new SqlParameter("@PNAME", PNAME));
            command.Parameters.Add(new SqlParameter("@PMONEY", PMONEY));
            command.Parameters.Add(new SqlParameter("@PNUM", PNUM));
            command.Parameters.Add(new SqlParameter("@SDATE", SDATE));
            command.Parameters.Add(new SqlParameter("@EDATE", EDATE));
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

    }
}

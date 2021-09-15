using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Management;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net.Mime;
using System.Web.UI;
namespace ACME
{
    public partial class CheckMoney : Form
    {
        private string FileName;
        System.Data.DataTable H1 = null;
        System.Data.DataTable H2 = null;
        private decimal sd;
        private decimal se;
        private decimal sc;
        private decimal sdk;
        private decimal USD;
        public CheckMoney()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataRow dr22 = null;
            DataRow dr32 = null;
            string 單號;
            string 客戶;
            string 總類;
            string 過帳日期;
            string 客戶名稱;
            decimal 台幣金額;
            decimal 已支付;
            string 美金單價;
            string 收款條件;
            string 文件類型;
            string 客戶代碼;
            string 應收總計;
            string 傳票備註;
            string 傳票號碼;
            string 採購單1;
            string 採購單2;
            string 日期;
            string INVOICENO;
            string LC;
            string 收貨採購單;
            string AP發票;
            string fh = "";
            string 到期日期;
            decimal 未支付;
            string shipping工單號碼;
            string DOCTYPE = "";
            string 日數;
            string BU = "";
            string PAYMENT;
            string UU = "";
            if (textBox1.Text == "")
            {
                fh = "1";
            }
            else
            {
                fh = "2";
            }

            string FD = util.SPACE(textBox3.Text);
            System.Data.DataTable dt = GetOrderDataAP(fh, FD);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }

            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
            DataRow dr = null;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {



                單號 = dt.Rows[i]["docentry"].ToString();
                文件類型 = dt.Rows[i]["文件類型"].ToString();
                dt1 = GetOrderDataAP1(單號, 文件類型);
                if (dt1.Rows.Count > 0)
                {
                    dr = dtCost.NewRow();
                    總類 = dt1.Rows[0]["總類"].ToString();
                    DOCTYPE = dt1.Rows[0]["DOCTYPE"].ToString();
                    過帳日期 = dt1.Rows[0]["過帳日期"].ToString();
                    台幣金額 = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]);
                    已支付 = Convert.ToDecimal(dt1.Rows[0]["已支付"]);
                    未支付 = Convert.ToDecimal(dt1.Rows[0]["未支付"]);
                    採購單1 = dt1.Rows[0]["採購單1"].ToString();
                    採購單2 = dt1.Rows[0]["採購單2"].ToString();
                    收款條件 = dt1.Rows[0]["收款條件"].ToString();
                    美金單價 = dt1.Rows[0]["美金單價"].ToString();
                    客戶代碼 = dt1.Rows[0]["客戶代碼"].ToString();
                    客戶名稱 = dt1.Rows[0]["客戶名稱"].ToString();
                    應收總計 = dt1.Rows[0]["應收總計"].ToString();
                    傳票備註 = dt1.Rows[0]["傳票備註"].ToString();
                    傳票號碼 = dt1.Rows[0]["傳票號碼"].ToString();
                    到期日期 = dt1.Rows[0]["到期日期"].ToString();
                    UU = dt1.Rows[0]["USD"].ToString();
                    shipping工單號碼 = dt1.Rows[0]["shipping工單號碼"].ToString();
                    LC = dt1.Rows[0]["LC"].ToString();
                    客戶 = dt1.Rows[0]["客戶名稱"].ToString();
                    美金單價 = dt1.Rows[0]["美金單價"].ToString();
                    日期 = dt1.Rows[0]["日期"].ToString();
                    INVOICENO = dt1.Rows[0]["INVOICENO"].ToString();
                    AP發票 = dt1.Rows[0]["AP發票"].ToString();
                    收貨採購單 = dt1.Rows[0]["收貨採購單"].ToString();
                    日數 = dt1.Rows[0]["日數"].ToString();
                    BU = dt1.Rows[0]["BU"].ToString();
                    PAYMENT = dt1.Rows[0]["PAYMENT"].ToString();
                    dr["AP發票"] = AP發票;
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();
                    if (dt1.Rows.Count > 0)
                    {
                        string DUP = "";
                        string DUP2 = "";
                        for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                        {


                            DataRow dd = dt1.Rows[j];
                            string SAPDOC = dd["採購單2"].ToString();
                            string SAPDOC2 = dd["採購單1"].ToString();
                            if (DUP != SAPDOC)
                            {
                                sb.Append(SAPDOC + "/");
                            }
                            if (DUP2 != SAPDOC2)
                            {
                                sb2.Append(SAPDOC2 + "/");
                            }
                            DUP = SAPDOC;
                            DUP2 = SAPDOC2;
                        }

                        sb.Remove(sb.Length - 1, 1);
                        if (sb2.Length > 0)
                        {
                            sb2.Remove(sb2.Length - 1, 1);
                            dr["採購單1"] = sb2.ToString();
                        }
             
                        dr["採購單2"] = sb.ToString();
 
                    }
                    if (總類 == "AR")
                    {
                        dr["收貨採購單"] = 收貨採購單;
                    }
                    else
                    {
                        dr["收貨採購單"] = 單號;
                    }

                    //if (AP發票 == "59005")
                    //{

                    //    MessageBox.Show("a");
                    //}
                    string usd = "";
                    dr["客戶代碼"] = 客戶代碼;
                    dr["客戶名稱"] = 客戶名稱;
                    dr["傳票備註"] = 傳票備註;
                    dr["傳票NO"] = 過帳日期 + "-" + 傳票號碼;
                    dr["日期"] = 日期;
                    dr["INVOICENO"] = INVOICENO;
                    dr["發票號碼"] = dt1.Rows[0]["發票號碼"].ToString();
                    dr["發票日期"] = dt1.Rows[0]["發票日期"].ToString();
                    dr["項目成本差異原因"] = dt1.Rows[0]["項目成本差異原因"].ToString();
                    dr["LC"] = LC;
                    dr["到期日期"] = 到期日期;
                    dr["未支付"] = 未支付;
                    dr["shipping工單號碼"] = shipping工單號碼;
                    int g1 = 日數.IndexOf("-");
                    if (g1 != -1)
                    {
                        g1 = 0;
                    }
                    dr["日數"] = 日數;
                    dr["BU"] = BU;
                    dr["PAYMENT"] = PAYMENT;
                    if (!String.IsNullOrEmpty(shipping工單號碼))
                    {
                        System.Data.DataTable SHIP = GetSHIP(shipping工單號碼);
                        if (SHIP.Rows.Count > 0)
                        {
                            dr["收貨地"] = SHIP.Rows[0]["收貨地"].ToString();
                            dr["目的地"] = SHIP.Rows[0]["目的地"].ToString();
                        }
                    }
                    H1 = GetCMONEY(收貨採購單);
                    H2 = GetCMONEY2(單號);
                    if (總類 == "AR")
                    {
                        sc = Convert.ToDecimal(台幣金額);
                    }
                    else
                    {
                        sc = (Convert.ToDecimal(台幣金額)) * -1;
                    }
                    sd = 0;
                    int sh = 0;
                    decimal sv = 0;
                    for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                    {
                        DataRow dd = dt1.Rows[j];
                        string sa = dd["數量"].ToString();
                        string sds = dd["美金單價"].ToString();
                        string sae = dd["稅率"].ToString();
                        decimal n;
                        if ((!String.IsNullOrEmpty(dd["數量"].ToString())) && (!String.IsNullOrEmpty(dd["美金單價"].ToString())) && (!String.IsNullOrEmpty(dd["稅率"].ToString())))
                        {

                            if (decimal.TryParse(sds, out n))
                            {
                                sd += Convert.ToDecimal(dd["數量"]) * Convert.ToDecimal(dd["美金單價"]) * Convert.ToDecimal(dd["稅率"]);
                                sh += Convert.ToInt32(dd["數量"]);
                                sv = Convert.ToDecimal(dd["美金單價"]);

                                Double x = Convert.ToDouble(sd);
                                sd = Convert.ToDecimal(C1Round(x, 2).ToString());
                                if (sd != 0)
                                {
                                    se = Convert.ToDecimal(dt1.Rows[0]["台幣金額"]) / sd;
                                }

                                if (總類 == "AR")
                                {
                                    usd = sd.ToString();
                                }
                                else
                                {
                                    usd = (sd * -1).ToString();
                                }
                                if (sh > 1000) 
                                {

                                }
                                dr["總數量"] = Convert.ToDecimal(sh).ToString("#,##0.00"); ;
                                dr["美金單價2"] = sv.ToString("#,##0");
                                dr["匯率"] = se.ToString("#,##0.000");

                                if (H1.Rows.Count > 0)
                                {
                                    decimal RATE = Convert.ToDecimal(H1.Rows[0]["RATE"]);
                                    dr["匯率"] = RATE.ToString("#,##0.000");
                                }
                            }
                            else
                            {
                                sd = 0;
                                usd = "0";
                                dr["匯率"] = 0;
                            }

                        }
                        else
                        {
                            sd = 0;
                            usd = "0";
                            dr["匯率"] = 0;
                        }


                        dr["應收帳款2"] = sc.ToString("#,##0");
                        if (收貨採購單 == "29255")
                        {
                            usd = "1521.45";
                        }
                        if (收貨採購單 == "29256")
                        {
                            usd = "3,260.25";
                        }
                        if (收貨採購單 == "29262")
                        {
                            usd = "448.88";
                        }

                        if (!String.IsNullOrEmpty(UU))
                        {
                            usd = UU;
                        }
                        if (decimal.TryParse(usd, out n))
                        {
                            dr["美金應收帳款2"] = Convert.ToDecimal(usd).ToString("#,##0.00");
                        }



                        if (H1.Rows.Count > 0)
                        {
                            decimal USD2 = Convert.ToDecimal(H1.Rows[0]["USD"]);
                            dr["美金應收帳款2"] = Convert.ToDecimal(USD2).ToString("#,##0.00");
                        }

                        if (H2.Rows.Count > 0)
                        {
                            decimal USD2 = Convert.ToDecimal(H2.Rows[0]["AMOUNT"]);
                            dr["美金應收帳款2"] = Convert.ToDecimal(USD2).ToString("#,##0.00");
                        }
                        sdk = 0;
                        USD = 0;
                        dt2 = GetOrderDataAP2(單號, 文件類型);
                        for (int k = 0; k <= dt2.Rows.Count - 1; k++)
                        {
                            DataRow ddk = dt2.Rows[k];
                            if ((!String.IsNullOrEmpty(ddk["金額"].ToString())))
                            {
                                sdk += Convert.ToDecimal(ddk["金額"]);

                                if (總類 == "AR")
                                {
                                    USD = sdk;
                                }
                                else
                                {
                                    USD = (sdk * -1);
                                }

                            }
                            else
                            {
                                sdk = 0;
                                USD = 0;

                            }

                            decimal da = Convert.ToDecimal(usd) - USD;
                            dr["美金應收帳款2"] = da.ToString();

                        }


                        string F1 = dr["匯率"].ToString();

                        if (F1 == "1.000")
                        {
                            dr["美金應收帳款2"] = "0";
                        }
                    }
                    if (總類 == "AR貸項" && DOCTYPE == "S")
                    {
                        dr["美金單價2"] = "0";
                        dr["總數量"] = "0";
                    }

                    dtCost.Rows.Add(dr);


                }

                if (dtCost.Rows.Count > 0)
                {
                    //     dtCost.DefaultView.Sort = "BU,客戶代碼";
                    dataGridView1.DataSource = dtCost;
                    string g = dtCost.Compute("Sum(應收帳款2)", null).ToString();
                    string gk = dtCost.Compute("Sum(美金應收帳款2)", null).ToString();
                    decimal n;
                    if (decimal.TryParse(g, out n) && decimal.TryParse(gk, out n))
                    {
                        decimal sh = Convert.ToDecimal(g);
                        decimal shk = Convert.ToDecimal(gk);
                        label3.Text = "台幣合計:" + sh.ToString("#,##0");
                        label6.Text = "美金合計:" + shk.ToString("#,##0.0000");
                    }
                }
            }
            System.Data.DataTable dtCost2 = MakeTableCombine2();

            string 廠商1;
            string 廠商名稱;
            for (int l = 0; l <= dtCost.Rows.Count - 1; l++)
            {

                DataRow drFind;

                DataRow dz = dtCost.Rows[l];
                廠商1 = dz["客戶代碼"].ToString();

                廠商名稱 = dz["客戶名稱"].ToString();


                drFind = dtCost2.Rows.Find(廠商1);

                if (drFind == null)
                {
                    dr22 = dtCost2.NewRow();
                    string das = dtCost.Compute("Sum(應收帳款2)", "客戶代碼='" + 廠商1 + "'").ToString();



                    dr22["客戶代碼"] = 廠商1;
                    dr22["客戶名稱"] = 廠商名稱;
                    dr22["應收帳款"] = Convert.ToDecimal(das);


                    dtCost2.Rows.Add(dr22);
                }
                dtCost2.DefaultView.Sort = "客戶代碼";
                dataGridView2.DataSource = dtCost2;
            }



             System.Data.DataTable dtCost3 = MakeTableCombine3();

            string DBU = "";

            for (int l = 0; l <= dtCost.Rows.Count - 1; l++)
            {
                DataRow dz = dtCost.Rows[l];
                string BU3 = dz["BU"].ToString();
                DataRow drFind;
                if (l == 0)
                {
                    dr32 = dtCost3.NewRow();
                    string Tdas = dtCost.Compute("Sum(美金應收帳款2)", null).ToString();
                    string Tdas2 = dtCost.Compute("Sum(應收帳款2)", null).ToString();
                    string TdQ = dtCost.Compute("Sum(總數量)", null).ToString();
                    decimal TTdas = Convert.ToDecimal(Tdas);
                    decimal TTdas2 = Convert.ToDecimal(Tdas2);
                    decimal TTQ = Convert.ToDecimal(TdQ);
                    dr32["BU"] = "總計";
                    dr32["美金應收帳款2"] = TTdas.ToString("#,##0.00");
                    dr32["應收帳款2"] = TTdas2.ToString("#,##0");
                    dr32["總數量"] = TTQ.ToString("#,#");
                    dr32["INVOICENO"] = (TTdas2 / TTdas).ToString("0.0000");

                    dtCost3.Rows.Add(dr32);
                }

                drFind = dtCost2.Rows.Find(BU3);

                //  dt1.Rows[0]["客戶代碼"].ToString();
                dr32 = dtCost3.NewRow();
                dr32["客戶代碼"] = dz["客戶代碼"].ToString();
                dr32["傳票NO"] = dz["傳票NO"].ToString();
                dr32["到期日期"] = dz["到期日期"].ToString();
                dr32["AP發票"] = dz["AP發票"].ToString();
                dr32["採購單1"] = dz["採購單1"].ToString();
                dr32["採購單2"] = dz["採購單2"].ToString();
                dr32["收貨採購單"] = dz["收貨採購單"].ToString();
                string rr = dz["總數量"].ToString();
                if (String.IsNullOrEmpty(rr))
                {
                    rr = "0";
                }
                dr32["總數量"] = Convert.ToDecimal(rr);
                dr32["美金單價2"] = dz["美金單價2"].ToString();
                dr32["傳票備註"] = dz["傳票備註"].ToString();
                string str = dz["美金應收帳款2"].ToString();
                dr32["美金應收帳款2"] = Convert.ToDecimal(dz["美金應收帳款2"]).ToString("#,##0.00");

                dr32["應收帳款2"] = Convert.ToDecimal(dz["應收帳款2"]).ToString("#,##0");

                dr32["INVOICENO"] = dz["INVOICENO"].ToString();
                dr32["日期"] = dz["日期"].ToString();

                dr32["發票號碼"] = dz["發票號碼"].ToString();
                dr32["發票日期"] = dz["發票日期"].ToString();

                dr32["匯率"] = dz["匯率"].ToString();

                dr32["項目成本差異原因"] = dz["項目成本差異原因"].ToString();

                //dr32["LC"] = dz["LC"].ToString();
                dr32["shipping工單號碼"] = dz["shipping工單號碼"].ToString();

                dr32["收貨地"] = dz["收貨地"].ToString();

                dr32["目的地"] = dz["目的地"].ToString();


                dr32["日數"] = dz["日數"].ToString();

                dr32["BU"] = dz["BU"].ToString();
                dr32["PAYMENT"] = dz["PAYMENT"].ToString();
                dtCost3.Rows.Add(dr32);
                if (dtCost.Rows.Count - 1 != l)
                {
                    DBU = dtCost.Rows[l + 1]["BU"].ToString();
                }
                if (dtCost.Rows.Count - 1 == l)
                {
                    if (drFind == null)
                    {
                        dr32 = dtCost3.NewRow();
                        string das = dtCost.Compute("Sum(美金應收帳款2)", "BU='" + BU3 + "'").ToString();
                        string das2 = dtCost.Compute("Sum(應收帳款2)", "BU='" + BU3 + "'").ToString();
                        decimal Hdas = Convert.ToDecimal(das);
                        decimal Hdas2 = Convert.ToDecimal(das2);
                        dr32["BU"] = BU3 + "合計";
                        dr32["美金應收帳款2"] = Hdas.ToString("#,##0.00");
                        dr32["應收帳款2"] = Hdas2.ToString("#,##0");
                        dr32["INVOICENO"] = (Hdas2 / Hdas).ToString("0.0000");
                        dtCost3.Rows.Add(dr32);


                        dr32 = dtCost3.NewRow();
                        string Tdas = dtCost.Compute("Sum(美金應收帳款2)", null).ToString();
                        string Tdas2 = dtCost.Compute("Sum(應收帳款2)", null).ToString();
                        decimal TTdas = Convert.ToDecimal(Tdas);
                        decimal TTdas2 = Convert.ToDecimal(Tdas2);
                        dr32["BU"] = "總計";
                        dr32["美金應收帳款2"] = TTdas.ToString("#,##0.00"); 
                        dr32["應收帳款2"] = TTdas2.ToString("#,##0"); 
                        dr32["INVOICENO"] = (TTdas2 / TTdas).ToString("0.0000");

                        dtCost3.Rows.Add(dr32);
                    }
                }

                if (BU3 != DBU)
                {
                    if (drFind == null)
                    {
                        dr32 = dtCost3.NewRow();
                        string das = dtCost.Compute("Sum(美金應收帳款2)", "BU='" + BU3 + "'").ToString();
                        string das2 = dtCost.Compute("Sum(應收帳款2)", "BU='" + BU3 + "'").ToString();
                        string das3 = dtCost.Compute("Sum(總數量)", "BU='" + BU3 + "'").ToString();
                        decimal Hdas = Convert.ToDecimal(das);
                        decimal Hdas2 = Convert.ToDecimal(das2);
                        decimal Hdas3 = Convert.ToDecimal(das3);

                        dr32["總數量"] =  Hdas3;
                        dr32["BU"] = BU3 + "合計";
                        dr32["美金應收帳款2"] = Hdas.ToString("#,##0.00"); 
                        dr32["應收帳款2"] = Hdas2.ToString("#,##0"); 
                        if (Hdas != 0)
                        {
                            dr32["INVOICENO"] = (Hdas2 / Hdas).ToString("0.0000");
                        }
                        dtCost3.Rows.Add(dr32);

                    }
                }



                dataGridView3.DataSource = dtCost3;
            }
        }
        private string AddComma(string str) 
        {
            string Num ="";
            return Num;

        }
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("應收帳款", typeof(Decimal));


            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["客戶代碼"];
            dt.PrimaryKey = colPk;

            return dt;
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("客戶名稱", typeof(string));
            dt.Columns.Add("傳票NO", typeof(string));
            dt.Columns.Add("AP發票", typeof(string));
            dt.Columns.Add("採購單1", typeof(string));
            dt.Columns.Add("採購單2", typeof(string));
            dt.Columns.Add("收貨採購單", typeof(string));

            dt.Columns.Add("總數量", typeof(Decimal));
            dt.Columns.Add("美金單價2", typeof(string));

            dt.Columns.Add("傳票備註", typeof(string));
            dt.Columns.Add("美金應收帳款2", typeof(Decimal));
            dt.Columns.Add("應收帳款2", typeof(Decimal));
            dt.Columns.Add("未支付", typeof(Decimal));

            //dt.Columns.Add("美金應收帳款2", typeof(string));
           //dt.Columns.Add("應收帳款2", typeof(string));
            //dt.Columns.Add("未支付", typeof(string));

            dt.Columns.Add("INVOICENO", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("項目成本差異原因", typeof(string));
            dt.Columns.Add("LC", typeof(string));
            dt.Columns.Add("到期日期", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("日數", typeof(string));
            dt.Columns.Add("收貨地", typeof(string));
            dt.Columns.Add("目的地", typeof(string));
            dt.Columns.Add("shipping工單號碼", typeof(string));
            dt.Columns.Add("PAYMENT", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableCombine3()
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("客戶代碼", typeof(string));
            dt.Columns.Add("傳票NO", typeof(string));
            dt.Columns.Add("到期日期", typeof(string));
            dt.Columns.Add("AP發票", typeof(string));
            dt.Columns.Add("採購單1", typeof(string));
            dt.Columns.Add("採購單2", typeof(string));
            dt.Columns.Add("收貨採購單", typeof(string));

            dt.Columns.Add("總數量", typeof(Decimal));
            //dt.Columns.Add("總數量", typeof(string));

            dt.Columns.Add("美金單價2", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("傳票備註", typeof(string));
            
            dt.Columns.Add("美金應收帳款2", typeof(Decimal));
            //dt.Columns.Add("美金應收帳款2", typeof(string));
            dt.Columns.Add("應收帳款2", typeof(Decimal));
            //dt.Columns.Add("應收帳款2", typeof(string));

            dt.Columns.Add("INVOICENO", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("發票號碼", typeof(string));
            dt.Columns.Add("發票日期", typeof(string));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("項目成本差異原因", typeof(string));
           // dt.Columns.Add("LC", typeof(string));
            dt.Columns.Add("收貨地", typeof(string));
            dt.Columns.Add("目的地", typeof(string));
            dt.Columns.Add("shipping工單號碼", typeof(string));
            dt.Columns.Add("日數", typeof(string));
            dt.Columns.Add("PAYMENT", typeof(string));
            return dt;
        }
        private System.Data.DataTable GetOrderDataAP(string jh, string FDF)
        {
            string FD = "";
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("             select distinct 'AR' 總類,t0.docentry,t0.objtype 文件類型,T0.CARDCODE");
            sb.Append("             from OPCH t0");
            sb.Append("             left join ORPC t1 on(cast(t0.docentry as varchar)=cast(t1.u_acme_arap as varchar)  and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2   )");
            sb.Append("             LEFT JOIN ocrd T9 ON (T0.cardcode=T9.cardcode)    ");
            sb.Append("             left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("             [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='18' )");
            sb.Append("             left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("             [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'   and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='18' )");
            sb.Append("             LEFT JOIN (SELECT DISTINCT T1.BASETYPE,T1.BASEENTRY,T0.DOCTOTAL FROM ORPC T0");
            sb.Append("             LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("             WHERE BASETYPE=18 and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 ) T6 ON (T0.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              where  Convert(varchar(8),t0.docdate,112)  between '20071231' and  @DocDate2 ");
            if (!checkBox1.Checked)
            {
                sb.Append(" and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) <>  0 ");
                sb.Append("  and ((isnull(t0.doctotal,0)-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0)) - isnull(t1.doctotal,0)) <> 0  and t0.transid not in (select je from ACMESQLSP.DBO.account_je) ");

            }//  AUO+ADP全部
            if (jh == "1")
            {
                if (comboBox1.SelectedValue.ToString() == "NAUO")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='S'  AND SUBSTRING(T0.CARDCODE,1,5) NOT IN ('S0001','S0623')   ");

                }
                else if (comboBox1.SelectedValue.ToString() == "UCARD")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='U'     ");

                }
                else if (comboBox1.SelectedValue.ToString() == "ADP+AUO全部")
                {
                    sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");

                }
                else
                {
                    sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
                }
            }
            else
            {
                sb.Append(" and T0.CARDCODE in (" + textBox1.Text.ToString() + ")  ");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0) <>  0 ");
                sb.Append("  and ((isnull(t0.doctotal,0)-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.DOCTOTAL,0)) - isnull(t1.doctotal,0)) <> 0   ");
            }
            if (FDF != "")
            {
                sb.Append(" and T0.U_ACME_PayGUI  like '%" + FDF + "%'  ");
            }
            sb.Append(" union all");
            sb.Append("               select distinct 'AR貸項' 總類,t0.docentry,t0.objtype 文件類型,T0.CARDCODE from ORPC t0");
            sb.Append("               left join OPCH t1 on(cast(t1.docentry as varchar)=cast(t0.u_acme_arap as varchar) and  Convert(varchar(8),t1.docdate,112)  between '20071231' and @DocDate2  ) ");
            sb.Append("               LEFT JOIN ocrd  T9 ON (T0.cardcode=T9.cardcode) ");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("               [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    GROUP BY T1.[DocEntry],invtype )  t4 on (t0.docentry=t4.docentry and t4.invtype='19'  )");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("               [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 GROUP BY T1.[DocEntry],invtype  ) t5 on (t0.docentry=t5.docentry and t5.invtype='19'  )");
            sb.Append("               left join (SELECT SUBSTRING(COMMENTS,4,10) DocEntry,(T0.[TRSFRSUM]) SumApplied FROM  [dbo].[OVPM] T0  ");
            sb.Append("               WHERE    T0.[Canceled] = 'N'   AND SUBSTRING(COMMENTS,1,3) = 'ARD'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    ) t6 on (t0.docentry=t6.docentry)");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[ORCT] T0  INNER  JOIN ");
            sb.Append("               [dbo].[RCT2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'  and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2    GROUP BY T1.[DocEntry],invtype )  t7 on (cast(t0.u_acme_arap as varchar)=cast(t7.docentry as varchar)  and t7.invtype='18'  )");
            sb.Append("               left join (SELECT T1.[DocEntry] DocEntry,invtype,SUM(T1.[SumApplied]) SumApplied FROM  [dbo].[OVPM] T0  INNER  JOIN ");
            sb.Append("               [dbo].[VPM2] T1  ON  T1.[DocNum] = T0.DocEntry   WHERE    T0.[Canceled] = 'N'    and  Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2  GROUP BY T1.[DocEntry],invtype  ) t8 on (cast(t0.u_acme_arap as varchar)=cast(t8.docentry as varchar) and t8.invtype='18'  )");
            sb.Append("   where   Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 and t0.doctotal-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0) <>  0    and ((isnull(t0.doctotal,0)-isnull(t4.sumapplied,0)-isnull(t5.sumapplied,0)-isnull(t6.sumapplied,0)) - (isnull(t1.doctotal,0)-isnull(t7.sumapplied,0)-isnull(t8.sumapplied,0))) <> 0  AND t0.transid not in (select je from ACMESQLSP.DBO.account_je) ");
            if (!checkBox1.Checked)
            {
                sb.Append("  AND (CAST(T0.DOCTOTAL-ISNULL(t4.sumapplied,0)-isnull(t5.sumapplied,0) AS INT)) <> 0  ");
                sb.Append("  AND T0.DOCENTRY NOT IN (select DISTINCT T0.DOCENTRY from RPC1 T0");
                sb.Append("   LEFT JOIN PCH1 T1 ON (T0.BASEENTRY=T1.DOCENTRY AND T0.LINENUM=T1.BASELINE ) WHERE T0.BASETYPE=18 AND Convert(varchar(8),t0.docdate,112)  between '20071231' and @DocDate2 AND T0.DOCENTRY NOT IN (select DOC from ACMESQLSP.DBO.account_je2)) ");
            }
              if (jh == "1")
            {

                if (comboBox1.SelectedValue.ToString() == "NAUO")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='S'  AND SUBSTRING(T0.CARDCODE,1,5) <> 'S0001'   ");

                }
                else if (comboBox1.SelectedValue.ToString() == "UCARD")
                {
                    sb.Append(" and SUBSTRING(T0.CARDCODE,1,1)='U'     ");

                }
                else if (comboBox1.SelectedValue.ToString() == "ADP+AUO全部")
                {
                    sb.Append("  AND SUBSTRING(T0.CARDCODE,1,5)  IN ('S0001','S0623')  ");

                }
                else
                {
                    sb.Append(" and T0.CARDCODE like '%" + comboBox1.SelectedValue.ToString() + "%'  ");
                }
            }
            else
            {
                sb.Append(" and T0.CARDCODE in (" + textBox1.Text.ToString() + ")  ");
            }

              if (FDF != "")
              {
                  sb.Append(" and T0.U_ACME_PayGUI  like '%" + FDF + "%'  ");
              }
              sb.Append(" ORDER BY   T0.CARDCODE,t0.docentry ");
              //T0.CARDCODE
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@DocDate4", FD));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "oinv");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderDataAP1(string aa, string bb)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' 總類,T0.DOCTYPE,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,  ");
            sb.Append(" T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額   ");
            sb.Append(" ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,  ");
            sb.Append(" T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR) 美金單價,T0.JRNLMEMO 應收總計,T0.JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單2,T8.BaseEntry 採購單1,T0.U_ACME_INV INVOICENO,");
            sb.Append(" T0.U_ACME_LC LC,T7.docentry 收貨採購單,t0.docentry AP發票, ");
            sb.Append(" Convert(varchar(10),T0.U_PC_BSDAT,111) 發票日期,T0.U_PC_BSINV 發票號碼,T0.U_ACME_COSTMARK 項目成本差異原因, ");
            sb.Append(" CASE T0.U_ACME_SHIPMENT WHEN '1' THEN Convert(varchar(10),T0.DOCDUEDATE,111) ELSE '' END 到期日期,T0.U_SHIPPING_NO shipping工單號碼 ");
            sb.Append(" ,DATEDIFF(DAY,Convert(varchar(10),T0.U_PC_BSDAT,111),CASE T0.U_ACME_SHIPMENT WHEN '1' THEN Convert(varchar(10),T0.DOCDUEDATE,111) ELSE '' END ) 日數 ");
            sb.Append(" ,SUBSTRING(T0.CARDCODE,CHARINDEX('-', T0.CARDCODE)+1,10) BU,T0.U_ACME_PayGUI PAYMENT,T0.U_acme_pi USD FROM OPCH T0    ");
            sb.Append(" LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)  ");
            sb.Append(" LEFT JOIN OPDN T9 ON (T7.DOCENTRY=T9.DOCENTRY)");
            sb.Append(" LEFT JOIN POR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)  ");
            sb.Append(" where t1.basetype='20' and  T0.CARDCODE LIKE '%S0001%'   ");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");

            sb.Append("              union all");
            sb.Append(" SELECT 'AR' 總類,T0.DOCTYPE,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼,  ");
            sb.Append(" T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額   ");
            sb.Append(" ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量,  ");
            sb.Append(" T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR) 美金單價,T0.JRNLMEMO 應收總計,T0.JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單2,T8.BaseEntry 採購單1,T0.U_ACME_INV INVOICENO");
            sb.Append(" ,T0.U_ACME_LC LC,T7.docentry 收貨採購單,t0.docentry AP發票,");
            sb.Append(" Convert(varchar(10),T0.U_PC_BSDAT,111)  發票日期,T0.U_PC_BSINV 發票號碼,T0.U_ACME_COSTMARK 項目成本差異原因, Convert(varchar(10),T0.DOCDUEDATE,111)  到期日期,T0.U_SHIPPING_NO shipping工單號碼 ");
            sb.Append(" ,DATEDIFF(DAY,Convert(varchar(10),T0.U_PC_BSDAT,111),Convert(varchar(10),T0.DOCDUEDATE,111)  ) 日數 ");
            sb.Append("  ,SUBSTRING(T0.CARDCODE,CHARINDEX('-', T0.CARDCODE)+1,10) BU,T0.U_ACME_PayGUI  PAYMENT,T0.U_acme_pi USD  FROM OPCH T0    ");
            sb.Append(" LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry    ");
            sb.Append(" LEFT JOIN PDN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append(" LEFT JOIN OPDN T9 ON (T7.DOCENTRY=T9.DOCENTRY)");
            sb.Append(" LEFT JOIN POR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)  ");
            sb.Append(" where t1.basetype='20' and  T0.CARDCODE NOT LIKE '%S0001%'   ");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
      
            sb.Append("              union all");
            sb.Append(" SELECT 'AR' 總類,T0.DOCTYPE,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼, ");
            sb.Append(" T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額  ");
            sb.Append(" ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量, ");
            sb.Append(" T0.COMMENTS 備註,t0.u_acme_pay 收款條件,CAST(T8.PRICE AS VARCHAR)  美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,T8.docentry 採購單2,T8.BaseEntry 採購單1,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111) 發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_SHIPPING_NO shipping工單號碼");
            sb.Append(" ,DATEDIFF(DAY,Convert(varchar(10),T0.U_PC_BSDAT,111), Convert(varchar(10),T0.DOCDUEDATE,111) ) 日數");
            sb.Append("  ,SUBSTRING(T0.CARDCODE,CHARINDEX('-', T0.CARDCODE)+1,10) BU,T0.U_ACME_PayGUI PAYMENT,T0.U_acme_pi USD FROM OPCH T0    ");
            sb.Append(" LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" LEFT JOIN POR1 T8 ON (T8.docentry=T1.baseentry AND T8.linenum=T1.baseline) ");
            sb.Append(" where t1.basetype='22' ");
            sb.Append("              and t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append(" SELECT 'AR' 總類,T0.DOCTYPE,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼, ");
            sb.Append(" T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額  ");
            sb.Append(" ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT) 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量, ");
            sb.Append(" T0.COMMENTS 備註,t0.u_acme_pay 收款條件,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期,'' 採購單2,'' 採購單1,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_PC_BSDAT,111)  發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_SHIPPING_NO shipping工單號碼");
            sb.Append(" ,DATEDIFF(DAY,Convert(varchar(10),T0.U_PC_BSDAT,111), Convert(varchar(10),T0.DOCDUEDATE,111) ) 日數");
            sb.Append("  ,SUBSTRING(T0.CARDCODE,CHARINDEX('-', T0.CARDCODE)+1,10) BU,T0.U_ACME_PayGUI PAYMENT,T0.U_acme_pi USD  FROM OPCH T0    ");
            sb.Append(" LEFT JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append(" where  t1.basetype =-1 ");
            sb.Append("  and  t0.docentry=@docentry and t0.objtype=@bb ");
            sb.Append("                            union all");
            sb.Append(" SELECT 'AR貸項' 總類,T0.DOCTYPE,CAST(T0.DOCENTRY AS VARCHAR) 單號,1+t1.vatprcnt/100 稅率,cast(t1.price as int) 台幣單價,substring(Convert(varchar(10),t0.docdate,111),6,6) 過帳日期,T0.TRANSID 傳票號碼, ");
            sb.Append(" T0.[Cardcode] 客戶代碼,T0.[CardName] 客戶名稱,T1.ITEMCODE 產品編號,Substring (T1.[ItemCode],2,8) 品名,CAST(T0.doctotal AS INT) 台幣金額  ");
            sb.Append(" ,cast(T0.PAIDtodate as int) 已支付,CAST(T0.DOCTOTAL-T0.PAIDtodate AS INT)*-1 未支付,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS DECIMAL(14,4)) end 數量, ");
            sb.Append(" T0.COMMENTS 備註,t0.u_acme_pay 收款條件,T1.u_acme_inv 美金單價,T0.JRNLMEMO 應收總計,JRNLMEMO 傳票備註,Convert(varchar(10),T0.u_acme_invoice,111) 日期 ,'' 採購單2,'' 採購單1");
            sb.Append(" ,T0.U_ACME_INV INVOICENO,U_ACME_LC LC,'' 收貨採購單,t0.docentry AP發票,Convert(varchar(10),T0.U_RP_BSDAT,111) 發票日期,U_PC_BSINV 發票號碼,U_ACME_COSTMARK 項目成本差異原因,Convert(varchar(10),T0.DOCDUEDATE,111) 到期日期,T0.U_SHIPPING_NO shipping工單號碼 ");
            sb.Append(" ,DATEDIFF(DAY,Convert(varchar(10),T0.U_RP_BSDAT,111), Convert(varchar(10),T0.DOCDUEDATE,111) ) 日數");
            sb.Append("  ,SUBSTRING(T0.CARDCODE,CHARINDEX('-', T0.CARDCODE)+1,10) BU,T0.U_ACME_PayGUI  PAYMENT,T0.U_acme_pi USD  FROM ORPC T0    ");
            sb.Append(" LEFT JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry   ");
            sb.Append("                         where  t0.docentry=@docentry and t0.objtype=@bb");
       
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
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
        private System.Data.DataTable GetOrderDataAP2(string cc, string dd)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT isnull(T0.U_USD,0) 金額,T0.DOCENTRY 來源,T0.DOCNUM 單號 FROM VPM2  T0");
            sb.Append(" inner join OVPM t1 on (t0.docnum=t1.docnum) ");
            sb.Append(" WHERE  t1.canceled <> 'Y' AND T0.DOCENTRY=@cc and t0.invtype=@dd ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@dd", dd));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT receivePlace 收貨地,goalPlace 目的地 FROM SHIPPING_MAIN  WHERE SHIPPINGCODE=@SHIPPINGCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetCMONEY(string MEMO)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PARAM_NO USD,PARAM_DESC RATE FROM PARAMS WHERE PARAM_KIND='CMONEY' AND MEMO=@MEMO");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetCMONEY2(string DOCENTRY)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT AMOUNT FROM Account_CHECKMONEY WHERE DOCENTRY=@DOCENTRY ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

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
            else if (tabControl1.SelectedIndex == 2)
            {
                ExcelReport.GridViewToExcelSHARONS(dataGridView3);
            }
        }

        private void CheckPaid_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.MoneyBU("checkmoney"), "DataTEXT", "DataValue");
            textBox2.Text = GetMenu.Day();
            label6.Text = "";
            label3.Text = "";

            for (int i = 0; i < clbMailAddress.Items.Count; i++)
            {
                clbMailAddress.SetItemChecked(i, true);
            }


        }
        public double C1Round(double value, int digit)
        {
            double vt = Math.Pow(10, digit);
            double vx = value * vt;

            vx += 0.5;
            return (Math.Floor(vx) / vt);
        }
        private void button3_Click(object sender, EventArgs e)
        {

            string cs;
            APS3 frm1 = new APS3();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                cs = frm1.q;
                if (!String.IsNullOrEmpty(cs))
                {

                    textBox1.Text = cs;


                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CheckMoney2 frm1 = new CheckMoney2();

            frm1.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
  
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


                WriteExcelAP(FileName);

                MessageBox.Show("匯入成功");

            }
        }
        public void UPPCARD(string PAYMENT, string DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OPCH SET U_ACME_PayGUI=@PAYMENT WHERE DOCENTRY=@DOCENTRY UPDATE ORPC SET U_ACME_PayGUI=@PAYMENT WHERE DOCENTRY=@DOCENTRY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PAYMENT", PAYMENT));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


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

        private void WriteExcelAP(string ExcelFile)
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



            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string DOCENTRY = "";
                string PAYMENT = "";
                string APP = "";
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    DOCENTRY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    PAYMENT = range.Text.ToString().Trim().TrimEnd();
                    PAYMENT = util.SPACE(PAYMENT);
                    if (!String.IsNullOrEmpty(DOCENTRY))
                    {
                        int n;
                        if (int.TryParse(DOCENTRY, out n))
                        {
                            UPPCARD(PAYMENT, DOCENTRY);
                        }
                    }
                    // AddAP(DOCENTRY);
                }




            }
            finally
            {


                try
                {
                    // excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        public static void Print(string filePath)
        {
            // Status = PrintJobStatus.Printing;
            // Message = string.Empty;
            try
            {
                //     logger.Debug($"Printing... {filePath}");
                ProcessStartInfo info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = filePath;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;

                Process p = new Process();
                p.StartInfo = info;
                p.Start();

                p.WaitForInputIdle();
                //以下邏輯克服無法得知Acrobat Reader或Foxit Reader是否列印完成的問題
                //最多等待180秒（假設所有檔案可在3分鐘內印完）
                var timeOut = DateTime.Now.AddSeconds(180);
                bool printing = false; //是否開始列印
                bool done = false; //是否列印完成
                //取純檔名部分，跟PrintQueue進行比對
                string pureFileName = Path.GetFileName(filePath);
                //限定最大等待時間
                while (DateTime.Now.CompareTo(timeOut) < 0)
                {
                    if (!printing)
                    {
                        //未開始列印前發現檔名相同的列印工作
                        if (CheckPrintQueue(pureFileName))
                        {
                            printing = true;
                            //        Console.WriteLine($"[{pureFileName}]列印中...");
                        }
                    }
                    else
                    {
                        //已開始列印後，同檔名列印工作消失表示列印完成
                        if (!CheckPrintQueue(pureFileName))
                        {
                            done = true;
                            //     Console.WriteLine($"[{pureFileName}]列印完成");
                            break;
                        }
                    }
                    System.Threading.Thread.Sleep(100);
                }
                try
                {
                    //若程序尚未關閉，強制關閉之
                    if (false == p.CloseMainWindow())
                        p.Kill();
                }
                catch
                {
                }
                if (!done)
                {
                    //    Console.WriteLine($"無法確認報表[{pureFileName}]列印狀態！");
                }
            }
            catch (Exception ex)
            {
                //        Console.WriteLine($"Error: {DateTime.Now:HH:mm:ss} {ex.Message}");
            }
        }

        private static bool CheckPrintQueue(string file)
        {
            string Job_Name = "";
            string searchQuery =
                "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher printJobs = new ManagementObjectSearcher(searchQuery);
            foreach (ManagementObject mo in printJobs.Get())
            {
                Job_Name = mo.Properties["Document"].Value.ToString();
            }

            if (Job_Name == file)
            { return true; }
            else
                return false;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "INVOICENO")
                {
                    string 收貨採購單 = dataGridView1.CurrentRow.Cells["收貨採購單"].Value.ToString();

                    System.Data.DataTable gg1 = GetOPTW(收貨採購單);
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();

                        string aa = path + "\\" + 路徑;

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);

                    }

                }
            }
            catch { }
        }

        public System.Data.DataTable GetOPTWFGG()
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT T1.TrgetEntry ID, U_ACME_PayGUI PAY FROM OPDN T0 LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE ISNULL(U_ACME_PayGUI,'') <> ''   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetOPTW(string docentry)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct cast(t3.TRGTPATH as nvarchar(80))  [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,T3.FILENAME+'.'+Fileext 檔案名稱 from oclg t2     ");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)     ");
            sb.Append(" where  T2.DOCTYPE='20'  ");
            sb.Append(" and   t2.docentry=@docentry and  T3.[FILENAME] LIKE '%INV%' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            catch { }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void button6_Click(object sender, EventArgs e)
        {
                   DialogResult result;
            result = MessageBox.Show("請確認否要列印", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                DELETEFILE();
                for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                {

                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[i];
                    string 收貨採購單 = row.Cells["收貨採購單"].Value.ToString();

                    System.Data.DataTable gg1 = GetOPTW(收貨採購單);
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();

                        string aa = path + "\\" + 路徑;


                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);
                      // System.Diagnostics.Process.Start(aa);
                       // System.Diagnostics.Process.Start(NewFileName);
                       Print(NewFileName);


                    }
                }

                for (int i = dataGridView3.SelectedRows.Count - 1; i >= 0; i--)
                {

                    DataGridViewRow row;

                    row = dataGridView3.SelectedRows[i];
                    string 收貨採購單 = row.Cells["收貨採購單3"].Value.ToString();

                    System.Data.DataTable gg1 = GetOPTW(收貨採購單);
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();

                        string aa = path + "\\" + 路徑;


                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);
                        // System.Diagnostics.Process.Start(aa);
                 //       System.Diagnostics.Process.Start(NewFileName);
                        //   Print(NewFileName);
                        Print(NewFileName);

                    }
                }
            }
     
        }
        private void DELETEFILE()
        {
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "INVOICENO3")
                {
                    string 收貨採購單 = dataGridView3.CurrentRow.Cells["收貨採購單3"].Value.ToString();

                    System.Data.DataTable gg1 = GetOPTW(收貨採購單);
                    if (gg1.Rows.Count > 0)
                    {
                        string path = gg1.Rows[0]["path"].ToString();
                        string 路徑 = gg1.Rows[0]["路徑"].ToString();
                        string 檔案名稱 = gg1.Rows[0]["檔案名稱"].ToString();

                        string aa = path + "\\" + 路徑;

                        string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                        string filename = 檔案名稱;
                        string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                        System.IO.File.Copy(aa, NewFileName, true);

                        System.Diagnostics.Process.Start(NewFileName);

                    }

                }
            }
            catch { }
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\MailTemplates\\CheckMoney.html";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border:3px #F5F5DC groove;'>");
            sb.AppendLine("<tr>");
            for (int i = 0; i < 19; i++)
            {
                //總計
                sb.AppendLine("<td bgcolor=\"#FFFF00\">" + dataGridView3.Rows[0].Cells[i].Value + "</td>");
            }
            sb.AppendLine("</tr>");

            sb.AppendLine("<tr>");
            for (int i = 0; i < 19; i++)
            {
                //欄位名稱
                sb.AppendLine("<td bgcolor=\"#726E6D\"><span style=\"color: white;\">" + dataGridView3.Columns[i].HeaderText + "</span></td>");
            }
            sb.AppendLine("</tr>");


            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (i == 0)
                {
                    continue;
                }
                if (dataGridView3.Rows[i].Cells[0].Value.ToString() == "" || dataGridView3.Rows[i].Cells[0].Value == null)
                {
                    //小計
                    sb.AppendLine("<tr bgcolor=\"#98AFC7\">");
                    for (int j = 0; j < 19; j++)
                    {
                        //只到成本差異原因,所以19
                        sb.AppendLine("<td>" + dataGridView3.Rows[i].Cells[j].Value + "</td>");
                    }
                    sb.AppendLine("</tr>");

                    continue;
                }
                if (i % 2 == 0)
                {
                    sb.AppendLine("<tr bgcolor=\"#C0C0C0\">");
                    for (int j = 0; j < 19; j++)
                    {
                        //只到成本差異原因,所以19
                        sb.AppendLine("<td>" + dataGridView3.Rows[i].Cells[j].Value + "</td>");
                    }
                    sb.AppendLine("</tr>");
                }
                else
                {
                    sb.AppendLine("<tr bgcolor=\"#E5E4E2\">");
                    for (int j = 0; j < 19; j++)
                    {
                        //只到成本差異原因,所以19
                        sb.AppendLine("<td>" + dataGridView3.Rows[i].Cells[j].Value + "</td>");
                    }
                    sb.AppendLine("</tr>");
                }

            }

            sb.AppendLine("</table>");

            template = template.Replace("##Payment##", textBox3.Text);

            template = template.Replace("##Template##", sb.ToString());

            string SlpName = globals.UserID;

            string MailToAddress = "";
            
            string strSubject = textBox3.Text;

            for (int i = 0; i < clbMailAddress.Items.Count; i++) 
            {
                if (clbMailAddress.GetItemChecked(i)) 
                {
                    MailToAddress += clbMailAddress.Items[i].ToString() + ";";
                }
            }

            MailToAddress = MailToAddress.TrimEnd(';');

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
                MessageBox.Show("信件已寄出");
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
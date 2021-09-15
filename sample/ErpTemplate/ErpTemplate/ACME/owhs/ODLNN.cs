using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
namespace ACME
{
    public partial class ODLNN : Form
    {
        string DATE1 = "";
        string DATE2 = GetMenu.Day();
        string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        public ODLNN()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            DELETEFILE2();

            System.Data.DataTable N2 = null;
            try
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                FileName = lsAppDir + "\\Excel\\wh\\放貨單.xls";

                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row2;
                    
                    StringBuilder sb5 = new StringBuilder();
                    StringBuilder sb6 = new StringBuilder();
                    StringBuilder sb7 = new StringBuilder();
                    StringBuilder sb8 = new StringBuilder();
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row2 = dataGridView1.SelectedRows[i];
                        sb8.Append("'" + row2.Cells["ID"].Value.ToString() + " " + row2.Cells["LINENUM"].Value.ToString() + "',");
                        sb5.Append("'" + row2.Cells["ID"].Value.ToString() +  "',");
                        sb6.Append("'" + row2.Cells["WHNO"].Value.ToString() + "',");



                    }


                    sb5.Remove(sb5.Length - 1, 1);
                    sb6.Remove(sb6.Length - 1, 1);
             
                    sb8.Remove(sb8.Length - 1, 1); 
                    DataGridViewRow row;
                    row = dataGridView1.SelectedRows[0];

                    string T1 = "";
                    string T12 = "";
                    if (checkBox1.Checked)
                    {
                        T1 = sb5.ToString();
                        T12 = sb8.ToString();
                    }
                    else
                    {
                        T1 = row.Cells["ID"].Value.ToString();
                        T12 = row.Cells["ID"].Value.ToString();
                    }
                    string ID = row.Cells["ID"].Value.ToString();
                    string TT1 = row.Cells["LINENUM"].Value.ToString();
                    string SALES3 = row.Cells["SALES2"].Value.ToString();
                    string WHNO = row.Cells["WHNO"].Value.ToString();
                    string WHNO2 = "";
                    StringBuilder sb3 = new StringBuilder();
                    StringBuilder sb4 = new StringBuilder();
                    if (checkBox1.Checked)
                    {
                        System.Data.DataTable NWH2 = GetWHNO2(T1);
                        if (NWH2.Rows.Count > 0)
                        {

    
                            WHNO = sb6.ToString();

                            System.Data.DataTable NWH3 = GetWHNO3(sb6.ToString());
                            for (int i = 0; i <= NWH3.Rows.Count - 1; i++)
                            {
                                sb7.Append("'" + NWH3.Rows[i]["WHNO"].ToString() + "'/");
                     
                            }
                            sb7.Remove(sb7.Length - 1, 1);
                            WHNO2 = sb7.ToString();
                        }
           
                    }
                    else
                    {
                        WHNO2 = WHNO;
                    }
                    string Login = fmLogin.LoginID.ToString();
                    string CARDNAME = row.Cells["CARDNAME"].Value.ToString();
                    CARDNAME = CARDNAME.Replace("/", "");
                    string CARDCODE = row.Cells["客戶編號"].Value.ToString();
                    string BU = row.Cells["BU"].Value.ToString();
                    string OWHS = row.Cells["倉庫"].Value.ToString();
                    System.Data.DataTable N1 = Getprepare2(WHNO);
                    System.Data.DataTable NPO = GetPO1(WHNO);
                    System.Data.DataTable NPO2 = GetPO2(WHNO);
                    System.Data.DataTable NPO3 = GetPO3(WHNO);
                    string FLOW = "";
                    string FLOW2 = "";
                    if (BU == "LED")
                    {
                        FLOW = "異常放貨流程(LED)";
                        FLOW2 = "LED總經理審核";
                    }
                    else
                    {
                        FLOW = "異常放貨流程(TFT)";
                        FLOW2 = "TFT總經理審核";
                    }

                    System.Data.DataTable J2 = GetMAN(ID, FLOW, FLOW2);
                    System.Data.DataTable J3 = GetSALES(ID);
                    System.Data.DataTable J4 = GetWHP(Login);
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();
                    for (int i = 0; i <= N1.Rows.Count - 1; i++)
                    {

                        DataRow d = N1.Rows[i];

                        sb.Append("'" + d["SAPDOC"].ToString() + "',");

                        sb2.Append(d["SAPDOC"].ToString() + "/");

                    }
                    sb.Remove(sb.Length - 1, 1);
                    sb2.Remove(sb2.Length - 1, 1);
                    string T2 = sb.ToString();
                    string T3 = sb2.ToString();

                    string OutPutFile = "";
                    System.Data.DataTable ty = GetPO(WHNO);
                    ViewDATE(WHNO);
                    System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                    System.Data.DataTable H2 = GetQTY2(T1);
                    if (H2.Rows.Count > 0)
                    {
                        string QTY = H2.Rows[0][0].ToString();
                        if (H1.Rows.Count > 0)
                        {
                            string OWHS1 = "";
                            int LEN = OWHS.Length;
                            if (LEN<=3)
                            {
                                OWHS1 = OWHS.Trim();
                            }
                            else
                            {
                                OWHS1 = OWHS.Trim().Replace("倉", "").Replace("-", "");
                            }


                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                              OWHS1 + "放貨單(" + CARDNAME + ")--" + QTY + "PCS.xls";

                        }
                        else
                        {
               
                                OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                  DATE2 + CARDNAME + "放貨單--" + QTY + ".xls";
                            
                        }



                    }
                    StringBuilder sbb = new StringBuilder();
                    string j4 = "";
                    if (ty.Rows.Count > 0)
                    {
                        for (int i = 0; i <= ty.Rows.Count - 1; i++)
                        {

                            DataRow d = ty.Rows[i];


                            sbb.Append("PO#" + d["po"].ToString() + "*" + d["quantity"].ToString() + "片/");


                        }

                        sbb.Remove(sbb.Length - 1, 1);
                        j4 = sbb.ToString();
                    }

                    string gj = j4 +
         Environment.NewLine + "第1項為FOC" +
          Environment.NewLine + "AUS-INVOICE NO#";
    
                    string MANAGER = J2.Rows[0][0].ToString().ToUpper();
                    string SALES = J3.Rows[0][0].ToString().ToUpper();


                    if (SALES == "業務-許建峰")
                    {
                        SALES = "業務-許心如";
                    }
                    string WHP = "製單: " + J4.Rows[0][0].ToString();
                    int h1 = SALES.Length;
                    string SALES2 = SALES.Substring(3, h1 - 3);
                    string AR = "";
               
                    string 發票日期 = "";
                    string 發票號碼 = "";
                    string 發票聯式 = "";
                    DataTable DTAR = GetAR(WHNO);
                    if (DTAR.Rows.Count > 0)
                    {
                        AR = DTAR.Rows[0]["AR"].ToString();

                        發票日期 = DTAR.Rows[0]["發票日期"].ToString();
                        發票號碼 = DTAR.Rows[0]["發票號碼"].ToString();
                        發票聯式 = DTAR.Rows[0]["發票聯式"].ToString();
                    }


                    //Excel的樣版檔
                    string ExcelTemplate = FileName;
         

                    //輸出檔


                    string B1 = "//acmew08r2ap//table//SIGN//MANAGER//";
                    string B2 = "//acmew08r2ap//table//SIGN//SALES//";

                    string bb = "";

                        bb = "進金生實業股份有限公司";

                    string WHNONAME = "放貨單" + WHNO2;
                   
                    if (NPO.Rows.Count> 0)
                    {
                        string AA = CARDNAME.Substring(0, 2);
                        N2 = Getprepare2S1(WHNO, sb2.ToString(), WHNONAME, "銷售訂單:", bb, DATE1, 發票聯式, 發票日期, 發票號碼, gj, T1,T12);
                        if (N2.Rows.Count > 0)
                        {


                            if (NPO2.Rows.Count == 0)
                            {

                                ExcelReport.ExcelHelenPIC(N2, lsAppDir + "\\Excel\\wh\\放貨單南京2.xls", OutPutFile, AA, B1 + MANAGER + ".JPG", B2 + SALES + ".JPG", "B", "Y", "Y");
                            }
                            else if (NPO3.Rows.Count == 0)
                            {

                                ExcelReport.ExcelHelenPIC(N2, lsAppDir + "\\Excel\\wh\\放貨單南京3.xls", OutPutFile, AA, B1 + MANAGER + ".JPG", B2 + SALES + ".JPG", "C", "Y", "Y");
                            }
                            else
                            {

                                ExcelReport.ExcelHelenPIC(N2, lsAppDir + "\\Excel\\wh\\放貨單南京.xls", OutPutFile, AA, B1 + MANAGER + ".JPG", B2 + SALES + ".JPG", "A", "Y", "Y");
                            }

                            UpdateAPLC5(WHNO);
                            UpdateCHECK(WHNO);
                        }
                        else
                        {
                            MessageBox.Show("沒有資料");
                        
                        }
                    }
                    else
                    {

                        N2 = Getprepare2S(WHNO, sb2.ToString(), WHNONAME, "銷售訂單:", bb, DATE1, 發票聯式, 發票日期, 發票號碼, gj, T1, T12);
                        if (N2.Rows.Count > 0)
                        {
                            ExcelReport.ExcelReportOutputLA(N2, lsAppDir + "\\Excel\\wh\\放貨單.xls", OutPutFile, B1 + MANAGER + ".JPG", B2 + SALES + ".JPG","Y");
                            UpdateCHECK(WHNO);
                            UpdateAPLC5(WHNO);
                        }
                        else
                        {
                            MessageBox.Show("沒有資料");

                        }
                    }

             

  

                }




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void UpdateAPLC5(string SHIPPINGCODE)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update wh_main set add5=@aa where shippingcode=@bb AND add5 <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@aa", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@bb", SHIPPINGCODE));

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

        private System.Data.DataTable GetWHNO2(string ID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append(" SELECT DISTINCT WHNO  FROM ACME_ODLNN1 WHERE  ID IN (" + ID + ") ");
            }
            else
            {
                sb.Append(" SELECT DISTINCT WHNO  FROM ACME_ODLNN1 WHERE  ID=@ID ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetWHNO3(string ID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append(" SELECT DISTINCT WHNO  FROM ACME_ODLNN1 WHERE  WHNO IN (" + ID + ") ");
            }
            else
            {
                sb.Append(" SELECT DISTINCT WHNO  FROM ACME_ODLNN1 WHERE  ID=@ID ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        private System.Data.DataTable GetQTY2(string ID)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append(" SELECT SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) FROM ACME_ODLNN1 WHERE  ID IN (" + ID + ") ");
            }
            else
            {
                sb.Append(" SELECT SUM(CAST(ISNULL(QUANTITY,0) AS DECIMAL)) FROM ACME_ODLNN1 WHERE  ID=@ID ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }
        public  System.Data.DataTable GetPO(string shippingcode)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append(" select distinct po,quantity from wh_item where shippingcode in (" + shippingcode + ") and isnull(po,'') <> '' ");

            }
            else
            {
                sb.Append(" select distinct po,quantity from wh_item where shippingcode=@shippingcode and isnull(po,'') <> '' ");
              
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public System.Data.DataTable GetPOS(string shippingcode)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append(" select createname from WH_MAIN where shippingcode in (" + shippingcode + ") ");

            }
            else
            {
                sb.Append(" select createname from WH_MAIN where shippingcode=@shippingcode ");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(FileName1);
                string[] filenames2 = Directory.GetDirectories(FileName1);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }


            }
            catch { }
        }
        private void DELETEFILE2()
        {
            try
            {
                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string FileName1 = lsAppDir1 + "\\Excel\\temp\\wh";
                string[] filenames = Directory.GetFiles(FileName1);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }


            }
            catch { }
        }

        public static System.Data.DataTable GetMAN(string ID, string FLOW_DESC, string S_STEP_ID)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT  S_USER_ID  FROM dbo.SYS_TODOLIST WHERE FLOW_DESC=@FLOW_DESC AND S_STEP_ID=@S_STEP_ID ");
            sb.Append(" and ltrim(SUBSTRING(FORM_PRESENTATION,5,10)) = @ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@FLOW_DESC", FLOW_DESC));
            command.Parameters.Add(new SqlParameter("@S_STEP_ID", S_STEP_ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetWHP(string hometel)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select lastname+firstname from ohem where  hometel=@hometel ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@hometel", hometel));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable Getprepare2S(string docentry, string cc, string ee, string AR發票, string bb, string DATE, string 發票聯式, string 發票日期, string 發票號碼, string gj, string id, string id2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select  t3.SAPDOC 單號,t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話                    ");
            sb.Append(" ,shipment 送貨地址,t0.bucntctprsn 業務人員,t3.itemcode 產品編號,");
            sb.Append(" (case when itmsgrpcod =1032 AND SUBSTRING(T3.itemcode,1,4) <> 'ACME' then T0.frgnname else T0.dscription end) 品名規格 ,                 ");
            sb.Append(" t0.pino 料號,T3.QUANTITY 出貨數量,T0.倉管,T0.倉別,t0.ARType 發票方式, '客戶名稱:'+t0.cardname 銷售客戶名稱,");
            sb.Append(" t0.grade 等級,t0.ver 版本,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO ,");
            sb.Append(" t0.cardcode PCS,ROHS='ROHS',AU='AUS' ,''  PO料號,'' PO,SEQNO   ");
            sb.Append(" ,@DATE 日期,@gj 備註,@ee 文件名稱,@發票聯式 發票聯式,@發票日期 發票日期,@發票號碼 發票號碼");
            sb.Append(" ,AR=@cc,單據=@ee,AR發票=@AR發票,公司=@bb                   ");
            sb.Append(" from acmesqleep.dbo.acme_odlnn1 T3  LEFT JOIN ( SELECT T0.shippingcode,T1.ITEMCODE,MAX(T1.frgnname) frgnname,MAX(T1.dscription) dscription,");
            sb.Append(" MAX(t1.docentry) docentry,MAX(t0.cardname)  cardname,MAX(arriveday) arriveday  ,MAX(cfs) cfs,MAX(shipment) shipment,MAX(bucntctprsn) bucntctprsn,	");
            sb.Append(" MAX(T1.pino) pino,'製單: '+MAX(T0.createName) 倉管,MAX(shipping_obu) 倉別,MAX(ARType) ARType,MAX(grade) grade,MAX(ver) ver,MAX(oBUBillTo) oBUBillTo,MAX(oBUShipTo) oBUShipTo ,MAX(t1.cardcode) cardcode,MAX(t1.SEQNO) SEQNO   ");
            sb.Append(" FROM wh_main T0  left join wh_item t1 on (t0.shippingcode=t1.shippingcode) ");
            sb.Append(" GROUP BY T0.shippingcode,T1.ITEMCODE) T0 ON (T3.WHNO=T0.shippingcode AND T3.ITEMCODE=T0.ITEMCODE)");
            sb.Append(" INNER join acmesql02.dbo.oitm t2 on (T3.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS)");

            if (checkBox1.Checked)
            {
                sb.Append(" where   cast(t3.id as varchar)+' '+cast(T3.LINENUM as varchar)  in (" + id2 + ")  ORDER BY   t3.id,T3.LINENUM");

            }
            else
            {
                sb.Append("       where t3.WHNO=@aa and t3.id=@id  ORDER BY  T3.LINENUM ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            command.Parameters.Add(new SqlParameter("@發票聯式", 發票聯式));
            command.Parameters.Add(new SqlParameter("@發票日期", 發票日期));
            command.Parameters.Add(new SqlParameter("@發票號碼", 發票號碼));
            command.Parameters.Add(new SqlParameter("@gj", gj));
            command.Parameters.Add(new SqlParameter("@id", id));
            command.Parameters.Add(new SqlParameter("@id2", id2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable Getprepare2S1(string docentry, string cc, string ee, string AR發票, string bb, string DATE, string 發票聯式, string 發票日期, string 發票號碼, string gj, string id, string id2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select distinct itemremark 單據類別,t1.docentry 單號, T3.LINENUM");
            sb.Append(" ,@DATE 日期,t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話");
            sb.Append("         ,shipment 送貨地址,t0.bucntctprsn 業務人員,t1.itemcode 產品編號,case when itmsgrpcod =1032 AND SUBSTRING(t1.itemcode,1,4) <> 'ACME' THEN t1.frgnname else t1.dscription end 品名規格");
            sb.Append("       , t1.pino 料號,T3.QUANTITY 出貨數量,t1.nowqty 現有數量,@gj 備註,'製單: '+createName 倉管,shipping_obu 倉別,");
            sb.Append("         @ee 文件名稱,t0.ARType 發票方式,@發票聯式 發票聯式,@發票日期 發票日期,@發票號碼 發票號碼,t1.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("           '客戶名稱:'+t0.cardname 銷售客戶名稱,t1.grade 等級,t1.ver 版本,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO,AR=@cc,單據=@ee,AR發票=@AR發票,公司=@bb ");
            sb.Append("   ,''''+U_CUSTITEMCODE  PO料號,U_CUSTDOCENTRY PO");
            sb.Append("          from wh_main t0 left join wh_item t1 on (t0.shippingcode=t1.shippingcode) ");
            sb.Append(" left join acmesql02.dbo.oitm t2 on (t1.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("  LEFT JOIN acmesqleep.dbo.acme_odlnn1 T3 ON (T1.shippingcode=T3.WHNO AND T1.SEQNO=T3.WHLINE) ");
            if (checkBox1.Checked)
            {
                sb.Append(" where  cast(t3.id as varchar)+' '+cast(T3.LINENUM as varchar)  in (" + id2 + ")  ORDER BY   T3.LINENUM");
            }
            else
            {
                sb.Append(" where t0.shippingcode=@aa  AND id=@id ORDER BY  T3.LINENUM");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@cc", cc));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            command.Parameters.Add(new SqlParameter("@發票聯式", 發票聯式));
            command.Parameters.Add(new SqlParameter("@發票日期", 發票日期));
            command.Parameters.Add(new SqlParameter("@發票號碼", 發票號碼));
            command.Parameters.Add(new SqlParameter("@gj", gj));
            command.Parameters.Add(new SqlParameter("@id", id));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable Getprepare22(string docentry, string bb, string ee, string ff, string gg, string AR發票, string DATE, string 船務, string id, string 備註, string FAX,string id2)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("               select distinct  T3.LINENUM,");
            sb.Append(" t0.cardname 客戶名稱,arriveday 聯絡人員,cfs 聯絡電話,SEQNO");
            sb.Append("      ,t0.shipment 送貨地址,t0.bucntctprsn 業務人員,t0.itemcode 產品編號,T0.品名規格,");
            sb.Append("      t0.pino 料號,T3.QUANTITY 出貨數量,T0.倉管,T0.倉別,");
            sb.Append("      '放貨單'+t0.shippingcode 文件名稱,t0.ARType 發票方式,t0.ARTyp2 發票聯式,ARDate 發票日期,ARNumber 發票號碼,t0.cardcode PCS,ROHS='ROHS',AU='AUS',");
            sb.Append("        '客戶名稱:'+t0.cardname 銷售客戶名稱,t0.grade 等級,t0.ver 版本,'BILL TO: '+oBUBillTo BILLTO,'SHIP TO: '+oBUShipTo SHIPTO ");
            sb.Append(" ,備註=@備註,@DATE 日期,表頭=@dd,T1=@ee,T2=@ff,T3=@gg,單號=@AR發票,船務=@船務,FAX=@FAX");
            sb.Append("                      from acmesqleep.dbo.acme_odlnn1 T3 ");
            sb.Append(" LEFT JOIN (");
            sb.Append(" SELECT T0.shippingcode,T1.ITEMCODE,MAX(t1.docentry) docentry,MAX(t0.cardname)  cardname,MAX(arriveday) arriveday ");
            sb.Append(" ,MAX(cfs) cfs,MAX(shipment) shipment,MAX(bucntctprsn) bucntctprsn,MAX(case when itmsgrpcod =1032 AND SUBSTRING(t1.itemcode,1,4) <> 'ACME' then t1.frgnname else t1.dscription end) 品名規格");
            sb.Append(" ,MAX(T1.pino) pino,'製單: '+MAX(T0.createName) 倉管,MAX(shipping_obu) 倉別,MAX(ARType) ARType,MAX(grade) grade,MAX(ver) ver,MAX(oBUBillTo) oBUBillTo,MAX(oBUShipTo) oBUShipTo");
            sb.Append(" ,MAX(t1.cardcode) cardcode,MAX(t1.SEQNO) SEQNO,MAX(ARNumber) ARNumber,MAX(ARDate) ARDate,MAX(ARTyp2) ARTyp2   FROM wh_main T0");
            sb.Append("  left join wh_item t1 on (t0.shippingcode=t1.shippingcode) ");
            sb.Append(" left join acmesql02.dbo.oitm t2 on (t1.itemcode=t2.itemcode COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            if (checkBox1.Checked)
            {
                sb.Append(" WHERE  T0.ShippingCode IN (" + docentry + ")");

            }
            sb.Append(" GROUP BY T0.shippingcode,T1.ITEMCODE) T0 ON (T3.WHNO=T0.shippingcode AND T3.ITEMCODE=T0.ITEMCODE)");
     
            //    sb.Append(" where t0.shippingcode=@aa  AND id=@id ORDER BY  T3.LINENUM");

                if (checkBox1.Checked)
                {
                    sb.Append(" where  cast(t3.id as varchar)+' '+cast(T3.LINENUM as varchar)  in (" + id2 + ")  ORDER BY   T3.LINENUM");
                  //  sb.Append(" where   t3.id=@id2  ORDER BY T3.LINENUM");

                }
                else
                {
                    sb.Append("       where T3.WHNO=@aa and id=@id  ORDER BY  T3.LINENUM ");
                }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@dd", bb));
            command.Parameters.Add(new SqlParameter("@ee", ee));
            command.Parameters.Add(new SqlParameter("@ff", ff));
            command.Parameters.Add(new SqlParameter("@gg", gg));
            command.Parameters.Add(new SqlParameter("@AR發票", AR發票));
            command.Parameters.Add(new SqlParameter("@DATE", DATE));
            command.Parameters.Add(new SqlParameter("@船務", 船務));
            command.Parameters.Add(new SqlParameter("@id", id));
            command.Parameters.Add(new SqlParameter("@備註", 備註));
            command.Parameters.Add(new SqlParameter("@FAX", FAX));
            command.Parameters.Add(new SqlParameter("@id2", id2));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static System.Data.DataTable GetAR(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select T5.DOCENTRY AR,Convert(varchar(10),T6.u_in_bsdat,111) as 發票日期,T6.u_in_bsinv as 發票號碼,");
            sb.Append("                                      發票聯式 = case T6.u_in_bsty1");
            sb.Append("                                  when '0' then '三聯式發票'  when '1' then '三聯式收銀機發票' ");
            sb.Append("                            when '2' then '二聯式發票' when '3' then '二聯式收銀機發票'  ");
            sb.Append("                              when '4' then '電子計算機發票' when '5' then '免用統一發票' end from   wh_item2 t1 ");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.RDR1 T2 ON (T1.DOCENTRY=T2.DOCENTRY AND T1.LINENUM=T2.LINENUM)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.ORDR T3 ON (T2.DOCENTRY=T3.DOCENTRY)");
            sb.Append(" left join ACMESQL02.DBO.dln1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum  and t4.basetype='17')");
            sb.Append(" left join ACMESQL02.DBO.INV1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum  and t5.basetype='15')");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OINV T6 ON (T5.DOCENTRY=T6.DOCENTRY) WHERE T1.SHIPPINGCODE=@SHIPPINGCODE");




            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public static DataTable Getocrdnew2(string docentry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql;


            sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode billzipcode FROM acmesql02.dbo.ORDR T0 LEFT JOIN  acmesql02.dbo.CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  acmesql02.dbo.CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')   where t0.docentry in (" + docentry + ")  ";


            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ordr");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ordr"];
        }


        private System.Data.DataTable Getprepareend(int  FLAG)
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select DISTINCT T0.ID,T1.LINENUM,ltrim(rtrim(A10)) 金額,CASE WHEN ISNULL(T1.WHNO,'') ='' THEN T0.WHNO ELSE T1.WHNO END WHNO  ,SALES2,SHIPPING_OBU 倉庫,CARDNAME,T0.SA BU,CARDCODE 客戶編號,T2.UPDATE_DATE 簽核日期,UPDATE_TIME 簽核時間,T1.ITEMCODE 產品編號,T1.QUANTITY 數量,T1.CHECKED 已匯出,T3.createName    from ACME_ODLNN T0   ");
            sb.Append(" LEFT JOIN ACME_ODLNN1 T1 ON (T0.ID=T1.ID)  ");
            sb.Append(" LEFT JOIN (  ");
            sb.Append(" SELECT MAX(UPDATE_TIME) UPDATE_TIME,MAX(UPDATE_DATE) UPDATE_DATE,ltrim(SUBSTRING(FORM_PRESENTATION,5,10)) ID FROM  acmesqleep.dbo.SYS_TODOLIST    ");
            sb.Append(" WHERE  flow_desc ='異常放貨流程(TFT)'   ");
            sb.Append(" GROUP BY ltrim(SUBSTRING(FORM_PRESENTATION,5,10))  ");
            sb.Append(" ) T2 ON (T0.ID=T2.ID) ");
            sb.Append(" LEFT JOIN ( SELECT SHIPPING_OBU,SHIPPINGCODE,createName FROM ACMESQLSP.DBO.WH_MAIN) T3 ON (CASE WHEN ISNULL(T1.WHNO,'') ='' THEN T0.WHNO ELSE T1.WHNO END=T3.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.FLOWFLAG='Z'  ");
            if (FLAG == 1)
            {
             sb.Append("  AND CAST(T2.UPDATE_DATE AS DATETIME)    BETWEEN  DATEADD(M,-1,GETDATE()) and GETDATE() ");
            }
            else
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    sb.Append("  and T0. ID BETWEEN @AA AND @BB ");

                }

                if (textBox3.Text != "")
                {
                    sb.Append("  and CASE WHEN ISNULL(T1.WHNO,'') ='' THEN T0.WHNO ELSE T1.WHNO END =@CC ");

                }
            }
            sb.Append(" order by UPDATE_DATE desc,UPDATE_TIME DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox3.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public  System.Data.DataTable Getprepare2(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            if (checkBox1.Checked)
            {
                sb.Append(" SELECT DISTINCT DOCENTRY SAPDOC FROM wh_item WHERE SHIPPINGCODE IN (" + ID + ") ");
            }
            else
            {
                sb.Append(" SELECT DISTINCT DOCENTRY SAPDOC FROM wh_item WHERE SHIPPINGCODE = @ID ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public  System.Data.DataTable GetPO1(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ") AND ISNULL(U_CUSTITEMCODE,'')+ISNULL(U_CUSTDOCENTRY,'') <> '' ");
            }
            else
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(U_CUSTITEMCODE,'')+ISNULL(U_CUSTDOCENTRY,'') <> '' ");
            }



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        public System.Data.DataTable GetPO2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ") AND ISNULL(U_CUSTITEMCODE,'')<> '' ");
            }
            else
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(U_CUSTITEMCODE,'')<> ''  ");
            }



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public System.Data.DataTable GetPO3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (checkBox1.Checked)
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + ")  AND ISNULL(U_CUSTDOCENTRY,'')<> '' ");
            }
            else
            {
                sb.Append("  SELECT SHIPPINGCODE FROM wh_item  WHERE SHIPPINGCODE=@SHIPPINGCODE  AND ISNULL(U_CUSTDOCENTRY,'')<> '' ");
            }



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable Getprepare2G(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT invoice INV FROM wh_item WHERE SHIPPINGCODE=@ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
      
        public static System.Data.DataTable GetSALES(string ID)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select CASE WHEN DESCRIPTION LIKE '%中國%' THEN '業務-許心如' ELSE GROUPNAME END GROUPNAME from ACME_ODLNN T0 LEFT JOIN GROUPS T1 ON (T0.SALES=T1.GROUPID) where ID=@ID ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ID", ID));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }




        private void ODLNN_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getprepareend(1);

        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Getprepareend(0);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            DELETEFILE2();

            System.Data.DataTable N2 = null;
            try
            {



                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                DataGridViewRow row2;
                StringBuilder sb8 = new StringBuilder();
                StringBuilder sb9 = new StringBuilder();
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row2 = dataGridView1.SelectedRows[i];
                        sb9.Append("'" + row2.Cells["ID"].Value.ToString() + " " + row2.Cells["LINENUM"].Value.ToString() + "',");
                        sb8.Append("'" + row2.Cells["WHNO"].Value.ToString() + "',");

                    }
                    sb8.Remove(sb8.Length - 1, 1);
                    sb9.Remove(sb9.Length - 1, 1);

      
          
                    DataGridViewRow row;

                    row = dataGridView1.SelectedRows[0];
                    string OutPutFile = "";
                    string T1 = row.Cells["ID"].Value.ToString();
                    string TT1 = row.Cells["LINENUM"].Value.ToString();
                    string SALES3 = row.Cells["SALES2"].Value.ToString();
                    string Login = fmLogin.LoginID.ToString();
                    string WHNO = row.Cells["WHNO"].Value.ToString();
                    string CARDCODE = row.Cells["客戶編號"].Value.ToString();
                    string CARDNAME = row.Cells["CARDNAME"].Value.ToString();
                    CARDNAME = CARDNAME.Replace("/", "");
                    string OWHS = row.Cells["倉庫"].Value.ToString();
                    System.Data.DataTable H1 = GetMenu.GetOWHS3(OWHS);
                    System.Data.DataTable H2 = GetQTY2(T1);

                    if (H2.Rows.Count > 0)
                    {
                        string QTY = H2.Rows[0][0].ToString();
                        if (H1.Rows.Count > 0)
                        {
                            int LEN = OWHS.Length;
                            string OWHS1 = "";
                            if (LEN <=3)
                            {
                                OWHS1 = OWHS.Trim();
                            }
                            else
                            {
                                OWHS1 = OWHS.Trim().Replace("倉", "").Replace("-", "");
                            }


                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                              OWHS1 + "放貨單(" + CARDNAME + ")--" + QTY + "PCS.xls";

                        }
                        else
                        {

                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                       DateTime.Now.ToString("yyyyMMdd") + CARDNAME + Path.GetFileName(FileName) + ".xls";
                        }



                    }
                    string T12 = "";
                    if (checkBox1.Checked)
                    {
                        WHNO = sb8.ToString();
                        T12 = sb9.ToString();
                    }
                    System.Data.DataTable N1 = Getprepare2(WHNO);
                    System.Data.DataTable N1G = Getprepare2G(WHNO);
                    System.Data.DataTable J3 = GetSALES(T1);

                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();
                    StringBuilder sb5 = new StringBuilder();

                    string T2 = "";
                    string T3 = "";
                    string T5 = "";
                    string SALES  = "業務-許心如";

                    if (N1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= N1.Rows.Count - 1; i++)
                        {

                            DataRow d = N1.Rows[i];

                            sb.Append("'" + d["SAPDOC"].ToString() + "',");

                            sb2.Append(d["SAPDOC"].ToString() + "/");


                        }

                        sb.Remove(sb.Length - 1, 1);
                        sb2.Remove(sb2.Length - 1, 1);

                        T2 = sb.ToString();
                        T3 = sb2.ToString();
                    }

                    if (N1G.Rows.Count > 0)
                    {

                        for (int i = 0; i <= N1G.Rows.Count - 1; i++)
                        {

                            DataRow d = N1G.Rows[i];



                            sb5.Append(d["INV"].ToString() + "/");
                        }

                        sb5.Remove(sb5.Length - 1, 1);

                        T5 = "AUS-INV#" + sb5.ToString();

                    }


                    

                    string JOBNO = "";
                    //船務jobno
                    System.Data.DataTable dt3 = GetSHIP(WHNO);
                    if (dt3.Rows.Count > 0)
                    {
                        if (dt3.Rows[0]["ITEMREMARK"].ToString() == "銷售訂單")
                        {

                            StringBuilder sb4 = new StringBuilder();
                            StringBuilder sb3 = new StringBuilder();
                            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
                            {
                                string DOCENTRY = dt3.Rows[i]["DOCENTRY"].ToString();
                                string LINENUM = dt3.Rows[i]["LINENUM"].ToString();
                                sb4.Append("'" + DOCENTRY +' '+ LINENUM + "',");

                            }
                            sb4.Remove(sb4.Length - 1, 1);
                            string A = sb4.ToString();
                            System.Data.DataTable SS = GetSH(A);
                            if (SS.Rows.Count > 0)
                            {
                                for (int i = 0; i <= SS.Rows.Count - 1; i++)
                                {
                                    string CODE = SS.Rows[i]["CODE"].ToString();

                                    sb3.Append(CODE + ",");

                                }
                                sb3.Remove(sb3.Length - 1, 1);
                                JOBNO = sb3.ToString();
                            }

                        }
                    }
                    string bb = "";
                    string ee = "";
                    string ff = "";
                    string gg = "";
                    //1349-00
                    //if (CARDCODE.Trim() == "0511-00" )
                    //{
                    //    bb = "CHOICE INT'L LTD";
                    //    ee = "60 market squarem p.o. box364,Belize";
                    //    ff = "BELIZE CITY,BELIZE";
                    //    gg = "";
                    //    SALES = "CHO" + SALES + G;
                    //}
                    //else if (CARDCODE.Trim() == "1349-00")
                    //{
                    //    bb = "Infinite Power Group Inc.";
                    //    ee = "60 market squarem p.o. box364,Belize";
                    //    ff = "BELIZE CITY,BELIZE";
                    //    gg = "";
                    //    SALES = "INF" + SALES + G;
                    //}
                    //else if (CARDCODE.Trim() == "1349-00")
                    //{
                    //    bb = "TOP GARDEN INT'L LTD";
                    //    ee = "60 market squarem p.o. box364,Belize";
                    //    ff = "BELIZE CITY,BELIZE";
                    //    gg = "";
                    //    SALES = "CHO" + SALES;
                    //}
                    //else
                    //{
                        bb = "進金生實業股份有限公司";
                        ee = "5F.-3, No.257, Sinhu 2nd Rd.,";
                        ff = "Nei-hu District, Taipei Taiwan";
                        gg = "TEL:886-2-8791-2868 FAX:886-02-8791-2869";
                        SALES = "業務-許心如吳昭憲Tony";
               //     }

                    ViewDATE(WHNO);
                    string B2 = "//acmew08r2ap//table//SIGN//SALESOUT//";
                 
                    //Excel的樣版檔
                    string ExcelTemplate = FileName;
                    string FAX = "";
                    if (globals.DBNAME == "達睿生")
                    {
                        FAX = "*請於收到貨後,簽回傳真至FAX:0755-25911201,謝謝!";
                    }
                    else
                    {
                        FAX = "*請於收到貨後,簽回傳真至FAX:+886-2-8791-2869,謝謝!";
                    }

                    N2 = Getprepare22(WHNO, bb, ee, ff, gg, sb2.ToString(), DATE1, JOBNO, T1, T5, FAX, T12);
                        ExcelReport.ExcelReportOutputLA2(N2, lsAppDir + "\\Excel\\wh\\進金生.xls", OutPutFile, B2 + SALES + ".JPG");
                        UpdateCHECK(WHNO);
                        UpdateAPLC5(WHNO);
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public System.Data.DataTable GetSHIP(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT ITEMREMARK,DOCENTRY,LINENUM FROM WH_ITEM4 T1 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ITEMREMARK in ('銷售訂單') ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_item"];
        }
        private System.Data.DataTable GetSH(string DocEntry)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT SHIPPINGCODE CODE from shipping_item T0");
            sb.Append("  where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) ");
            sb.Append(" IN (" + DocEntry + ") AND ITEMREMARK='銷售訂單' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            return ds.Tables[0];

        }

        private void UpdateCHECK(string WHNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            if (checkBox1.Checked)
            {
                sb.Append(" UPDATE  ACME_ODLNN1 SET CHECKED='True'  WHERE WHNO IN (" + WHNO + ")  ");
            }
            else
            {
                sb.Append(" UPDATE  ACME_ODLNN1 SET CHECKED='True'  WHERE WHNO=@WHNO ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@WHNO", WHNO));

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
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
             
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\wh\\異常放貨.xls";


                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                DateTime before2month = DateTime.Now.AddMonths(-1);
                string dd = before2month.ToString("yyyy");
                string d2 = before2month.ToString("yyyyMM");
                System.Data.DataTable k1 = GetODLNN(d2);
                System.Data.DataTable k2 = GetODLNN2(dd);
                //產生 Excel Report
                ExcelReport.ODLNN(k2, ExcelTemplate, OutPutFile, k1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable GetWHP1(string hometel)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" select lastname+firstname from ohem where  hometel=@hometel ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@hometel", hometel));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetODLNN(string DATE)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CASE 欄位 WHEN '2' THEN '' ELSE SALES2 END 業務,SALES2,CARDNAME 客戶名稱,異常筆數,異常金額 FROM (");
            sb.Append(" SELECT '1' 欄位,SALES2,CARDNAME,COUNT(*) 異常筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 異常金額 FROM ACME_ODLNN T0");
            sb.Append(" WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z' and substring(DOCDATE,1,6)=@DATE   GROUP BY  SALES2,CARDNAME");
            sb.Append("              UNION ALL");
            sb.Append("              SELECT '3','','總計',COUNT(*) 筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 金額 FROM ACME_ODLNN T0");
            sb.Append("              WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z' and substring(DOCDATE,1,6)=@DATE ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2',SALES2,'小計',COUNT(*) 筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 金額 FROM ACME_ODLNN T0");
            sb.Append(" WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z' and substring(DOCDATE,1,6)=@DATE  GROUP BY  SALES2 ) AS A");
            sb.Append(" ORDER BY SALES2 DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DATE", DATE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        public static System.Data.DataTable GetODLNN2(string DATE)
        {
            string strCn = "Data Source=acmesap;Initial Catalog=AcmeSqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CASE 欄位 WHEN '2' THEN '' ELSE SALES2 END 業務,SALES2,CARDNAME 客戶名稱,異常筆數,異常金額 FROM (");
            sb.Append(" SELECT '1' 欄位,SALES2,CARDNAME,COUNT(*) 異常筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 異常金額 FROM ACME_ODLNN T0");
            sb.Append(" WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z' and substring(DOCDATE,1,4)=@DATE   GROUP BY  SALES2,CARDNAME");
            sb.Append(" UNION ALL");
            sb.Append("  SELECT '3','','總計',COUNT(*) 筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 金額 FROM ACME_ODLNN T0");
            sb.Append("  WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z'  and substring(DOCDATE,1,4)=@DATE ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '2',SALES2,'小計',COUNT(*) 筆數,SUM(CAST(REPLACE(A10,',','') AS DECIMAL)) 金額 FROM ACME_ODLNN T0");
            sb.Append(" WHERE ISNULL(A10,'0') <> '0'  AND FLOWFLAG='Z'  and substring(DOCDATE,1,4)=@DATE GROUP BY  SALES2 ) AS A");
            sb.Append(" ORDER BY SALES2 desc");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DATE", DATE));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }




        private void ViewDATE(string SHIPPINGCODE)
        {


            System.Data.DataTable N1 = GetORDRDATE(SHIPPINGCODE);

            if (N1.Rows.Count > 0)
            {
                DataRow drw3 = N1.Rows[0];
                DATE1 = drw3["DATE"].ToString();
            }
            else
            {
                DATE1 = DateTime.Now.ToString("yyyy") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");
            }

            DATE2 = DATE1.Replace("/", "");
        }



        public static System.Data.DataTable GetORDRDATE(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select CASE U_ACME_WORKDAY WHEN '內銷' THEN  CONVERT(VARCHAR(10) ,u_ACME_SHIPDAY, 111 ) WHEN '進口轉內銷' THEN  CONVERT(VARCHAR(10) ,u_ACME_SHIPDAY, 111 ) ");
            sb.Append(" ELSE CONVERT(VARCHAR(10) ,GETDATE(), 111 ) END DATE");
            sb.Append("   FROM RDR1 T0");
            sb.Append(" INNER JOIN ACMESQLSP.DBO.WH_ITEM T1 ON (T0.DOCENTRY=T1.DOCENTRY AND T0.LINENUM=T1.LINENUM)");
            sb.Append(" WHERE T1.SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
    }
}
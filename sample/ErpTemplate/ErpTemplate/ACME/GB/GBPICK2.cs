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
    public partial class GBPICK2 : Form
    {

        System.Data.DataTable dtCost = null;
        int scrollPosition = 0;
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string PICKDAY1 = "";
        string PICKDAY2 = "";
        string PICKDAY3 = "";
        string PP1 = "";
        int RR = 0;
        public GBPICK2()
        {
            InitializeComponent();
        }

    
        private System.Data.DataTable DT()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT GBTYPE 批發,SHIPPINGCODE 撿貨單號,DOCDATE 撿貨日期 from GB_PICK  WHERE ISNULL(CHECKED,'') <> 'Checked' and substring(docdate,1,4)>year(getdate())-2  ORDER BY SHIPPINGCODE DESC ");
        
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT2(string SHIPPINGCODE, string GTYPE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT BILLNO  訂單單號 from GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  and GTYPE=@GTYPE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@GTYPE", GTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT22(string SHIPPINGCODE, string GTYPE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT BILLNO  訂單單號2 from GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  and GTYPE=@GTYPE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@GTYPE", GTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT3(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ITEMCODE 產品編號,ITEMNAME 產品名稱,CAST(QTY AS INT) 數量 from GB_PICK2 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable DT4(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ITEMCODE 產品編號,ITEMNAME 產品名稱,PACK1 箱1,PACK3 公斤,LINE 欄號 from GB_PICK2 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable DT5(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT sum(PACK1) 箱1 from GB_PICK2 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable DT6(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (PACK1) 箱1 from GB_PICK2 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void GBPICK2_Load(object sender, EventArgs e)
        {
            DELETEFILE();
            dataGridView1.DataSource = DT();

            dataGridView1.Rows[0].Selected = true;
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir ;
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    FileInfo filess = new FileInfo(file);
                    string dd = filess.Name.ToString();
                    int ad = dd.LastIndexOf(".");
                    string PanelName = dd.Substring(ad, dd.Length - ad).ToString();
                    if (PanelName == ".csv")
                    {
                        File.Delete(file);
                    }

                }
            }
            catch { }
        }


        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("BILLNO", typeof(string));
            dt.Columns.Add("ITEMCODE", typeof(string));
            dt.Columns.Add("ITEMNAME", typeof(string));
            dt.Columns.Add("ITEMNAME1", typeof(string));
            dt.Columns.Add("QTY2", typeof(string));
            dt.Columns.Add("CustAddress", typeof(string));
            dt.Columns.Add("LinkMan", typeof(string));
            dt.Columns.Add("LinkTelephone", typeof(string));
            dt.Columns.Add("UserDef1", typeof(string));
            dt.Columns.Add("PreInDate", typeof(string));
            dt.Columns.Add("UserDef2", typeof(string));
            dt.Columns.Add("AMT", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }

        private System.Data.DataTable DTDEADLINE()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ITEMCODE", typeof(string));
            dt.Columns.Add("ITEMNAME", typeof(string));
            dt.Columns.Add("BARCODEID", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("DOCDATE", typeof(string));
            dt.Columns.Add("DEADDATE", typeof(string));
            return dt;
        }
        private void TOTAL2(string ID)
        {
            dtCost = MakeTableCombine();
            DataRow dr = null;
        
            System.Data.DataTable DT1 = GetEzCat(ID);
          
          
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {

                int PACK1 = Convert.ToInt16(DT1.Rows[i]["PACK1"].ToString());
                for (int s = 1; s <= PACK1; s++)
                {
                    string a = PACK1.ToString();
                    string a1 = s.ToString();

                    string BILLNO = DT1.Rows[i]["BILLNO"].ToString().Trim();
                    string CARDCODE = DT1.Rows[i]["CARDCODE"].ToString().Trim();
                        dr = dtCost.NewRow();
                        System.Data.DataTable dt = GetGROUP(ID, BILLNO);
                        System.Data.DataTable dt2 = GetGROUP2(ID, BILLNO);
                        System.Data.DataTable dt3 = GetGROUP3(ID, BILLNO);
                        string TRADE = "";
                        string CLASS = "";
    
                        StringBuilder sb = new StringBuilder();

                        if (dt.Rows.Count > 0)
                        {

                             TRADE = dt3.Rows[0]["TRADE"].ToString();
                            for (int S = 0; S <= dt.Rows.Count - 1; S++)
                            {

                                DataRow dd = dt.Rows[S];
                                sb.Append(dd["GBGROUP"].ToString() + "/");
                            }

                            sb.Remove(sb.Length - 1, 1);

                            CLASS = sb.ToString();
                        }

                        string EZ = "";
                        System.Data.DataTable GG = CHOEZCAT(CARDCODE);
                        if (GG.Rows.Count > 0)
                        {
                            EZ = "_" + CHOEZCAT(CARDCODE).Rows[0][0].ToString();
                        }


                        dr["BILLNO"] = BILLNO +  EZ;
                            dr["ITEMCODE"] = CLASS;
                            dr["ITEMNAME"] = CLASS;
                            dr["ITEMNAME1"] = CLASS;
                            dr["QTY2"] = DT1.Rows[i]["QTY2"].ToString().Trim();
                            dr["AMT"] = DT1.Rows[i]["AMT"].ToString().Trim();
                  

                        dr["CustAddress"] = DT1.Rows[i]["CustAddress"].ToString().Trim();
                        dr["LinkMan"] = DT1.Rows[i]["LinkMan"].ToString().Trim();
                        dr["LinkTelephone"] = "'" + DT1.Rows[i]["LinkTelephone"].ToString().Trim();
                        dr["UserDef1"] = DT1.Rows[i]["UserDef1"].ToString().Trim();
                        dr["PreInDate"] = DT1.Rows[i]["PreInDate"].ToString().Trim();
                        dr["UserDef2"] = DT1.Rows[i]["UserDef2"].ToString().Trim();

                        dr["備註"] = DT1.Rows[i]["CUSTMEMO"].ToString().Trim() + "**到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！";

                        dtCost.Rows.Add(dr);

                //    }
                }
                
            }

            //System.Data.DataTable DT2 = GetEzCat2(ID);
            //for (int i = 0; i <= DT2.Rows.Count - 1; i++)
            //{
            //    int ROW = Convert.ToInt16(DT2.Rows[i]["ROW"].ToString());
            
                    
            //        dr = dtCost.NewRow();

            //        if (DT2.Rows.Count > 1)
            //        {
            //            dr["BILLNO"] = DT2.Rows[i]["BILLNO"].ToString().Trim() + "-" + ROW.ToString();
                       
                  
            //        }
            //        else
            //        {
            //            dr["BILLNO"] = DT2.Rows[i]["BILLNO"].ToString().Trim();

            //        }
            //        dr["ITEMCODE"] = DT2.Rows[i]["ITEMNAME"].ToString().Trim();
            //        dr["ITEMNAME"] = DT2.Rows[i]["ITEMNAME"].ToString().Trim();
            //        dr["ITEMNAME1"] = DT2.Rows[i]["ITEMNAME"].ToString().Trim();
            //        dr["QTY2"] = DT2.Rows[i]["QTY2"].ToString().Trim();
            //        dr["CustAddress"] = DT2.Rows[i]["CustAddress"].ToString().Trim();
            //        dr["LinkMan"] = DT2.Rows[i]["LinkMan"].ToString().Trim();
            //        dr["LinkTelephone"] = "'" + DT2.Rows[i]["LinkTelephone"].ToString().Trim();

            //        dr["UserDef1"] = DT2.Rows[i]["UserDef1"].ToString().Trim();
            //        dr["PreInDate"] = DT2.Rows[i]["PreInDate"].ToString().Trim();
            //        dr["UserDef2"] = DT2.Rows[i]["UserDef2"].ToString().Trim();
            //        dr["AMT"] = DT2.Rows[i]["AMT"].ToString().Trim();
            //        dr["備註"] = "到貨前請先聯絡收件人，並提醒收件人馬上冷凍，謝謝！！";
           
            //        dtCost.Rows.Add(dr);


                
            //}

        }


        public System.Data.DataTable GetEzCat(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("   SELECT BILLNO,MAX(CARDCODE) CARDCODE, ");
            sb.Append("                           SUM(QTY2) QTY2,");
            sb.Append("         MAX(CustAddress) CustAddress,");
            sb.Append("                         MAX(LinkMan) LinkMan,");
            sb.Append("                            MAX(replace(linktelephone,'''',''))  LinkTelephone,");
            sb.Append(" MAX(UserDef1)  UserDef1,");
            sb.Append(" MAX(PreInDate)  PreInDate,");
            sb.Append(" CASE MAX(UserDef2) WHEN '13:00以前' THEN 1 WHEN '13:00 以前' THEN 1 WHEN '中午前' THEN 1  WHEN '12-17時' THEN 2  WHEN '14:00 - 18:00' THEN 2 WHEN '17-20時' THEN 3 WHEN '不指定' THEN 4  WHEN '任何時段' THEN 4  END   UserDef2,");
            sb.Append("               MAX(AMT)            AMT,MAX(PACK1) PACK1,MAX(CUSTMEMO) CUSTMEMO");
            sb.Append("               FROM GB_PICK2  where shippingcode=@shippingcode  AND GTYPE IN ('零售','禮盒') ");
            sb.Append(" GROUP BY BILLNO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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

        public System.Data.DataTable MANHWAR2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT 'S' 模式,SHIPPINGCODE 訂單編號,RTRIM(UserDef1) 出貨日期,RTRIM(PreInDate) 到貨日期,LINE+1 訂單明細ID,");
            sb.Append(" ITEMCODE 商品編號,ITEMNAME 商品名稱,ITEMNAME 規格,QTY 數量,");
            sb.Append(" AMT2 訂單未稅金額,0 訂單稅額,AMT2  未稅金額,0 稅額,AMT2 含稅金額");
            sb.Append(" ,LINKMAN 收件人名稱,'' 郵遞區號,CUSTADDRESS 收件人地址,LINKTELEPHONE 行動電話");
            sb.Append(" ,'' 日間聯絡電話,'' 夜間聯絡電話,CASE UserDef2  WHEN '不指定' THEN 1  WHEN '中午前' THEN 2 WHEN '12-17時' THEN 3");
            sb.Append(" WHEN '17-20時' THEN 4 END 到貨時段,AMT3  代收金額,'' 訂單備註,'' 品項備註,Convert(varchar(10),Getdate(),111)+Convert(varchar(12),Getdate(),114) 拋單日期 ");
            sb.Append("  FROM GB_PICK2 WHERE SHIPPINGCODE=@shippingcode");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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

        public System.Data.DataTable MANFONTAI(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

   

            sb.Append("                            SELECT 'RGN' 貨主編號,BILLNO 貨主單號,MAX(ROWNO) 序號,MAX(CARDCODE) 客戶端代號   ");
            sb.Append("                                           ,MAX(RTRIM(UserDef1)) 訂購日期,MAX(RTRIM(PreInDate)) 預計到貨日,ITEMCODE 商品編號,MAX(ITEMNAME) 商品名稱,   ");
            sb.Append("                                           'X' 倉別,DEADDATE 指定效期,'' 指定, SUM(QTY) 訂購數量,'X' 商品單價,MAX(CASE WHEN CARDCODE='TW90144-94' THEN 'FT' ELSE 'Tcat' END) 配送方式   ");
            sb.Append("                                           ,MAX(LINKMAN) 收貨人姓名,MAX(CUSTADDRESS) 收貨人地址,MAX(LINKTELEPHONE) 收貨人聯絡電話,   ");
            sb.Append("                                           '' 日間聯絡電話,'' 夜間聯絡電話,MAX(CASE UserDef2  WHEN '不指定' THEN 1  WHEN '中午前' THEN 2 WHEN '12-17時' THEN 3    ");
            sb.Append("                                                         WHEN '17-20時' THEN 4 END) 到貨時段,SUM(AMT)  代收金額,'' 宅配單備註,'' 品項備註,CONVERT(VARCHAR(8) , GETDATE(), 112 ) 單日期   ");
            sb.Append("                                            FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE  AND ITEMNAME <> 'DM'    ");
            sb.Append("                             GROUP BY BILLNO,ITEMCODE,DEADDATE ORDER BY BILLNO  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        public System.Data.DataTable GetGROUP(string shippingcode,string BILLNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


 
            sb.Append(" select DISTINCT GBGROUP from GB_PICK2  where GBGROUP='朝貢豬' AND SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" UNION ALL");
            sb.Append(" select DISTINCT GBGROUP from GB_PICK2  where GBGROUP='朝貢雞' AND SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" UNION ALL");
            sb.Append(" select DISTINCT GBGROUP from GB_PICK2  where GBGROUP='大力蝦' AND SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" UNION ALL");
            sb.Append(" select DISTINCT GBGROUP from GB_PICK2  where GBGROUP='白金蝦' AND SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" UNION ALL");
            sb.Append(" select DISTINCT GBGROUP from GB_PICK2  where GBGROUP NOT IN ('白金蝦','大力蝦','朝貢雞','朝貢豬') AND SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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

        public System.Data.DataTable GetGROUP2(string shippingcode, string BILLNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("   select  QTY from GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO AND ROWNO=0  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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
        public System.Data.DataTable GetGROUP3(string shippingcode, string BILLNO)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("   select  TRADE from GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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
        public System.Data.DataTable GetEzCat2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT BILLNO, QTY2  QTY2,CustAddress,ITEMNAME+'  '+PACK3 +'公斤' ITEMNAME,");
            sb.Append("                    LinkMan, '0'+replace(linktelephone,'''','') LinkTelephone,UserDef1,PreInDate, Rank() OVER (ORDER BY LINE) ROW,");
            sb.Append("   CASE UserDef2 WHEN '中午前' THEN 1 WHEN '12-17時' THEN 2 WHEN '17-20時' THEN 3 WHEN '不指定' THEN 4  END UserDef2, AMT,PACK1,PACK3");
            sb.Append("                FROM GB_PICK2  where shippingcode=@shippingcode AND GTYPE='批發'");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
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
        private void FF(string 撿貨單號)
        {
            TOTAL2(撿貨單號);
            INSEREZ(撿貨單號, "匯出EZCAT");
            dataGridView6.DataSource = dtCost;
            string t1 = "ezcat" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            ExcelReport.GridViewToCSVCATPOTATO(dataGridView6, t1);
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ExcelReport.DELETEFILE();
                          scrollPosition = e.RowIndex;

                      if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                      {
                          DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                          if (column.Name == "colEdit2")
                          {
                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {
                                    string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();
                                  System.Data.DataTable H1 = GPSZ(撿貨單號, "匯出EZCAT");
                                  if (H1.Rows.Count > 0)
                                  {
                                      DialogResult result;
                                      result = MessageBox.Show("匯出EZCAT已匯過，請確認是否要匯出", "請確認", MessageBoxButtons.YesNo);
                                      if (result == DialogResult.Yes)
                                      {
                                          FF(撿貨單號);

                                      }

                                  }
                                  else
                                  {
                                      FF(撿貨單號);
                                  }
                 
                              }
                          }
 
                          if (column.Name == "逢泰EzCat")
                          {
                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {
                                  string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();
                                  dataGridView6.DataSource = MANHWAR2(撿貨單號);
                                  string t1 = "RGN_SHIPMENT_" + DateTime.Now.ToString("yyyyMMddHHmmssFFF") + ".sht";
                                  ExcelReport.GridViewToCSVCATPOTATO2(dataGridView6, t1);
                              }
                          }

                          if (column.Name == "逢泰出貨主檔")
                          {
                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {
                                  string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();
                                  System.Data.DataTable G1 = MANFONTAI(撿貨單號);
                                  if (G1.Rows.Count > 0)
                                  {
                                      System.Data.DataTable dtCost = MakeTableFONTAI();
                                      DataRow dr = null;
                                      for (int i = 0; i <= G1.Rows.Count - 1; i++)
                                      {
                                          DataRow dd = G1.Rows[i];
                                          dr = dtCost.NewRow();
                                          dr["貨主編號"] = dd["貨主編號"].ToString();
                                          dr["貨主單號"] = dd["貨主單號"].ToString();
                                          dr["序號"] = dd["序號"].ToString();
                                          dr["客戶端代號"] = dd["客戶端代號"].ToString();
                                          dr["訂購日期"] = dd["訂購日期"].ToString();
                                          dr["預計到貨日"] = dd["預計到貨日"].ToString();
                                          dr["商品編號"] = dd["商品編號"].ToString();
                               string ITEMCODE= dd["商品名稱"].ToString();
                                dr["商品名稱"] = ITEMCODE;
                                          dr["倉別"] = dd["倉別"].ToString();
                                          dr["指定"] = dd["指定"].ToString();
                                          dr["訂購數量"] = dd["訂購數量"].ToString();
                                          dr["商品單價"] = dd["商品單價"].ToString();
                                          dr["配送方式"] = dd["配送方式"].ToString();
                                    
                                          dr["收貨人姓名"] = dd["收貨人姓名"].ToString();
                                          dr["收貨人地址"] = dd["收貨人地址"].ToString();
                                          dr["收貨人聯絡電話"] = dd["收貨人聯絡電話"].ToString();
                                          dr["日間聯絡電話"] = dd["日間聯絡電話"].ToString();
                                          dr["夜間聯絡電話"] = dd["夜間聯絡電話"].ToString();
                                          dr["到貨時段"] = dd["到貨時段"].ToString();
                                          dr["代收金額"] = dd["代收金額"].ToString();
                                          dr["宅配單備註"] = dd["宅配單備註"].ToString();
                                          dr["品項備註"] = dd["品項備註"].ToString();
                                          dr["單日期"] = dd["單日期"].ToString();
                                          dr["指定效期"] = dd["指定效期"].ToString();
                       
                                          dtCost.Rows.Add(dr);


                                      }

                                      string FileName = string.Empty;
                                      string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                                      FileName = lsAppDir + "\\Excel\\GW\\逢泰出貨.xls";

                                      //Excel的樣版檔
                                      string ExcelTemplate = FileName;
                                 
                                      string OutPutFile = lsAppDir + "\\Excel\\temp\\" + "RGN_SALE_" +
                                            DateTime.Now.ToString("yyyyMMddHHmmss") + "_shipments.xls";

                                      //產生 Excel Report
                                      ExcelReport.ExcelReportOutputFONTAI(dtCost, ExcelTemplate, OutPutFile);
                                  }
                              }
                          }
                          if (column.Name == "結案")
                          {
                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {

                                  string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();
                             

                                  DialogResult result;
                                  result = MessageBox.Show("單號 " + 撿貨單號 + " 請確認是否結案", "YES/NO", MessageBoxButtons.YesNo);
                                  if (result == DialogResult.Yes)
                                  {
                                      GetMenu.UPDATEPICKCHECK(撿貨單號);

                                      dataGridView1.DataSource = DT();
                                  }
                              }
                          }
                          if (column.Name == "備貨單")
                          {
                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {
                                  string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();
                                  string 批發 = Convert.ToString(row["批發"]).Trim();


                                  string FileName = string.Empty;
                                  string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                                  System.Data.DataTable T2 = DDS(撿貨單號);
                                  if (T2.Rows.Count > 0)
                                  {
                                      System.Data.DataTable dtCost = MakeTable();
                                      DataRow dr = null;
                                      for (int i = 0; i <= T2.Rows.Count - 1; i++)
                                      {
                                          DataRow dd = T2.Rows[i];
                                          dr = dtCost.NewRow();
                                          dr["序號"] = dd["序號"].ToString();

                                          string ITEMCODE = dd["ITEMCODE"].ToString();

                                string ITEMNAME = dd["ITEMNAME"].ToString().Replace("加工品_", "");
                                dr["ITEMCODE"] = ITEMCODE;
                                          dr["QTY"] = dd["QTY"].ToString();
                                dr["ITEMNAME"] = ITEMNAME;
                                          dr["DOCDATE"] = dd["DOCDATE"].ToString();
                                          System.Data.DataTable TT1 = DDSUMQTY(撿貨單號);
                                          if (TT1.Rows.Count > 0)
                                          {
                                              dr["總數量"] = TT1.Rows[0][0].ToString();
                                          }

                                          dr["DEADDATE"] = dd["指定效期"].ToString();
                               
                                          dtCost.Rows.Add(dr);
                                      }
                                      string OutPutFile = "";
                            FileName = lsAppDir + "\\Excel\\GW\\小包裝備貨.xlsx";

                            OutPutFile = lsAppDir + "\\Excel\\temp\\" +
            DateTime.Now.ToString("yyyyMMddHHmmss") + "retail.xlsx";

                                      string ExcelTemplate = FileName;
                                      WriteExcelProduct(FileName, OutPutFile, 撿貨單號, 批發, dtCost);


                                  }
                              }
                          }
      
 
                          if (column.Name == "簡訊")
                          {

                              DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;

                              if (row != null)
                              {
                                  string 撿貨單號 = Convert.ToString(row["撿貨單號"]).Trim();

                                  System.Data.DataTable T1 = DDMESSAGE(撿貨單號);
                                  string FileName = string.Empty;
                                  string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                                  FileName = lsAppDir + "\\Excel\\GW\\簡訊.xls";

                                  string ExcelTemplate = FileName;

                                  string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                        DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                                  //產生 Excel Report
                                  ExcelReport.ExcelReportOutput(T1, ExcelTemplate, OutPutFile, "N");

                              }


                          }
                      }
        }

        public void GETMANDAY(string 撿貨日期)
        {
            撿貨日期 = 撿貨日期.Substring(0, 4) + "/" + 撿貨日期.Substring(4, 2) + "/" + 撿貨日期.Substring(6, 2);
            DateTime D1 = Convert.ToDateTime(撿貨日期);
            DateTime D2;
            DateTime D3 ;
            PICKDAY1 = Convert.ToString(Convert.ToInt16(D1.ToString("MM"))) + "月" + Convert.ToString(Convert.ToInt16(D1.ToString("dd"))) + "日";
            int DAY = (int)D1.DayOfWeek;

            if (DAY == 1 || DAY == 3)
            {
                 D2 = D1.AddDays(2);
                 D3 = D1.AddDays(3);
            }
            else if (DAY == 2 || DAY == 4)
            {
                 D2 = D1.AddDays(1);
                 D3 = D1.AddDays(2);
            }
            else
            {
                 D2 = D1.AddDays(3);
                 D3 = D1.AddDays(4);
            }

            PICKDAY2 = Convert.ToString(Convert.ToInt16(D2.ToString("MM"))) + "月" + Convert.ToString(Convert.ToInt16(D2.ToString("dd"))) + "日";
            PICKDAY3 = Convert.ToString(Convert.ToInt16(D3.ToString("MM"))) + "月" + Convert.ToString(Convert.ToInt16(D3.ToString("dd"))) + "日";
        }
        private System.Data.DataTable DTMAIN( string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT CASE WHEN ISNULL(PACKNO,'') <> '' THEN PACKNO ELSE  BILLNO END BILLNO  FROM GB_PICK2  WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private System.Data.DataTable DTANN(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT LINKMAN FROM GB_PICK2  WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT BILLNO,LINKMAN FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable DTBILLNOMAN()
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(ADDRID AS INT) ADDRID,CASE AddrID WHEN 038  THEN '新店中華' WHEN 039 THEN '新店門市'  ELSE REPLACE(REPLACE(REPLACE(LinkMan,'店',''),'棉花田-',''),'門市','') END LINK,WALKADDR  FROM comCustAddress ");
            sb.Append(" WHERE ID='TW90144-94' AND WalkAddr IN(135,246) ORDER BY WalkAddr,CAST(ADDRID AS INT)");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
          
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable CHOEZCAT(string ID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.EngName  FROM comCustomer T0 LEFT JOIN comCustClass T1 ON (T0.ClassID =T1.ClassID AND T0.Flag =T1.Flag)  WHERE ID=@ID  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNOANN()
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                     SELECT CAST(ADDRID AS INT) ADDRID,REPLACE(REPLACE(REPLACE(LinkMan,'店',''),'安永鮮物-',''),'門市','') LINK,WALKADDR  FROM comCustAddress   ");
            sb.Append("                             WHERE ID='TW90146-89' AND AddrID <> 0 AND WalkAddr IN(135,246) ORDER BY WalkAddr,CAST(ADDRID AS INT)  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO2(string SHIPPINGCODE, string BILLNO, string ITEMCODE, string ITEMNAME, string DEADDATE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            //DEADDATE
           // sb.Append(" SELECT SUM(QTY2) QTY  FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO AND ITEMCODE=@ITEMCODE ");
            sb.Append("                SELECT  SUM(QTY) QTY  FROM (   ");
            sb.Append("                  SELECT Convert(varchar(10),MAX(CAST(DOCDATE AS DATETIME)),111) DOCDATE,ITEMCODE, ");
            sb.Append("               ITEMNAME+ ");
            sb.Append("               CASE WHEN PRICE = 0 AND ITEMNAME <> 'DM'  THEN '<贈品>'  ");
            sb.Append("               WHEN T1.MEMO='短效品' THEN '<短效品>' ");
            sb.Append("               WHEN T1.MEMO='促銷品' THEN '<促銷品>' ");
            sb.Append("               ELSE '' END ITEMNAME,SUM(QTY) QTY,CASE WHEN PRICE=0 THEN '贈品: ' ");
            sb.Append("               WHEN T1.MEMO='短效品' THEN '短效品:' ");
            sb.Append("               WHEN T1.MEMO='促銷品' THEN '促銷品:' ");
            sb.Append("                ELSE '' END MEMO   FROM GB_PICK T0   ");
            sb.Append("                LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)    ");
            sb.Append("                WHERE T0.SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE  AND BILLNO=@BILLNO AND ISNULL(DEADDATE,'') =@DEADDATE GROUP BY ITEMCODE,ITEMNAME, ");
            sb.Append("               CASE WHEN PRICE=  0 AND ITEMNAME <> 'DM'  THEN '<贈品>'  ");
            sb.Append("               WHEN T1.MEMO='短效品' THEN '<短效品>' ");
            sb.Append("               WHEN T1.MEMO='促銷品' THEN '<促銷品>' ");
            sb.Append("               ELSE '' END,CASE WHEN PRICE=0 THEN '贈品: ' ");
            sb.Append("               WHEN T1.MEMO='短效品' THEN '短效品:' ");
            sb.Append("               WHEN T1.MEMO='促銷品' THEN '促銷品:' ");
            sb.Append("                ELSE '' END   ) AS A  WHERE  REPLACE(A.ITEMNAME,'加工品_','')=@ITEMNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@DEADDATE", DEADDATE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO2MAN(string SHIPPINGCODE, string PACKAGE, string ITEMCODE,string T1)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            if (T1 == "1")
            {
                sb.Append(" SELECT SUM(QTY2) QTY  FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND PACKAGE=@PACKAGE AND ITEMCODE=@ITEMCODE AND PRICE <> 0 ");

            }
            if (T1 == "2")
            {
                sb.Append(" SELECT SUM(QTY2) QTY  FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND PACKAGE=@PACKAGE AND ITEMCODE=@ITEMCODE AND PRICE = 0 ");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PACKAGE", PACKAGE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO3MAN(string SHIPPINGCODE, string PACKAGE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(QTY2) QTY  FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND PACKAGE=@PACKAGE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PACKAGE", PACKAGE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable DTBILLNO4MAN(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT COUNT(DISTINCT LINKMAN)  LINKMAN FROM GB_PICK2 WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO3(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("        SELECT  ITEMCODE, SUM(QTY2) QTY FROM GB_PICK2 T0 ");
            sb.Append("               WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" AND T0.ITEMNAME <> 'DM'     GROUP BY ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO33(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("        SELECT SUM(QTY2) QTY FROM GB_PICK2 T0 ");
            sb.Append("               WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO");
            sb.Append(" AND T0.ITEMNAME <> 'DM'  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNOP1(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT SUM(QTY2) QTY FROM GB_PICK2 T0  ");
            sb.Append("               WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO AND ISNULL(CHI,'') NOT IN ('011','012','013','018')");
            sb.Append("               AND T0.ITEMNAME <> 'DM'   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNOP2(string SHIPPINGCODE, string BILLNO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT SUM(QTY2) QTY FROM GB_PICK2 T0  ");
            sb.Append("               WHERE SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO AND ISNULL(CHI,'') IN ('011','012','013','018')");
            sb.Append("               AND T0.ITEMNAME <> 'DM'   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO4(string ProdID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PRODID,case when ClassName like '%零售%' THEN '零售' END GTYPE ,PackAmt1 包");
            sb.Append(" ,REPLACE(SUBSTRING(PRODDESC,CHARINDEX('##', ProdDesc)+2,10),'##','') PACK    FROM comProduct T0");
            sb.Append("  LEFT JOIN comProductClass T1 ON (T0.ClassID =T1.ClassID)   WHERE ProdDesc LIKE '%##%'");
            sb.Append(" and ProdID=@ProdID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTBILLNO5(string ProdID)
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PRODID,case when ClassName like '%零售%' THEN '零售' when ClassName like '%批發%' THEN '批發' END GTYPE ,PackAmt1 包,CtmWeight/1000 WEIGHT,PackAmt1 包   FROM comProduct T0");
            sb.Append("  LEFT JOIN comProductClass T1 ON (T0.ClassID =T1.ClassID)   WHERE ProdID=@ProdID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ProdID", ProdID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    
        private System.Data.DataTable DT(string BILLNO, string SHIPPINGCODE, string AMT)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();



            sb.Append("     SELECT '到貨日期 : '+SUBSTRING(PreInDate,1,4)+'/'+SUBSTRING(PreInDate,5,2)+'/'+SUBSTRING(PreInDate,7,2)  ");
            sb.Append("            到貨日期,'訂單號碼: '+CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END 訂單號碼,'運輸單號: '+ISNULL(T1.STORE,'') 運輸單號,  ");
            sb.Append("              '訂單日期: '+SUBSTRING(BILLDATE,1,4)+'/'+SUBSTRING(BILLDATE,5,2)+'/'+SUBSTRING(BILLDATE,7,2) 訂單日期  ");
            sb.Append("              , '收件姓名: '+LinkMan 收貨人,'收貨地址: '+");
            sb.Append("			  REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CustAddress,'1','*'),'2','*'),'3','*'),'4','*') ,'5','*') ,'6','*') ,'7','*') ,'8','*') ,'9','*') ,'0','*')   收貨地址");
            sb.Append("			  ,'聯繫電話: '+REPLACE(LinkTelephone,SUBSTRING(LinkTelephone,LEN(LinkTelephone)/2+1,3),'***') 聯繫電話 ");
            sb.Append("              ,RANK() OVER (ORDER BY REPLACE(T1.ITEMNAME,'_','')  DESC) AS [NO],T1.ITEMCODE 產品編號,T1.ITEMNAME 品名規格  ");
            sb.Append("              ,CASE WHEN T1.UNIT='KG' THEN CAST(CAST(T1.QTY AS DECIMAL(10,1)) AS VARCHAR) ELSE CAST(T1.QTY2 AS VARCHAR) END  箱,'KG' KG,'訂購人: '+LTRIM(SUBSTRING(ORDMAN,0,CHARINDEX('TEL', ORDMAN))) 訂購人,LTRIM(SUBSTRING(ORDMAN,CHARINDEX('TEL', ORDMAN),30))  TEL,T1.MEMO 備註,T1.UNIT FROM GB_PICK T0  ");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE and CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END=@BILLNO   ORDER BY NO  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable DTF(string BILLNO, string SHIPPINGCODE, string AMT)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();



            sb.Append("        SELECT '到貨日期 : '+SUBSTRING(PreInDate,1,4)+'/'+SUBSTRING(PreInDate,5,2)+'/'+SUBSTRING(PreInDate,7,2)   ");
            sb.Append("                         到貨日期,'訂單號碼: '+CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END 訂單號碼,'運輸單號: '+ISNULL(T1.STORE,'') 運輸單號,   ");
            sb.Append("                           '訂單日期: '+SUBSTRING(BILLDATE,1,4)+'/'+SUBSTRING(BILLDATE,5,2)+'/'+SUBSTRING(BILLDATE,7,2) 訂單日期   ");
            sb.Append("                           , '收件姓名: '+LinkMan 收貨人,'收貨地址: '+ ");
            sb.Append("             			  REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CustAddress,'1','*'),'2','*'),'3','*'),'4','*') ,'5','*') ,'6','*') ,'7','*') ,'8','*') ,'9','*') ,'0','*')   收貨地址 ");
            sb.Append("             			  ,'聯繫電話: '+REPLACE(LinkTelephone,SUBSTRING(LinkTelephone,LEN(LinkTelephone)/2+1,3),'***') 聯繫電話  ");
            sb.Append("                           ,CAST(ROWNO AS VARCHAR) AS [NO],T1.ITEMCODE 產品編號,T1.ITEMNAME 品名規格   ");
            sb.Append("                           ,CASE WHEN T1.UNIT='KG' THEN CAST(CAST(T1.QTY AS DECIMAL(10,1)) AS VARCHAR) ELSE CAST(T1.QTY2 AS VARCHAR) END  箱,'KG' KG,'訂購人: '+LTRIM(SUBSTRING(ORDMAN,0,CHARINDEX('TEL', ORDMAN))) 訂購人,LTRIM(SUBSTRING(ORDMAN,CHARINDEX('TEL', ORDMAN),30))  TEL,T1.MEMO 備註,LTRIM(RTRIM(T1.UNIT)) UNIT,T1.ROWNO");
            sb.Append("						 FROM GB_PICK T0   ");
            sb.Append("              LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)   ");
            sb.Append("              WHERE T0.SHIPPINGCODE=@SHIPPINGCODE and CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END=@BILLNO");
            sb.Append("			  UNION ALL");
            sb.Append("			          SELECT '到貨日期 : '+SUBSTRING(PreInDate,1,4)+'/'+SUBSTRING(PreInDate,5,2)+'/'+SUBSTRING(PreInDate,7,2)   ");
            sb.Append("                         到貨日期,'訂單號碼: '+CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END 訂單號碼,'運輸單號: '+ISNULL(T1.STORE,'') 運輸單號,   ");
            sb.Append("                           '訂單日期: '+SUBSTRING(BILLDATE,1,4)+'/'+SUBSTRING(BILLDATE,5,2)+'/'+SUBSTRING(BILLDATE,7,2) 訂單日期   ");
            sb.Append("                           , '收件姓名: '+LinkMan 收貨人,'收貨地址: '+ ");
            sb.Append("             			  REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CustAddress,'1','*'),'2','*'),'3','*'),'4','*') ,'5','*') ,'6','*') ,'7','*') ,'8','*') ,'9','*') ,'0','*')   收貨地址 ");
            sb.Append("             			  ,'聯繫電話: '+REPLACE(LinkTelephone,SUBSTRING(LinkTelephone,LEN(LinkTelephone)/2+1,3),'***') 聯繫電話  ");
            sb.Append("                           ,'' [NO],T2.COMBSUBID 產品編號,'      '+T3.PRODNAME 小料號");
            sb.Append("                           , T2.AMOUNT 箱, 'KG' KG,'訂購人: '+LTRIM(SUBSTRING(ORDMAN,0,CHARINDEX('TEL', ORDMAN))) 訂購人,LTRIM(SUBSTRING(ORDMAN,CHARINDEX('TEL', ORDMAN),30))  TEL,T1.MEMO 備註,'包',T1.ROWNO FROM GB_PICK T0   ");
            sb.Append("              LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)   ");
            sb.Append("			  LEFT JOIN  otherDB.CHIComp02.DBO.comProdCombine T2 ON(T1.ITEMCODE=T2.PRODID)");
            sb.Append("			  			  LEFT JOIN  otherDB.CHIComp02.DBO.COMPRODUCT T3 ON(T2.COMBSUBID=T3.PRODID)");
            sb.Append("              WHERE T0.SHIPPINGCODE=@SHIPPINGCODE and CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END=@BILLNO");
            sb.Append("			  AND T3.PRODNAME <> ''");
            sb.Append("			     ORDER BY ROWNO,產品編號");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTAMT(string BILLNO, string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(AMT) AMT,MAX(CustAddress) CUST   FROM GB_PICK T0  ");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)   ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE  and CASE WHEN ISNULL(T1.PACKNO,'') <> '' THEN T1.PACKNO ELSE  T1.BILLNO END=@BILLNO  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DTANN(string LINKMAN, string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT SUBSTRING(PreInDate,1,4)+'/'+SUBSTRING(PreInDate,5,2)+'/'+SUBSTRING(PreInDate,7,2) 到貨日期");
            sb.Append("              ,'訂單號碼: '+T1.BILLNO 訂單號碼,'運輸單號: '+ISNULL(T1.STORE,'') 運輸單號,");
            sb.Append("            '訂單日期: '+SUBSTRING(BILLDATE,1,4)+'/'+SUBSTRING(BILLDATE,5,2)+'/'+SUBSTRING(BILLDATE,7,2) 訂單日期");
            sb.Append("            , '收貨人: '+LinkMan 收貨人,'收貨地址: '+CustAddress 收貨地址,'聯繫電話: '+LinkTelephone 聯繫電話");
            sb.Append("            ,RANK() OVER (ORDER BY T1.ITEMCODE) AS [NO],T1.ITEMCODE 產品編號,T1.ITEMNAME 品名規格");
            sb.Append(" ,'付款方式: '+CASE TRADE WHEN '貨到付款' THEN  '貨到付款' +CAST(AMT AS VARCHAR) ELSE TRADE END 付款方式");
            sb.Append(" ,T1.QTY2 箱,'KG' KG,'訂購人: '+LTRIM(SUBSTRING(ORDMAN,0,CHARINDEX('TEL', ORDMAN))) 訂購人,LTRIM(SUBSTRING(ORDMAN,CHARINDEX('TEL', ORDMAN),30))  TEL,T1.MEMO 備註,''''+BARCODEID 國際條碼 FROM GB_PICK T0");
            sb.Append("            LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("                 WHERE T0.SHIPPINGCODE=@SHIPPINGCODE and T1.LINKMAN=@LINKMAN ORDER BY NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@LINKMAN", LINKMAN));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln1");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public System.Data.DataTable DDS(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCDATE,ITEMCODE,ITEMNAME,QTY,RANK() OVER (ORDER BY ITEMCODE,指定效期 DESC) AS 序號,指定效期 FROM (   ");
            sb.Append(" SELECT Convert(varchar(10),MAX(CAST(DOCDATE AS DATETIME)),111) DOCDATE,ITEMCODE, ");
            sb.Append(" ITEMNAME+ ");
            sb.Append(" CASE WHEN PRICE = 0 AND ITEMNAME <> 'DM'  THEN '<贈品>'  ");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '<短效品>' ");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '<促銷品>' ");
            sb.Append(" ELSE '' END ITEMNAME,SUM(QTY) QTY,T1.DEADDATE 指定效期    FROM GB_PICK T0   ");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)    ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE GROUP BY ITEMCODE,ITEMNAME,T1.DEADDATE, ");
            sb.Append(" CASE WHEN PRICE=  0 AND ITEMNAME <> 'DM'  THEN '<贈品>'  ");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '<短效品>' ");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '<促銷品>' ");
            sb.Append(" ELSE '' END  ) AS A ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable DDSMAN(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,");
            sb.Append(" MAX(GBGROUP)+'-'+ITEMNAME+");
            sb.Append(" CASE WHEN PRICE=  0 THEN '<贈品>' ");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '<短效品>'");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '<促銷品>'");
            sb.Append(" ELSE '' END ITEMNAME");
            sb.Append(" ,BARCODEID,SUM(QTY) QTY,T0.DOCDATE,");
            sb.Append(" CASE WHEN PRICE=0 THEN '贈品: '");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '短效品:'");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '促銷品:'");
            sb.Append("  ELSE '' END MEMO FROM GB_PICK T0    ");
            sb.Append("  LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)   ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append(" GROUP BY ITEMCODE,ITEMNAME,BARCODEID,");
            sb.Append(" CASE WHEN PRICE=0 THEN '贈品: '");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '短效品:'");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '促銷品:'  ELSE '' END,CASE WHEN PRICE=  0 THEN '<贈品>' ");
            sb.Append(" WHEN T1.MEMO='短效品' THEN '<短效品>'");
            sb.Append(" WHEN T1.MEMO='促銷品' THEN '<促銷品>' ELSE '' END,T0.DOCDATE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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

        public System.Data.DataTable DDSMANF(string ITEMCODE,string TYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT Convert(varchar(8),DEADLINE,112)   DEADLINE FROM GB_DEADLINE WHERE ITEMCODE=@ITEMCODE  ");
           
            if (TYPE == "D1")
            {
                sb.Append("    AND D1='TRUE'  ");
            }

            if (TYPE == "D2")
            {
                sb.Append("    AND D2='TRUE'  ");
            }

            if (TYPE == "D3")
            {
                sb.Append("    AND D3='TRUE'  ");
            }

            if (TYPE == "D4")
            {
                sb.Append("    AND D4='TRUE'  ");
            }
            if (TYPE == "D5")
            {
                sb.Append("    AND D5='TRUE'  ");
            }
            if (TYPE == "D6")
            {
                sb.Append("    AND D6='TRUE'  ");
            }
            if (TYPE == "D7")
            {
                sb.Append("    AND D7='TRUE'  ");
            }
            if (TYPE == "D8")
            {
                sb.Append("    AND D8='TRUE'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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

        public System.Data.DataTable DDSUMQTY(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT SUM(QTY) QQ   FROM GB_PICK2 T0   WHERE T0.SHIPPINGCODE='GP20170203001X'  AND  ITEMNAME <> 'DM' AND T0.SHIPPINGCODE=@SHIPPINGCODE  ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable DDMESSAGE(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();



            sb.Append("               SELECT DISTINCT LINKMAN 姓名,SUBSTRING(REPLACE(LinkTelephone,'-',''),1,4)+'-'+SUBSTRING(REPLACE(LinkTelephone,'-',''),5,3)+'-'+SUBSTRING(REPLACE(LinkTelephone,'-',''),8,3) 手機門號 ");
            sb.Append("               ,LINKMAN+'會員' 參數一,'' 電子郵件,'' 傳送日期, ");
            sb.Append("               ''''+CAST(CAST(SUBSTRING(UserDef1,5,2) AS INT) AS VARCHAR)+'/'+CAST(CAST(SUBSTRING(UserDef1,7,2) AS INT) AS VARCHAR) 參數二, ");
            sb.Append("               ''''+CAST(CAST(SUBSTRING(PREINDATE,5,2) AS INT) AS VARCHAR)+'/'+CAST(CAST(SUBSTRING(PREINDATE,7,2) AS INT) AS VARCHAR) 參數三, ");
            sb.Append("               USERDEF2 參數四,''''+STORE 參數五 ");
            sb.Append("               FROM GB_PICK2 WHERE LEN(REPLACE(LinkTelephone,'-',''))=10 AND SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append("                AND SUBSTRING(REPLACE(LinkTelephone,'-',''),1,2) = '09' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable DDSBILLNO(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT BILLNO,LINKMAN  FROM GB_PICK2 T0 ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE  ORDER BY BILLNO");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public System.Data.DataTable GPSZ(string SHIPPINGCODE, string EZTYPE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SHIPPINGCODE  FROM GB_PICKEZ WHERE SHIPPINGCODE=@SHIPPINGCODE AND EZTYPE=@EZTYPE ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@EZTYPE", EZTYPE));
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
        public System.Data.DataTable DD2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ITEMCODE,SUM(QTY2) QTY2 FROM GB_PICK T0");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append(" GROUP BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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

        public System.Data.DataTable DD23(string ITEMCODE, string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(QTY2) QTY,SUM(CAST(PACK3 AS DECIMAL(10,2))) KG FROM GB_PICK T0");
            sb.Append(" LEFT JOIN GB_PICK2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE AND ITEMCODE=@ITEMCODE ");
            sb.Append(" GROUP BY ITEMCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        private void WriteExcelProduct(string ExcelFile, string OutPutFile, string SHIPPINGCODE, string TYPE, System.Data.DataTable OrderData2)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            if (TYPE == "批發")
            {
                excelSheet.Name = GetMenu.Day() + "大宗出貨需求";
            }
            else if (TYPE == "零售" || TYPE == "禮盒" || TYPE == "")
            {
                excelSheet.Name = GetMenu.Day() + "小包裝出貨需求";
            }

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iRowCnt2 = 56;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            int iColCnt2 = 7;
            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;

                int Qty = 0;

                DataRow dr;
                DataRow drFind;
                object Cell_From = null;
                object Cell_To = null;

                string sTemp2 = string.Empty;
                string sTemp3 = string.Empty;
                string DEADDATE = string.Empty;
                string FieldValue2 = string.Empty;
                bool IsDetail2 = false;
                int DetailRow2 = 0;

            
                        for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                        {
                            for (int iField = 1; iField <= iColCnt; iField++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                                sTemp2 = (string)range.Text;
                                sTemp2 = sTemp2.Trim();

                                if (ExcelReport.CheckSerial(OrderData2, sTemp2, ref FieldValue2))
                                {
                                    range.Value2 = FieldValue2;
                                }

                                if (IsDetailRow(sTemp2))
                                {
                                    IsDetail2 = true;
                                    DetailRow2 = iRecord;
                                    break;
                                }

                            }

                        }

                        if (DetailRow2 != 0)
                        {

                            for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
                            {

                                //最後一筆不作
                                if (aRow != OrderData2.Rows.Count - 1)
                                {

                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, 1]);
                                    range.EntireRow.Copy(oMissing);

                                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                        oMissing);
                                }


                                for (int iField = 1; iField <= iColCnt; iField++)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, iField]);
                                    //range.Select();
                                    sTemp2 = (string)range.Text;
                                    sTemp2 = sTemp2.Trim();

                                    FieldValue2 = "";
                                    SetRow(OrderData2, aRow, sTemp2, ref FieldValue2);

                                    range.Value2 = FieldValue2;


                                }

                                DetailRow2++;
                            }
                        }

                        System.Data.DataTable G1 = DTBILLNO(SHIPPINGCODE);
             
                        for (int i = 0; i <= G1.Rows.Count - 1; i++)
                        {
                            string BILLNO = G1.Rows[i]["BILLNO"].ToString();
                            string LINKMAN = G1.Rows[i]["LINKMAN"].ToString();
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, i + 6]);
                            range.Value2 = BILLNO;
                            range.Font.Name = "微軟正黑體";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                            range.Columns.AutoFit();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, i + 6]);
                            range.Value2 = LINKMAN;
                            range.Font.Name = "微軟正黑體";
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;


                            System.Data.DataTable G2 = OrderData2;
                            for (int s = 0; s <= G2.Rows.Count - 1; s++)
                            {
                                //12345
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, 2]);
                                sTemp2 = (string)range.Text;
                                sTemp2 = sTemp2.Trim();

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, 3]);
                                sTemp3 = (string)range.Text;
                                sTemp3 = sTemp3.Trim();

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, 4]);
                                DEADDATE = (string)range.Text;
                                DEADDATE = DEADDATE.Trim();

                 
                                System.Data.DataTable G3 = DTBILLNO2(SHIPPINGCODE, BILLNO, sTemp2, sTemp3, DEADDATE);
                                if (G3.Rows.Count > 0)
                                {

                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, i + 6]);
                                    string QTY = G3.Rows[0]["QTY"].ToString();
                                    range.Value2 = QTY;
                                    range.Font.Name = "微軟正黑體";
                                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                    range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                                    range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                    range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;

                                }

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, 3]);
                                sTemp2 = (string)range.Text;
                                sTemp2 = sTemp2.Trim();
                            
                                    for (int L = 1; L <= G1.Rows.Count +5; L++)
                                    {
                                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 3, L]);
                                            int K1 = sTemp2.IndexOf("雞");
                                            if (K1 != -1)
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
                                            }
                                            int K2 = sTemp2.IndexOf("豬");
                                            if (K2 != -1)
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                                            }
                                            int K3 = sTemp2.IndexOf("蝦");
                                            int K31 = sTemp2.IndexOf("魚");
                                            if (K3 != -1 || K31 != -1)
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                                            }
                                            int K4 = sTemp2.IndexOf("加工");
                                            if (K4 != -1)
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
                                            }

                                            int K5 = sTemp2.IndexOf("DM");
                                            if (K5 != -1)
                                            {
                                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                                            }
                                    }
                                
                                
                            }

                             System.Data.DataTable G44 = DTBILLNO33(SHIPPINGCODE, BILLNO);
                             if (G44.Rows.Count > 0)
                             {
                                 PP1 = "";
                                 PPF(SHIPPINGCODE, BILLNO);

                                 range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[G2.Rows.Count + 3, i + 6]);
                                 range.Value2 = G44.Rows[0]["QTY"].ToString();
                                 range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                 range.Font.Name = "微軟正黑體";
                                 range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                 range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                                 range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                 range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;

                                 range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[G2.Rows.Count + 4, i + 6]);
                                 range.Value2 = PP1;
                                 range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                 range.Font.Name = "微軟正黑體";
                                 range.Font.Size = 10;



                                 range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[G2.Rows.Count + 5, i + 6]);
                                 System.Data.DataTable GP1 = DTBILLNOP1(SHIPPINGCODE, BILLNO);
                                 if (GP1.Rows.Count > 0)
                                 {
                                     range.Value2 = GP1.Rows[0]["QTY"].ToString();
                                     range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                     range.Font.Name = "微軟正黑體";
                            
                                 }
                                 else
                                 {
                                     range.Value2 = "";
                                 }

                                 range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[G2.Rows.Count + 6, i + 6]);
                                 System.Data.DataTable GP2 = DTBILLNOP2(SHIPPINGCODE, BILLNO);
                                 if (GP2.Rows.Count > 0)
                                 {
                                     range.Value2 = GP2.Rows[0]["QTY"].ToString();
                                     range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                     range.Font.Name = "微軟正黑體";
                            
                                 }
                                 else
                                 {
                                     range.Value2 = "";
                                 }

                             }
                        }
       

                     
                   
                
                //0523
                    System.Data.DataTable H1 = DTMAIN(SHIPPINGCODE);
        
                      
                  if (H1.Rows.Count > 1)
                  {
                      excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                      excelSheet.Activate();

                      Cell_From = "A1";
                      Cell_To = "I26";
                      excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);

                      for (int i = 1; i <= H1.Rows.Count - 1; i++)
                      {
                          
                              excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                   
                      }


                  }

                  if (H1.Rows.Count > 1)
                  {
                      for (int i = 0; i <= H1.Rows.Count - 2; i++)
                      {
                          int h = i + 2;
                          excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(h);

               

                          excelSheet.Paste(oMissing, oMissing);
                          excelSheet.get_Range(Cell_From, Cell_To).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats,
                                                    Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                          range.ColumnWidth = 4;
                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 2]);
                          range.ColumnWidth = 26.13;
                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 3]);
                          range.ColumnWidth  = 19.13;
                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 4]);
                          range.ColumnWidth = 6.13;
                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 5]);
                          range.ColumnWidth = 6.13;
                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 6]);
                          range.ColumnWidth = 17;
                       

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 1]);
                          range.RowHeight = 37.5;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 1]);
                          range.RowHeight = 33;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, 1]);
                          range.RowHeight = 21;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[8, 1]);
                          range.RowHeight = 21;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[9, 1]);
                          range.RowHeight = 21;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[12, 1]);
                          range.RowHeight = 33;

                          range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, 1]);
                          range.RowHeight = 18;


        

                          
                      }
                  }
                  if (H1.Rows.Count > 1)
                  {
                      for (int i = 0; i <= H1.Rows.Count - 2; i++)
                      {
                          int h = i + 2;
                          excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(h);

                          excelSheet.PageSetup.LeftMargin = 1.5;
                          excelSheet.PageSetup.RightMargin = 1.6;
                          excelSheet.PageSetup.TopMargin = 1;
                          excelSheet.PageSetup.BottomMargin = 0;
                      }
                  }

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                if (H1.Rows.Count > 0)
                {
                    for (int i = 0; i <= H1.Rows.Count - 1; i++)
                    {
                        string BILLNO = H1.Rows[i]["BILLNO"].ToString();
                        System.Data.DataTable OrderData = null;

             
                        string AMT = "0";
                        string CUST = "";
                        System.Data.DataTable DTA = DTAMT(BILLNO, SHIPPINGCODE);
                        if (DTA.Rows.Count > 0)
                        {
                            AMT = DTA.Rows[0][0].ToString();
                            CUST = DTA.Rows[0][1].ToString();
                        }
                        if (checkBox1.Checked)
                        {
                            OrderData = DTF(BILLNO, SHIPPINGCODE, AMT);
                        }
                        else
                        {
                            OrderData = DT(BILLNO, SHIPPINGCODE, AMT);
                        }
                        int K = i + 2;
                        int K1 = i + 1;
                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(K);
                        excelSheet.Activate();

                        excelSheet.Name = BILLNO;

                        string B3 = "//acmew08r2ap//table//SIGN//USER//RG.JPG";
                        excelSheet.Shapes.AddPicture(B3, Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoTrue, 5, 5, 80, 60);
                        for (int iRecord = 1; iRecord <= iRowCnt2; iRecord++)
                        {
                      
                            for (int iField = 1; iField <= iColCnt2; iField++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                                range.Select();
                                sTemp = (string)range.Text;
                                sTemp = sTemp.Trim();

                                if (CheckSerial(OrderData, sTemp, ref FieldValue))
                                {
                                    range.Value2 = FieldValue;
                                }

                                if (IsDetailRow(sTemp))
                                {
                                    IsDetail = true;
                                    DetailRow = iRecord;
                                    break;
                                }

                            }

                        }


                        if (DetailRow != 0)
                        {

                            for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                            {

                                //最後一筆不作
                                if (aRow != OrderData.Rows.Count - 1)
                                {

                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                                    range.EntireRow.Copy(oMissing);

                                    range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                        oMissing);
                                }

                      
                               for (int iField = 1; iField <= 6; iField++)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                                    range.Select();
                                    sTemp = (string)range.Text;
                                    sTemp = sTemp.Trim();

                                    FieldValue = "";

                                    SetRow(OrderData, aRow, sTemp, ref FieldValue);
                                    if (FieldValue == "")
                                    {
                                        FieldValue = " ";
                                    }
                                    range.Value2 = FieldValue.Replace("加工品_", "");


                                }

                                DetailRow++;
                            }

                        }

                    }

                }


              
            }
            finally
            {



                try
                {

                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                System.Diagnostics.Process.Start(OutPutFile);


            }

        }


        private void WriteExcelProductMAN(string ExcelFile, string OutPutFile, string SHIPPINGCODE, System.Data.DataTable OrderData2, string PICKDAY1, string PICKDAY2, string PICKDAY3)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false ;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            excelSheet.Name = "統計表" + GetMenu.Day().Substring(4, 4);
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iRowCnt2 = 56;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;
            int iColCnt2 = 7;
            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string SERIAL_NO;

                int Qty = 0;

                DataRow dr;
                DataRow drFind;
                object Cell_From = null;
                object Cell_To = null;

                string sTemp2 = string.Empty;
      
                string FieldValue2 = string.Empty;
                bool IsDetail2 = false;
                int DetailRow2 = 0;

                //第一行要



                    for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                    {
                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                            sTemp2 = (string)range.Text;
                            sTemp2 = sTemp2.Trim();

                            if (ExcelReport.CheckSerial(OrderData2, sTemp2, ref FieldValue2))
                            {
                                range.Value2 = FieldValue2;
                            }

                            if (IsDetailRow(sTemp2))
                            {
                                IsDetail2 = true;
                                DetailRow2 = iRecord;
                                break;
                            }

                        }

                    }

                    if (DetailRow2 != 0)
                    {

                        for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
                        {
                            if (aRow != OrderData2.Rows.Count - 1)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, 1]);
                                range.EntireRow.Copy(oMissing);

                                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                    oMissing);
                            }


                            for (int iField = 1; iField <= iColCnt; iField++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, iField]);
                                sTemp2 = (string)range.Text;
                                sTemp2 = sTemp2.Trim();

                                FieldValue2 = "";
                                SetRow(OrderData2, aRow, sTemp2, ref FieldValue2);

                                range.Value2 = FieldValue2;


                            }

                            DetailRow2++;
                        }
                    }

                    System.Data.DataTable G1 = DTBILLNOMAN();

                    for (int i = 0; i <= G1.Rows.Count - 1; i++)
                    {
                        string ADDRID = G1.Rows[i]["ADDRID"].ToString();
                        string LINK = G1.Rows[i]["LINK"].ToString();
                        string WALKADDR = G1.Rows[i]["WALKADDR"].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[3, i + 6]);
                        range.Value2 = LINK;
                        range.Font.Name = "微軟正黑體";
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        range.WrapText = true;
                        range.ColumnWidth  = 3.38;
                        if (LINK == "三民" || LINK == "士林")
                        {
                            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        }

                        if (i == G1.Rows.Count - 1)
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        }
                 
          
                        if (WALKADDR == "135")
                        {
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 255));
                        }
                        if (WALKADDR == "246")
                        {
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        int F1 = LINK.IndexOf("宜蘭");
                        int F2 = LINK.IndexOf("基隆");
                        if (F1 != -1 || F2 != -1)
                        {
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 6]);
                        range.Value2 = ADDRID;
                        range.Font.Name = "微軟正黑體";
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);



                        if (LINK == "三民" || LINK == "士林")
                        {
                            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        }

                        if (i == G1.Rows.Count - 1)
                        {
                            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        }
                 
                        if (LINK == "三民" || LINK == "士林")
                        {
                            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, i + 6]);

                            if (LINK == "士林")
                            {
                                range.Value2 = "到貨日每週一.三.五";
                            }
                            if (LINK == "三民")
                            {
                                range.Value2 = "到貨日每週二.四.六";
                            }

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 5]);
                            range.Value2 = PICKDAY1;

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, i + 6]);

                            if (LINK == "士林")
                            {
                                range.Value2 = PICKDAY2;
                            }
                            if (LINK == "三民")
                            {
                                range.Value2 = PICKDAY3;
                            }
                        }

                        System.Data.DataTable G2 = OrderData2;
                        string ITEM = "";
                        string ADD = "";
                        string DESC = "";
                        for (int s = 0; s <= G2.Rows.Count - 1; s++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 5, 1]);
                            sTemp2 = (string)range.Text;
                            ITEM = sTemp2.Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 5, 3]);
                            sTemp2 = (string)range.Text;
                            DESC = sTemp2.Trim();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[4, i + 6]);
                            sTemp2 = (string)range.Text;
                            ADD = sTemp2.Trim();

                            string T1 = "1";
                            int FS2 = DESC.IndexOf("贈品");
                            if (FS2 != -1)
                            {
                                T1="2";
                            }
                            System.Data.DataTable G3 = DTBILLNO2MAN(SHIPPINGCODE, ADD, ITEM, T1);
                            if (G3.Rows.Count > 0)
                            {

                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 5, i + 6]);

                                range.Value2 = G3.Rows[0]["QTY"].ToString();
                                range.Font.Name = "微軟正黑體";
                                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                                if (LINK == "三民" || LINK == "士林")
                                {
                                    range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                }

                                if (i == G1.Rows.Count - 1)
                                {
                                    range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                                }
                            }


                            int FS = DESC.IndexOf("贈品");
                            if (FS != -1)
                            {
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                for (int F = 1; F <= 5; F++)
                                {
                                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[s + 5, F]);
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                }
                            }
                        }
        
                        System.Data.DataTable G4 = DTBILLNO3MAN(SHIPPINGCODE, ADDRID);
                        if (G4.Rows.Count > 0)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[G2.Rows.Count + 5, i + 6]);
                            range.Value2 = G4.Rows[0]["QTY"].ToString();
                            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.Font.Name = "微軟正黑體";
                            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                            if (LINK == "三民" || LINK == "士林")
                            {
                                range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            }

                            if (i == G1.Rows.Count - 1)
                            {
                                range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                            }
                        }
                    }

                        //取出欄位值 - 科目
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[OrderData2.Rows.Count + 9, 5]);
                    range.Select();
                  System.Data.DataTable G5 = DTBILLNO4MAN(SHIPPINGCODE);
                  if (G5.Rows.Count > 0)
                  {
                      range.Value2 = G5.Rows[0][0].ToString();
                  }
                        



                    

            }
            finally
            {



                try
                {

                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }
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

                System.Diagnostics.Process.Start(OutPutFile);


            }

        }

        public static bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }
            //}
            return false;
        }
        public static void SetRow(System.Data.DataTable OrderData, int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[iRow][FieldName]);
            }

        }
        public static bool CheckSerial(System.Data.DataTable OrderData, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //第一個固定欄位(工單號碼)
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("ITEMCODE", typeof(string));
            dt.Columns.Add("ITEMNAME", typeof(string));
            dt.Columns.Add("DOCDATE", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("DEADDATE", typeof(string));
            dt.Columns.Add("總數量", typeof(string));
            //總數量
            return dt;
        }

        private System.Data.DataTable MakeTableFONTAI()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("貨主編號", typeof(string));
            dt.Columns.Add("貨主單號", typeof(string));
            dt.Columns.Add("序號", typeof(string));
            dt.Columns.Add("客戶端代號", typeof(string));
            dt.Columns.Add("訂購日期", typeof(string));
            dt.Columns.Add("預計到貨日", typeof(string));
            dt.Columns.Add("商品編號", typeof(string));
            dt.Columns.Add("商品名稱", typeof(string));
            dt.Columns.Add("倉別", typeof(string));
            dt.Columns.Add("指定效期", typeof(string));
            dt.Columns.Add("指定", typeof(string));
            dt.Columns.Add("訂購數量", typeof(string));
            dt.Columns.Add("商品單價", typeof(string));
            dt.Columns.Add("配送方式", typeof(string));
            dt.Columns.Add("收貨人姓名", typeof(string));
            dt.Columns.Add("收貨人地址", typeof(string));
            dt.Columns.Add("收貨人聯絡電話", typeof(string));
            dt.Columns.Add("日間聯絡電話", typeof(string));
            dt.Columns.Add("夜間聯絡電話", typeof(string));
            dt.Columns.Add("到貨時段", typeof(string));
            dt.Columns.Add("代收金額", typeof(string));
            dt.Columns.Add("宅配單備註", typeof(string));
            dt.Columns.Add("品項備註", typeof(string));
            dt.Columns.Add("單日期", typeof(string));

            
            return dt;
        }
        private void INSEREZ(string ShippingCode, string EZTYPE)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO  [dbo].[GB_PICKEZ]");
            sb.Append("            ([ShippingCode],[EZTYPE],[EZDATE],[EZUSER])");
            sb.Append("      VALUES (@ShippingCode,@EZTYPE,@EZDATE,@EZUSER)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@EZTYPE", EZTYPE));
            command.Parameters.Add(new SqlParameter("@EZDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@EZUSER", fmLogin.LoginID.ToString().Trim().ToUpper()));
  
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

        private void PPF(string SHIPPINGCODE,string BILLNO)

        {
            System.Data.DataTable G4 = DTBILLNO3(SHIPPINGCODE, BILLNO);

            if (G4.Rows.Count > 0)
            {
                double PACKS = 0;
                for (int s = 0; s <= G4.Rows.Count - 1; s++)
                {
                    string ITEMCODE = G4.Rows[s]["ITEMCODE"].ToString();
                    double QTY = Convert.ToDouble(G4.Rows[s]["QTY"]);

                    System.Data.DataTable G5 = DTBILLNO4(ITEMCODE);
                    System.Data.DataTable G6 = DTBILLNO5(ITEMCODE);
                    if (G5.Rows.Count > 0)
                    {
                        PACKS += Convert.ToDouble(G5.Rows[0]["PACK"]) * QTY;

                    }
                    else
                    {
                        string GTYPE = G6.Rows[0]["GTYPE"].ToString();
                        double 包 = Convert.ToDouble(G6.Rows[0]["包"]);
                        double WEIGHT = Convert.ToDouble(G6.Rows[0]["WEIGHT"]);
                        if (GTYPE == "批發")
                        {
                            PACKS += WEIGHT * 2277;
                        }
                        else
                        {
                            if (WEIGHT == 0)
                            {
                                PACKS += WEIGHT * 3795;
                            }
                            else
                            {
                                PACKS += 包 * 1000;
                            }
                        }

                    }
                }



                string Q2 = "";
                double Q1 = PACKS / 45543;
                int RESAULT = Convert.ToInt16(Math.Floor(Q1));
                int Q3 = Convert.ToInt32(Math.Round((((Q1 - Convert.ToDouble(RESAULT)) * 45543)), 0, MidpointRounding.AwayFromZero));
            
                if (Q3 < 5760)
                {
                    Q2 = "S1(綠提盒)";
                }
                else if (Q3 >= 5760 && Q3 < 11550)
                {
                    Q2 = "M2(牛皮)";
                }
                else if (Q3 >= 11550 && Q3 < 19227)
                {
                    Q2 = "M1(牛皮)";
                }
                else if (Q3 >= 19227 && Q3 < 45543)
                {
                    if (RESAULT != 0)
                    {
                        RESAULT = RESAULT + 1;
                    }
                    else
                    {
                        Q2 = "L(牛皮)";
                    }
                }

                if (RESAULT != 0)
                {
                    PP1 = "L(牛皮)*" + RESAULT.ToString() + "+" + Q2;
                    RR = RESAULT;
                }
                else
                {

                    PP1 = Q2;
                }
            }
        }
   
    }
}

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
using System.Net.Mail;
using System.Reflection;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
namespace ACME
{
    public partial class POTAOCHI : Form
    {
        System.Data.DataTable dtGetAcmeStageG = null;
        System.Data.DataTable dtGetAcmeStageG2 = null;
        System.Data.DataTable dtGetAcmeStageG3 = null;
        private StreamWriter sw;
        Attachment data = null;
        string GlobalMailContent = "";
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public POTAOCHI()
        {
            InitializeComponent();
        }
        public System.Data.DataTable D1(string CreateDate, string CreateDate2)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("               SELECT DISTINCT A.BillNO 訂購憑單ID,FullName 訂購人,LinkTelephone 訂購人電話,U.TaxNo 統一編號,CAST(A.SumQty AS DECIMAL(10,2)) 數量,CAST(A.SumAmtATax AS INT) 金額,A.BillDate 訂單日期,A.UserDef1 取貨日期");
            sb.Append("                         ,CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END 交易方式,");
            sb.Append("                         ltrim(substring(A.Remark,CHARINDEX('付款人', A.Remark)+4,20)) 付款人,CASE WHEN A.Remark  LIKE '%2.%' THEN ");
            sb.Append("                         ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,CHARINDEX('2.', A.Remark)-CHARINDEX('1.紙箱DM:', A.Remark)-7)) ");
            sb.Append("                         ELSE ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,20))");
            sb.Append("                         END 有無含DM,");
            sb.Append("                         A.CustBillNo 輔助系統ID  FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append("                         On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append("                              left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1");
            sb.Append("                              Left Join comProduct B On B.ProdID=G.ProdID ");
            sb.Append("            WHERE A.Flag =2 AND  CAST(A.BillDate AS VARCHAR) BETWEEN @CreateDate  AND @CreateDate2 ");
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END = '" + comboBox2.SelectedValue.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar));
            command.Parameters["@CreateDate"].Value = CreateDate;
            command.Parameters.Add(new SqlParameter("@CreateDate2", SqlDbType.VarChar));
            command.Parameters["@CreateDate2"].Value = CreateDate2;
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

        public System.Data.DataTable GETOUTID(string ID)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("     SELECT ORDERPIN FROM GB_POTATO WHERE cast(ID as varchar)=@ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ID", ID));
  
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
        public System.Data.DataTable DD1(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT DISTINCT O.BillNO 銷貨憑單ID FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append("                               On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append("                               left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500 ");
            sb.Append("                  WHERE A.Flag =2 AND  A.BillNO=@BillNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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
        public System.Data.DataTable DD2(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("               SELECT DISTINCT S.UDef2 快遞單號,");
            sb.Append("                         CASE WHEN S.Remark  LIKE '%3.實%' THEN ");
            sb.Append("                         ltrim(substring(S.Remark,CHARINDEX('貨運:', S.Remark)+3,CHARINDEX('3.實', S.Remark)-CHARINDEX('貨運:', S.Remark)-3)) ");
            sb.Append("                         ELSE ltrim(substring(S.Remark,CHARINDEX('貨運:', S.Remark)+3,20))");
            sb.Append("                         END 貨運  FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append("                         On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append("                         left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500 ");
            sb.Append("                             left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)");
            sb.Append("                              left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1");
            sb.Append("                              Left Join comProduct B On B.ProdID=G.ProdID ");
            sb.Append("            WHERE A.Flag =2 AND  A.BillNO=@BillNO ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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

        public System.Data.DataTable DD3(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("         SELECT CAST(SUM(ISNULL(MLAmount,0)) AS INT) 未稅金額,CAST(SUM(ISNULL(MLTaxAmt,0)) AS INT) 稅額  FROM ComProdRec WHERE BillNO =@BillNO AND FLAG=500 ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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


        public System.Data.DataTable DD4(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("         SELECT CAST(SUM(ISNULL(Amount,0)) AS INT) 未稅金額,CAST(SUM(ISNULL(TaxAmt,0)) AS INT) 稅額  FROM OrdBillSub WHERE BillNO =@BillNO AND FLAG=2 ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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
        public System.Data.DataTable DD5(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append("         SELECT CAST(ISNULL(Amount,0) AS INT) 未稅金額 FROM OrdBillSub WHERE BillNO =@BillNO AND FLAG=2 AND ProdID='FREIGHT01' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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
        public static System.Data.DataTable download2()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("     select   Convert(varchar(10),DATEADD(D,0,MAX(a.date_time)),112)     a  from     ");
            sb.Append("       (   select   top  3 *   From   acmesqlsp.dbo.Y_2004   ");
            sb.Append("           where   IsRestDay   =   0   ");
            sb.Append("           and   Convert(varchar(10),date_time,112)    >=    '" + DateTime.Now.ToString("yyyyMMdd") + "' ");
            sb.Append("           order   by   date_time    )   as a   ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


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
        public System.Data.DataTable DD1(string CreateDate, string CreateDate2,string A)
        {
            System.Data.DataTable T1 = download2();
            string DATE = T1.Rows[0][0].ToString();


            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                         SELECT DISTINCT A.BillNO 訂購憑單ID,FullName 訂購人,LinkTelephone 訂購人電話,A.LinkMan 收貨人,A.LinkTelephone 收貨人電話,U.InvoiceHead 收貨人公司,");
            sb.Append("                         TaxNo 統一編號,CAST(A.SumQty AS DECIMAL(10,2)) 數量,CAST(A.SumAmtATax AS INT) 總計,A.CustAddress 交貨地點,");
            sb.Append("                         A.BillDate 訂單日期,A.UserDef1 取貨日期");
            sb.Append("                         ,CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END 交易方式,");
            sb.Append("                         CASE WHEN A.Remark  LIKE '%7.%' THEN ");
            sb.Append("                         ltrim(substring(A.Remark,CHARINDEX('付款人', A.Remark)+4,CHARINDEX('7.', A.Remark)-CHARINDEX('付款人', A.Remark)-4)) ");
            sb.Append("                         ELSE ltrim(substring(A.Remark,CHARINDEX('付款人', A.Remark)+4,20))");
            sb.Append("                         END 付款人,CASE WHEN A.Remark  LIKE '%2.%' THEN ");
            sb.Append("                         ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,CHARINDEX('2.', A.Remark)-CHARINDEX('1.紙箱DM:', A.Remark)-7)) ");
            sb.Append("                         ELSE ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,20))");
            sb.Append("                         END 有無含紙箱DM,A.UserDef2 運送時段     ");
            sb.Append("             ,CASE WHEN A.Remark  LIKE '%6.%' THEN ");
            sb.Append("             ltrim(substring(A.Remark,CHARINDEX('PO:', A.Remark)+3,CHARINDEX('6.', A.Remark)-CHARINDEX('PO:', A.Remark)-3)) ");
            sb.Append("             ELSE ltrim(substring(A.Remark,CHARINDEX('PO:', A.Remark)+3,20))");
            sb.Append("             END 輔助系統ID  ");
            sb.Append("                ,CASE WHEN A.Remark  LIKE '%1.%' THEN ");
            sb.Append("             ltrim(substring(A.Remark,0,CHARINDEX('1.', A.Remark))) ");
            sb.Append("             END 備註");
            sb.Append("             FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append("                         On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append("                              left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1");
            sb.Append("                              Left Join comProduct B On B.ProdID=G.ProdID ");
            sb.Append("                         WHERE A.Flag =2  AND ISNULL(A.UserDef1,'') <> '' ");
            if (A == "1")
            {
                sb.Append("  AND  CAST(A.BillDate AS VARCHAR) BETWEEN @CreateDate AND @CreateDate2");
            }
            if (A == "2")
            {
                sb.Append("   AND  A.UserDef1  <=  '" + DATE + "' AND    ISNULL(S.UDef2,'')='' AND  ISNULL(A.UserDef1,'') > '20140205' ");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CreateDate", CreateDate));
            command.Parameters.Add(new SqlParameter("@CreateDate2", CreateDate2));
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
        private System.Data.DataTable MakeTableCombineGG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂購憑單ID", typeof(string));
            dt.Columns.Add("銷貨憑單ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("統一編號", typeof(string));
            dt.Columns.Add("數量", typeof(decimal));
            dt.Columns.Add("未稅金額", typeof(Int32));
            dt.Columns.Add("稅額", typeof(Int32));
            dt.Columns.Add("總計", typeof(Int32));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("預交日期", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
            dt.Columns.Add("有無含DM", typeof(string));
            dt.Columns.Add("快遞單號", typeof(string));
            dt.Columns.Add("貨運", typeof(string));
            dt.Columns.Add("輔助系統ID", typeof(string));
            dt.Columns.Add("外部網站", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombineGG2()
        {

 


            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂購憑單ID", typeof(string));
            dt.Columns.Add("銷貨憑單ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人公司", typeof(string));
            dt.Columns.Add("統一編號", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("產品", typeof(string));
            dt.Columns.Add("單價", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交貨地點", typeof(string));
            dt.Columns.Add("訂單日期", typeof(string));
            dt.Columns.Add("運送時段", typeof(string));
            dt.Columns.Add("預交日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("實際到貨日期", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
            dt.Columns.Add("輔助系統ID", typeof(string));
            dt.Columns.Add("外部網站", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂購憑單ID", typeof(string));
            dt.Columns.Add("外部網站訂單ID", typeof(string));
            dt.Columns.Add("銷貨憑單ID", typeof(string));
            dt.Columns.Add("訂購人", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(string));
            dt.Columns.Add("運費", typeof(string));
            dt.Columns.Add("總計", typeof(string));
            dt.Columns.Add("交易方式", typeof(string));
            dt.Columns.Add("付款人", typeof(string));
     
            return dt;
        }
        private void TOTAL2GG(System.Data.DataTable dt)
        {
            dtGetAcmeStageG = MakeTableCombineGG();

            System.Data.DataTable DT1 = dt;
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtGetAcmeStageG.NewRow();
                string ID = DT1.Rows[i]["訂購憑單ID"].ToString().Trim();
                dr["訂購憑單ID"] = ID;
       
                dr["訂購人"] = DT1.Rows[i]["訂購人"].ToString().Trim();
                dr["訂購人電話"] = DT1.Rows[i]["訂購人電話"].ToString().Trim();
                dr["統一編號"] = DT1.Rows[i]["統一編號"].ToString().Trim();
                dr["數量"] = Convert.ToDecimal(DT1.Rows[i]["數量"].ToString().Trim());
                dr["訂單日期"] = DT1.Rows[i]["訂單日期"].ToString().Trim();
                dr["取貨日期"] = DT1.Rows[i]["取貨日期"].ToString().Trim();
                StringBuilder sbM = new StringBuilder();
                System.Data.DataTable M1 = D3PREINDATE(ID);
                if (M1.Rows.Count > 0)
                {
                    for (int s = 0; s <= M1.Rows.Count - 1; s++)
                    {

                        DataRow dd = M1.Rows[s];


                        sbM.Append(dd["PreInDate"].ToString() + "/");


                    }

                    sbM.Remove(sbM.Length - 1, 1);

                    dr["預交日期"] = sbM.ToString();
                }
          
                dr["交易方式"] = DT1.Rows[i]["交易方式"].ToString().Trim();
                string PAY = DT1.Rows[i]["付款人"].ToString().Trim();
                dr["付款人"] = PAY;
                int T1 = PAY.IndexOf("7.");
                if (T1 != -1)
                {
                    dr["付款人"] = PAY.Substring(0, T1);
                }
                dr["有無含DM"] = DT1.Rows[i]["有無含DM"].ToString().Trim();
                string ASID = DT1.Rows[i]["輔助系統ID"].ToString().Trim();
                dr["輔助系統ID"] = ASID;
                System.Data.DataTable J1 = GETOUTID(ASID);
                if (J1.Rows.Count > 0)
                {
                    dr["外部網站"] = J1.Rows[0][0].ToString().Trim();
                }
                
                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();
                StringBuilder sb3 = new StringBuilder();
                System.Data.DataTable dtt = DD1(ID);
                Int32 R1 = 0;
                Int32 R2 = 0;
                Int32 R3 = 0;

                try
                {
                    if (dtt.Rows.Count > 0)
                    {
                        for (int s = 0; s <= dtt.Rows.Count - 1; s++)
                        {

                            DataRow dd = dtt.Rows[s];
                            string G1 = dd["銷貨憑單ID"].ToString().Trim();
                            if (G1 == "DN1070810003")
                            {
                                MessageBox.Show("A");
                            }
                            sb.Append(G1 + "/");

                            System.Data.DataTable dtt4 = DD4(ID);
                            DataRow dd1 = dtt4.Rows[0];
                            R1 = Convert.ToInt32(dd1["未稅金額"]);
                            R2 = Convert.ToInt32(dd1["稅額"]);

                            System.Data.DataTable dtt5 = DD5(ID);
                            if (dtt5.Rows.Count > 0)
                            {
                                int GF1 = Convert.ToInt16(Convert.ToDouble(dtt5.Rows[0][0].ToString()) * 1.05);
                                int GF2 = Convert.ToInt16(dtt5.Rows[0][0].ToString());
                                R2 = GF1 - GF2;
                            }
  
                        }

                        sb.Remove(sb.Length - 1, 1);
                        R3 += R1 + R2;
                        dr["銷貨憑單ID"] = sb.ToString();
                        dr["未稅金額"] = R1;
                        dr["稅額"] = R2;
                        dr["總計"] = R3;
                    }
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + dr["銷貨憑單ID"].ToString());
                }

                System.Data.DataTable dtt2 = DD2(ID);
                if (dtt2.Rows.Count > 0)
                {
                    for (int s = 0; s <= dtt2.Rows.Count - 1; s++)
                    {

                        DataRow dd = dtt2.Rows[s];
                        sb2.Append(dd["快遞單號"].ToString().Trim() + "/");
                        sb3.Append(dd["貨運"].ToString().Trim() + "/");
                    }

                    sb2.Remove(sb2.Length - 1, 1);
                    sb3.Remove(sb3.Length - 1, 1);

                    dr["銷貨憑單ID"] = sb.ToString();
                    dr["快遞單號"] = sb2.ToString();
                    dr["貨運"] = sb3.ToString();
                }

                dtGetAcmeStageG.Rows.Add(dr);
            }

           
        }
        private void TOTAL2GG2(System.Data.DataTable dt)
        {
            dtGetAcmeStageG3 = MakeTableCombineGG2();

            System.Data.DataTable DT1 = dt;
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtGetAcmeStageG3.NewRow();

                dr["訂購憑單ID"] = DT1.Rows[i]["訂購憑單ID"].ToString().Trim();
                dr["銷貨憑單ID"] = DT1.Rows[i]["銷貨憑單ID"].ToString().Trim();
                dr["訂購人"] = DT1.Rows[i]["訂購人"].ToString().Trim();
                dr["訂購人電話"] = DT1.Rows[i]["訂購人電話"].ToString().Trim();
                dr["收貨人"] = DT1.Rows[i]["收貨人"].ToString().Trim();
                dr["收貨人電話"] = DT1.Rows[i]["收貨人電話"].ToString().Trim();
                dr["收貨人公司"] = DT1.Rows[i]["收貨人公司"].ToString().Trim();
                dr["統一編號"] = DT1.Rows[i]["統一編號"].ToString().Trim();
                dr["類型"] = DT1.Rows[i]["類型"].ToString().Trim();
                dr["數量"] = DT1.Rows[i]["數量"].ToString().Trim();
                dr["產品"] = DT1.Rows[i]["產品"].ToString().Trim();
                dr["單價"] = DT1.Rows[i]["單價"].ToString().Trim();
                dr["運費"] = DT1.Rows[i]["運費"].ToString().Trim();
                dr["總計"] = DT1.Rows[i]["總計"].ToString().Trim();
                dr["交貨地點"] = DT1.Rows[i]["交貨地點"].ToString().Trim();
                dr["訂單日期"] = DT1.Rows[i]["訂單日期"].ToString().Trim();

                dr["運送時段"] = DT1.Rows[i]["運送時段"].ToString().Trim();
                dr["預交日期"] = DT1.Rows[i]["預交日期"].ToString().Trim();
                dr["取貨日期"] = DT1.Rows[i]["取貨日期"].ToString().Trim();
                dr["實際到貨日期"] = DT1.Rows[i]["實際到貨日期"].ToString().Trim();
                dr["交易方式"] = DT1.Rows[i]["交易方式"].ToString().Trim();
                dr["付款人"] = DT1.Rows[i]["付款人"].ToString().Trim();
                dr["備註"] = DT1.Rows[i]["備註"].ToString().Trim();


                string ASID = DT1.Rows[i]["輔助系統ID"].ToString().Trim();
                dr["輔助系統ID"] = ASID;
                System.Data.DataTable J1 = GETOUTID(ASID);
                if (J1.Rows.Count > 0)
                {
                    dr["外部網站"] = J1.Rows[0][0].ToString().Trim();
                }



                dtGetAcmeStageG3.Rows.Add(dr);
            }


        }
        private void ACCR(System.Data.DataTable dt)
        {
            dtGetAcmeStageG2 = MakeTableCombine();

            System.Data.DataTable DT1 = dt;
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
            
                
                dr = dtGetAcmeStageG2.NewRow();

                string ID = DT1.Rows[i]["訂購憑單ID"].ToString().Trim();
                string ROWNO = DT1.Rows[i]["ROWNO"].ToString().Trim();
                dr["訂購憑單ID"] = ID;
                StringBuilder sb = new StringBuilder();
                Int32 R1 = 0;
                Int32 R2 = 0;
                Int32 R3 = 0;
                if (ROWNO == "1")
                {
                    System.Data.DataTable dtt = DD1(ID);
                    try
                    {
                        if (dtt.Rows.Count > 0)
                        {
                            for (int s = 0; s <= dtt.Rows.Count - 1; s++)
                            {

                                DataRow dd = dtt.Rows[s];
                                string G1 = dd["銷貨憑單ID"].ToString().Trim();
                                sb.Append(G1 + "/");

                                System.Data.DataTable dtt4 = DD4(ID);
                                DataRow dd1 = dtt4.Rows[0];

                                R1 = Convert.ToInt32(dd1["未稅金額"]);
                                R2 = Convert.ToInt32(dd1["稅額"]);

                                System.Data.DataTable dtt5 = DD5(ID);
                                if (dtt5.Rows.Count > 0)
                                {
                                    int GF1 = Convert.ToInt16(Convert.ToDouble(dtt5.Rows[0][0].ToString()) * 1.05);
                                    int GF2 = Convert.ToInt16(dtt5.Rows[0][0].ToString());
                                    R2 = GF1 - GF2;
                                }

                            }

                            sb.Remove(sb.Length - 1, 1);
                            R3 += R1 + R2;
                            dr["銷貨憑單ID"] = sb.ToString();
                            dr["運費"] = R2;
                            dr["總計"] = R3;
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + dr["銷貨憑單ID"].ToString());
                    }
                }
                dr["訂購人"] = DT1.Rows[i]["訂購人"].ToString().Trim();
                dr["數量"] = DT1.Rows[i]["數量"].ToString().Trim();
                dr["單價"] = DT1.Rows[i]["單價"].ToString().Trim();

                dr["交易方式"] = DT1.Rows[i]["交易方式"].ToString().Trim();
           
                string PAY=DT1.Rows[i]["付款人"].ToString().Trim();
                dr["付款人"] = PAY;
                int T1=PAY.IndexOf("7.");
                if (T1 != -1)
                {
                    dr["付款人"] = PAY.Substring(0, T1);
                }
                string ASID = DT1.Rows[i]["輔助系統ID"].ToString().Trim(); ;
                System.Data.DataTable J1 = GETOUTID(ASID);
                if (J1.Rows.Count > 0)
                {
                    dr["外部網站訂單ID"] = J1.Rows[0][0].ToString().Trim();
                }


                dtGetAcmeStageG2.Rows.Add(dr);
            }


        }
        public System.Data.DataTable D2(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ProdID 產品編號,ProdName 產品名稱,CAST(Quantity AS decimal(10,2)) 數量,CAST(Price AS INT) 單價, CAST(Amount AS INT) 金額  FROM OrdBillSub WHERE BillNO=@BillNO AND Flag =2");
            

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

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



        public System.Data.DataTable GETACC(string CreateDate, string CreateDate2)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable T1 = download2();
            string DATE = T1.Rows[0][0].ToString();

            sb.Append("                      SELECT  A.BillNO 訂購憑單ID,FullName 訂購人, G.ProdName 產品,G.Quantity 數量,G.Price 單價,G.SerNO ROWNO,");
            sb.Append("                                                                  CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END 交易方式,  ");
            sb.Append("                                                                 ltrim(substring(A.Remark,CHARINDEX('付款人', A.Remark)+4,20)) 付款人,   A.CustBillNo 輔助系統ID ");
            sb.Append("                                                               FROM  OrdBillMain A    ");
            sb.Append("                                                               Inner Join OrdBillSub G On G.Flag=A.Flag  And G.BillNO=A.BillNO   ");
            sb.Append("                                                                      left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1  ");
            sb.Append("                                                                      Left Join comProduct B On B.ProdID=G.ProdID     ");
            sb.Append("                                                               WHERE       A.Flag =2   AND  CAST(A.BillDate AS VARCHAR) BETWEEN @CreateDate  AND @CreateDate2   ");
            if (comboBox2.SelectedValue.ToString() != "Please-Select")
            {
                sb.Append("and CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END = '" + comboBox2.SelectedValue.ToString() + "'  ");
            }

            sb.Append(" ORDER BY FullName,A.BillNO,G.SerNO  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar));
            command.Parameters["@CreateDate"].Value = CreateDate;
            command.Parameters.Add(new SqlParameter("@CreateDate2", SqlDbType.VarChar));
            command.Parameters["@CreateDate2"].Value = CreateDate2;


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

        public System.Data.DataTable download12DT(string CreateDate, string CreateDate2, string A)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            System.Data.DataTable T1 = download2();
            string DATE = T1.Rows[0][0].ToString();

            sb.Append("        SELECT  A.BillNO 訂購憑單ID,O.BillNO  銷貨憑單ID,FullName 訂購人,LinkTelephone 訂購人電話,A.LinkMan 收貨人,A.LinkTelephone 收貨人電話,U.InvoiceHead 收貨人公司, ");
            sb.Append("                                                    TaxNo 統一編號,    CASE SUBSTRING(G.ProdID,1,3) WHEN 'MCK' THEN '雞' WHEN 'MPK' THEN '豬' END 類型,G.ProdName 產品,            CASE WHEN B.ProdID IN ('MCK010101','MCK020101') THEN CAST(G.Quantity AS DECIMAL(10,2))*6 ELSE   CAST(G.Quantity AS DECIMAL(10,2)) END  數量,G.Price 單價,CASE G.SerNO WHEN 1 THEN ISNULL(T1.FEE,0) END 運費,CASE G.SerNO WHEN 1 THEN ISNULL(T2.AMT,0)+ISNULL(T1.FEE,0) END 總計, ");
            sb.Append("                                                A.CustAddress 交貨地點, ");
            sb.Append("                                                    A.BillDate 訂單日期,A.UserDef1 取貨日期,G.PreInDate 預交日期,S.UDef1 實際到貨日期, ");
            sb.Append("                                                    CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END 交易方式, ");
            sb.Append("                                                   ltrim(substring(A.Remark,CHARINDEX('付款人', A.Remark)+4,20)) 付款人,CASE WHEN A.Remark  LIKE '%2.%' THEN  ");
            sb.Append("                                                    ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,CHARINDEX('2.', A.Remark)-CHARINDEX('1.紙箱DM:', A.Remark)-7))  ");
            sb.Append("                                                    ELSE ltrim(substring(A.Remark,CHARINDEX('1.紙箱DM:', A.Remark)+7,20)) ");
            sb.Append("                                                    END 有無含紙箱DM,A.UserDef2 運送時段,a.CustBillNo 輔助系統ID   ");
            sb.Append("                                           ,CASE WHEN A.Remark  LIKE '%1.%' THEN ltrim(substring(A.Remark,0,CHARINDEX('1.', A.Remark))) END 備註 ");
            sb.Append("                                                 FROM  OrdBillMain A   ");
            sb.Append("                                                 Inner Join OrdBillSub G On G.Flag=A.Flag  And G.BillNO=A.BillNO  ");
            sb.Append("                                                   left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500  ");
            sb.Append("                                                       left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500) ");
            sb.Append("                                                        left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1 ");
            sb.Append("                                                        Left Join comProduct B On B.ProdID=G.ProdID    ");
            sb.Append("                                                        LEFT JOIN (SELECT  max(CAST(Amount*1.05 AS INT))FEE,BillNO FROM OrdBillSub  WHERE SUBSTRING(ProdID,1,3) = 'FRE' group by BillNO) T1 ON (A.BillNO=T1.BillNO) ");
            sb.Append("                                                                                        LEFT JOIN (SELECT BillNO,CAST(SUM(Quantity*Price) AS INT) AMT  FROM OrdBillSub  WHERE SUBSTRING(ProdID,1,3) <> 'FRE' group by BillNO) T2 ON (A.BillNO=T2.BillNO) ");
            sb.Append("                                                 WHERE       A.Flag =2   ");
            if (A == "1")
            {
                if (comboBox2.SelectedValue.ToString() != "Please-Select")
                {
                    sb.Append(" and CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN A.GatherOther END = '" + comboBox2.SelectedValue.ToString() + "'  ");
                }
                sb.Append("  AND  CAST(A.BillDate AS VARCHAR) BETWEEN @CreateDate AND @CreateDate2 AND SUBSTRING(G.ProdID,1,3) <> 'FRE'");
            }
            if (A == "2")
            {
                sb.Append("   AND  G.PreInDate  <=  '" + DATE + "' AND    ISNULL(S.UDef2,'')='' AND  ISNULL(A.UserDef1,'') > '20140205'  ");
                sb.Append("                                     AND G.QtyRemain <> 0 AND  SUBSTRING(G.ProdID,1,3) <> 'FRE'");

             
            }

            sb.Append(" ORDER BY A.BillNO,G.SerNO  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
  

            command.Parameters.Add(new SqlParameter("@CreateDate", SqlDbType.VarChar));
            command.Parameters["@CreateDate"].Value = CreateDate;
            command.Parameters.Add(new SqlParameter("@CreateDate2", SqlDbType.VarChar));
            command.Parameters["@CreateDate2"].Value = CreateDate2;


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
        public System.Data.DataTable D3(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  DISTINCT A.LinkMan 聯絡人,A.LinkTelephone 電話,U.InvoiceHead 公司,A.CustAddress 地址,A.UserDef2 運送時段,A.UserDef1 取貨日期,S.UDef1 實際到貨日期");
            sb.Append("  FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append(" On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append(" left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500 ");
            sb.Append("     left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)");
            sb.Append("      left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1");
            sb.Append("      Left Join comProduct B On B.ProdID=G.ProdID ");
            sb.Append(" WHERE A.Flag =2 AND  A.BillNO =@BillNO");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

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
        public System.Data.DataTable D3PREINDATE(string BillNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT PreInDate FROM OrdBillSub WHERE BillNO=@BillNO AND Flag =2   ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));

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
        private void POTAOCHI_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox2, GetTRADE(), "DataValue", "DataValue");
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            comboBox1.Text = "快遞單號";
            comboBox2.Text = "Please-Select";
            TOTAL2GG(D1(textBox1.Text, textBox2.Text));
            dataGridView1.DataSource = dtGetAcmeStageG;

            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                if (dtGetAcmeStageG.Columns[i].DataType == typeof(Int32))
                {
                    SetDefaultStyle_Int(dataGridView1.Columns[i]);
                }
                else if (dtGetAcmeStageG.Columns[i].DataType == typeof(Decimal))
                {
                    SetDefaultStyle_Numeric(dataGridView1.Columns[i]);
                }
            }
        }

        public  System.Data.DataTable GetTRADE()
        {

            SqlConnection con = new SqlConnection(strCn);
            string sql = " SELECT * FROM (SELECT 'Please-Select' DataValue UNION ALL SELECT DISTINCT CASE GatherStyle WHEN 0 THEN '貨到付款' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN GatherOther END DataValue FROM OrdBillMain WHERE GatherOther <> '3' AND Flag =2 AND BillDate >20140101 ) AS A WHERE DataValue <> '' ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "oslp");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["oslp"];
        }
        private void SetDefaultStyle_Int(DataGridViewColumn Column)
        {
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle.Format = "#,##0";
            dataGridViewCellStyle.NullValue = null;
            Column.DefaultCellStyle = dataGridViewCellStyle;
        }
        private void SetDefaultStyle_Numeric(DataGridViewColumn Column)
        {
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle.Format = "#,##0.00";
            dataGridViewCellStyle.NullValue = null;
            Column.DefaultCellStyle = dataGridViewCellStyle;
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {

                string ID = dataGridView1.SelectedRows[0].Cells["訂購憑單ID"].Value.ToString();

                dataGridView2.DataSource = D2(ID);
                dataGridView3.DataSource = D3(ID);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TOTAL2GG(D1(textBox1.Text, textBox2.Text));

            if (textBox3.Text != "")
            {
                dtGetAcmeStageG.DefaultView.RowFilter = " [外部網站] = ('" + textBox3.Text + "') ";
            }
            dataGridView1.DataSource = dtGetAcmeStageG;
              
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                string F = opdf.FileName;

                GetExcelContentGD4(F, comboBox1.Text);


            }
        }
        private void GetExcelContentGD4(string ExcelFile, string T1)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2 = "";
            string id3 = "";

            int u = 0;
            int v = 0;
            //for (int b = 5; b <= 20; b++)
            //{
            //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[7, b]);
            //    range.Select();
            //    id = range.Text.ToString();

            //}


            for (int jj = 1; jj <= 30; jj++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, jj]);
                range.Select();
                id = range.Text.ToString();

                if (id.Trim().ToUpper() == "銷貨憑單ID")
                {

                    u = jj;
                }

                if (id.Trim() == T1)
                {

                    v = jj;
                }

            }


            if (u == 0 || v == 0)
            {
                MessageBox.Show("Excel格式有誤");
                return;

            }


            try
            {


                for (int j = 2; j <= iRowCnt; j++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, u]);
                    range.Select();
                    id2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, v]);
                    range.Select();
                    id3 = range.Text.ToString().Trim();



                    if (!String.IsNullOrEmpty(id2))
                    {

                        UPDATESAP(id3, id2, T1);

                    }


                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            MessageBox.Show("匯入成功");
        }
        public void UPDATESAP(string OrdNo, string ID, string TYPE)
        {
            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = null;
            if (TYPE == "快遞單號")
            {

                command = new SqlCommand("UPDATE  COMBILLACCOUNTS SET UDef2=@OrdNo WHERE Flag=500 AND FundBillNo=@FundBillNo ", connection);
            }

            if (TYPE == "實際到貨日期")
            {

                command = new SqlCommand("UPDATE  COMBILLACCOUNTS SET UDef1=@OrdNo WHERE Flag=500 AND FundBillNo=@FundBillNo ", connection);
            }
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@OrdNo", OrdNo));
            command.Parameters.Add(new SqlParameter("@ID", ID));



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

        private void button3_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\GW\\生技出貨排程.xls";
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = "";

            OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-生技出貨排程" + ".xls";
            TOTAL2GG2(download12DT(textBox1.Text, textBox2.Text, "1"));
            ExcelReport.ExcelReportOutput(dtGetAcmeStageG3, ExcelTemplate, OutPutFile, "N");
           
            //dtGetAcmeStageG
 
        }
        private void DELETEFILE()
        {
            try
            {
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

        private void SendByMailList()
        {


            string MailDate = DateTime.Now.ToString("yyyyMMdd");

            string strSubject = "";
            string SlpName = "LleytonChen";
            string MailAddress = "LleytonChen@acmepoint.com";
            string MailContent = GlobalMailContent;

            string BuGroup = "聿豐實業";

            string SysCode = "GD_POTATO";
            System.Data.DataTable dt = GetACME_MAILLIST(SysCode);
          

                strSubject = "生技出貨排程";

                DELETEFILE();


                System.Data.DataTable H1 = download12DT(textBox1.Text, textBox2.Text, "2");

                if (H1.Rows.Count > 0)
                {

                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\GW\\生技出貨排程.xls";
                    string ExcelTemplate = FileName;

                    //輸出檔
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-生技出貨排程" + ".xls";

                    //產生 Excel Report
                    ExcelReport.ExcelReportOutputJ2(H1, ExcelTemplate, OutPutFile);

                    //  UpdatINV();
                }
                else
                {
                    MessageBox.Show("沒有資料");
                    return;
                }

                DataRow dr;

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    dr = dt.Rows[i];


                    SlpName = Convert.ToString(dr["UserCode"]);

                    MailAddress = Convert.ToString(dt.Rows[i]["UserMail"]);
    
                    if (string.IsNullOrEmpty(MailAddress))
                    {
                        MailAddress = "lleytonchen@acmepoint.com";
                    }


                    MailContent = string.Format("時間->{0} 朝貢雞排程完成", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));


                    MailTest2(strSubject, SlpName, MailAddress, MailContent);

                


      
                //log
                MailAddress = "lleytonchen@acmepoint.com";
                SlpName = "lleytonchen";

                MailContent = string.Format("時間->{0} 朝貢雞排程完成", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                strSubject = string.Format("{0}朝貢雞排程", BuGroup);








            }
                MessageBox.Show("寄送成功");
        }

        public System.Data.DataTable GetACME_MAILLIST(string SysCode)
        {
            SqlConnection connection = globals.Connection;

            string sql = "SELECT UserCode,UserMail FROM ACME_ARES_MAIL where Active='Y' and SysCode=@SysCode";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SysCode", SysCode));
            //command.Parameters.Add(new SqlParameter("@MailDate", MailDate));
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ACME_MAIL");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["ACME_MAIL"];
        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void MailTest2(string strSubject, string SlpName, string MailAddress, string MailContent)
        {
            MailMessage message = new MailMessage();
            string FROM = fmLogin.LoginID.ToString() + "@acmepoint.com";
            message.From = new MailAddress(FROM, "系統發送");
            message.To.Add(new MailAddress(MailAddress));



            string template;
            StreamReader objReader;


            objReader = new StreamReader(GetExePath() + "\\MailTemplates\\POTATO.htm");

            template = objReader.ReadToEnd();
            objReader.Close();
            template = template.Replace("##FirstName##", SlpName);
            template = template.Replace("##LastName##", "");
            template = template.Replace("##Company##", "聿豐實業");
            template = template.Replace("##Content##", MailContent);

            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = template;
            //格式為 Html
            message.IsBodyHtml = true;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                string m_File = "";

                m_File = file;
                data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                //附件资料
                ContentDisposition disposition = data.ContentDisposition;


                // 加入邮件附件
                message.Attachments.Add(data);

            }

            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";
       
            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            //client.Send(message);

            try
            {
                client.Send(message);
                data.Dispose();
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);
                    }
                    else
                    {
                        SetMsg(String.Format("Failed to deliver message to {0}",
                            ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                        ex.ToString()));
            }

        }
        private void WriteToLog(StreamWriter sw, string Msg)
        {
            // StreamWriter sw = new StreamWriter("file.html", true, Encoding.UTF8);//creating html file
            sw.Write(Msg);
            // sw.Close();
        }
        private void SetMsg(string Msg)
        {
            label1.Text = "處理訊息:" + Msg;
            label1.Refresh();
            WriteToLog(sw, label1.Text + "\r\n");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("確定是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                SendByMailList();
                
            }

        
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\GW\\生技出貨排程2.xls";
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = "";

            OutPutFile = lsAppDir + "\\Excel\\temp\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-生技出貨排程" + ".xls";
            ACCR(GETACC(textBox1.Text, textBox2.Text));
            ExcelReport.ExcelReportOutput(dtGetAcmeStageG2, ExcelTemplate, OutPutFile, "N");
        }

     
    }
}

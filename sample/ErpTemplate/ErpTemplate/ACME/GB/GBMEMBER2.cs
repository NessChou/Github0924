using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class GBMEMBER2 : Form
    {

        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBMEMBER2()
        {
            InitializeComponent();
        }

        private void GBMEMBER_Load(object sender, EventArgs e)
        {
            textBox5.Text = GetMenu.DFirst();
            textBox6.Text = GetMenu.DLast();

            AddGB_MEMBER();
            System.Data.DataTable F1 = GetTOTAL();
            System.Data.DataTable F2 = GetTOTAL2();
            int L1 = Convert.ToInt32(F1.Rows[0][0]);
            int L2 = 0;
            for (int i = 0; i <= F2.Rows.Count - 1; i++)
            {
                L2 += Convert.ToInt32(F2.Rows[i][0]);
                if (L1 > L2)
                {
                    AddGB_MEMBER2(F2.Rows[i][1].ToString(), Convert.ToInt32(F2.Rows[i][0]), F2.Rows[i][2].ToString(), F2.Rows[i][3].ToString());
                }
          
            }


        }


        private System.Data.DataTable GetMEM()
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" Select A.BILLNO,A.CustBillNo CUSTNO,A.CustomerID CUSTID,L.ClassName 客戶類別,case when a.CustomerID = 'tw90146-16' then T11.LinkManProf  ELSE  M.AddField1 END 來源,CASE CHARINDEX('-',FullName) WHEN 0 THEN B.FullName ELSE substring(B.FullName,0,CHARINDEX('-', B.FullName)) END    簡稱, ");
            sb.Append(" B.Telephone1  訂購人電話,A.LinkMan 收貨人,  A.LinkTelephone 收貨人電話,A.CustAddress 收貨人地址,A.BillDate 訂購日期,isnull((A.SumAmtATax),0) 訂購金額,A.UserDef1 取貨日期 ");
            sb.Append(" ,case when l.ClassID  in (19,20,21,26,27,28,29) THEN 'P' ELSE '' END FTYPE,S.PersonName  SALES  From OrdBillMain A     ");
            sb.Append(" Left Join comCustomer B On B.Flag=A.Flag-1 And B.ID=A.CustomerID        ");
            sb.Append(" Left Join comCustDesc M On B.ID =M.ID and M.Flag =1  ");
            sb.Append(" LEFT JOIN comCustAddress T11 ON (A.AddressID=T11.AddrID AND A.CustomerID=T11.ID )   ");
            sb.Append(" Left Join comCustClass L On L.ClassID =b.ClassID and L.Flag =1 ");
            sb.Append(" LEFT JOIN comPerson S ON (B.PersonID =S.PersonID)");
            sb.Append(" Where A.Flag=2 and L.ClassName not in ('棉花田','進金生集團','批發','安永鮮物','分切廠','公司企業')  AND A.SumBTaxAmt <> 0 AND A.BillStatus <> 2  ");
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("  AND (B.FullName  LIKE  '%" + textBox1.Text + "%' OR A.LinkMan  LIKE  '%" + textBox1.Text + "%')   ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("  AND (S.PersonName   LIKE  '%" + textBox3.Text + "%' OR S.ENGNAME  LIKE  '%" + textBox3.Text + "%')   ");
            }
            sb.Append("  ORDER BY A.LinkMan  ");
            //A.LinkMan 
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
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


        private System.Data.DataTable GetMEM2(string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                                                                               Select  COUNT (DISTINCT L.ClassName)   G   From OrdBillMain A     ");
            sb.Append("                                                                            Left Join comCustomer B On B.Flag=A.Flag-1 And B.ID=A.CustomerID        ");
            sb.Append("                                                                                         Left Join comCustClass L On L.ClassID =b.ClassID and L.Flag =1 ");
            sb.Append("                                                Where A.Flag=2 and L.ClassName not in ('棉花田','進金生集團','批發','安永鮮物','分切廠','公司企業')  AND A.SumBTaxAmt <> 0 AND A.BillStatus <> 2  ");
            if (textBox5.Text != "" && textBox6.Text != "")
            {

                sb.Append("  AND A.BillDate  between @BillDate1 and @BillDate2 ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("  AND (B.FullName  LIKE  '%" + textBox1.Text + "%' OR A.LinkMan  LIKE  '%" + textBox1.Text + "%')   ");
            }

            sb.Append("       and A.LinkMan=@LinkMan  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillDate1", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BillDate2", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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

        private System.Data.DataTable GetMEMS(string BillNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("                                  SELECT InvoProdName+'_'+CAST(CAST(T0.Quantity AS INT) AS VARCHAR) PROD    FROM OrdBillSub T0 ");
            sb.Append("                                               Left Join comProduct T1 On T0.ProdID=T1.ProdID    ");
            sb.Append("                                  WHERE Flag=2 and  BillNO  =@BillNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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

        private System.Data.DataTable GetMEMS2(string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                        SELECT RIVACOUPON FROM GB_POTATO WHERE CAST(ID AS VARCHAR)=@ID");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
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
        private System.Data.DataTable GetTEL(string ORDNAME)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT ORDTEL FROM GB_POTATO WHERE ORDNAME=@ORDNAME AND ISNULL(ORDTEL,'') <> ''  ORDER BY ID DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ORDNAME", ORDNAME));
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
        private System.Data.DataTable GetMEMS3(string CustomerID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("                                  SELECT BILLDATE    FROM OrdBillMAIN      WHERE Flag=2 AND CustomerID=@CustomerID   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
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

        private System.Data.DataTable GetTOTAL()
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("        Select SUM(A.SumAmtATax)*0.5 AMT   From OrdBillMain A ");
            sb.Append("                        left join comCustomer U On  U.ID=a.CustomerID  AND U.Flag =1    ");
            sb.Append("                            Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1   ");
            sb.Append("                      Where A.Flag=2 and l.ENGNAME  in ('c','C_Web') AND L.ClassID <> '013' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
        private System.Data.DataTable GetTOTAL2()
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("              SELECT SUM(AMT) AMT,LINKMAN,MAX(客戶類別) 客戶類別,MAX(來源) 來源 FROM (     Select SUM(A.SumAmtATax) AMT,CASE CHARINDEX('-',FullName) WHEN 0 THEN FullName ELSE substring(FullName,0,CHARINDEX('-', FullName)) END LINKMAN,max(L.ClassName) 客戶類別,max(case when a.CustomerID = 'tw90146-16' then T11.LinkManProf  ELSE  M.AddField1 END) 來源   From OrdBillMain A  ");
            sb.Append("                                      left join comCustomer U On  U.ID=a.CustomerID  AND U.Flag =1     ");
            sb.Append("                                          Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1    ");
            sb.Append("                                          Left Join comCustDesc M On u.ID =M.ID and M.Flag =1  ");
            sb.Append("                                            LEFT JOIN comCustAddress T11 ON (A.AddressID=T11.AddrID AND A.CustomerID=T11.ID )   ");
            sb.Append("                           Where A.Flag=2 and l.ENGNAME  = ('c') AND L.ClassID <> '013' GROUP BY L.ClassName,CASE CHARINDEX('-',FullName) WHEN 0 THEN FullName ELSE substring(FullName,0,CHARINDEX('-', FullName)) END  ");
            sb.Append("                           UNION ALL ");
            sb.Append("                             Select SUM(A.SumAmtATax) AMT,A.LinkMan,max(L.ClassName) 客戶類別,max(case when a.CustomerID = 'tw90146-16' then T11.LinkManProf  ELSE  M.AddField1 END) 來源    From OrdBillMain A  ");
            sb.Append("                                      left join comCustomer U On  U.ID=a.CustomerID  AND U.Flag =1     ");
            sb.Append("                                      Left Join comCustDesc M On u.ID =M.ID and M.Flag =1  ");
            sb.Append("                                          Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1    ");
            sb.Append("                                            LEFT JOIN comCustAddress T11 ON (A.AddressID=T11.AddrID AND A.CustomerID=T11.ID )   ");
            sb.Append("                           Where A.Flag=2 and l.ENGNAME  = ('C_Web') AND L.ClassID <> '013' GROUP BY A.LinkMan ) AS A GROUP BY LINKMAN ");
            sb.Append("                           ORDER BY SUM(AMT) DESC ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
        private System.Data.DataTable GetTOTAL3()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("         SELECT ID 排名,CUSTTYPE,CUSTTYPE2,CUSTNAME 客戶,AMOUNT 金額 FROM GB_MEMBER ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
        private System.Data.DataTable GetMEMS4(string CustomerID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("                                                                          select COUNT(*) 訂購筆數 from (    Select distinct a.BillDate,a.LinkMan 收貨人  From OrdBillMain A     Where A.Flag=2    AND a.CustomerID =@CustomerID   group by a.BillDate,a.LinkMan ) as a    ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));

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
        private System.Data.DataTable GetRANKVIP(string CUSTNAME,string LINKMAN)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CUSTNAME FROM GB_MEMBER WHERE CUSTNAME=@CUSTNAME OR CUSTNAME=@LINKMAN");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CUSTNAME", CUSTNAME));
            command.Parameters.Add(new SqlParameter("@LINKMAN", LINKMAN));
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
        private System.Data.DataTable GetRANK1(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
        
            sb.Append("     SELECT COUNT(*) FROM (");
            sb.Append("     Select  SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BillDate  From OrdBillMain A     Where A.Flag=2    ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("      AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-6,GETDATE()),112) AND CONVERT(varchar(6),DATEADD(M,-1,GETDATE()),112)");
            sb.Append("     ) AS F HAVING COUNT(*) >=6");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK12(string CustomerID, string FTYPE, string LinkMan,int NUM)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT COUNT(*) FROM (");
            sb.Append("     Select  SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BillDate  From OrdBillMain A     Where A.Flag=2    ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }

            sb.Append("     ) AS F HAVING COUNT(*) >=@NUM");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
            command.Parameters.Add(new SqlParameter("@NUM", NUM));
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
        private System.Data.DataTable GetRANK2(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2 ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-6,GETDATE()),112)");
            sb.Append("   AND CONVERT(varchar(6),DATEADD(M,-3,GETDATE()),112)  ) AS A HAVING COUNT(*) >=1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK3(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2    ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-2,GETDATE()),112)");
            sb.Append("   AND CONVERT(varchar(6),DATEADD(M,-1,GETDATE()),112)  ) AS A HAVING COUNT(*) >=1");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK4(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2    ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-12,GETDATE()),112)");
            sb.Append("   AND CONVERT(varchar(6),GETDATE(),112)  ) AS A HAVING COUNT(*) >=2");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK41(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2   ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-2,GETDATE()),112)");
            sb.Append("   AND CONVERT(varchar(6),GETDATE(),112)  ) AS A HAVING COUNT(*) >=0");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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

        private System.Data.DataTable GetRANK6(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2    ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) BETWEEN CONVERT(varchar(6),DATEADD(M,-2,GETDATE()),112)");
            sb.Append("   AND CONVERT(varchar(6),GETDATE(),112) ) AS A HAVING COUNT(*) >=1");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK7(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2     ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  AND SUBSTRING(CAST(A.BillDate AS VARCHAR),1,6) < CONVERT(varchar(6),DATEADD(M,-2,GETDATE()),112)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable GetRANK8(string CustomerID, string FTYPE, string LinkMan)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT COUNT(*)  FROM (");
            sb.Append(" Select  a.BillNO,A.BillDate  From OrdBillMain A     Where A.Flag=2     ");
            if (FTYPE == "P")
            {
                sb.Append("   AND a.LinkMan =@LinkMan ");
            }
            else
            {
                sb.Append("     AND a.CustomerID =@CustomerID ");
            }
            sb.Append("  ) AS A HAVING COUNT(*) =1");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
            command.Parameters.Add(new SqlParameter("@LinkMan", LinkMan));
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
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("客戶等級", typeof(string));
            dt.Columns.Add("業務", typeof(string));
            dt.Columns.Add("客戶類別", typeof(string));
            dt.Columns.Add("來源", typeof(string));
            dt.Columns.Add("簡稱", typeof(string));
            dt.Columns.Add("訂購人電話", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("收貨人地址", typeof(string));
            dt.Columns.Add("訂購日期", typeof(string));
            dt.Columns.Add("取貨日期", typeof(string));
            dt.Columns.Add("訂購品項", typeof(string));
            dt.Columns.Add("訂購金額", typeof(decimal));
            dt.Columns.Add("優惠碼", typeof(string));
            dt.Columns.Add("第一次訂購日期", typeof(string));
            dt.Columns.Add("訂購筆數", typeof(string));
            dt.Columns.Add("多平台", typeof(int));
            
            												

            return dt;
        }

        private System.Data.DataTable MakeTable2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("客戶類別", typeof(string));
            dt.Columns.Add("收貨人", typeof(string));
            dt.Columns.Add("收貨人電話", typeof(string));
            dt.Columns.Add("多平台", typeof(int));



            return dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
       
        }

        private void button2_Click(object sender, EventArgs e)
        {
         
            DataRow dr = null;
            System.Data.DataTable dtCost = MakeTable();
            System.Data.DataTable dt = GetMEM();

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                string  BILLNO=dt.Rows[i]["BILLNO"].ToString();
                string SALES = dt.Rows[i]["SALES"].ToString();
                string CUSTNO = dt.Rows[i]["CUSTNO"].ToString();
                string CUSTID = dt.Rows[i]["CUSTID"].ToString();
                string FTYPE = dt.Rows[i]["FTYPE"].ToString();
                string REMAN = dt.Rows[i]["收貨人"].ToString();
                string CUSTTYPE = dt.Rows[i]["客戶類別"].ToString();
                string CUSTNAME = dt.Rows[i]["簡稱"].ToString();
                System.Data.DataTable dtF = GetMEM2(REMAN);
                if (dtF.Rows.Count > 0)
                {
                    dr["多平台"] = dtF.Rows[0][0].ToString();
                }
                dr["客戶類別"] = CUSTTYPE;
                dr["業務"] = SALES;
                dr["客戶等級"] = GetRANK(CUSTID, FTYPE, REMAN, CUSTNAME);
               
                dr["來源"] = dt.Rows[i]["來源"].ToString();
              
                dr["簡稱"] = CUSTNAME;
                string ORDTEL = dt.Rows[i]["訂購人電話"].ToString();
                if (String.IsNullOrEmpty(ORDTEL))
                {
                    System.Data.DataTable G1 = GetTEL(CUSTNAME);
                    if (G1.Rows.Count > 0)
                    {
                        ORDTEL = G1.Rows[0][0].ToString();
                    }
                }
                dr["訂購人電話"] = ORDTEL;
                dr["收貨人"] = REMAN;
                dr["收貨人電話"] = dt.Rows[i]["收貨人電話"].ToString();
                dr["收貨人地址"] = dt.Rows[i]["收貨人地址"].ToString();
                dr["訂購日期"] = dt.Rows[i]["訂購日期"].ToString();
                dr["取貨日期"] = dt.Rows[i]["取貨日期"].ToString();
                System.Data.DataTable T1 = GetMEMS(BILLNO);
                StringBuilder sb = new StringBuilder();
                if (T1.Rows.Count > 0)
                {

                    for (int F = 0; F <= T1.Rows.Count - 1; F++)
                        {

                            DataRow dd = T1.Rows[F];


                            sb.Append(dd["PROD"].ToString() + "/");


                        }

                        sb.Remove(sb.Length - 1, 1);

                        dr["訂購品項"] = sb.ToString();
                }

                dr["訂購金額"] = dt.Rows[i]["訂購金額"].ToString();
                System.Data.DataTable T2 = GetMEMS2(CUSTNO);
                if (T2.Rows.Count > 0)
                {
                    dr["優惠碼"] = T2.Rows[0][0].ToString();

                }
                System.Data.DataTable T3 = GetMEMS3(CUSTID);
                if (T3.Rows.Count > 0)
                {
                    dr["第一次訂購日期"] = T3.Rows[0][0].ToString();

                }
                System.Data.DataTable T4 = GetMEMS4(CUSTID);
                if (T4.Rows.Count > 0)
                {
                    dr["訂購筆數"] = T4.Rows[0][0].ToString();

                }

                dtCost.Rows.Add(dr);



            }
            if (textBox2.Text != "")
            {
                dtCost.DefaultView.RowFilter = " 訂購品項  LIKE  '%" + textBox2.Text + "%' ";

            }
            if (checkBox1.Checked)
            {
                dtCost.DefaultView.RowFilter = " 多平台 > 1 ";

            }
            dataGridView1.DataSource = dtCost;

            System.Data.DataTable TT1 = GetTOTAL3();
            dataGridView2.DataSource = TT1;
        }
        private string GetRANK(string CUSTID, string FTYPE, string LinkMan, string CUSTNAME)
        {

            System.Data.DataTable RANK1 = GetRANK1(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK12 = GetRANK12(CUSTID, FTYPE, LinkMan,40);
            System.Data.DataTable RANK2 = GetRANK2(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK22 = GetRANK12(CUSTID, FTYPE, LinkMan, 20);
            System.Data.DataTable RANK3 = GetRANK3(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK4 = GetRANK4(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK41 = GetRANK41(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK42 = GetRANK12(CUSTID, FTYPE, LinkMan, 10);
            System.Data.DataTable RANKVIP = GetRANKVIP(CUSTNAME, LinkMan);
            System.Data.DataTable RANK6 = GetRANK6(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK7 = GetRANK7(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK8 = GetRANK8(CUSTID, FTYPE, LinkMan);
            System.Data.DataTable RANK9 = GetRANK12(CUSTID, FTYPE, LinkMan, 2);
            string RANK = string.Empty;

            if (RANK1.Rows.Count > 0 || RANK12.Rows.Count > 0 || RANKVIP.Rows.Count >0)
            {
                RANK = "VIP";
            }
            else if ((RANK2.Rows.Count > 0 && RANK3.Rows.Count > 0) || (RANK22.Rows.Count > 0))
            {
                RANK  = "A";
            }
            else if ((RANK4.Rows.Count > 0 && RANK41.Rows.Count == 0) || (RANK42.Rows.Count > 0))
            {
                RANK = "B";
            }
        
            else if (RANK6.Rows.Count > 0 && RANK7.Rows.Count == 0)
            {
                RANK = "NEW";
            }
            else if (RANK8.Rows.Count > 0)
            {
                RANK = "D";
            }
            else if (RANK9.Rows.Count > 0)
            {
                RANK = "C";
            }
            return RANK;
        }

       
        public void AddGB_MEMBER()
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("TRUNCATE TABLE GB_MEMBER ", connection);
            command.CommandType = CommandType.Text;



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
        public void AddGB_MEMBER2(string CUSTNAME, int AMOUNT, string CUSTTYPE, string CUSTTYPE2)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand(" Insert into GB_MEMBER(CUSTNAME,AMOUNT,CUSTTYPE,CUSTTYPE2) values(@CUSTNAME,@AMOUNT,@CUSTTYPE,@CUSTTYPE2)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CUSTNAME", CUSTNAME));
            command.Parameters.Add(new SqlParameter("@AMOUNT", AMOUNT));
            command.Parameters.Add(new SqlParameter("@CUSTTYPE", CUSTTYPE));
            command.Parameters.Add(new SqlParameter("@CUSTTYPE2", CUSTTYPE2));
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

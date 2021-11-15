using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Transactions;
using System.Configuration;

namespace ACME
{
    class GetMenu
    {
        private static string strCnSP = "Data Source=acmesap;Initial Catalog=acmesqlSP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        private static string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql02;Persist Security Info=True;User ID=sapdbo;Password=@rmas";

        public static object[] GetMenuList1DRS()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };
            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD  ";

            SPlLookupDRS dialog = new SPlLookupDRS();



            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetWorker()
        {
            string[] FieldNames = new string[] { "UserCode", "UserName" };

            string[] Captions = new string[] { "英文名", "中文名" };

            string SqlScript = "SELECT UserCode, UserName  FROM ACME_TASK_WORKER  ";

            SPlLookup dialog = new SPlLookup();


            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCar()
        {
            string[] FieldNames = new string[] { "CarName", "Phone", "GPSPhone" };

            string[] Captions = new string[] { "姓名", "電話", "GPS電話" };

            string SqlScript = "SELECT CarName, Phone,GPSPhone  FROM WH_CAR  ";

            SPlLookup dialog = new SPlLookup();

            //  SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetJoy1(string aa)
        {
            string[] FieldNames = new string[] { "port" };

            string[] Captions = new string[] { aa };

            string SqlScript = " select port from Account_Temp7 WHERE PORTTYPE = '" + aa + "' order by ltrim(substring(port,CHARINDEX(',', port)+1,10)),port  ";

            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuList()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD WHERE FROZENFOR <> 'Y'  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetSOLAR()
        {
            string[] FieldNames = new string[] { "PRJCODE", "PRJNAME" };

            string[] Captions = new string[] { "專案代碼", "專案名稱" };

            string SqlScript = "SELECT PRJCODE,PRJNAME FROM OPRJ WHERE U_BU='SOLAR' ";


            MultiValueDialog dialog = new MultiValueDialog();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.KeyFieldName = "PRJCODE";
            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOPOR(string aa)
        {
            string[] FieldNames = new string[] { "CARDCODE", "CARDNAME", "DOCDATE", "DSCRIPTION", "QUANTITY", "PRICE", "TOTAL", "PROJECT", "PRJNAME", "PYMNTGROUP" };

            string[] Captions = new string[] { "客戶編號", "客戶名稱", "採購日期", "採購項目", "採購數量", "採購單價", "採購金額", "專案代碼", "專案名稱", "付款條件" };
            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT T1.CARDCODE,T1.CARDNAME,CONVERT(VARCHAR(8),T1.DOCDATE,112) DOCDATE,DSCRIPTION,CAST(QUANTITY AS VARCHAR) QUANTITY,CAST(PRICE AS VARCHAR) PRICE,CAST(QUANTITY*PRICE AS INT) TOTAL");
            sb.Append("       ,T0.PROJECT,T2.PRJNAME,T3.PYMNTGROUP FROM POR1 T0");
            sb.Append("      LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("      LEFT JOIN OPRJ T2 ON (T0.PROJECT=T2.PRJCODE)");
            sb.Append("      LEFT JOIN OCTG T3 ON (T1.GROUPNUM=T3.GROUPNUM)");
            sb.Append("  WHERE T1.DOCENTRY='" + aa + "'  AND LINESTATUS='O' ");

            SOLARLookup dialog = new SOLARLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] GetGBOITM(string TYPE)
        {
            string[] FieldNames = new string[] { "ITEMCODE", "ITEMNAME", "PRICE", "UNIT" };

            string[] Captions = new string[] { "產品編號", "產品名稱", "價格", "單位" };
            string SqlScript = "";
            if (TYPE == "內部")
            {

                SqlScript = "  SELECT ProdID ITEMCODE,InvoProdName ITEMNAME,SalesPriceA  PRICE,UNIT    FROM comProduct WHERE ClassID IN ('ARG100','ARP200','ARC200','ARS200','ARS210') AND UDef1 <= Convert(varchar(8),GETDATE(),112) AND CASE ISNULL(UDef2,'') WHEN '' THEN '30140101' ELSE UDef2 END    >= Convert(varchar(8),GETDATE(),112)   ";
            }
            if (TYPE == "外部")
            {
                SqlScript = "  SELECT ProdID ITEMCODE,InvoProdName ITEMNAME,SuggestPrice  PRICE,UNIT    FROM comProduct WHERE ClassID IN ('ARG100','ARP200','ARC200','ARS200','ARS210') AND UDef1 <= Convert(varchar(8),GETDATE(),112) AND CASE ISNULL(UDef2,'') WHEN '' THEN '30140101' ELSE UDef2 END    >= Convert(varchar(8),GETDATE(),112)  AND ProdID IN (SELECT ProdID  FROM stkEnSaleSub where  EnSalePrice  <> 0  ) ";

            }
            if (TYPE == "大宗")
            {
                SqlScript = "  SELECT ProdID ITEMCODE,ProdName  ITEMNAME,SuggestPrice  PRICE,UNIT    FROM comProduct WHERE  SUBSTRING(ClassID,1,2) ='AW' AND UDef1 <= Convert(varchar(8),GETDATE(),112) AND CASE ISNULL(UDef2,'') WHEN '' THEN '30140101' ELSE UDef2 END    >= Convert(varchar(8),GETDATE(),112)   ";

            }
            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetGBOITMO(string aa)
        {
            string[] FieldNames = new string[] { "ITEMCODE", "ITEMNAME", "PRICE" };

            string[] Captions = new string[] { "產品編號", "產品名稱", "價格" };

            string SqlScript = "SELECT  ITEMCODE,ITEMNAME,PRICE  FROM GB_OITM WHERE ITEMOI ='" + aa + "' AND ISNULL(ACTIVE,'') <> 'N' ORDER BY ITEMCODE  ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static string GetWhMainDelivery(string ShippingCode)
        {
            //分批放貨時從備貨單選擇分批放貨轉放貨單,放貨單匯出文件在這邊選擇匯出哪一批文件
            string[] FieldNames = new string[] { "SeqDelivery", "ItemQty" };

            string[] Captions = new string[] { "排序", "料號數量 " };

            string SqlScript = @"SELECT SeqDelivery , (SELECT ItemCode + '-' + Quantity+'pcs ' 
FROM WH_ITEM T2
 WHERE T2.SeqDelivery = T1.SeqDelivery  for xml path('')
) ItemQty FROM WH_ITEM T1 WHERE shippingcode = '{0}'
GROUP BY SEQDELIVERY
ORDER by SeqDelivery";

            SqlScript = string.Format(SqlScript, ShippingCode);
            SOLARLookup dialog = new SOLARLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            dialog.LookUpConnection = new SqlConnection(globals.ConnectionString); 
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues[0].ToString();

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHIPCARD()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " SELECT ID,FULLNAME FROM comCustomer WHERE FLAG=1 ";


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetCHIPCARD2()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " SELECT ID,FULLNAME FROM comCustomer WHERE FLAG=2  ";


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHI()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer  ";


            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUSTGB()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = "   SELECT  ID,FullName FROM comCustomer  ";


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetCHICUSTGB2(string AA)
        {
            string[] FieldNames = new string[] { "PRODID", "PRODNAME" };

            string[] Captions = new string[] { "料號", "名稱" };

            StringBuilder sb = new StringBuilder();
            sb.Append("                SELECT PRODID,PRODNAME FROM (");
            sb.Append("                SELECT  J.PRODID,J.PRODNAME ,   CASE ");
            sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='S' AND SUBSTRING(K.ClassID,1,1)='A'  THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'      ");
            sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬'  WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '加工品' END 品項");
            sb.Append("                                                                                      FROM comProduct J ");
            sb.Append("                   Left Join comProductClass K On J.ClassID =K.ClassID   ");
            sb.Append("              WHERE K.CLASSID IN ('ARP200','ARC200','ARS200','ARS210','BPK010','BPK020','BPK030')");
            sb.Append("                      ) AS A WHERE ISNULL(品項,'') ='" + AA + "' ");

            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetCHICUSTGB3(string CustomerID)
        {
            string[] FieldNames = new string[] { "BillNO", "BillDate", "LinkMan", "CustAddress" };

            string[] Captions = new string[] { "訂單編號", "訂購日期", "到貨人", "地址" };

            StringBuilder sb = new StringBuilder();
            sb.Append("                     select BillNO,BillDate,LinkMan,CustAddress from OrdBillMain  WHERE Flag =2 AND CustomerID='" + CustomerID + "' ");


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetCHICUSTGB4(string CustomerID)
        {
            string[] FieldNames = new string[] { "BillNO", "BillDate", "LinkMan", "CustAddress" };

            string[] Captions = new string[] { "訂單編號", "訂購日期", "到貨人", "地址" };

            StringBuilder sb = new StringBuilder();

            sb.Append("                      select DISTINCT   T0.BillNO,T0.BillDate,T0.LinkMan,CustAddress from OrdBillMain T0  ");
            sb.Append("                               LEFT JOIN comCustAddress T1 ON (T0.AddressID=T1.AddrID AND T0.CustomerID=T1.ID )  ");
            sb.Append("                              where  REPLACE(REPLACE(REPLACE(REPLACE(T1.Telephone,'(',''),')',''),'&',''),'-','') LIKE '" + CustomerID + "'  ");

            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=1  ";


            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUSTM()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=2  ";


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUSTM2()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=1  ";


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST12()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=1 ";


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST13()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=1 ";


            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST14()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer where FLAG=1 ";


            CHOLookup5 dialog = new CHOLookup5();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST2GB()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "供應商代碼", "供應商名稱" };

            string SqlScript = "select ID,FullName from comCustomer where FLAG=2";


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST2()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "供應商代碼", "供應商名稱" };

            string SqlScript = "select ID,FullName from comCustomer where FLAG=2";


            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST222()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "供應商代碼", "供應商名稱" };

            string SqlScript = "select ID,FullName from comCustomer where FLAG=2";


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST223()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "供應商代碼", "供應商名稱" };

            string SqlScript = "select ID,FullName from comCustomer where FLAG=2";


            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHICUST224()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "供應商代碼", "供應商名稱" };

            string SqlScript = "select ID,FullName from comCustomer where FLAG=2";


            CHOLookup5 dialog = new CHOLookup5();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHI2()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer  ";


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHI4()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer  ";


            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetCHI5()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer   ";


            CHOLookup5 dialog = new CHOLookup5();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetCHI6()
        {
            string[] FieldNames = new string[] { "ID", "FullName" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱" };

            string SqlScript = " select ID,FullName from comCustomer  ";


            CHOLookup6 dialog = new CHOLookup6();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] JOJO()
        {
            string[] FieldNames = new string[] { "ITEMCODE", "ITEMNAME" };

            string[] Captions = new string[] { "產品編號", "產品名稱" };

            string SqlScript = "SELECT ITEMCODE ,ITEMNAME FROM OITM T0 INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T0.ItemCode where AND  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組'   ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuListtest()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD  ";


            MultiValueDialog dialog = new MultiValueDialog();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetMenuListC()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD where cardtype='C'  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetMenuListS()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD where substring(cardcode,1,1) IN ('S','U')  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuListU2()
        {
            string[] FieldNames = new string[] { "CARDCODE", "CARDNAME", "BANKNAME", "ACCOUNT", "BANKCODE", "PYMNTGROUP" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "銀行名稱", "銀行帳號", "銀行代碼", "付款條件" };

            string SqlScript = "SELECT T0.CARDCODE,T0.CARDNAME,BANKNAME+DFLBRANCH BANKNAME,DFLACCOUNT ACCOUNT,T0.BANKCODE,PYMNTGROUP FROM OCRD  T0 LEFT JOIN ODSC T1 ON (T0.BANKCODE =T1.BANKCODE)  LEFT JOIN OCTG T2 ON (T0.GROUPNUM =T2.GROUPNUM)   where substring(cardcode,1,1) IN ('S','U')   ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuListU()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax", "address", "mailaddres" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "billto", "shipto" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD where SUBSTRING(Cardcode,1,1) in ('U') UNION ALL SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax,address,mailaddres FROM OCRD where SUBSTRING(Cardcode,1,1) in ('S') ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOitmz()
        {
            string[] FieldNames = new string[] { "itemcode", "itemname" };

            string[] Captions = new string[] { "產品編號", "產品名稱" };

            string SqlScript = "select itemcode,itemname from oitm where substring(itemcode,1,1)='Z' ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOslp()
        {
            string[] FieldNames = new string[] { "slpname", "memo" };

            string[] Captions = new string[] { "業務名稱", "部門" };

            string SqlScript = "select slpname,memo from oslp ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOitm()
        {
            string[] FieldNames = new string[] { "itemcode", "itemname" };

            string[] Captions = new string[] { "產品編號", "產品名稱" };

            string SqlScript = "select itemcode,itemname from oitm ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOitmGB()
        {
            string[] FieldNames = new string[] { "ProdID", "ProdName" };

            string[] Captions = new string[] { "產品編號", "產品名稱" };

            string SqlScript = "SELECT ProdID,ProdName FROM comProduct  WHERE  SUBSTRING(ProdID,1,2) IN ('MC','MP')   or SUBSTRING(ProdID,1,3) IN ('MSR') ";


            CHOLookup3 dialog = new CHOLookup3();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetJobNO()
        {
            string[] FieldNames = new string[] { "shippingcode", "tradeCondition" };

            string[] Captions = new string[] { "JOBNO", "貿易條件" };

            string SqlScript = "select shippingcode,tradeCondition from shipping_main ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOcrd()
        {
            string[] FieldNames = new string[] { "cardcode", "cardname" };

            string[] Captions = new string[] { "客戶編號", "客戶名稱" };

            string SqlScript = "select cast(docentry as varchar),cardcode,cardname from ocrd ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetMenuListS1()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax FROM OCRD where cardtype='C' ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetU()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "Cntctprsn", "Phone1", "Fax" };

            string[] Captions = new string[] { "廠商代碼", "廠商名稱", "聯絡人", "電話", "傳真" };

            string SqlScript = "SELECT Cardcode, CardName, Cntctprsn,Phone1,Fax FROM OCRD where substring(cardcode,1,1) IN ('U') ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetFMD()
        {
            string[] FieldNames = new string[] { "TRANSID", "MEMO", "usersign" };

            string[] Captions = new string[] { "傳票號碼", "備註", "使用者" };

            string SqlScript = "SELECT CAST(TRANSID AS VARCHAR) TRANSID,MEMO,usersign usersign FROM OJDT WHERE TRANSID NOT IN (SELECT U_BSREN FROM [@CADMEN_FMD]) ORDER BY cast(TRANSID as int) DESC ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] GETUIN()
        {
            string[] FieldNames = new string[] { "CARDCODE", "CARDNAME" };

            string[] Captions = new string[] { "廠商代碼", "廠商名稱" };

            string SqlScript = "SELECT CARDCODE,CARDNAME FROM Shipping_OQUT4 WHERE ISNULL(CARDCODE,'') <> ''  ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GETOUITEM()
        {
            string[] FieldNames = new string[] { "ITEMCODE", "ITEMNAME" };

            string[] Captions = new string[] { "項目編號", "項目名稱" };

            string SqlScript = "SELECT ITEMCODE,ITEMNAME FROM  Shipping_OQUT5 WHERE ISNULL(ITEMCODE,'') <> ''  ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GETOUITEM2(string aa)
        {
            string[] FieldNames = new string[] { "ITEMCODE", "ITEMNAME" };

            string[] Captions = new string[] { "項目編號", "項目名稱" };

            string SqlScript = "SELECT ITEMCODE,ITEMNAME FROM  OITM WHERE SUBSTRING(ITEMCODE,1,2)='ZA'  AND SUBSTRING(ITEMNAME,1,1)='" + aa + "'  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] RmaCardcode(string cardname)
        {
            string[] FieldNames = new string[] { "CustCode", "CustName", "ConctPrsn", "Tel", "fax", "BillTo", "shiplTo" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "BillTo", "shiplTo" };

            string SqlScript = "SELECT CustCode,CustName,ConctPrsn,Tel,fax,BillTo,shiplTo  FROM Rma_Cardcode where cast(custname as varchar) like  '%" + cardname + "%'  ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] RmrCardcode()
        {
            string[] FieldNames = new string[] { "CustCode", "CustName", "ConctPrsn", "Tel", "fax", "BillTo", "shiplTo" };

            string[] Captions = new string[] { "客戶代碼", "客戶名稱", "聯絡人", "電話", "傳真", "BillTo", "shiplTo" };

            string SqlScript = "SELECT CustCode,CustName,ConctPrsn,Tel,fax,BillTo,shiplTo  FROM Rma_Cardcode ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] RmrRRS()
        {
            string[] FieldNames = new string[] { "PARAM_NO", "PARAM_DESC" };

            string[] Captions = new string[] { "起運地", "起運地" };

            string SqlScript = "SELECT PARAM_NO,PARAM_DESC FROM PARAMS WHERE PARAM_KIND='RRS' ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] RmrRRSH()
        {
            string[] FieldNames = new string[] { "PARAM_NO", "PARAM_DESC" };

            string[] Captions = new string[] { "SHIPPER", "SHIPPER" };

            string SqlScript = "SELECT PARAM_NO,PARAM_DESC FROM PARAMS WHERE PARAM_KIND='RRSH' ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] RmrRRM()
        {
            string[] FieldNames = new string[] { "PARAM_NO", "PARAM_DESC" };

            string[] Captions = new string[] { "備註", "備註" };

            string SqlScript = "SELECT PARAM_NO,PARAM_DESC FROM PARAMS WHERE PARAM_KIND='RRM' ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] RmrRRT()
        {
            string[] FieldNames = new string[] { "PARAM_NO", "PARAM_DESC" };

            string[] Captions = new string[] { "貨運", "貨運" };

            string SqlScript = "SELECT PARAM_NO,PARAM_DESC FROM PARAMS WHERE PARAM_KIND='RRT' ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] RmrCONS()
        {
            string[] FieldNames = new string[] { "PARAM_NO", "PARAM_DESC" };

            string[] Captions = new string[] { "廠區", "廠區" };

            string SqlScript = "SELECT PARAM_NO,PARAM_DESC FROM PARAMS WHERE PARAM_KIND='CONS' ";


            SPlLookup dialog = new SPlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] Getowhs()
        {
            string[] FieldNames = new string[] { "whscode", "whsname" };

            string[] Captions = new string[] { "倉庫代碼", "倉庫名稱" };

            string SqlScript = "SELECT whscode,whsname FROM owhs where whsname is not null order by substring(whscode,0,2) desc  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] Getohem()
        {
            string[] FieldNames = new string[] { "name", "jobtitle" };

            string[] Captions = new string[] { "名稱", "職稱" };

            string SqlScript = "SELECT  t0.lastname+t0.firstname [name],t1.[name] jobtitle from ohem t0 left join oudp t1 on (t0.dept=t1.code)  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] Getowtr1(string aa, string bb, string cc)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "AR發票")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111)  as  過帳日期  FROM oinv T0 where T0.cardcode='" + aa + "' ");


            }
            else if (bb == "銷售訂單")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM ORDR T0 inner join (select distinct docentry from rdr1 where 1=1  ");
                if (cc == "b")
                {
                    sb.Append("    and linestatus='O' ");
                }
                sb.Append(" ) T1 on (t0.docentry=t1.docentry) where T0.cardcode='" + aa + "' order by T0.docentry desc ");


            }
            else if (bb == "庫存調撥-借出")
            {
                sb.Append("SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='1' and T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "發貨單")
            {
                sb.Append("SELECT DISTINCT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期,T0.DOCENTRY FROM Oige T0 LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE ISNULL(T1.BASEREF,'') = '' ORDER BY T0.DOCENTRY DESC   ");

            }
            else if (bb == "庫存調撥-撥倉")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='3'  ");

            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Opor T0 inner join (select distinct docentry from por1 where 1=1  ");

                if (cc == "b")
                {
                    sb.Append("    and linestatus='O' ");
                }
                sb.Append(" ) T1 on (t0.docentry=t1.docentry) where T0.cardcode='" + aa + "' ");
            }
            else if (bb == "採購報價")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM OPQT T0 inner join (select distinct docentry from PQT1 where 1=1  ");

                if (cc == "b")
                {
                    sb.Append("    and linestatus='O' ");
                }
                sb.Append(" ) T1 on (t0.docentry=t1.docentry) where T0.cardcode='" + aa + "' ");
            }
            else if (bb == "收貨採購單")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Opdn T0 where  T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "採購退貨")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM ORPD T0 where  T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "AP貸項")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM ORPC T0 where  T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "庫存調撥-借出還回")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Owtr T0 where t0.u_acme_kind='2' and T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "AR貸項通知單")
            {
                sb.Append(" SELECT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期 FROM Orin T0 where  T0.cardcode='" + aa + "'  ");

            }
            else if (bb == "收貨單")
            {
                sb.Append(" SELECT DISTINCT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期,T0.DOCENTRY FROM OIGN T0 LEFT JOIN IGN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE ISNULL(T1.BASEREF,'') = '' ORDER BY T0.DOCENTRY DESC   ");

            }
            else if (bb == "生產訂單")
            {
                sb.Append(" SELECT cast(T0.DOCENTRY as varchar) as 單號,Convert(varchar(10),t0.POSTdate,111) as  過帳日期 FROM OWOR T0 where  T0.cardcode='" + aa + "' ORDER BY DOCENTRY DESC   ");

            }
            else if (bb == "生產發貨")
            {
                sb.Append(" SELECT DISTINCT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期,T0.DOCENTRY FROM Oige T0  LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) LEFT JOIN OWOR T2 ON (T1.BASEREF=T2.DOCENTRY) WHERE ISNULL(T1.BASEREF,'') <> '' AND  T2.cardcode='" + aa + "' ORDER BY T0.DOCENTRY DESC   ");

            }
            else if (bb == "生產收貨")
            {
                sb.Append(" SELECT DISTINCT cast(T0.Docnum as varchar) as 單號,Convert(varchar(10),t0.docdate,111) as  過帳日期,T0.DOCENTRY FROM Oige T0  LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) LEFT JOIN OWOR T2 ON (T1.BASEREF=T2.DOCENTRY) WHERE ISNULL(T1.BASEREF,'') <> '' AND  T2.cardcode='" + aa + "' ORDER BY T0.DOCENTRY DESC   ");

            }
            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrCHO(string aa, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrTAO(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT T0.FundBillNo 單號,T0.BILLDATE 過帳日期 FROM comBillAccounts T0 WHERE T0.Flag=600  AND T0.CustID ='" + aa + "'");

            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetBOWCHO(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT T0.BorrowNO  單號,T0.BorrowDATE 過帳日期 FROM stkBorrowMain T0 WHERE T0.CUSTOMERID ='" + aa + "' ");
            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrDIAO()
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MoveNO 單號,MoveDate 過帳日期 FROM stkMoveMAIN ORDER BY MOVENO DESC");

            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHO2(string aa, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrTAO2(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT T0.FundBillNo 單號,T0.BILLDATE 過帳日期 FROM comBillAccounts T0 WHERE T0.Flag=600  AND T0.CustID ='" + aa + "'");


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetBOWINF(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT T0.BorrowNO  單號,T0.BorrowDATE 過帳日期 FROM stkBorrowMain T0 WHERE T0.CUSTOMERID ='" + aa + "' ");

            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrDIAO2()
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MoveNO 單號,MoveDate 過帳日期 FROM stkMoveMAIN ORDER BY MOVENO DESC");


            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrDIAOT()
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MoveNO 單號,MoveDate 過帳日期 FROM stkMoveMAIN ORDER BY MOVENO DESC ");

            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHO2T(string aa, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetBOWTOP(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT T0.BorrowNO  單號,T0.BorrowDATE 過帳日期 FROM stkBorrowMain T0 WHERE T0.CUSTOMERID ='" + aa + "' ");

            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrTAO3(string aa)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT T0.FundBillNo 單號,T0.BILLDATE 過帳日期 FROM comBillAccounts T0 WHERE T0.Flag=600  AND T0.CustID ='" + aa + "'");


            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHOAD(string aa, string STATUS, string FLAG)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG='" + FLAG + "' AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup5 dialog = new CHOLookup5();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrCHOAT(string aa, string STATUS, string FLAG)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG='" + FLAG + "' AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup6 dialog = new CHOLookup6();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHOCHO(string aa, string STATUS, string FLAG)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG='" + FLAG + "' AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrCHOCHO2(string aa, string STATUS, string FLAG)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG='" + FLAG + "' AND CUSTOMERID='" + aa + "' ");
            if (STATUS == "0")
            {
                sb.Append(" AND BillStatus = 0 ");

            }

            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHO1(string aa, string bb, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "銷售單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=4 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "銷貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=500 AND custid='" + aa + "' ");

            }
            else if (bb == "進貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=100 AND custid='" + aa + "' ");

            }
            else if (bb == "銷退單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=600 AND custid='" + aa + "' ");

            }
            else if (bb == "調撥單")
            {
                sb.Append(" SELECT MoveNO  單號,MOVEdate 過帳日期 FROM StkMoveMain WHERE FLAG=400 ORDER BY MoveNO DESC  ");

            }
            else if (bb == "借出單")
            {
                sb.Append(" SELECT BorrowNO  單號,BorrowDate 過帳日期 FROM stkBorrowMain WHERE FLAG=10 AND CUSTOMERID='" + aa + "'  ");

            }
            else if (bb == "還回單")
            {
                sb.Append(" SELECT ReturnNO  單號,ReturnDate 過帳日期 FROM stkReturnMain WHERE FLAG=12 AND CUSTOMERID='" + aa + "'  ");

            }
            CHOLookup dialog = new CHOLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrCHO2(string aa, string bb, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "銷售單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=4 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "銷貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=500 AND custid='" + aa + "' ");

            }
            else if (bb == "進貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=100 AND custid='" + aa + "' ");

            }
            else if (bb == "銷退單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=600 AND custid='" + aa + "' ");

            }
            else if (bb == "調撥單")
            {
                sb.Append(" SELECT MoveNO  單號,MOVEdate 過帳日期 FROM StkMoveMain WHERE FLAG=400 ORDER BY MoveNO DESC  ");

            }
            else if (bb == "借出單")
            {
                sb.Append(" SELECT BorrowNO  單號,BorrowDate 過帳日期 FROM stkBorrowMain WHERE FLAG=10 AND CUSTOMERID='" + aa + "'  ");

            }
            else if (bb == "還回單")
            {
                sb.Append(" SELECT ReturnNO  單號,ReturnDate 過帳日期 FROM stkReturnMain WHERE FLAG=12 AND CUSTOMERID='" + aa + "'  ");

            }
            CHOLookup2 dialog = new CHOLookup2();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHO4(string aa, string bb, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "銷售單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=4 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "銷貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=500 AND custid='" + aa + "' ");

            }
            else if (bb == "進貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=100 AND custid='" + aa + "' ");

            }
            else if (bb == "銷退單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=600 AND custid='" + aa + "' ");

            }
            else if (bb == "調撥單")
            {
                sb.Append(" SELECT MoveNO  單號,MOVEdate 過帳日期 FROM StkMoveMain WHERE FLAG=400 ORDER BY MoveNO DESC  ");

            }
            else if (bb == "借出單")
            {
                sb.Append(" SELECT BorrowNO  單號,BorrowDate 過帳日期 FROM stkBorrowMain WHERE FLAG=10 AND CUSTOMERID='" + aa + "'  ");

            }
            else if (bb == "還回單")
            {
                sb.Append(" SELECT ReturnNO  單號,ReturnDate 過帳日期 FROM stkReturnMain WHERE FLAG=12 AND CUSTOMERID='" + aa + "'  ");

            }
            CHOLookup4 dialog = new CHOLookup4();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetowtrCHO6(string aa, string bb, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "銷售單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=4 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "銷貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=500 AND custid='" + aa + "' ");

            }
            else if (bb == "進貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=100 AND custid='" + aa + "' ");

            }
            else if (bb == "銷退單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=600 AND custid='" + aa + "' ");

            }
            else if (bb == "調撥單")
            {
                sb.Append(" SELECT MoveNO  單號,MOVEdate 過帳日期 FROM StkMoveMain WHERE FLAG=400 ORDER BY MoveNO DESC  ");

            }
            else if (bb == "借出單")
            {
                sb.Append(" SELECT BorrowNO  單號,BorrowDate 過帳日期 FROM stkBorrowMain WHERE FLAG=10 AND CUSTOMERID='" + aa + "'  ");

            }
            else if (bb == "還回單")
            {
                sb.Append(" SELECT ReturnNO  單號,ReturnDate 過帳日期 FROM stkReturnMain WHERE FLAG=12 AND CUSTOMERID='" + aa + "'  ");

            }
            CHOLookup6 dialog = new CHOLookup6();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetowtrCHO5(string aa, string bb, string STATUS)
        {
            string[] FieldNames = new string[] { "單號", "過帳日期" };

            string[] Captions = new string[] { "單號", "過帳日期" };
            StringBuilder sb = new StringBuilder();
            if (bb == "銷售單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=2 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "採購單")
            {
                sb.Append(" SELECT BillNO 單號,BILLDATE 過帳日期 FROM OrdBillMain WHERE FLAG=4 AND CUSTOMERID='" + aa + "' ");
                if (STATUS == "b")
                {
                    sb.Append(" AND BillStatus = 0 ");

                }
            }
            else if (bb == "銷貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=500 AND custid='" + aa + "' ");

            }
            else if (bb == "進貨單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=100 AND custid='" + aa + "' ");

            }
            else if (bb == "銷退單")
            {
                sb.Append(" SELECT FundBillNo  單號,billdate 過帳日期 FROM comBillAccounts WHERE FLAG=600 AND custid='" + aa + "' ");

            }
            else if (bb == "調撥單")
            {
                sb.Append(" SELECT MoveNO  單號,MOVEdate 過帳日期 FROM StkMoveMain WHERE FLAG=400 ORDER BY MoveNO DESC  ");

            }
            else if (bb == "借出單")
            {
                sb.Append(" SELECT BorrowNO  單號,BorrowDate 過帳日期 FROM stkBorrowMain WHERE FLAG=10 AND CUSTOMERID='" + aa + "'  ");

            }
            else if (bb == "還回單")
            {
                sb.Append(" SELECT ReturnNO  單號,ReturnDate 過帳日期 FROM stkReturnMain WHERE FLAG=12 AND CUSTOMERID='" + aa + "'  ");

            }
            CHOLookup5 dialog = new CHOLookup5();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = sb.ToString();

            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] Get0itm()
        {
            string[] FieldNames = new string[] { "itemcode", "itemname" };

            string[] Captions = new string[] { "產品編號", "品名規格" };

            string SqlScript = "SELECT itemcode,itemname FROM oitm  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }


        public static object[] GetMenuListSu()
        {
            string[] FieldNames = new string[] { "Cardcode", "CardName", "CardFName", "GroupCode", "Currency", "LicTradNum" };

            string[] Captions = new string[] { "客戶代碼", "名稱", "外文名稱", "群組", "供應商", "統編" };

            string SqlScript = "SELECT Cardcode,CardName,CardFName,GroupCode,Currency,LicTradNum FROM OCRD";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static object[] GetMenuInvo()
        {
            string[] FieldNames = new string[] { "InvoiceSeq", "WHInvoice", "InvoicePlace", "Memo", "StratDay", "EndDay" };

            string[] Captions = new string[] { "代號", "倉庫", "地點", "備註", "啟用日", "結束日" };

            string SqlScript = "SELECT * FROM invoiceseq";


            inLookup dialog = new inLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }





        public static DataTable GetOWOR(string originnum)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.DOCENTRY,T1.LINENUM,T1.ITEMCODE,T2.ITEMNAME,CAST(T1.PLANNEDQTY AS INT) 數量,CAST(T2.AVGPRICE*1.1 AS DECIMAL(10,4)) 單價,CAST(CAST(T2.AVGPRICE*1.1 AS DECIMAL(10,4))*T1.PLANNEDQTY AS DECIMAL(15,4)) 金額,originnum 銷售單號,Convert(varchar(8),T1.U_ACME_SHIPPDAY,112)  離倉日期,T1.VISORDER   from OWOR T0 ");
            sb.Append(" LEFT JOIN WOR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE) ");
            sb.Append(" where Convert(nvarchar(8),t1.u_acme_work,112)+cast(T0.Docnum as nvarchar)+T1.WAREHOUSE  in (" + originnum + ") ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }


        public static DataTable GetinvoiceM(string shippingcode)
        {

            SqlConnection MyConnection;

            MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM invoiceM WHERE shippingcode=@shippingcode  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static DataTable GetOWOR2(string originnum)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select top 1 t0.Docnum,t1.ItemCode,t1.Dscription,t0.NumAtCard,t1.Quantity,t1.Price,t1.linenum,t1.totalfrgn,t0.u_acme_tardeterm 貿易條件");
            sb.Append(" ,t0.u_beneficiary 最終客戶,T1.U_PAY 付款,T1.U_SHIPDAY 押出貨日,T1.U_SHIPSTATUS 貨況,T1.U_MARK 特殊嘜頭,T1.U_MEMO 注意事項,cast(u_acme_forwarder as nvarchar(max))  FORWARDER,u_acme_byair 運輸方式,t0.u_acme_shipform1 shipform,t0.u_acme_shipto1 shipto,T0.U_ACME_PAY 付款方式  from rdr1 t1");
            sb.Append(" left join ordr t0 on (t1.docentry=t0.docentry) ");

            sb.Append(" where t0.docentry= '" + originnum + "' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable GetO(string aa)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select cardcode from ocrd where cardname=@aa");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", aa));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable Gettt(string DocEntry, string TYPE)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.CURRENCY, t1.doctype 類型,t0.LineNum LineNum,  ");
            sb.Append(" SUBSTRING(t0.dscription,0,20) 描述,t1.docentry,  ");
            sb.Append(" t0.itemcode,t0.quantity,t0.price,  ");
            sb.Append(" case when t0.vatprcnt = 0 then ROUND(T0.totalfrgn,3) else  ROUND(T0.totalfrgn*1.05,3)  end  gtotalfc, ");
            sb.Append(" CASE WHEN  isnull(t0.u_shipprice,0) <> 0 THEN t0.u_shipprice ELSE isnull(t0.rate,0) END rate, ");
            sb.Append(" (case when t0.vatprcnt = 0 then ROUND(T0.totalfrgn,3) else  ROUND(T0.totalfrgn*1.05,3)  end)*(CASE WHEN  isnull(t0.u_shipprice,0) <> 0 THEN t0.u_shipprice ELSE isnull(t0.rate,0) END) gtotal,  ");
            sb.Append(" (CASE T6.ITMSGRPCOD WHEN 1033 THEN case when t0.vatprcnt = 0 then ROUND(T0.LINEtotal,0) else ROUND(T0.LINEtotal*1.05,0) end ELSE ROUND(t0.gtotal,0) END) gtotalC,  ");
            sb.Append(" t0.vatprcnt,Convert(varchar(10),t0.u_acme_work,112) shipdate,  ");
            sb.Append(" numatcard cardcode,case when T1.cardcode  in ('0511-00','0257-00') then u_beneficiary  else T1.cardname end cardname ,case when isnull(t5.docentry,'') ='' then t7.docentry else t5.docentry end oinv,t1.cardcode 客戶編號,T0.PRICEAFVAT,T8.CardFName 英文名稱 from rdr1 t0   ");
            sb.Append(" left join ordr t1 on (t0.docentry=t1.docentry)  ");
            sb.Append(" left join dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum )  ");
            sb.Append(" left join inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and  t5.basetype='15')     ");
            sb.Append(" LEFT JOIN OITM T6 ON (T0.ITEMCODE=T6.ITEMCODE)    ");
            sb.Append(" left join inv1 t7 on (t7.baseentry=T0.docentry and  t7.baseline=t0.linenum and  t7.basetype='17')     ");
            sb.Append(" left join OCRD T8 on (T1.CARDCODE=T8.CARDCODE)  ");
            if (TYPE == "ORDR")
            {
                sb.Append("      where t1.docentry=@DocEntry ");
            }

            if (TYPE == "OINV")
            {
                sb.Append("      where t5.docentry=@DocEntry ");
            }

            if (TYPE == "OINV3")
            {
                sb.Append("      where t7.docentry=@DocEntry ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable GetttCHO(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("          SELECT 'I' 類型,T1.RowNO LineNum,T0.BillNO docentry,T1.ProdID itemcode,T1.Quantity quantity,T1.Price price,ROUND((T1.AMOUNT+T1.TaxAmt),3) gtotalfc,ExchRate rate,round(((T1.AMOUNT+T1.TaxAmt)*ExchRate),0) gtotal");
            sb.Append("                ,CASE T1.TaxAmt WHEN 0 THEN 0 ELSE 5 END vatprcnt,PreInDate shipdate,'' cardcode,FullName cardname,T7.BillNO oinv,T0.CustomerID  客戶編號   FROM OrdBillMain T0");
            sb.Append("                  Inner Join OrdBillSub T1 On T0.Flag=T1.Flag And T0.BillNO=T1.BillNO ");
            sb.Append("                  Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID)");
            sb.Append("                  LEFT JOIN comProdRec T7 ON (T1.BillNO =T7.FromNO AND T1.RowNO=T7.FromRow)");
            sb.Append("                   where t0.Flag =2 AND T0.BillNO=@DocEntry ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static DataTable Getttt(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t1.doctype 類型,SUBSTRING(t0.dscription,0,20) 描述,t1.docentry,itemcode,quantity,price,gtotalfc,rate,gtotal,vatprcnt,Convert(varchar(10),u_acme_work,112) shipdate,numatcard cardcode,cardname from rin1 t0 left join orin t1 on (t0.docentry=t1.docentry)  where t1.docentry=@DocEntry");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }


        public static System.Data.DataTable GetWHPACK4(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select MAX(CAST(PLATENO AS INT)) PLATENO, SUM(CAST(CARTONNO  AS INT)) CARTONNO,");
            sb.Append(" MAX(CAST(CARTONNO  AS INT))+ MAX(CAST(PLATENO AS INT))  TOTAL");
            sb.Append(" from WH_PACK2  WHERE  SHIPPINGCODE =@SHIPPINGCODE ");
            sb.Append(" GROUP BY FLAG1");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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
        public static System.Data.DataTable GetWHPACKH2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT PACKMEMO  FROM WH_MAIN WHERE SHIPPINGCODE IN (" + SHIPPINGCODE + "  ) AND PACKMEMO <> ''  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        public static System.Data.DataTable GetWHMARK(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT PACKMEMO  FROM WH_MAIN WHERE SHIPPINGCODE =@SHIPPINGCODE AND PACKMEMO <> '' ");
            sb.Append("  UNION ALL ");
            sb.Append("  SELECT PACKMEMO  FROM AcmeSqlSPCHOICE.DBO.WH_MAIN WHERE SHIPPINGCODE =@SHIPPINGCODE AND PACKMEMO <> '' ");
            sb.Append("   UNION ALL ");
            sb.Append("  SELECT PACKMEMO  FROM AcmeSqlSPINFINITE.DBO.WH_MAIN WHERE SHIPPINGCODE =@SHIPPINGCODE AND PACKMEMO <> '' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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
        public static DataTable GetDOWNLOAD2(string shippingcode, string seq)
        {

            SqlConnection MyConnection;

            MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select filename,path FROM DOWNLOAD2 WHERE shippingcode=@shippingcode AND seq=@seq");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable KingsMenuTble(string DocEntry, string ca)
        {

            SqlConnection MyConnection;
            if (ca == "Checked")
            {
                MyConnection = globals.lpConnection;

            }
            else
            {
                MyConnection = globals.shipConnection;
            }
            string sql = "select * from rdr1 left join ordr on (rdr1.docentry=ordr.docentry) where ordr.Docnum=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static DataTable GetPI(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct(docentry) docentry  from shipping_item where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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

        public static DataTable GetSAME(string forecastDay, string receivePlace, string GOALPLACE, string tradeCondition, string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT T0.shippingcode JOBNO,SUBSTRING(CARDCODE,7,3) BU,T1.WHSCODE 倉庫 FROM SHIPPING_MAIN T0");
            sb.Append(" LEFT JOIN LcInstro T1 ON (T0.shippingcode=T1.SHIPPINGCODE)");
            sb.Append("  WHERE forecastDay=@forecastDay AND receivePlace=@receivePlace AND GOALPLACE=@GOALPLACE AND tradeCondition=@tradeCondition and T0.shippingcode <> @shippingcode AND ");
            sb.Append(" substring(cardcode,1,5)='S0001' and receiveDay='sea'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@forecastDay", forecastDay));
            command.Parameters.Add(new SqlParameter("@receivePlace", receivePlace));
            command.Parameters.Add(new SqlParameter("@GOALPLACE", GOALPLACE));
            command.Parameters.Add(new SqlParameter("@tradeCondition", tradeCondition));
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

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


        public static DataTable GetPO(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select distinct(pino) pino from shipping_item where shippingcode=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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
        public static DataTable GetOwhs2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from owhs where docentry=@docentry ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["owhs"];
        }
        public static DataTable GetOwhb2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from owhb where docentry=@docentry ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "owhr");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["owhr"];
        }
        public static DataTable Owhs(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from owhr where qty='0' and docentry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "owhr");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["owhr"];
        }
        public static DataTable Owhb2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from owhb where qty='0' and docentry=@DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "owhb");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["owhb"];
        }
        public static DataTable GetOrdr2(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select Docnum,ItemCode,Dscription,Price,T1.linenum,totalfrgn,T1.QUANTITY-ISNULL(T2.QUANTITY,0) QTY,VISORDER from por1 T1 left join opor T0 on (T1.docentry=T0.docentry)");
            sb.Append(" left join  (SELECT DOCENTRY,LINENUM,SUM(QUANTITY) QUANTITY FROM ACMESQLSP.DBO.SHIPPING_ITEM WHERE ITEMREMARK='採購訂單'");
            sb.Append(" GROUP BY DOCENTRY,LINENUM) T2 ON (CAST(T1.DOCENTRY AS VARCHAR)=CAST(T2.DOCENTRY AS VARCHAR)  AND T1.LINENUM=T2.LINENUM)");
            sb.Append(" where  cast(T1.docentry as varchar)+' '+cast(T1.LINENUM as varchar) IN (" + DocEntry + ") order by T1.Docentry,T1.LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static DataTable GetOrdr2Q(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select Docnum,ItemCode,Dscription,Price,T1.linenum,totalfrgn,T1.QUANTITY-ISNULL(T2.QUANTITY,0) QTY,VISORDER from PQT1 T1 left join OPQT T0 on (T1.docentry=T0.docentry)");
            sb.Append(" left join  (SELECT DOCENTRY,LINENUM,SUM(QUANTITY) QUANTITY FROM ACMESQLSP.DBO.SHIPPING_ITEM WHERE ITEMREMARK='採購報價'");
            sb.Append(" GROUP BY DOCENTRY,LINENUM) T2 ON (CAST(T1.DOCENTRY AS VARCHAR)=CAST(T2.DOCENTRY AS VARCHAR)  AND T1.LINENUM=T2.LINENUM)");
            sb.Append(" where  cast(T1.docentry as varchar)+' '+cast(T1.LINENUM as varchar) IN (" + DocEntry + ") order by T1.Docentry,T1.LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable GetOrdr2DRS(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select Docnum,ItemCode,Dscription,Price,T1.linenum,totalfrgn,T1.QUANTITY-ISNULL(T2.QUANTITY,0) QTY,VISORDER from por1 T1 left join opor T0 on (T1.docentry=T0.docentry)");
            sb.Append(" left join  (SELECT DOCENTRY,LINENUM,SUM(QUANTITY) QUANTITY FROM ACMESQLSPDRS.DBO.SHIPPING_ITEM WHERE ITEMREMARK='採購訂單'");
            sb.Append(" GROUP BY DOCENTRY,LINENUM) T2 ON (T1.DOCENTRY=T2.DOCENTRY AND T1.LINENUM=T2.LINENUM)");
            sb.Append(" where  cast(T1.docentry as varchar)+' '+cast(T1.LINENUM as varchar) IN (" + DocEntry + ") order by T1.Docentry,T1.LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable Getopdn2(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;


            string sql = "select * from pdn1 left join opdn on (pdn1.docentry=opdn.docentry) where opdn.Docnum=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static DataTable GetCHO(string ID)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.MEMO shipbuilding,T1.[Address] shipstreet,T1.Telephone shipblock,T1.FaxNo shipcity,T1.LinkMan shipzipcode");
            sb.Append(" ,T2.MEMO billbuilding,T2.[Address] billstreet,T2.Telephone billblock,T2.FaxNo billcity,T2.LinkMan billzipcode");
            sb.Append("    FROM   comCustDesc T0");
            sb.Append(" LEFT JOIN comCustAddress T1 ON (T0.DeliverAddrID=T1.AddrID and T0.ID=T1.ID AND T0.Flag =T1.Flag )");
            sb.Append(" LEFT JOIN comCustAddress T2 ON (T0.EngAddrID=T2.AddrID and T0.ID=T2.ID AND T0.Flag =T2.Flag )");
            sb.Append(" WHERE T0.ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable whs1(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select sum(orgqty) orgqty from whs1 where Docentry=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "whs1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["whs1"];
        }
        public static DataTable GetOwtr(string DocEntry)
        {

            SqlConnection MyConnection;


            MyConnection = globals.shipConnection;


            string sql = "select * from wtr1 left join owtr on (wtr1.docentry=owtr.docentry) where owtr.Docnum=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static DataTable GetOige(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;


            string sql = "select * from ige1 left join oige on (ige1.docentry=oige.docentry) where oige.Docnum=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ige1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ige1"];
        }
        public static DataTable GetOigN(string DocEntry)
        {

            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;


            string sql = "select * from ign1 left join oign on (ign1.docentry=oign.docentry) where oign.Docnum=@Docentry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ige1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ige1"];
        }
        public static DataTable GetSumInvoice(string SHIPPINGCODE, string InvoiceNo, string InvoiceNo_seq)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT InvoiceNo,InvoiceNo_seq,SUM(amount) AS amount  FROM invoiced  where SHIPPINGCODE=@SHIPPINGCODE and InvoiceNo=@InvoiceNo and InvoiceNo_seq=@InvoiceNo_seq group by InvoiceNo,InvoiceNo_seq";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@InvoiceNo", InvoiceNo));
            command.Parameters.Add(new SqlParameter("@InvoiceNo_seq", InvoiceNo_seq));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoiced");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoiced"];
        }
        public static DataTable GetPIno(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT  distinct(docentry) docentry  from shipping_item where shippingcode=@shippingcode";
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

        public static DataTable Getocrdnew1(string docentry, string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql;

            if (aa == "銷售訂單")
            {
                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'')  billzipcode FROM ORDR T0 LEFT JOIN  CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry ";

            }
            else if (aa == "AR貸項")
            {
                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'') billzipcode FROM ORIN T0 LEFT JOIN  CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry ";
            }
            else if (aa == "AP貸項")
            {
                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'') billzipcode FROM ORPC T0 LEFT JOIN  CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry ";
            }
            else if (aa == "採購退貨")
            {
                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'') billzipcode FROM ORPD T0 LEFT JOIN  CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry ";
            }
            else if (aa == "採購訂單")
            {
                sql = "select address,address2 from opor where docnum = @docentry";
            }
            else
            {
                sql = "select address,address2 from owtr where docnum = @docentry";
            }

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


        public static DataTable Getocrdnew2(string docentry, string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql;

            if (aa == "銷售訂單" || aa == "ALL" || aa=="三角")
            {

                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'') billzipcode FROM dbo.ORDR T0 LEFT JOIN  dbo.CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN dbo.CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry ";

            }
            else if (aa == "採購單")
            {

                sql = "select address,address2 from dbo.opor where docnum = @docentry";


            }
            else if (aa == "收貨採購單")
            {

                sql = "select address,address2 from dbo.opdn where docnum = @docentry";

            }
            else if (aa == "AR發票")
            {

                sql = "SELECT T0.DOCENTRY,T1.address shipaddress,T1.building shipbuilding,T1.street+ISNULL(T1.COUNTY,'') shipstreet ,T1.block shipblock ,T1.city shipcity ,T1.zipcode+ISNULL(T1.U_USERNAME,'') shipzipcode ,T2.address billaddress,T2.building billbuilding,T2.street+ISNULL(T2.COUNTY,'') billstreet,T2.block billblock,T2.city billcity,T2.zipcode+ISNULL(T2.U_USERNAME,'') billzipcode FROM dbo.OINV T0 LEFT JOIN  dbo.CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.shiptocode=T1.ADDRESS and T1.adrestype='S')  LEFT JOIN  dbo.CRD1 T2 ON (T0.CARDCODE=T2.CARDCODE AND T0.PAYTOCODE=T2.ADDRESS and T2.adrestype='B')  where t0.docnum = @docentry";

            }

            else if (aa.Contains("調撥"))
            {

                sql = "select U_PO_ADD,U_PO_ADD2 from dbo.owtr where docnum = @docentry";
            }
            else
            {
                sql = "select address,address2 from dbo.owtr where docnum = @docentry";
            }

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
        public static DataTable Getaddress(string docentry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql;


            sql = "select address,mailaddres,phone1,fax,cntctprsn,cardname cardname from ocrd where cardcode= @docentry";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }

        public static DataTable Getaddress2(string CARDCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql;


            sql = "SELECT T1.BUILDING 公司全稱,T1.STREET+ISNULL(T1.COUNTY,'') 地址,T1.BLOCK 電話,T1.CITY 傳真,T1.ZIPCODE+ISNULL(T1.U_USERNAME,'') 大名 FROM OCRD T0 LEFT JOIN CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.billtodef=T1.address AND  adrestype='b') where T0.CARDCODE=@CARDCODE UNION ALL SELECT T1.BUILDING 公司全稱,T1.STREET+ISNULL(T1.COUNTY,'') 地址,T1.BLOCK 電話,T1.CITY 傳真,T1.ZIPCODE+ISNULL(T1.U_USERNAME,'') 大名 FROM OCRD T0 LEFT JOIN CRD1 T1 ON (T0.CARDCODE=T1.CARDCODE AND T0.billtodef=T1.address AND  adrestype='S') where T0.CARDCODE=@CARDCODE ";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }
        public static DataTable GetShipmain(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select * from shipping_main where shippingcode = @shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "shipping_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["shipping_main"];
        }
        public static DataTable GetHR_Cardcode(string cardcode)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select address from ocrd where cardcode = @cardcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }




        public static DataTable GetUsername(string aa)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select * from [right] where username=@aa ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "odam");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["odam"];
        }
        public static DataTable GetOcrd(string cardcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string sql = "select * from ocrd where cardcode=@cardcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "ocrd");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["ocrd"];
        }
        public static DataTable Getinvoice(string shippingcode, string invoiceno, string invoiceno_seq)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select * from invoicem a left join invoiced b on (a.shippingcode=b.shippingcode and a.invoiceno=b.invoiceno and a.invoiceno_seq=b.invoiceno_seq) where a.shippingcode=@shippingcode and a.invoiceno=@invoiceno and a.invoiceno_seq=@invoiceno_seq";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("invoiceno_seq", invoiceno_seq));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }



        public static object[] GetMenuOw()
        {
            string[] FieldNames = new string[] { "Docnum", "itemcode" };

            string[] Captions = new string[] { "調撥單號", "品名" };

            string SqlScript = "SELECT cast(owtr.Docnum as varchar) as Docnum,wtr1.itemcode as itemcode FROM owtr left join wtr1 on (owtr.docentry=wtr1.docentry) order by cast(docnum as int) desc ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuOg()
        {
            string[] FieldNames = new string[] { "Docnum", "itemcode" };

            string[] Captions = new string[] { "調撥單號", "品名" };

            string SqlScript = "SELECT cast(oige.Docnum as varchar) as Docnum,ige1.itemcode as itemcode FROM oige left join ige1 on (oige.docentry=ige1.docentry) where oige.u_acme_kind='3' order by cast(docnum as int) desc ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetMenuOgN()
        {
            string[] FieldNames = new string[] { "Docnum", "itemcode" };

            string[] Captions = new string[] { "收貨單號", "品名" };

            string SqlScript = "SELECT cast(oign.Docnum as varchar) as Docnum,ign1.itemcode as itemcode FROM oign left join ign1 on (oign.docentry=ign1.docentry)  order by cast(docnum as int) desc ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static System.Data.DataTable GetSA(string PINO)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CASE WHEN T2.SlpCode='2' THEN '' ELSE (T2.[SlpName]) END 業務, CASE WHEN T2.SlpCode='2' THEN '' ELSE (T3.[lastName]+T3.[firstName]) END 業管,  CASE WHEN T2.SlpCode='2' THEN '' ELSE T2.U_EMAIL END 業務信箱,T3.EMAIL SA信箱  ");
            sb.Append("              FROM ORDR T0  ");
            sb.Append("              INNER JOIN OSLP T2 ON T0.SlpCode = T2.SlpCode ");
            sb.Append("              INNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  ");
            sb.Append("              WHERE    CAST(T0.DOCENTRY AS VARCHAR)=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", PINO));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }

        public static System.Data.DataTable GetSA2(string SHIPPINGCODE)
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                 SELECT T1.EMAIL,T3.EMAIL FROM WH_MAIN  T0   ");
            sb.Append("                            LEFT JOIN ACMESQL02.DBO.OHEM T1 ON (T0.BuCntctPrsn=T1.lastName +T1.firstName COLLATE Chinese_Taiwan_Stroke_CI_AS)  ");
            sb.Append("                              		    LEFT JOIN ACMESQL02.DBO.OCRD T4 ON (T0.CARDCODE =T4.CARDCODE COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("             		    LEFT JOIN ACMESQL02.DBO.OHEM T3 ON (T4.DfTcnician =T3.EMPID) ");

            sb.Append("  WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rdr1"];
        }
        public static System.Data.DataTable GetSHIPOHEM(string USER)
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT HOMETEL FROM OHEM WHERE (WORKCOUNTR='CN' OR homeTel IN ('APPLECHEN'))  AND HOMETEL=@HOMETEL AND ISNULL(TERMDATE,'') =''   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", USER));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetSHIPOHEM2(string USER)
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT HOMETEL FROM OHEM WHERE (WORKCOUNTR='CN')  AND HOMETEL=@HOMETEL AND ISNULL(TERMDATE,'') =''   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", USER));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                connection.Close();

            }

            return ds.Tables[0];

        }

        public static DataTable GETMEOCARD(string CARDCODE, string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT  SUBSTRING(NOTIFYMEMO,0,CHARINDEX(CHAR(13) , NOTIFYMEMO)) MEMO FROM SHIPPING_MAIN WHERE CARDCODE=@CARDCODE AND SHIPPINGCODE <>@SHIPPINGCODE AND SUBSTRING(NOTIFYMEMO,1,4) ='注意事項' ORDER BY SHIPPINGCODE DESC";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        public static DataTable Getmark(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select mark from mark where shippingcode=@shippingcode ORDER BY CAST(SEQ AS INT)";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }


        public static DataTable Getmark2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select mark as containerseals from mark");
            sb.Append(" where shippingcode=@shippingcode");
            sb.Append(" union all");
            sb.Append(" select containerseals  from ladingd ");
            sb.Append(" where shippingcode=@shippingcode");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        public static DataTable GetRmamark(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select mark from rma_mark where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rma_mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rma_mark"];
        }


        public static DataTable Getmeas(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select cfscode+'       '+unit as b  from cfs where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "cfs");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["cfs"];
        }
        public static DataTable Getgross(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select CASE WHEN substring(MAX(columnTotal),CHARINDEX('ONLY.',MAX(columnTotal))+5,40)='' THEN REPLACE(substring(MAX(columnTotal),CHARINDEX('(',MAX(columnTotal))+1,40),')','') ELSE substring(MAX(columnTotal),CHARINDEX('ONLY.',MAX(columnTotal))+5,40) END  PLTS, CAST(SUM(gross) AS VARCHAR)+'KGS' KGS,CAST(SUM(CAST(CBM AS DECIMAL(10,2))) AS VARCHAR)+'CBM' CBM 　 from PackingListM  where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PackingListM");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PackingListM"];
        }
        public static DataTable Getinvoicem(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from invoicem where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        public static DataTable Getpackm(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from packinglistm where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        public static DataTable GetLcInstro1(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from Shipping_Item where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Shipping_Item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Shipping_Item"];
        }

        public static DataTable GetAP_Prsn(string id)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from AP_Prsn where id=@id";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@id", id));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "AP_Prsn");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["AP_Prsn"];
        }

        public static DataTable Getemployee(string name)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select * from acmesql02.dbo.ohem where hometel=@name";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@name", name));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "employee");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["employee"];
        }
        public static DataTable Getemployee2(string name)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select * from acmesql02.dbo.ohem WHERE REPLACE(REPLACE(lastName +firstName,'(',''),')','') =@name";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@name", name));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "employee");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["employee"];
        }
        public static DataTable GetWHNAIL(string WHNAME)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT SEMAIL,CEMAIL FROM ACMESQLSP.DBO.WH_MAIL where (WHNAME=@WHNAME OR REPLACE(REPLACE(WHNAME,'倉',''),'-','')=@WHNAME)";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHNAME", WHNAME));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "employee");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["employee"];
        }
        public static DataTable Getshipinvo(string shippingcode, string invoiceno, string invoiceno_seq)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from invoiced where shippingcode=@shippingcode and invoiceno=@invoiceno and invoiceno_seq=@invoiceno_seq";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@invoiceno", invoiceno));
            command.Parameters.Add(new SqlParameter("@invoiceno_seq", invoiceno_seq));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoiceno");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoiceno"];
        }
        public static DataTable GetPacking(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select top 1 plno from PackingListM where shippingcode=@shippingcode ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoiceno");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoiceno"];
        }
        public static DataTable GetPacking2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select  plno from PackingListM where shippingcode=@shippingcode ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoiceno");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoiceno"];
        }
        public static DataTable Getwhs1(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select sum(cast(numebr as int)) as number from whs1 where docentry=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "whs1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["whs1"];
        }
        public static DataTable Getinv1(string docnum)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select cast(sum(amt) as varchar) as aa from PLC1 where docnum = @docnum ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docnum", docnum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "PLC1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["PLC1"];
        }
        public static object[] GetAR(string aa)
        {
            string[] FieldNames = new string[] { "Docnum", "LINENUM", "itemcode", "dscription", "quantity", "price" };

            string[] Captions = new string[] { "AR單號", "欄號", "項目號碼", "項目說明", "數量", "單價" };

            string SqlScript = "SELECT cast(T0.Docnum as varchar) as Docnum,cast(T1.LINENUM as varchar) LINENUM,T1.itemcode itemcode ,T1.dscription dscription,T1.quantity quantity,T1.price price FROM OINV  T0 left join inv1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  where t0.cardcode='" + aa + "' and t0.doctype='I' ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static object[] GetOINV(string aa)
        {
            string[] FieldNames = new string[] { "DocNum", "Cardname" };

            string[] Captions = new string[] { "AR單號", "客戶" };

            string SqlScript = "select cast(DocNum as nvarchar) DocNum,Cardname from OINV T0 WHERE t0.cardcode='" + aa + "' and t0.doctype='I'  ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static DataTable GetAR2(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   select T0.linetotal,T1.DocNum DOCNUM,T0.VatSumsy VatSumsy,T1.U_CHI_NO U_CHI_NO,T0.ItemCode ");
            sb.Append("           ItemCode,T0.Dscription Dscription,CAST(T0.Quantity AS INT) Quantity,");
            sb.Append("           isnull(SUM(T2.QTY),0) QTY,T0.Price Price,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE  END RATE,T0.Vatprcnt Vatprcnt,T0.GtotalFc GtotalFc");
            sb.Append("           ,T0.VatSumfrgn VatSumfrgn,T0.LineNum LineNum,t1.U_ACME_INV inv,t1.u_acme_rate1 匯率,Convert(varchar(10),t1.u_acme_invoice,112)  日期");
            sb.Append("            from PCH1 T0 left join OPCH T1 on (T0.docentry=T1.docentry) ");
            sb.Append("           left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO ");
            sb.Append("           AND T0.LINENUM = T2.LINENUM AND T2.PKIND='AP發票')");
            sb.Append("           where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ")");
            sb.Append("           GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity");
            sb.Append("           ,T0.Price,T0.Vatprcnt,T0.GtotalFc,T0.VatSumfrgn,T0.LineNum,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE  END,T0.linetotal,T0.VatSumsy,t1.U_ACME_INV,t1.u_acme_rate1,t1.u_acme_invoice");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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
        public static DataTable GetAPLC(string LcNo)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("        SELECT DOCNUM  FROM APLC WHERE LcNo =@LcNo  ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LcNo", LcNo));
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
        public static DataTable GetAR22(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                 select T0.linetotal,T1.DocNum DOCNUM,T0.VatSumsy VatSumsy,T1.U_CHI_NO U_CHI_NO,T0.ItemCode  ");
            sb.Append("                         ItemCode,T0.Dscription Dscription,CAST(T0.Quantity AS INT) Quantity, ");
            sb.Append("                         isnull(SUM(T2.QTY),0) QTY,T0.Price Price,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE  END RATE,T0.Vatprcnt Vatprcnt,T0.GtotalFc GtotalFc ");
            sb.Append("                         ,T0.VatSumfrgn VatSumfrgn,T0.LineNum LineNum,t1.U_ACME_INV inv,t1.u_acme_rate1 匯率,Convert(varchar(10),t1.u_acme_invoice,112)  日期 ");
            sb.Append("                          from PDN1 T0 left join OPDN T1 on (T0.docentry=T1.docentry)  ");
            sb.Append("                         left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO  ");
            sb.Append("                         AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購') ");
            sb.Append("           where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ")");
            sb.Append("                         GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity ");
            sb.Append("                         ,T0.Price,T0.Vatprcnt,T0.GtotalFc,T0.VatSumfrgn,T0.LineNum,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE  END,T0.linetotal,T0.VatSumsy,t1.U_ACME_INV,t1.u_acme_rate1,t1.u_acme_invoice ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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

        public static DataTable Get12(string Docentry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                         select (T0.Docentry),(T0.LINENUM),(T0.itemcode),(T0.dscription),(T0.PRICE),cast(T0.quantity as int) quantity,sum(isnull(T2.QTY,0)) QTY,cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) as AA,t1.U_ACME_INV INV,REPLACE(MAX(T1.CardName),'友達光電股份有限公司','')+'還款' 客戶,MAX(T1.CardCode) 客戶編號,MAX(T1.CardName) 客戶名稱 ");
            sb.Append("                             from acmesql02.dbo.PDN1 T0  ");
            sb.Append("                            left join acmesql02.dbo.OPDN T1 on (T0.docentry=T1.docentry) ");
            sb.Append("                            left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購')  ");
            sb.Append("                            where CAST(T1.Docentry AS VARCHAR)=@Docentry ");
            sb.Append("               group by t0.DocEntry,T0.LINENUM,T0.itemcode,T0.dscription,T0.PRICE,T0.quantity,t1.U_ACME_INV ");
            if (GetMenu.Day() != "20190614")
            {
                sb.Append("               having cast(T0.quantity as int)-sum(isnull(T2.QTY,0)) <> 0 ");
            }
            sb.Append("               order by T0.Docentry desc ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Docentry", Docentry));
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
        public static DataTable GetAR3(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("           select T0.linetotal,T1.DocNum DOCNUM,T0.VatSumsy VatSumsy,T1.U_CHI_NO U_CHI_NO,T0.ItemCode ");
            sb.Append("                   ItemCode,T0.Dscription Dscription,CAST(T0.Quantity AS INT) Quantity,");
            sb.Append("                   isnull(SUM(T2.QTY),0) QTY,T5.Price Price,CASE T4.RATE WHEN 0 THEN 1 ELSE T4.RATE END  RATE,T4.Vatprcnt Vatprcnt,T4.GtotalFc GtotalFc");
            sb.Append("                   ,T4.VatSumfrgn VatSumfrgn,T4.LineNum LineNum,t1.U_ACME_INV inv,cast(case isnull(t0.linetotal,0) when 0 then 0 else case t5.totalfrgn when 0 then 0 else t0.linetotal/t5.totalfrgn end end as decimal(10,4)) 匯率,Convert(varchar(10),t1.u_acme_invoice,112)  日期");
            sb.Append("                    from PCH1 T0 left join OPCH T1 on (T0.docentry=T1.docentry) ");
            sb.Append("                   left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO ");
            sb.Append("                   AND T0.LINENUM = T2.LINENUM AND T2.PKIND='AP發票')");
            sb.Append(" left join PDN1 t4 on (t0.baseentry=T4.docentry and  t0.baseline=t4.linenum )");
            sb.Append(" left join Por1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum )");
            sb.Append("           where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ")");
            sb.Append("                   GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity");
            sb.Append("                   ,T5.Price,T4.Vatprcnt,T4.GtotalFc,T4.VatSumfrgn,T4.LineNum,CASE T4.RATE WHEN 0 THEN 1 ELSE T4.RATE END,T0.linetotal,T0.VatSumsy,t1.U_ACME_INV,t0.linetotal,t5.totalfrgn,t1.u_acme_invoice");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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

        public static DataTable GetAR32(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("               select T0.linetotal,T1.DocNum DOCNUM,T0.VatSumsy VatSumsy,T1.U_CHI_NO U_CHI_NO,T0.ItemCode  ");
            sb.Append("                                 ItemCode,T0.Dscription Dscription,CAST(T0.Quantity AS INT) Quantity, ");
            sb.Append("                                 isnull(SUM(T2.QTY),0) QTY,T5.Price Price,CASE T0.RATE WHEN 0 THEN 1 ELSE T0.RATE END  RATE,T0.Vatprcnt Vatprcnt,T0.GtotalFc GtotalFc ");
            sb.Append("                                 ,T0.VatSumfrgn VatSumfrgn,T0.LineNum LineNum,t1.U_ACME_INV inv,cast(case isnull(t0.linetotal,0) when 0 then 0 else case t5.totalfrgn when 0 then 0 else t0.linetotal/t5.totalfrgn end end as decimal(10,4)) 匯率,Convert(varchar(10),t1.u_acme_invoice,112)  日期 ");
            sb.Append("                                  from PDN1 T0 left join OPDN T1 on (T0.docentry=T1.docentry)  ");
            sb.Append("                                 left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO  ");
            sb.Append("                                 AND T0.LINENUM = T2.LINENUM AND T2.PKIND='收貨採購') ");
            sb.Append("               left join Por1 t5 on (t0.baseentry=T5.docentry and  t0.baseline=t5.linenum ) ");
            sb.Append("           where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ")");
            sb.Append("                                 GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity ");
            sb.Append("                                 ,T5.Price,T0.Vatprcnt,T0.GtotalFc,T0.VatSumfrgn,T0.LineNum,CASE T0.RATE WHEN 0 THEN 1 ELSE T0.RATE END,T0.linetotal,T0.VatSumsy,t1.U_ACME_INV,t0.linetotal,t5.totalfrgn,t1.u_acme_invoice ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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
        public static DataTable GetSHICAR(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T1.SHIPPINGCODE,ISNULL(SUM(NET),0) NET, ISNULL(SUM(GROSS),0) GROSS,ISNULL(SUM(ISNULL(CAST(ISNULL(SAYTOTAL,0) AS DECIMAL(10,2)),0)),0) PACKAGE  ");
            sb.Append("                             ,MAX(T1.CARDNAME) CARDNAME,ISNULL(sum(T0.quantity),0) QTY,max(add7) OWNER ,MAX(T1.DOCTYPE) 類別,MAX(T1.PINO) DOC,MAX(T1.SENDGOODS) CBM,MAX(ADD1) ADD1   FROM Shipping_main  T1  ");
            sb.Append("               LEFT JOIN PackingListM T0 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("  WHERE T1.SHIPPINGCODE in ( " + SHIPPINGCODE + ") GROUP BY T1.SHIPPINGCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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

        public static DataTable GetSHICART(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T1.SHIPPINGCODE,ISNULL(SUM(NET),0) NET, ISNULL(SUM(GROSS),0) GROSS,ISNULL(SUM(ISNULL(CAST(ISNULL(SAYTOTAL,0) AS DECIMAL(10,2)),0)),0) PACKAGE  ");
            sb.Append("                             ,MAX(T1.CARDNAME) CARDNAME,ISNULL(sum(T0.quantity),0) QTY,max(add7) OWNER ,MAX(T1.DOCTYPE) 類別,MAX(T1.PINO) DOC,ISNULL(MAX(T1.SENDGOODS),0) CBM   FROM Shipping_main  T1  ");
            sb.Append("               LEFT JOIN PackingListM T0 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("  WHERE T1.SHIPPINGCODE =@SHIPPINGCODE GROUP BY T1.SHIPPINGCODE  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        public static DataTable GetSHICARSA(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT T3.HOMETEL SA FROM ORDR T0 iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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

        public static DataTable GetSHICARSA2(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT T3.HOMETEL SA FROM OPOR T0 iNNER JOIN OHEM T3 ON T0.OwnerCode = T3.empID  WHERE DOCENTRY=@DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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
        public static DataTable GetSHICAR2(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               select ISNULL(CASE WHEN CHARINDEX('-', PACKAGENO)<>0  ");
            sb.Append("               THEN CAST(SUBSTRING(PACKAGENO,CHARINDEX('-', PACKAGENO)+1,LEN(PACKAGENO)-CHARINDEX('-', PACKAGENO)) AS INT)-CAST(SUBSTRING(PACKAGENO,0,CHARINDEX('-', PACKAGENO)) AS INT)+1  ");
            sb.Append("               ELSE CASE WHEN ISNULL(PACKAGENO,0)<> 0 THEN 1 ELSE ISNULL(PACKAGENO,0) END  END,0)");
            sb.Append("               PACKAGE,REPLACE(MeasurmentCM,'@','') CM,MeasurmentCM CM2 from PackingListD ");
            sb.Append(" WHERE SHIPPINGCODE in (SELECT JOBNO FROM Shipping_CAR2 WHERE SHIPPINGCODE=@SHIPPINGCODE)");
            sb.Append("  AND ISNULL(MeasurmentCM,'') <> ''");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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
        public static DataTable GetSHICAR3(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT MeasurmentCM CM,SUM(CAST(PACKAGE AS INT)) PACKAGE  FROM Shipping_CAR3  WHERE SHIPPINGCODE=@SHIPPINGCODE GROUP BY MeasurmentCM");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
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


        public static void GetSHICARUP(string CHOPrice, string CHOAmount, string ShippingCode, string ITEMCODE, string QUANTITY, string ADD8, string CHOMEMO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE Shipping_CAR2 SET QTY=@QTY , Net=@Net,Gross=@CHOMEMO WHERE ShippingCode=@ShippingCode AND ITEMCODE=@ITEMCODE AND QUANTITY=@QUANTITY     ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CHOPrice", CHOPrice));
            command.Parameters.Add(new SqlParameter("@CHOAmount", CHOAmount));
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QUANTITY", QUANTITY));
            command.Parameters.Add(new SqlParameter("@ADD8", ADD8));
            command.Parameters.Add(new SqlParameter("@CHOMEMO", CHOMEMO));
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



        public static object[] GetOP(string aa)
        {
            string[] FieldNames = new string[] { "Docnum", "LINENUM", "itemcode", "dscription", "quantity", "price" };

            string[] Captions = new string[] { "採購單", "欄號", "項目號碼", "項目說明", "數量", "單價" };

            string SqlScript = "SELECT cast(T0.Docnum as varchar) as Docnum,cast(T1.LINENUM as varchar) LINENUM,T1.itemcode itemcode ,T1.dscription dscription,T1.quantity quantity,T1.price price FROM OPOR T0 left join POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)   where t0.cardcode='" + aa + "' and t0.doctype='I' ";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }
        public static DataTable PackOP(string shippingcode)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "select 'quantity :'+isnull(cast(a.quantity as varchar),'')+' net :'+isnull(cast(a.net as varchar),'')+' gross :'+isnull(cast(a.gross as varchar),'')+' package :'+isnull(cast(a.saytotal as varchar),'')+' 20呎 :'+isnull(cast(b.boardCount as varchar),'')+' 40呎 :'+isnull(cast(b.boardDeliver as varchar),'')+' 併櫃/CBM :'+isnull(b.sendGoods,'') as aa FROM acmesqlsp.dbo.PackingListM a left join acmesqlsp.dbo.shipping_main b on (a.shippingcode=b.shippingcode) where a.shippingcode=@shippingcode ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " POR1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" POR1 "];
        }

        public static DataTable GetOP2(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.linetotal,T1.DocNum DOCNUM,T1.U_CHI_NO U_CHI_NO,T0.ItemCode ");
            sb.Append(" ItemCode,T0.Dscription Dscription,T0.VatSumsy VatSumsy,CAST(T0.Quantity AS INT) Quantity,");
            sb.Append(" isnull(SUM(T2.QTY),0) QTY,T0.Price Price,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE END  RATE,T0.Vatprcnt Vatprcnt,T0.GtotalFc GtotalFc");
            sb.Append(" ,T0.VatSumfrgn VatSumfrgn,T0.LineNum LineNum");
            sb.Append("  from POR1 T0 left join OPOR T1 on (T0.docentry=T1.docentry) ");
            sb.Append(" left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO ");
            sb.Append(" AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購單')");
            sb.Append(" where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ") and t0.docentry > 400 ");

            sb.Append(" GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity");
            sb.Append(" ,T0.Price,T0.Vatprcnt,T0.GtotalFc,T0.VatSumfrgn,T0.LineNum,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE END,T0.linetotal,T0.VatSumsy");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " POR1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" POR1 "];
        }

        public static DataTable GetOP2Q(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.linetotal,T1.DocNum DOCNUM,T1.U_CHI_NO U_CHI_NO,T0.ItemCode ");
            sb.Append(" ItemCode,T0.Dscription Dscription,T0.VatSumsy VatSumsy,CAST(T0.Quantity AS INT) Quantity,");
            sb.Append(" isnull(SUM(T2.QTY),0) QTY,T0.Price Price,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE END  RATE,T0.Vatprcnt Vatprcnt,T0.GtotalFc GtotalFc");
            sb.Append(" ,T0.VatSumfrgn VatSumfrgn,T0.LineNum LineNum");
            sb.Append("  from PQT1 T0 left join OPQT T1 on (T0.docentry=T1.docentry) ");
            sb.Append(" left join acmesqlSP.dbo.PLC1 T2 on (T0.Docentry=T2.DONNO ");
            sb.Append(" AND T0.LINENUM = T2.LINENUM AND T2.PKIND='採購報價')");
            sb.Append(" where cast(T0.docentry as varchar)+' '+cast(T0.LINENUM as varchar) IN (" + DocEntry + ")  ");

            sb.Append(" GROUP BY T1.DocNum,T1.U_CHI_NO,T0.ItemCode,T0.Dscription,T0.Quantity");
            sb.Append(" ,T0.Price,T0.Vatprcnt,T0.GtotalFc,T0.VatSumfrgn,T0.LineNum,CASE ISNULL(T0.RATE,0) WHEN 0 THEN 1 ELSE T0.RATE END,T0.linetotal,T0.VatSumsy");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " POR1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" POR1 "];
        }
        public static DataTable download(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from download where filename=@DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable downloadCAR(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from Shipping_CARDownload where filename=@DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable GETSUIP()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select PARAM_NO  from RMA_PARAMS  WHERE PARAM_KIND ='SUIP'";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable OQUTDownload(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from shipping_OQUTDownload where filename=@DocEntry ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable OQUTDownload2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from shipping_OQUTDownload2 where filename=@DocEntry ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable download3(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from download3 where filename=@DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static void DELdownload2(string filename)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE download2 where filename=@filename ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@filename", filename));


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
        public static DataTable download2(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from download2 where filename=@DocEntry ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable getaa(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select shippingcode from shipping_item where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable getCAR(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT FLAG1 FROM  WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(FLAG1,'') <>''";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static DataTable GetSHICARSACHOICE(string DOCENTRY)
        {

            string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
            SqlConnection MyConnection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT P.PersonName  FROM [ordBillMain] T0  left join comPerson P ON(T0.Salesman = P.PersonID)  WHERE BILLNO = @DOCENTRY");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
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

        public static DataTable getPACK(string SHIPPINGCODE)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "SELECT DISTINCT FLAG1 FROM  WH_PACK2 WHERE SHIPPINGCODE=@SHIPPINGCODE";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" download "];
        }
        public static object[] GetAccountdate()
        {
            string[] FieldNames = new string[] { "aa", "bb" };

            string[] Captions = new string[] { "過帳日期", "數量" };

            string SqlScript = "select cast(Convert(varchar(8),docdate,112) as varchar) as aa,cast(count(*) as varchar)  as bb from inv1 group by docdate";


            SqlLookup dialog = new SqlLookup();

            dialog.Captions = Captions;
            dialog.FieldNames = FieldNames;

            dialog.SqlScript = SqlScript;
            try
            {


                if (dialog.ShowDialog() == DialogResult.OK)
                {



                    object[] LookupValues = dialog.LookupValues;
                    return LookupValues;

                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dialog.Dispose();
            }
        }

        public static System.Data.DataTable GetMail(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select Docentry,Dscription,Quantity,T1.LC,seqno,Checked,T0.shippingcode,RED,T0.DocNum  from    dbo.LcInstro  T0 INNER JOIN dbo.LcInstro1  T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) where T0.shippingcode=@shippingcode OR T0.SHIPPINGCODE='1'";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }

        public static System.Data.DataTable GetMailSA(string shippingcode, string SEQNO)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT REMARK FROM SHIPPING_ITEM WHERE SHIPPINGCODE=@shippingcode AND SEQNO=@SEQNO ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable GetIN(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select distinct docentry from dbo.LcInstro1 where shippingcode=@shippingcode ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable Getdet(string WHSCODE)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT LOCATION FROM SHIPPING_WHS WHERE WHSCODE=@WHSCODE ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSCODE", WHSCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable GetTIFF2(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("                     SELECT SUM(CAST(quantity AS INT)) QTY,CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                           SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                           SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                          AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                          Substring (T1.[ItemCode],2,8) END+'.'+Substring([ItemCode],12,1) MODEL ");
            sb.Append("                            FROM dbo.LcInstro  T0");
            sb.Append(" INNER JOIN dbo.LcInstro1  T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM )");
            sb.Append("               WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append("               GROUP BY CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND  ");
            sb.Append("                           SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND  ");
            sb.Append("                           SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%' ");
            sb.Append("                          AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE  ");
            sb.Append("                          Substring (T1.[ItemCode],2,8) END,Substring([ItemCode],12,1)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }

        public static System.Data.DataTable GetTIFF2AD(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(CAST(quantity AS INT)) QTY,T1.Dscription  MODEL ");
            sb.Append(" FROM dbo.LcInstro  T0 ");
            sb.Append(" INNER JOIN dbo.LcInstro1  T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE AND T0.DOCNUM=T1.DOCNUM ) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE  GROUP BY T1.Dscription  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable GetTIFF(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT CASE CHARINDEX('HONG', unloadCargo) WHEN 0  THEN  ltrim(substring(unloadCargo,CHARINDEX(',', unloadCargo)+1,10)) ELSE 'HK' END  國家,CASE  WHEN CHARINDEX('KONG',RTRIM(LTRIM(unloadCargo))) <> 0 THEN 'HK' WHEN CHARINDEX(',', unloadCargo)=0 THEN unloadCargo  ELSE ltrim(substring(unloadCargo,0,CHARINDEX(',', unloadCargo))) END 地名 FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE  ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable GetMail2(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from dbo.wh_item where shippingcode=@shippingcode OR SHIPPINGCODE='1'";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }

        public static System.Data.DataTable Getfee(string feeDoccur)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind=@feeDoccur order by DataValue ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@feeDoccur", feeDoccur));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static System.Data.DataTable Getfee2(string feeDoccur)
        {
            SqlConnection connection = globals.Connection;
            string sql = "SELECT '' DataValue,'' DataText UNION ALL SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind=@feeDoccur  ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@feeDoccur", feeDoccur));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "LcInstro1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["LcInstro1"];
        }
        public static string DFirst()
        {
            DateTime DFirst =
    DateTime.Parse(DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + "1");
            DateTime DLast;
            if (DFirst.Month != 12)
            {
                DLast = DFirst.AddMonths(1).AddDays(-1);
            }
            else
            {
                DLast = DFirst.AddDays(30);
            }
            return DFirst.ToString("yyyyMMdd");
        }
        public static string DLast()
        {
            DateTime DFirst =
    DateTime.Parse(DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + "1");
            DateTime DLast;
            if (DFirst.Month != 12)
            {
                DLast = DFirst.AddMonths(1).AddDays(-1);
            }
            else
            {
                DLast = DFirst.AddDays(30);
            }
            return DLast.ToString("yyyyMMdd");
        }
        public static string DLast3()
        {
            DateTime DFirst =
    DateTime.Parse(DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + "1");
            DateTime DLast;
            if (DFirst.Month != 12)
            {
                DLast = DFirst.AddMonths(1).AddDays(-1);
            }
            else
            {
                DLast = DFirst.AddDays(30);
            }
            return DLast.ToString("yyyy") + "/" + DLast.ToString("MM") + "/" + DLast.ToString("dd");
        }
        public static string DLast2(string yearmonth)
        {
            string year = yearmonth.Substring(0, 4);
            string month = yearmonth.Substring(4, 2);

            DateTime DFirst =
    DateTime.Parse(year + "-" + month + "-" + "1");
            DateTime DLast;
            if (DFirst.Month != 12)
            {
                DLast = DFirst.AddMonths(1).AddDays(-1);
            }
            else
            {
                DLast = DFirst.AddDays(30);
            }
            return DLast.ToString("yyyyMMdd");
        }
        public static string DLast3(string yearmonth)
        {
            string year = yearmonth.Substring(0, 4);
            string month = yearmonth.Substring(4, 2);

            DateTime DFirst =
    DateTime.Parse(year + "-" + month + "-" + "1");
            DateTime DLast;
            if (DFirst.Month != 12)
            {
                DLast = DFirst.AddMonths(1).AddDays(-1);
            }
            else
            {
                DLast = DFirst.AddDays(30);
            }
            return DLast.ToString("yyyy") + "/" + DLast.ToString("MM") + "/" + DLast.ToString("dd");
        }
        public static string SHIP(DateTime SHIPDATE)
        {
            int DATE1 = Convert.ToInt16(SHIPDATE.ToString("dd"));

            if (DATE1 < 11)
            {
                SHIPDATE.AddMonths(-1);
            }

            return SHIPDATE.ToString("yyyyMM");
        }
        public static string Day()
        {

            return DateTime.Now.ToString("yyyyMMdd");
        }

        public static string DayS(string DATE)
        {

            return DATE.Substring(0, 4) + "/" + DATE.Substring(4, 2) + "/" + DATE.Substring(6, 2);

        }

        public static string DaySWHNO(string DATE)
        {
            string DD = DATE.Substring(4, 2) + "/" + DATE.Substring(6, 2);
            if (DATE.Substring(4, 1) == "0")
            {
                DD = DATE.Substring(5, 1) + "/" + DATE.Substring(6, 2);
            }

            DD = DD.Replace("/0", "");
            return DD;

        }
        public static string DAYTIME()
        {

            return DateTime.Now.ToString("yyyyMMddhhmm");
        }
        public static string DayYEAR()
        {

            return DateTime.Now.ToString("yyyy");
        }

        public static DataTable Account_LCDownload(string LCID, string ID2)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  filepath from  Account_LCDownload");
            sb.Append("  WHERE LCID=@LCID and ID2=@ID2  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@LCID", LCID));

            command.Parameters.Add(new SqlParameter("@ID2", ID2));

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

        public static DataTable GETENG()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT RMA_ENG DataValue FROM RMA_ENGINNER WHERE ACTIVE='Y' AND AKIND IN ('ENG','CUST') AND ISNULL(PLACE,'') <> 'SZ' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            // command.Parameters.Add(new SqlParameter("@PLACE", PLACE));



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
        public static DataTable GETENG2(string RMA_ENG)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT PLACE,engname FROM RMA_ENGINNER WHERE RMA_ENG=@RMA_ENG ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@RMA_ENG", RMA_ENG));



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

        public static DataTable GETDEFECTCODE(string DEFECTCODE, string DTYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select * FROM RMA_DEFECTCODE WHERE DEFECTCODE=@DEFECTCODE AND DTYPE=@DTYPE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DEFECTCODE", DEFECTCODE));

            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));



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
        public static DataTable GETDEFECTCODE2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select * FROM RMA_DEFECTCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;





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
        public static System.Data.DataTable BU()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='BU' order by DataValue  ";

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
        public static System.Data.DataTable Status()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='結案' order by DataValue  ";

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
        public static System.Data.DataTable Year()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM acmesqlsp.dbo.RMA_PARAMS where param_kind='shipyear' order by DataValue desc ";

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

        public static System.Data.DataTable Year2017()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='shipyear' AND PARAM_NO >2015 order by DataValue desc ";

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
        public static System.Data.DataTable Month()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='shipmonth' order by cast(PARAM_NO as int)";

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
        public static System.Data.DataTable Month2()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataTex　FROM RMA_PARAMS where param_kind='shipmonth'　UNION ALL SELECT '',0 UNION ALL SELECT 'ALL未結PO',0 order by cast(PARAM_DESC as int)";

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
        public static System.Data.DataTable GetBU(string KIND)
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by DataValue";

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
        public static System.Data.DataTable GetBUGB(string KIND)
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM ACMESQLSP.DBO.RMA_PARAMS where param_kind='" + KIND + "' order by ID";

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

        public static System.Data.DataTable GetWHPLATE(string MEMO)
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM PARAMS where param_kind='PLATE' AND MEMO='" + MEMO + "' order by DataValue";

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
        public static System.Data.DataTable GetBUPLATE()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM PARAMS where param_kind='PLATEBU'  order by int";

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
        public static System.Data.DataTable MoneyBU(string KIND)
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' ";

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
        public static System.Data.DataTable SolarBU()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='solar' order by id ";

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
        public static System.Data.DataTable GetOslp1()
        {

            SqlConnection con = globals.shipConnection;
            string sql = "select slpname as DataValue  from oslp where ISNULL(memo,'') <> ''  UNION ALL SELECT 'Please-Select'   as DataValue  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM' ORDER BY DataValue";


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

        public static System.Data.DataTable GetOhem()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "select [lastName]+[firstName] as DataValue   FROM OHEM WHERE (WORKCOUNTY='Y' or EMPID=113) AND ISNULL(TERMDATE,'') = ''  UNION ALL SELECT 'Please-Select'   as DataValue ORDER BY DataValue ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ohem");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ohem"];
        }

        public static Int32 DaySpan(string EndDate, string BaseDate)
        {
            TimeSpan s = StrToDate(BaseDate) - StrToDate(EndDate);

            return s.Days;

        }
        public static string SPLITDOC(string aa)
        {
            string[] arrurl = aa.Split(new Char[] { ',' });

            StringBuilder sb2 = new StringBuilder();
            foreach (string i in arrurl)
            {
                sb2.Append("'" + i + "',");
            }
            sb2.Remove(sb2.Length - 1, 1);
            return sb2.ToString();
        }
        public static DateTime StrToDate(string sDate)
        {

            UInt16 Year = Convert.ToUInt16(sDate.Substring(0, 4));
            UInt16 Month = Convert.ToUInt16(sDate.Substring(4, 2));
            UInt16 Day = Convert.ToUInt16(sDate.Substring(6, 2));

            return new DateTime(Year, Month, Day);
        }

        public static System.Data.DataTable GETORTT()
        {

            SqlConnection connection = new SqlConnection(strCnSP);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT YTEMP FROM HR_TEMP WHERE DOCDATE=@DOCDATE AND USERS=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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
        public static System.Data.DataTable GETORTT2(string DOCDATE)
        {

            SqlConnection connection = new SqlConnection(strCnSP);

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT DOCDATE 日期,YTEMP 體溫 FROM HR_TEMP WHERE USERS=@USERS AND SUBSTRING(DOCDATE,1,6)=@DOCDATE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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
        public static void InsertLog(string LoginID, string Event, string Detail, string Docdate)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO Shipping_Log (LoginID,Event,Detail,Docdate) VALUES (@LoginID,@Event,@Detail,@Docdate)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@LoginID", LoginID));
            command.Parameters.Add(new SqlParameter("@Event", Event));
            command.Parameters.Add(new SqlParameter("@Detail", Detail));
            command.Parameters.Add(new SqlParameter("@Docdate", Docdate));


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
        //public static void InsertLOGIN(string LOGINID, string SHIPPINGCODE)
        //{

        //    SqlConnection connection = globals.Connection;
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(" INSERT INTO SHIPPING_LOGIN (LOGINID,DATETIME,SHIPPINGCODE) VALUES (@LOGINID,@DATETIME,@SHIPPINGCODE)");
        //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
        //    command.CommandType = CommandType.Text;
        //    SqlDataAdapter da = new SqlDataAdapter(command);

        //    command.Parameters.Add(new SqlParameter("@LOGINID", LOGINID));
        //    command.Parameters.Add(new SqlParameter("@DATETIME", DateTime.Now.ToString("yyyyMMddHHmmss")));
        //    command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


        //    try
        //    {

        //        try
        //        {
        //            connection.Open();
        //            command.ExecuteNonQuery();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        connection.Close();
        //    }


        //}

        public static void InsertEXEXPORT(string NAME, string SQLPARAM, string EXUSER, string DOCDATE, string PATH)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO ACME_EXEXPORT(NAME,SQLPARAM,EXUSER,DOCDATE,PATH) VALUES (@NAME,@SQLPARAM,@EXUSER,@DOCDATE,@PATH)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@NAME", NAME));
            command.Parameters.Add(new SqlParameter("@SQLPARAM", SQLPARAM));
            command.Parameters.Add(new SqlParameter("@EXUSER", EXUSER));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@PATH", PATH));

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
        //public static void TRUNLOGIN()
        //{

        //    SqlConnection connection = globals.Connection;
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(" truncate table acmesqlsp.dbo.SHIPPING_LOGIN truncate table acmesqlspdrs.dbo.SHIPPING_LOGIN");
        //    SqlCommand command = new SqlCommand(sb.ToString(), connection);
        //    command.CommandType = CommandType.Text;
        //    SqlDataAdapter da = new SqlDataAdapter(command);




        //    try
        //    {

        //        try
        //        {
        //            connection.Open();
        //            command.ExecuteNonQuery();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        connection.Close();
        //    }


        //}
        public static void DELETELOGIN(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" DELETE SHIPPING_LOGIN  WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


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

        public bool IsNatural_Number(string str)
        {
            System.Text.RegularExpressions.Regex reg1 = new System.Text.RegularExpressions.Regex(@"^[A-Za-z0-9]+$");
            return reg1.IsMatch(str);
        }
        public static void Insertcho(string PARAM_NO, string PARAM_DESC)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO RMA_PARAMS (PARAM_KIND,PARAM_NO,PARAM_DESC) VALUES ('CHOICE',@PARAM_NO,@PARAM_DESC)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@PARAM_NO", PARAM_NO));
            command.Parameters.Add(new SqlParameter("@PARAM_DESC", PARAM_DESC));


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
        public static void Updatecho(string CHOPrice, string CHOAmount, string ShippingCode, string ITEMCODE, string QUANTITY, string ADD8, string CHOMEMO)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE shipping_item SET CHOPrice=@CHOPrice , CHOAmount=@CHOAmount,CHOMEMO=@CHOMEMO WHERE ShippingCode=@ShippingCode AND ITEMCODE=@ITEMCODE AND QUANTITY=@QUANTITY     ");
            sb.Append(" UPDATE shipping_MAIN SET ADD8=@ADD8  WHERE ShippingCode=@ShippingCode   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@CHOPrice", CHOPrice));
            command.Parameters.Add(new SqlParameter("@CHOAmount", CHOAmount));
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QUANTITY", QUANTITY));
            command.Parameters.Add(new SqlParameter("@ADD8", ADD8));
            command.Parameters.Add(new SqlParameter("@CHOMEMO", CHOMEMO));
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
        public static void Updatecho2(string ShippingCode, string ITEMCODE, string QUANTITY)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE shipping_item SET CHOPrice=0, CHOAmount=0  WHERE ShippingCode=@ShippingCode AND ITEMCODE=@ITEMCODE AND QUANTITY=@QUANTITY     ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@QUANTITY", QUANTITY));

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

        public static void Updatecho3(string ShippingCode, string ITEMCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE shipping_item SET CHOPrice=0, CHOAmount=0  WHERE ShippingCode=@ShippingCode AND ITEMCODE=@ITEMCODE    ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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
        public static System.Data.DataTable GETRMAWH()
        {
            SqlConnection con = globals.Connection;

            string sql = "SELECT PARAM_NO FROM dbo.PARAMS WHERE PARAM_KIND ='RMAWH' ORDER BY PARAM_NO ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        public static System.Data.DataTable Getwarehouse()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "SELECT whscode DataValue,street DataText  FROM owhs where isnull(street,'') <> ''  order by cast(isnull(block,1000) as int)  ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        public static System.Data.DataTable GetwarehouseWS()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "SELECT whscode DataValue,street DataText  FROM owhs where isnull(street,'') <> '' and whscode not in ('TW006','TW007','RM001','RM002','RM004','BW001','BW002','BW003','Z0005','CC002','LB001','OT001','TW002','TW004','TW005') order by cast(isnull(block,1000) as int)  ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        public static System.Data.DataTable Getwarehouse1()
        {
            SqlConnection con = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT whscode as DataValue,street as DataText,BLOCK FROM owhs where street is not null");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '全部倉','全部倉',NULL");
            sb.Append("  order by BLOCK desc,whscode");

            SqlDataAdapter da = new SqlDataAdapter(sb.ToString(), con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        public static System.Data.DataTable GETEMP()
        {
            SqlConnection con = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT homeTel DataText FROM OHEM ");
            sb.Append(" WHERE homeTel IN ( 'APPLECHEN', 'SUNNYWANG', 'HELENWU', 'SHIRLEYJUAN', 'NANCYTSAI', 'BETTYTSENG', 'DAVIDHUANG','shirleyJUAN','JOYCHEN','EVAHSU')");
            sb.Append(" ORDER BY homeTel");
            SqlDataAdapter da = new SqlDataAdapter(sb.ToString(), con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }
        public static System.Data.DataTable GetwarehouseCHI()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "select WareHouseID DataValue,case ShortName when '' then Memo else ShortName end DataText from comWareHouse WHERE EngName = 'V'";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }

        public static System.Data.DataTable GetwarehouseAD()
        {
            SqlConnection con = globals.shipConnection;

            string sql = "select WareHouseID DataValue,ShortName DataText from comWareHouse ";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "owhs");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["owhs"];
        }


        public static DataTable GETDRSINV(string shippingcode)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT SUM(QUANTITY) 數量,T0.SHIPPINGCODE 工單號碼,INVOICENO+INVOICENO_SEQ INV FROM SHIPPING_ITEM T0");
            sb.Append(" LEFT JOIN INVOICED T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@shippingcode");
            sb.Append(" GROUP BY INVOICENO+INVOICENO_SEQ,T0.SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }




        public static DataTable GETDRSINV2(string shippingcode)
        {
            StringBuilder sb = new StringBuilder();

            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT DISTINCT CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("         SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("        AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append(" Substring (T1.[ItemCode],2,8) END MODEL FROM SHIPPING_ITEM T1");
            sb.Append(" WHERE T1.SHIPPINGCODE=@shippingcode");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        public static void UPDATEPICK(string SHIPPINGCODE, string BILLNO, int PACK1)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE GB_PICK2 SET PACK1=@PACK1 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO   ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@PACK1", PACK1));


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
        public static void UPDATEPICKCHECK(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE GB_PICK SET CHECKED='Checked',CHECKEDDATE=@CHECKEDDATE WHERE  SHIPPINGCODE=@SHIPPINGCODE    ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@CHECKEDDATE", GetMenu.Day()));


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
        public static void UPDATEPICK2(string SHIPPINGCODE, string BILLNO, string PACK3, string LINE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE GB_PICK2 SET PACK3=@PACK3 WHERE  SHIPPINGCODE=@SHIPPINGCODE AND BILLNO=@BILLNO AND LINE=@LINE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
            command.Parameters.Add(new SqlParameter("@PACK3", PACK3));
            command.Parameters.Add(new SqlParameter("@LINE", LINE));
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

        public static System.Data.DataTable Getdata(string DOCKIND)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCPATH from WH_SAPDOC where USERS=@USERS AND DOCKIND=@DOCKIND ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim()));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "auogd4");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static void Add(string DOCPATH, string DOCKIND)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO [WH_SAPDOC]");
            sb.Append("            (USERS,DOCPATH,DOCKIND)");
            sb.Append("      VALUES(@USERS,@DOCPATH,@DOCKIND)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim()));
            command.Parameters.Add(new SqlParameter("@DOCPATH", DOCPATH));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
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
        public static System.Data.DataTable GetOWHS3(string WHSNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT * FROM OWHS WHERE ZIPCODE='三角' AND REPLACE(REPLACE(WHSNAME,'(',''),')','')=@WHSNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@WHSNAME", WHSNAME));
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

        public static void UP(string DOCPATH, string DOCKIND)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE [WH_SAPDOC] SET DOCPATH=@DOCPATH WHERE USERS=@USERS AND DOCKIND=@DOCKIND");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString().Trim()));
            command.Parameters.Add(new SqlParameter("@DOCPATH", DOCPATH));
            command.Parameters.Add(new SqlParameter("@DOCKIND", DOCKIND));
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
        public static System.Data.DataTable GetSAALL()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE  ISNULL(TERMDATE,'') = '' AND (workCountr ='CN' AND ISNULL(JOBTITLE,'') = ('業助') ) ORDER BY HOMETEL  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        public static System.Data.DataTable GetWHSHIP()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE JOBTITLE='船務'  AND ISNULL(TERMDATE,'') =''   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetWHCN()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE dept in (5,7)  AND ISNULL(TERMDATE,'') =''  AND workCountR='CN' AND homeTel <> 'ErinChou'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetWHSTOCK()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("   SELECT EMAIL FROM OHEM WHERE JOBTITLE='船務倉管'  AND ISNULL(TERMDATE,'') =''  AND EMPID <> 135　 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetWHSA()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT HOMETEL FROM OHEM WHERE (JOBTITLE='業助' OR HOMETEL  IN ('MIMICHEN','REBECCALIN','applechen','joychen','ShirleyJuan','evahsu','MILLYGENG','bettytseng','SUNNYWANG')) AND ISNULL(TERMDATE,'') = '' ORDER BY HOMETEL ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetWHSA2()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ' 後勤' HOMETEL");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT HOMETEL FROM OHEM WHERE (JOBTITLE='業助' OR HOMETEL  IN ('MIMICHEN','REBECCALIN','MILLYGENG')) AND ISNULL(TERMDATE,'') = ''");
            sb.Append(" ORDER BY HOMETEL");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public static System.Data.DataTable GetWHSA3()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT EMAIL FROM OHEM WHERE (HOMEZIP='Y') AND ISNULL(TERMDATE,'') = ''");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetWHSAALL()
        {

            SqlConnection connection = new SqlConnection(strCn02);

            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT CASE WHEN HOMETEL IN ('AppleChen','ViviWeng') THEN HOMETEL+'(GT)' ELSE HOMETEL END  HOMETEL FROM OHEM WHERE  ISNULL(TERMDATE,'') = '' AND (workCountr ='CN' AND ISNULL(JOBTITLE,'') IN ('業助','','船務倉管')) OR (EMPID IN ('21','11','103')) ORDER BY HOMETEL");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public static System.Data.DataTable GetSAPRevenue(string DocDate1, string DocDate2)
        {
            //合計 AS 銷售金額
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            //AR服務
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )   總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' ");
            sb.Append(" and (((T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%' )  and isnull(t2.u_acme_arap,'') <> 'xx' ) OR (T1.[Account]='22610103' AND (U_LOCATION)='XX' ))");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            sb.Append(" SELECT '貸項' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112)  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            //貸項服務
            sb.Append(" union all");
            sb.Append(" SELECT '貸項-服務' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account 科目代號,");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%'  or T1.[Account] like '4190%' or T1.[Account] like '4210%'   )  ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            if ((globals.DBNAME != "達睿生"))
            {
                sb.Append(" SELECT 'AR' as 單別,T7.DOCENTRY DocNum,T0.[TransId],");
                sb.Append(" MAX(T7.AcctCode)  科目代號,");
                sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,0 總金額,MAX(T0.[RefDate]) 日期,");
                sb.Append(" 0  總成本,");
                sb.Append(" 0  - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
                sb.Append(" inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
                sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
                sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
                sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
                sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
                sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
                sb.Append("  AND T2.[DocTotal] = 0 	AND T7.DOCENTRY <> 49540		 ");
                sb.Append(" GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
                sb.Append(" union all");
            }

            sb.Append("              SELECT 'AR預' as 單別,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account 科目代號,");
            sb.Append("              T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T6.[DOCDATE])  日期,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  總成本,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("              INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("              INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");

            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            sb.Append("              and T6.DOCDATE BETWEEN Convert(varchar(8),@DocDate1,112) and  Convert(varchar(8),@DocDate2,112) ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            sb.Append(" union all");
            //20120419 AR貸項沒有收入
            sb.Append("             SELECT '貸項' as 單別,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T6.AcctCode)  科目代號,");
            sb.Append("             T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,0 總金額,MAX(T0.[RefDate]) 日期,");
            sb.Append("             SUM(T1.[Debit] - T1.[Credit])  總成本,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) 總毛利");
            sb.Append("             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("             INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("             INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("             INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append("             WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("             GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            //20150916 AR預開貸項服務
            sb.Append("               SELECT '貸項-服務' as 單別,T2.[DocNum],T0.[TransId], ");
            sb.Append("               T1.Account 科目代號, ");
            sb.Append("               T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組, ");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) 總金額,MAX(T6.[DOCDATE])  日期,");
            sb.Append("               SUM(T1.[Debit] - T1.[Credit])  總成本, ");
            sb.Append("               (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) 總毛利 ");
            sb.Append("               FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               INNER JOIN ORIN T2 ON T0.TransId = T2.TransId  ");
            sb.Append("               INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("               INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE  ");
            sb.Append("               INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("        INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0 ");
            sb.Append("               LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13 ");
            sb.Append("               GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.U_ACME_ARAP=T6.BASEENTRY) ");
            sb.Append("               WHERE T2.[DocType] ='S' AND T1.ACCOUNT='22610103' AND U_LOCATION='XX'");
            sb.Append("              and T6.DOCDATE BETWEEN Convert(varchar(8),@DocDate1,112) and  Convert(varchar(8),@DocDate2,112) ");
            sb.Append("     GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            sb.Append(" union all");
            //20151006  折讓貸項
            sb.Append("                        SELECT 'JE' as 單別,T0.TransId,T0.[TransId],  ");
            sb.Append("                             T1.Account 科目代號,  ");
            sb.Append("                           T1.REF1 業務員編號, T3.SlpName  姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append("                             SUM(T1.debit)*(-1) 總金額,MAX(T0.REFDATE) 日期,  ");
            sb.Append("                             0  總成本,  ");
            sb.Append("                                     SUM(T1.debit)*(-1) 總毛利  ");
            sb.Append("                             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("                             INNER JOIN OSLP T3 ON T1.REF1 = T3.SlpCode  ");
            sb.Append(" INNER JOIN OCRD T2 ON T1.U_REMARK1 = T2.CARDCODE");
            sb.Append("               INNER JOIN OCRG T5 ON T2.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("                             WHERE T1.ACCOUNT='41900101' and isnull(T1.REF2,'')  ='xx'");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("                             GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1 ");
            sb.Append(" union all");
            //20190827  成本調整
            sb.Append("                        SELECT 'JE2' as 單別,T0.TransId+T1.Line_ID,T0.[TransId],  ");
            sb.Append("                             T1.Account 科目代號,  ");
            sb.Append("                           T1.REF1 業務員編號, T3.SlpName  姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組,");
            sb.Append("                             0 總金額,MAX(T0.REFDATE) 日期,  ");
            sb.Append("                             SUM(T1.debit)   總成本,  ");
            sb.Append("                                     SUM(T1.debit)*-1 總毛利  ");
            sb.Append("                             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("                             INNER JOIN OSLP T3 ON T1.REF1 = T3.SlpCode  ");
            sb.Append(" INNER JOIN OCRD T2 ON T1.U_REMARK1 = T2.CARDCODE");
            sb.Append("               INNER JOIN OCRG T5 ON T2.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("                             WHERE T1.ACCOUNT='51100101' and isnull(T1.REF2,'')  ='xx'");
            sb.Append("  and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112) ");
            sb.Append("                             GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1,T1.Line_ID ");
            sb.Append(" union all");
            //20171012  無成本
            sb.Append(" SELECT DISTINCT  'AR' as 單別,T2.[DocNum],T0.[TransId], ");
            sb.Append(" '41100102' 科目代號, ");
            sb.Append(" T2.SlpCode 業務員編號, T3.SlpName 姓名,MAX(SUBSTRING(GROUPNAME,4,13)) 客戶群組, ");
            sb.Append(" 0 總金額,MAX(T0.[RefDate]) 日期, ");
            sb.Append(" 0   總成本, ");
            sb.Append(" 0 總毛利 ");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId  ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE  ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE  ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T0.TransId in (338377,464446))  ");
            sb.Append(" and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112)  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];


        }


    }
}

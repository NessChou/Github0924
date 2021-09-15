using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ACME
{
    class PayFormat
    {


        public static System.Data.DataTable PackOPCH(string AA, string B1, string B2, string COM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT  'A/P號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append("              T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)SS,ISNULL(TG.dd,0)dd,ISNULL(TG.ee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,T0.COMMENTS,aa=('" + B1 + "'),bb=('" + B2 + "'),COM=('" + COM + "')");
            sb.Append("              ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,TRANSID 傳票,'' 科目名稱  FROM OPCH  T0 ");
            sb.Append("              INNER JOIN PCH1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append("              LEFT JOIN OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append("              ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM OPCH AA WHERE docentry  IN (" + AA + ")");
            sb.Append("              ) TG");
            sb.Append(" WHERE   t1.docentry  IN (" + AA + ")");


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
            return ds.Tables["OPOR"];

        }

        public static System.Data.DataTable PackOPCH2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM OPCH T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
        public static System.Data.DataTable PackOPDN(string AA, string B1, string B2, string COM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT  '收採號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append("              T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)SS,ISNULL(TG.dd,0)dd,ISNULL(TG.ee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,T0.COMMENTS,aa=('" + B1 + "'),bb=('" + B2 + "'),COM=('" + COM + "')");
            sb.Append("              ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM OPDN  T0 ");
            sb.Append("              INNER JOIN PDN1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append("              LEFT JOIN acmesql02.dbo.OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append("              ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM OPDN AA WHERE docentry  IN (" + AA + ")");
            sb.Append("              ) TG");
            sb.Append(" WHERE   t1.docentry  IN (" + AA + ")");


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
            return ds.Tables["OPOR"];

        }
        public static System.Data.DataTable PackOPDN2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PDN1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PDN1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM OPDN T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }

        public static System.Data.DataTable PackOPOR(string AA, string B1, string B2, string COM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT  '採購單號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append("              T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)SS,ISNULL(TG.dd,0)dd,ISNULL(TG.ee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,T0.COMMENTS,aa=('" + B1 + "'),bb=('" + B2 + "'),COM=('" + COM + "'),TT='請全額付清貨款，手續費進金生支付'");
            sb.Append("              ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM OPOR  T0 ");
            sb.Append("              INNER JOIN POR1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append("              LEFT JOIN acmesql02.dbo.OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append("              ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM OPOR AA WHERE docentry  IN (" + AA + ")");
            sb.Append("              ) TG");
            sb.Append(" WHERE   t1.docentry  IN (" + AA + ")");


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
            return ds.Tables["OPOR"];

        }
        public static System.Data.DataTable PackOPOR2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM POR1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM POR1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM OPOR T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
        public static System.Data.DataTable PackORIN(string AA, string B1, string B2, string COM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT  'AR貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int) 金額,t1.totalsumsy,t1.linevat,t1.totalsumsy+t1.linevat 加總,");
            sb.Append("              T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)SS,ISNULL(TG.dd,0)dd,ISNULL(TG.ee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,T0.COMMENTS,aa=('" + B1 + "'),bb=('" + B2 + "'),COM=('" + COM + "'),C1='備註:',COMMENTS C2");
            sb.Append("              ,t5.[name]  ,t0.vatsum s2,t0.doctotal-t0.vatsum s3,t0.doctotal s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM ORIN  T0 ");
            sb.Append("              INNER JOIN RIN1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append("              LEFT JOIN acmesql02.dbo.OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append("              ,(SELECT CAST(SUM(doctotal) AS INT) SS, CAST(SUM(vatsum) AS INT) dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT)) as ee   FROM ORIN AA WHERE docentry  IN (" + AA + ")");
            sb.Append("              ) TG");
            sb.Append(" WHERE   t1.docentry  IN (" + AA + ")");


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
            return ds.Tables["OPOR"];

        }
        public static System.Data.DataTable PackORIN2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0) 加總 FROM RIN1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM RIN1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM ORIN T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
        public static System.Data.DataTable PackORPC(string AA, string B1, string B2, string COM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("      SELECT  'AP貸項號碼' DOCTYPE,Convert(varchar(10),T0.[DocDate],111)  採購日期,CAST(T0.[Docnum] AS VARCHAR) 單據號碼,isnull(T0.U_SHIPPING_NO,'')+isnull(T1.U_SHIPPING_NO,'')   工單號碼,T1.ACCTCODE 科目,T1.[Dscription] 費用名稱, cast(T1.[Price] as int) 單價, cast(t0.doctotal as int)*-1 金額,t1.totalsumsy*-1 totalsumsy,t1.linevat*-1 linevat,(t1.totalsumsy+t1.linevat)*-1 加總, ");
            sb.Append("              T1.[Currency] 幣別,T0.[CardCode] 客戶編號,ISNULL(TG.SS,0)SS,ISNULL(TG.dd,0)dd,ISNULL(TG.ee,0) ee,T0.[CardName] 客戶名稱,T6.[lictradnum] 統一編號,T3.PYMNTGROUP,T0.COMMENTS,aa=('" + B1 + "'),bb=('" + B2 + "'),COM=('" + COM + "'),C1='備註:',COMMENTS C2");
            sb.Append("                                         ,t5.[name]  ,t0.doctotal*-1 s1,t0.vatsum*-1 s2,(t0.doctotal-t0.vatsum)*-1 s3,t0.doctotal*-1 s4,T7.ocrNAME 部門,t1.totalsumsy+t1.linevat 加總,'' 科目名稱 FROM ORPC  T0   ");
            sb.Append("              INNER JOIN RPC1 T1  ON T0.DocEntry = T1.DocEntry ");
            sb.Append("              INNER JOIN OCRD T2 ON T0.CARDCODE = T2.cardcode ");
            sb.Append("              LEFT JOIN OCTG T3 ON (T3.GROUPNUM=T0.GROUPNUM) ");
            sb.Append(" LEFT JOIN ohem T4 ON (T4.empid=T0.ownercode) ");
            sb.Append(" LEFT JOIN oudp T5 ON (T4.dept=T5.code)");
            sb.Append("              LEFT JOIN acmesql02.dbo.OCRD T6 ON (T0.cardcode=T6.cardcode)");
            sb.Append(" LEFT JOIN OOCR T7 ON (T1.OCRCODE=T7.OCRCODE)");
            sb.Append("              ,(SELECT CAST(SUM(doctotal) AS INT)*-1 SS, CAST(SUM(vatsum) AS INT)*-1 dd,(CAST(SUM(doctotal) AS INT)-CAST(SUM(vatsum) AS INT))*-1 as ee   FROM ORPC AA WHERE docentry  IN (" + AA + ")");
            sb.Append("              ) TG");
            sb.Append(" WHERE   t1.docentry  IN (" + AA + ")");


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
            return ds.Tables["OPOR"];

        }
        public static System.Data.DataTable PackORPC2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM RPC1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM RPC1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM ORPC T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
    }
}

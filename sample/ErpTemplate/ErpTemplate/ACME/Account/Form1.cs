using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

//Account_Temp612020 ID 10881

namespace ACME
{
    public partial class Form1 : Form
    {
        string YEAROITM = "";
        string COM = "";
        string DOCTYPE = "";
        string ss1 = "";
        string ss2 = "";
        string ss3 = "";
        string ss4 = "";
        string ss5 = "";
        string ss6 = "";
        string ss7 = "";
        string ss8 = "";
        string ss9 = "";
        string ss10 = "";
        string ss101 = "";
        string ss11 = "";
        string ss12 = "";
        string BASEDOC = "";
        string BASELINE = "";
        string FileName;
        string NewFileName;
        Double DJ, DJ2, DJ3, DJ33, HJ, SALES22 = 0;
        System.Data.DataTable dtCostDD3 = null;
        System.Data.DataTable dtCostEun = null;
        int hh, hhR, hhC, hhG = 0;
        int hS, hSR, hSC, hSG = 0;
        int hCS, hCR, hCC, hCG = 0;
        int hD, hDR, hDC, hDG = 0;
        int hE, hER, hEC, hEG = 0;
        int hF, hFR, hFC, hFG = 0;
        string A1 = "";
        string A2 = "";

        int j = 0;
        int p = 0;
        public Form1()
        {
            InitializeComponent();
        }


        private string DateToStr(DateTime Date)
        {

            return Date.ToString("yyyyMMdd");
        }





        private System.Data.DataTable GetESCO(string DocDate1, string DocDate2,int W)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT  T1.U_REMARK1  �Ȥ�s��,ISNULL(T2.CARDNAME,T1.U_REMARK1) �Ȥ�W��,ISNULL(T3.SlpName,'���i��(Airy)')  �m�W,0 �ƶq,");
            sb.Append(" 0 �渹�`���J, SUM(T1.[Debit] - T1.[Credit])   �渹�`����, T0.TransId ���J�渹, T1.Account ��إN��");
            sb.Append(" , ISNULL(MAX(T6.GROUPCODE),116)  �Ȥ�s��, SUM(T1.[Debit] - T1.[Credit]) ���ئ���,0 ���B,Convert(varchar(10),MAX(CAST(t0.REFDATE AS DATETIME)),111)  ���");
            sb.Append("                                           FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("  LEFT JOIN OCRD T6 ON T1.U_REMARK1 = T6.CARDCODE    ");
            sb.Append("                                           LEFT JOIN OSLP T3 ON T6.SLPCODE = T3.SlpCode   ");
            sb.Append("               LEFT JOIN OCRD T2 ON T1.U_REMARK1 = T2.CARDCODE ");
            sb.Append("                                           WHERE T1.ACCOUNT IN ('51101007','52200101','52200102','51101002','51101004','51101005','51100103') AND ISNULL(T1.U_REMARK1,'') <> ''");
            if (W == 7)
            {
                sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate3  and Convert(varchar(8),T0.RefDate,112) <= @DocDate4 ");
            }
            else
            {
                sb.Append("   and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112)  ");
            }
            sb.Append("                                           GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1,T6.SLPCODE,T1.U_REMARK1,T2.CARDNAME");

            sb.Append("  UNION ALL ");
            sb.Append("SELECT  T1.U_REMARK1  �Ȥ�s��,ISNULL(T2.CARDNAME,T1.U_REMARK1) �Ȥ�W��,ISNULL(T3.SlpName,'���i��(Airy)')  �m�W,0 �ƶq, ");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) �渹�`���J, 0   �渹�`����, T0.TransId ���J�渹, T1.Account ��إN�� ");
            sb.Append(", ISNULL(MAX(T6.GROUPCODE),116)  �Ȥ�s��, 0 ���ئ���, SUM(T1.[Debit] - T1.[Credit])  ���B,Convert(varchar(10),MAX(CAST(t0.REFDATE AS DATETIME)),111)  ��� ");
            sb.Append("FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("LEFT JOIN OCRD T6 ON T1.U_REMARK1 = T6.CARDCODE     ");
            sb.Append("LEFT JOIN OSLP T3 ON T6.SLPCODE = T3.SlpCode    ");
            sb.Append("LEFT JOIN OCRD T2 ON T1.U_REMARK1 = T2.CARDCODE  ");
            sb.Append("WHERE T1.ACCOUNT IN ('22610103') AND ISNULL(T1.U_REMARK1,'') <> ''   ");
            if (W == 7)
            {
                sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate3  and Convert(varchar(8),T0.RefDate,112) <= @DocDate4 ");
            }
            else
            {
                sb.Append("   and T0.[RefDate] >= Convert(varchar(8),@DocDate1,112) and T0.[RefDate] <= Convert(varchar(8),@DocDate2,112)  ");
            }
            sb.Append("                                           GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1,T6.SLPCODE,T1.U_REMARK1,T2.CARDNAME");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            string df = textBox3.Text.Substring(0, 6);
            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            command.Parameters.Add(new SqlParameter("@DocDate3", df));
            command.Parameters.Add(new SqlParameter("@DocDate4", textBox3.Text));

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

        private System.Data.DataTable GetOCRD(string DOCENTRY)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT SHIPTOCODE,CARDCODE FROM OINV WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            string df = textBox3.Text.Substring(0, 6);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable GetOCRDORIN(string DOCENTRY)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT SHIPTOCODE,CARDCODE FROM ORIN WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            string df = textBox3.Text.Substring(0, 6);
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        private System.Data.DataTable GetOCRD2(string CARDCODE, string ADDRESS)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("    SELECT U_TERRITORY FROM CRD1 WHERE ADRESTYPE='S' AND CARDCODE=@CARDCODE AND ADDRESS=@ADDRESS ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            string df = textBox3.Text.Substring(0, 6);
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@ADDRESS", ADDRESS));
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

        private System.Data.DataTable GetSAPRevenueTempLED()
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");  
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='S' ");
            sb.Append(" and (((T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%' )  and isnull(t2.u_acme_arap,'') <> 'xx' ) OR (T1.[Account]='22610103' AND (U_LOCATION)='XX' ))");
            sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");

            sb.Append(" union all");
            sb.Append(" SELECT '�U��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");  
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            //�U���A��
            sb.Append(" union all");
            sb.Append(" SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId],"); 
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%' or T1.[Account] like '4190%' or T1.[Account] like '4210%' )  ");
            sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as ��O,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T7.AcctCode)  ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append("              0 �`���B,");
            sb.Append("             0  �`����,");
            sb.Append("            0  - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");  
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append("              AND T2.[DocTotal] = 0  	AND T7.DOCENTRY <> 49540	");
            sb.Append("            GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName,T4.ocrname ");

            sb.Append(" union all");
            sb.Append("              SELECT 'AR�w' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("  Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");  
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            sb.Append("             and Convert(varchar(6),T6.DOCDATE,112) =@DocDate1  and Convert(varchar(8),T6.DOCDATE,112) <= @DocDate2 ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname");
            sb.Append(" union all");
            sb.Append("                         SELECT '�U��' as ��O,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("                        MAX(T6.AcctCode)  ��إN��,");
            sb.Append("                         T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����,");
            sb.Append("                         0 �`���B,");
            sb.Append("                              0  �`����,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("                         FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("             INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("             Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append("                         WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("  and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2 ");
            sb.Append("                         AND T2.[DocTotal] = 0 ");
            sb.Append("                       GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName,T4.ocrname ");
            sb.Append(" union all");
            //20150916 AR�w�}�U���A��
            sb.Append("            SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId], ");
            sb.Append("               T1.Account ��إN��, ");
            sb.Append("               T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����, ");
            sb.Append("               SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B, ");
            sb.Append("               SUM(T1.[Debit] - T1.[Credit])  �`����, ");
            sb.Append("               (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ��� ");
            sb.Append("               FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               INNER JOIN ORIN T2 ON T0.TransId = T2.TransId  ");
            sb.Append("               INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("                Left join oocr t4 on (t1.profitcode=t4.ocrcode)  ");
            sb.Append("               INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0 ");
            sb.Append("               LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13 ");
            sb.Append("               GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.U_ACME_ARAP=T6.BASEENTRY) ");
            sb.Append("               WHERE T2.[DocType] ='S'  AND T1.ACCOUNT='22610103' AND U_LOCATION='XX'    ");
            sb.Append("               and Convert(varchar(6),T6.DOCDATE,112) =@DocDate1  and Convert(varchar(8),T6.DOCDATE,112) <= @DocDate2");
            sb.Append("               GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname  ");
            sb.Append(" union all");
            //20151006  �����U��
            sb.Append("                        SELECT 'JE' as ��O,T0.TransId,T0.[TransId],  ");
            sb.Append("                             T1.Account ��إN��,  ");
            sb.Append("                           T1.REF1 �~�ȭ��s��, T3.SlpName  �m�W,T4.ocrname ����,");
            sb.Append("                             SUM(T1.debit)*(-1) �`���B,  ");
            sb.Append("                             0  �`����,  ");
            sb.Append("                                     SUM(T1.debit)*(-1) �`��Q,MAX(T0.REFDATE) ���  ");
            sb.Append("                             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("                             INNER JOIN OSLP T3 ON T1.REF1 = T3.SlpCode  ");
            sb.Append("   Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append("                             WHERE T1.ACCOUNT='41900101' and isnull(T1.REF2,'')  ='xx'");
            sb.Append("             and Convert(varchar(6),T0.REFDATE,112) =@DocDate1  and Convert(varchar(8),T0.REFDATE,112) <= @DocDate2");
            sb.Append("                             GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1,T4.ocrname ");
            sb.Append(" union all");
            //20190827  �����վ�
            sb.Append("                        SELECT 'JE2' as ��O,T0.TransId+T1.Line_ID,T0.[TransId],  ");
            sb.Append("                             T1.Account ��إN��,  ");
            sb.Append("                           T1.REF1 �~�ȭ��s��, T3.SlpName  �m�W,T4.ocrname ����,");
            sb.Append("                             SUM(T1.debit)�`���B,  ");
            sb.Append("                             0  �`����,  ");
            sb.Append("                                     SUM(T1.debit) �`��Q,MAX(T0.REFDATE) ���  ");
            sb.Append("                             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("                             INNER JOIN OSLP T3 ON T1.REF1 = T3.SlpCode  ");
            sb.Append("   Left join oocr t4 on (t1.profitcode=t4.ocrcode) ");
            sb.Append("                             WHERE T1.ACCOUNT='51100101' and isnull(T1.REF2,'')  ='xx'");
            sb.Append("             and Convert(varchar(6),T0.REFDATE,112) =@DocDate1  and Convert(varchar(8),T0.REFDATE,112) <= @DocDate2");
            sb.Append("                             GROUP BY T0.[TransId],T1.[Account] ,T3.SlpName,T1.REF1,T4.ocrname,T1.Line_ID ");
            sb.Append(" union all");
            //20171012  �L���J
            sb.Append(" SELECT DISTINCT 'AR' as ��O,T2.[DocNum],T0.[TransId], ");
            sb.Append(" '41100102' ��إN��, ");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,T4.ocrname ����, ");
            sb.Append(" 0 �`���B, ");
            sb.Append(" 0  �`����, ");
            sb.Append(" 0 �`��Q,MAX(T0.[RefDate]) ��� ");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId  ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append(" Left join oocr t4 on (t1.profitcode=t4.ocrcode)    ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T0.TransId IN (338377,464446))  ");
            sb.Append(" and Convert(varchar(6),T0.RefDate,112) =@DocDate1  and Convert(varchar(8),T0.RefDate,112) <= @DocDate2  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName,T4.ocrname ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            string df = textBox3.Text.Substring(0, 6);
            string df1 = textBox3.Text.Substring(4, 2);
            command.Parameters.Add(new SqlParameter("@DocDate1", df));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@aa", df1));
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
        private System.Data.DataTable GetSAPRevenueTemp3()
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account]");
            sb.Append(" ,T2.SlpCode , T3.SlpName ");
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%'  ) and isnull(t2.u_acme_arap,'') <> 'xx' ");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account]");
            sb.Append(" ,T2.SlpCode , T3.SlpName ");
            sb.Append(" union all");
            sb.Append(" SELECT '�U��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1)  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            //�U���A��
            sb.Append(" union all");
            sb.Append(" SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%'  or T1.[Account] like '4190%' or T1.[Account] like '4210%' )  ");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as ��O,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T7.AcctCode)  ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              0 �`���B,");
            sb.Append("              0  �`����,");
            sb.Append("            0  - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1) ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("            GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR�w' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            sb.Append("              and (Convert(varchar(6), DATEADD(month,1,T6.DOCDATE),112) >=@DocDate1 and  Convert(varchar(6),T6.DOCDATE,112) <=@DocDate1) ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName ");
            sb.Append(" union all");
            sb.Append("                  SELECT '�U��' as ��O,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("                 MAX(T6.AcctCode)  ��إN��,");
            sb.Append("                  T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("                  0 �`���B,");
            sb.Append("                   0 �`����,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("                  FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("                             INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("                  INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("                  WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append("              and T0.[RefDate] = '2012.03.28'");
            sb.Append(" and (Convert(varchar(6), DATEADD(month, 1, T0.RefDate),112) >=@DocDate1 and  Convert(varchar(6),T0.RefDate,112) <=@DocDate1) ");
            sb.Append("                GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", textBox6.Text));
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
        private System.Data.DataTable GetSAPRevenueTemp3q(string q, string year)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%'  )  and isnull(t2.u_acme_arap,'') <> 'xx' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append(" SELECT '�U��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX'  and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%')  ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");
            //�U���A��
            sb.Append(" union all");
            sb.Append(" SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%' or T1.[Account] like '4190%' or T1.[Account] like '4210%' )  ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as ��O,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T7.AcctCode)  ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              0 �`���B,");
            sb.Append("              0  �`����,");
            sb.Append("            0  - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T0.RefDate,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("            GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode , T3.SlpName ");

            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");

            sb.Append(" union all");
            sb.Append("                        SELECT '�U��' as ��O,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("                       MAX(T6.AcctCode)  ��إN��,");
            sb.Append("                        T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("                        0 �`���B,");
            sb.Append("                       0  �`����,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("                        FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId");
            sb.Append("                         INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("                        INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("                        WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101')");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year2+'10' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'01' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'04' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),T6.DOCDATE,112)  between @year+'07' and @year+'12') ");
            }
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("                        GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode , T3.SlpName ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            int f = Convert.ToInt32(year);
            string year2 = Convert.ToString(f - 1);
            command.Parameters.Add(new SqlParameter("@year", year));
            command.Parameters.Add(new SqlParameter("@year2", year2));
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


        private System.Data.DataTable GetSAPRevenueTemp3y()
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%' ) AND  isnull(t2.u_acme_arap,'') <> 'xx' ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append(" SELECT '�U��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='I'  AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2)  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");
            sb.Append(" union all");

            sb.Append(" SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%' or T1.[Account] like '4190%' or T1.[Account] like '4210%' )  ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2) ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");

            sb.Append(" union all");
            sb.Append("              SELECT 'AR' as ��O,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T7.AcctCode)  ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              0 �`���B,");
            sb.Append("              0  �`����,");
            sb.Append("            0  - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("             inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2) ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("            GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR�w' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            sb.Append("              and (Convert(varchar(4),T6.DOCDATE,112) >=@DocDate1 and  Convert(varchar(4),T6.DOCDATE,112) <=@DocDate2) ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName");
            sb.Append(" union all");
            sb.Append("                     SELECT '�U��' as ��O,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("                    MAX(T6.AcctCode)  ��إN��,");
            sb.Append("                     T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,");
            sb.Append("                     0 �`���B,");
            sb.Append("                      0 �`����,");
            sb.Append("             0-SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("                     FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("                               INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append("                     INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("                     WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append(" and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2) ");
            sb.Append("                     AND T2.[DocTotal] = 0 ");
            sb.Append("                   GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            sb.Append("               SELECT DISTINCT 'AR' as ��O,T2.[DocNum],T0.[TransId], ");
            sb.Append("               '41100101' ��إN��, ");
            sb.Append("               T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,0 �`���B,0 �`����, ");
            sb.Append("             0  �`��Q,MAX(T0.[RefDate]) ��� ");
            sb.Append("               FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               INNER JOIN OINV T2 ON T0.TransId = T2.TransId  ");
            sb.Append("               INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("               WHERE T2.[DocType] ='I' and (T0.TransId IN (338377,464446))  ");
            sb.Append("              and (Convert(varchar(4),T0.RefDate,112) >=@DocDate1 and  Convert(varchar(4),T0.RefDate,112) <=@DocDate2)  ");
            sb.Append("               GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode,T3.SlpName ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            string f = textBox6.Text.Substring(0, 4);
            int g = Convert.ToInt16(f) - 1;
            string f2 = g.ToString();
            command.Parameters.Add(new SqlParameter("@DocDate1", f2));
            command.Parameters.Add(new SqlParameter("@DocDate2", f));
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
        private System.Data.DataTable GetSAPRevenueear(string dd)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT 'AR' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  ) ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            //AR�A��
            sb.Append(" union all");
            sb.Append(" SELECT 'AR-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit]) * ( -1 )   �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%'  or T1.[Account] like '4210%' )  and isnull(t2.u_acme_arap,'') <> 'xx' ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            sb.Append(" SELECT '�U��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' AND ISNULL(U_ACME_SERIAL,'') <> 'XX' and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4190%') ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  ");
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            //�U���A��
            sb.Append(" union all");
            sb.Append(" SELECT '�U��-�A��' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append(" T1.Account ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,");
            sb.Append(" SUM(T2.[DocTotal] - T2.[VatSum])*(-1) �`���B,");
            sb.Append(" SUM(T1.[Debit] - T1.[Credit])  �`����,");
            sb.Append(" (SUM(T2.[DocTotal] - T2.[VatSum])*(-1)) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='S' AND ISNULL(U_ACME_SERIAL,'') <> 'XX'  and (T1.[Account] like '4110%' or T1.[Account] like '4170%' or T1.[Account] like '4270%'  or T1.[Account] like '4190%' or T1.[Account] like '4210%'  )  ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  "); 
            sb.Append(" GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode , T3.SlpName");

            sb.Append(" union all");
            sb.Append(" SELECT 'AR' as ��O,T7.DOCENTRY DocNum,T0.[TransId],");
            sb.Append(" MAX(T7.AcctCode)  ��إN��,");
            sb.Append(" T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,0 �`���B,");
            sb.Append(" 0  �`����,");
            sb.Append(" 0  - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" INNER JOIN ODLN T2 ON T0.TransId = T2.TransId ");
            sb.Append(" INNER JOIN DLN1 T6 ON T2.DOCENTRY = T6.DOCENTRY ");
            sb.Append(" inner join inv1 T7 on (T7.baseentry=T6.docentry and  T7.baseline=T6.linenum and T7.basetype='15'  )");
            sb.Append(" INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append(" INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append(" INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  ");
            sb.Append("  AND T2.[DocTotal] = 0 ");
            sb.Append(" GROUP BY T7.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");
            sb.Append("              SELECT 'AR�w' as ��O,T2.[DocNum],T0.[TransId],");
            sb.Append("              T1.Account ��إN��,");
            sb.Append("              T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) �`���B,");
            sb.Append("              SUM(T1.[Debit] - T1.[Credit]) * ( -1 )  �`����,");
            sb.Append("              SUM(T2.[DocTotal] - T2.[VatSum]) - SUM(T1.[Debit] - T1.[Credit]) �`��Q,MAX(T6.[DOCDATE]) ���");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode ");
            sb.Append("              INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("              INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append(" INNER JOIN (SELECT DISTINCT BASEENTRY,T1.DOCDATE FROM DLN1 T0");
            sb.Append(" LEFT JOIN ODLN T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE T0.BASETYPE=13");
            sb.Append(" GROUP BY BASEENTRY,T1.DOCDATE) T6 ON (T2.DOCENTRY=T6.BASEENTRY)");
            sb.Append("              WHERE T2.[DocType] ='I' and T1.[Account] IN ('22610103','41100101','42100101') ");
            sb.Append("              and Convert(varchar(4),T6.DOCDATE,112) =@DocDate1  ");
            sb.Append("              GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName");
            sb.Append(" union all");
            sb.Append("             SELECT '�U��' as ��O,T2.DOCENTRY DocNum,T0.[TransId],");
            sb.Append("             MAX(T6.AcctCode)  ��إN��,");
            sb.Append("             T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��,0 �`���B,");
            sb.Append("             0  �`����,");
            sb.Append("              0-SUM(T1.[Debit] - T1.[Credit])  �`��Q,MAX(T0.[RefDate]) ���");
            sb.Append("             FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("            INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("             INNER JOIN RIN1 T6 ON T2.DOCENTRY = T6.DOCENTRY");
            sb.Append("             INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode   ");
            sb.Append("             INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE ");
            sb.Append("             INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE ");
            sb.Append("             WHERE T2.[DocType] ='I' and (T1.[Account] = '51100101') ");
            sb.Append(" and Convert(varchar(4),T0.RefDate,112) =@DocDate1  ");
            sb.Append("              AND T2.[DocTotal] = 0 ");
            sb.Append("             GROUP BY T2.DOCENTRY,T0.[TransId],T2.SlpCode,T3.SlpName ");
            sb.Append(" union all");

            //�L���J
            sb.Append("               SELECT DISTINCT 'AR' as ��O,T2.[DocNum],T0.[TransId], ");
            sb.Append("                   '41100101' ��إN��, ");
            sb.Append("               T2.SlpCode �~�ȭ��s��, T3.SlpName �m�W,MAX(SUBSTRING(GROUPNAME,4,13)) �Ȥ�s��, ");
            sb.Append("               0 �`���B, ");
            sb.Append("               0 �`����, ");
            sb.Append("               0 �`��Q,MAX(T0.[RefDate]) ��� ");
            sb.Append("               FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId  ");
            sb.Append("               INNER JOIN OINV T2 ON T0.TransId = T2.TransId  ");
            sb.Append("               INNER JOIN OSLP T3 ON T2.SlpCode = T3.SlpCode  ");
            sb.Append("               INNER JOIN OCRD T4 ON T2.CARDCODE = T4.CARDCODE  ");
            sb.Append("               INNER JOIN OCRG T5 ON T4.GROUPCODE = T5.GROUPCODE  ");
            sb.Append("               WHERE T2.[DocType] ='I' and (T0.TransId IN (338377,464446))  ");
            sb.Append("               and Convert(varchar(4),T0.RefDate,112) =@DocDate1   ");
            sb.Append("               GROUP BY T2.[DocNum],T0.[TransId],T1.[Account],T2.SlpCode ,T3.SlpName ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocDate1", dd));



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
        private System.Data.DataTable GetSAPDoc(string DocKind, Int32 DocNum, string AcctCode,string A1,string A2)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "��f")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");


            }
            else if (DocKind == "�P�h")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "�U��")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                //201810                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   (UpdInvnt='I' ) ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   (UpdInvnt='I'  OR  T1.DocEntry ='44955' ) ");



            }
            else if (DocKind == "AR�w")
            {

                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 ");
                sb.Append("INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry ");
                sb.Append("INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)");
                sb.Append("LEFT JOIN ODLN T4 ON (T3.DOCENTRY=T4.DOCENTRY)");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   T1.UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0 and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2     ");


            }
            else if (DocKind == "JE")
            {

                sb.Append("                  SELECT  T0.U_REMARK1 as CardCode, T2.CARDNAME as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price],  ");
                sb.Append("                    T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice],'103' GROUPCODE FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID   ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.U_REMARK1 = T2.CARDCODE");
                sb.Append(" WHERE T1.TransID =@DocEntry   AND T0.REF2='XX'  ");

            }
            else if (DocKind == "JE2")
            {

                sb.Append("                  SELECT  T0.U_REMARK1 as CardCode, T2.CARDNAME as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price],  ");
                sb.Append("                   0 as  [LineTotal],T0.[Debit] - T0.[Credit]  as [StockPrice],'103' GROUPCODE FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID   ");
                sb.Append(" INNER JOIN OCRD T2 ON T0.U_REMARK1 = T2.CARDCODE");
                sb.Append(" WHERE T1.TransID+T0.Line_ID =@DocEntry   AND T0.REF2='XX'  ");

            }
            else if (DocKind == "�U��-�A��")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //�[�J��¦�渹 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR-�A��")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                //�[�J��¦�渹 
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");

                sb.Append("T0.[LineTotal], T0.[StockPrice],T2.GROUPCODE FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  INNER JOIN OCRD T2 ON T1.CARDCODE = T2.CARDCODE ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            if (DocKind != "JE" && DocKind != "JE2")
            {
                //20081009 �W�[ ��إN��
                sb.Append("AND  T0.AcctCode =@AcctCode   ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            //20081009 �W�[ ��إN��
            command.Parameters.Add(new SqlParameter("@AcctCode", AcctCode));
            command.Parameters.Add(new SqlParameter("@A1", A1));
            command.Parameters.Add(new SqlParameter("@A2", A2));
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


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }

      


        //���o����
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum, Int32 LineNum)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "��f")
            {

                sb.Append(" SELECT LINENUM,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append(" T0.[LineTotal], T0.[StockPrice],T2.�`���� ");
                sb.Append(" FROM DLN1 T0 INNER JOIN ODLN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append(" INNER JOIN (SELECT SUM([Debit]-[Credit]) �`����,TransId FROM JDT1 WHERE [Account]='51100101' GROUP BY TransId) T2 ");
                sb.Append(" ON(T1.TransId=T2.TransId)");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");
                sb.Append("and   T0.LineNum =@LineNum   ");



            }
            else if (DocKind == "�P�h")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RDN1 T0 INNER JOIN ORDN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "�U��")
            {

                sb.Append("SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");

                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }
            else if (DocKind == "AR")
            {

                sb.Append(" SELECT * FROM (SELECT T0.ACCTCODE,SELECT T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[BaseEntry], T0.[BaseLine],");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   AND   UpdInvnt='I'   ");
                sb.Append("UNION ALL   ");
             

            }
            else if (DocKind == "AR�w")
            {



                sb.Append("SELECT T0.ACCTCODE,T1.[CardCode], T1.[CardName], T0.[ItemCode], T0.[Dscription], T3.[Quantity], T0.[Price], ");
                sb.Append("T3.DOCENTRY AS BaseEntry,T3.LINENUM AS BaseLine,");
                sb.Append("T3.[LineTotal], T0.[StockPrice] FROM INV1 T0 INNER JOIN OINV T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("LEFT JOIN DLN1 T3 ON (T3.BASEENTRY=T0.DOCENTRY AND T3.BASELINE=T0.LINENUM)     LEFT JOIN ODLN T4 ON (T3.DOCENTRY=T4.DOCENTRY) ");
                sb.Append(" WHERE T1.DocEntry =@DocEntry   AND   T1.UpdInvnt='C' and  ISNULL(T3.DOCENTRY,0) <> 0   and T3.BASETYPE=13 AND T4.DOCDATE BETWEEN @A1 AND @A2  ");
            }
            else if (DocKind == "JE")
            {

                sb.Append("SELECT  T0.Account as CardCode, T0.LineMemo as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price], ");
                sb.Append("  T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice] FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID  ");
                sb.Append("WHERE T1.TransID =@DocEntry   ");

            }
            else if (DocKind == "JE2")
            {

                sb.Append("SELECT  T0.Account as CardCode, T0.LineMemo as CardName, '' as [ItemCode], '' as [Dscription], 0 as [Quantity],  0 as [Price], ");
                sb.Append("  T0.[Debit] - T0.[Credit]  as  [LineTotal], 0 as [StockPrice] FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.TransID = T1.TransID  ");
                sb.Append("WHERE T1.TransID+T0.Line_ID =@DocEntry   ");

            }
            else if (DocKind == "�U��-�A��")
            {

                sb.Append("SELECT T1.[CardCode],T1.[CardName],T0.AcctCode as ItemCode, T0.[Dscription], T0.[Quantity], T0.[Price], ");
                sb.Append("T0.[LineTotal], T0.[StockPrice] FROM RIN1 T0 INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry  ");
                sb.Append("WHERE T1.DocEntry =@DocEntry   ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));

            command.Parameters.Add(new SqlParameter("@A1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@A2", textBox2.Text));
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


        //����ڥ����ҳs�����������B
        private System.Data.DataTable GetSAPDocByLine(string DocKind, Int32 DocNum)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();


            if (DocKind == "��f")
            {


            }
            else if (DocKind == "�P�h")
            {


            }
            else if (DocKind == "�U��")
            {

                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) �`���� ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");


            }
            else if (DocKind == "AR" || DocKind == "AR-�A��" | DocKind == "AR�w")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) �`���� ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");



            }
            else if (DocKind == "JE")
            {

            }
            else if (DocKind == "JE2")
            {

            }
            else if (DocKind == "�U��-�A��")
            {
                sb.Append(" SELECT SUM(T1.[Debit] - T1.[Credit]) �`���� ");
                sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
                sb.Append(" INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
                sb.Append(" WHERE T2.DocEntry =@DocEntry  AND T1.[Account] ='51100101' ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //

            command.Parameters.Add(new SqlParameter("@DocEntry", DocNum));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OINV");
                connection.Close();
            }
            finally
            {
                connection.Close();
            }


            //bindingSource2.DataSource = ds.Tables[0];
            //dataGridView7.DataSource = bindingSource2;

            return ds.Tables[0];


        }


        //�ʺA���͸�Ƶ��c
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("�渹", typeof(Int32));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���~�W��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(Int32));
            dt.Columns.Add("���", typeof(decimal));
            dt.Columns.Add("���B", typeof(Int32));
            dt.Columns.Add("���ئ���", typeof(Int32));
            dt.Columns.Add("�渹�`����", typeof(Int32));
            return dt;
        }


        private System.Data.DataTable MakeTableOCRD()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�P�����B", typeof(decimal));
            dt.Columns.Add("�P������", typeof(decimal));
            dt.Columns.Add("�ƶq", typeof(decimal));
            dt.Columns.Add("�Q��", typeof(decimal));
            dt.Columns.Add("�Q���", typeof(string));
            dt.Columns.Add("COM", typeof(string));
            dt.Columns.Add("��a", typeof(string));
            return dt;
        }

        private System.Data.DataTable MakeTableOCRDACC()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�P�����B", typeof(decimal));
            dt.Columns.Add("�P������", typeof(decimal));
            dt.Columns.Add("�ƶq", typeof(decimal));
            dt.Columns.Add("�Q��", typeof(decimal));
            dt.Columns.Add("�Q���", typeof(string));
            dt.Columns.Add("��إN��", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableOCRDF()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�P�����B", typeof(decimal));
            dt.Columns.Add("�P������", typeof(decimal));
            dt.Columns.Add("�ƶq", typeof(decimal));
            dt.Columns.Add("�Q��", typeof(decimal));
            dt.Columns.Add("�Q���", typeof(string));
            dt.Columns.Add("COM", typeof(string));
            dt.Columns.Add("��a", typeof(string));
            dt.Columns.Add("SHIPTO", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableRevenue()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("�渹", typeof(Int32));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���~�W��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(Int32));
            dt.Columns.Add("���", typeof(decimal));
            dt.Columns.Add("���B", typeof(Int32));
            dt.Columns.Add("���ئ���", typeof(Int32));
            dt.Columns.Add("�渹�`����", typeof(Int32));
            dt.Columns.Add("��¦�渹", typeof(Int32));
            dt.Columns.Add("��¦�C", typeof(Int32));
            return dt;
        }


        //�ʺA���͸�Ƶ��c
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���J���", typeof(string));
            dt.Columns.Add("���J�渹", typeof(Int32));

            dt.Columns.Add("�������", typeof(string));
            dt.Columns.Add("�����渹", typeof(Int32));

            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));

            //20081008
            //�~�ȭ�
            dt.Columns.Add("�~�ȭ��s��", typeof(string));
            dt.Columns.Add("�m�W", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));

            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���~�W��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(Int32));
            dt.Columns.Add("���", typeof(decimal));
            dt.Columns.Add("���B", typeof(Int32));

            dt.Columns.Add("���ئ���", typeof(Int32));   //�������ɼg�J�����
            dt.Columns.Add("�渹�`����", typeof(Int32)); //�������ɼg�J�����

            dt.Columns.Add("�渹�`���J", typeof(Int32));
            dt.Columns.Add("��¦�渹", typeof(Int32));
            dt.Columns.Add("��¦�C", typeof(Int32));





            return dt;
        }
        //�ʺA���͸�Ƶ��c
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���J���", typeof(string));
            dt.Columns.Add("���J�渹", typeof(Int32));

            dt.Columns.Add("�������", typeof(string));
            dt.Columns.Add("�����渹", typeof(Int32));

            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));

            //20081008
            //�~�ȭ�   
            dt.Columns.Add("�~�ȭ��s��", typeof(string));
            dt.Columns.Add("�m�W", typeof(string));


            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���~�W��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(Int32));
            dt.Columns.Add("���", typeof(decimal));
            dt.Columns.Add("���B", typeof(Int32));

            dt.Columns.Add("���ئ���", typeof(Int32));   //�������ɼg�J�����
            dt.Columns.Add("�渹�`����", typeof(Int32)); //�������ɼg�J�����

            dt.Columns.Add("�渹�`���J", typeof(Int32));
            dt.Columns.Add("��¦�渹", typeof(Int32));
            dt.Columns.Add("��¦�C", typeof(Int32));

            dt.Columns.Add("���", typeof(DateTime));
            dt.Columns.Add("��إN��", typeof(string));




            return dt;
        }

        private void button3_Click(object sender, EventArgs e)
        {


            System.Data.DataTable dt = GetMenu.GetSAPRevenue(textBox1.Text, textBox2.Text);

            System.Data.DataTable dtCost = MakeTableCombine();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string ���;
            string ��إN��;

            Int32 �渹;


            Int32 ��¦�渹;
            Int32 ��¦�C;

            //20080904
            //�ŧi DuplicateKey ���ˬd
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                ��� = Convert.ToString(dt.Rows[i]["��O"]);
                �渹 = Convert.ToInt32(dt.Rows[i]["DocNum"]);

                ��إN�� = Convert.ToString(dt.Rows[i]["��إN��"]);

                dtDoc = GetSAPDoc(���, �渹, ��إN��, textBox1.Text, textBox2.Text);



                ��¦�渹 = -1;
                ��¦�C = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();


                    dr["���J���"] = ���;
                    dr["���J�渹"] = �渹;


                    dr["�Ȥ�s��"] = "'"+Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["�Ȥ�W��"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["���~�s��"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["���~�W��"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);
                    dr["�ƶq"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["���"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["���B"] = Convert.ToInt64(dtDoc.Rows[j]["LineTotal"]);
                    dr["�渹�`����"] = 0;
                    dr["���ئ���"] = 0;


       

                    //20081008
                    //�~�ȭ�
                    dr["�~�ȭ��s��"] = Convert.ToString(dt.Rows[i]["�~�ȭ��s��"]);
                    dr["�m�W"] = Convert.ToString(dt.Rows[i]["�m�W"]);
                    dr["�Ȥ�s��"] = Convert.ToString(dt.Rows[i]["�Ȥ�s��"]);

                    if (��� == "AR" || ��� == "�U��" || ��� == "AR�w")
                    {



                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ��¦�渹 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["��¦�渹"] = ��¦�渹;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            ��¦�C = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["��¦�C"] = ��¦�C;
                        }

                    }

                    //�`���J�g�b�̫�@��
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (��� == "AR" || ��� == "AR-�A��" || ��� == "AR�w")
                        {
                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]);


                        }
                        else
                        {
                           
                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]) * (-1);
                        }
                    }

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  �ƶq�令 �t��
                        dr["�ƶq"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }


                    //�̾�  ��¦�渹 & ��¦�C ���o����
                    //�p�G��ڥ����S����¦�渹 & ��¦�C�N�{�C����

                    //20080916 AR ������ �y�� ������|
                    if (��� == "AR" || ��� == "AR�w")
                    {
                        //0303
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (��¦�渹.ToString() == "3169" && �渹.ToString() == "3429")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (��¦�渹.ToString() == "3167" && �渹.ToString() == "3404")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }

                            
                            dtDocLine = GetSAPDocByLine("��f", ��¦�渹, ��¦�C);

                            dr["�������"] = "��f";
                            dr["�����渹"] = ��¦�渹;

                            if (dtDocLine.Rows.Count == 1)
                            {
                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                                                  * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["�渹�`����"] = 0;
                                    if (�渹 != DuplicateKey)
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }
                                //20091204�@��h
                                if (��¦�渹.ToString() == "5394" && �渹.ToString() == "5673")
                                {

                                    dr["�渹�`����"] = 2111964;
                                    dr["�渹�`���J"] = 0;

                                }
                                //2010331�h��@
                                if (�渹.ToString() == "6975")
                                {
                                    dr["�渹�`����"] = 0;

                                }
                                //2010409�q����AR
                                if (�渹.ToString() == "7022")
                                {
                                    dr["�渹�`����"] = "5476";
                                    dr["���ئ���"] = "5476";

                                }
                                //20150506 AR��AR�w�@�s
                                if (��¦�渹.ToString() == "26223" && ��¦�C.ToString() == "0")
                                {
                                    dr["�渹�`����"] = "1005608";

                                }
                                //20111102 �Ӷ����f�h��1
                                System.Data.DataTable GT = TF(��¦�渹.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["�Ǹ�"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (�渹.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["�渹�`����"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                            //���������Ӧۦܩ����

                        }
                        //�S����¦�渹
                        else
                        {
                            //������Ƭ��ۤw
                            dr["�������"] = ���;
                            dr["�����渹"] = �渹;


                            dtDocLine = GetSAPDocByLine(���, �渹);

                            if (dtDocLine != null)
                            {

                                if (dtDocLine.Rows.Count == 1)
                                {
                                    dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                                   * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                        {
                                            dr["�渹�`����"] = 0;
                                        }
                                        else
                                        {
                                            //�Ϧ^�h��P�f����
                                            System.Data.DataTable dtSalesCost = GetSalesCost(�渹.ToString());
                                            try
                                            {
                                                dr["�渹�`����"] = Convert.ToInt32(dtSalesCost.Rows[0]["�`����"]);
                                            }
                                            catch
                                            {
                                                dr["�渹�`����"] = 0;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    //Rows.Count =0 �������s
                                    dr["���ئ���"] = 0;
                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                        }

                    }

                    // 3 ��רҨS���ӷ��渹

                    //20081007 �W�[�P�h..�������t

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////�n�P�_�ӷ������
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));

                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;

                                    //20081231
                                    if (�渹 != DuplicateKey)
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }
                            }


                        }
                        else
                        {
                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                        // dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }
                        }
                    }
                    string g1 = dr["�渹�`���J"].ToString();
                    dtCost.Rows.Add(dr);

                }
            }

            dtCost.DefaultView.Sort = "�Ȥ�s�� DESC";
            bindingSource1.DataSource = dtCost;
            dataGridView8.DataSource = bindingSource1.DataSource;

            if (checkBox1.Checked)
            {
                ACME.Form1Rpt31 frm = new ACME.Form1Rpt31();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            if (checkBox2.Checked)
            {
                ACME.Form1Rpt3c frm = new ACME.Form1Rpt3c();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            else
            {

                ACME.Form1Rpt3 frm = new ACME.Form1Rpt3();
                frm.dt = dtCost;
                frm.ShowDialog();

            }



        }
        private System.Data.DataTable MakeTableOrder_Item()
        {
            System.Data.DataTable dt = new System.Data.DataTable();



            //20081008
            dt.Columns.Add("���~�s��", typeof(string));
            dt.Columns.Add("���~�W��", typeof(string));

            dt.Columns.Add("�P�����B", typeof(Int32));
            dt.Columns.Add("�P������", typeof(Int32));
            dt.Columns.Add("�Q��", typeof(Int32));
            dt.Columns.Add("�Q���", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["���~�s��"];
            dt.PrimaryKey = colPk;


            //�g�J���
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "�q��i��";
            //dt.Rows.Add(dr);


            return dt;
        }
        private System.Data.DataTable MakeTableOrder_Sales()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            //dt.Columns.Add("�Ȥ�s��", typeof(string));
            //dt.Columns.Add("�Ȥ�W��", typeof(string));
            //20081008
            dt.Columns.Add("�~�ȭ��s��", typeof(string));
            dt.Columns.Add("�m�W", typeof(string));

            dt.Columns.Add("�P�����B", typeof(Int32));
            dt.Columns.Add("�P������", typeof(Int32));
            dt.Columns.Add("�Q��", typeof(Int32));
            dt.Columns.Add("�Q���", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["�~�ȭ��s��"];
            dt.PrimaryKey = colPk;


            //�g�J���
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "�q��i��";
            //dt.Rows.Add(dr);


            return dt;
        }



   

        //���o����
        private System.Data.DataTable GetSalesCost(string BaseRef)
        {
            //�X�p AS �P����B
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT (T1.[Debit] - T1.[Credit])  �`����");
            sb.Append(" FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append(" WHERE T0.TransType=13 and  T0.BaseRef=@BaseRef and T1.[Account] like '5110%' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            //
            command.Parameters.Add(new SqlParameter("@BaseRef", BaseRef));

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

        private void button5_Click_1(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            string YEAR = A1.Substring(0, 4);
            int M1 = Convert.ToInt16(A1.Substring(4, 2));
            int M2 = Convert.ToInt16(A2.Substring(4, 2));

            if (radioButton1.Checked)
            {
                System.Data.DataTable DT1 = null;
    
                if (M2 - M1 != 0 && (globals.DBNAME != "�F�ͥ�"))
                {
                    DT1 = GetItemmDD(COM, "Account_Temp6" + YEAR);                    
                }
                else
                {
                    Category(8, "", "Account_Temp6");
                    DT1 = GetItemmDD(COM, "Account_Temp6");
                }
                System.Data.DataTable dtCost = MakeTableOCRD();
                DataRow dr = null;
   
                for (int i = 0; i <= DT1.Rows.Count - 1; i++)
                {
                    DataRow dd = DT1.Rows[i];
                    dr = dtCost.NewRow();
                    string CARDCODE = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�s��"] = CARDCODE;
                    dr["�Ȥ�s��"] = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�W��"] = dd["�Ȥ�W��"].ToString();
                    dr["�ƶq"] = Convert.ToDecimal(dd["�ƶq"]);
                    dr["�P�����B"] = Convert.ToDecimal(dd["�P�����B"]);
                    dr["�P������"] = Convert.ToDecimal(dd["�P������"]);
                    dr["�Q��"] = Convert.ToDecimal(dd["�Q��"]);
                    dr["�Q���"] = dd["�Q���"].ToString();
                    dr["COM"] = dd["COM"].ToString();
                    StringBuilder sb2 = new StringBuilder();
                    System.Data.DataTable DT2 = GetItemmDDOCRD(CARDCODE, "Account_Temp61");
                    if (DT2.Rows.Count > 0)
                    {
                        for (int il = 0; il <= DT2.Rows.Count - 1; il++)
                        {
                            string COUNTRY = DT2.Rows[il]["COUNTRY"].ToString();
                            sb2.Append(COUNTRY + "/");

                        }
                        sb2.Remove(sb2.Length - 1, 1);
                        dr["��a"] = sb2.ToString();
                    }
                    dtCost.Rows.Add(dr);
                }
                ACME.Form1Rpt4 frm4 = new ACME.Form1Rpt4();
                frm4.dt = dtCost;
                frm4.s = textBox1.Text;
                frm4.q = textBox2.Text;
                frm4.ShowDialog();
            }
            else if (radioButton2.Checked)
            {

                System.Data.DataTable G1 = null;
                if (M2 - M1 != 0 && (globals.DBNAME != "�F�ͥ�"))
                {

                    G1 = GetItemmDDS("Account_Temp6" + YEAR);
                    
                }
                else
                {
                    Category(8, "", "Account_Temp6");

                    G1 = GetItemmDDS("Account_Temp6");
                }


                try
                {

                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\ACC\\�Ȥ�h�I�ڤ覡.xls";


                    //Excel���˪���
                    string ExcelTemplate = FileName;

                    //��X��
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                          DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                    ExcelReport.ExcelReportOutput(G1, ExcelTemplate, OutPutFile, "N");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
            else if (radioButton3.Checked)
            {
                ViewBatchPayment3();

            }
            else if (radioButton6.Checked)
            {
                A1 = textBox1.Text.Substring(0, 4) + "0101";
                A2 = textBox1.Text.Substring(0, 4) + "1231";
                Category(8, "", "Account_Temp6");
                System.Data.DataTable T1 = GetItemmDDQ();

                try
                {

                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\ACC\\�Ȥ�h�u.xls";


                    //Excel���˪���
                    string ExcelTemplate = FileName;

                    //��X��
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                          DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                    ExcelReport.ExcelReportOutput(T1, ExcelTemplate, OutPutFile, "N");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetMenu.GetSAPRevenue(textBox1.Text, textBox2.Text);

            System.Data.DataTable dtCost = MakeTableCombine();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string ���;
            string ��إN��;

            Int32 �渹;


            Int32 ��¦�渹;
            Int32 ��¦�C;

            //20080904
            //�ŧi DuplicateKey ���ˬd
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                ��� = Convert.ToString(dt.Rows[i]["��O"]);
                �渹 = Convert.ToInt32(dt.Rows[i]["DocNum"]);


                ��إN�� = Convert.ToString(dt.Rows[i]["��إN��"]);


                //if (�渹 == 1756)
                //{
                //    MessageBox.Show("");
                //}

                //20080904 �W�C�קK�P�f�������л{�C
                //�@�i�榳�h�ؾP�f���J,�P�f�����u���@��
                //�@�k:

                dtDoc = GetSAPDoc(���, �渹, ��إN��, textBox1.Text, textBox2.Text);


                ��¦�渹 = -1;
                ��¦�C = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();

                    //  dr["��إN��"] = ��إN��;
                    dr["���J���"] = ���;
                    dr["���J�渹"] = �渹;


                    dr["�Ȥ�s��"] = "'"+Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["�Ȥ�W��"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["���~�s��"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["���~�W��"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);
                    dr["�ƶq"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["���"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]);
                    dr["�渹�`����"] = 0;
                    dr["���ئ���"] = 0;


                    //20081008
                    //�~�ȭ�
                    dr["�~�ȭ��s��"] = Convert.ToString(dt.Rows[i]["�~�ȭ��s��"]);
                    dr["�m�W"] = Convert.ToString(dt.Rows[i]["�m�W"]);
                    dr["�Ȥ�s��"] = Convert.ToString(dt.Rows[i]["�Ȥ�s��"]);

                    if (��� == "AR" || ��� == "�U��" || ��� == "AR�w")
                    {



                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ��¦�渹 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["��¦�渹"] = ��¦�渹;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            ��¦�C = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["��¦�C"] = ��¦�C;
                        }

                    }

                    //�`���J�g�b�̫�@��
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (��� == "AR" || ��� == "AR-�A��" || ��� == "AR�w")
                        {
                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]);
                        }
                        else
                        {

                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]) * (-1);
                        }
                    }

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  �ƶq�令 �t��
                        dr["�ƶq"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }


                    //�̾�  ��¦�渹 & ��¦�C ���o����
                    //�p�G��ڥ����S����¦�渹 & ��¦�C�N�{�C����

                    //20080916 AR ������ �y�� ������|
                    if (��� == "AR" || ��� == "AR�w")
                    {


                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (��¦�渹.ToString() == "3169" && �渹.ToString() == "3429")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (��¦�渹.ToString() == "3167" && �渹.ToString() == "3404")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            dtDocLine = GetSAPDocByLine("��f", ��¦�渹, ��¦�C);

                            dr["�������"] = "��f";
                            dr["�����渹"] = ��¦�渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                               * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["�渹�`����"] = 0;
                                    if (�渹 != DuplicateKey)
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }
                                //20091204�@��h
                                if (��¦�渹.ToString() == "5394" && �渹.ToString() == "5673")
                                {

                                    dr["�渹�`����"] = 2111964;
                                    dr["�渹�`���J"] = 0;
                                }

                                //2010331�h��@
                                if (�渹.ToString() == "6975")
                                {
                                    dr["�渹�`����"] = 0;

                                }

                                //2010409�q����AR
                                if (�渹.ToString() == "7022")
                                {
                                    dr["�渹�`����"] = "5476";
                                    dr["���ئ���"] = "5476";
                                }

                                //20150506 AR��AR�w�@�s
                                if (��¦�渹.ToString() == "26223" && ��¦�C.ToString() == "0")
                                {
                                    dr["�渹�`����"] = "1005608";


                                }

                                //20111102 �Ӷ����f�h��1
                                System.Data.DataTable GT = TF(��¦�渹.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["�Ǹ�"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (�渹.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["�渹�`����"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                            //���������Ӧۦܩ����

                        }
                        //�S����¦�渹
                        else
                        {
                            //������Ƭ��ۤw
                            dr["�������"] = ���;
                            dr["�����渹"] = �渹;
                            dtDocLine = GetSAPDocByLine(���, �渹);
                            if (dtDocLine.Rows.Count == 1)
                            {
                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            //�Ϧ^�h��P�f����
                                            System.Data.DataTable dtSalesCost = GetSalesCost(�渹.ToString());
                                            //111
                                            dr["�渹�`����"] = Convert.ToInt32(dtSalesCost.Rows[0]["�`����"]);
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                        }

                    }

                    // 3 ��רҨS���ӷ��渹

                    //20081007 �W�[�P�h..�������t

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////�n�P�_�ӷ������
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));

                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }


                        }
                        else
                        {
                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);
                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                        // dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }
                        }
                    }

                    dtCost.Rows.Add(dr);

                }
            }


            dtCost.DefaultView.Sort = "�~�ȭ��s�� DESC";
            //bindingSource1.DataSource = dtCost;
            //dataGridView8.DataSource = bindingSource1.DataSource;


            //20081008 ���ʳ���
            //�Ȥ�O



            if (checkBox2.Checked)
            {
                ACME.Form1Salesc frm = new ACME.Form1Salesc();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            else
            {

                ACME.Form1Sales frm = new ACME.Form1Sales();
                frm.dt = dtCost;
                frm.ShowDialog();
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //20080904 �[�J�S�������P�f���J���P�f����
            //�p�G�O��f�N�P�h...
            System.Data.DataTable dt = GetMenu.GetSAPRevenue(textBox1.Text, textBox2.Text);

            System.Data.DataTable dtCost = MakeTableCombine();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string ���;
            string ��إN��;

            Int32 �渹;


            Int32 ��¦�渹;
            Int32 ��¦�C;

            //20080904
            //�ŧi DuplicateKey ���ˬd
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                ��� = Convert.ToString(dt.Rows[i]["��O"]);
                �渹 = Convert.ToInt32(dt.Rows[i]["DocNum"]);

                ��إN�� = Convert.ToString(dt.Rows[i]["��إN��"]);



                dtDoc = GetSAPDoc(���, �渹, ��إN��, textBox1.Text, textBox2.Text);


                //if (�渹 == 23116)
                //{
                //    MessageBox.Show("");
                //}

                //20080904 �W�C�קK�P�f�������л{�C
                //�@�i�榳�h�ؾP�f���J,�P�f�����u���@��
                //�@�k:

                ��¦�渹 = -1;
                ��¦�C = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();

                    //  dr["��إN��"] = ��إN��;
                    dr["���J���"] = ���;
                    dr["���J�渹"] = �渹;


                    dr["�Ȥ�s��"] = "'"+Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["�Ȥ�W��"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["���~�s��"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["���~�W��"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);
                    dr["�ƶq"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["���"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]);
                    dr["�渹�`����"] = 0;
                    dr["���ئ���"] = 0;


                    //20081008
                    //�~�ȭ�
                    dr["�~�ȭ��s��"] = Convert.ToString(dt.Rows[i]["�~�ȭ��s��"]);
                    dr["�m�W"] = Convert.ToString(dt.Rows[i]["�m�W"]);
                    dr["�Ȥ�s��"] = Convert.ToString(dt.Rows[i]["�Ȥ�s��"]);

                    if (��� == "AR" || ��� == "�U��" || ��� == "AR�w")
                    {



                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ��¦�渹 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["��¦�渹"] = ��¦�渹;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            ��¦�C = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["��¦�C"] = ��¦�C;
                        }

                    }

                    //�`���J�g�b�̫�@��
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (��� == "AR" || ��� == "AR-�A��" || ��� == "AR�w")
                        {
                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]);
                        }
                        else
                        {

                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]) * (-1);
                        }
                    }

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  �ƶq�令 �t��
                        dr["�ƶq"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }


                    //�̾�  ��¦�渹 & ��¦�C ���o����
                    //�p�G��ڥ����S����¦�渹 & ��¦�C�N�{�C����

                    //20080916 AR ������ �y�� ������|
                    if (��� == "AR" || ��� == "AR�w")
                    {

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (��¦�渹.ToString() == "3169" && �渹.ToString() == "3429")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (��¦�渹.ToString() == "3167" && �渹.ToString() == "3404")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }

                            dtDocLine = GetSAPDocByLine("��f", ��¦�渹, ��¦�C);

                            dr["�������"] = "��f";
                            dr["�����渹"] = ��¦�渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                               * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["�渹�`����"] = 0;
                                    if (�渹 != DuplicateKey)
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }


                                //20091204�@��h
                                if (��¦�渹.ToString() == "5394" && �渹.ToString() == "5673")
                                {

                                    dr["�渹�`����"] = 2111964;
                                    dr["�渹�`���J"] = 0;

                                }
                                //2010331�h��@
                                if (�渹.ToString() == "6975")
                                {
                                    dr["�渹�`����"] = 0;

                                }

                                //2010409�q����AR
                                if (�渹.ToString() == "7022")
                                {
                                    dr["�渹�`����"] = "5476";
                                    dr["���ئ���"] = "5476";
                                }

                                //20150506 AR��AR�w�@�s
                                if (��¦�渹.ToString() == "26223" && ��¦�C.ToString() == "0")
                                {
                                    dr["�渹�`����"] = "1005608";


                                }

                                //20111102 �Ӷ����f�h��1
                                System.Data.DataTable GT = TF(��¦�渹.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["�Ǹ�"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (�渹.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["�渹�`����"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                            //���������Ӧۦܩ����

                        }
                        //�S����¦�渹
                        else
                        {
                            //������Ƭ��ۤw
                            dr["�������"] = ���;
                            dr["�����渹"] = �渹;
                            dtDocLine = GetSAPDocByLine(���, �渹);
                            if (dtDocLine.Rows.Count == 1)
                            {
                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //�Ϧ^�h��P�f����
                                        System.Data.DataTable dtSalesCost = GetSalesCost(�渹.ToString());
                                        try
                                        {
                                            dr["�渹�`����"] = Convert.ToInt32(dtSalesCost.Rows[0]["�`����"]);
                                        }
                                        catch
                                        {
                                            dr["�渹�`����"] = 0;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                        }

                    }

                    // 3 ��רҨS���ӷ��渹

                    //20081007 �W�[�P�h..�������t

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////�n�P�_�ӷ������
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));

                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }


                        }
                        else
                        {
                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                        // dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }
                        }
                    }

                    dtCost.Rows.Add(dr);

                }
            }


            //bindingSource1.DataSource = dtCost;
            //dataGridView8.DataSource = bindingSource1.DataSource;

            if (checkBox1.Checked)
            {
                ACME.Form1Item1 frm = new ACME.Form1Item1();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            if (checkBox2.Checked)
            {
                ACME.Form1Itemc frm = new ACME.Form1Itemc();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            else
            {
                ACME.Form1Item frm = new ACME.Form1Item();
                frm.dt = dtCost;
                frm.ShowDialog();
            }
            //�~�ȧO



        }





        private void ViewBatchPayment()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("         SELECT T.CARDCODE �Ȥ�s��,T.CARDNAME �Ȥ�W��,SUM(�P�����B) �P�����B,SUM(�P�����B)+SUM(�h�^���B)+SUM(�������B) �`��P���B,SUM(�h�^���B) �h�^���B,SUM(�������B) �������B FROM ");
            sb.Append(" (            SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              Convert(int,sum(LINETOTAL)) �P�����B,0 �h�^���B,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4110%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              Convert(int,sum(LINETOTAL)) * (-1)  �P�����B ,0 �h�^���B,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4110%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName]");
            sb.Append(" union all");
            sb.Append("             SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              0 �P�����B, Convert(int,sum(LINETOTAL)) �h�^���B,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4170%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4170%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              0  �P�����B ,convert(int,sum(LINETOTAL)) * (-1) �h�^���B,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4170%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4170%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName]");
            sb.Append(" union all");
            sb.Append("             SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              0 �P�����B, 0 �h�^���B,Convert(int,sum(LINETOTAL)) �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4190%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4190%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              0  �P�����B ,0 �h�^���B,Convert(int,sum(LINETOTAL))*(-1) �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4190%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4190%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName] ) T");
            sb.Append("            GROUP BY T.CARDCODE,T.CARDNAME");
            sb.Append("              ORDER BY �`��P���B DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //��J���F�W��
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " OINV");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView8.DataSource = bindingSource1;


            ACME.Form1Rpt4C frm4 = new ACME.Form1Rpt4C();
            frm4.dt = ds.Tables[0];
            frm4.s = textBox1.Text;
            frm4.q = textBox2.Text;
            frm4.ShowDialog();

        }

        private void ViewBatchPayment3()
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT T.CARDCODE �Ȥ�s��,T.CARDNAME �Ȥ�W��,SUM(�P�����B) �P�����B,SUM(cast(�P��ƶq as int)) �P���ƶq,SUM(�P�����B)+SUM(�h�^���B)+SUM(�������B) �`��P���B,SUM(cast(�P��ƶq as int))+SUM(cast(�h�^�ƶq as int)) �`��P�ƶq,SUM(�h�^���B) �h�^���B,SUM(cast(�h�^�ƶq as int)) �h�^�ƶq,SUM(�������B) �������B FROM ");
            sb.Append(" (            SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              Convert(int,sum(LINETOTAL)) �P�����B,SUM([quantity]) �P��ƶq,0 �h�^���B,0 �h�^�ƶq,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4110%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              Convert(int,sum(LINETOTAL)) * (-1)  �P�����B ,SUM([quantity])*(-1)  �P��ƶq,0 �h�^���B,0 �h�^�ƶq,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4110%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4110%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName]");
            sb.Append(" union all");
            sb.Append("             SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              0 �P�����B,0 �P��ƶq, Convert(int,sum(LINETOTAL)) �h�^���B,SUM([quantity]) �h�^�ƶq,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4170%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4170%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              0  �P�����B ,0 �P��ƶq,convert(int,sum(LINETOTAL)) * (-1) �h�^���B,SUM([quantity])* (-1) �h�^�ƶq,0 �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4170%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4170%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName]");
            sb.Append(" union all");
            sb.Append("             SELECT  T2.[CardCode], T2.[CardName],");
            sb.Append("              0 �P�����B,0 �P��ƶq, 0 �h�^���B,0 �h�^�ƶq,Convert(int,sum(LINETOTAL)) �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN OINV T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN INV1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4190%') ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4190%' ) ");
            sb.Append("              GROUP BY     T2.[CardCode], T2.[CardName]");
            sb.Append("              union all");
            sb.Append("              SELECT   T2.[CardCode], T2.[CardName],");
            sb.Append("              0  �P�����B,0 �P��ƶq,0 �h�^���B,0 �h�^�ƶq,Convert(int,sum(LINETOTAL))*(-1) �������B");
            sb.Append("              FROM OJDT T0  INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ");
            sb.Append("              INNER JOIN ORIN T2 ON T0.TransId = T2.TransId ");
            sb.Append("              INNER JOIN RIN1 T5 ON T2.DocEntry = T5.DocEntry ");
            sb.Append("              WHERE T2.[DocType] ='I' and (T1.[Account] like '4190%' ) ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and Convert(varchar(10),(T0.RefDate),112) between  @DocDate1 and @DocDate2 ");
            }
            sb.Append("              and  (T5.[AcctCode] like '4190%' )");
            sb.Append("              GROUP BY  T2.[CardCode], T2.[CardName] ) T");
            sb.Append("            GROUP BY T.CARDCODE,T.CARDNAME");
            sb.Append("              ORDER BY �`��P���B DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //��J���F�W��
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " OINV");
            }
            finally
            {
                MyConnection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView8.DataSource = bindingSource1;


            ACME.Form1Rpt4D frm4 = new ACME.Form1Rpt4D();
            frm4.dt = ds.Tables[0];
            frm4.s = textBox1.Text;
            frm4.q = textBox2.Text;
            frm4.ShowDialog();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            YEAROITM = "2021";
            if (globals.GroupID.ToString().Trim() != "EEP")
            {
                button29.Visible = false;
                button7.Visible = false;
     
            }
  if (globals.DBNAME == "�F�ͥ�")
            {
                COM = "�F�ͥͬ�޵o�i�]�`�`�^�������q";
            }
            else
            {
                COM = "�i���͹�~�ѥ��������q";
            }

            if (globals.GroupID.ToString().Trim() == "ACCS" )
            {
                Close();
            }

            System.Data.DataTable TACO2 = TACO("EUNICE");
            textBox5.Text = TACO2.Rows[0]["PARAM_NO"].ToString();

            textBox6.Text = DateTime.Now.ToString("yyyyMM");
            comboBox6.Text = "��";
            if (globals.GroupID.ToString().Trim() == "SHI")
            {
                checkBox1.Visible = false;
                button4.Visible = false;
                button6.Visible = false;
                button5.Visible = false;
                radioButton1.Visible = false;
                radioButton2.Visible = false;
                radioButton3.Visible = false;

             
            }

            comboBox1.Text = "�Ȥ�O";
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            textBox7.Text = GetMenu.DFirst();
            textBox8.Text = GetMenu.DLast();
            textBox3.Text = DateTime.Now.ToString("yyyyMMdd");

            textBox10.Text = GetMenu.DFirst();
            textBox11.Text = GetMenu.DLast();

            textBox12.Text = GetMenu.DFirst();
            textBox13.Text = GetMenu.DLast();

            textBox16.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox14.Text = GetMenu.DFirst();
            textBox15.Text = GetMenu.DLast();

            textBox19.Text = GetMenu.DFirst();
            textBox20.Text = GetMenu.DLast();

            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Year(), "DataValue", "DataValue");

            UtilSimple.SetLookupBinding(comboBox3, GetMenu.Year(), "DataValue", "DataValue");

            UtilSimple.SetLookupBinding(comboBox4, GetMenu.Year(), "DataValue", "DataValue");

            UtilSimple.SetLookupBinding(comboBox7, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox5, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox8, GetMenu.Year2017(), "DataValue", "DataValue");

            UtilSimple.SetLookupBinding(comboBox9, GetMenu.Month2(), "DataValue", "DataValue");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {
                    FileName = openFileDialog1.FileName;

                    if (comboBox1.Text == "�Ȥ�O")
                    {
                        GetExcelProduct(FileName, 2, 1, 22, 3);
                    }
                    else if (comboBox1.Text == "�~�ȧO")
                    {
                        GetExcelProduct(FileName, 2, 1, 21, 3);
                    }
                    else if (comboBox1.Text == "���~�O")
                    {
                        GetExcelProduct(FileName, 2, 1, 20, 3);
                    }

                    else if (comboBox1.Text == "�Ȥ����Ʀ�")
                    {
                        GetExcelProduct(FileName, 1, 1, 24, 6);
                    }
                    MessageBox.Show("�����ɮ�->" + NewFileName);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }
        private void GetExcelProduct(string ExcelFile, int a, int b, int c, int d)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //���o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);


            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;


                for (int iRecord = iRowCnt; iRecord >= d; iRecord--)
                {


                    //���X���� - ���
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, a]);
                    range.Select();
                    sTemp = (string)range.Text;




                    if (sTemp == "")
                    {
                        for (int i = b; i <= c; i++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, i]);
                            range.Select();
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        }
                    }



                }

            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
               "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";


                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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


            }

        }
        private void Category(int year, string ff, string TABLE)
        {
            if (ff != "W")
            {
                AddAUOGD1(TABLE);
            }
            System.Data.DataTable dt = null;
            if (year == 2||year == 6)
            {
               
                dt = GetSAPRevenueear(ff);

         
            }
         
            else if (year == 3)
            {
                dt = GetSAPRevenueTemp3();

            }
            else if (year == 4)
            {
                string q = util.quarter(textBox6.Text);
                string year2 = textBox6.Text.Substring(0, 4);
                dt = GetSAPRevenueTemp3q(q, year2);

            }
            else if (year == 5)
            {
                dt = GetSAPRevenueTemp3y();

            }
            else if (year == 7)
            {
  
                dt = GetSAPRevenueTempLED();

            }
            else if (year == 8)
            {

                dt = GetMenu.GetSAPRevenue(A1, A2);

            }
            System.Data.DataTable dtCost = MakeTableCombine2();

            DataRow dr = null;

            System.Data.DataTable dtDoc = null;


            System.Data.DataTable dtDocLine = null;

            string ���;
            string ��إN��;

            Int32 �渹;
            DateTime ���;

            Int32 ��¦�渹;
            Int32 ��¦�C;

            //20080904
            //�ŧi DuplicateKey ���ˬd
            Int32 DuplicateKey = 0;


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {

                ��� = Convert.ToString(dt.Rows[i]["��O"]);
                �渹 = Convert.ToInt32(dt.Rows[i]["DocNum"]);
                ��� = Convert.ToDateTime(dt.Rows[i]["���"]);
                ��إN�� = Convert.ToString(dt.Rows[i]["��إN��"]);

                //if (�渹.ToString().Trim() == "47188")
                //{
                //    MessageBox.Show("A");
                //}
                dtDoc = GetSAPDoc(���, �渹, ��إN��, A1, A2);


                ��¦�渹 = -1;
                ��¦�C = -1;


                for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();
                   
                    dr["���J���"] = ���;
                    dr["���J�渹"] = �渹;


                    dr["���"] = ���;
                    dr["��إN��"] = ��إN��;
                    dr["�Ȥ�s��"] = "'"+Convert.ToString(dtDoc.Rows[j]["CardCode"]);
                    dr["�Ȥ�W��"] = Convert.ToString(dtDoc.Rows[j]["CardName"]);
                    dr["���~�s��"] = Convert.ToString(dtDoc.Rows[j]["ItemCode"]);
                    dr["���~�W��"] = Convert.ToString(dtDoc.Rows[j]["Dscription"]);

       

                    if (year == 7)
                    { 
                        dr["�Ȥ�s��"] = Convert.ToString(dt.Rows[i]["����"]);
                    }
                    else
                    {
                        dr["�Ȥ�s��"] = Convert.ToString(dtDoc.Rows[j]["GROUPCODE"]);
                    }
                    string D = dtDoc.Rows[j]["LineTotal"].ToString();
                    dr["�ƶq"] = Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]);
                    dr["���"] = Convert.ToDecimal(dtDoc.Rows[j]["Price"]);
                    dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]);
                    dr["�渹�`����"] = 0;
                    dr["���ئ���"] = 0;


                    //�~�ȭ�
                    dr["�~�ȭ��s��"] = Convert.ToString(dt.Rows[i]["�~�ȭ��s��"]);
                    dr["�m�W"] = Convert.ToString(dt.Rows[i]["�m�W"]);



                    if (��� == "AR" || ��� == "�U��" || ��� == "AR�w")
                    {



                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ��¦�渹 = Convert.ToInt32(dtDoc.Rows[j]["BaseEntry"]);
                            dr["��¦�渹"] = ��¦�渹;
                        }

                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseLine"]))
                        {
                            ��¦�C = Convert.ToInt32(dtDoc.Rows[j]["BaseLine"]);
                            dr["��¦�C"] = ��¦�C;
                        }

                    }

       
                    //�`���J�g�b�̫�@��
                    if (j == dtDoc.Rows.Count - 1)
                    {
                        if (��� == "AR" || ��� == "AR-�A��" || ��� == "AR�w")
                        {
                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]);

                        }
                    
                       
                        else
                        {

                            dr["�渹�`���J"] = Convert.ToInt32(dt.Rows[i]["�`����"]) * (-1);
                        }
                    }

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h" || ��� == "JE")
                    {
                        dr["���B"] = Convert.ToInt32(dtDoc.Rows[j]["LineTotal"]) * (-1);
                        //20081007  �ƶq�令 �t��
                        dr["�ƶq"] = Convert.ToInt32(dtDoc.Rows[j]["Quantity"]) * (-1);
                    }

                    if (��� == "AR" || ��� == "AR�w")
                    {
                        //0303
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            if (��¦�渹.ToString() == "3169" && �渹.ToString() == "3429")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }
                            if (��¦�渹.ToString() == "3167" && �渹.ToString() == "3404")
                            {
                                dr["���ئ���"] = 0;
                                dr["�渹�`����"] = 0;
                                dtCost.Rows.Add(dr);
                                continue;
                            }

                            dtDocLine = GetSAPDocByLine("��f", ��¦�渹, ��¦�C);

                            dr["�������"] = "��f";
                            dr["�����渹"] = ��¦�渹;

                            if (dtDocLine.Rows.Count == 1)
                            {
                                string AA1 = dtDocLine.Rows[0]["StockPrice"].ToString();
                                string AA2 = dtDocLine.Rows[0]["Quantity"].ToString();

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDocLine.Rows[0]["StockPrice"])
                                               * Convert.ToDecimal(dtDocLine.Rows[0]["Quantity"]));
          
                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    //20080904
                                    dr["�渹�`����"] = 0;
                                    if (�渹 != DuplicateKey)
                                    {
                                        string AA3 = dtDocLine.Rows[0]["�`����"].ToString();
                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }
                                //20091204�@��h
                                if (��¦�渹.ToString() == "5394" && �渹.ToString() == "5673")
                                {

                                    dr["�渹�`����"] = 2111964;
                                    dr["�渹�`���J"] = 0;

                                }
                                //2010331�h��@
                                if (�渹.ToString() == "6975")
                                {
                                    dr["�渹�`����"] = 0;

                                }
                                //2010409�q����AR
                                if (�渹.ToString() == "7022")
                                {
                                    dr["�渹�`����"] = "5476";
                                    dr["���ئ���"] = "5476";

                                }
                                
                                //20150506 AR��AR�w�@�s
                                if (��¦�渹.ToString() == "26223" && ��¦�C.ToString() == "0")
                                {
                                    dr["�渹�`����"] = "1005608";
                 

                                }

                                if (��¦�渹.ToString() == "26441" && Convert.ToInt16(dtDoc.Rows[j]["Quantity"]) == 4)
                                {
                                    dr["�渹�`����"] = "261458";


                                }

                                if (�渹.ToString() == "47188" && Convert.ToInt16(dtDoc.Rows[j]["Quantity"]) == 55)
                                {
                                    dr["�渹�`����"] = "720368";


                                }
                                System.Data.DataTable GT = TF(��¦�渹.ToString());

                                if (GT.Rows.Count > 0)
                                {
                                    for (int n = 0; n <= GT.Rows.Count - 1; n++)
                                    {
                                        string g2 = GT.Rows[n]["�Ǹ�"].ToString();
                                        string g3 = GT.Rows[n]["AR"].ToString();
                                        if (�渹.ToString() == g3)
                                        {
                                            if (g2 != "1")
                                            {
                                                dr["�渹�`����"] = "0";
                                            }

                                        }
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;

                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                            //���������Ӧۦܩ����

                        }
                        //�S����¦�渹
                        else
                        {
                            //������Ƭ��ۤw
                            dr["�������"] = ���;
                            dr["�����渹"] = �渹;


                            dtDocLine = GetSAPDocByLine(���, �渹);

                            if (dtDocLine != null)
                            {

                                if (dtDocLine.Rows.Count == 1)
                                {
                                    dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                                   * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"]));

                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                        {
                                            dr["�渹�`����"] = 0;
                                        }
                                        else
                                        {
                                            //�Ϧ^�h��P�f����
                                            System.Data.DataTable dtSalesCost = GetSalesCost(�渹.ToString());
                                            try
                                            {
                                                dr["�渹�`����"] = Convert.ToInt32(dtSalesCost.Rows[0]["�`����"]);
                                            }
                                            catch
                                            {
                                                dr["�渹�`����"] = 0;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    //Rows.Count =0 �������s
                                    dr["���ئ���"] = 0;
                                    if (j == dtDoc.Rows.Count - 1)
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }

                        }

                    }

                    // 3 ��רҨS���ӷ��渹

                    //20081007 �W�[�P�h..�������t

                    if (��� == "�U��" || ��� == "�U��-�A��" || ��� == "�P�h")
                    {
                        if (!Convert.IsDBNull(dtDoc.Rows[j]["BaseEntry"]))
                        {
                            ////�n�P�_�ӷ������
                            //MessageBox.Show(Convert.ToString(dtDoc.Rows[j]["BaseEntry"]));
                         
                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;

                                    //20081231
                                    if (�渹 != DuplicateKey)
                                    {

                                        dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }
                                    DuplicateKey = �渹;
                                }
                            }


                        }
                        else
                        {

                          
                            dtDocLine = GetSAPDocByLine(���, �渹);

                            //������Ƭ��ۤw

                            dr["�������"] = ���;

                            dr["�����渹"] = �渹;

                            if (dtDocLine.Rows.Count == 1)
                            {

                                dr["���ئ���"] = Convert.ToInt32(Convert.ToDecimal(dtDoc.Rows[j]["StockPrice"])
                                               * Convert.ToDecimal(dtDoc.Rows[j]["Quantity"])) * (-1);

                                if (j == dtDoc.Rows.Count - 1)
                                {

                                    if (Convert.IsDBNull(dtDocLine.Rows[0]["�`����"]))
                                    {
                                        dr["�渹�`����"] = 0;
                                    }
                                    else
                                    {
                                        //20081231
                                        if (�渹 != DuplicateKey)
                                        {

                                            dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                        }
                                        DuplicateKey = �渹;

                                        // dr["�渹�`����"] = Convert.ToInt32(dtDocLine.Rows[0]["�`����"]);
                                    }

                                }
                            }
                            else
                            {
                                //Rows.Count =0 �������s
                                dr["���ئ���"] = 0;
                                if (j == dtDoc.Rows.Count - 1)
                                {
                                    dr["�渹�`����"] = 0;
                                }
                            }
                        }
                    }
                 
                    dtCost.Rows.Add(dr);


                }

      
            }

            System.Data.DataTable ESCOT1 = GetESCO(textBox1.Text, textBox2.Text, year);
            if (ESCOT1.Rows.Count > 0)
            {

                for (int j = 0; j <= ESCOT1.Rows.Count - 1; j++)
                {

                    dr = dtCost.NewRow();

                    string BU = Convert.ToString(ESCOT1.Rows[j]["�Ȥ�s��"]);
                    string CARDCODE = Convert.ToString(ESCOT1.Rows[j]["�Ȥ�s��"]);
                    dr["�Ȥ�s��"] = "'" + Convert.ToString(ESCOT1.Rows[j]["�Ȥ�s��"]);
                    dr["�Ȥ�W��"] = Convert.ToString(ESCOT1.Rows[j]["�Ȥ�W��"]); ;
                    dr["�m�W"] = "'" + Convert.ToString(ESCOT1.Rows[j]["�m�W"]);
                    dr["�ƶq"] = Convert.ToString(ESCOT1.Rows[j]["�ƶq"]);
                    dr["���~�s��"] = "";
                    dr["�渹�`���J"] = Convert.ToInt32(ESCOT1.Rows[j]["�渹�`���J"]);
                    dr["�渹�`����"] = Convert.ToInt32(ESCOT1.Rows[j]["�渹�`����"]);
                    dr["���J�渹"] = Convert.ToInt32(ESCOT1.Rows[j]["���J�渹"]);
                    dr["��إN��"] = Convert.ToString(ESCOT1.Rows[j]["��إN��"]);
                    dr["�Ȥ�s��"] = BU;

                    if (year == 7)
                    {
                        dr["�Ȥ�s��"] = "�`��A�ȨƷ~��";
                        if (BU == "103" || CARDCODE=="U0516")
                        {
                            dr["�Ȥ�s��"] = "TFT�Ʒ~��";
                        }
                    }
                    dr["���ئ���"] = Convert.ToInt32(ESCOT1.Rows[j]["���ئ���"]);             
                    dr["���B"] = Convert.ToInt32(ESCOT1.Rows[j]["���B"]);
                    dr["���"] = Convert.ToDateTime(Convert.ToString(ESCOT1.Rows[j]["���"]));
                    dtCost.Rows.Add(dr);
            
                }
            }

            dataGridView8.DataSource = dtCost;
            for (int i = 0; i <= dataGridView8.Rows.Count - 1; i++)
            {

                DOCTYPE = dataGridView8.Rows[i].Cells["���J���"].Value.ToString().Trim();
                ss1 = dataGridView8.Rows[i].Cells["�Ȥ�s��"].Value.ToString();
                ss2 = dataGridView8.Rows[i].Cells["�Ȥ�W��"].Value.ToString();
                ss3 = dataGridView8.Rows[i].Cells["�m�W"].Value.ToString();
                ss4 = dataGridView8.Rows[i].Cells["�ƶq"].Value.ToString();
                ss5 = dataGridView8.Rows[i].Cells["�渹�`���J"].Value.ToString();
                ss6 = dataGridView8.Rows[i].Cells["�渹�`����"].Value.ToString();
                ss7 = dataGridView8.Rows[i].Cells["���J�渹"].Value.ToString();
                ss8 = dataGridView8.Rows[i].Cells["��إN��"].Value.ToString();
                ss9 = dataGridView8.Rows[i].Cells["�Ȥ�s��"].Value.ToString();
                ss10 = dataGridView8.Rows[i].Cells["���ئ���"].Value.ToString();
                ss101 = dataGridView8.Rows[i].Cells["���B"].Value.ToString();
                ss11 = dataGridView8.Rows[i].Cells["���~�s��"].Value.ToString();
                ss12 = dataGridView8.Rows[i].Cells["���~�W��"].Value.ToString();
                BASEDOC = dataGridView8.Rows[i].Cells["��¦�渹"].Value.ToString();
                BASELINE = dataGridView8.Rows[i].Cells["��¦�C"].Value.ToString();
         
                if (ss7 == "56452"||ss7 == "56453"|| ss7 == "58960")
                {
                    ss9 = "103";
                }
                //20150324 �Ӷ���
                if (ss4 == "1")
                {
                    if (ss7 == "24018")
                    {
                        ss5 = "31600";
                    }
                    if (ss7 == "25636")
                    {
                        ss5 = "30000";
                    }
                }
                 
            //౽n
                    if (ss7 == "26976")
                    {
                        if (ss4 == "206")
                        {
                            ss6 = "575134";
                        }

                        if (ss4 == "66")
                        {
                            ss6 = "143783";
                        }

                        if (ss4 == "110")
                        {
                            ss6 = "239639";
                        }
                    }
                    if (ss7 == "28510")
                    {
                        if (ss4 == "108")
                        {
                            ss6 = "221455";
                        }
                    }
           
                //20151207
                    if (ss7 == "30027")
                    {
                        if (ss4 == "1")
                        {
                            ss101 = "5790";
                        }
                    }
                //2050901
                    if (ss7 == "27312")
                    {
                        if (ss4 == "264")
                        {
                            ss6 = "571912";
                            ss101 = "608889";
                            
                        }
                    }
                if (ss7 == "27487")
                {
                    if (ss4 == "100")
                    {
                        ss101 = "265338";
                    }
                    if (ss4 == "900")
                    {
                        ss101 = "2388046";
                        ss6 = "2450000";
                    }
                    if (ss4 == "200")
                    {
                        ss101 = "530677";
                    }
                }
                if (ss7 == "30013")
                {
                    if (ss4 == "160")
                    {
                        ss6 = "436545";
                    }

                    if (ss4 == "16")
                    {
                        ss101 = "38166";
                    }
                }


                //20160513

                if (ss7 == "29133")
                {
                    if (ss4 == "384")
                    {
                        ss6 = "787394";
                    }
                }

                if (ss7 == "30993")
                {
                    if (ss4 == "325")
                    {
                        ss6 = "1220106";
                    }
                }
                if (ss7 == "31526")
                {
                    if (ss4 == "240")
                    {
                        ss6 = "654818";
                    }
                }
                if (ss7 == "31702")
                {
                    if (ss4 == "3212")
                    {
                        ss6 = "9366132";

                        ss101 = "9484876";
                    }
                }
                if (ss7 == "29133")
                {
                    if (ss4 == "960")
                    {
                        ss6 = "2070704";
                    }
                }
                //�{�ɤ�
                if (ss7 == "29010")
                {
                    if (ss101 == "79625")
                    {
                        ss6 = "69647";
                    }
                }


                //����
                if (ss7 == "32778")
                {

                    ss6 = "0";
      
                }

                //20150625 �Ӷ���t1��
                if (ss7 == "22848")
                {
                    ss101 = "3648959";
                }
                if (ss7 == "27445")
                {
                    ss101 = "1107541";
                }


                if (ss7 == "31193")
                {
                    if (ss101 == "13272")
                    {
                        ss6 = "7256";
                    }
                }
                if (ss7 == "31194")
                {
                    if (ss101 == "26543")
                    {
                        ss6 = "14513";
                    }
                    if (ss101 == "4976")
                    {
                        ss101 = "4977";
                    }
                }
                if (ss7 == "31195")
                {
                    if (ss101 == "39815")
                    {
                        ss6 = "21769";
                    }
                }

                //33318
                if (ss7 == "33318")
                {
                    ss101 = "19489";
                    ss5 = "19489";
                }
                if (ss7 == "1532"|| ss7 == "1533")
                {
                    if (ss2 == "YSHENG  HOLDINGS  LIMITED")
                    {
                        ss101 = "0";
                        ss5 = "0";
                    }
                }

                //20180104
                if (ss7 == "40574")
                {
                    if (ss4 == "12")
                    {
                        ss6 = "41747";
                        ss10 = "41747";
                    }
                }

                //20181218
                if (ss7 == "45051")
                {
                    if (ss4 == "1130")
                    {
                        ss6 = "3121647";
                        ss10 = "3121647";
                    }
                }

                //20190827
                if (ss7 == "411579")
                {
                    ss101 = "388000";
                    ss5 = "388000";
                }
                if (ss7 == "411580")
                {
                    ss6 = "157563";
                    ss10 = "157563";
                }
                if (ss7 == "411581")
                {
                    ss6 = "230437";
                    ss10 = "230437";
                }
                if (ss7 == "464404")
                {
                    ss6 = "14198843";
                    ss10 = "14198843";
                    ss9 = "116";
                }
                if (ss7 == "464414")
                {
                    ss6 = "1787684";
                    ss10 = "1787684";
                    ss9 = "116";
                }

                if (ss7 == "464446")
                {
                    ss101 = "17809524";
                    ss5 = "17809524";
                }
                //27445     
                DateTime dd = Convert.ToDateTime(dataGridView8.Rows[i].Cells["���"].Value);

                if (String.IsNullOrEmpty(ss6))
                {
                    ss6 = "0";
                }
                if (String.IsNullOrEmpty(ss5))
                {
                    ss5 = "0";
                }

                int GF = Convert.ToInt32(ss4);

                string DOCENTRY = ss7;
                string COUNTRY = "";
                if (DOCTYPE=="AR"||DOCTYPE =="AR-�A��")
                {
                    System.Data.DataTable G1 = GetOCRD(ss7);
                    if (G1.Rows.Count > 0)
                    {
                        string SHIPCODE = G1.Rows[0][0].ToString();
                        string CARDCODE = G1.Rows[0][1].ToString();

                        System.Data.DataTable G2 = GetOCRD2(CARDCODE, SHIPCODE);

                        if (G2.Rows.Count > 0)
                        {

                            COUNTRY = G2.Rows[0][0].ToString();
                        }
                    }

                }
                else
                {
                    System.Data.DataTable G1 = GetOCRDORIN(ss7);
                    if (G1.Rows.Count > 0)
                    {
                        string SHIPCODE = G1.Rows[0][0].ToString();
                        string CARDCODE = G1.Rows[0][1].ToString();

                        System.Data.DataTable G2 = GetOCRD2(CARDCODE, SHIPCODE);

                        if (G2.Rows.Count > 0)
                        {

                            COUNTRY = G2.Rows[0][0].ToString();
                        }
                    }
                }

                if (TABLE == "Account_Temp61" + YEAROITM || TABLE == "Account_Temp61")
                {
                    AddAUOGD61(TABLE, ss11, ss12, ss1, ss2, ss3, ss4, ss101, ss10, dd, ss7, ss8, ss9, ss5, ss6, BASEDOC, BASELINE, COUNTRY);
                }
                else
                {
                    AddAUOGD(TABLE, ss1, ss2, ss3, ss4, ss101, ss6, dd, ss7, ss8, ss9, COUNTRY);
                }

            }
         
        }



        private void EunLED()
        {
            Category(7, "", "Account_Temp6");


            dtCostDD3 = MakeTableFcstWeek();
            Eun2("2", "TFT�Ʒ~��");
            DJ = 0;
            DJ2 = 0;
            DJ3 = 0;
            DJ33 = 0;
            HJ = 0;
            SALES22 = 0;
            Eun2("2", "LED-Lighting �Ʒ~��");
            DJ = 0;
            DJ2 = 0;
            DJ3 = 0;
            DJ33 = 0;
            HJ = 0;
            SALES22 = 0;
            Eun2("2", "LED-Chip Package �Ʒ~��");
            DJ = 0;
            DJ2 = 0;
            DJ3 = 0;
            DJ33 = 0;
            HJ = 0;
            SALES22 = 0;
            Eun2("2", "�`��A�ȨƷ~��");
            DJ = 0;
            DJ2 = 0;
            DJ3 = 0;
            DJ33 = 0;
            HJ = 0;
            SALES22 = 0;
            Eun2("2", "���Ʒ~��");

        }
        private void Eun2(string TYPE,string GROUP)
        {
            string �Ȥ�s��t;
            string �Ȥ�W��t;
            string �~��t;
            string BUt;
            System.Data.DataTable dtemp5 = GetTemp5("N", GROUP);
            if (dtemp5.Rows.Count < 1)
            {
                return;
            }

            DataRow drtemp5 = null;
            System.Data.DataTable dtWeek = MakeTableWeek();
            GetMonthWeekStartDate(textBox3.Text, dtWeek);
            string ff = "";
            if (TYPE == "1")
            {
                if (GROUP == "103")
                {
                    ff = "TFT";
                }
                else if (GROUP == "104")
                {
                    ff = "LED";
                }
                else if (GROUP == "105")
                {
                    ff = "SOLAR";
                }
            }
            if (TYPE == "2")
            {
                ff = GROUP;
            }
            drtemp5 = dtCostDD3.NewRow();
            drtemp5["�~��"] = ff;
            dtCostDD3.Rows.Add(drtemp5);
            for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
            {
                j = i + 1;
                drtemp5 = dtCostDD3.NewRow();
                �Ȥ�s��t = dtemp5.Rows[i]["�Ȥ�s��"].ToString();
                �Ȥ�W��t = dtemp5.Rows[i]["�Ȥ�W��"].ToString();
      
                �~��t = dtemp5.Rows[i]["�~��"].ToString();
                BUt = dtemp5.Rows[i]["BU"].ToString();
                drtemp5["row"] = j.ToString();
                drtemp5["�~��"] = �~��t;
                drtemp5["BU"] = BUt;
                drtemp5["�Ȥ�s��"] = "'"+�Ȥ�s��t;
                drtemp5["�Ȥ�W��"] = �Ȥ�W��t;
                
               
                for (int w = 0; w <= p; w++)
                {
                    string sg = "";
                    string sg2 = "";
                    string sg3 = "";

                    int h = 0;
                    int hR = 0;
                    int hC = 0;
                    int hG = 0;
                    string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                    string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                    string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                    System.Data.DataTable dh = null;
                    dh = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "Q");
                    sg = dh.Rows[0]["�ƶq"].ToString();
                    drtemp5[WeekName + "_Q"] = sg;

                    System.Data.DataTable dh2 = null;
                    dh2 = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "R");
                    sg2 = dh2.Rows[0]["���B"].ToString();
                    drtemp5[WeekName + "_R"] = sg2;

                    System.Data.DataTable dh3 = null;
                    dh3 = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "C");
              
                    sg3 = dh3.Rows[0]["����"].ToString();
                    drtemp5[WeekName + "_C"] = sg3;

                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    int �P�����B = Convert.ToInt32(sg2);
                    double �Q��� = (�Q�� / (�P�����B)) * 100;


                    drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                    h = Convert.ToInt32(sg);
                    hR = Convert.ToInt32(sg2);
                    hC = Convert.ToInt32(sg3);
                    hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    //��
                    string aa = Convert.ToString(�Q���);
                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                    {
                        drtemp5[WeekName + "_P"] = "0%";
                    }
                    else
                    {
                        drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                    }
                    hh += h;
                    hhR += hR;
                    hhC += hC;
                    hhG += hG;

                }

                drtemp5["WT_Q"] = hh;
                drtemp5["WT_R"] = hhR;
                drtemp5["WT_C"] = hhC;
                drtemp5["WT_G"] = hhG;
                DJ = hhR;
                hh = 0;
                hhR = 0;
                hhC = 0;
                hhG = 0;


                double �Q���1 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
                string aa2 = Convert.ToString(�Q���1);
                if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr" || aa2 == "�D�ƭ�" || aa2 == "��" || aa2 == "-��")
                {
                    drtemp5["WT_P"] = 0;
                }
                else
                {
                    drtemp5["WT_P"] = �Q���1.ToString("#0.00") + "%";
                }


                System.Data.DataTable dh11 = null;
                dh11 = GetTemp5_1TOTAL();
                HJ = Convert.ToDouble(dh11.Rows[0]["���B"].ToString());

                double SALES = (DJ / HJ) * 100;
                drtemp5["SALES2"] = SALES.ToString("#0.00") + "%";
                SALES22 += SALES;
                drtemp5["SALES22"] = SALES22.ToString("#0.00") + "%";
                dtCostDD3.Rows.Add(drtemp5);
            }


            drtemp5 = dtCostDD3.NewRow();
            drtemp5["�~��"] = "";
            drtemp5["�Ȥ�s��"] = "";
            drtemp5["�Ȥ�W��"] = "3rd parties sub total";
            for (int w = 0; w <= p; w++)
            {
                string sg = "";
                string sg2 = "";
                string sg3 = "";

                int h = 0;
                int hR = 0;
                int hC = 0;
                int hG = 0;

                string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                System.Data.DataTable dh = null;
                dh = GetTemp5_1t(StartDate, EndDate, "Q", "N", GROUP);
                sg = dh.Rows[0]["�ƶq"].ToString();
                drtemp5[WeekName + "_Q"] = sg;

                System.Data.DataTable dh2 = null;
                dh2 = GetTemp5_1t(StartDate, EndDate, "R", "N", GROUP);
                sg2 = dh2.Rows[0]["���B"].ToString();
                drtemp5[WeekName + "_R"] = sg2;

                System.Data.DataTable dh3 = null;
                dh3 = GetTemp5_1t(StartDate, EndDate, "C", "N", GROUP);
                sg3 = dh3.Rows[0]["����"].ToString();
                drtemp5[WeekName + "_C"] = sg3;

                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                int �P�����B = Convert.ToInt32(sg2);
                double �Q��� = (�Q�� / (�P�����B)) * 100;


                drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                h = Convert.ToInt32(sg);
                hR = Convert.ToInt32(sg2);
                hC = Convert.ToInt32(sg3);
                hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                string aa = Convert.ToString(�Q���);
                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                {
                    drtemp5[WeekName + "_P"] = 0;
                }
                else
                {
                    drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                }

                hS += h;
                hSR += hR;
                hSC += hC;
                hSG += hG;
                string pp = (w + 1).ToString();
            }
            drtemp5["WT_Q"] = hS;
            drtemp5["WT_R"] = hSR;
            drtemp5["WT_C"] = hSC;
            drtemp5["WT_G"] = hSG;

            hS = 0;
            hSR = 0;
            hSC = 0;
            hSG = 0;



            double �Q���2 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
            string aa3 = Convert.ToString(�Q���2);
            if (aa3 == "���L�a�j" || aa3 == "�t�L�a�j" || aa3 == "���O�@�ӼƦr" || aa3 == "�D�ƭ�")
            {
                drtemp5["WT_P"] = 0;
            }
            else
            {
                drtemp5["WT_P"] = �Q���2.ToString("#0.00") + "%";
            }

            dtCostDD3.Rows.Add(drtemp5);

            if (GROUP == "103")
            {
                //CHOICE
                System.Data.DataTable dtemp6 = GetTemp5("", GROUP);




                for (int i = 0; i <= dtemp6.Rows.Count - 1; i++)
                {
                    drtemp5 = dtCostDD3.NewRow();
                    �Ȥ�s��t = dtemp6.Rows[i]["�Ȥ�s��"].ToString();
                    �Ȥ�W��t = dtemp6.Rows[i]["�Ȥ�W��"].ToString();
                    �~��t = dtemp6.Rows[i]["�~��"].ToString();
                    drtemp5["�~��"] = �~��t;
                    drtemp5["�Ȥ�s��"] = "'"+�Ȥ�s��t;
                    drtemp5["�Ȥ�W��"] = �Ȥ�W��t;



                    for (int w = 0; w <= p; w++)
                    {
                        string sg = "";
                        string sg2 = "";
                        string sg3 = "";

                        int h = 0;
                        int hR = 0;
                        int hC = 0;
                        int hG = 0;

                        string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                        string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                        string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                        System.Data.DataTable dh = null;
                        dh = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "Q");
                        sg = dh.Rows[0]["�ƶq"].ToString();
                        drtemp5[WeekName + "_Q"] = sg;

                        System.Data.DataTable dh2 = null;
                        dh2 = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "R");
                        sg2 = dh2.Rows[0]["���B"].ToString();
                        drtemp5[WeekName + "_R"] = sg2;

                        System.Data.DataTable dh3 = null;
                        dh3 = GetTemp5_1(StartDate, EndDate, �Ȥ�s��t, �~��t, "C");
                        sg3 = dh3.Rows[0]["����"].ToString();
                        drtemp5[WeekName + "_C"] = sg3;

                        double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                        int �P�����B = Convert.ToInt32(sg2);
                        double �Q��� = (�Q�� / (�P�����B)) * 100;


                        drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                        h = Convert.ToInt32(sg);
                        hR = Convert.ToInt32(sg2);
                        hC = Convert.ToInt32(sg3);
                        hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                        string aa = Convert.ToString(�Q���);
                        if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                        {
                            drtemp5[WeekName + "_P"] = 0;
                        }
                        else
                        {
                            drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                        }
                        hCS += h;
                        hCR += hR;
                        hCC += hC;
                        hCG += hG;

                    }
                    drtemp5["WT_Q"] = hCS;
                    drtemp5["WT_R"] = hCR;
                    drtemp5["WT_C"] = hCC;
                    drtemp5["WT_G"] = hCG;

                    DJ3 = hCR;

                    hCS = 0;
                    hCR = 0;
                    hCC = 0;
                    hCG = 0;

                    double �Q���1 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
                    string aa2 = Convert.ToString(�Q���1);
                    if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr" || aa2 == "�D�ƭ�" || aa2 == "��" || aa2 == "-��")
                    {
                        drtemp5["WT_P"] = 0;
                    }
                    else
                    {
                        drtemp5["WT_P"] = �Q���1.ToString("#0.00") + "%";
                    }

                    double SALES = (DJ3 / HJ) * 100;
                    drtemp5["SALES2"] = SALES.ToString("#0.00") + "%";

                    dtCostDD3.Rows.Add(drtemp5);
                }

                //choice�[�`
                drtemp5 = dtCostDD3.NewRow();
                drtemp5["�~��"] = "";
                drtemp5["�Ȥ�s��"] = "";
                drtemp5["�Ȥ�W��"] = "Related Parties sub total";

                for (int w = 0; w <= p; w++)
                {
                    string sg = "";
                    string sg2 = "";
                    string sg3 = "";


                    int h = 0;
                    int hR = 0;
                    int hC = 0;
                    int hG = 0;

                    string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                    string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                    string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                    System.Data.DataTable dh = null;
                    dh = GetTemp5_1t(StartDate, EndDate, "Q", "", "");
                    sg = dh.Rows[0]["�ƶq"].ToString();
                    drtemp5[WeekName + "_Q"] = sg;

                    System.Data.DataTable dh2 = null;
                    dh2 = GetTemp5_1t(StartDate, EndDate, "R", "", "");
                    sg2 = dh2.Rows[0]["���B"].ToString();
                    drtemp5[WeekName + "_R"] = sg2;

                    System.Data.DataTable dh3 = null;
                    dh3 = GetTemp5_1t(StartDate, EndDate, "C", "", "");
                    sg3 = dh3.Rows[0]["����"].ToString();
                    drtemp5[WeekName + "_C"] = sg3;

                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    int �P�����B = Convert.ToInt32(sg2);
                    double �Q��� = (�Q�� / (�P�����B)) * 100;


                    drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                    h = Convert.ToInt32(sg);
                    hR = Convert.ToInt32(sg2);
                    hC = Convert.ToInt32(sg3);
                    hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                    string aa = Convert.ToString(�Q���);
                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                    {
                        drtemp5[WeekName + "_P"] = 0;
                    }
                    else
                    {
                        drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                    }

                    hD += h;
                    hDR += hR;
                    hDC += hC;
                    hDG += hG;


                }
                drtemp5["WT_Q"] = hD;
                drtemp5["WT_R"] = hDR;
                drtemp5["WT_C"] = hDC;
                drtemp5["WT_G"] = hDG;

                DJ33 = hDR;
                hD = 0;
                hDR = 0;
                hDC = 0;
                hDG = 0;


                double �Q���5 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
                string aa5 = Convert.ToString(�Q���5);
                if (aa5 == "���L�a�j" || aa5 == "�t�L�a�j" || aa5 == "���O�@�ӼƦr" || aa5 == "�D�ƭ�" || aa5 == "��" || aa5 == "-��")
                {
                    drtemp5["WT_P"] = 0;
                }
                else
                {
                    drtemp5["WT_P"] = �Q���5.ToString("#0.00") + "%";
                }


                double SALES33 = (DJ33 / HJ) * 100;
                drtemp5["SALES2"] = SALES33.ToString("#0.00") + "%";


                dtCostDD3.Rows.Add(drtemp5);


                //�[�`
                drtemp5 = dtCostDD3.NewRow();
                drtemp5["�~��"] = "";
                drtemp5["�Ȥ�s��"] = "";
                drtemp5["�Ȥ�W��"] = "Grand Total";

                for (int w = 0; w <= p; w++)
                {
                    string sg = "";
                    string sg2 = "";
                    string sg3 = "";

                    int h = 0;
                    int hR = 0;
                    int hC = 0;
                    int hG = 0;

                    string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                    string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                    string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                    System.Data.DataTable dh = null;
                    dh = GetTemp5_1t(StartDate, EndDate, "Q", "f", GROUP);
                    sg = dh.Rows[0]["�ƶq"].ToString();
                    drtemp5[WeekName + "_Q"] = sg;

                    System.Data.DataTable dh2 = null;
                    dh2 = GetTemp5_1t(StartDate, EndDate, "R", "f", GROUP);
                    sg2 = dh2.Rows[0]["���B"].ToString();
                    drtemp5[WeekName + "_R"] = sg2;

                    System.Data.DataTable dh3 = null;
                    dh3 = GetTemp5_1t(StartDate, EndDate, "C", "f", GROUP);
                    sg3 = dh3.Rows[0]["����"].ToString();
                    drtemp5[WeekName + "_C"] = sg3;

                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    int �P�����B = Convert.ToInt32(sg2);
                    double �Q��� = (�Q�� / (�P�����B)) * 100;


                    drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    h = Convert.ToInt32(sg);
                    hR = Convert.ToInt32(sg2);
                    hC = Convert.ToInt32(sg3);
                    hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                    string aa = Convert.ToString(�Q���);
                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                    {
                        drtemp5[WeekName + "_P"] = 0;
                    }
                    else
                    {
                        drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                    }
                    hE += h;
                    hER += hR;
                    hEC += hC;
                    hEG += hG;

                }
                drtemp5["WT_Q"] = hE;
                drtemp5["WT_R"] = hER;
                drtemp5["WT_C"] = hEC;
                drtemp5["WT_G"] = hEG;
                hE = 0;
                hER = 0;
                hEC = 0;
                hEG = 0;


                double �Q���55 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
                string aa55 = Convert.ToString(�Q���55);
                if (aa55 == "���L�a�j" || aa55 == "�t�L�a�j" || aa55 == "���O�@�ӼƦr" || aa5 == "�D�ƭ�" || aa5 == "��" || aa5 == "-��")
                {
                    drtemp5["WT_P"] = 0;
                }
                else
                {
                    drtemp5["WT_P"] = �Q���55.ToString("#0.00") + "%";
                }

                dtCostDD3.Rows.Add(drtemp5);


                drtemp5 = dtCostDD3.NewRow();
                dtCostDD3.Rows.Add(drtemp5);

                drtemp5 = dtCostDD3.NewRow();
                drtemp5["�Ȥ�W��"] = "�~��";
                dtCostDD3.Rows.Add(drtemp5);



            }
            //SALES
            System.Data.DataTable dtemp7 = GetTemp5Sales(GROUP);




            for (int i = 0; i <= dtemp7.Rows.Count - 1; i++)
            {
                drtemp5 = dtCostDD3.NewRow();

                �~��t = dtemp7.Rows[i]["�~��"].ToString();
                BUt = dtemp7.Rows[i]["BU"].ToString();
                drtemp5["BU"] = BUt;
                drtemp5["�~��"] = "";
                drtemp5["�Ȥ�s��"] = "";
                drtemp5["�Ȥ�W��"] = �~��t;
                drtemp5["SALES2"] = "";
                for (int w = 0; w <= p; w++)
                {
                    string sg = "";
                    string sg2 = "";
                    string sg3 = "";

                    int h = 0;
                    int hR = 0;
                    int hC = 0;
                    int hG = 0;

                    string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                    string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                    string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                    System.Data.DataTable dh = null;
                    dh = GetTemp5_1SALES(StartDate, EndDate, �~��t, "Q", GROUP);
                    sg = dh.Rows[0]["�ƶq"].ToString();
                    drtemp5[WeekName + "_Q"] = sg;

                    System.Data.DataTable dh2 = null;
                    dh2 = GetTemp5_1SALES(StartDate, EndDate, �~��t, "R", GROUP);
                    sg2 = dh2.Rows[0]["���B"].ToString();
                    drtemp5[WeekName + "_R"] = sg2;

                    System.Data.DataTable dh3 = null;
                    dh3 = GetTemp5_1SALES(StartDate, EndDate, �~��t, "C", GROUP);
                    sg3 = dh3.Rows[0]["����"].ToString();
                    drtemp5[WeekName + "_C"] = sg3;

                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    int �P�����B = Convert.ToInt32(sg2);
                    double �Q��� = (�Q�� / (�P�����B)) * 100;


                    drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                    h = Convert.ToInt32(sg);
                    hR = Convert.ToInt32(sg2);
                    hC = Convert.ToInt32(sg3);
                    hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                    string aa = Convert.ToString(�Q���);
                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                    {
                        drtemp5[WeekName + "_P"] = 0;
                    }
                    else
                    {
                        drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                    }
                    hF += h;
                    hFR += hR;
                    hFC += hC;
                    hFG += hG;


                }
                drtemp5["WT_Q"] = hF;
                drtemp5["WT_R"] = hFR;
                drtemp5["WT_C"] = hFC;
                drtemp5["WT_G"] = hFG;
                DJ2 = hFR;
                hF = 0;
                hFR = 0;
                hFC = 0;
                hFG = 0;

                double �Q���1 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
                string aa2 = Convert.ToString(�Q���1);
                if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr" || aa2 == "�D�ƭ�" || aa2 == "��" || aa2 == "-��")
                {
                    drtemp5["WT_P"] = 0;
                }
                else
                {
                    drtemp5["WT_P"] = �Q���1.ToString("#0.00") + "%";
                }

                double SALES = (DJ2 / HJ) * 100;
                drtemp5["SALES2"] = SALES.ToString("#0.00") + "%";
                dtCostDD3.Rows.Add(drtemp5);
            }

            drtemp5 = dtCostDD3.NewRow();
            drtemp5["�~��"] = "";
            drtemp5["�Ȥ�s��"] = "";
            drtemp5["�Ȥ�W��"] = "�[�`";

            drtemp5["SALES2"] = "";

            for (int w = 0; w <= p; w++)
            {
                string sg = "";
                string sg2 = "";
                string sg3 = "";

                int h = 0;
                int hR = 0;
                int hC = 0;
                int hG = 0;

                string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);

                string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);

                System.Data.DataTable dh = null;
                dh = GetTemp5_1t(StartDate, EndDate, "Q", "N", GROUP);
                sg = dh.Rows[0]["�ƶq"].ToString();
                drtemp5[WeekName + "_Q"] = sg;

                System.Data.DataTable dh2 = null;
                dh2 = GetTemp5_1t(StartDate, EndDate, "R", "N", GROUP);
                sg2 = dh2.Rows[0]["���B"].ToString();
                drtemp5[WeekName + "_R"] = sg2;

                System.Data.DataTable dh3 = null;
                dh3 = GetTemp5_1t(StartDate, EndDate, "C", "N", GROUP);
                sg3 = dh3.Rows[0]["����"].ToString();
                drtemp5[WeekName + "_C"] = sg3;

                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                int �P�����B = Convert.ToInt32(sg2);
                double �Q��� = (�Q�� / (�P�����B)) * 100;


                drtemp5[WeekName + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                h = Convert.ToInt32(sg);
                hR = Convert.ToInt32(sg2);
                hC = Convert.ToInt32(sg3);
                hG = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);

                string aa = Convert.ToString(�Q���);
                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr" || aa == "�D�ƭ�" || aa == "��" || aa == "-��")
                {
                    drtemp5[WeekName + "_P"] = 0;
                }
                else
                {
                    drtemp5[WeekName + "_P"] = �Q���.ToString("#0.00") + "%";
                }
                hF += h;
                hFR += hR;
                hFC += hC;
                hFG += hG;

            }

            drtemp5["WT_Q"] = hF;
            drtemp5["WT_R"] = hFR;
            drtemp5["WT_C"] = hFC;
            drtemp5["WT_G"] = hFG;

            hF = 0;
            hFR = 0;
            hFC = 0;
            hFG = 0;

            double �Q���6 = (Convert.ToDouble(drtemp5["WT_G"]) / Convert.ToInt32(drtemp5["WT_R"])) * 100;
            string aa6 = Convert.ToString(�Q���6);
            if (aa6 == "���L�a�j" || aa6 == "�t�L�a�j" || aa6 == "���O�@�ӼƦr" || aa6 == "�D�ƭ�" || aa6 == "��" || aa6 == "-��")
            {
                drtemp5["WT_P"] = 0;
            }
            else
            {
                drtemp5["WT_P"] = �Q���6.ToString("#0.00") + "%";
            }

            dtCostDD3.Rows.Add(drtemp5);
            //
        }

        private void Eun22()
        {

            dtCostEun = MakeTableEUN32();
            DataRow drtemp5 = null;
            System.Data.DataTable dtWeek = MakeTableWeek();
            GetMonthWeekStartDate(textBox3.Text, dtWeek);


            for (int w = 0; w <= p; w++)
            {

                drtemp5 = dtCostEun.NewRow();


                string StartDate = Convert.ToString(dtWeek.Rows[w]["StartDate"]);
                string EndDate = Convert.ToString(dtWeek.Rows[w]["EndDate"]);
                string WeekName = Convert.ToString(dtWeek.Rows[w]["WeekName"]);
                string sg = "";
                string sg2 = "";
                string sg3 = "";
                string sg4 = "";
                string sg5 = "";
                string sg6 = "";
                string sg7 = "";
                string sg8 = "";
                string sg9 = "";
                drtemp5["WEEK"] = textBox3.Text.Substring(0, 4) + "-" + textBox3.Text.Substring(6, 2) + "-" + WeekName;

                System.Data.DataTable dh = null;
                dh = GetTemp5Eun(StartDate, EndDate, "Q", "N");
                if (dh.Rows.Count < 1)
                {
                    sg = "0";
                }
                else
                {
                    sg = dh.Rows[0]["�ƶq"].ToString();
                }

                System.Data.DataTable dh2 = null;
                dh2 = GetTemp5Eun(StartDate, EndDate, "R", "N");
                if (dh2.Rows.Count < 1)
                {
                    sg2 = "0";
                }
                else
                {
                    sg2 = dh2.Rows[0]["���B"].ToString();
                }
                System.Data.DataTable dh3 = null;
                dh3 = GetTemp5Eun(StartDate, EndDate, "C", "N");
                if (dh3.Rows.Count < 1)
                {
                    sg2 = "0";
                }
                else
                {
                    sg3 = dh3.Rows[0]["����"].ToString();
                }

                drtemp5["1_Q"] = sg;
                drtemp5["1_R"] = sg2;
                drtemp5["1_C"] = sg3;

                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                int �P�����B = Convert.ToInt32(sg2);
                double �Q��� = (�Q�� / (�P�����B));

                drtemp5["1_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                
                string aa = Convert.ToString(�Q���);
                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "�D�ƭ�" || aa == "���O�@�ӼƦr" || aa == "��" || aa == "-��")
                {
                    drtemp5["1_P"] = 0;
                }
                else
                {
                    drtemp5["1_P"] = �Q���;
                }

                System.Data.DataTable dh4 = null;
                dh4 = GetTemp5Eun(StartDate, EndDate, "Q", "Y");
                if (dh4.Rows.Count < 1)
                {
                    sg4 = "0";
                }
                else
                {
                    sg4 = dh4.Rows[0]["�ƶq"].ToString();
                }
                System.Data.DataTable dh5 = null;
                dh5 = GetTemp5Eun(StartDate, EndDate, "R", "Y");
                if (dh5.Rows.Count < 1)
                {
                    sg5 = "0";
                }
                else
                {
                    sg5 = dh5.Rows[0]["���B"].ToString();
                }
                System.Data.DataTable dh6 = null;
                dh6 = GetTemp5Eun(StartDate, EndDate, "C", "Y");
                if (dh6.Rows.Count < 1)
                {
                    sg6 = "0";
                }
                else
                {
                    sg6 = dh6.Rows[0]["����"].ToString();
                }

                drtemp5["2_Q"] = sg4;
                drtemp5["2_R"] = sg5;
                drtemp5["2_C"] = sg6;

                double �Q��2 = Convert.ToInt32(sg5) - Convert.ToInt32(sg6);
                int �P�����B2 = Convert.ToInt32(sg5);
                double �Q���2 = (�Q��2 / (�P�����B2));

                drtemp5["2_G"] = Convert.ToInt32(sg5) - Convert.ToInt32(sg6);

                string aa2 = Convert.ToString(�Q���2);

                //�D�ƭ�
                if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "�D�ƭ�" || aa2 == "���O�@�ӼƦr" || aa2 == "��" || aa2 == "-��")
                {
                    drtemp5["2_P"] = 0;
                }
                else
                {
                    drtemp5["2_P"] = �Q���2;
                }


                System.Data.DataTable dh7 = null;
                dh7 = GetTemp5Eun(StartDate, EndDate, "Q", "");
                if (dh7.Rows.Count < 1)
                {
                    sg7 = "0";
                }
                else
                {
                    sg7 = dh7.Rows[0]["�ƶq"].ToString();
                }

                System.Data.DataTable dh8 = null;
                dh8 = GetTemp5Eun(StartDate, EndDate, "R", "");
                if (dh8.Rows.Count < 1)
                {
                    sg8 = "0";
                }
                else
                {
                    sg8 = dh8.Rows[0]["���B"].ToString();
                }

                System.Data.DataTable dh9 = null;
                dh9 = GetTemp5Eun(StartDate, EndDate, "C", "");
                if (dh9.Rows.Count < 1)
                {
                    sg9 = "0";
                }
                else
                {
                    sg9 = dh9.Rows[0]["����"].ToString();
                }
                drtemp5["3_Q"] = sg7;
                drtemp5["3_R"] = sg8;
                drtemp5["3_C"] = sg9;

                double �Q��3 = Convert.ToInt32(sg8) - Convert.ToInt32(sg9);
                int �P�����B3 = Convert.ToInt32(sg8);
                double �Q���3 = (�Q��3 / (�P�����B3));

                drtemp5["3_G"] = Convert.ToInt32(sg8) - Convert.ToInt32(sg9);

                string aa3 = Convert.ToString(�Q���3);
                if (aa3 == "���L�a�j" || aa3 == "�t�L�a�j" || aa3 == "�D�ƭ�" || aa3 == "���O�@�ӼƦr" || aa3 == "��" || aa3 == "-��")
                {
                    drtemp5["3_P"] = 0;
                }
                else
                {
                    drtemp5["3_P"] = �Q���3;
                }

                dtCostEun.Rows.Add(drtemp5);

            }

        }



        public void AddAUOGD1(string TABLE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table " + TABLE + " ", connection);
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


        public void UPDATE2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE RMA_PARAMS SET PARAM_NO=@PARAM_NO  WHERE PARAM_KIND='EUNICE' ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", textBox5.Text.Trim().ToString()));
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
        public void UPDATEBALANCE(string DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE  OINV  SET U_BALANCE2=DOCTOTAL WHERE  DOCENTRY=@DOCENTRY ", connection);
            command.CommandType = CommandType.Text;
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
        public void AddAUOGD(string TABLE, string CARDCODE, string CARDNAME, string SALES, string GQty, string GTotal, string GValue, DateTime ���, string DOCENTRY, string ACCOUNT, string CARDGROUP, string COUNTRY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into " + TABLE + "(CARDCODE,CARDNAME,SALES,GQty,GTotal,GValue,DDATE,DOCENTRY,ACCOUNT,CARDGROUP,COUNTRY) values(@CARDCODE,@CARDNAME,@SALES,@GQty,@GTotal,@GValue,@DDATE,@DOCENTRY,@ACCOUNT,@CARDGROUP,@COUNTRY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@GQty", GQty));
            command.Parameters.Add(new SqlParameter("@GTotal", GTotal));
            command.Parameters.Add(new SqlParameter("@GValue", GValue));
            command.Parameters.Add(new SqlParameter("@DDATE", ���));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", CARDGROUP));
            command.Parameters.Add(new SqlParameter("@COUNTRY", COUNTRY));

            
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

        public void AddAUOGD61(string TABLE, string ITEMCODE, string ITEMNAME, string CARDCODE, string CARDNAME, string SALES, string GQty, string GTotal, string GValue, DateTime ���, string DOCENTRY, string ACCOUNT, string CARDGROUP, string GSUMTotal, string GSUMValue, string BASEDOC, string BASELINE, string COUNTRY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into " + TABLE + "(ITEMCODE,ITEMNAME,CARDCODE,CARDNAME,SALES,GQty,GTotal,GValue,DDATE,DOCENTRY,ACCOUNT,CARDGROUP,GSUMTotal,GSUMValue,BASEDOC,BASELINE,COUNTRY) values(@ITEMCODE,@ITEMNAME,@CARDCODE,@CARDNAME,@SALES,@GQty,@GTotal,@GValue,@DDATE,@DOCENTRY,@ACCOUNT,@CARDGROUP,@GSUMTotal,@GSUMValue,@BASEDOC,@BASELINE,@COUNTRY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@GQty", GQty));
            command.Parameters.Add(new SqlParameter("@GTotal", GTotal));
            command.Parameters.Add(new SqlParameter("@GValue", GValue));
            command.Parameters.Add(new SqlParameter("@DDATE", ���));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@ACCOUNT", ACCOUNT));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", CARDGROUP));
            command.Parameters.Add(new SqlParameter("@GSUMTotal", GSUMTotal));
            command.Parameters.Add(new SqlParameter("@GSUMValue", GSUMValue));
            command.Parameters.Add(new SqlParameter("@BASEDOC", BASEDOC));
            command.Parameters.Add(new SqlParameter("@BASELINE", BASELINE));
            command.Parameters.Add(new SqlParameter("@COUNTRY", COUNTRY));
            
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



       
        private void AddToDtWeek(System.Data.DataTable dt, string WeekName, string StartDate, string EndDate)
        {
            DataRow dr;
            dr = dt.NewRow();
            dr["WeekName"] = WeekName;
            dr["StartDate"] = StartDate;
            dr["EndDate"] = EndDate;
            dt.Rows.Add(dr);
        }
        private void GetMonthWeekStartDate(string CurrentMonth, System.Data.DataTable dtWeek)
        {
            string Month = CurrentMonth.Substring(0, 6);

            int iear = Convert.ToInt32(Month.Substring(0, 4));
            int iMonth = Convert.ToInt32(Month.Substring(4, 2));

            string EndDayOfMonth = Month + DateTime.DaysInMonth(iear, iMonth).ToString("00");
            string StartDayOfMonth = Month + "01";

            //�Ĥ@�g�������I
            DateTime dt = GetMenu.StrToDate(StartDayOfMonth);

            int iWeek = 1;

            string sDate = StartDayOfMonth;
            string eDate = "";



            while (dt <= GetMenu.StrToDate(EndDayOfMonth))
            {
                if (dt.DayOfWeek == DayOfWeek.Sunday)
                {
                    eDate = dt.ToString("yyyyMMdd");
                    AddToDtWeek(dtWeek, "W" + iWeek.ToString("0"), sDate, eDate);

                    sDate = dt.AddDays(1).ToString("yyyyMMdd");

                    iWeek++;
                }
                dt = dt.AddDays(1);
            }

            if (eDate != EndDayOfMonth)
            {
                eDate = EndDayOfMonth;
                AddToDtWeek(dtWeek, "W" + iWeek.ToString("0"), sDate, eDate);
            }
            p = dtWeek.Rows.Count - 1;
        }
        private System.Data.DataTable MakeTableFcstWeek()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("row", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�~��", typeof(string));

            dt.Columns.Add("W1_Q", typeof(decimal));
            dt.Columns.Add("W1_R", typeof(decimal));
            dt.Columns.Add("W1_C", typeof(decimal));
            dt.Columns.Add("W1_G", typeof(decimal));
            dt.Columns.Add("W1_P", typeof(string));
            dt.Columns.Add("W2_Q", typeof(decimal));
            dt.Columns.Add("W2_R", typeof(decimal));
            dt.Columns.Add("W2_C", typeof(decimal));
            dt.Columns.Add("W2_G", typeof(decimal));
            dt.Columns.Add("W2_P", typeof(string));
            dt.Columns.Add("W3_Q", typeof(decimal));
            dt.Columns.Add("W3_R", typeof(decimal));
            dt.Columns.Add("W3_C", typeof(decimal));
            dt.Columns.Add("W3_G", typeof(decimal));
            dt.Columns.Add("W3_P", typeof(string));
            dt.Columns.Add("W4_Q", typeof(decimal));
            dt.Columns.Add("W4_R", typeof(decimal));
            dt.Columns.Add("W4_C", typeof(decimal));
            dt.Columns.Add("W4_G", typeof(decimal));
            dt.Columns.Add("W4_P", typeof(string));
            dt.Columns.Add("W5_Q", typeof(decimal));
            dt.Columns.Add("W5_R", typeof(decimal));
            dt.Columns.Add("W5_C", typeof(decimal));
            dt.Columns.Add("W5_G", typeof(decimal));
            dt.Columns.Add("W5_P", typeof(string));
            dt.Columns.Add("W6_Q", typeof(decimal));
            dt.Columns.Add("W6_R", typeof(decimal));
            dt.Columns.Add("W6_C", typeof(decimal));
            dt.Columns.Add("W6_G", typeof(decimal));
            dt.Columns.Add("W6_P", typeof(string));
            dt.Columns.Add("WT_Q", typeof(decimal));
            dt.Columns.Add("WT_R", typeof(decimal));
            dt.Columns.Add("WT_C", typeof(decimal));
            dt.Columns.Add("WT_G", typeof(decimal));
            dt.Columns.Add("WT_P", typeof(string));
            dt.Columns.Add("SALES2", typeof(string));
            dt.Columns.Add("SALES22", typeof(string));
            return dt;

        }

        private System.Data.DataTable MakeTableFcstear()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("row", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("�~��", typeof(string));
            dt.Columns.Add("1_Q", typeof(decimal));
            dt.Columns.Add("1_R", typeof(decimal));
            dt.Columns.Add("1_C", typeof(decimal));
            dt.Columns.Add("1_G", typeof(decimal));
            dt.Columns.Add("1_P", typeof(string));
            dt.Columns.Add("2_Q", typeof(decimal));
            dt.Columns.Add("2_R", typeof(decimal));
            dt.Columns.Add("2_C", typeof(decimal));
            dt.Columns.Add("2_G", typeof(decimal));
            dt.Columns.Add("2_P", typeof(string));
            dt.Columns.Add("3_Q", typeof(decimal));
            dt.Columns.Add("3_R", typeof(decimal));
            dt.Columns.Add("3_C", typeof(decimal));
            dt.Columns.Add("3_G", typeof(decimal));
            dt.Columns.Add("3_P", typeof(string));
            dt.Columns.Add("4_Q", typeof(decimal));
            dt.Columns.Add("4_R", typeof(decimal));
            dt.Columns.Add("4_C", typeof(decimal));
            dt.Columns.Add("4_G", typeof(decimal));
            dt.Columns.Add("4_P", typeof(string));
            dt.Columns.Add("5_Q", typeof(decimal));
            dt.Columns.Add("5_R", typeof(decimal));
            dt.Columns.Add("5_C", typeof(decimal));
            dt.Columns.Add("5_G", typeof(decimal));
            dt.Columns.Add("5_P", typeof(string));
            dt.Columns.Add("6_Q", typeof(decimal));
            dt.Columns.Add("6_R", typeof(decimal));
            dt.Columns.Add("6_C", typeof(decimal));
            dt.Columns.Add("6_G", typeof(decimal));
            dt.Columns.Add("6_P", typeof(string));
            dt.Columns.Add("7_Q", typeof(decimal));
            dt.Columns.Add("7_R", typeof(decimal));
            dt.Columns.Add("7_C", typeof(decimal));
            dt.Columns.Add("7_G", typeof(decimal));
            dt.Columns.Add("7_P", typeof(string));
            dt.Columns.Add("8_Q", typeof(decimal));
            dt.Columns.Add("8_R", typeof(decimal));
            dt.Columns.Add("8_C", typeof(decimal));
            dt.Columns.Add("8_G", typeof(decimal));
            dt.Columns.Add("8_P", typeof(string));
            dt.Columns.Add("9_Q", typeof(decimal));
            dt.Columns.Add("9_R", typeof(decimal));
            dt.Columns.Add("9_C", typeof(decimal));
            dt.Columns.Add("9_G", typeof(decimal));
            dt.Columns.Add("9_P", typeof(string));
            dt.Columns.Add("10_Q", typeof(decimal));
            dt.Columns.Add("10_R", typeof(decimal));
            dt.Columns.Add("10_C", typeof(decimal));
            dt.Columns.Add("10_G", typeof(decimal));
            dt.Columns.Add("10_P", typeof(string));
            dt.Columns.Add("11_Q", typeof(decimal));
            dt.Columns.Add("11_R", typeof(decimal));
            dt.Columns.Add("11_C", typeof(decimal));
            dt.Columns.Add("11_G", typeof(decimal));
            dt.Columns.Add("11_P", typeof(string));
            dt.Columns.Add("12_Q", typeof(decimal));
            dt.Columns.Add("12_R", typeof(decimal));
            dt.Columns.Add("12_C", typeof(decimal));
            dt.Columns.Add("12_G", typeof(decimal));
            dt.Columns.Add("12_P", typeof(string));
            dt.Columns.Add("T_Q", typeof(decimal));
            dt.Columns.Add("T_R", typeof(decimal));
            dt.Columns.Add("T_C", typeof(decimal));
            dt.Columns.Add("T_G", typeof(decimal));
            dt.Columns.Add("T_P", typeof(string));
            dt.Columns.Add("SALES2", typeof(string));
            dt.Columns.Add("SALES22", typeof(string));
            return dt;

        }
        private System.Data.DataTable MakeTableQuar()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("row", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("�~��", typeof(string));
            dt.Columns.Add("�Ȥ�s��", typeof(string));
            dt.Columns.Add("�Ȥ�W��", typeof(string));
            dt.Columns.Add("1_Q", typeof(decimal));
            dt.Columns.Add("1_R", typeof(decimal));
            dt.Columns.Add("1_C", typeof(decimal));
            dt.Columns.Add("1_G", typeof(decimal));
            dt.Columns.Add("1_P", typeof(string));
            dt.Columns.Add("2_Q", typeof(decimal));
            dt.Columns.Add("2_R", typeof(decimal));
            dt.Columns.Add("2_C", typeof(decimal));
            dt.Columns.Add("2_G", typeof(decimal));
            dt.Columns.Add("2_P", typeof(string));
            dt.Columns.Add("3_Q", typeof(decimal));
            dt.Columns.Add("3_R", typeof(decimal));
            dt.Columns.Add("3_C", typeof(decimal));
            dt.Columns.Add("3_G", typeof(decimal));
            dt.Columns.Add("3_P", typeof(string));
            dt.Columns.Add("4_Q", typeof(decimal));
            dt.Columns.Add("4_R", typeof(decimal));
            dt.Columns.Add("4_C", typeof(decimal));
            dt.Columns.Add("4_G", typeof(decimal));
            dt.Columns.Add("4_P", typeof(string));
            dt.Columns.Add("T_Q", typeof(decimal));
            dt.Columns.Add("T_R", typeof(decimal));
            dt.Columns.Add("T_C", typeof(decimal));
            dt.Columns.Add("T_G", typeof(decimal));
            dt.Columns.Add("T_P", typeof(string));

            return dt;

        }
        private System.Data.DataTable MakeTableMQY()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NAME", typeof(string));


            dt.Columns.Add("1_M", typeof(string));
            dt.Columns.Add("2_M", typeof(string));
            dt.Columns.Add("3_M", typeof(string));
            dt.Columns.Add("4_M", typeof(string));
            dt.Columns.Add("5_M", typeof(string));
            dt.Columns.Add("6_M", typeof(string));
            dt.Columns.Add("7_M", typeof(string));
            dt.Columns.Add("8_M", typeof(string));
            dt.Columns.Add("9_M", typeof(string));
            dt.Columns.Add("10_M", typeof(string));
            dt.Columns.Add("11_M", typeof(string));
            dt.Columns.Add("12_M", typeof(string));
            dt.Columns.Add("1_Q", typeof(string));
            dt.Columns.Add("2_Q", typeof(string));
            dt.Columns.Add("3_Q", typeof(string));
            dt.Columns.Add("4_Q", typeof(string));
            dt.Columns.Add("Y", typeof(string));

            return dt;

        }
        private System.Data.DataTable MakeTableEUN32()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("WEEK", typeof(string));

            dt.Columns.Add("1_Q", typeof(decimal));
            dt.Columns.Add("1_R", typeof(decimal));
            dt.Columns.Add("1_C", typeof(decimal));
            dt.Columns.Add("1_G", typeof(decimal));
            dt.Columns.Add("1_P", typeof(decimal));
            dt.Columns.Add("2_Q", typeof(decimal));
            dt.Columns.Add("2_R", typeof(decimal));
            dt.Columns.Add("2_C", typeof(decimal));
            dt.Columns.Add("2_G", typeof(decimal));
            dt.Columns.Add("2_P", typeof(decimal));
            dt.Columns.Add("3_Q", typeof(decimal));
            dt.Columns.Add("3_R", typeof(decimal));
            dt.Columns.Add("3_C", typeof(decimal));
            dt.Columns.Add("3_G", typeof(decimal));
            dt.Columns.Add("3_P", typeof(decimal));

            return dt;

        }
        private System.Data.DataTable MakeTableAvg()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("����", typeof(string));
            dt.Columns.Add("����", typeof(string));
            dt.Columns.Add("�ƶq", typeof(string));
            return dt;

        }
        private System.Data.DataTable MakeTableAcc()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("��ئW��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(string));
            dt.Columns.Add("Debit", typeof(string));
            dt.Columns.Add("Credit", typeof(string));
            dt.Columns.Add("Balance", typeof(string));
            dt.Columns.Add("����", typeof(string));

            return dt;

        }
        private System.Data.DataTable MakeTableOINV()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("���", typeof(string));
            dt.Columns.Add("��ئW��", typeof(string));
            dt.Columns.Add("�ƶq", typeof(string));
            dt.Columns.Add("Debit", typeof(string));
            dt.Columns.Add("Credit", typeof(string));
            dt.Columns.Add("Balance", typeof(string));
            dt.Columns.Add("����", typeof(string));

            return dt;

        }
        System.Data.DataTable GetTemp5(string choice, string GROUP)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT T0.CARDCODE �Ȥ�s��,T0.CARDNAME �Ȥ�W��,T0.SALES �~��,MAX(T1.BU) BU FROM Account_Temp6 T0 ");
            sb.Append(" LEFT JOIN Account_TempSALES T1 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T1.SALES) ");
            sb.Append(" where 1=1");


            if (choice == "N")
            {
                sb.Append(" and  T0.CARDCODE not in  ('0511-00','0257-00') ");
            }
            if (choice == "")
            {
                sb.Append(" and  T0.CARDCODE  in  ('0511-00','0257-00') ");
            }

            if (GROUP != "")
            {

                if (GROUP == "���Ʒ~��")
                {
                    sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��' )");
                }
                else
                {
                    sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                }
            }
            sb.Append(" group by  T0.CARDCODE,T0.CARDNAME,T0.SALES ");
            sb.Append(" order by sum(cast(gtotal as int))  desc");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



            command.Parameters.Add(new SqlParameter("@CARDGROUP", GROUP));
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
        System.Data.DataTable GetTemp61(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT T0.CARDCODE �Ȥ�s��,T0.CARDNAME �Ȥ�W��,T0.SALES �~��,MAX(T1.BU) BU  FROM   " + Account_Temp6 + "   T0");
            sb.Append(" LEFT JOIN Account_TempSALES T1 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T1.SALES)             ");
            sb.Append(" group by  T0.CARDCODE,T0.CARDNAME,T0.SALES");
            sb.Append(" order by sum(cast(gtotal as int))  desc    ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox2.Text));

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

        System.Data.DataTable GetTemp5Sales(string GROUP)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  T0.SALES �~��,MAX(T1.BU)�@BU FROM Account_Temp6�@T0");
            sb.Append("  LEFT JOIN Account_TempSALES T1 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T1.SALES) ");
            sb.Append("   where CARDCODE not in  ('0511-00','0257-00')");

            if (GROUP != "")
            {
                if (GROUP == "���Ʒ~��")
                {
                    sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��' ) ");
                }
                else
                {
                    sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                }
            }

            sb.Append("   group by T0.SALES  order by sum(cast(gtotal as int)) desc ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@CARDGROUP", GROUP));
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

        System.Data.DataTable GetTemp5_1(string startdate, string enddate, string cardcode, string sales, string group)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and cardcode=@cardcode and sales=@sales ");
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and cardcode=@cardcode and sales=@sales ");

            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and cardcode=@cardcode and sales=@sales ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startdate", startdate));
            command.Parameters.Add(new SqlParameter("@enddate", enddate));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@sales", sales));
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
        System.Data.DataTable GetTemp5Eun(string startdate, string enddate, string group, string choice)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate  ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "Y")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }

            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "Y")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }
            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate  ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "Y")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startdate", startdate));
            command.Parameters.Add(new SqlParameter("@enddate", enddate));

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
        System.Data.DataTable GetTemp5_1M(int MONTH, string cardcode, string sales, string group,string ITEMCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp61 where YEAR(ddate)=@YEAR AND month(ddate) = @MONTH and cardcode=@cardcode and sales=@sales AND ITEMCODE=@ITEMCODE  ");
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp61 where YEAR(ddate)=@YEAR AND month(ddate) = @MONTH and cardcode=@cardcode and sales=@sales AND ITEMCODE=@ITEMCODE ");

            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp61 where YEAR(ddate)=@YEAR AND  month(ddate) = @MONTH and cardcode=@cardcode and sales=@sales AND ITEMCODE=@ITEMCODE ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@sales", sales));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox2.Text));
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

        System.Data.DataTable GetMQY(string group, int ACCOUNT, string YEAR, int MONTH)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ISNULL(CAST( SUM(T0.[credit])-SUM(T0.[debit]) AS DECIMAL),0) ���B ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId WHERE");
            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (ACCOUNT == 4)
            {
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4')");
            }
            if (ACCOUNT == 5)
            {
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('5')");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));

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

        System.Data.DataTable GetMQY2(string group, string YEAR, int MONTH, string ACCTNAME)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T3.ACCTNAME ���,ISNULL(CAST( SUM(T0.[credit])-SUM(T0.[debit]) AS DECIMAL),0) ���B ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId");
            sb.Append("            Left join OACT t3 on (t0.[Account]=t3.ACCTCODE)  WHERE  ");           
            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR AND T3.ACCTNAME=@ACCTNAME  ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  AND T3.ACCTNAME=@ACCTNAME ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Y2")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR AND T3.ACCTNAME=@ACCTNAME    ");
            }
            sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4')");
            sb.Append(" GROUP BY ACCOUNT,T3.ACCTNAME ORDER BY t0.[Account] ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@ACCTNAME", ACCTNAME));

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



        System.Data.DataTable GetMQY3(string group, string YEAR, int MONTH)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(���B)+SUM(����) ���B,CASE WHEN SUM(���B) =0 THEN 0   ELSE ((SUM(���B)+SUM(����))/SUM(���B))*100  END ��Q from (  ");
            sb.Append(" SELECT ISNULL(CAST( SUM(T0.[credit])-SUM(T0.[debit]) AS DECIMAL),0) ���B,0 ���� ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId WHERE");
            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
           
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4')");
            



            sb.Append(" UNION ALL  ");
            sb.Append(" SELECT 0,ISNULL(CAST( SUM(T0.[credit])-SUM(T0.[debit]) AS DECIMAL),0) ���B ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId WHERE");
            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
       
         
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('5') ) as aa");
            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));

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

        System.Data.DataTable GetMQY4(string group, string YEAR, int MONTH, string ocrname, int ACCOUNT)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  SUM(T0.[credit])-SUM(T0.[debit]) ���B ,ISNULL(ocrname,'') ����");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId");
            sb.Append("              Left join oocr t2 on (t0.profitcode=t2.ocrcode) WHERE  ");


            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname  ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  AND ISNULL(ocrname,'')=@ocrname ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Y2")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname    ");
            }
            if (ACCOUNT == 4)
            {
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4')");
            }
            if (ACCOUNT == 5)
            {
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('5')");
            }
            sb.Append("              GROUP BY ocrname,profitcode order by profitcode");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@ocrname", ocrname));

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



        System.Data.DataTable GetMQY5(string group, string YEAR, int MONTH, string ocrname)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT  SUM(���J)+SUM(����) ��Q,����,profitcode,CASE WHEN SUM(���J) =0 THEN 0   ELSE ((SUM(���J)+SUM(����))/SUM(���J))*100  END ��Q�v  FROM ( ");
            sb.Append(" SELECT  SUM(T0.[credit])-SUM(T0.[debit]) ���J ,0 ����,ISNULL(ocrname,'') ����,profitcode ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId");
            sb.Append("              Left join oocr t2 on (t0.profitcode=t2.ocrcode) WHERE  ");


            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname  ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  AND ISNULL(ocrname,'')=@ocrname ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Y2")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname    ");
            }
        
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4')");
            
 
            sb.Append("              GROUP BY ocrname,profitcode ");

            sb.Append("       UNION ALL ");

            sb.Append(" SELECT  0  ���J,SUM(T0.[credit])-SUM(T0.[debit]) ���� ,ISNULL(ocrname,'') ����,profitcode ");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId");
            sb.Append("              Left join oocr t2 on (t0.profitcode=t2.ocrcode) WHERE  ");


            if (group == "M")
            {
                sb.Append("               MONTH(T0.[RefDate]) = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname  ");
            }
            if (group == "Q")
            {


                sb.Append("  CASE WHEN month(T0.[RefDate]) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(T0.[RefDate]) BETWEEN 10 AND 12 THEN 4 END = @MONTH AND   YEAR(T0.[RefDate]) = @YEAR  AND ISNULL(ocrname,'')=@ocrname ");
            }
            if (group == "Y")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR   ");
            }
            if (group == "Y2")
            {
                sb.Append("               YEAR(T0.[RefDate]) = @YEAR AND ISNULL(ocrname,'')=@ocrname    ");
            }
          
                sb.Append(" AND  T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('5')");

                sb.Append("              GROUP BY ocrname,profitcode ) AS AA GROUP BY ����,profitcode  ORDER BY profitcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", YEAR));
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@ocrname", ocrname));

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
        System.Data.DataTable GetTemp5_1Q(int MONTH, string cardcode, string sales, string group,string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM " + Account_Temp6 + " where   cardcode=@cardcode and sales=@sales ");
                sb.Append(" and CASE WHEN month(ddate) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(ddate) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(ddate) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(ddate) BETWEEN 10 AND 12 THEN 4 END = @MONTH  ");
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM " + Account_Temp6 + " where cardcode=@cardcode and sales=@sales ");
                sb.Append(" and CASE WHEN month(ddate) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(ddate) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(ddate) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(ddate) BETWEEN 10 AND 12 THEN 4 END = @MONTH  ");
            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM " + Account_Temp6 + " where cardcode=@cardcode and sales=@sales ");
                sb.Append(" and CASE WHEN month(ddate) BETWEEN 1 AND 3 THEN 1 ");
                sb.Append("            WHEN month(ddate) BETWEEN 4 AND 6 THEN 2");
                sb.Append("            WHEN month(ddate) BETWEEN 7 AND 9 THEN 3");
                sb.Append("            WHEN month(ddate) BETWEEN 10 AND 12 THEN 4 END = @MONTH ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

    
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@cardcode", cardcode));
            command.Parameters.Add(new SqlParameter("@sales", sales));
          
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


       
        System.Data.DataTable GetTemp5_1SALES(string startdate, string enddate, string sales, string group, string GROUP)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and sales=@sales   and  CARDCODE not in  ('0511-00','0257-00') ");
                if (GROUP != "")
                {
                    if (GROUP == "���Ʒ~��")
                    {
                        sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��' ) ");
                    }
                    else
                    {
                        sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                    }
                }
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and sales=@sales  and  CARDCODE not in  ('0511-00','0257-00')");
                if (GROUP != "")
                {
                    if (GROUP == "���Ʒ~��")
                    {
                        sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��') ");
                    }
                    else
                    {
                        sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                    }
                }
            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate and  sales=@sales  and  CARDCODE not in  ('0511-00','0257-00') ");
                if (GROUP != "")
                {
                    if (GROUP == "���Ʒ~��")
                    {
                        sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��' ) ");
                    }
                    else
                    {
                        sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                    }
                }
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startdate", startdate));
            command.Parameters.Add(new SqlParameter("@enddate", enddate));
            command.Parameters.Add(new SqlParameter("@sales", sales));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", GROUP));
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

        System.Data.DataTable GetTemp5_1SALESM(int MONTH, string sales, string group)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp6 where  month(ddate)= @MONTH and sales=@sales   and  CARDCODE not in  ('0511-00','0257-00') ");
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp6 where  month(ddate)= @MONTH and sales=@sales  and  CARDCODE not in  ('0511-00','0257-00')");

            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp6 where  month(ddate)= @MONTH and  sales=@sales  and  CARDCODE not in  ('0511-00','0257-00') ");

            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MONTH", MONTH));
            command.Parameters.Add(new SqlParameter("@sales", sales));
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
        System.Data.DataTable GetTemp5_1TOTAL()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT isnull(sum(cast(GTOTAL as float)),0) ���B FROM Account_Temp6 where CARDCODE not in  ('0511-00','0257-00')   ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        System.Data.DataTable GetTemp5_1t(string startdate, string enddate, string group, string choice, string GROUP1)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            if (group == "Q")
            {
                sb.Append(" SELECT isnull(sum(cast(GQTY as int)),0) �ƶq FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate  ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }

                if (GROUP1 == "���Ʒ~��")
                {
                    sb.Append(" and ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��') ");
                }
                else
                {
                    sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                }
            }
            if (group == "R")
            {
                sb.Append(" SELECT isnull(sum(cast(GTOTAL as int)),0) ���B FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate   ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }

                if (GROUP1 == "���Ʒ~��")
                {
                    sb.Append(" and  ( CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��' )");
                }
                else
                {
                    sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                }
            }
            if (group == "C")
            {
                sb.Append(" SELECT isnull(sum(cast(GVALUE as int)),0) ���� FROM Account_Temp6 where  Convert(varchar(8),ddate,112) between @startdate and @enddate   ");
                if (choice == "N")
                {
                    sb.Append(" and  CARDCODE not in  ('0511-00','0257-00') ");
                }
                if (choice == "")
                {
                    sb.Append(" and  CARDCODE  in  ('0511-00','0257-00') ");
                }

                if (GROUP1 == "���Ʒ~��")
                {
                    sb.Append(" and   (CARDGROUP LIKE '%���%'  or CARDGROUP='�Ӷ���Ʒ~��') ");
                }
                else
                {
                    sb.Append(" and   CARDGROUP=@CARDGROUP  ");
                }
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startdate", startdate));
            command.Parameters.Add(new SqlParameter("@enddate", enddate));
            command.Parameters.Add(new SqlParameter("@CARDGROUP", GROUP1));
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

        private System.Data.DataTable MakeTableWeek()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("WeekName", typeof(string));
            dt.Columns.Add("StartDate", typeof(string));
            dt.Columns.Add("EndDate", typeof(string));


            return dt;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                A1 = comboBox2.Text + "0101";
                A2 = comboBox2.Text + "1231";
             //   Category(6, comboBox2.SelectedValue.ToString(), "Account_Temp61");

                string YEAR = comboBox8.Text;
                string �Ȥ�s��t;
                string �Ȥ�W��t;
                string �~��t;
                string ���~�s��;
                string ���~�W��;
                string MODEL;
                System.Data.DataTable dtemp5 = GetTemp61("Account_Temp61" + YEAR);
                System.Data.DataTable dtCostDD = MakeTableFcstear();

                DataRow drtemp5 = null;
     
                for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
                {
                    j = i + 1;
                    drtemp5 = dtCostDD.NewRow();
                    �Ȥ�s��t = dtemp5.Rows[i]["�Ȥ�s��"].ToString();
                    �Ȥ�W��t = dtemp5.Rows[i]["�Ȥ�W��"].ToString();
                    �~��t = dtemp5.Rows[i]["�~��"].ToString();
                    ���~�s�� = dtemp5.Rows[i]["���~�s��"].ToString();
                  //  ���~�W�� = dtemp5.Rows[i]["���~�W��"].ToString();
                    MODEL = dtemp5.Rows[i]["MODEL"].ToString();
                    drtemp5["row"] = j.ToString();
                    drtemp5["�~��"] = �~��t;
                    drtemp5["�Ȥ�s��"] = "'" + �Ȥ�s��t;
                    drtemp5["�Ȥ�W��"] = �Ȥ�W��t;
                    drtemp5["MODEL"] = MODEL;
                    drtemp5["���~�s��"] = ���~�s��;
              //      drtemp5["���~�W��"] = ���~�W��;

                    for (int y = 1; y <= 12; y++)
                    {
                        string sg = "";
                        string sg2 = "";
                        string sg3 = "";

                        System.Data.DataTable dh = null;
                        dh = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "Q", ���~�s��);
                        sg = dh.Rows[0]["�ƶq"].ToString();
                        drtemp5[y + "_Q"] = sg;

                        System.Data.DataTable dh2 = null;
                        dh2 = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "R", ���~�s��);
                        sg2 = dh2.Rows[0]["���B"].ToString();
                        drtemp5[y + "_R"] = sg2;

                        System.Data.DataTable dh3 = null;
                        dh3 = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "C", ���~�s��);
                        sg3 = dh3.Rows[0]["����"].ToString();
                        drtemp5[y + "_C"] = sg3;

                        double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                        int �P�����B = Convert.ToInt32(sg2);
                        double �Q��� = (�Q�� / (�P�����B)) * 100;


                        drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                        string aa = Convert.ToString(�Q���);
                        if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
                        {
                            drtemp5[y + "_P"] = 0;
                        }
                        else
                        {
                            drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
                        }

                    }

                    drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    + Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    + Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
                    DJ = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    + Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    + Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
                    drtemp5["T_R"] = DJ;
                    drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    + Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    + Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
                    drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    + Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    + Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);


                    double �Q���1 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
                    string aa2 = Convert.ToString(�Q���1);
                    if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr")
                    {
                        drtemp5["T_P"] = 0;
                    }
                    else
                    {
                        drtemp5["T_P"] = �Q���1.ToString("#0.00") + "%";
                    }


        
                   dtCostDD.Rows.Add(drtemp5);
                }


    //            drtemp5 = dtCostDD.NewRow();
    //            drtemp5["�~��"] = "";
    //            drtemp5["�Ȥ�s��"] = "";
    //            drtemp5["�Ȥ�W��"] = "3rd parties sub total";
    //            for (int y = 1; y <= 12; y++)
    //            {

    //                string sg = "";
    //                string sg2 = "";
    //                string sg3 = "";
    //                System.Data.DataTable dh = null;
    //                dh = GetTemp5_1tM(y, "Q");
    //                sg = dh.Rows[0]["�ƶq"].ToString();
    //                drtemp5[y + "_Q"] = sg;

    //                System.Data.DataTable dh2 = null;
    //                dh2 = GetTemp5_1tM(y, "R");
    //                sg2 = dh2.Rows[0]["���B"].ToString();
    //                drtemp5[y + "_R"] = sg2;

    //                System.Data.DataTable dh3 = null;
    //                dh3 = GetTemp5_1tM(y, "C");
    //                sg3 = dh3.Rows[0]["����"].ToString();
    //                drtemp5[y + "_C"] = sg3;

    //                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                int �P�����B = Convert.ToInt32(sg2);
    //                double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                string aa = Convert.ToString(�Q���);
    //                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                {
    //                    drtemp5[y + "_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                }

    //            }

    //            drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //+ Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //+ Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //            drtemp5["T_R"] = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //+ Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //+ Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //            drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //+ Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //+ Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //            drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //+ Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //+ Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);


    //            double �Q���2 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //            string aa3 = Convert.ToString(�Q���2);
    //            if (aa3 == "���L�a�j" || aa3 == "�t�L�a�j" || aa3 == "���O�@�ӼƦr")
    //            {
    //                drtemp5["T_P"] = 0;
    //            }
    //            else
    //            {
    //                drtemp5["T_P"] = �Q���2.ToString("#0.00") + "%";
    //            }

    //            dtCostDD.Rows.Add(drtemp5);


    //            //CHOICE
    //            System.Data.DataTable dtemp6 = GetTemp5("", "");




    //            for (int i = 0; i <= dtemp6.Rows.Count - 1; i++)
    //            {
    //                drtemp5 = dtCostDD.NewRow();
    //                �Ȥ�s��t = dtemp6.Rows[i]["�Ȥ�s��"].ToString();
    //                �Ȥ�W��t = dtemp6.Rows[i]["�Ȥ�W��"].ToString();
    //                �~��t = dtemp6.Rows[i]["�~��"].ToString();
    //                drtemp5["�~��"] = �~��t;
    //                drtemp5["�Ȥ�s��"] = "'" + �Ȥ�s��t;
    //                drtemp5["�Ȥ�W��"] = �Ȥ�W��t;



    //                for (int y = 1; y <= 12; y++)
    //                {
    //                    string sg = "";
    //                    string sg2 = "";
    //                    string sg3 = "";

    //                    System.Data.DataTable dh = null;
    //                    dh = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "Q");
    //                    sg = dh.Rows[0]["�ƶq"].ToString();
    //                    drtemp5[y + "_Q"] = sg;

    //                    System.Data.DataTable dh2 = null;
    //                    dh2 = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "R");
    //                    sg2 = dh2.Rows[0]["���B"].ToString();
    //                    drtemp5[y + "_R"] = sg2;

    //                    System.Data.DataTable dh3 = null;
    //                    dh3 = GetTemp5_1M(y, �Ȥ�s��t, �~��t, "C");
    //                    sg3 = dh3.Rows[0]["����"].ToString();
    //                    drtemp5[y + "_C"] = sg3;

    //                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                    int �P�����B = Convert.ToInt32(sg2);
    //                    double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                    drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                    string aa = Convert.ToString(�Q���);
    //                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                    {
    //                        drtemp5[y + "_P"] = 0;
    //                    }
    //                    else
    //                    {
    //                        drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                    }

    //                }

    //                drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //+ Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //+ Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //                DJ3 = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //    + Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //    + Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //                drtemp5["T_R"] = DJ3;
    //                drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //    + Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //    + Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //                drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //    + Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //    + Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);

    //                double �Q���1 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //                string aa2 = Convert.ToString(�Q���1);
    //                if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr")
    //                {
    //                    drtemp5["T_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5["T_P"] = �Q���1.ToString("#0.00") + "%";
    //                }

    //                double SALES = (DJ3 / HJ) * 100;
    //                drtemp5["SALES2"] = SALES.ToString("#0.00") + "%";

    //                dtCostDD.Rows.Add(drtemp5);
    //            }

    //            //choice�[�`
    //            drtemp5 = dtCostDD.NewRow();
    //            drtemp5["�~��"] = "";
    //            drtemp5["�Ȥ�s��"] = "";
    //            drtemp5["�Ȥ�W��"] = "Related Parties sub total";


    //            for (int y = 1; y <= 12; y++)
    //            {

    //                string sg = "";
    //                string sg2 = "";
    //                string sg3 = "";

    //                System.Data.DataTable dh = null;
    //                dh = GetTemp5_1tM(y, "Q", "");
    //                sg = dh.Rows[0]["�ƶq"].ToString();
    //                drtemp5[y + "_Q"] = sg;

    //                System.Data.DataTable dh2 = null;
    //                dh2 = GetTemp5_1tM(y, "R", "");
    //                sg2 = dh2.Rows[0]["���B"].ToString();
    //                drtemp5[y + "_R"] = sg2;

    //                System.Data.DataTable dh3 = null;
    //                dh3 = GetTemp5_1tM(y, "C", "");
    //                sg3 = dh3.Rows[0]["����"].ToString();
    //                drtemp5[y + "_C"] = sg3;

    //                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                int �P�����B = Convert.ToInt32(sg2);
    //                double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                string aa = Convert.ToString(�Q���);
    //                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                {
    //                    drtemp5[y + "_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                }

    //            }

    //            drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //+ Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //+ Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //            DJ33 = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //+ Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //+ Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //            drtemp5["T_R"] = DJ33;
    //            drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //+ Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //+ Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //            drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //+ Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //+ Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);




    //            double �Q���5 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //            string aa5 = Convert.ToString(�Q���5);
    //            if (aa5 == "���L�a�j" || aa5 == "�t�L�a�j" || aa5 == "���O�@�ӼƦr")
    //            {
    //                drtemp5["T_P"] = 0;
    //            }
    //            else
    //            {
    //                drtemp5["T_P"] = �Q���5.ToString("#0.00") + "%";
    //            }


    //            double SALES33 = (DJ33 / HJ) * 100;
    //            drtemp5["SALES2"] = SALES33.ToString("#0.00") + "%";


    //            dtCostDD.Rows.Add(drtemp5);


    //            //�[�`
    //            drtemp5 = dtCostDD.NewRow();
    //            drtemp5["�~��"] = "";
    //            drtemp5["�Ȥ�s��"] = "";
    //            drtemp5["�Ȥ�W��"] = "Grand Total";

    //            for (int y = 1; y <= 12; y++)
    //            {
    //                string sg = "";
    //                string sg2 = "";
    //                string sg3 = "";

    //                System.Data.DataTable dh = null;
    //                dh = GetTemp5_1tM(y, "Q");
    //                sg = dh.Rows[0]["�ƶq"].ToString();
    //                drtemp5[y + "_Q"] = sg;

    //                System.Data.DataTable dh2 = null;
    //                dh2 = GetTemp5_1tM(y, "R");
    //                sg2 = dh2.Rows[0]["���B"].ToString();
    //                drtemp5[y + "_R"] = sg2;

    //                System.Data.DataTable dh3 = null;
    //                dh3 = GetTemp5_1tM(y, "C");
    //                sg3 = dh3.Rows[0]["����"].ToString();
    //                drtemp5[y + "_C"] = sg3;

    //                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                int �P�����B = Convert.ToInt32(sg2);
    //                double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                string aa = Convert.ToString(�Q���);
    //                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                {
    //                    drtemp5[y + "_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                }

    //            }

    //            drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //          + Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //          + Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //            drtemp5["T_R"] = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //+ Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //+ Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //            drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //+ Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //+ Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //            drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //+ Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //+ Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);

    //            double �Q���55 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //            string aa55 = Convert.ToString(�Q���55);
    //            if (aa55 == "���L�a�j" || aa55 == "�t�L�a�j" || aa55 == "���O�@�ӼƦr")
    //            {
    //                drtemp5["T_P"] = 0;
    //            }
    //            else
    //            {
    //                drtemp5["T_P"] = �Q���55.ToString("#0.00") + "%";
    //            }

    //            dtCostDD.Rows.Add(drtemp5);


                //drtemp5 = dtCostDD.NewRow();
                //dtCostDD.Rows.Add(drtemp5);

                //drtemp5 = dtCostDD.NewRow();
                //drtemp5["�Ȥ�W��"] = "�~��";
                //dtCostDD.Rows.Add(drtemp5);




    //            //SALES
    //            System.Data.DataTable dtemp7 = GetTemp5Sales("");




    //            for (int i = 0; i <= dtemp7.Rows.Count - 1; i++)
    //            {
    //                drtemp5 = dtCostDD.NewRow();

    //                �~��t = dtemp7.Rows[i]["�~��"].ToString();
    //                drtemp5["�~��"] = "";
    //                drtemp5["�Ȥ�s��"] = "";
    //                drtemp5["�Ȥ�W��"] = �~��t;
    //                drtemp5["SALES2"] = "";
    //                for (int y = 1; y <= 12; y++)
    //                {
    //                    string sg = "";
    //                    string sg2 = "";
    //                    string sg3 = "";


    //                    System.Data.DataTable dh = null;
    //                    dh = GetTemp5_1SALESM(y, �~��t, "Q");
    //                    sg = dh.Rows[0]["�ƶq"].ToString();
    //                    drtemp5[y + "_Q"] = sg;

    //                    System.Data.DataTable dh2 = null;
    //                    dh2 = GetTemp5_1SALESM(y, �~��t, "R");
    //                    sg2 = dh2.Rows[0]["���B"].ToString();
    //                    drtemp5[y + "_R"] = sg2;

    //                    System.Data.DataTable dh3 = null;
    //                    dh3 = GetTemp5_1SALESM(y, �~��t, "C");
    //                    sg3 = dh3.Rows[0]["����"].ToString();
    //                    drtemp5[y + "_C"] = sg3;

    //                    double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                    int �P�����B = Convert.ToInt32(sg2);
    //                    double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                    drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                    string aa = Convert.ToString(�Q���);
    //                    if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                    {
    //                        drtemp5[y + "_P"] = 0;
    //                    }
    //                    else
    //                    {
    //                        drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                    }

    //                }

    //                drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //                        + Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //                        + Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //                DJ2 = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //    + Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //    + Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //                drtemp5["T_R"] = DJ2;
    //                drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //    + Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //    + Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //                drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //    + Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //    + Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);

    //                double �Q���1 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //                string aa2 = Convert.ToString(�Q���1);
    //                if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr")
    //                {
    //                    drtemp5["T_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5["T_P"] = �Q���1.ToString("#0.00") + "%";
    //                }

    //                double SALES = (DJ2 / HJ) * 100;
    //                drtemp5["SALES2"] = SALES.ToString("#0.00") + "%";
    //                dtCostDD.Rows.Add(drtemp5);
    //            }

    //            drtemp5 = dtCostDD.NewRow();
    //            drtemp5["�~��"] = "";
    //            drtemp5["�Ȥ�s��"] = "";
    //            drtemp5["�Ȥ�W��"] = "�[�`";

    //            drtemp5["SALES2"] = "";

    //            for (int y = 1; y <= 12; y++)
    //            {
    //                string sg = "";
    //                string sg2 = "";
    //                string sg3 = "";


    //                System.Data.DataTable dh = null;
    //                dh = GetTemp5_1tM(y, "Q");
    //                sg = dh.Rows[0]["�ƶq"].ToString();
    //                drtemp5[y + "_Q"] = sg;

    //                System.Data.DataTable dh2 = null;
    //                dh2 = GetTemp5_1tM(y, "R");
    //                sg2 = dh2.Rows[0]["���B"].ToString();
    //                drtemp5[y + "_R"] = sg2;

    //                System.Data.DataTable dh3 = null;
    //                dh3 = GetTemp5_1tM(y, "C");
    //                sg3 = dh3.Rows[0]["����"].ToString();
    //                drtemp5[y + "_C"] = sg3;

    //                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                int �P�����B = Convert.ToInt32(sg2);
    //                double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                string aa = Convert.ToString(�Q���);
    //                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                {
    //                    drtemp5[y + "_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                }

    //            }

    //            drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    //                    + Convert.ToInt32(drtemp5["4_Q"]) + Convert.ToInt32(drtemp5["5_Q"]) + Convert.ToInt32(drtemp5["6_Q"]) + Convert.ToInt32(drtemp5["7_Q"]) + Convert.ToInt32(drtemp5["8_Q"])
    //                    + Convert.ToInt32(drtemp5["9_Q"]) + Convert.ToInt32(drtemp5["10_Q"]) + Convert.ToInt32(drtemp5["11_Q"]) + Convert.ToInt32(drtemp5["12_Q"]);
    //            DJ = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    //+ Convert.ToInt32(drtemp5["4_R"]) + Convert.ToInt32(drtemp5["5_R"]) + Convert.ToInt32(drtemp5["6_R"]) + Convert.ToInt32(drtemp5["7_R"]) + Convert.ToInt32(drtemp5["8_R"])
    //+ Convert.ToInt32(drtemp5["9_R"]) + Convert.ToInt32(drtemp5["10_R"]) + Convert.ToInt32(drtemp5["11_R"]) + Convert.ToInt32(drtemp5["12_R"]);
    //            drtemp5["T_R"] = DJ;
    //            drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    //+ Convert.ToInt32(drtemp5["4_C"]) + Convert.ToInt32(drtemp5["5_C"]) + Convert.ToInt32(drtemp5["6_C"]) + Convert.ToInt32(drtemp5["7_C"]) + Convert.ToInt32(drtemp5["8_C"])
    //+ Convert.ToInt32(drtemp5["9_C"]) + Convert.ToInt32(drtemp5["10_C"]) + Convert.ToInt32(drtemp5["11_C"]) + Convert.ToInt32(drtemp5["12_C"]);
    //            drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    //+ Convert.ToInt32(drtemp5["4_G"]) + Convert.ToInt32(drtemp5["5_G"]) + Convert.ToInt32(drtemp5["6_G"]) + Convert.ToInt32(drtemp5["7_G"]) + Convert.ToInt32(drtemp5["8_G"])
    //+ Convert.ToInt32(drtemp5["9_G"]) + Convert.ToInt32(drtemp5["10_G"]) + Convert.ToInt32(drtemp5["11_G"]) + Convert.ToInt32(drtemp5["12_G"]);

    //            double �Q���6 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
    //            string aa6 = Convert.ToString(�Q���6);
    //            if (aa6 == "���L�a�j" || aa6 == "�t�L�a�j" || aa6 == "���O�@�ӼƦr")
    //            {
    //                drtemp5["T_P"] = 0;
    //            }
    //            else
    //            {
    //                drtemp5["T_P"] = �Q���6.ToString("#0.00") + "%";
    //            }

    //            dtCostDD.Rows.Add(drtemp5);


                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\���J����by��.xls";


                //Excel���˪���
                string ExcelTemplate = FileName;

                //��X��
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                string FF = comboBox2.SelectedValue.ToString();
                //���� Excel Report
                ExcelReport.ACC2(dtCostDD, ExcelTemplate, OutPutFile, FF);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            A1 = comboBox3.Text + "0101";
            A2 = comboBox3.Text + "1231";

            Category(2, comboBox3.SelectedValue.ToString(), "Account_Temp6");

            dataGridView8.DataSource = GetCust();
            ExcelReport.GridViewToExcel(dataGridView8);

        }
        System.Data.DataTable GetCust()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.CARDCODE �Ȥ�s��,MAX(T0.CARDNAME) �Ȥ�W��,SUM(T0.GQTY) �ƶq,SUM(T0.GTOTAL) ���B");
            sb.Append(" ,max(T4.pymntgroup) �I�ڱ���,MAX(T1.STREET) �a�},MAX(T2.PHONE1) �q�� FROM Account_Temp6  T0");
            sb.Append(" LEFT JOIN (select CARDCODE,MAX(STREET) STREET from ACMESQL02.DBO.crd1 where adrestype='S' ");
            sb.Append(" AND ISNULL(STREET,'') <> ''");
            sb.Append(" GROUP BY CARDCODE) T1 ON (T0.CARDCODE=T1.CARDCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OCRD T2 ON (T0.CARDCODE=T2.CARDCODE COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" left join ACMESQL02.DBO.octg t4 on(t2.groupnum=t4.groupnum)");
            sb.Append(" GROUP BY T0.CARDCODE");
            sb.Append(" ORDER BY SUM(T0.GTOTAL) DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



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

        System.Data.DataTable GetSup()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT A.CARDCODE �t�ӽs��,MAX(T0.CARDNAME) �t�ӦW��,CAST(ISNULL(SUM(�ƶq),0) AS INT) �ƶq,CAST(ISNULL(SUM(���B),0) AS FLOAT) ���B");
            sb.Append(" ,max(T4.pymntgroup) �I�ڱ���,MAX(T3.STREET) �a�},MAX(T0.PHONE1) �q�� FROM (SELECT t0.cardcode,SUM(t1.QUANTITY) �ƶq,SUM(T1.LINETOTAL) ���B FROM OPCH T0");
            sb.Append(" LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  WHERE SUBSTRING(T0.CARDCODE,1,1)='S' and year(t0.docdate)=@aa  ");
            sb.Append(" GROUP BY t0.cardcode");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT t0.cardcode,SUM(t1.QUANTITY)*-1 �ƶq,SUM(T1.LINETOTAL)*-1 ���B FROM ORPC T0");
            sb.Append(" LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) WHERE SUBSTRING(T0.CARDCODE,1,1)='S' and year(t0.docdate)=@aa ");
            sb.Append(" GROUP BY t0.cardcode ) AS A");
            sb.Append(" LEFT JOIN OCRD T0 ON (A.CARDCODE=T0.CARDCODE)");
            sb.Append(" LEFT JOIN (select CARDCODE,MAX(STREET) STREET from crd1 where adrestype='S' ");
            sb.Append(" AND ISNULL(STREET,'') <> ''");
            sb.Append(" GROUP BY CARDCODE) T3 ON (A.CARDCODE=T3.CARDCODE)");
            sb.Append(" left join ACMESQL02.DBO.octg t4 on(T0.groupnum=t4.groupnum)");
            sb.Append(" GROUP BY A.cardcode");
            sb.Append(" ORDER BY SUM(���B) DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", comboBox4.SelectedValue.ToString()));

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
        System.Data.DataTable GetAccount()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select account �|�p���,sum(gqty) �ƶq,sum(gtotal) ���B from Account_Temp6 group by account ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;


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
        System.Data.DataTable GetItem(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select itemcode ���� from oitm where itemcode in ( " + cs + ")");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

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
        System.Data.DataTable GetItemm(string cs)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select ITEMCODE ���~�Ƹ�,Convert(varchar(8),T0.DOCDATE,112) �L�b���,");
            sb.Append(" CASE TRANSTYPE ");
            sb.Append(" WHEN 13 THEN 'AR' WHEN 14 THEN 'AR�U��' WHEN 15 THEN '��f' WHEN 16 THEN '�P��h�f' ");
            sb.Append(" WHEN 18 THEN 'AP' WHEN 19 THEN 'AP�U��' WHEN 20 THEN '���f���ʳ�' WHEN 20 THEN '���ʰh�f'");
            sb.Append(" WHEN 59 THEN '���f��' WHEN 60 THEN '�o�f��' WHEN 67 THEN '�w�s�ռ�' ");
            sb.Append(" ELSE '' END TRANSTYPE,BASE_REF,WAREHOUSE,CAST(INQTY AS INT) ���f�q,CAST(OUTQTY AS INT) �o�f�q");
            sb.Append(" ,CAST((SELECT SUM(INQTY-OUTQTY) A FROM OINM T1 where T1.itemcode =T0.ITEMCODE AND T1.TRANSNUM <= T0.TRANSNUM) AS INT) �l�B from oinm T0  where itemcode in ( " + cs + ") ");
            sb.Append(" and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 "); 
            sb.Append(" ORDER BY ITEMCODE,DOCDATE,DOCTIME");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox11.Text));
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
        System.Data.DataTable GetItemmDD(string COM, string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CARDCODE �Ȥ�s��,SUBSTRING(GROUPNAME,4,10) �Ȥ�s��,CARDNAME �Ȥ�W��,SUM(GTOTAL) �P�����B,SUM(GVALUE) �P������");
            sb.Append(" ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) �Q��");
            sb.Append(" ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)  �Q���");
            sb.Append(" ,COM=('" + COM + "') from " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)");
            sb.Append(" WHERE  Convert(varchar(8),DDATE,112) BETWEEN @DocDate1 AND @DocDate2");
            sb.Append(" GROUP BY CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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

        System.Data.DataTable GetItemmDDACC(string COM, string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select CARDCODE �Ȥ�s��,SUBSTRING(GROUPNAME,4,10) �Ȥ�s��,CARDNAME �Ȥ�W��,SUM(GTOTAL) �P�����B,SUM(GVALUE) �P������");
            sb.Append(" ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) �Q��,ACCOUNT+'-'+ACCTNAME COLLATE  Chinese_Taiwan_Stroke_CI_AS ��إN��");
            sb.Append(" ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)  �Q���");
            sb.Append(" ,COM=('" + COM + "') from " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)");
            sb.Append(" LEFT JOIN AcmeSql02.DBO.OACT T2 ON (T0.ACCOUNT=T2.ACCTCODE  COLLATE  Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" WHERE  Convert(varchar(8),DDATE,112) BETWEEN @DocDate1 AND @DocDate2");
            sb.Append(" GROUP BY CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME,ACCOUNT,ACCTNAME");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        System.Data.DataTable GetItemmDDF(string COM, string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select T0.CARDCODE �Ȥ�s��,SUBSTRING(GROUPNAME,4,10) �Ȥ�s��,T0.CARDNAME �Ȥ�W��,SUM(GTOTAL) �P�����B,SUM(GVALUE) �P������");
            sb.Append(" ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) �Q��");
            sb.Append(" ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)  �Q���");
            sb.Append(" ,T2.u_acme_shipto1 SHIPTO,COM=('" + COM + "') from " + Account_Temp6 + " T0");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.OINV T2 ON (T0.DOCENTRY=T2.DOCENTRY) ");
            sb.Append(" WHERE  Convert(varchar(8),DDATE,112) BETWEEN @DocDate1 AND @DocDate2");
            sb.Append(" GROUP BY T0.CARDCODE,SUBSTRING(GROUPNAME,4,10),T0.CARDNAME,T2.u_acme_shipto1");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        System.Data.DataTable GetItemmSALES(string COM, string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               select T2.SEQ,CARDCODE �Ȥ�s��, ltrim(substring(REPLACE(REPLACE(T0.SALES,'.',''),'''',''),0,CHARINDEX('(', REPLACE(REPLACE(T0.SALES,'.',''),'''',''))))  �Ȥ�s��,CARDNAME �Ȥ�W��,SUM(GTOTAL) �P�����B,SUM(GVALUE) �P������ ");
            sb.Append("               ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) �Q�� ");
            sb.Append("               ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)  �Q��� ,T2.BU");
            sb.Append("                ,COM=('" + COM + "') from " + Account_Temp6 + " T0 ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE) ");
            sb.Append("               LEFT JOIN Account_TempSALES T2 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T2.SALES)");
            sb.Append("               WHERE  Convert(varchar(4),DDATE,112) =@YEAR ");
            if (comboBox9.Text != "")
            {
                sb.Append("    AND MONTH(DDATE)<=@MONTH      ");
            }
            sb.Append("               AND SUBSTRING(GROUPNAME,4,10)='TFT' ");
            sb.Append("               GROUP BY CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME,REPLACE(REPLACE(T0.SALES,'.',''),'''',''),T2.BU, T2.SEQ");
            sb.Append("               ORDER BY T2.SEQ,REPLACE(REPLACE(T0.SALES,'.',''),'''','')");
            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox8.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", comboBox9.Text));
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
        System.Data.DataTable GETNANCY(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("                            select CAST(T2.SEQ AS varchar)  SEQ,CASE T2.BU  WHEN '' THEN 'TFT-OTHERS' ELSE T2.BU END BU  ,SUM(GQTY) �ƶq, ");
            sb.Append("							SUM(GTOTAL) �`��P���B,SUM(GVALUE) �`��P����  ");
            sb.Append("                          ,SUM(GTOTAL-GVALUE) �P��Q��  ");
            sb.Append("                            ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)+'%'    �Q��� ");
            sb.Append("                     from " + Account_Temp6 + " T0  ");
            sb.Append("                            LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)  ");
            sb.Append("                            LEFT JOIN Account_TempSALES T2 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T2.SALES) ");
            sb.Append("                            WHERE  Convert(varchar(4),DDATE,112) =@YEAR");
            sb.Append("							                      AND SUBSTRING(GROUPNAME,4,10)='TFT'  AND ISNULL(T2.SEQ,'') <> ''");
            sb.Append("                            GROUP BY SUBSTRING(GROUPNAME,4,10),T2.BU, CAST(T2.SEQ AS varchar)  ");
            sb.Append("							UNION ALL");
            sb.Append("							                            select 'TOTAL' SEQ,'TOTAL',SUM(GQTY) �ƶq, ");
            sb.Append("							SUM(GTOTAL) �`��P���B,SUM(GVALUE) �`��P����  ");
            sb.Append("                          ,SUM(GTOTAL-GVALUE) �P��Q��  ");
            sb.Append("                            ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)+'%'    �Q��� ");
            sb.Append("                     from " + Account_Temp6 + " T0  ");
            sb.Append("                            LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)  ");
            sb.Append("                            LEFT JOIN Account_TempSALES T2 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T2.SALES) ");
            sb.Append("                            WHERE  Convert(varchar(4),DDATE,112) =@YEAR");
            sb.Append("							                      AND SUBSTRING(GROUPNAME,4,10)='TFT'  AND ISNULL(T2.SEQ,'') <> ''");
            sb.Append("                            GROUP BY SUBSTRING(GROUPNAME,4,10)");
            sb.Append("                            ORDER BY SEQ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox8.Text));
         
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
        System.Data.DataTable GETNANCY2(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("                            select SUBSTRING(GROUPNAME,4,10) [GROUP] ,SUM(GQTY) �ƶq, ");
            sb.Append("							SUM(GTOTAL) �`��P���B,SUM(GVALUE) �`��P����  ");
            sb.Append("                          ,SUM(GTOTAL-GVALUE) �P��Q��  ");
            sb.Append("                            ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR) +'%'   �Q��� ");
            sb.Append("                     from " + Account_Temp6 + " T0  ");
            sb.Append("                            LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)  ");
            sb.Append("                            WHERE  Convert(varchar(4),DDATE,112) =@YEAR");
            sb.Append("							                      AND SUBSTRING(GROUPNAME,4,10) = ('ESCO')  ");
            sb.Append("                            GROUP BY SUBSTRING(GROUPNAME,4,10)");
            sb.Append("							UNION ALL");
            sb.Append("							                            select SUBSTRING(GROUPNAME,4,10) [GROUP] ,SUM(GQTY) �ƶq, ");
            sb.Append("							SUM(GTOTAL) �`��P���B,SUM(GVALUE) �`��P����  ");
            sb.Append("                          ,SUM(GTOTAL-GVALUE) �P��Q��  ");
            sb.Append("                            ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR) +'%'   �Q��� ");
            sb.Append("                     from " + Account_Temp6 + " T0  ");
            sb.Append("                            LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)  ");
            sb.Append("                            WHERE  Convert(varchar(4),DDATE,112) =@YEAR");
            sb.Append("							                      AND SUBSTRING(GROUPNAME,4,10) = ('TFT')  ");
            sb.Append("                            GROUP BY SUBSTRING(GROUPNAME,4,10)");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox8.Text));

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
        System.Data.DataTable GetItemmSALESDATE(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("               SELECT Convert(varchar(8),MAX(dateadd(day ,-1, dateadd(m, datediff(m,0,DDATE)+1,0))),112) DDATE   from " + Account_Temp6 + "  ");
            if (comboBox9.Text != "")
            {
                sb.Append("    WHERE MONTH(DDATE)<=@MONTH      ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MONTH", comboBox9.Text));
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
        System.Data.DataTable GetSALESOITM(string COM,string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("              select T2.SEQ,CARDCODE �Ȥ�s��, ltrim(substring(REPLACE(REPLACE(T0.SALES,'.',''),'''',''),0,CHARINDEX('(', REPLACE(REPLACE(T0.SALES,'.',''),'''',''))))  �Ȥ�s��,CARDNAME �Ȥ�W��,");
            sb.Append("                                   T0.ITEMCODE ���~�s��,T0.ITEMNAME ���~�W��,SUM(GTOTAL) �P�����B,SUM(GVALUE) �P������  ");
            sb.Append("                             ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) �Q��  ");
            sb.Append("                             ,CAST(CAST(ROUND(CASE WHEN SUM(GTOTAL-GVALUE) = 0 THEN 0 WHEN  SUM(GTOTAL) =0 THEN 0 ELSE SUM(GTOTAL-GVALUE)/SUM(GTOTAL)* 100  END,2) AS DECIMAL(10,2)) AS VARCHAR)  �Q��� ,T2.BU ");
            sb.Append("       ,COM=('" + COM + "')      from " + Account_Temp6 + " T0 ");
            sb.Append("                             LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE)  ");
            sb.Append("                             LEFT JOIN Account_TempSALES T2 ON (REPLACE(REPLACE(T0.SALES,'.',''),'''','')=T2.SALES) ");
            sb.Append("               WHERE  Convert(varchar(4),DDATE,112) =@YEAR ");
            if (comboBox9.Text != "")
            {
                sb.Append("    AND MONTH(DDATE)<=@MONTH      ");
            }
            sb.Append("                             AND SUBSTRING(GROUPNAME,4,10)='TFT'  ");
            sb.Append("                             GROUP BY CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME,REPLACE(REPLACE(T0.SALES,'.',''),'''',''),T2.BU, T2.SEQ,T0.ITEMCODE,T0.ITEMNAME");
            sb.Append("               ORDER BY T2.SEQ,REPLACE(REPLACE(T0.SALES,'.',''),'''',''),CARDCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@YEAR", comboBox8.Text));
            command.Parameters.Add(new SqlParameter("@MONTH", comboBox9.Text));
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
        System.Data.DataTable GetItemmDDOCRD(string CARDCODE, string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT COUNTRY  from " + Account_Temp6 + "  WHERE CARDCODE=@CARDCODE AND ISNULL(COUNTRY,'') <> '' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
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
        System.Data.DataTable GetItemmDDS(string Account_Temp6)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();


            sb.Append("               select T0.CARDCODE �Ȥ�s��,SUBSTRING(GROUPNAME,4,10) �Ȥ�s��,CARDNAME �Ȥ�W��,SUM(GTOTAL) ���B,SUM(GVALUE) ���� ");
            sb.Append("               ,SUM(GQTY) �ƶq,SUM(GTOTAL-GVALUE) ��Q,MAX(TERM) �I�ڤ覡,CASE WHEN ISNULL(MAX(T3.CARDCODE),'')='' THEN 'V' END  �s�Ȥ�");
            sb.Append("  from   " + Account_Temp6 + "   T0 ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE) ");
            sb.Append("          LEFT JOIN (SELECT MAX(U_ACME_PAY) TERM,DOCENTRY FROM ACMESQL02.DBO.OINV WHERE ISNULL(U_ACME_PAY,'') <> '' GROUP BY DOCENTRY)");
            sb.Append(" T2 ON (T0.DOCENTRY=T2.DOCENTRY)");
            sb.Append(" LEFT JOIN (SELECT DISTINCT CARDCODE  FROM ACMESQL02.DBO.OINV ");
            sb.Append("   WHERE  Convert(varchar(8),DOCDATE,112) NOT BETWEEN @DocDate1 AND @DocDate2 ) T3");
            sb.Append(" ON(REPLACE(T0.CARDCODE,'''','')=T3.CARDCODE  COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append("               WHERE  Convert(varchar(8),DDATE,112) BETWEEN @DocDate1 AND @DocDate2 ");
            sb.Append("               GROUP BY T0.CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandTimeout = 0;
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        System.Data.DataTable GetItemmDDQ()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT �Ȥ�s��,�Ȥ�s��,�Ȥ�W��,�ƶq,���B,����,�Q��,Q1,Q2,Q3,Q4 FROM (Select �Ȥ�s�� ,�Ȥ�s��,�Ȥ�W��,[1] Q1,[2] Q2,[3] Q3,[4]  Q4");
            sb.Append(" from (");
            sb.Append("    select CARDCODE �Ȥ�s��,SUBSTRING(GROUPNAME,4,10) �Ȥ�s��,CARDNAME �Ȥ�W��,");
            sb.Append("    SUM(GTOTAL-GVALUE) �Q��,");
            sb.Append("              DATEPART(QQ,DDATE) Q");
            sb.Append("               from Account_Temp6 T0 ");
            sb.Append("               LEFT JOIN ACMESQL02.DBO.OCRG T1 ON (T0.CARDGROUP=T1.GROUPCODE) ");
            sb.Append("               GROUP BY CARDCODE,SUBSTRING(GROUPNAME,4,10),CARDNAME,DATEPART(QQ,DDATE) ) T");
            sb.Append(" PIVOT");
            sb.Append(" (");
            sb.Append(" SUM(�Q��)");
            sb.Append(" FOR Q IN");
            sb.Append(" ( [1],[2],[3],[4])");
            sb.Append(" ) AS pvt ) AS T0");
            sb.Append(" LEFT JOIN (SELECT CARDCODE,SUM(GQTY) �ƶq,SUM(GTOTAL) ���B,SUM(GVALUE) ����,SUM(GTOTAL-GVALUE)  �Q�� FROM  Account_Temp6 GROUP BY CARDCODE ) T1 ");
            sb.Append(" ON (T0.�Ȥ�s�� = T1.CARDCODE)  ORDER BY �Ȥ�s��,�Q�� DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
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
        System.Data.DataTable GetDIST()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT SUBSTRING(T2.GROUPNAME,4,6) �s��,''''+A.CARDCODE �Ȥ�s��,A.CardName �Ȥ�W��,SUM(�ƶq) �ƶq,SUM(�P����B) �P����B,SUM(����) �P�⦨��,SUM(�P����B)-SUM(����) �P���Q FROM (         ");
            sb.Append("    SELECT T0.DOCENTRY, T0.CARDCODE,T0.CardName ,");
            sb.Append("               SUM(T2.Quantity) �ƶq,MAX(T0.DOCTOTAL-T0.vatsumsy) �P����B");
            sb.Append("               ,SUM(CAST(Round(T2.StockPrice*T2.Quantity,0) AS INT)) ���� FROM OINV T0 ");
            sb.Append("                     INNER JOIN INV1 T2 ON T0.DocEntry = T2.DocEntry  INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T2.ItemCode ");
            sb.Append("                     WHERE T0.[DocType] ='I' ");
            sb.Append("                     and  ISNULL(TA.U_GROUP,'') <> 'Z&R-�O�����s��'  AND T0.U_IN_BSTYC <> '1' ");
            sb.Append(" AND Convert(varchar(8),T0.[DocDate],112) BETWEEN @DocDate1 AND @DocDate2");
            sb.Append(" GROUP BY  T0.CARDCODE,T0.CardName,T0.DOCENTRY) AS A");
            sb.Append(" LEFT JOIN OCRD T1 ON (A.CARDCODE=T1.CARDCODE)");
            sb.Append(" LEFT JOIN OCRG T2 ON (T1.GROUPCODE=T2.GROUPCODE)");
            sb.Append(" GROUP BY A.CARDCODE,A.CardName,T2.GROUPNAME");
            sb.Append(" ORDER BY T2.GROUPNAME,SUM(�P����B)-SUM(����) DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text));
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
        System.Data.DataTable GetACC()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.[Account] ���,T3.ACCTNAME ��ئW��, SUM(T0.[debit]) Debit, SUM(T0.[credit]) Credit,SUM(T0.[debit])-SUM(T0.[credit]) Balance ,t0.profitcode ����");
            sb.Append("              FROM  [dbo].[JDT1] T0  INNER  JOIN [dbo].[OJDT] T1  ON  T1.[TransId] = T0.TransId  ");
            sb.Append("            Left join OACT t3 on (t0.[Account]=t3.ACCTCODE)    ");
            sb.Append("              WHERE   T0.[TransType] <> '-3'  AND  T0.[TransType] <> '-2'   and substring(T0.[Account],1,1) in ('4','5')");
            sb.Append(" and Convert(varchar(8),T0.REFDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T0.[Account],t0.profitcode,T3.ACCTNAME");
            sb.Append(" order by T0.[Account]");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox12.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox13.Text));
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

        System.Data.DataTable GetBAALANCE(string DOCENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCENTRY FROM OINV WHERE DOCENTRY=@DOCENTRY");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
     

        System.Data.DataTable GetAvg(string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '' ���� ,0  ����,0  �ƶq FROM OITM T0 WHERE T0.ITEMCODE=@aa");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT T0.ITEMCODE ����,AVGPRICE  ����,ONHAND  �ƶq FROM OITM T0 WHERE T0.ITEMCODE =@aa ");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '���ʳ�',AVG(PRICE*DOCRATE) ,SUM(QUANTITY)  FROM POR1 T0 LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE ITEMCODE ='[%0]' AND LINESTATUS='O'");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '�[�v',(ONHAND*AVGPRICE+T1.QUANTITY*T1.PRICE)/(ONHAND+T1.QUANTITY)  ����,ONHAND+T1.QUANTITY  �ƶq FROM OITM T0");
            sb.Append(" LEFT JOIN (SELECT SUM(QUANTITY) QUANTITY,AVG(PRICE*DOCRATE) PRICE,MAX(ITEMCODE) ITEMCODE FROM POR1 T0 ");
            sb.Append(" LEFT JOIN OPOR T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE ITEMCODE ='[%0]' AND LINESTATUS='O') T1 ON(T0.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" WHERE T0.ITEMCODE =@aa");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));

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
        System.Data.DataTable GetOINV(string ACCTCODE, string OCRCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("             SELECT SUM(�ƶq) �ƶq,���,���� FROM (          ");
            sb.Append("             SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM OINV T0 ");
            sb.Append("      LEFT JOIN INV1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                        UNION ALL");
            sb.Append("                           SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM ORIN T0 ");
            sb.Append("      LEFT JOIN RIN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                      UNION ALL");
            sb.Append("             SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM ODLN T0 ");
            sb.Append("      LEFT JOIN DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                      UNION ALL");
            sb.Append("                       SELECT SUM(T1.QUANTITY)*-1 �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM ORDN T0 ");
            sb.Append("      LEFT JOIN RDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                    UNION ALL");
            sb.Append("             SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM OPCH T0 ");
            sb.Append("      LEFT JOIN PCH1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                    UNION ALL");
            sb.Append("               SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM ORPC T0 ");
            sb.Append("      LEFT JOIN RPC1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                    UNION ALL");
            sb.Append("             SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM OIGN T0 ");
            sb.Append("      LEFT JOIN IGN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                    UNION ALL");
            sb.Append("             SELECT SUM(T1.QUANTITY) �ƶq,T2.ACCOUNT ���,ISNULL(T2.PROFITCODE,'') ���� FROM OIGE T0 ");
            sb.Append("      LEFT JOIN IGE1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)  ");
            sb.Append("       INNER JOIN JDT1 T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append("              WHERE T2.ACCOUNT = @ACCTCODE and ISNULL(T2.PROFITCODE,'')=@OCRCODE ");
            sb.Append("              and Convert(varchar(8),T0.DOCDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("              GROUP BY T2.ACCOUNT,T2.PROFITCODE");
            sb.Append("                    UNION ALL");
            sb.Append("                 SELECT * FROM ( SELECT SUM(T0.CMPLTQTY) �ƶq,T1.ACCOUNT ���,'11131' ���� FROM OWOR T0 ");
            sb.Append("                     INNER JOIN JDT1 T1 ON (T0.TRANSID=T1.TRANSID)");
            sb.Append("                          WHERE T1.ACCOUNT = @ACCTCODE  ");
            sb.Append("                          and Convert(varchar(8),T1.REFDATE,112) between @DocDate1 and @DocDate2 ");
            sb.Append("                          GROUP BY T1.ACCOUNT ) AS B ");
            sb.Append(" WHERE ����=@OCRCODE ");
            sb.Append(" ) AS AA");
            sb.Append("             GROUP BY ���,����");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@ACCTCODE", ACCTCODE));
            command.Parameters.Add(new SqlParameter("@OCRCODE", OCRCODE));
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox12.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox13.Text));
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
        private void button9_Click(object sender, EventArgs e)
        {
            A1 = comboBox4.Text + "0101";
            A2 = comboBox4.Text + "1231";

            Category(2, comboBox4.SelectedValue.ToString(), "Account_Temp6");

            dataGridView8.DataSource = GetSup();
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            A1 = comboBox5.Text + "0101";
            A2 = comboBox5.Text + "1231";

            Category(2, comboBox5.SelectedValue.ToString(), "Account_Temp6");

            dataGridView8.DataSource = GetAccount();
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt = GetItem(textBox4.Text);

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("�L���Ƹ�");
                    return;
                }
                System.Data.DataTable dtCost = MakeTableAvg();
                System.Data.DataTable dtDoc = null;
                DataRow dr = null;
                string ����;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {

                    ���� = Convert.ToString(dt.Rows[i]["����"]);




                    dtDoc = GetAvg(����);
                    for (int j = 0; j <= dtDoc.Rows.Count - 1; j++)
                    {
                        dr = dtCost.NewRow();

                        if (j == 0)
                        {
                            dr["����"] = Convert.ToString(dtDoc.Rows[j]["����"]);
                            dr["����"] = "����";
                            dr["�ƶq"] = "�ƶq";
                        }
                        else
                        {

                            dr["����"] = Convert.ToString(dtDoc.Rows[j]["����"]);
                            dr["����"] = Convert.ToString(dtDoc.Rows[j]["����"]);
                            dr["�ƶq"] = Convert.ToString(dtDoc.Rows[j]["�ƶq"]);
                        }

                        dtCost.Rows.Add(dr);


                    }

                    for (int H = 0; H <= 1; H++)
                    {
                        dr = dtCost.NewRow();
                        dr["����"] = "";
                        dr["����"] = "";
                        dr["�ƶq"] = "";

                        dtCost.Rows.Add(dr);
                    }
                }

                dataGridView1.DataSource = dtCost;
                ExcelReport.GridViewToExcel(dataGridView1);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            APS2 frm1 = new APS2();
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = frm1.q;
            }
        }




        private void button14_Click(object sender, EventArgs e)
        {

            UPDATE2();

            MessageBox.Show("��s���\");

        }


        System.Data.DataTable TACO(string TABLE)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT PARAM_NO FROM RMA_PARAMS WHERE PARAM_KIND=@TABLE ");



            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TABLE", TABLE));
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

        System.Data.DataTable TF(string TRGETENTRY)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT RANK() OVER (ORDER BY DOCENTRY DESC) AS �Ǹ�,DOCENTRY AR,TRGETENTRY ��f FROM  INV1 WHERE  TRGETENTRY IN (");
            sb.Append(" select docentry  from dln1 where BASEtype='13'");
            sb.Append(" GROUP BY DOCENTRY HAVING COUNT (DISTINCT BASEENTRY) >1) AND    DOCENTRY NOT IN (SELECT DOCENTRY FROM OINV WHERE DOCTOTAL=0) AND TRGETENTRY=@TRGETENTRY  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRGETENTRY", TRGETENTRY));
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
        private void button15_Click(object sender, EventArgs e)
        {
            System.Data.DataTable DFF = null;
            string gh = "";
            string gh2 = "";
            string g2 = "";
            string g3 = "";
            string year2 = textBox6.Text.Substring(0, 4);
            int hg = Convert.ToInt16(year2) - 1;
            if (comboBox6.Text == "��")
            {

                string ff2 = year2 + "/" + textBox6.Text.Substring(4, 2) + "/01";
                DateTime gf = Convert.ToDateTime(ff2);
                g2 = gf.AddMonths(-1).ToString("yyyyMM");
                g3 = textBox6.Text;
                gh = "�i���͹�~(��)���q-" + textBox6.Text.Substring(4, 2) + "���禬�P�W����";
                gh2 = "�P�W�����W(��)";
                Category(3, "", "Account_Temp6");
                DFF = GetAA();
            }
            else if (comboBox6.Text == "�u")
            {
                string q = util.quarter(textBox6.Text);


                if (q == "1")
                {

                    g2 = hg.ToString() + "�~��4�u";
                }
                else
                {
                    int hgg = Convert.ToInt16(q) - 1;
                    g2 = year2 + "�~��" + hgg.ToString() + "�u";
                }
                g3 = year2 + "�~��" + q.ToString() + "�u";
                Category(4, "", "Account_Temp6");



                gh = "�i���͹�~(��)���q-��" + q + "�u�禬�P�W�u���";
                gh2 = "�P�W�u����W(��)";

                DFF = GetAAq(q, year2);
            }
            else if (comboBox6.Text == "�~")
            {
                g2 = hg.ToString() + "�~";
                g3 = year2 + "�~";
                gh2 = "�P�h�~����W(��)";
                gh = "�i���͹�~(��)���q-" + year2 + "�~�禬�P�h�~���";
                Category(5, "", "Account_Temp6");
                DFF = GetAAy();

            }

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\�禬���.xls";
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);



            ExcelReport.EUN(FileName, OutPutFile, DFF, g2, g3, gh, gh2);
        }



        System.Data.DataTable GetAA()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'TFT-3rd parties sub total-���J',isnull(sum(gtotal),0) 'A',(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH) 'B' ");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(6), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(6), DATEADD(month, 1, dDate),112) =@MONTH");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MONTH", textBox6.Text));


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

        System.Data.DataTable GetAAq(string q, string year)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'TFT-3rd parties sub total-���J',isnull(sum(gtotal),0) 'A',(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append("   ) 'B'  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }

            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append("  ) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append("  ) from Account_Temp6 where cardgroup='103'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'LED-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='104' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='104'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='105' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='105'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'LED-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='104'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='105'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='103'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'LED-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='104'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'10' and @year+'12') ");
            }
            sb.Append(" ) from Account_Temp6 where cardgroup='105'");
            if (q == "1")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year2+'10' and @year2+'12') ");
            }
            else if (q == "2")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'01' and @year+'03') ");
            }
            else if (q == "3")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'04' and @year+'06') ");
            }
            else if (q == "4")
            {
                sb.Append(" and (Convert(varchar(6),dDate,112)  between @year+'07' and @year+'09') ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            int f = Convert.ToInt32(year);
            string year2 = Convert.ToString(f - 1);
            command.Parameters.Add(new SqlParameter("@year", year));
            command.Parameters.Add(new SqlParameter("@year2", year2));


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
        System.Data.DataTable GetAAy()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select 'TFT-3rd parties sub total-���J',isnull(sum(gtotal),0) 'A',(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH) 'B' ");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-���J',isnull(sum(gtotal),0),(select isnull(sum(gtotal),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-����',isnull(sum(GVALUE),0),(select isnull(sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" UNION ALL");
            sb.Append(" ");
            sb.Append(" select 'TFT-3rd parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode not in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Related Parties sub total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103' and cardcode  in ('0257-00' , '0511-00') ");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'TFT-Grand Total-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='103' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='103'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'LED-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='104' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='104'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");
            sb.Append(" union all");
            sb.Append(" select 'SOLAR-��Q',isnull(sum(GTOTAL)-sum(GVALUE),0),(select isnull(sum(GTOTAL)-sum(GVALUE),0) from Account_Temp6 where cardgroup='105' ");
            sb.Append(" and  Convert(varchar(4), dDate,112) =@MONTH)");
            sb.Append("  from Account_Temp6 where cardgroup='105'");
            sb.Append(" and  Convert(varchar(4), DATEADD(year, 1, dDate),112) =@MONTH");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MONTH", textBox6.Text.Substring(0, 4)));


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

        private void button16_Click(object sender, EventArgs e)
        {
            string ����;
            string �渹;
            string ���;
            string ��ئW��;
            decimal �ƶq = 0;
            string Debit;
            string Credit;
            string Balance;
            string ����;
            //try
            //{
            System.Data.DataTable dt = GetACC();



            System.Data.DataTable dtCost = MakeTableAcc();
            System.Data.DataTable dtDoc = null;
            DataRow dr = null;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();

                ��� = Convert.ToString(dt.Rows[i]["���"]);
                ��ئW�� = Convert.ToString(dt.Rows[i]["��ئW��"]);
                //   �ƶq = Convert.ToString(dt.Rows[i]["�ƶq"]);
                Debit = Convert.ToString(dt.Rows[i]["Debit"]);
                Credit = Convert.ToString(dt.Rows[i]["Credit"]);
                Balance = Convert.ToString(dt.Rows[i]["Balance"]);
                ���� = Convert.ToString(dt.Rows[i]["����"]);


                dr["���"] = ���;
                dr["��ئW��"] = ��ئW��;

                dr["Debit"] = Debit;
                dr["Credit"] = Credit;
                dr["Balance"] = Balance;

                // dtCost.Rows.Add(dr);
                �ƶq = 0;

                dtDoc = GetOINV(���, ����);

                if (dtDoc.Rows.Count > 0)
                {

                    �ƶq = Convert.ToDecimal(dtDoc.Rows[0]["�ƶq"]);
                }

                dr["�ƶq"] = �ƶq.ToString();
                if (���� == "11111")
                {
                    ���� = "TFT";
                }
                if (���� == "11121")
                {
                    ���� = "LED";
                }
                if (���� == "11131")
                {
                    ���� = "SOLAR";
                }

                dr["����"] = ����;
                dtCost.Rows.Add(dr);

            }

            dataGridView8.DataSource = dtCost;
        
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {

            //    Category(6, comboBox8.SelectedValue.ToString(), "Account_Temp61");
                string Account_Temp6 = "Account_Temp6" + comboBox8.Text;

                string �Ȥ�s��t;
                string �Ȥ�W��t;
                string �~��t;
   
                System.Data.DataTable dtemp5 = GetTemp61(Account_Temp6);
                System.Data.DataTable dtCostDD = MakeTableQuar();

                DataRow drtemp5 = null;
                for (int i = 0; i <= dtemp5.Rows.Count - 1; i++)
                {
                    j = i + 1;
                    drtemp5 = dtCostDD.NewRow();
                    �Ȥ�s��t = dtemp5.Rows[i]["�Ȥ�s��"].ToString();
                    �Ȥ�W��t = dtemp5.Rows[i]["�Ȥ�W��"].ToString();
                    �~��t = dtemp5.Rows[i]["�~��"].ToString();
       
                    drtemp5["row"] = j.ToString();
                    drtemp5["BU"] = dtemp5.Rows[i]["BU"].ToString();
                    drtemp5["�~��"] = �~��t;
                    drtemp5["�Ȥ�s��"] = "'" + �Ȥ�s��t;
                    drtemp5["�Ȥ�W��"] = �Ȥ�W��t;
       
                 //   drtemp5["���~�W��"] = ���~�W��;
                    for (int y = 1; y <= 4; y++)
                    {
                        string sg = "";
                        string sg2 = "";
                        string sg3 = "";

                        System.Data.DataTable dh = null;
                        dh = GetTemp5_1Q(y, �Ȥ�s��t, �~��t, "Q", Account_Temp6);
                        sg = dh.Rows[0]["�ƶq"].ToString();
                        drtemp5[y + "_Q"] = sg;

                        System.Data.DataTable dh2 = null;
                        dh2 = GetTemp5_1Q(y, �Ȥ�s��t, �~��t, "R", Account_Temp6);
                        sg2 = dh2.Rows[0]["���B"].ToString();
                        drtemp5[y + "_R"] = sg2;

                        System.Data.DataTable dh3 = null;
                        dh3 = GetTemp5_1Q(y, �Ȥ�s��t, �~��t, "C", Account_Temp6);
                        sg3 = dh3.Rows[0]["����"].ToString();
                        drtemp5[y + "_C"] = sg3;

                        double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                        int �P�����B = Convert.ToInt32(sg2);
                        double �Q��� = (�Q�� / (�P�����B)) * 100;


                        drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
                        string aa = Convert.ToString(�Q���);
                        if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
                        {
                            drtemp5[y + "_P"] = 0;
                        }
                        else
                        {
                            drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
                        }



                    }

                    drtemp5["T_Q"] = Convert.ToInt32(drtemp5["1_Q"]) + Convert.ToInt32(drtemp5["2_Q"]) + Convert.ToInt32(drtemp5["3_Q"])
    + Convert.ToInt32(drtemp5["4_Q"]);
                    DJ = Convert.ToInt32(drtemp5["1_R"]) + Convert.ToInt32(drtemp5["2_R"]) + Convert.ToInt32(drtemp5["3_R"])
    + Convert.ToInt32(drtemp5["4_R"]);
                    drtemp5["T_R"] = DJ;
                    drtemp5["T_C"] = Convert.ToInt32(drtemp5["1_C"]) + Convert.ToInt32(drtemp5["2_C"]) + Convert.ToInt32(drtemp5["3_C"])
    + Convert.ToInt32(drtemp5["4_C"]);
                    drtemp5["T_G"] = Convert.ToInt32(drtemp5["1_G"]) + Convert.ToInt32(drtemp5["2_G"]) + Convert.ToInt32(drtemp5["3_G"])
    + Convert.ToInt32(drtemp5["4_G"]);


                    double �Q���1 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt32(drtemp5["T_R"])) * 100;
                    string aa2 = Convert.ToString(�Q���1);
                    if (aa2 == "���L�a�j" || aa2 == "�t�L�a�j" || aa2 == "���O�@�ӼƦr")
                    {
                        drtemp5["T_P"] = 0;
                    }
                    else
                    {
                        drtemp5["T_P"] = �Q���1.ToString("#0.00") + "%";
                    }


                    dtCostDD.Rows.Add(drtemp5);





                }
    //            drtemp5 = dtCostDD.NewRow();
    //            drtemp5["�~��"] = "";
    //            drtemp5["�Ȥ�s��"] = "";
    //            drtemp5["�Ȥ�W��"] = "�[�`";

    //       //     drtemp5["SALES2"] = "";

    //            for (int y = 1; y <= 4; y++)
    //            {
    //                string sg = "";
    //                string sg2 = "";
    //                string sg3 = "";


    //                System.Data.DataTable dh = null;
    //                dh = GetTemp5_1tQ(y, "Q");
    //                sg = dh.Rows[0]["�ƶq"].ToString();
    //                drtemp5[y + "_Q"] = sg;

    //                System.Data.DataTable dh2 = null;
    //                dh2 = GetTemp5_1tQ(y, "R");
    //                sg2 = dh2.Rows[0]["���B"].ToString();
    //                drtemp5[y + "_R"] = sg2;

    //                System.Data.DataTable dh3 = null;
    //                dh3 = GetTemp5_1tQ(y, "C");
    //                sg3 = dh3.Rows[0]["����"].ToString();
    //                drtemp5[y + "_C"] = sg3;

    //                double �Q�� = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                int �P�����B = Convert.ToInt32(sg2);
    //                double �Q��� = (�Q�� / (�P�����B)) * 100;


    //                drtemp5[y + "_G"] = Convert.ToInt32(sg2) - Convert.ToInt32(sg3);
    //                string aa = Convert.ToString(�Q���);
    //                if (aa == "���L�a�j" || aa == "�t�L�a�j" || aa == "���O�@�ӼƦr")
    //                {
    //                    drtemp5[y + "_P"] = 0;
    //                }
    //                else
    //                {
    //                    drtemp5[y + "_P"] = �Q���.ToString("#0.00") + "%";
    //                }

    //            }

    //            drtemp5["T_Q"] = Convert.ToInt64(drtemp5["1_Q"]) + Convert.ToInt64(drtemp5["2_Q"]) + Convert.ToInt64(drtemp5["3_Q"])
    //                    + Convert.ToInt64(drtemp5["4_Q"]);
    //            string F = drtemp5["1_R"].ToString();
    //            string F2 = drtemp5["2_R"].ToString();
    //            string F3 = drtemp5["3_R"].ToString();
    //            string F4 = drtemp5["4_R"].ToString();
    //            Int64 FG = Convert.ToInt64(drtemp5["1_R"]) + Convert.ToInt64(drtemp5["2_R"]) + Convert.ToInt64(drtemp5["3_R"])
    //+ Convert.ToInt64(drtemp5["4_R"]);
    //            drtemp5["T_R"] = FG;
    //            drtemp5["T_C"] = Convert.ToInt64(drtemp5["1_C"]) + Convert.ToInt64(drtemp5["2_C"]) + Convert.ToInt64(drtemp5["3_C"])
    //+ Convert.ToInt64(drtemp5["4_C"]);
    //            drtemp5["T_G"] = Convert.ToInt64(drtemp5["1_G"]) + Convert.ToInt64(drtemp5["2_G"]) + Convert.ToInt64(drtemp5["3_G"])
    //+ Convert.ToInt64(drtemp5["4_G"]);

    //            double �Q���6 = (Convert.ToDouble(drtemp5["T_G"]) / Convert.ToInt64(drtemp5["T_R"])) * 100;
    //            string aa6 = Convert.ToString(�Q���6);
    //            if (aa6 == "���L�a�j" || aa6 == "�t�L�a�j" || aa6 == "���O�@�ӼƦr")
    //            {
    //                drtemp5["T_P"] = 0;
    //            }
    //            else
    //            {
    //                drtemp5["T_P"] = �Q���6.ToString("#0.00") + "%";
    //            }

    //            dtCostDD.Rows.Add(drtemp5);

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\���J����by�u.xlsx";


                //Excel���˪���
                string ExcelTemplate = FileName;

                //��X��
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                string FF = comboBox2.SelectedValue.ToString();
                //���� Excel Report
                ExcelReport.TEMP61(dtCostDD, ExcelTemplate, OutPutFile, FF);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = Get2();
            ExcelReport.GridViewToExcel(dataGridView8);
        }
        private System.Data.DataTable Get2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUBSTRING(T3.GROUPNAME,3,15) BU,t1.boardCountNo �T������,t0.ShippingCode �u�渹�X,item �O��,amount ���B,t0.cardname ������,subcompany �l���q,DocDate ���,doccur ���O,doccur1 �ײv FROM dbo.Shipping_Fee T0 ");
            sb.Append(" LEFT JOIN SHIPPING_MAIN T1 ON(T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" left join acmesql02.dbo.ocrd t2 on (t2.cardcode=t1.cardcode COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" LEFT JOIN acmesql02.dbo.OCRG T3 ON (T2.GROUPCODE = T3.GROUPCODE) ");
            sb.Append(" WHERE  T0.INSDATE BETWEEN @AA AND @BB and  isnull(feecheck,'False')='true' ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AA", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox8.Text));
   
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

        private System.Data.DataTable SHIP()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.CLOSEDAY ������,T1.ITEMCODE �Ƹ�,T1.QUANTITY �ƶq,T1.ITEMPRICE ���,T1.ITEMAMOUNT ���B FROM ACMESQLSP.DBO.SHIPPING_MAIN T0 ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SHIPPING_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE IN (SELECT U_SHIPPING_NO COLLATE Chinese_Taiwan_Stroke_CI_AS from ACMESQL02.DBO.OPOR WHERE CARDCODE='U0019') AND T0.CLOSEDAY BETWEEN @AA AND @BB");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AA", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox8.Text));

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


        private void button19_Click_1(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("�п���ɮ�");
                }
                else
                {

                    AUALICE(opdf.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void AUALICE(string ExcelFile)
        {
            StringBuilder sb = new StringBuilder();
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
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
       

            for (int i = 1; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString().Trim();




                try
                {

                    if (!String.IsNullOrEmpty(id))
                    {




                        sb.Append("'" + id + "',");
                        

                    

                       // AddProduct("", id, id2, id3, id4, id5, id6, id7, id8, id9, id10, id11, id12, id13, id14, id15, id16, id17, id18, id19, id20, id21, id22, id23, id24, id25);

                    }


                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
            MessageBox.Show("�פJ���\");
            sb.Remove(sb.Length - 1, 1);


            textBox9.Text = sb.ToString();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetItemm(textBox9.Text);
            dataGridView8.DataSource = dt;
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        private void button21_Click(object sender, EventArgs e)
        {

            string ����;
            string �渹 ;
            string ��� ;
            string ��ئW��;
            decimal  �ƶq=0 ;
            string Debit;
            string Credit;
            string Balance ;
            string ����;
            //try
            //{
                System.Data.DataTable dt = GetACC();



                System.Data.DataTable dtCost = MakeTableAcc();
                System.Data.DataTable dtDoc = null;
                DataRow dr = null;

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dr = dtCost.NewRow();

                     ��� = Convert.ToString(dt.Rows[i]["���"]);
                     ��ئW�� = Convert.ToString(dt.Rows[i]["��ئW��"]);
                  //   �ƶq = Convert.ToString(dt.Rows[i]["�ƶq"]);
                     Debit = Convert.ToString(dt.Rows[i]["Debit"]);
                     Credit = Convert.ToString(dt.Rows[i]["Credit"]);
                     Balance = Convert.ToString(dt.Rows[i]["Balance"]);
                     ���� = Convert.ToString(dt.Rows[i]["����"]);

                
                     dr["���"] = ���;
                     dr["��ئW��"] = ��ئW��;
                   
                     dr["Debit"] = Debit;
                     dr["Credit"] = Credit;
                     dr["Balance"] = Balance;
                    
                    // dtCost.Rows.Add(dr);
                     �ƶq = 0;
                  
                         dtDoc = GetOINV(���, ����);
                    
                    if(dtDoc.Rows.Count > 0 )
                    {

                        �ƶq = Convert.ToDecimal(dtDoc.Rows[0]["�ƶq"]);
                         }
                
                     dr["�ƶq"] = �ƶq.ToString();
                     if (���� == "11111")
                     {
                         ���� = "TFT";
                     }
                     if (���� == "11121")
                     {
                         ���� = "LED";
                     }
                     if (���� == "11131")
                     {
                         ���� = "SOLAR";
                     }

                     dr["����"] = ����;
                     dtCost.Rows.Add(dr);
      
                }

                dataGridView8.DataSource = dtCost;
                ExcelReport.GridViewToExcel(dataGridView8);

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }




        private void button22_Click(object sender, EventArgs e)
        {
            //try
            //{
                A1 = textBox3.Text.Substring(0, 6) + "01";
                A2 = textBox3.Text;
                EunLED();
                
                Eun22();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\���J����by�g.xls";


                //Excel���˪���
                string ExcelTemplate = FileName;

                //��X��
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                string dh1 = textBox3.Text.Trim();
                string FIRSTDATE = dh1.Substring(0, 4) + "/" + dh1.Substring(4, 2) + "/01";
                string LASTDATE;
                if (dh1 == DateTime.Now.ToString("yyyyMM"))
                {
                    LASTDATE = GetMenu.Day();
                }
                else
                {
                    LASTDATE = GetMenu.DLast2(textBox3.Text);
                }

                LASTDATE = LASTDATE.Substring(0, 4) + "/" + LASTDATE.Substring(4, 2) + "/" + LASTDATE.Substring(6, 2);

                //���� Excel Report
                ExcelReport.ACC(dtCostDD3, ExcelTemplate, OutPutFile, FIRSTDATE, LASTDATE, dtCostEun);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void button23_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = SHIP();
            ExcelReport.GridViewToExcel(dataGridView8);
        }

 

        private void button25_Click(object sender, EventArgs e)
        {
            string FD = textBox16.Text;
            string DD = FD.Substring(6, 2);

            int FG = Convert.ToInt16(DD);
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\�Ȥ����Ʀ�.xls";


            //Excel���˪���
            string ExcelTemplate = FileName;

            //��X��
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //���� Excel Report
            ExcelReport.ExcelReportOutput(GetSHIP(FD, FG), ExcelTemplate, OutPutFile, "N");
        }

        public System.Data.DataTable GetSHIP(string D1,int FG)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT add9 ���渹�X,arriveDay �������,ITEMCODE �Ƹ�,DSCRIPTION �~�W,CAST(ITEMPRICE AS DECIMAL(10,2)) ���,CAST(T1.QUANTITY AS INT) �ƶq,CAST(CAST(ITEMPRICE AS DECIMAL(10,2))*T1.QUANTITY AS INT) ���B");
            sb.Append(" ,CASE ISNULL(RATE,0) WHEN 0 THEN (SELECT TOP 1 CAST(BUY AS DECIMAL(10,2)) FROM SHIPBUY) ELSE CAST(RATE AS DECIMAL(10,2)) END �ײv");
            sb.Append(" ,(CASE ISNULL(RATE,0) WHEN 0 THEN (SELECT TOP 1 CAST(BUY AS DECIMAL(10,2)) FROM SHIPBUY) ELSE CAST(RATE AS DECIMAL(10,2)) END)*CAST(CAST(ITEMPRICE AS DECIMAL(10,2))*T1.QUANTITY AS INT) �x�����B  FROM SHIPPING_MAIN T0 ");
            sb.Append(" LEFT JOIN SHIPPING_ITEM T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.ortt T2 ON (T0.CLOSEDAY=Convert(varchar(10),ratedate,112) AND CURRENCY='USD')");
            sb.Append(" WHERE substring(add9,1,1)='A' and BoardCountNo='�i�f' ");
            if (FG <= 15)
            {
                sb.Append(" AND arriveDay between Convert(varchar(6),dateadd(month,-1,@D1),112)+'01' AND @D1");
            }
            else
            {
                sb.Append(" AND arriveDay between Convert(varchar(6),dateadd(month,0,@D1),112)+'01' AND @D1");

            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@D1", D1));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["data"];
        }

        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dtCostDD = MakeTableMQY();

                DataRow drtemp5 = null;

                string YRAR = comboBox7.Text;


                drtemp5 = dtCostDD.NewRow();
                drtemp5["NAME"] = "��ئW��";
                for (int y = 1; y <= 12; y++)
                {
                    drtemp5[y + "_M"] = "'" + comboBox7.Text + " / " + y.ToString();
                }
                for (int j = 1; j <= 4; j++)
                {


                    drtemp5[j + "_Q"] = comboBox7.Text + " / " + "��" + j + "�u";
                }

                drtemp5["Y"] = comboBox7.Text + "��~��";
                dtCostDD.Rows.Add(drtemp5);


                drtemp5 = dtCostDD.NewRow();
                drtemp5["NAME"] = "�P�f���J�b�B";
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = GetMQY("M",4, YRAR, y);
  
                        drtemp5[y + "_M"] = dh.Rows[0]["���B"].ToString();
                    }
                    for (int j = 1; j <= 4; j++)
                    {
      
                        System.Data.DataTable dhQ = GetMQY("Q",4, YRAR, j);
                        drtemp5[j + "_Q"] = dhQ.Rows[0]["���B"].ToString();
                    }


                    System.Data.DataTable dhY = GetMQY("Y",4, YRAR, 0);

                    drtemp5["Y"] = dhY.Rows[0]["���B"].ToString();
                    dtCostDD.Rows.Add(drtemp5);


                    System.Data.DataTable dg2 = GetMQY2("Y", YRAR, 0,"123");
                    for (int i = 0; i <= dg2.Rows.Count-1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        string NAME = dg2.Rows[i]["���"].ToString();
                        drtemp5["NAME"] = NAME;
                        for (int y = 1; y <= 12; y++)
                        {

                            System.Data.DataTable dg3 = GetMQY2("M", YRAR, y, NAME);
                            if (dg3.Rows.Count > 0)
                            {
                                drtemp5[y + "_M"] = dg3.Rows[0]["���B"].ToString();
                            }
                          
                        }
                        for (int j = 1; j <= 4; j++)
                        {

                            System.Data.DataTable dg4 = GetMQY2("Q", YRAR, j, NAME);
                            if (dg4.Rows.Count > 0)
                            {
                                drtemp5[j + "_Q"] = dg4.Rows[0]["���B"].ToString();
                            }
                        }

                        System.Data.DataTable dg5 = GetMQY2("Y2", YRAR, 0, NAME);


                        drtemp5["Y"] = dg5.Rows[0]["���B"].ToString();
                        dtCostDD.Rows.Add(drtemp5);
                    }




                    drtemp5 = dtCostDD.NewRow();
                    drtemp5["NAME"] = "�P�f����";
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = GetMQY("M", 5, YRAR, y);

                        drtemp5[y + "_M"] = dh.Rows[0]["���B"].ToString();
                    }
                    for (int j = 1; j <= 4; j++)
                    {

                        System.Data.DataTable dhQ = GetMQY("Q", 5, YRAR, j);
                        drtemp5[j + "_Q"] = dhQ.Rows[0]["���B"].ToString();
                    }


                    System.Data.DataTable dhY2 = GetMQY("Y", 5, YRAR, 0);

                    drtemp5["Y"] = dhY2.Rows[0]["���B"].ToString();
                    dtCostDD.Rows.Add(drtemp5);


                    drtemp5 = dtCostDD.NewRow();
                    drtemp5["NAME"] = "�P�f��Q";
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = GetMQY3("M", YRAR, y);

                        drtemp5[y + "_M"] = dh.Rows[0]["���B"].ToString();
                    }
                    for (int j = 1; j <= 4; j++)
                    {

                        System.Data.DataTable dhQ = GetMQY3("Q", YRAR, j);
                        drtemp5[j + "_Q"] = dhQ.Rows[0]["���B"].ToString();
                    }


                    System.Data.DataTable dhY3 = GetMQY3("Y", YRAR, 0);

                    drtemp5["Y"] = dhY3.Rows[0]["���B"].ToString();
                    dtCostDD.Rows.Add(drtemp5);


                    drtemp5 = dtCostDD.NewRow();
                    drtemp5["NAME"] = "��Q�v";
                    for (int y = 1; y <= 12; y++)
                    {

                        System.Data.DataTable dh = GetMQY3("M", YRAR, y);

                        drtemp5[y + "_M"] = "'" + Convert.ToDecimal(dh.Rows[0]["��Q"]).ToString("#0.00") + "%";
                    }
                    for (int j = 1; j <= 4; j++)
                    {

                        System.Data.DataTable dhQ = GetMQY3("Q", YRAR, j);
                        drtemp5[j + "_Q"] = "'" + Convert.ToDecimal(dhQ.Rows[0]["��Q"]).ToString("#0.00") + "%";
         
                    }


                    System.Data.DataTable dhY4 = GetMQY3("Y", YRAR, 0);

                    drtemp5["Y"] = "'" + Convert.ToDecimal(dhY4.Rows[0]["��Q"]).ToString("#0.00") + "%";
                    dtCostDD.Rows.Add(drtemp5);


                    drtemp5 = dtCostDD.NewRow();
                    dtCostDD.Rows.Add(drtemp5);
                    drtemp5 = dtCostDD.NewRow();
                    dtCostDD.Rows.Add(drtemp5);



                    drtemp5 = dtCostDD.NewRow();
                    drtemp5["NAME"] = "����";
                    for (int y = 1; y <= 12; y++)
                    {
                        drtemp5[y + "_M"] = "'"+comboBox7.Text + " / " + y.ToString();
                    }
                    for (int j = 1; j <= 4; j++)
                    {


                        drtemp5[j + "_Q"] = comboBox7.Text + " / " + "��" + j + "�u";
                    }

                    drtemp5["Y"] = comboBox7.Text + "��~��";
                    dtCostDD.Rows.Add(drtemp5);


                    System.Data.DataTable dj2 = GetMQY4("Y", YRAR, 0, "1234", 4);
                    for (int i = 0; i <= dj2.Rows.Count - 1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        string NAME = dj2.Rows[i]["����"].ToString();
                        drtemp5["NAME"] = NAME.Replace("�Ʒ~��", "") + " -���J";
                        for (int y = 1; y <= 12; y++)
                        {

                            System.Data.DataTable dg3 = GetMQY4("M", YRAR, y, NAME, 4);
                            if (dg3.Rows.Count > 0)
                            {
                                drtemp5[y + "_M"] = dg3.Rows[0]["���B"].ToString();
                            }

                        }
                        for (int j = 1; j <= 4; j++)
                        {

                            System.Data.DataTable dg4 = GetMQY4("Q", YRAR, j, NAME, 4);
                            if (dg4.Rows.Count > 0)
                            {
                                drtemp5[j + "_Q"] = dg4.Rows[0]["���B"].ToString();
                            }
                        }

                        System.Data.DataTable dg5 = GetMQY4("Y2", YRAR, 0, NAME, 4);


                        drtemp5["Y"] = dg5.Rows[0]["���B"].ToString();
                        dtCostDD.Rows.Add(drtemp5);
                    }


                    drtemp5 = dtCostDD.NewRow();
                    dtCostDD.Rows.Add(drtemp5);

                    System.Data.DataTable dj4 = GetMQY4("Y", YRAR, 0, "1234", 5);
                    for (int i = 0; i <= dj4.Rows.Count - 1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        string NAME = dj4.Rows[i]["����"].ToString();
                        drtemp5["NAME"] = NAME.Replace("�Ʒ~��", "") + " -����";
                        for (int y = 1; y <= 12; y++)
                        {

                            System.Data.DataTable dg3 = GetMQY4("M", YRAR, y, NAME, 5);
                            if (dg3.Rows.Count > 0)
                            {
                                drtemp5[y + "_M"] = dg3.Rows[0]["���B"].ToString();
                            }

                        }
                        for (int j = 1; j <= 4; j++)
                        {

                            System.Data.DataTable dg4 = GetMQY4("Q", YRAR, j, NAME, 5);
                            if (dg4.Rows.Count > 0)
                            {
                                drtemp5[j + "_Q"] = dg4.Rows[0]["���B"].ToString();
                            }
                        }

                        System.Data.DataTable dg5 = GetMQY4("Y2", YRAR, 0, NAME, 5);


                        drtemp5["Y"] = dg5.Rows[0]["���B"].ToString();
                        dtCostDD.Rows.Add(drtemp5);
                    }


                    drtemp5 = dtCostDD.NewRow();
                    dtCostDD.Rows.Add(drtemp5);


                    System.Data.DataTable dj3 = GetMQY5("Y", YRAR, 0, "1234");
                    for (int i = 0; i <= dj3.Rows.Count - 1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        string NAME = dj3.Rows[i]["����"].ToString();
                        drtemp5["NAME"] = NAME.Replace("�Ʒ~��", "") + " -��Q";
                        for (int y = 1; y <= 12; y++)
                        {

                            System.Data.DataTable dg3 = GetMQY5("M", YRAR, y, NAME);
                            if (dg3.Rows.Count > 0)
                            {
                                drtemp5[y + "_M"] = dg3.Rows[0]["��Q"].ToString();
                            }

                        }
                        for (int j = 1; j <= 4; j++)
                        {

                            System.Data.DataTable dg4 = GetMQY5("Q", YRAR, j, NAME);
                            if (dg4.Rows.Count > 0)
                            {
                                drtemp5[j + "_Q"] = dg4.Rows[0]["��Q"].ToString();
                            }
                        }

                        System.Data.DataTable dg5 = GetMQY5("Y2", YRAR, 0, NAME);


                        drtemp5["Y"] = dg5.Rows[0]["��Q"].ToString();
                        dtCostDD.Rows.Add(drtemp5);
                    }

                    drtemp5 = dtCostDD.NewRow();
                    dtCostDD.Rows.Add(drtemp5);

                    System.Data.DataTable dj5 = GetMQY5("Y", YRAR, 0, "1234");
                    for (int i = 0; i <= dj5.Rows.Count - 1; i++)
                    {
                        drtemp5 = dtCostDD.NewRow();
                        string NAME = dj5.Rows[i]["����"].ToString();

                        drtemp5["NAME"] = NAME.Replace("�Ʒ~��", "") + " -��Q�v";
                        for (int y = 1; y <= 12; y++)
                        {

                            System.Data.DataTable dg3 = GetMQY5("M", YRAR, y, NAME);
                            if (dg3.Rows.Count > 0)
                            {

                                drtemp5[y + "_M"] = "'" + Convert.ToDecimal(dg3.Rows[0]["��Q�v"]).ToString("#0.00") + "%";
                            }

                        }
                        for (int j = 1; j <= 4; j++)
                        {

                            System.Data.DataTable dg4 = GetMQY5("Q", YRAR, j, NAME);
                            if (dg4.Rows.Count > 0)
                            {
                                drtemp5[j + "_Q"] = "'" + Convert.ToDecimal(dg4.Rows[0]["��Q�v"]).ToString("#0.00") + "%";
                            }

                            System.Data.DataTable dg5 = GetMQY5("Y2", YRAR, 0, NAME);


                            drtemp5["Y"] = "'" + Convert.ToDecimal(dg5.Rows[0]["��Q�v"]).ToString("#0.00") + "%";
                           
                        }
                        dtCostDD.Rows.Add(drtemp5);
                    }
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\�P�f�l�q.xls";


                //Excel���˪���
                string ExcelTemplate = FileName;

                //��X��
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                string FF = comboBox2.SelectedValue.ToString();
                //���� Excel Report
                ExcelReport.ExcelReportOutput(dtCostDD, ExcelTemplate, OutPutFile, FF);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;

    
                LookupValues = GetMenu.GetMenuListS();
            
            if (LookupValues != null)
            {
                textBox17.Text = Convert.ToString(LookupValues[0]);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;


            LookupValues = GetMenu.GetMenuListS();

            if (LookupValues != null)
            {
                textBox18.Text = Convert.ToString(LookupValues[0]);
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = GetAP();
            ExcelReport.GridViewToExcel(dataGridView8);
        }

        System.Data.DataTable GetAP()
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.SHORTNAME �t�ӽs��,T3.CARDNAME �t�ӦW��, Convert(varchar(8),T0.REFDATE,112)  �L�b���,T0.TRANSID ���,[dbo].[fun_SAPDOC](T0.TRANSTYPE)  �ӷ�,T0.BASEREF 'Orinin no',T1.CONTRAACT �R�P���,T1.DEBIT �ɶ�,T1.CREDIT �U��");
            sb.Append(" ,T2.AMT �|�B,T1.LINEMEMO �Ƶ� FROM OJDT T0");
            sb.Append(" LEFT JOIN (SELECT  T0.TRANSID,MAX(T0.SHORTNAME) SHORTNAME");
            sb.Append(" ,MAX(CASE SUBSTRING(T0.CONTRAACT,1,1) WHEN 'U' THEN '' WHEN 'S' THEN '' ELSE T0.CONTRAACT END ) CONTRAACT ");
            sb.Append(" ,MAX(CASE SUBSTRING(T0.CONTRAACT,1,1) WHEN 'U' THEN '' WHEN 'S' THEN '' ELSE LINEMEMO END ) LINEMEMO");
            sb.Append(" ,MAX(CASE SUBSTRING(T0.CONTRAACT,1,1) WHEN 'U' THEN 0 WHEN 'S' THEN 0 ELSE DEBIT END ) DEBIT");
            sb.Append(" ,MAX(CASE SUBSTRING(T0.CONTRAACT,1,1) WHEN 'U' THEN 0 WHEN 'S' THEN 0 ELSE CREDIT END ) CREDIT");
            sb.Append("  FROM JDT1 T0 GROUP BY T0.TRANSID) T1 ON (T0.TRANSID=T1.TRANSID)");
            sb.Append(" LEFT JOIN (SELECT  T0.TRANSID,SUM(DEBIT-CREDIT) AMT FROM JDT1 T0 WHERE T0.ACCOUNT='12640101' GROUP BY T0.TRANSID) T2 ON (T0.TRANSID=T2.TRANSID)");
            sb.Append(" Left JOIN OCRD T3 on (T1.SHORTNAME=T3.CARDCODE)");
        
            sb.Append(" WHERE   Convert(varchar(8),T0.REFDATE,112) between @CC AND @DD ");
            if (textBox17.Text != "" && textBox17.Text != "")
            {
                sb.Append(" AND T1.SHORTNAME between @AA AND @BB ");

            }
            else
            {
                sb.Append(" AND  SUBSTRING(T1.SHORTNAME,1,1) IN ('S','U') ");
            }

            sb.Append(" ORDER BY T1.SHORTNAME,T0.TRANSID");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox17.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox18.Text));
            command.Parameters.Add(new SqlParameter("@CC", textBox19.Text));
            command.Parameters.Add(new SqlParameter("@DD", textBox20.Text));
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

        private void radioButton4_Click(object sender, EventArgs e)
        {
            textBox17.Text = "S0";
            textBox18.Text = "SZ";
        }

        private void radioButton5_Click(object sender, EventArgs e)
        {
            textBox17.Text = "U0";
            textBox18.Text = "UZ";
        }

        private void button24_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetDIST();
            dataGridView8.DataSource = dt;
            ExcelReport.GridViewToExcel(dataGridView8);
        }



        private void button29_Click(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            Category(8, "W", "Account_Temp6" + "2021");

            MessageBox.Show("�w����");
        }

        private void button7_Click(object sender, EventArgs e)
        {

            A1 = textBox1.Text;
            A2 = textBox2.Text;
            Category(8, "W", "Account_Temp61" + YEAROITM);

            MessageBox.Show("�w����");
        }

        private void button30_Click(object sender, EventArgs e)
        {
            string DOC=textBox21.Text;
            System.Data.DataTable H1 = GetBAALANCE(DOC);
            if (H1.Rows.Count > 0)
            {
                UPDATEBALANCE(textBox21.Text);
                MessageBox.Show("�w����");
            }
            else
            {
                MessageBox.Show("SAP�S�����u��");
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {

            string YEAR = comboBox8.Text;


            ACME.Form1Rpt4SALES frm4 = new ACME.Form1Rpt4SALES();


            frm4.dt = GetItemmSALES(COM, "Account_Temp6" + YEAR);
            frm4.s = YEAR + "0101";

            frm4.q = GetItemmSALESDATE("Account_Temp6" + YEAR).Rows[0][0].ToString();
            frm4.ShowDialog();
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
        private void button32_Click(object sender, EventArgs e)
        {

            string YEAR = comboBox8.Text;


            ACME.SALESOITM frm4 = new ACME.SALESOITM();


            frm4.dt = GetSALESOITM(COM, "Account_Temp61" + YEAR);
            frm4.s = YEAR + "0101";
            frm4.q = GetItemmSALESDATE("Account_Temp6" + YEAR).Rows[0][0].ToString();
            frm4.ShowDialog();
        }

        private void button33_Click(object sender, EventArgs e)
        {

        }

        private void button33_Click_1(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            string YEAR = A1.Substring(0, 4);
            int M1 = Convert.ToInt16(A1.Substring(4, 2));
            int M2 = Convert.ToInt16(A2.Substring(4, 2));

            if (radioButton1.Checked)
            {
                System.Data.DataTable DT1 = null;
                ACME.Form1Rpt4F frm4 = new ACME.Form1Rpt4F();
                if (M2 - M1 != 0 && (globals.DBNAME != "�F�ͥ�"))
                {
                    DT1 = GetItemmDDF(COM, "Account_Temp61" + YEAR);
                }
                else
                {
                    Category(8, "", "Account_Temp61");
                    DT1 = GetItemmDDF(COM, "Account_Temp61");
                }
                System.Data.DataTable dtCost = MakeTableOCRDF();
                DataRow dr = null;

                for (int i = 0; i <= DT1.Rows.Count - 1; i++)
                {
                    DataRow dd = DT1.Rows[i];
                    dr = dtCost.NewRow();
                    string CARDCODE = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�s��"] = CARDCODE;
                    dr["�Ȥ�s��"] = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�W��"] = dd["�Ȥ�W��"].ToString();
                    dr["�ƶq"] = Convert.ToDecimal(dd["�ƶq"]);
                    dr["�P�����B"] = Convert.ToDecimal(dd["�P�����B"]);
                    dr["�P������"] = Convert.ToDecimal(dd["�P������"]);
                    dr["�Q��"] = Convert.ToDecimal(dd["�Q��"]);
                    dr["�Q���"] = dd["�Q���"].ToString();
                    dr["COM"] = dd["COM"].ToString();

                    //SHIPTO
                    StringBuilder sb2 = new StringBuilder();
                    System.Data.DataTable DT2 = GetItemmDDOCRD(CARDCODE, "Account_Temp61");
                    if (DT2.Rows.Count > 0)
                    {
                        for (int il = 0; il <= DT2.Rows.Count - 1; il++)
                        {
                            string COUNTRY = DT2.Rows[il]["COUNTRY"].ToString();
                            sb2.Append(COUNTRY + "/");

                        }
                        sb2.Remove(sb2.Length - 1, 1);
                        dr["��a"] = sb2.ToString();
                    }
                    dr["SHIPTO"] = dd["SHIPTO"].ToString();
                    dtCost.Rows.Add(dr);
                }
                frm4.dt = dtCost;
                frm4.s = textBox1.Text;
                frm4.q = textBox2.Text;
                frm4.ShowDialog();
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            string YEAR = A1.Substring(0, 4);
            int M1 = Convert.ToInt16(A1.Substring(4, 2));
            int M2 = Convert.ToInt16(A2.Substring(4, 2));
            System.Data.DataTable DT1 = null;
            ACME.Form1RptACC frm4 = new ACME.Form1RptACC();
            if (M2 - M1 != 0 && (globals.DBNAME != "�F�ͥ�"))
            {
                DT1 = GetItemmDDACC(COM, "Account_Temp6" + YEAR);
            }
            else
            {
                Category(8, "", "Account_Temp6");
                DT1 = GetItemmDDACC(COM, "Account_Temp6");
            }
            System.Data.DataTable dtCost = MakeTableOCRDACC();
            DataRow dr = null;

            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                DataRow dd = DT1.Rows[i];
                dr = dtCost.NewRow();
                string CARDCODE = dd["�Ȥ�s��"].ToString();
                dr["�Ȥ�s��"] = CARDCODE;
                dr["�Ȥ�s��"] = dd["�Ȥ�s��"].ToString();
                dr["�Ȥ�W��"] = dd["�Ȥ�W��"].ToString();
                dr["�ƶq"] = Convert.ToDecimal(dd["�ƶq"]);
                dr["�P�����B"] = Convert.ToDecimal(dd["�P�����B"]);
                dr["�P������"] = Convert.ToDecimal(dd["�P������"]);
                dr["�Q��"] = Convert.ToDecimal(dd["�Q��"]);
                dr["�Q���"] = dd["�Q���"].ToString();
                dr["��إN��"] = dd["��إN��"].ToString();
                dtCost.Rows.Add(dr);
            }
            frm4.dt = dtCost;
            frm4.s = textBox1.Text;
            frm4.q = textBox2.Text;
            frm4.ShowDialog();
        }

        private void button35_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\NANCY.xlsx";

            string YEAR = comboBox8.Text;
            //Excel���˪���
            string ExcelTemplate = FileName;

            //��X��
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //���� Excel Report
            ExcelReport.ExcelNANCY(GETNANCY("Account_Temp6" + YEAR), ExcelTemplate, OutPutFile, GETNANCY2("Account_Temp6" + YEAR));
        }

        private void button36_Click(object sender, EventArgs e)
        {
            A1 = textBox1.Text;
            A2 = textBox2.Text;
            string YEAR = A1.Substring(0, 4);
            int M1 = Convert.ToInt16(A1.Substring(4, 2));
            int M2 = Convert.ToInt16(A2.Substring(4, 2));


                System.Data.DataTable DT1 = null;
                ACME.Form1Rpt4F frm4 = new ACME.Form1Rpt4F();
                if (M2 - M1 != 0 && (globals.DBNAME != "�F�ͥ�"))
                {
                    DT1 = GetItemmDDF(COM, "Account_Temp61" + YEAR);
                }
                else
                {
                    Category(8, "", "Account_Temp61");
                    DT1 = GetItemmDDF(COM, "Account_Temp61");
                }
                System.Data.DataTable dtCost = MakeTableOCRDF();
                DataRow dr = null;

                for (int i = 0; i <= DT1.Rows.Count - 1; i++)
                {
                    DataRow dd = DT1.Rows[i];
                    dr = dtCost.NewRow();
                    string CARDCODE = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�s��"] = CARDCODE;
                    dr["�Ȥ�s��"] = dd["�Ȥ�s��"].ToString();
                    dr["�Ȥ�W��"] = dd["�Ȥ�W��"].ToString();
                    dr["�ƶq"] = Convert.ToDecimal(dd["�ƶq"]);
                    dr["�P�����B"] = Convert.ToDecimal(dd["�P�����B"]);
                    dr["�P������"] = Convert.ToDecimal(dd["�P������"]);
                    dr["�Q��"] = Convert.ToDecimal(dd["�Q��"]);
                    dr["�Q���"] = dd["�Q���"].ToString();
                    dr["COM"] = dd["COM"].ToString();

                    //SHIPTO
                    StringBuilder sb2 = new StringBuilder();
                    System.Data.DataTable DT2 = GetItemmDDOCRD(CARDCODE, "Account_Temp61");
                    if (DT2.Rows.Count > 0)
                    {
                        for (int il = 0; il <= DT2.Rows.Count - 1; il++)
                        {
                            string COUNTRY = DT2.Rows[il]["COUNTRY"].ToString();
                            sb2.Append(COUNTRY + "/");

                        }
                        sb2.Remove(sb2.Length - 1, 1);
                        dr["��a"] = sb2.ToString();
                    }
                    dr["SHIPTO"] = dd["SHIPTO"].ToString();
                    dtCost.Rows.Add(dr);
                }
                frm4.dt = dtCost;
                frm4.s = textBox1.Text;
                frm4.q = textBox2.Text;
                frm4.ShowDialog();
       
        }


     



    }







}


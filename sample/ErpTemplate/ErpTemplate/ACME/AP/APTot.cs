using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace ACME
{
    public partial class APTot : Form
    {
        System.Data.DataTable dtCost;
        public APTot()
        {
            InitializeComponent();
        }
        private int ss;
        private System.Data.DataTable GetSAPRevenue1()
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select  case t1.targettype when 20 then t1.trgetentry else '' end 收採單號,t0.comments SAP備註, cast(month(t0.docdate) as varchar)+'月訂單' 訂單月份,''進貨數, case substring(t0.cardname,1,2) when '友達' then 'AUO-'+substring(T0.cardcode,7,3) else t0.cardname end 廠商,T1.U_ACME_DSCRIPTION PAYMENT ");
            sb.Append(" ,cast(t1.gtotalfc as decimal(12,2)) USD,cast(t1.totalfrgn as decimal(12,2)) USD1,cast(t1.gtotalfc-t1.totalfrgn as decimal(12,2)) USD2,t0.docentry docentry,t1.linenum linenum,  ");
            sb.Append(" T11.U_TMODEL 品名 ,T11.U_GRADE GD,U_VERSION Ver,cast(t1.quantity as int) 數量,cast(t1.price as decimal(12,2)) 單價,cast(t1.rate as decimal(16,4)) 匯率");
            sb.Append(" ,CASE WHEN  TV.REPLY IS NULL THEN t0.U_acme_place1 ELSE TV.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS 出貨地,CASE WHEN  TU.REPLY IS NULL THEN t0.u_acme_meth+t0.u_acme_cus ELSE TU.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS  出貨方式,t0.cardname BU,case when t3.clgcode is not null then 'Link' end appath ");
            sb.Append("  ,SUBSTRING(CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("               SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("               SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("              AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            sb.Append("       Substring (T1.[ItemCode],2,8) END,1,4) 尺寸");
            sb.Append(" ,t4.reply 備註,t0.u_acme_discoo 折讓,t5.color1,t5.ROWNAME,TS.shipping_no ,T1.ITEMCODE  from opor t0");
            sb.Append(" left join por1 t1 on (t0.docentry=t1.docentry)");
            sb.Append("   left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
            sb.Append(" left join OCRD TD on (t0.CARDCODE=TD.CARDCODE)");
            if (comboBox1.SelectedValue.ToString() == "其他")
            {
                sb.Append(" left join oclg t3 on (t1.docentry=t3.docentry and t3.doctype='22')");
            }
            else
            {
                sb.Append(" left join oclg t3 on (t1.trgetentry=t3.docentry and t3.doctype='20')");
            }
            sb.Append(" left join acmesqlsp.dbo.aplc2 t4 on (t1.docentry=t4.docentry and t1.linenum=t4.linenum)");
            sb.Append(" left join acmesqlsp.dbo.aplc3 t5 on (t1.docentry=t5.docentry and t1.linenum=t5.linenum) ");
            sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC4 TU ON (T1.DOCENTRY=TU.DOCENTRY AND T1.LINENUM=TU.LINENUM)");
            sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC5 TV ON (T1.DOCENTRY=TV.DOCENTRY AND T1.LINENUM=TV.LINENUM)");
            sb.Append(" left join (select MAX(T1.u_shipping_no) shipping_no,MAX(T1.DOCDATE) DOCDATE,ITEMCODE,isnull(substring(max(Convert(varchar(10),t1.u_acme_invoice,111)),6,6)+'-','')+isnull(max(t1.u_acme_inv),'')  invoice,t0.baseentry,t0.baseQTY ");
            sb.Append(" from pdn1 t0 left join opdn t1 on (t0.docentry=t1.docentry)");
            sb.Append(" group by baseentry,baseQTY,ITEMCODE) ts");
            sb.Append(" on (t1.docentry=tS.baseentry and  ts.BASEQTY=t1.QUANTITY AND TS.ITEMCODE=T1.ITEMCODE)");
            sb.Append(" where  T0.CANCELED <> 'Y'  AND ISNULL(CAST(TRGETENTRY AS VARCHAR),'')+LINESTATUS <> 'C'  ");
            if (textBox2.Text != "")
            {
                sb.Append("  and t1.trgetentry='" + textBox2.Text.ToString() + "'");
            }
            else if (textBox1.Text != "")
            {
                sb.Append("  and TS.invoice like '%" + textBox1.Text.ToString() + "%' ");
            }
            else if (textBox3.Text != "")
            {
                sb.Append("  and t0.docentry='" + textBox3.Text.ToString() + "'");
            }
            else if (textBox4.Text != "")
            {
                sb.Append("  and TS.shipping_no ='" + textBox4.Text.ToString() + "'");
            }
            else
            {
          
                    sb.Append("  and    year(T0.docdate)='" + comboBox2.SelectedValue.ToString() + "'    ");
                    if (comboBox3.Text  != "")
                    {
                        //ALL未結PO
                        if (comboBox3.Text == "ALL未結PO")
                        {
                            //DOCSTATUS='C'
                            sb.Append(" AND T1.LINESTATUS='O'  ");
                        }
                        else
                        {
                            sb.Append(" and month(T0.docdate)='" + comboBox3.SelectedValue.ToString() + "'   ");
                        }
                    }

            //if (comboBox1.SelectedValue.ToString() == "零件")
            //{
            //    sb.Append("  AND  substring(t0.cardname,1,2) not in ('友達')  AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE NOT IN ('S0007','S0006','S0029','S0028','S0040','S0046','S0077','S0070','S0082','S0080')    ");

            //}
            //else
                if (comboBox1.SelectedValue.ToString() == "其他")
                {
                    sb.Append("  and TD.GROUPCODE = 101  and (substring(T0.cardcode,1,5) <> 'S0001' or substring(T0.cardcode,1,5) <> 'S0623') AND T0.CARDCODE NOT IN ('U0361','U0193')  ");
                    //  sb.Append("  and T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080')  ");

                }
                else if (comboBox1.SelectedValue.ToString() == "AUO")
                {
                    sb.Append("   and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
                    sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                    sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                    sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                    sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
                    sb.Append("                      AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");

                }
                else if (comboBox1.SelectedValue.ToString() == "全部")
                {
                    sb.Append("     and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
                    sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                    sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                    sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                    sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
                    sb.Append("                     AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");
                    sb.Append("                      OR (T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080'))  ");
                    sb.Append("                      OR (substring(t0.cardname,1,2) not in ('友達')  AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE NOT IN ('S0007','S0006','S0029','S0028','S0040','S0046','S0077','S0070','S0082','S0080'))");
                }
                else if (comboBox1.SelectedValue.ToString() == "豐藝")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0028'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "凌巨")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0465'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "達運")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0462'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "博豐")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0005'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "瑞威")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0014'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "晟統")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0329'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "展驛")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0121'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "香港譽天")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0509'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "RCH")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0575'  ");
                }
                else if (comboBox1.SelectedValue.ToString() == "天馬")
                {
                    sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0207-03'  ");
                }
                else
                {
                    sb.Append("               and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR  substring(t1.itemcode,1,3) in ('KTC') OR ");
                    sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                    sb.Append("         SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                    sb.Append("         SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                    sb.Append("        AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%' )");

                  //  sb.Append(" AND  substring(t0.cardname,11,12) ='" + comboBox1.SelectedValue.ToString() + "'  ");
                    sb.Append(" AND  SUBSTRING(T0.CARDCODE,7,3) ='" + comboBox1.SelectedValue.ToString() + "'  ");
                }
            }
            // sb.Append(" UNION ALL ");
            // sb.Append("             select   case t1.targettype when 20 then t1.trgetentry else '' end 收採單號,t0.comments SAP備註, cast(month(t0.docdate) as varchar)+'月訂單' 訂單月份,''進貨數, case substring(t0.cardname,1,2) when '友達' then 'AUO-'+substring(T0.cardcode,7,3) else t0.cardname end 廠商,T1.U_ACME_DSCRIPTION PAYMENT, ");
            // sb.Append("               cast(t1.gtotalfc as decimal(12,2)) USD,cast(t1.totalfrgn as decimal(12,2)) USD1,cast(t1.gtotalfc-t1.totalfrgn as decimal(12,2)) USD2,t0.docentry docentry,t1.linenum linenum,   ");
            // sb.Append("              OT.U_TMODEL 品名 , ");
            // sb.Append("               CASE (Substring(T1.[ItemCode],11,1))  ");
            // sb.Append("                when 'A' then 'A' when 'B' then 'B' when '0' then 'Z'  ");
            // sb.Append("               when '1' then 'P' when '2' then 'N' when '3' then 'V'  ");
            // sb.Append("               when '4' then 'U' when '5' then 'NN' ELSE 'X' ");
            // sb.Append(" END GD,(Substring(T1.[ItemCode],12,3)) Ver,cast(t1.quantity as int) 數量,cast(t1.price as decimal(12,2)) 單價,cast(t1.rate as decimal(16,4)) 匯率");
            // sb.Append(" ,CASE WHEN  TV.REPLY IS NULL THEN t0.U_acme_place1 ELSE TV.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS 出貨地,CASE WHEN  TU.REPLY IS NULL THEN t0.u_acme_meth+t0.u_acme_cus ELSE TU.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS  出貨方式,t0.cardname BU,case when t3.clgcode is not null then 'Link' end appath ");
            // sb.Append("  ,SUBSTRING(CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            // sb.Append("               SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            // sb.Append("               SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
            // sb.Append("              AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
            // sb.Append("       Substring (T1.[ItemCode],2,8) END,1,4) 尺寸");
            // sb.Append(" ,t4.reply 備註,t0.u_acme_discoo 折讓,t5.color1,TS.shipping_no,T1.ITEMCODE from opor t0");
            // sb.Append(" left join por1 t1 on (t0.docentry=t1.docentry)");
            // sb.Append("               left join OITM OT on (t1.ITEMCODE=OT.ITEMCODE) ");
            // sb.Append(" left join OCRD TD on (t0.CARDCODE=TD.CARDCODE)");
            // if (comboBox1.SelectedValue.ToString() == "其他")
            // {
            //     sb.Append(" left join oclg t3 on (t1.docentry=t3.docentry and t3.doctype='22')");
            // }
            // else
            // {
            //     sb.Append(" left join oclg t3 on (t1.trgetentry=t3.docentry and t3.doctype='20')");
            // }
            // sb.Append(" left join acmesqlsp.dbo.aplc2 t4 on (t1.docentry=t4.docentry and t1.linenum=t4.linenum)");
            // sb.Append(" left join acmesqlsp.dbo.aplc3 t5 on (t1.docentry=t5.docentry and t1.linenum=t5.linenum) ");
            // sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC4 TU ON (T1.DOCENTRY=TU.DOCENTRY AND T1.LINENUM=TU.LINENUM)");
            // sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC5 TV ON (T1.DOCENTRY=TV.DOCENTRY AND T1.LINENUM=TV.LINENUM)");
            // sb.Append(" left join (select MAX(T1.u_shipping_no) shipping_no,MAX(T1.DOCDATE) DOCDATE,ITEMCODE,isnull(substring(max(Convert(varchar(10),t1.u_acme_invoice,111)),6,6)+'-','')+isnull(max(t1.u_acme_inv),'')  invoice,t0.baseentry,t0.baseQTY ");
            // sb.Append(" from pdn1 t0 left join opdn t1 on (t0.docentry=t1.docentry)");
            // sb.Append(" group by baseentry,baseQTY,ITEMCODE) ts");
            // sb.Append(" on (t1.docentry=tS.baseentry and  ts.BASEQTY=t1.QUANTITY AND TS.ITEMCODE=T1.ITEMCODE)");
            // sb.Append(" where  T0.CANCELED <> 'Y'  AND  TS.DOCDATE is  null  AND ISNULL(CAST(TRGETENTRY AS VARCHAR),'')+LINESTATUS <> 'C'   ");
            // if (textBox2.Text != "")
            // {
            //     sb.Append("  and t1.trgetentry='" + textBox2.Text.ToString() + "'");
            // }
            // else if (textBox1.Text != "")
            // {
            //     sb.Append("  and TS.invoice like '%" + textBox1.Text.ToString() + "%' ");
            // }
            // else if (textBox3.Text != "")
            // {
            //     sb.Append("  and t0.docentry='" + textBox3.Text.ToString() + "'");
            // }
            // else if (textBox4.Text != "")
            // {
            //     sb.Append("  and TS.shipping_no ='" + textBox4.Text.ToString() + "'");
            // }
            // else
            // {

            //     sb.Append("  and    year(ts.docdate)='" + comboBox2.SelectedValue.ToString() + "'    ");
            //     if (comboBox3.Text != "")
            //     {
            //         sb.Append(" and month(ts.docdate)='" + comboBox3.SelectedValue.ToString() + "'   ");
            //     }

            //// if (comboBox1.SelectedValue.ToString() == "零件")
            //// {
            ////     sb.Append("  AND  substring(t0.cardname,1,2) not in ('友達')  AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE NOT IN ('S0007','S0006','S0029','S0028','S0040','S0046','S0077','S0070','S0082','S0080')    ");

            //// }
            //     //else
            //         if (comboBox1.SelectedValue.ToString() == "其他")
            //     {
            //         sb.Append("  and TD.GROUPCODE = 101  and substring(T0.cardcode,1,5) <> 'S0001'  AND T0.CARDCODE NOT IN ('U0361','U0193')  ");
            //         //     sb.Append("  and T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080')  ");

            //     }
            //     else if (comboBox1.SelectedValue.ToString() == "AUO")
            // {
            //     sb.Append("   and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
            //     sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            //     sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            //     sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
            //     sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
            //     sb.Append("                      AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");

            // }
            // else if (comboBox1.SelectedValue.ToString() == "全部")
            // {
            //     sb.Append("     and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
            //     sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            //     sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            //     sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
            //     sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
            //     sb.Append("                     AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");
            //     sb.Append("                      OR (T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080'))  ");
            //     sb.Append("                      OR (substring(t0.cardname,1,2) not in ('友達')  AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE NOT IN ('S0007','S0006','S0029','S0028','S0040','S0046','S0077','S0070','S0082','S0080'))");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "豐藝")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0028'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "凌巨")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0465'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "達運")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0462'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "博豐")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0005'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "瑞威")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0014'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "晟統")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0329'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "展驛")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0121'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "香港譽天")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0509'  ");
            // }
            // else if (comboBox1.SelectedValue.ToString() == "RCH")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0575'  ");
            // }
            //    else if (comboBox1.SelectedValue.ToString() == "天馬")
            // {
            //     sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0207-03'  ");
            // }
            // else
            // {
            //     sb.Append("               and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
            //     sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            //     sb.Append("         SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            //     sb.Append("         SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
            //     sb.Append("        AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%' )");
            //     sb.Append(" AND  substring(t0.cardname,11,12) ='" + comboBox1.SelectedValue.ToString() + "'  ");

            // }

            //  }
            if (comboBox3.Text != "") 
            {
                sb.Append(" UNION ");
                //採購報價單
                sb.Append(" select  case t1.targettype when 20 then t1.trgetentry else '' end 收採單號,t0.comments SAP備註, cast(month(t0.docdate) as varchar)+'月訂單' 訂單月份,''進貨數, case substring(t0.cardname,1,2) when '友達' then 'AUO-'+substring(T0.cardcode,7,3) else t0.cardname end 廠商,T1.U_ACME_DSCRIPTION PAYMENT ");
                sb.Append(" ,cast(t1.gtotalfc as decimal(12,2)) USD,cast(t1.totalfrgn as decimal(12,2)) USD1,cast(t1.gtotalfc-t1.totalfrgn as decimal(12,2)) USD2,t0.docentry docentry,t1.linenum linenum,  ");
                sb.Append(" T11.U_TMODEL 品名 ,T11.U_GRADE GD,U_VERSION Ver,cast(t1.quantity as int) 數量,cast(t1.price as decimal(12,2)) 單價,cast(t1.rate as decimal(16,4)) 匯率");
                sb.Append(" ,CASE WHEN  TV.REPLY IS NULL THEN t0.U_acme_place1 ELSE TV.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS 出貨地,CASE WHEN  TU.REPLY IS NULL THEN t0.u_acme_meth+t0.u_acme_cus ELSE TU.REPLY END COLLATE Chinese_Taiwan_Stroke_CI_AS  出貨方式,t0.cardname BU,case when t3.clgcode is not null then 'Link' end appath ");
                sb.Append("  ,SUBSTRING(CASE WHEN SUBSTRING(T1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                sb.Append("               SUBSTRING(T1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                sb.Append("               SUBSTRING(T1.ITEMCODE,3,1) LIKE '[0-9]%'");
                sb.Append("              AND SUBSTRING(T1.ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring (T1.[ItemCode],1,9)  ELSE ");
                sb.Append("       Substring (T1.[ItemCode],2,8) END,1,4) 尺寸");
                sb.Append(" ,t4.reply 備註,t0.u_acme_discoo 折讓,t5.color1,t5.ROWNAME,T0.U_Shipping_no shipping_no ,T1.ITEMCODE  from opqt t0");
                sb.Append(" left join pqt1 t1 on (t0.docentry=t1.docentry)");
                sb.Append("   left JOIN OITM T11 ON T1.ITEMCODE = T11.ITEMCODE  ");
                sb.Append(" left join OCRD TD on (t0.CARDCODE=TD.CARDCODE)");
                if (comboBox1.SelectedValue.ToString() == "其他")
                {
                    sb.Append(" left join oclg t3 on (t1.docentry=t3.docentry and t3.doctype='22')");
                }
                else
                {
                    sb.Append(" left join oclg t3 on (t1.trgetentry=t3.docentry and t3.doctype='20')");
                }
                sb.Append(" left join acmesqlsp.dbo.aplc2 t4 on (t1.docentry=t4.docentry and t1.linenum=t4.linenum)");
                sb.Append(" left join acmesqlsp.dbo.aplc3 t5 on (t1.docentry=t5.docentry and t1.linenum=t5.linenum) ");
                sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC4 TU ON (T1.DOCENTRY=TU.DOCENTRY AND T1.LINENUM=TU.LINENUM)");
                sb.Append(" LEFT JOIN  ACMESQLSP.DBO.APLC5 TV ON (T1.DOCENTRY=TV.DOCENTRY AND T1.LINENUM=TV.LINENUM)");
                sb.Append(" where  T0.CANCELED <> 'Y'  AND ISNULL(CAST(TRGETENTRY AS VARCHAR),'')+LINESTATUS <> 'C'  ");
                if (textBox2.Text != "")
                {
                    sb.Append("  and t1.trgetentry='" + textBox2.Text.ToString() + "'");
                }
                else if (textBox1.Text != "")
                {
                    sb.Append("  and TS.invoice like '%" + textBox1.Text.ToString() + "%' ");
                }
                else if (textBox3.Text != "")
                {
                    sb.Append("  and t0.docentry='" + textBox3.Text.ToString() + "'");
                }
                else if (textBox4.Text != "")
                {
                    sb.Append("  and TS.shipping_no ='" + textBox4.Text.ToString() + "'");
                }
                else
                {

                    sb.Append("  and    year(T0.docdate)='" + comboBox2.SelectedValue.ToString() + "'    ");
                    if (comboBox3.Text != "")
                    {
                        //ALL未結PO
                        if (comboBox3.Text == "ALL未結PO")
                        {
                            //DOCSTATUS='C'
                            sb.Append(" AND T1.LINESTATUS='O'  ");
                        }
                        else
                        {
                            sb.Append(" and month(T0.docdate)='" + comboBox3.SelectedValue.ToString() + "'   ");
                        }
                    }
                    if (comboBox1.SelectedValue.ToString() == "其他")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  and (substring(T0.cardcode,1,5) <> 'S0001' or substring(T0.cardcode,1,5) <> 'S0623') AND T0.CARDCODE NOT IN ('U0361','U0193')  ");
                        //  sb.Append("  and T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080')  ");

                    }
                    else if (comboBox1.SelectedValue.ToString() == "AUO")
                    {
                        sb.Append("   and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
                        sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                        sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                        sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                        sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
                        sb.Append("                      AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");

                    }
                    else if (comboBox1.SelectedValue.ToString() == "全部")
                    {
                        sb.Append("     and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR substring(t1.itemcode,1,3) in ('KTC') OR ");
                        sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                        sb.Append("     SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                        sb.Append("     SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                        sb.Append("    AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%')");
                        sb.Append("                     AND  substring(t0.cardname,11,12) in ('AV','PID','DD','GD','NB','TV')");
                        sb.Append("                      OR (T0.CARDCODE IN ('S0028','S0040','S0046','S0077','S0070','S0082','S0080'))  ");
                        sb.Append("                      OR (substring(t0.cardname,1,2) not in ('友達')  AND SUBSTRING(T0.CARDCODE,1,1)='S' AND T0.CARDCODE NOT IN ('S0007','S0006','S0029','S0028','S0040','S0046','S0077','S0070','S0082','S0080'))");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "豐藝")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0028'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "凌巨")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0465'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "達運")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0462'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "博豐")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0005'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "瑞威")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0014'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "晟統")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0329'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "展驛")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0121'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "香港譽天")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0509'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "RCH")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0575'  ");
                    }
                    else if (comboBox1.SelectedValue.ToString() == "天馬")
                    {
                        sb.Append("  and TD.GROUPCODE = 101  AND T0.CARDCODE ='S0207-03'  ");
                    }
                    else
                    {
                        sb.Append("               and  ( substring(t1.itemcode,1,1) in ('T','I','V','O','A') OR  substring(t1.itemcode,1,3) in ('KTC') OR ");
                        sb.Append(" SUBSTRING(t1.ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
                        sb.Append("         SUBSTRING(t1.ITEMCODE,2,1) LIKE '[0-9]%' AND ");
                        sb.Append("         SUBSTRING(t1.ITEMCODE,3,1) LIKE '[0-9]%'");
                        sb.Append("        AND SUBSTRING(t1.ITEMCODE,4,1) LIKE '[0-9]%' )");

                        //  sb.Append(" AND  substring(t0.cardname,11,12) ='" + comboBox1.SelectedValue.ToString() + "'  ");
                        sb.Append(" AND  SUBSTRING(T0.CARDCODE,7,3) ='" + comboBox1.SelectedValue.ToString() + "'  ");
                    }
                }
            }
            
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 3600;
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
        private System.Data.DataTable GetSAPRevenue2(string aa,string bb)
        {
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select isnull(cast(cast(T1.quantity as int) as varchar),'') 交貨數量,isnull(cast(t2.quantity as int),0) 退貨數量,isnull(substring(Convert(varchar(10),t0.u_acme_invoice,111),6,6)+'-','') 時間,");
            sb.Append(" isnull(cast(T0.u_acme_inv as varchar),'') inv from opdn t0");
            sb.Append(" left join pdn1 t1 on(t0.docentry=t1.docentry)");
            sb.Append(" left join rpd1 t2 on (t2.baseentry=t1.docentry and t2.baseline=t1.linenum and t2.basetype='20')");
            sb.Append(" where t1.baseentry=@aa and t1.baseline=@bb and t1.basetype='22'");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            //

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
        private void APTot_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.BU(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Year(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetMenu.Month2(), "DataValue", "DataValue");

            comboBox2.Text = DateTime.Now.ToString("yyyy");
            comboBox3.Text = DateTime.Now.ToString("MM"); 
          
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "APPATH")
                {
                    if (comboBox1.SelectedValue.ToString() == "其他")
                    {
                        string sd = dataGridView1.CurrentRow.Cells["docentry"].Value.ToString();
                    
                        System.Data.DataTable dt1 = oclgohters(sd);

                        for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                        {
                            DataRow drw = dt1.Rows[j];



                            System.Diagnostics.Process.Start(drw["path"].ToString() + "\\" + drw["路徑"].ToString());

                        }

                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }
                    else
                    {
                  
                        string sd = dataGridView1.CurrentRow.Cells["收採單號"].Value.ToString();
         
                        System.Data.DataTable dt1 = oclg(sd);

                        for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                        {
                            DataRow drw = dt1.Rows[j];



                            System.Diagnostics.Process.Start(drw["path"].ToString() + "\\" + drw["路徑"].ToString());

                        }

                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;


                    }
                }
                if (dgv.Columns[e.ColumnIndex].Name == "船務工單")
                {
                    if (dgv.Columns[e.ColumnIndex].Name == "船務工單")
                    {
                        string sd = dataGridView1.CurrentRow.Cells["船務工單"].Value.ToString();

                        System.Data.DataTable dt1 = GetOPDN(sd);

                        for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                        {
                            DataRow drw = dt1.Rows[j];



                            System.Diagnostics.Process.Start(drw["path"].ToString() + "\\" + drw["路徑"].ToString());

                        }

                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);
            }
        }
        public static System.Data.DataTable GetOPDN(string shippingcode)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t4.docentry 收貨採購單號,t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑,FILENAME 檔案名稱 from oclg t2");
            sb.Append("  LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append(" inner join opdn t4 on(t2.docentry=t4.docentry)");
            sb.Append(" where  t2.doctype='20' and t4.u_shipping_no=@shippingcode");
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
        public System.Data.DataTable oclg(string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑 from oclg t2");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append(" where DOCENTRY=@aa and t2.doctype='20'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", aa));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "oclg");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["oclg"];
        }
        public System.Data.DataTable oclgohters(string aa)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t3.TRGTPATH [path],'\'+CAST(T3.[FILENAME]  AS nVARCHAR(80) )+'.'+Fileext 路徑 from oclg t2");
            sb.Append(" LEFT JOIN ATC1 T3 ON (T2.ATCENTRY=T3.ABSENTRY)");
            sb.Append(" where DOCENTRY=@aa and t2.doctype='22'");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", aa));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "oclg");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["oclg"];
        }
        private void button1_Click(object sender, EventArgs e)
        {
          
      
            string 訂單月份;
            string 廠商;
            string Payment;
            string USD;
            string USD1;
            string USD2;
            string docentry;
            string 品名;
            string GD;
            string Ver;
            string 數量;
            string 單價;
            string 匯率;
            string 出貨地;
            string 出貨方式;
            string 進貨數;
            string 折讓;
            string 備註;
            string APPATH;
            string color1;
            string ROWNAME;
            string linenum;
            string BU;
 
            string SAP備註;
            string 尺寸;
            string 船務工單;
            string 收採單號;
            System.Data.DataTable dt = GetSAPRevenue1();
            System.Data.DataTable dt1 = null;
       
              dtCost = MakeTableCombine();
               DataRow dr = null;
               for (int i = 0; i <= dt.Rows.Count - 1; i++)
               {
                   docentry = dt.Rows[i]["docentry"].ToString();
                   linenum = dt.Rows[i]["linenum"].ToString();
                   dt1 = GetSAPRevenue2(docentry, linenum);
                   dr = dtCost.NewRow();
                   尺寸 = dt.Rows[i]["尺寸"].ToString();
                   string ITEMCODE = dt.Rows[i]["ITEMCODE"].ToString();
                   if (!String.IsNullOrEmpty(ITEMCODE))
                   {
                       string T1 = ITEMCODE.Substring(0, 2);
                       if (T1 == "OM")
                       {
                           尺寸 = "O" + 尺寸.Replace("M", "");
                       }
                       訂單月份 = dt.Rows[i]["訂單月份"].ToString();
                       廠商 = dt.Rows[i]["廠商"].ToString();
                       Payment = dt.Rows[i]["Payment"].ToString();
                       USD = dt.Rows[i]["USD"].ToString();
                       USD1 = dt.Rows[i]["USD1"].ToString();
                       USD2 = dt.Rows[i]["USD2"].ToString();
                       SAP備註 = dt.Rows[i]["SAP備註"].ToString();
                       品名 = dt.Rows[i]["品名"].ToString();
                       GD = dt.Rows[i]["GD"].ToString();
                       Ver = dt.Rows[i]["Ver"].ToString();
                       數量 = dt.Rows[i]["數量"].ToString();
                       單價 = dt.Rows[i]["單價"].ToString();
                       匯率 = dt.Rows[i]["匯率"].ToString();
                       出貨地 = dt.Rows[i]["出貨地"].ToString();
                       出貨方式 = dt.Rows[i]["出貨方式"].ToString();
                       進貨數 = dt.Rows[i]["進貨數"].ToString();
                       折讓 = dt.Rows[i]["折讓"].ToString();
                       備註 = dt.Rows[i]["備註"].ToString();
                       APPATH = dt.Rows[i]["APPATH"].ToString();
                       color1 = dt.Rows[i]["color1"].ToString();
                       ROWNAME = dt.Rows[i]["ROWNAME"].ToString();
                       船務工單 = dt.Rows[i]["shipping_no"].ToString();
                       BU = dt.Rows[i]["BU"].ToString();
                       收採單號 = dt.Rows[i]["收採單號"].ToString();
                       dr["訂單月份"] = 訂單月份;
                       dr["廠商"] = 廠商;
                       dr["Payment"] = Payment;
                       dr["USD"] = USD;
                       dr["USD1"] = USD1;
                       dr["USD2"] = USD2;
                       dr["docentry"] = docentry;
                       dr["品名"] = 品名;
                       dr["GD"] = GD;
                       dr["Ver"] = Ver;
                       dr["尺寸"] = 尺寸;

                       dr["單價"] = 單價;
                       dr["匯率"] = 匯率;
                       dr["出貨地"] = 出貨地;
                       dr["出貨方式"] = 出貨方式;
                       dr["備註"] = 備註;
                       dr["折讓"] = 折讓;
                       dr["SAP備註"] = SAP備註;
                       dr["APPATH"] = APPATH;
                       dr["color1"] = color1;
                       dr["ROWNAME"] = ROWNAME;
                       //
                       dr["linenum"] = linenum;
                       dr["BU"] = BU;
                       dr["船務工單"] = 船務工單;
                       dr["收採單號"] = 收採單號;
                       for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                       {
                           DataRow dv = dt1.Rows[j];
                           dr["進貨數"] += dv["時間"].ToString() + "/" + dv["inv"].ToString() + "*" + dv["交貨數量"].ToString() + "/";

                           ss = 0;
                           if (dv["退貨數量"].ToString() != "0")
                           {
                               ss += Convert.ToInt32(dv["退貨數量"]);
                           }
                       }
                       dr["數量"] = (Convert.ToInt32(數量) - ss);
                       if (dr["數量"].ToString() != "0")
                       {
                           dtCost.Rows.Add(dr);
                       }
                   }
               }


           

            dataGridView1.DataSource = dtCost;

        }
 
                private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("訂單月份", typeof(string));
            dt.Columns.Add("廠商", typeof(string));
            dt.Columns.Add("Payment", typeof(string));
            dt.Columns.Add("USD", typeof(string));
            dt.Columns.Add("USD1", typeof(string));
            dt.Columns.Add("USD2", typeof(string));
            dt.Columns.Add("docentry", typeof(string));
            dt.Columns.Add("尺寸", typeof(string));
                    
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("GD", typeof(string));
            dt.Columns.Add("Ver", typeof(string));
            dt.Columns.Add("數量", typeof(Int32));
            dt.Columns.Add("單價", typeof(string));
            dt.Columns.Add("匯率", typeof(string));
            dt.Columns.Add("出貨地", typeof(string));
            dt.Columns.Add("出貨方式", typeof(string));
            dt.Columns.Add("進貨數", typeof(string));
            dt.Columns.Add("折讓", typeof(string));
            dt.Columns.Add("船務工單", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("APPATH", typeof(string));
            dt.Columns.Add("color1", typeof(string));
            dt.Columns.Add("ROWNAME", typeof(string));
            //
            dt.Columns.Add("linenum", typeof(string));
            dt.Columns.Add("BU", typeof(string));
            dt.Columns.Add("SAP備註", typeof(string));
            dt.Columns.Add("收採單號", typeof(string));
            return dt;
        }
    
    
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcelAP(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {
        
                DataGridViewRow row;
            
                row = dataGridView1.Rows[i];
                string a0 = row.Cells["docentry"].Value.ToString();
                string a1 = row.Cells["linenum"].Value.ToString();
                string a3 = row.Cells["備註"].Value.ToString();
                string a4 = row.Cells["出貨地"].Value.ToString();
                string a6 = row.Cells["出貨方式"].Value.ToString();
                string a8 = row.Cells["折讓"].Value.ToString();
                string a10 = row.Cells["PAYMENT"].Value.ToString();
                //PAYMENT

                //更新備註
                System.Data.DataTable aa = Getshipitem2(a0,a1);
                if (aa.Rows.Count > 0)
                {
                
                    DeleteDetailSQL(a0, a1);
                    UpdateMasterSQL(a0, a1, a3);
               
                }
                else if (aa.Rows.Count.ToString() == "0" && a3 != "")
                {
                    DeleteDetailSQL(a0, a1);
                    UpdateMasterSQL(a0, a1, a3);
                }



                System.Data.DataTable bb = GetOPOR4(a0, a1);

                string a5 = bb.Rows[0]["出貨地"].ToString();

                if (a4 != a5)
                {
                    DeleteAPLC5(a0, a1);
                    UpdateAPLC5(a0, a1, a4);
                }


           
                string a7 = bb.Rows[0]["出貨方式"].ToString();

                if (a6 != a7)
                {
                    DeleteAPLC4(a0, a1);
                    UpdateAPLC4(a0, a1, a6);
                }



                string a9 = bb.Rows[0]["折讓"].ToString();
                if (a9 != a8)
                {
                    UpdateDISC(a8, a0);
                }



                string a11 = bb.Rows[0]["PAYMENT"].ToString();
                if (a10 != a11)
                {
                    UpdateDSC(a10, a0, a1);
                }
            }

            button1_Click(null, new EventArgs());
         
            


        }

        private void DeleteDetailSQL(string DocEntry, string LineNum)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  APLC2 where [DocEntry]=@DocEntry and [LineNum]=@LineNum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
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
        private void DeleteAPLC4(string DocEntry, string LineNum)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  APLC4 where [DocEntry]=@DocEntry and [LineNum]=@LineNum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
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
        private void DeleteAPLC5(string DocEntry, string LineNum)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  APLC5 where [DocEntry]=@DocEntry and [LineNum]=@LineNum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
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
        private void DeleteDetailSQL2(string DocEntry, string LineNum)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" delete  APLC3 where [DocEntry]=@DocEntry and [LineNum]=@LineNum ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
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
        private void UpdateMasterSQL(string DocEntry, string LineNum, string Reply)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO APLC2 (DocEntry,LineNum,Reply) VALUES (@DocEntry,@LineNum,@Reply)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
            command.Parameters.Add(new SqlParameter("@Reply", Reply));
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

        private void UpdateAPLC4(string DocEntry, string LineNum, string Reply)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO APLC4 (DocEntry,LineNum,Reply) VALUES (@DocEntry,@LineNum,@Reply)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
            command.Parameters.Add(new SqlParameter("@Reply", Reply));
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
        private void UpdateAPLC5(string DocEntry, string LineNum, string Reply)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO APLC5 (DocEntry,LineNum,Reply) VALUES (@DocEntry,@LineNum,@Reply)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
            command.Parameters.Add(new SqlParameter("@Reply", Reply));
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
        private void UpdateMasterSQL2(string DocEntry, string LineNum, string ShipTo, string ROWNAME)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" INSERT INTO APLC3 (DocEntry,LineNum,color1,ROWNAME) VALUES (@DocEntry,@LineNum,@ShipTo,@ROWNAME)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));
            command.Parameters.Add(new SqlParameter("@ShipTo", ShipTo));
            command.Parameters.Add(new SqlParameter("@ROWNAME", ROWNAME));
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
        public static System.Data.DataTable Getshipitem2(string docentry, string linenum)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select reply from aplc2 where docentry=@docentry and linenum=@linenum  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Rma_Item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Rma_Item"];
        }
        public static System.Data.DataTable GetAPLC4(string docentry, string linenum)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select reply from aplc4 where docentry=@docentry and linenum=@linenum  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "aplc4");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["aplc4"];
        }

        public static System.Data.DataTable GetOPOR4(string docentry, string linenum)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select U_ACME_PLACE1 出貨地,U_ACME_METH+U_acme_cus 出貨方式,u_acme_discoo 折讓,U_ACME_DSCRIPTION PAYMENT from OPOR T0 LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY) where T0.docentry=@docentry and linenum=@linenum  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@docentry", docentry));
            command.Parameters.Add(new SqlParameter("@linenum", linenum));
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
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
             
                    DataGridViewRow row;
                
                    for (int i = 0; i <= dataGridView1.SelectedRows.Count - 1; i++)
                    {
                        row = dataGridView1.SelectedRows[i];
                        string a0 = row.Cells["docentry"].Value.ToString();
                        string a1 = row.Cells["linenum"].Value.ToString();
                        string a3;
                        string a4 = comboBox4.Text;
                        DeleteDetailSQL2(a0, a1);


                        if (radioButton1.Checked)
                        {
                            a3 = "1";
                            UpdateMasterSQL2(a0, a1, a3, a4);
                        }
                        else if (radioButton2.Checked)
                        {
                            a3 = "2";
                            UpdateMasterSQL2(a0, a1, a3, a4);
                        }
                        else if (radioButton3.Checked)
                        {
                            a3 = "3";
                            UpdateMasterSQL2(a0, a1, a3, a4);
                        }
                        else if (radioButton4.Checked)
                        {
                            a3 = "4";
                            UpdateMasterSQL2(a0, a1, a3, a4);
                        }

                        else
                        {
                            MessageBox.Show("請選擇顏色");

                        }
                    }
            }
            catch
            {
                MessageBox.Show("請產生資料");

            
            }


            button1_Click(null, new EventArgs());
            
        }



    

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
                
            if   (e.RowIndex   >=   dataGridView1.Rows.Count) 
                                return; 
                        DataGridViewRow   dgr   =   dataGridView1.Rows[e.RowIndex]; 
                        try 
                        {
                            string ROW = dgr.Cells["ROWNAME"].Value.ToString();
                                if   (dgr.Cells[ "color1"].Value.ToString()=="1") 
                                {
                                    if (!String.IsNullOrEmpty(ROW))
                                    {
                                        dataGridView1.Rows[e.RowIndex].Cells[ROW].Style.BackColor = Color.Yellow;
                                    }
                                    else
                                    {
                                        dgr.DefaultCellStyle.BackColor = Color.Yellow;
                                    }
                                }
                                else if (dgr.Cells["color1"].Value.ToString() == "2")
                                {

                                    if (!String.IsNullOrEmpty(ROW))
                                    {
                                        dataGridView1.Rows[e.RowIndex].Cells[ROW].Style.BackColor = Color.YellowGreen;
                                    }
                                    else
                                    {
                                        dgr.DefaultCellStyle.BackColor = Color.YellowGreen;
                                    }

                                }
                                else if (dgr.Cells["color1"].Value.ToString() == "3")
                                {
                                    if (!String.IsNullOrEmpty(ROW))
                                    {
                                        dataGridView1.Rows[e.RowIndex].Cells[ROW].Style.BackColor = Color.Pink;
                                    }
                                    else
                                    {
                                        dgr.DefaultCellStyle.BackColor = Color.Pink;
                                    }
               
                                }
                                else if (dgr.Cells["color1"].Value.ToString() == "4")
                                {
                                    if (!String.IsNullOrEmpty(ROW))
                                    {
                                        dataGridView1.Rows[e.RowIndex].Cells[ROW].Style.BackColor = Color.White;
                                    }
                                    else
                                    {
                                        dgr.DefaultCellStyle.BackColor = Color.White;
                                    }

                                } 
                        } 
                        catch   (Exception ex) 
                        { 
                                MessageBox.Show(ex.Message); 
                        } 

        }



    



        private void UpdateDISC(string u_acme_discoo, string DOCENTRY)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE OPOR SET u_acme_discoo=@u_acme_discoo WHERE DOCENTRY=@DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@u_acme_discoo", u_acme_discoo));
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


        private void UpdateDSC(string U_ACME_DSCRIPTION, string DOCENTRY, string LINENUM)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET U_ACME_DSCRIPTION=@U_ACME_DSCRIPTION WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@U_ACME_DSCRIPTION", U_ACME_DSCRIPTION));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {


                dataGridView1.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value = "";

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text  = "";
            textBox3.Text = "";
            textBox4.Text = "";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox1.Text = "";
            textBox4.Text = "";
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox1.Text = "";
        }

     
   
     

    }
}